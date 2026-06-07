#!/usr/bin/env python3
"""
Headless Slack Message Fetcher — 自動讀取 private channel 訊息
前提：需先執行一次 slack_login_once.py 存下 session

被 monitor_daemon.py 自動呼叫，每 15 分鐘執行一次
"""
import json, re, sys
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
AUTH_FILE = BASE_DIR / "slack_auth_state.json"
OUTPUT_FILE = BASE_DIR / "slack_latest_msgs.json"

CHANNELS = {
    "ring-compliance-prd-reviews": "C09T0NHULCA",
    "blink-compliance-prd-reviews": "C09T0G17V8S",
}

SLACK_BASE = "https://app.slack.com/client/EHPH7AM3L"


def fetch_messages():
    if not AUTH_FILE.exists():
        print("❌ No auth state. Run slack_login_once.py first!")
        return []

    messages = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(storage_state=str(AUTH_FILE))
        page = context.new_page()

        for ch_name, ch_id in CHANNELS.items():
            try:
                url = f"{SLACK_BASE}/{ch_id}"
                page.goto(url, wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(5000)  # Wait for messages to load

                # Try to extract token and use API
                token = page.evaluate("""() => {
                    try {
                        // Check boot_data
                        if (window.boot_data && window.boot_data.api_token) 
                            return window.boot_data.api_token;
                        // Check localStorage
                        for (let i = 0; i < localStorage.length; i++) {
                            const val = localStorage.getItem(localStorage.key(i));
                            if (val) {
                                const m = val.match(/xoxc-[0-9]+-[0-9]+-[0-9]+-[a-f0-9]+/);
                                if (m) return m[0];
                            }
                        }
                        // Check redux store
                        const el = document.querySelector('[data-app-env]');
                        if (el) {
                            const env = JSON.parse(el.getAttribute('data-app-env'));
                            if (env && env.token) return env.token;
                        }
                    } catch(e) {}
                    return null;
                }""")

                if token:
                    # Use API for clean data
                    result = page.evaluate(f"""async () => {{
                        const resp = await fetch('/api/conversations.history', {{
                            method: 'POST',
                            headers: {{'Content-Type': 'application/x-www-form-urlencoded'}},
                            body: 'token={token}&channel={ch_id}&limit=30'
                        }});
                        return await resp.json();
                    }}""")
                    if result.get('ok'):
                        for msg in result.get('messages', []):
                            msg['_channel'] = ch_name
                            messages.append(msg)
                        print(f"  ✅ #{ch_name}: {len(result.get('messages',[]))} msgs (API)")
                    else:
                        raise Exception(f"API error: {result.get('error')}")
                else:
                    # Fallback: scrape text from DOM
                    page.wait_for_timeout(3000)
                    texts = page.evaluate("""() => {
                        const msgs = [];
                        const items = document.querySelectorAll('[data-qa="virtual-list-item"], .c-message_kit__blocks');
                        items.forEach(el => {
                            const text = el.innerText.trim();
                            if (text && text.length > 5) msgs.push(text);
                        });
                        return msgs;
                    }""")
                    for text in texts:
                        messages.append({"text": text, "_channel": ch_name, "_scraped": True})
                    print(f"  ✅ #{ch_name}: {len(texts)} msgs (scraped)")

            except Exception as e:
                print(f"  ⚠️ #{ch_name}: {e}")

        # Save updated auth state (session refresh)
        context.storage_state(path=str(AUTH_FILE))
        browser.close()

    # Save messages
    OUTPUT_FILE.write_text(json.dumps(messages, indent=2, ensure_ascii=False))
    print(f"  💾 Saved {len(messages)} total messages")
    return messages


if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Headless Slack fetch...")
    msgs = fetch_messages()
    if msgs:
        # Show recent PRD-related messages
        prd_msgs = [m for m in msgs if 'PRD' in m.get('text', '').upper() or 'confluence' in m.get('text', '').lower()]
        if prd_msgs:
            print(f"\n  PRD-related messages: {len(prd_msgs)}")
            for m in prd_msgs[:5]:
                print(f"    • {m.get('text','')[:80]}")
