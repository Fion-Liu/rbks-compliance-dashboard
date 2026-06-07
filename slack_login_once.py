#!/usr/bin/env python3
"""
Slack 一次性登入 — 存下 session 後就能 headless 自動讀取

只需要執行一次！登入後 session 會存在 slack_auth_state.json
之後 monitor_daemon.py 就能 headless 自動讀 Slack private channels
"""
from playwright.sync_api import sync_playwright
from pathlib import Path
import json

BASE_DIR = Path(__file__).parent
AUTH_FILE = BASE_DIR / "slack_auth_state.json"
SLACK_URL = "https://ring.enterprise.slack.com"


def main():
    print("🔐 Slack 一次性登入")
    print("   會開啟瀏覽器，請手動登入 Ring Slack")
    print("   登入成功後 session 會自動儲存\n")

    with sync_playwright() as p:
        browser = p.firefox.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto(SLACK_URL, timeout=60000)
        print("⏳ 請在瀏覽器中登入 Ring Slack...")
        print("   登入完成後，等待頁面載入到 channel 畫面")
        print("   你有 5 分鐘的時間完成登入...")

        # Wait for any of these indicators that we're logged in
        try:
            page.wait_for_url("**slack.com/client/**", timeout=300000)
        except:
            try:
                page.wait_for_url("**slack.com/**", timeout=10000)
            except:
                pass

        # Extra wait for page to fully load
        page.wait_for_timeout(5000)
        print(f"   目前網址: {page.url}")

        # Save auth state
        context.storage_state(path=str(AUTH_FILE))
        print(f"✅ Session 已存到: {AUTH_FILE}")
        print("   之後 headless 自動讀取不需要再登入！")

        browser.close()


if __name__ == "__main__":
    main()
