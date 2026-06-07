#!/usr/bin/env python3
"""
Email-based Slack Monitor — 透過 Outlook for Mac 本地資料庫偵測 Slack 通知

流程:
  Slack channel 新訊息 → Slack 寄 email → Outlook 收到 → 本腳本掃描 → 更新 dashboard

讀取: ~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite
"""

import json, re, sqlite3, os
from pathlib import Path
from datetime import datetime, timedelta
import time

BASE_DIR = Path(__file__).parent
SCHEDULES_FILE = BASE_DIR / "consolidated_schedules.json"
SLACK_CACHE_FILE = BASE_DIR / "slack_cache.json"
ALERT_FILE = BASE_DIR / "confluence_update_alerts.txt"
EMAIL_CACHE_FILE = BASE_DIR / "email_scan_cache.json"

OUTLOOK_DB = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/Outlook/"
    "Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite"
)

# Slack notification patterns
SLACK_SENDERS = ["slack", "notification@slack", "no-reply@slack"]
CHANNEL_KEYWORDS = ["compliance-prd-review", "prd-review", "compliance-prd"]

# Project name extraction
SKIP_WORDS = {
    'The','This','That','Please','Review','New','Update','Link','Here','See',
    'Check','For','From','With','PRD','EVT','DVT','NPI','Compliance','Blink',
    'Ring','Hello','Hi','Thanks','Note','Slack','Channel','Team','All','Done',
    'Message','Posted','Sent','Reply','Thread','Email','Notification',
    'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday',
    'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec',
    'Amazon','Safety','Energy','GMA','EMC','Test','Testing','Report',
    'Schedule','Milestone','Program','Plan','EXTERNAL','CAUTION',
}


def load_cache():
    if EMAIL_CACHE_FILE.exists():
        try:
            return json.loads(EMAIL_CACHE_FILE.read_text())
        except:
            pass
    return {"last_scan_time": 0, "processed_ids": []}


def save_cache(cache):
    EMAIL_CACHE_FILE.write_text(json.dumps(cache, indent=2))


def get_known_projects():
    projects = set()
    if SCHEDULES_FILE.exists():
        projects.update(json.loads(SCHEDULES_FILE.read_text()).keys())
    if SLACK_CACHE_FILE.exists():
        try:
            sc = json.loads(SLACK_CACHE_FILE.read_text())
            projects.update(sc.get("known_projects", []))
        except:
            pass
    return projects


def scan_outlook_for_slack():
    """Scan Outlook.sqlite for Slack notification emails"""
    if not os.path.exists(OUTLOOK_DB):
        print("  ❌ Outlook database not found")
        return []

    cache = load_cache()
    last_time = cache.get("last_scan_time", 0)
    processed = set(cache.get("processed_ids", []))

    # Connect to Outlook DB (read-only to avoid conflicts)
    conn = sqlite3.connect(f"file:{OUTLOOK_DB}?mode=ro", uri=True)
    cur = conn.cursor()

    # Search for recent emails from Slack or containing PRD/compliance channel refs
    # Outlook stores time as Apple epoch (seconds since 2001-01-01)
    # We look for emails from last 24 hours
    time_cutoff = last_time if last_time > 0 else (time.time() - 978307200 - 86400)

    cur.execute('''
        SELECT Record_RecordID, Message_NormalizedSubject, Message_SenderList, 
               Message_Preview, Message_TimeReceived
        FROM Mail 
        WHERE Message_TimeReceived > ?
        AND (
            Message_SenderList LIKE '%slack%'
            OR Message_NormalizedSubject LIKE '%prd-review%'
            OR Message_NormalizedSubject LIKE '%compliance-prd%'
            OR Message_Preview LIKE '%ring-compliance-prd%'
            OR Message_Preview LIKE '%blink-compliance-prd%'
        )
        ORDER BY Message_TimeReceived DESC
        LIMIT 50
    ''', (time_cutoff,))

    rows = cur.fetchall()
    conn.close()

    # Process results
    new_projects = []
    known = get_known_projects()

    for row in rows:
        record_id, subject, sender, preview, recv_time = row
        msg_id = str(record_id)

        if msg_id in processed:
            continue
        processed.add(msg_id)

        # Combine subject and preview for scanning
        text = f"{subject or ''} {preview or ''}"

        # Check if it's related to our channels
        text_lower = text.lower()
        is_relevant = (
            any(kw in text_lower for kw in CHANNEL_KEYWORDS)
            or ('slack' in (sender or '').lower() and 'prd' in text_lower)
            or ('confluence' in text_lower and 'prd' in text_lower)
        )

        if not is_relevant:
            continue

        # Extract project info
        info = extract_project(text)
        if info["name"] and info["name"] not in known:
            info["source"] = "outlook_slack_email"
            info["email_subject"] = subject
            info["detected_at"] = datetime.now().isoformat()
            new_projects.append(info)
            known.add(info["name"])

    # Update cache
    if rows:
        max_time = max(r[4] for r in rows)
        cache["last_scan_time"] = max_time
    cache["processed_ids"] = list(processed)[-200:]
    save_cache(cache)

    return new_projects


def extract_project(text):
    """Extract project name and PRD link from email text"""
    info = {"name": None, "prd_link": None}

    # Confluence links
    links = re.findall(r'https?://confluence\.atl\.ring\.com/[^\s<>"\']+', text)
    if links:
        info["prd_link"] = links[0]
        # Try to get project name from link
        m = re.search(r'/([A-Z][a-z]+(?:\+[A-Z][a-z]+)*)\+?-?\+?(?:PRD|Product)', links[0])
        if m:
            info["name"] = m.group(1).replace('+', ' ')

    # PRD for/: ProjectName
    if not info["name"]:
        m = re.search(r'(?:PRD|prd)\s*(?:for|:|–|-|review)?\s*([A-Z][a-zA-Z]{2,15})', text)
        if m and m.group(1) not in SKIP_WORDS:
            info["name"] = m.group(1)

    # Code names (capitalized words)
    if not info["name"]:
        words = re.findall(r'\b([A-Z][a-z]{2,15})\b', text)
        candidates = [w for w in words if w not in SKIP_WORDS]
        if candidates:
            info["name"] = candidates[0]

    return info


def process_new_projects(new_projects):
    """Add new projects and generate alerts"""
    # Add to schedules
    schedules = json.loads(SCHEDULES_FILE.read_text()) if SCHEDULES_FILE.exists() else {}
    for p in new_projects:
        if p["name"] not in schedules:
            schedules[p["name"]] = {
                "evt": "TBD", "dvt": "TBD", "ok2mp": "TBD",
                "ship": "TBD", "announce": "TBD", "street": "TBD",
                "prd_link": p.get("prd_link", ""),
                "source": "slack_email"
            }
    SCHEDULES_FILE.write_text(json.dumps(schedules, indent=2, ensure_ascii=False))

    # Update slack cache
    sc = json.loads(SLACK_CACHE_FILE.read_text()) if SLACK_CACHE_FILE.exists() else {"known_projects": []}
    for p in new_projects:
        if p["name"] not in sc["known_projects"]:
            sc["known_projects"].append(p["name"])
    SLACK_CACHE_FILE.write_text(json.dumps(sc, indent=2))

    # Alert file
    alert = f"\n{'='*50}\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] 📧 新案件偵測 (via Slack email)\n"
    for p in new_projects:
        alert += f"  🆕 {p['name']} - PRD: {p.get('prd_link','N/A')}\n"
    alert += f"  📋 請更新: Confluence NPI Overview + SharePoint Excel\n{'='*50}\n"
    with open(ALERT_FILE, "a") as f:
        f.write(alert)


if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 📧 Outlook Slack Email Monitor")

    new_projects = scan_outlook_for_slack()
    print(f"  New projects found: {len(new_projects)}")

    if new_projects:
        process_new_projects(new_projects)
        for p in new_projects:
            print(f"  🆕 {p['name']} - {p.get('prd_link', 'no link')}")
    else:
        print("  ✅ No new projects. Monitoring continues...")
