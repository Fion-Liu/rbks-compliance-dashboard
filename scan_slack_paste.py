#!/usr/bin/env python3
"""
Slack Message Scanner — 貼上 Slack 訊息文字即可自動掃描新案件

使用方式：
  1. 在 Slack channel 全選複製訊息
  2. 貼到 slack_messages.txt
  3. 執行: python3 scan_slack_paste.py
  
  或者直接用剪貼簿：
  pbpaste | python3 scan_slack_paste.py --stdin
"""

import json, re, sys
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
INPUT_FILE = BASE_DIR / "slack_messages.txt"
SCHEDULES_FILE = BASE_DIR / "consolidated_schedules.json"
SLACK_CACHE_FILE = BASE_DIR / "slack_cache.json"
ALERT_FILE = BASE_DIR / "confluence_update_alerts.txt"

# Known internal project code name patterns (short, capitalized)
SKIP_WORDS = {
    'The','This','That','Please','Review','New','Update','Link','Here','See',
    'Check','For','From','With','PRD','EVT','DVT','NPI','Compliance','Blink',
    'Ring','Hello','Hi','Thanks','Note','Slack','Channel','Team','All','Done',
    'Yes','No','Will','Can','Did','Has','Are','Was','Not','But','Also','Just',
    'Need','Make','Let','Get','Got','Any','Now','Today','Monday','Tuesday',
    'Wednesday','Thursday','Friday','Action','Item','Status','Share','Doc',
    'Meeting','FYI','Reminder','Question','Answer','Follow','Kick','Off',
    'Start','End','Date','Week','Month','Jan','Feb','Mar','Apr','May','Jun',
    'Jul','Aug','Sep','Oct','Nov','Dec','Amazon','Safety','Energy','GMA',
    'EMC','Test','Testing','Report','Schedule','Milestone','Program','Plan',
}


def load_known_projects():
    projects = set()
    if SCHEDULES_FILE.exists():
        projects.update(json.loads(SCHEDULES_FILE.read_text()).keys())
    if SLACK_CACHE_FILE.exists():
        cache = json.loads(SLACK_CACHE_FILE.read_text())
        projects.update(cache.get("known_projects", []))
    return projects


def extract_projects_from_text(text):
    """Scan text for project names + PRD links"""
    found = []
    lines = text.split('\n')

    for line in lines:
        # Look for PRD links
        links = re.findall(r'https?://confluence\.atl\.ring\.com/[^\s<>|]+', line)
        
        # Look for project names near PRD mentions
        prd_names = re.findall(r'(?:PRD|prd)\s*(?:for|:|–|-|review)?\s*([A-Z][a-zA-Z]+)', line)
        
        # Look for capitalized code names
        code_names = re.findall(r'\b([A-Z][a-z]{2,15})\b', line)
        code_names = [w for w in code_names if w not in SKIP_WORDS]
        
        # Combine findings
        names = prd_names + code_names
        if names:
            for name in set(names):
                entry = {"name": name, "prd_link": links[0] if links else None}
                # Avoid duplicates
                if not any(f["name"] == name for f in found):
                    found.append(entry)

    return found


def main():
    # Read input
    if "--stdin" in sys.argv:
        text = sys.stdin.read()
    elif INPUT_FILE.exists():
        text = INPUT_FILE.read_text()
    else:
        print(f"❌ 請先把 Slack 訊息貼到: {INPUT_FILE}")
        print(f"   或使用: pbpaste | python3 scan_slack_paste.py --stdin")
        return

    print(f"📋 掃描 {len(text)} 字元...")
    
    known = load_known_projects()
    found = extract_projects_from_text(text)
    
    # Filter to new projects only
    new_projects = [p for p in found if p["name"] not in known]
    
    print(f"\n📊 結果:")
    print(f"   掃描到案件: {len(found)}")
    print(f"   已知案件: {len(found) - len(new_projects)}")
    print(f"   🆕 新案件: {len(new_projects)}")
    
    if new_projects:
        print(f"\n{'='*50}")
        print("🆕 新偵測案件 New Projects Detected:")
        print(f"{'='*50}")
        for p in new_projects:
            print(f"   • {p['name']}")
            if p.get('prd_link'):
                print(f"     PRD: {p['prd_link']}")
        
        # Save to cache
        cache = json.loads(SLACK_CACHE_FILE.read_text()) if SLACK_CACHE_FILE.exists() else {"known_projects": []}
        for p in new_projects:
            if p["name"] not in cache["known_projects"]:
                cache["known_projects"].append(p["name"])
        SLACK_CACHE_FILE.write_text(json.dumps(cache, indent=2))
        
        # Add to schedules (empty placeholder)
        schedules = json.loads(SCHEDULES_FILE.read_text()) if SCHEDULES_FILE.exists() else {}
        for p in new_projects:
            if p["name"] not in schedules:
                schedules[p["name"]] = {"prd_link": p.get("prd_link", ""), "source": "slack"}
        SCHEDULES_FILE.write_text(json.dumps(schedules, indent=2, ensure_ascii=False))
        
        # Generate alerts
        alert = f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Slack 掃描發現新案件:\n"
        for p in new_projects:
            alert += f"  • {p['name']} - PRD: {p.get('prd_link','N/A')}\n"
        alert += f"  📋 請更新 Confluence NPI Overview + SharePoint Excel\n"
        with open(ALERT_FILE, "a") as f:
            f.write(alert)
        
        print(f"\n✅ 已儲存。接下來需要:")
        print(f"   1. 更新 Dashboard: 告訴 Kiro 新案件資訊")
        print(f"   2. 更新 Confluence: https://confluence.atl.ring.com/spaces/RCC/pages/3355651641")
        print(f"   3. 更新 SharePoint Excel 人員分配")
    else:
        print("\n✅ 沒有新案件。")


if __name__ == "__main__":
    main()
