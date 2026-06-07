#!/usr/bin/env python3
"""
RBKS Slack Channel Monitor — PRD Review Auto-Detect
監控 Slack channels 偵測新 PRD/案件，自動更新 Dashboard 並通知更新 Confluence

Channels:
  - #blink-compliance-prd-reviews (C09T0G17V8S)
  - #ring-compliance-prd-reviews (C09T0NHULCA)
"""

import json, re
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
SLACK_CACHE_FILE = BASE_DIR / "slack_cache.json"
SCHEDULES_FILE = BASE_DIR / "consolidated_schedules.json"
ALERT_FILE = BASE_DIR / "confluence_update_alerts.txt"

CHANNELS = {
    "blink-compliance-prd-reviews": "C09T0G17V8S",
    "ring-compliance-prd-reviews": "C09T0NHULCA",
}


def load_slack_cache():
    if SLACK_CACHE_FILE.exists():
        try:
            return json.loads(SLACK_CACHE_FILE.read_text())
        except:
            pass
    return {"known_projects": [], "last_check": ""}


def save_slack_cache(cache):
    SLACK_CACHE_FILE.write_text(json.dumps(cache, indent=2, ensure_ascii=False))


def extract_project_from_message(text):
    """Extract project name and PRD link from a Slack message."""
    info = {"name": None, "prd_link": None, "product": ""}

    # Extract Confluence links
    links = re.findall(r'https?://confluence\.atl\.ring\.com/[^\s<>|]+', text)
    if links:
        info["prd_link"] = links[0]
        # Try project name from link path
        for link in links:
            m = re.search(r'/(\w+)\+?-?\+?(?:PRD|Program|Schedule)', link)
            if m:
                info["name"] = m.group(1).replace('+', ' ')

    # PRD: ProjectName or PRD for ProjectName
    if not info["name"]:
        m = re.search(r'(?:PRD|prd)\s*(?:for|:|–|-|\s)\s*([A-Z][a-zA-Z]+)', text)
        if m:
            info["name"] = m.group(1)

    # Look for code names (capitalized words)
    if not info["name"]:
        skip = {'The','This','That','Please','Review','New','Update','Link',
                'Here','See','Check','For','From','With','PRD','EVT','DVT',
                'NPI','Compliance','Blink','Ring','Hello','Hi','Thanks','Note'}
        words = re.findall(r'\b([A-Z][a-z]{2,15})\b', text)
        candidates = [w for w in words if w not in skip]
        if candidates:
            info["name"] = candidates[0]

    return info


def generate_confluence_alert(project_info):
    """Append alert to file for user to update Ring Confluence"""
    alert = (
        f"\n{'='*60}\n"
        f"🆕 新案件偵測 New Project: {project_info.get('name','?')}\n"
        f"   PRD: {project_info.get('prd_link','N/A')}\n"
        f"   來源: Slack PRD Review Channel\n"
        f"   時間: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"   \n"
        f"   📋 請更新 Ring Confluence NPI Overview:\n"
        f"   https://confluence.atl.ring.com/spaces/RCC/pages/3355651641\n"
        f"{'='*60}\n"
    )
    with open(ALERT_FILE, "a") as f:
        f.write(alert)
    return alert


if __name__ == "__main__":
    print("Slack Monitor — RBKS Compliance")
    print(f"Channels: {list(CHANNELS.keys())}")
    cache = load_slack_cache()
    print(f"Known projects: {len(cache.get('known_projects', []))}")
    print(f"\nAlert file: {ALERT_FILE}")
    print("This module is imported by monitor_daemon.py")
