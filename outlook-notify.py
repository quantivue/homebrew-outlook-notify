#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook Subfolder Notifier
Watches Exchange subfolders in Outlook for Mac for new unread messages
and fires macOS notifications — including sender and subject when available.

Solves the gap: Outlook rules silently move mail to subfolders with no alert.
"""

import rumps
import subprocess
import json
import os
import threading
import time

CONFIG_DIR = os.path.expanduser("~/.config/outlook-notify")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
POLL_SECONDS = 30
SEPARATOR = "|||"  # delimiter safe against folder names and email content

# Internal Outlook stores/system folders the user never sees in the sidebar
_SYSTEM_FOLDERS = {
    "", "On My Computer", "Saved Messages", "Temporary Items",
    "Auto-Saved Messages", "Outbox",
}


# ─── AppleScript helpers ──────────────────────────────────────────────────────

def run_applescript(script):
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout.strip()
    return None


def notify(title, body):
    t = title.replace("\\", "\\\\").replace('"', '\\"')
    b = body.replace("\\", "\\\\").replace('"', '\\"')
    run_applescript(f'display notification "{b}" with title "{t}"')


def get_all_folders():
    """
    Returns list of (folder_name, unread_count) tuples from Outlook.
    Uses a breadth-first queue to recurse into subfolders — 'every mail folder'
    alone only returns top-level folders.
    """
    script = f'''
tell application "Microsoft Outlook"
    set allRows to {{}}
    set queue to (every mail folder)
    repeat while (count of queue) > 0
        set f to item 1 of queue
        set queue to rest of queue
        set n to name of f
        if n does not start with "Placeholder" then
            set end of allRows to n & "{SEPARATOR}" & (unread count of f)
        end if
        repeat with child in (mail folders of f)
            set end of queue to child
        end repeat
    end repeat
    return allRows
end tell
'''
    raw = run_applescript(script)
    if not raw:
        return []

    # Parse, filter system folders, and deduplicate by name (first occurrence wins).
    # Duplicates arise when the same folder exists under multiple accounts/stores.
    seen = set()
    folders = []
    for item in raw.split(", "):
        item = item.strip()
        if SEPARATOR not in item:
            continue
        name, _, count_str = item.partition(SEPARATOR)
        name = name.strip()
        if not name or name in _SYSTEM_FOLDERS or name.startswith("Placeholder"):
            continue
        if name in seen:
            continue
        seen.add(name)
        try:
            folders.append((name, int(count_str.strip())))
        except ValueError:
            pass
    return folders


def try_get_newest_unread(folder_name):
    """
    Attempts to retrieve sender + subject of the newest unread message.
    New Outlook for Mac doesn't always populate message objects via AppleScript,
    so this may return (None, None) — handled gracefully by callers.
    """
    safe_name = folder_name.replace("\\", "\\\\").replace('"', '\\"')
    script = f'''
tell application "Microsoft Outlook"
    try
        set f to mail folder "{safe_name}"
        set unread_msgs to (messages of f whose is read is false)
        if (count of unread_msgs) > 0 then
            set m to item 1 of unread_msgs
            set s to sender of m as string
            set sub to subject of m as string
            return s & "{SEPARATOR}" & sub
        end if
    end try
    return ""
end tell
'''
    raw = run_applescript(script)
    if raw and SEPARATOR in raw:
        sender, _, subject = raw.partition(SEPARATOR)
        return sender.strip(), subject.strip()
    return None, None


def pick_folders(current_watched):
    """
    Opens a native macOS multi-select list dialog via AppleScript.
    Returns the new list of selected folder names, or None if cancelled.
    """
    folders = get_all_folders()
    if not folders:
        run_applescript('display alert "Outlook is not running or has no folders." as warning')
        return None

    names = sorted([name for name, _ in folders], key=str.lower)

    # Build AppleScript list literals
    as_list = "{" + ", ".join(f'"{n}"' for n in names) + "}"
    pre_selected = "{" + ", ".join(f'"{n}"' for n in names if n in current_watched) + "}"

    script = f'''
set result to choose from list {as_list} ¬
    with title "Outlook Notify" ¬
    with prompt "Select folders to watch for new mail:" ¬
    default items {pre_selected} ¬
    with multiple selections allowed ¬
    with empty selection allowed
if result is false then
    return "CANCELLED"
end if
set output to ""
repeat with i from 1 to count of result
    if i > 1 then set output to output & "{SEPARATOR}"
    set output to output & (item i of result)
end repeat
return output
'''
    raw = run_applescript(script)
    if raw is None or raw == "CANCELLED":
        return None
    if raw == "":
        return []
    return [name.strip() for name in raw.split(SEPARATOR) if name.strip()]


# ─── Config ───────────────────────────────────────────────────────────────────

def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH) as f:
                return json.load(f)
        except Exception:
            pass
    return {"watched": [], "last_counts": {}}


def save_config(config):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=2)


# ─── App ──────────────────────────────────────────────────────────────────────

class OutlookNotify(rumps.App):
    def __init__(self):
        super().__init__("📬", quit_button=None)
        self.config = load_config()
        self._lock = threading.Lock()

        with self._lock:
            n = len(self.config["watched"])

        self._status_item = rumps.MenuItem(f"Watching {n} folder(s)")

        # Menu is built once and never rebuilt — avoids the accumulation bug.
        # Only _status_item.title changes over time.
        self.menu = [
            self._status_item,
            None,
            rumps.MenuItem("Select Folders...", callback=self._on_select_folders),
            None,
            rumps.MenuItem("Quit", callback=rumps.quit_application),
        ]

        self._start_polling()

    # ── Folder picker ─────────────────────────────────────────────────────────

    def _on_select_folders(self, _):
        with self._lock:
            current = set(self.config["watched"])

        selected = pick_folders(current)
        if selected is None:
            return  # user cancelled — no change

        with self._lock:
            self.config["watched"] = selected
            save_config(self.config)
            n = len(selected)

        self._status_item.title = f"Watching {n} folder(s)"

    # ── Polling ───────────────────────────────────────────────────────────────

    def _start_polling(self):
        t = threading.Thread(target=self._poll_loop, daemon=True)
        t.start()

    def _poll_loop(self):
        # Seed counts on first run so we don't spam on startup
        folders = get_all_folders()
        with self._lock:
            for name, count in folders:
                self.config["last_counts"].setdefault(name, count)
            save_config(self.config)

        while True:
            time.sleep(POLL_SECONDS)
            self._check_new_mail()

    def _check_new_mail(self):
        folders = get_all_folders()
        if not folders:
            return  # Outlook not running or AppleScript failed

        with self._lock:
            watched = set(self.config["watched"])
            last = dict(self.config["last_counts"])

        for name, count in folders:
            prev = last.get(name, count)

            if name in watched and count > prev:
                delta = count - prev
                sender, subject = try_get_newest_unread(name)

                if sender and subject:
                    body = f"{sender}: {subject}"
                else:
                    body = f"{delta} new message{'s' if delta > 1 else ''}"

                notify(f"📬 {name}", body)

            last[name] = count

        with self._lock:
            self.config["last_counts"] = last
            save_config(self.config)


if __name__ == "__main__":
    OutlookNotify().run()
