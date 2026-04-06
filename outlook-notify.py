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
    Single AppleScript call for efficiency — one round-trip regardless of folder count.
    """
    script = f'''
tell application "Microsoft Outlook"
    set rows to {{}}
    repeat with f in every mail folder
        set n to name of f
        if n does not start with "Placeholder" then
            set end of rows to n & "{SEPARATOR}" & (unread count of f)
        end if
    end repeat
    return rows
end tell
'''
    raw = run_applescript(script)
    if not raw:
        return []

    folders = []
    # AppleScript returns a list as comma-separated items
    for item in raw.split(", "):
        item = item.strip()
        if SEPARATOR in item:
            name, _, count_str = item.partition(SEPARATOR)
            try:
                folders.append((name.strip(), int(count_str.strip())))
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
        self._cached_folders = []
        self._status_item = None
        self._build_menu(refresh_folders=True)
        self._start_polling()

    # ── Menu ──────────────────────────────────────────────────────────────────

    def _build_menu(self, refresh_folders=True):
        """
        Rebuild the full menu. Pass refresh_folders=True to re-query Outlook.
        On folder toggles we only update state in-place — no full rebuild needed.
        """
        if refresh_folders:
            self._cached_folders = get_all_folders()

        with self._lock:
            watched = set(self.config["watched"])

        self._status_item = rumps.MenuItem(f"Watching {len(watched)} folder(s)")

        # Build Watch Folders submenu
        watch_sub = rumps.MenuItem("Watch Folders")
        watch_sub.add(rumps.MenuItem("↻ Refresh List", callback=self._on_refresh_list))
        watch_sub.add(None)  # separator

        for name, _ in sorted(self._cached_folders, key=lambda x: x[0].lower()):
            item = rumps.MenuItem(name, callback=self._toggle_folder)
            item.state = 1 if name in watched else 0
            watch_sub.add(item)

        self.menu = [
            self._status_item,
            None,
            watch_sub,
            None,
            rumps.MenuItem("Quit", callback=rumps.quit_application),
        ]

    def _on_refresh_list(self, _):
        self._build_menu(refresh_folders=True)

    def _toggle_folder(self, sender):
        with self._lock:
            if sender.state:
                self.config["watched"] = [
                    f for f in self.config["watched"] if f != sender.title
                ]
                sender.state = 0
            else:
                self.config["watched"].append(sender.title)
                sender.state = 1
            n = len(self.config["watched"])
            save_config(self.config)

        # Update count label in-place — submenu stays open
        if self._status_item is not None:
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
