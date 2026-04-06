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
SEPARATOR = "|||"


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


def folder_unread_count(folder_name):
    """
    Returns the unread count for a named folder, or None if Outlook is not
    running or the folder doesn't exist / isn't accessible.
    """
    safe = folder_name.replace("\\", "\\\\").replace('"', '\\"')
    script = f'''
tell application "Microsoft Outlook"
    try
        set f to mail folder "{safe}"
        return unread count of f
    end try
    return -1
end tell
'''
    raw = run_applescript(script)
    if raw is None:
        return None
    try:
        val = int(raw)
        return None if val == -1 else val
    except ValueError:
        return None


def verify_folder(folder_name):
    """Returns True if Outlook can find the named folder."""
    return folder_unread_count(folder_name) is not None


def try_get_newest_unread(folder_name):
    """
    Attempts to retrieve sender + subject of the newest unread message.
    Returns (None, None) if unavailable.
    """
    safe = folder_name.replace("\\", "\\\\").replace('"', '\\"')
    script = f'''
tell application "Microsoft Outlook"
    try
        set f to mail folder "{safe}"
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
        self._build_menu()
        self._start_polling()

    # ── Menu ──────────────────────────────────────────────────────────────────

    def _build_menu(self):
        with self._lock:
            watched = list(self.config["watched"])

        self._status_item = rumps.MenuItem(f"Watching {len(watched)} folder(s)")

        watched_items = []
        for name in watched:
            item = rumps.MenuItem(f"✓ {name}", callback=self._remove_folder)
            item._folder_name = name
            watched_items.append(item)

        items = [self._status_item, None]
        if watched_items:
            items += watched_items + [None]
        items += [
            rumps.MenuItem("Add Folder...", callback=self._on_add_folder),
            None,
            rumps.MenuItem("Quit", callback=rumps.quit_application),
        ]

        self.menu = items

    def _remove_folder(self, sender):
        name = sender._folder_name
        with self._lock:
            self.config["watched"] = [f for f in self.config["watched"] if f != name]
            save_config(self.config)
        self._build_menu()

    def _on_add_folder(self, _):
        win = rumps.Window(
            title="Add Folder",
            message="Enter the exact Outlook folder name to watch:",
            default_text="",
            ok="Add",
            cancel="Cancel",
            dimensions=(300, 24),
        )
        response = win.run()
        if not response.clicked:
            return

        name = response.text.strip()
        if not name:
            return

        with self._lock:
            if name in self.config["watched"]:
                rumps.alert(title="Already watching", message=f'"{name}" is already in your watch list.')
                return

        # Verify the folder exists in Outlook before adding
        if not verify_folder(name):
            rumps.alert(
                title="Folder not found",
                message=f'Outlook couldn\'t find a folder named "{name}".\n\nCheck the exact spelling — it\'s case-sensitive.',
            )
            return

        with self._lock:
            self.config["watched"].append(name)
            save_config(self.config)

        self._build_menu()

    # ── Polling ───────────────────────────────────────────────────────────────

    def _start_polling(self):
        t = threading.Thread(target=self._poll_loop, daemon=True)
        t.start()

    def _poll_loop(self):
        # Seed counts so we don't spam notifications on startup
        with self._lock:
            watched = list(self.config["watched"])

        for name in watched:
            count = folder_unread_count(name)
            if count is not None:
                with self._lock:
                    self.config["last_counts"].setdefault(name, count)
        with self._lock:
            save_config(self.config)

        while True:
            time.sleep(POLL_SECONDS)
            self._check_new_mail()

    def _check_new_mail(self):
        with self._lock:
            watched = list(self.config["watched"])
            last = dict(self.config["last_counts"])

        for name in watched:
            count = folder_unread_count(name)
            if count is None:
                continue  # Outlook not running or folder gone

            prev = last.get(name, count)

            if count > prev:
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
