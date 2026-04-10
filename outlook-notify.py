#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook Subfolder Notifier
Polls Exchange subfolders via Mac Mail's AppleScript API (which has full
Exchange folder access) and fires macOS notifications when unread counts rise.

Outlook rules silently deposit mail into subfolders — this fills that gap.
Mail.app must be running (add it to Login Items so it starts at login).
"""

import rumps
import shutil
import subprocess
import json
import os
import threading
import time

CONFIG_DIR = os.path.expanduser("~/.config/outlook-notify")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
POLL_SECONDS = 30
SEPARATOR = "|||"
RECORD_SEP = "^^^"

# Resolve terminal-notifier once at import time.  shutil.which works when
# launched from a shell; the fallback paths cover launchd (no PATH inherited).
NOTIFIER_BIN = (
    shutil.which("terminal-notifier")
    or next(
        (p for p in [
            "/opt/homebrew/bin/terminal-notifier",   # Apple Silicon
            "/usr/local/bin/terminal-notifier",       # Intel
        ] if os.path.isfile(p)),
        None,
    )
)


# ─── AppleScript helper ───────────────────────────────────────────────────────

def run_applescript(script):
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout.strip()
    return None


# ─── Notification helpers ─────────────────────────────────────────────────────


def notify(title, body):
    if NOTIFIER_BIN is None:
        return
    try:
        subprocess.Popen([
            NOTIFIER_BIN,
            "-title", title,
            "-message", body,
            "-sender", "com.apple.mail",
            "-sound", "default",
        ])
    except OSError:
        pass  # binary vanished after startup — silently skip


def _mail_script(inner):
    """Wrap inner AppleScript in a Mail tell block with error handling."""
    return f'''
tell application "Mail"
    try
        {inner}
    end try
end tell
'''


def folder_unread_count(folder_name):
    """
    Returns the unread count for the named mailbox in Mail.app,
    or None if Mail isn't running or the mailbox doesn't exist.
    """
    safe = folder_name.replace("\\", "\\\\").replace('"', '\\"')
    script = _mail_script(f'''
        repeat with acct in every account
            repeat with mbx in (every mailbox of acct)
                if (name of mbx) is "{safe}" then
                    return unread count of mbx
                end if
            end repeat
        end repeat
        return -1
    ''')
    raw = run_applescript(script)
    if raw is None:
        return None
    try:
        val = int(raw)
        return None if val == -1 else val
    except ValueError:
        return None


def batch_unread_counts(folder_names):
    """
    Returns {folder_name: int} for all folders in a single osascript call.
    Folders that don't exist in Mail are omitted from the result.
    """
    if not folder_names:
        return {}
    # Build an AppleScript list literal of the target names
    safe_names = [n.replace("\\", "\\\\").replace('"', '\\"') for n in folder_names]
    as_list = ", ".join(f'"{s}"' for s in safe_names)
    script = _mail_script(f'''
        set targets to {{{as_list}}}
        set results to ""
        repeat with acct in every account
            repeat with mbx in (every mailbox of acct)
                set mName to (name of mbx) as text
                if mName is in targets then
                    set cnt to unread count of mbx
                    set results to results & mName & "{SEPARATOR}" & (cnt as text) & "{RECORD_SEP}"
                end if
            end repeat
        end repeat
        return results
    ''')
    raw = run_applescript(script)
    if not raw:
        return {}
    counts = {}
    for record in raw.split(RECORD_SEP):
        if SEPARATOR in record:
            name, _, cnt_str = record.partition(SEPARATOR)
            try:
                counts[name.strip()] = int(cnt_str.strip())
            except ValueError:
                pass
    return counts


def verify_folder(folder_name):
    """Returns True if Mail.app can find the named mailbox."""
    return folder_unread_count(folder_name) is not None


def try_get_newest_unread(folder_name):
    """
    Retrieves sender + subject of the newest unread message via Mail.app.
    Returns (None, None) if unavailable.
    """
    safe = folder_name.replace("\\", "\\\\").replace('"', '\\"')
    script = _mail_script(f'''
        set targetMbx to missing value
        repeat with acct in every account
            repeat with mbx in (every mailbox of acct)
                if (name of mbx) is "{safe}" then
                    set targetMbx to mbx
                    exit repeat
                end if
            end repeat
            if targetMbx is not missing value then exit repeat
        end repeat
        if targetMbx is not missing value then
            set unreadMsgs to (messages of targetMbx whose read status is false)
            if (count of unreadMsgs) > 0 then
                set m to item 1 of unreadMsgs
                return (sender of m) & "{SEPARATOR}" & (subject of m)
            end if
        end if
        return ""
    ''')
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
        self._status_item = None
        # Seed with an empty list so rumps initialises self.menu before _build_menu
        self.menu = []
        self._build_menu()
        self._start_polling()

    # ── Menu ──────────────────────────────────────────────────────────────────

    def _build_menu(self):
        with self._lock:
            watched = list(self.config["watched"])

        # Clear all existing menu items before rebuilding — prevents accumulation
        self.menu.clear()

        self._status_item = rumps.MenuItem(f"Watching {len(watched)} folder(s)")
        self.menu.add(self._status_item)
        self.menu.add(rumps.separator)

        for name in watched:
            item = rumps.MenuItem(f"✓ {name}", callback=self._remove_folder)
            item._folder_name = name
            self.menu.add(item)

        if watched:
            self.menu.add(rumps.separator)

        self.menu.add(rumps.MenuItem("Add Folder...", callback=self._on_add_folder))
        self.menu.add(rumps.separator)
        self.menu.add(rumps.MenuItem("Quit", callback=rumps.quit_application))

    def _remove_folder(self, sender):
        name = sender._folder_name
        with self._lock:
            self.config["watched"] = [f for f in self.config["watched"] if f != name]
            save_config(self.config)
        self._build_menu()

    def _on_add_folder(self, _):
        win = rumps.Window(
            title="Add Folder",
            message="Enter the exact Mail folder name to watch:",
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

        if not verify_folder(name):
            rumps.alert(
                title="Folder not found",
                message=(
                    f'Mail couldn\'t find a folder named "{name}".\n\n'
                    "Check the exact spelling — it's case-sensitive.\n"
                    "Make sure Mail.app is running."
                ),
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

        counts = batch_unread_counts(watched)
        with self._lock:
            for name, count in counts.items():
                self.config["last_counts"].setdefault(name, count)
            save_config(self.config)

        while True:
            time.sleep(POLL_SECONDS)
            self._check_new_mail()

    def _check_new_mail(self):
        with self._lock:
            watched = list(self.config["watched"])
            last = dict(self.config["last_counts"])

        # Single osascript call for all folders instead of one per folder
        counts = batch_unread_counts(watched)
        changed = False

        for name in watched:
            count = counts.get(name)
            if count is None:
                continue  # Mail not running or folder gone

            prev = last.get(name, count)

            if count > prev:
                delta = count - prev
                sender, subject = try_get_newest_unread(name)

                if sender and subject:
                    body = f"{sender}: {subject}"
                else:
                    body = f"{delta} new message{'s' if delta > 1 else ''}"

                notify(f"📬 {name}", body)

            if last.get(name) != count:
                changed = True
            last[name] = count

        if changed:
            with self._lock:
                # Merge rather than replace — preserves counts for folders
                # added by another thread while we were polling
                self.config["last_counts"].update(last)
                save_config(self.config)


if __name__ == "__main__":
    OutlookNotify().run()
