"""
MailPulse Tray
===============
A lightweight, zero-config Windows System Tray utility that tracks unread
emails across every account configured in a running Outlook Classic (desktop)
session, using MAPI/COM automation via pywin32.

Features
--------
- Zero-config: attaches to the already-running Outlook session and walks
  outlook.Stores to discover every configured account/mailbox automatically.
- System tray icon with a live badge showing total unread count across all
  accounts (pystray + Pillow, icon regenerated on every refresh).
- Tray menu: "Check Mail Now", "Account Breakdown", "Exit".
- Desktop toast notifications on new mail (windows-toasts, falling back to
  win10toast if unavailable) showing Account / Sender / Subject. Clicking a
  notification brings Outlook to the foreground and opens that exact email.
- Polls every 20 seconds. Because it always asks Outlook for the live
  UnReadItemCount, reading a message in Outlook is reflected automatically
  on the very next cycle -- no separate "sync" logic is needed.

This file is intentionally single-file and dependency-light so it can be
frozen into a single .exe with PyInstaller. See the accompanying README /
build instructions for packaging steps.
"""

import sys
import threading
import time
import traceback
from dataclasses import dataclass, field

import pythoncom
import win32com.client
import win32gui
import win32con
import win32process
import psutil  # only used to best-effort find/foreground the Outlook window
from PIL import Image, ImageDraw, ImageFont
import pystray

# --- Notifications -----------------------------------------------------
# Prefer windows-toasts (modern, supports click callbacks cleanly). Fall
# back to win10toast if it isn't installed.
_NOTIFY_BACKEND = None
try:
    from windows_toasts import WindowsToaster, Toast, ToastActivatedEventArgs
    _NOTIFY_BACKEND = "windows_toasts"
except ImportError:
    try:
        from win10toast import ToastNotifier
        _NOTIFY_BACKEND = "win10toast"
    except ImportError:
        _NOTIFY_BACKEND = None


POLL_INTERVAL_SECONDS = 20
APP_NAME = "MailPulse Tray"


# ------------------------------------------------------------------------
# Data model
# ------------------------------------------------------------------------

@dataclass
class AccountStatus:
    name: str
    unread: int = 0
    # EntryID/StoreID of unread items seen last cycle, used to detect
    # genuinely *new* unread mail (as opposed to a count that dropped).
    seen_entry_ids: set = field(default_factory=set)


# ------------------------------------------------------------------------
# Outlook / MAPI access
# ------------------------------------------------------------------------

class OutlookMonitor:
    """Wraps COM access to a running Outlook Classic session."""

    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.accounts = {}  # store name -> AccountStatus
        self._lock = threading.Lock()

    def connect(self):
        """Attach to the running Outlook instance via MAPI."""
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")

    def _iter_inbox_like_folders(self, store):
        """Yield the Inbox folder for a store (root of unread tracking).

        Outlook stores expose olFolderInbox per-store via GetDefaultFolder
        on a Namespace bound to that store's root, but the simplest and
        most reliable zero-config approach is to walk the store's root
        folder tree and sum UnReadItemCount across all mail folders,
        which also captures unread mail that rules moved out of Inbox.
        """
        root = store.GetRootFolder()
        yield from self._walk_folders(root)

    def _walk_folders(self, folder):
        try:
            yield folder
            for sub in folder.Folders:
                yield from self._walk_folders(sub)
        except Exception:
            # Some folders (e.g. public folders, permissions-restricted)
            # may refuse enumeration -- skip them rather than crash.
            return

    def poll(self):
        """Refresh unread counts for every account. Returns:
        (total_unread, dict[name -> unread], list[new_mail_events])

        new_mail_events is a list of dicts: {account, sender, subject, entry_id}
        for messages that are newly unread since the previous poll.
        """
        pythoncom.CoInitialize()
        try:
            if self.outlook is None:
                self.connect()

            new_events = []
            snapshot = {}

            with self._lock:
                for store in self.namespace.Stores:
                    try:
                        store_name = store.DisplayName
                    except Exception:
                        continue

                    total_unread = 0
                    current_unread_ids = set()

                    for folder in self._iter_inbox_like_folders(store):
                        try:
                            unread_count = folder.UnReadItemCount
                        except Exception:
                            continue
                        if unread_count <= 0:
                            continue
                        total_unread += unread_count

                        # Identify the actual unread items so we can detect
                        # "new" mail and notify with sender/subject.
                        try:
                            items = folder.Items
                            items = items.Restrict("[UnRead] = true")
                            for item in items:
                                try:
                                    entry_id = item.EntryID
                                    current_unread_ids.add(entry_id)
                                except Exception:
                                    continue
                        except Exception:
                            continue

                    prior = self.accounts.get(store_name)
                    prior_ids = prior.seen_entry_ids if prior else set()

                    # Only fire notifications after we have an established
                    # baseline (skip the very first poll to avoid a storm
                    # of notifications for pre-existing unread mail).
                    if prior is not None:
                        newly_unread_ids = current_unread_ids - prior_ids
                        for entry_id in newly_unread_ids:
                            try:
                                item = self.namespace.GetItemFromID(entry_id)
                                new_events.append({
                                    "account": store_name,
                                    "sender": self._safe_sender_name(item),
                                    "subject": getattr(item, "Subject", "(no subject)"),
                                    "entry_id": entry_id,
                                    "store_id": store.StoreID,
                                })
                            except Exception:
                                continue

                    self.accounts[store_name] = AccountStatus(
                        name=store_name,
                        unread=total_unread,
                        seen_entry_ids=current_unread_ids,
                    )
                    snapshot[store_name] = total_unread

            total = sum(snapshot.values())
            return total, snapshot, new_events
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _safe_sender_name(item):
        for attr in ("SenderName", "SenderEmailAddress"):
            try:
                val = getattr(item, attr)
                if val:
                    return val
            except Exception:
                continue
        return "Unknown Sender"

    def open_item_and_foreground(self, entry_id):
        """Open the given item in Outlook and bring Outlook to front."""
        pythoncom.CoInitialize()
        try:
            item = self.namespace.GetItemFromID(entry_id)
            item.Display()  # opens the Outlook item window
            self._foreground_outlook()
        except Exception:
            traceback.print_exc()
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _foreground_outlook():
        """Best-effort: bring the Outlook main/item window to the foreground."""
        def enum_handler(hwnd, results):
            if not win32gui.IsWindowVisible(hwnd):
                return
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                proc = psutil.Process(pid)
                if proc.name().lower() == "outlook.exe":
                    results.append(hwnd)
            except Exception:
                pass

        results = []
        win32gui.EnumWindows(enum_handler, results)
        for hwnd in results:
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                continue


# ------------------------------------------------------------------------
# Tray icon rendering
# ------------------------------------------------------------------------

def make_badge_icon(count: int) -> Image.Image:
    """Render a simple envelope-style icon with a numeric unread badge."""
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Simple envelope base
    envelope_color = (33, 115, 199, 255) if count > 0 else (120, 120, 120, 255)
    draw.rounded_rectangle([4, 14, 60, 50], radius=6, fill=envelope_color)
    draw.polygon([(4, 14), (32, 36), (60, 14)], fill=(255, 255, 255, 60))

    if count > 0:
        # Red badge circle in the top-right corner
        badge_radius = 20
        cx, cy = size - badge_radius + 4, badge_radius - 4
        draw.ellipse(
            [cx - badge_radius, cy - badge_radius, cx + badge_radius, cy + badge_radius],
            fill=(220, 40, 40, 255),
            outline=(255, 255, 255, 255),
            width=2,
        )
        text = str(count) if count < 100 else "99+"
        try:
            font = ImageFont.truetype("segoeuib.ttf", 20 if len(text) < 3 else 15)
        except Exception:
            font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text((cx - tw / 2 - bbox[0], cy - th / 2 - bbox[1]), text, fill="white", font=font)

    return img


# ------------------------------------------------------------------------
# Notifications
# ------------------------------------------------------------------------

class Notifier:
    """Thin wrapper that normalizes windows-toasts / win10toast differences."""

    def __init__(self, on_click):
        self.on_click = on_click
        self._toaster = None
        if _NOTIFY_BACKEND == "windows_toasts":
            self._toaster = WindowsToaster(APP_NAME)
        elif _NOTIFY_BACKEND == "win10toast":
            self._toaster = ToastNotifier()

    def notify(self, account, sender, subject, entry_id):
        title = f"{account}: New Mail"
        message = f"From: {sender}\n{subject}"

        if _NOTIFY_BACKEND == "windows_toasts":
            toast = Toast()
            toast.text_fields = [title, message]

            def _activated(_args: ToastActivatedEventArgs):
                self.on_click(entry_id)

            toast.on_activated = _activated
            self._toaster.show_toast(toast)
        elif _NOTIFY_BACKEND == "win10toast":
            # win10toast has no reliable click callback; notify only.
            # Run in a thread since it can block briefly.
            threading.Thread(
                target=self._toaster.show_toast,
                kwargs=dict(title=title, msg=message, duration=8, threaded=True),
                daemon=True,
            ).start()
        else:
            print(f"[Notification] {title} -- {message}")


# ------------------------------------------------------------------------
# Application
# ------------------------------------------------------------------------

class MailPulseApp:
    def __init__(self):
        self.monitor = OutlookMonitor()
        self.notifier = Notifier(on_click=self._handle_notification_click)
        self.icon = pystray.Icon(APP_NAME)
        self.icon.icon = make_badge_icon(0)
        self.icon.title = f"{APP_NAME} - starting..."
        self.icon.menu = pystray.Menu(
            pystray.MenuItem("Check Mail Now", self._on_check_now),
            pystray.MenuItem("Account Breakdown", self._on_account_breakdown),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", self._on_exit),
        )
        self._stop_event = threading.Event()
        self._last_snapshot = {}

    # -- background polling loop -----------------------------------------

    def _poll_loop(self):
        while not self._stop_event.is_set():
            self._do_poll()
            self._stop_event.wait(POLL_INTERVAL_SECONDS)

    def _do_poll(self):
        try:
            total, snapshot, new_events = self.monitor.poll()
        except Exception:
            traceback.print_exc()
            self.icon.title = f"{APP_NAME} - Outlook not available"
            return

        self._last_snapshot = snapshot
        self.icon.icon = make_badge_icon(total)
        self.icon.title = f"{APP_NAME} - {total} unread"

        for event in new_events:
            self.notifier.notify(
                account=event["account"],
                sender=event["sender"],
                subject=event["subject"],
                entry_id=event["entry_id"],
            )

    # -- menu handlers -----------------------------------------------------

    def _on_check_now(self, icon, item):
        threading.Thread(target=self._do_poll, daemon=True).start()

    def _on_account_breakdown(self, icon, item):
        threading.Thread(target=self._show_breakdown_window, daemon=True).start()

    def _show_breakdown_window(self):
        # Lightweight Tkinter popup so we avoid pulling in a heavier GUI dep.
        import tkinter as tk

        root = tk.Tk()
        root.title(f"{APP_NAME} - Account Breakdown")
        root.attributes("-topmost", True)
        root.resizable(False, False)

        tk.Label(root, text="Unread by Account", font=("Segoe UI", 12, "bold")).pack(
            padx=16, pady=(12, 6)
        )

        if not self._last_snapshot:
            tk.Label(root, text="No data yet -- click 'Check Mail Now'.").pack(padx=16, pady=8)
        else:
            for name, count in sorted(self._last_snapshot.items()):
                tk.Label(root, text=f"{name}: {count}", font=("Segoe UI", 10)).pack(
                    anchor="w", padx=16
                )
            total = sum(self._last_snapshot.values())
            tk.Label(root, text=f"\nTotal: {total}", font=("Segoe UI", 10, "bold")).pack(
                anchor="w", padx=16, pady=(4, 12)
            )

        tk.Button(root, text="Close", command=root.destroy).pack(pady=(0, 12))
        root.mainloop()

    def _on_exit(self, icon, item):
        self._stop_event.set()
        icon.stop()

    def _handle_notification_click(self, entry_id):
        threading.Thread(
            target=self.monitor.open_item_and_foreground, args=(entry_id,), daemon=True
        ).start()

    # -- entrypoint ----------------------------------------------------------

    def run(self):
        threading.Thread(target=self._poll_loop, daemon=True).start()
        self.icon.run()


def main():
    if sys.platform != "win32":
        print("MailPulse Tray requires Windows (Outlook COM automation).")
        sys.exit(1)

    app = MailPulseApp()
    app.run()


if __name__ == "__main__":
    main()
