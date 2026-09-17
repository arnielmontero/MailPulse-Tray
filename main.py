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
import pywintypes
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
class UnreadMessage:
    entry_id: str
    sender: str
    subject: str
    received: str  # pre-formatted display string


@dataclass
class AccountStatus:
    name: str
    unread: int = 0
    # EntryID/StoreID of unread items seen last cycle, used to detect
    # genuinely *new* unread mail (as opposed to a count that dropped).
    seen_entry_ids: set = field(default_factory=set)
    messages: list = field(default_factory=list)


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

    def _force_send_receive(self):
        """Actively trigger Outlook to fetch new mail from the server for
        every account, instead of only reading whatever Outlook already
        happens to have cached locally. Mirrors pressing Send/Receive All
        Folders (F9) in Outlook."""
        try:
            self.namespace.SendAndReceive(False)  # False = don't show dialog
        except Exception:
            # Fall back to iterating each sync group explicitly if the
            # simple call is unavailable for some Outlook configurations.
            try:
                for sync_object in self.namespace.SyncObjects:
                    try:
                        sync_object.Start()
                    except Exception:
                        continue
            except Exception:
                pass

    def _get_inbox_folder(self, store):
        """Return the Inbox folder for a store, matching what Outlook's
        folder-pane unread badge shows (Inbox only, not subfolders like
        Archive/Trash/spam/rule-filed folders)."""
        try:
            # olFolderInbox = 6
            return store.GetDefaultFolder(6)
        except Exception:
            return None

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

            self._force_send_receive()

            new_events = []
            snapshot = {}

            with self._lock:
                try:
                    stores = list(self.namespace.Stores)
                except pywintypes.com_error:
                    # The MAPI session can drop mid-sync (e.g. Outlook was
                    # restarted, or a transient server disconnect). Reconnect
                    # once and retry this poll rather than crashing the loop.
                    self.connect()
                    stores = list(self.namespace.Stores)

                for store in stores:
                    try:
                        store_name = store.DisplayName
                    except Exception:
                        continue

                    total_unread = 0
                    current_unread_ids = set()
                    messages = []

                    folder = self._get_inbox_folder(store)
                    if folder is not None:
                        try:
                            total_unread = folder.UnReadItemCount
                        except Exception:
                            total_unread = 0

                        # Identify the actual unread items so we can detect
                        # "new" mail, notify with sender/subject, and show
                        # a message list in the Account Breakdown popup.
                        try:
                            items = folder.Items
                            items = items.Restrict("[UnRead] = true")
                            items.Sort("[ReceivedTime]", True)
                            for item in items:
                                try:
                                    entry_id = item.EntryID
                                    current_unread_ids.add(entry_id)
                                    messages.append(UnreadMessage(
                                        entry_id=entry_id,
                                        sender=self._safe_sender_name(item),
                                        subject=getattr(item, "Subject", "(no subject)") or "(no subject)",
                                        received=self._format_received_time(item),
                                    ))
                                except Exception:
                                    continue
                        except Exception:
                            pass

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
                        messages=messages,
                    )
                    snapshot[store_name] = AccountStatus(
                        name=store_name,
                        unread=total_unread,
                        seen_entry_ids=current_unread_ids,
                        messages=messages,
                    )

            total = sum(status.unread for status in snapshot.values())
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

    @staticmethod
    def _format_received_time(item):
        try:
            received = item.ReceivedTime
            return received.strftime("%Y-%m-%d %H:%M")
        except Exception:
            return ""

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
        self._breakdown_window = None  # set while the breakdown popup is open

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
        # Only one breakdown window at a time; Tkinter must run its own
        # mainloop on a dedicated thread since pystray owns the main thread.
        if self._breakdown_window is not None:
            return
        threading.Thread(target=self._run_breakdown_window, daemon=True).start()

    def _run_breakdown_window(self):
        import tkinter as tk
        from tkinter import ttk

        root = tk.Tk()
        self._breakdown_window = root
        root.title(f"{APP_NAME} - Account Breakdown")
        root.attributes("-topmost", True)
        root.geometry("480x420")

        header = tk.Label(root, text="Unread by Account", font=("Segoe UI", 12, "bold"))
        header.pack(padx=16, pady=(12, 6), anchor="w")

        container = tk.Frame(root)
        container.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        total_label = tk.Label(root, text="", font=("Segoe UI", 10, "bold"))
        total_label.pack(anchor="w", padx=16, pady=(0, 4))

        def on_close():
            self._breakdown_window = None
            root.destroy()

        tk.Button(root, text="Close", command=on_close).pack(pady=(0, 12))
        root.protocol("WM_DELETE_WINDOW", on_close)

        def render():
            for child in scroll_frame.winfo_children():
                child.destroy()

            if not self._last_snapshot:
                tk.Label(scroll_frame, text="No data yet -- click 'Check Mail Now'.").pack(
                    anchor="w", pady=8
                )
                total_label.config(text="")
                return

            for name, status in sorted(self._last_snapshot.items()):
                account_label = tk.Label(
                    scroll_frame,
                    text=f"{name}  ({status.unread} unread)",
                    font=("Segoe UI", 10, "bold"),
                    anchor="w",
                )
                account_label.pack(fill="x", pady=(10, 2))

                if not status.messages:
                    tk.Label(scroll_frame, text="  (no unread mail)", fg="gray").pack(
                        anchor="w"
                    )
                for msg in status.messages:
                    row = tk.Frame(scroll_frame)
                    row.pack(fill="x", pady=1)
                    tk.Label(
                        row,
                        text=f"  {msg.received}",
                        font=("Segoe UI", 8),
                        fg="gray",
                        width=14,
                        anchor="w",
                    ).pack(side="left")
                    tk.Label(
                        row,
                        text=f"{msg.sender} — {msg.subject}",
                        font=("Segoe UI", 9),
                        anchor="w",
                        wraplength=320,
                        justify="left",
                    ).pack(side="left", fill="x", expand=True)

            total = sum(status.unread for status in self._last_snapshot.values())
            total_label.config(text=f"Total: {total}")

        def auto_refresh():
            if self._breakdown_window is None:
                return  # window was closed
            render()
            # Re-check slightly more often than the poll interval so the
            # popup picks up each new result promptly after it lands.
            root.after(2000, auto_refresh)

        auto_refresh()
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
