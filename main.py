"""
MailPulse Tray
===============
A lightweight Windows System Tray utility that tracks unread emails across
multiple cPanel (or any IMAP) accounts by connecting DIRECTLY over IMAP,
independent of whether Outlook Classic is running. Outlook Classic (via
win32com/MAPI) is used only for one thing: deep-linking -- opening the
exact email in Outlook when you click a notification.

Architecture
------------
- IMAP (imaplib, SSL port 993) is the source of truth for unread counts.
  This means the tray badge updates even when Outlook is completely closed.
- Every poll cycle (15-30s) each configured account is checked for UNSEEN
  messages. Because Outlook marks a message SEEN on the IMAP server the
  moment you read it there, the next poll automatically sees the lower
  UNSEEN count -- no separate "sync" logic needed, IMAP already reflects
  Outlook's read state.
- New UNSEEN messages (not seen on the previous poll) trigger a toast
  notification with Account / Sender / Subject. Clicking the toast
  launches/focuses Outlook Classic and opens that exact email by matching
  its Message-ID header against Outlook's PR_INTERNET_MESSAGE_ID property.
- Account credentials live in a local config.json (plaintext, per user
  request) next to the executable/script.

This file is single-file and dependency-light so it can be frozen into a
single .exe with PyInstaller. See README.md for packaging steps.
"""

import imaplib
import email
import email.utils
import json
import os
import sys
import threading
import traceback
from dataclasses import dataclass, field

from PIL import Image, ImageDraw, ImageFont
import pystray

# --- Outlook COM (deep-linking only; imported lazily/defensively so the
# IMAP fetching path works even if pywin32/Outlook isn't available) -------
try:
    import pythoncom
    import win32com.client
    import win32gui
    import win32con
    import win32process
    import psutil
    _OUTLOOK_COM_AVAILABLE = True
except ImportError:
    _OUTLOOK_COM_AVAILABLE = False

# --- Notifications -----------------------------------------------------
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


APP_NAME = "MailPulse Tray"
POLL_INTERVAL_SECONDS = 20
CONFIG_FILENAME = "config.json"


def _base_dir():
    """Directory the .exe (or script) lives in, so config.json sits next
    to it whether running from source or as a frozen PyInstaller build."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(_base_dir(), CONFIG_FILENAME)


# ------------------------------------------------------------------------
# Config
# ------------------------------------------------------------------------

@dataclass
class AccountConfig:
    name: str
    imap_server: str
    imap_port: int
    email: str
    password: str
    mailbox: str = "INBOX"
    use_ssl: bool = True


def load_accounts():
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(
            f"config.json not found at {CONFIG_PATH}. See config.example.json "
            "for the expected format."
        )
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        raw = json.load(f)

    accounts = []
    for entry in raw.get("accounts", []):
        accounts.append(AccountConfig(
            name=entry.get("name") or entry["email"],
            imap_server=entry["imap_server"],
            imap_port=int(entry.get("imap_port", 993)),
            email=entry["email"],
            password=entry["password"],
            mailbox=entry.get("mailbox", "INBOX"),
            use_ssl=bool(entry.get("use_ssl", True)),
        ))
    return accounts


# ------------------------------------------------------------------------
# Data model
# ------------------------------------------------------------------------

@dataclass
class UnreadMessage:
    uid: str
    message_id: str
    sender: str
    subject: str
    received: str


@dataclass
class AccountStatus:
    name: str
    unread: int = 0
    seen_uids: set = field(default_factory=set)
    messages: list = field(default_factory=list)
    error: str = ""


# ------------------------------------------------------------------------
# IMAP fetching -- independent of Outlook
# ------------------------------------------------------------------------

class ImapMonitor:
    """Polls one or more IMAP accounts directly, independent of Outlook."""

    def __init__(self, accounts):
        self.accounts = accounts
        self.status = {}  # name -> AccountStatus
        self._lock = threading.Lock()

    def poll(self):
        """Check every account for unread mail. Returns:
        (total_unread, dict[name -> AccountStatus], list[new_mail_events])

        new_mail_events: list of dicts with account/sender/subject/message_id
        for messages newly seen as unread since the previous poll.
        """
        new_events = []
        snapshot = {}

        with self._lock:
            for account in self.accounts:
                status = self._poll_account(account)
                prior = self.status.get(account.name)
                prior_uids = prior.seen_uids if prior else set()

                if prior is not None:
                    new_uids = status.seen_uids - prior_uids
                    for msg in status.messages:
                        if msg.uid in new_uids:
                            new_events.append({
                                "account": account.name,
                                "sender": msg.sender,
                                "subject": msg.subject,
                                "message_id": msg.message_id,
                            })

                self.status[account.name] = status
                snapshot[account.name] = status

        total = sum(s.unread for s in snapshot.values())
        return total, snapshot, new_events

    def _poll_account(self, account: AccountConfig) -> AccountStatus:
        try:
            if account.use_ssl:
                conn = imaplib.IMAP4_SSL(account.imap_server, account.imap_port)
            else:
                conn = imaplib.IMAP4(account.imap_server, account.imap_port)
            try:
                conn.login(account.email, account.password)
                conn.select(account.mailbox, readonly=True)

                typ, data = conn.search(None, "UNSEEN")
                if typ != "OK":
                    return AccountStatus(name=account.name, error="IMAP search failed")

                uids = data[0].split()
                messages = []
                seen_uids = set()

                # Fetch headers only (fast, no body download) for each
                # unseen message, newest first.
                for uid in reversed(uids):
                    uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
                    seen_uids.add(uid_str)
                    try:
                        typ, msg_data = conn.fetch(
                            uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE MESSAGE-ID)])"
                        )
                        if typ != "OK" or not msg_data or not msg_data[0]:
                            continue
                        header_bytes = msg_data[0][1]
                        parsed = email.message_from_bytes(header_bytes)

                        sender = email.utils.parseaddr(parsed.get("From", ""))[0] \
                            or email.utils.parseaddr(parsed.get("From", ""))[1] \
                            or "Unknown Sender"
                        subject = parsed.get("Subject", "(no subject)")
                        message_id = parsed.get("Message-ID", "").strip()
                        date_hdr = parsed.get("Date", "")
                        try:
                            dt = email.utils.parsedate_to_datetime(date_hdr)
                            received = dt.strftime("%Y-%m-%d %H:%M")
                        except Exception:
                            received = ""

                        messages.append(UnreadMessage(
                            uid=uid_str,
                            message_id=message_id,
                            sender=sender,
                            subject=subject,
                            received=received,
                        ))
                    except Exception:
                        continue

                return AccountStatus(
                    name=account.name,
                    unread=len(uids),
                    seen_uids=seen_uids,
                    messages=messages,
                )
            finally:
                try:
                    conn.logout()
                except Exception:
                    pass
        except Exception as exc:
            return AccountStatus(name=account.name, error=str(exc))


# ------------------------------------------------------------------------
# Outlook COM deep-linking (click-to-open only)
# ------------------------------------------------------------------------

class OutlookLauncher:
    """Used only to open a specific email in Outlook Classic when a
    notification is clicked. Never used to check/send/receive mail."""

    def __init__(self):
        self.outlook = None
        self.namespace = None

    def _connect(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")

    def open_by_message_id(self, message_id: str, subject_fallback: str = ""):
        """Launch/focus Outlook and Display() the item matching the given
        Message-ID header. Falls back to a Subject search if the
        Message-ID can't be matched (e.g. Outlook hasn't synced it yet)."""
        if not _OUTLOOK_COM_AVAILABLE:
            return
        pythoncom.CoInitialize()
        try:
            if self.outlook is None:
                self._connect()

            item = self._find_by_message_id(message_id) if message_id else None
            if item is None and subject_fallback:
                item = self._find_by_subject(subject_fallback)

            if item is not None:
                item.Display()

            self._foreground_outlook()
        except Exception:
            traceback.print_exc()
        finally:
            pythoncom.CoUninitialize()

    def _find_by_message_id(self, message_id: str):
        PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
        try:
            for store in self.namespace.Stores:
                inbox = store.GetDefaultFolder(6)  # olFolderInbox
                items = inbox.Items
                for item in items:
                    try:
                        mapi_id = item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
                    except Exception:
                        continue
                    if mapi_id and mapi_id.strip() == message_id.strip():
                        return item
        except Exception:
            pass
        return None

    def _find_by_subject(self, subject: str):
        try:
            for store in self.namespace.Stores:
                inbox = store.GetDefaultFolder(6)
                items = inbox.Items
                items.Sort("[ReceivedTime]", True)
                restricted = items.Restrict(
                    "[Subject] = \"" + subject.replace('"', '') + "\""
                )
                for item in restricted:
                    return item
        except Exception:
            pass
        return None

    @staticmethod
    def _foreground_outlook():
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
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    envelope_color = (33, 115, 199, 255) if count > 0 else (120, 120, 120, 255)
    draw.rounded_rectangle([4, 14, 60, 50], radius=6, fill=envelope_color)
    draw.polygon([(4, 14), (32, 36), (60, 14)], fill=(255, 255, 255, 60))

    if count > 0:
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
    def __init__(self, on_click):
        self.on_click = on_click
        self._toaster = None
        if _NOTIFY_BACKEND == "windows_toasts":
            self._toaster = WindowsToaster(APP_NAME)
        elif _NOTIFY_BACKEND == "win10toast":
            self._toaster = ToastNotifier()

    def notify(self, account, sender, subject, message_id):
        title = f"{account}: New Mail"
        message = f"From: {sender}\n{subject}"

        if _NOTIFY_BACKEND == "windows_toasts":
            toast = Toast()
            toast.text_fields = [title, message]

            def _activated(_args: ToastActivatedEventArgs):
                self.on_click(message_id, subject)

            toast.on_activated = _activated
            self._toaster.show_toast(toast)
        elif _NOTIFY_BACKEND == "win10toast":
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
    def __init__(self, accounts):
        self.monitor = ImapMonitor(accounts)
        self.outlook_launcher = OutlookLauncher() if _OUTLOOK_COM_AVAILABLE else None
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
        self._breakdown_window = None

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
            self.icon.title = f"{APP_NAME} - error checking mail"
            return

        self._last_snapshot = snapshot
        self.icon.icon = make_badge_icon(total)

        errors = [s.name for s in snapshot.values() if s.error]
        if errors:
            self.icon.title = f"{APP_NAME} - {total} unread ({len(errors)} account(s) failed)"
        else:
            self.icon.title = f"{APP_NAME} - {total} unread"

        for event in new_events:
            self.notifier.notify(
                account=event["account"],
                sender=event["sender"],
                subject=event["subject"],
                message_id=event["message_id"],
            )

    # -- menu handlers -----------------------------------------------------

    def _on_check_now(self, icon, item):
        threading.Thread(target=self._do_poll, daemon=True).start()

    def _on_account_breakdown(self, icon, item):
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

        tk.Label(root, text="Unread by Account", font=("Segoe UI", 12, "bold")).pack(
            padx=16, pady=(12, 6), anchor="w"
        )

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
                label_text = f"{name}  ({status.unread} unread)"
                if status.error:
                    label_text += "  [connection error]"
                tk.Label(
                    scroll_frame, text=label_text, font=("Segoe UI", 10, "bold"), anchor="w"
                ).pack(fill="x", pady=(10, 2))

                if status.error:
                    tk.Label(scroll_frame, text=f"  {status.error}", fg="red").pack(anchor="w")
                elif not status.messages:
                    tk.Label(scroll_frame, text="  (no unread mail)", fg="gray").pack(anchor="w")
                for msg in status.messages:
                    row = tk.Frame(scroll_frame)
                    row.pack(fill="x", pady=1)
                    tk.Label(
                        row, text=f"  {msg.received}", font=("Segoe UI", 8), fg="gray",
                        width=14, anchor="w",
                    ).pack(side="left")
                    tk.Label(
                        row, text=f"{msg.sender} — {msg.subject}", font=("Segoe UI", 9),
                        anchor="w", wraplength=320, justify="left",
                    ).pack(side="left", fill="x", expand=True)

            total = sum(s.unread for s in self._last_snapshot.values())
            total_label.config(text=f"Total: {total}")

        def auto_refresh():
            if self._breakdown_window is None:
                return
            render()
            root.after(2000, auto_refresh)

        auto_refresh()
        root.mainloop()

    def _on_exit(self, icon, item):
        self._stop_event.set()
        icon.stop()

    def _handle_notification_click(self, message_id, subject):
        if self.outlook_launcher is None:
            return
        threading.Thread(
            target=self.outlook_launcher.open_by_message_id,
            args=(message_id, subject),
            daemon=True,
        ).start()

    # -- entrypoint ----------------------------------------------------------

    def run(self):
        threading.Thread(target=self._poll_loop, daemon=True).start()
        self.icon.run()


def main():
    if sys.platform != "win32":
        print("MailPulse Tray requires Windows.")
        sys.exit(1)

    try:
        accounts = load_accounts()
    except Exception as exc:
        print(f"Failed to load {CONFIG_PATH}: {exc}")
        sys.exit(1)

    if not accounts:
        print("No accounts configured in config.json.")
        sys.exit(1)

    app = MailPulseApp(accounts)
    app.run()


if __name__ == "__main__":
    main()
