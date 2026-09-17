# MailPulse Tray

A Windows system tray utility that tracks unread mail across multiple
cPanel (or any IMAP) accounts by connecting **directly over IMAP** —
independent of whether Outlook Classic is running. Outlook Classic is used
only for one thing: opening the exact email when you click a notification.

## Architecture

- **IMAP is the source of truth.** Every 20 seconds the app connects
  directly to each configured account over IMAP SSL (port 993) and checks
  `UNSEEN` message counts. This works even if Outlook is fully closed.
- **Read-state sync is automatic.** When you read an email in Outlook,
  Outlook marks it `SEEN` on the IMAP server. The next IMAP poll sees the
  lower `UNSEEN` count automatically — no separate sync logic needed.
- **Outlook COM is only used for deep-linking.** When you click a toast
  notification, the app launches/focuses Outlook Classic and opens that
  exact email by matching its `Message-ID` header against Outlook's
  `PR_INTERNET_MESSAGE_ID` property (falls back to a Subject search if the
  message hasn't synced into Outlook yet).

## Requirements

- Windows 10/11
- Python 3.9+ (only needed to run from source / build the .exe)
- cPanel (or any IMAP-accessible) email account credentials
- Outlook Classic installed, for the click-to-open deep-link feature only
  (not required for mail checking/notifications to work)

## 1. Install dependencies

```bash
pip install -r requirements.txt
```

This installs `pywin32`, `pystray`, `Pillow`, `psutil`, `windows-toasts`,
and `pyinstaller`. `pywin32`/`psutil` are only used for the Outlook
deep-link feature — IMAP fetching itself uses only the Python standard
library (`imaplib`, `email`).

## 2. Configure accounts

Copy `config.example.json` to `config.json` (same folder as `main.py`, or
next to the compiled `.exe`) and fill in your real cPanel IMAP credentials:

```json
{
  "accounts": [
    {
      "name": "Support Inbox",
      "imap_server": "mail.example-cpanel-host.com",
      "imap_port": 993,
      "use_ssl": true,
      "email": "support@example.com",
      "password": "your-password-here",
      "mailbox": "INBOX"
    }
  ]
}
```

Add one object per account. `mailbox` defaults to `INBOX` if omitted.

> **Security note:** passwords are stored in plaintext in `config.json` by
> design (per project requirements) for simplicity. Keep this file private
> — it is already excluded via `.gitignore` and must never be committed or
> shared. If you'd prefer OS-level credential storage instead, this can be
> swapped for Windows Credential Manager (via the `keyring` package) later.

## 3. Run from source (optional, for testing)

```bash
python main.py
```

The tray icon appears within a few seconds showing the unread badge across
all configured accounts, refreshed every 20 seconds.

## 4. Build the standalone .exe

```bash
pyinstaller --noconsole --onefile --name MailPulseTray ^
  --hidden-import=win32timezone ^
  --hidden-import=win32com ^
  --hidden-import=win32com.client ^
  --hidden-import=pythoncom ^
  --hidden-import=pywintypes ^
  --collect-submodules win32com ^
  main.py
```

Or just run `build.bat` from the project directory.

- `--noconsole` — suppresses the terminal/console window (background tray
  app).
- `--onefile` — bundles everything into a single portable `.exe`.
- The `--hidden-import` / `--collect-submodules` flags are required
  because `pythoncom` / `pywintypes` are compiled DLL-backed modules that
  PyInstaller's static analysis misses by default.

The compiled executable is created at:

```
dist\MailPulseTray.exe
```

**Important:** copy your real `config.json` into the `dist\` folder next
to `MailPulseTray.exe` — the compiled app looks for `config.json` in the
same directory it runs from.

## 5. Run automatically at Windows startup

1. Press `Win + R`, type `shell:startup`, press Enter.
2. Copy `dist\MailPulseTray.exe` and `dist\config.json` to a permanent
   folder (e.g. `C:\Tools\MailPulseTray\`).
3. In the Startup folder, right-click and **Paste shortcut** pointing at
   the `.exe` in that permanent folder.
4. Log off and back on (or reboot) to confirm it launches silently in the
   tray.

To stop it from auto-running, delete the shortcut from the Startup folder.

## How it works

- `ImapMonitor.poll()` connects to each account with `imaplib.IMAP4_SSL`,
  logs in, selects the mailbox read-only, and runs `SEARCH UNSEEN`. Only
  headers (`From`, `Subject`, `Date`, `Message-ID`) are fetched — not full
  message bodies — so polling stays fast even with a busy inbox.
- New-mail detection compares the set of UNSEEN UIDs against the previous
  poll's set; only genuinely new UIDs trigger a toast notification
  (the very first poll never fires notifications, to avoid a startup
  storm for pre-existing unread mail).
- Clicking a notification calls `OutlookLauncher.open_by_message_id()`,
  which launches Outlook via `win32com.client.Dispatch` if it isn't
  running, searches Inbox items across all Outlook stores for a matching
  `PR_INTERNET_MESSAGE_ID`, and calls `item.Display()` on it — falling
  back to a Subject-based search if no Message-ID match is found.
- Because IMAP is the only thing polled for counts, the tray badge, the
  Account Breakdown popup, and notifications all work whether or not
  Outlook is open. Outlook is only launched on-demand, only when you
  click a notification.
