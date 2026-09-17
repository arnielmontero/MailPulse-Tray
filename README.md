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
`keyring`, and `pyinstaller`. `pywin32`/`psutil` are only used for the
Outlook deep-link feature — IMAP fetching itself uses only the Python
standard library (`imaplib`, `email`).

## 2. Configure accounts (via the Settings window — no file editing)

Accounts are added entirely through the app's UI:

1. Run `python main.py` (or the compiled `.exe`). On first run with no
   accounts configured, the **Settings** window opens automatically.
2. Click **Add Account** and fill in: Account Name, Email, IMAP Server,
   IMAP Port (993 for SSL), Mailbox (defaults to `INBOX`), Use SSL, and
   Password.
3. Click **Save**. You can add as many accounts as you like, and later
   reopen Settings any time from the tray menu (right-click the tray
   icon → **Settings**) to add, edit, or delete accounts.

**Where credentials actually live:**
- Server/email/port/mailbox metadata is stored in `accounts.json`, next to
  `main.py` or the compiled `.exe`. No passwords are written to this file.
- Each account's password is stored securely in **Windows Credential
  Manager**, via the `keyring` package, keyed by an internal account ID.
  You can inspect these entries yourself in Windows' *Credential Manager*
  control panel under "Generic Credentials" (look for entries prefixed
  `MailPulseTray`).

### Notification timing

The same Settings window has a **Notifications** section:

- **Check for new email every:** how often the app polls IMAP (also
  controls how quickly the tray badge and new-mail toasts update).
  Options: 15 sec, 20 sec, 30 sec, 1 min, 2 min, 5 min.
- **Unread follow-up reminder:** an optional recurring toast reminding you
  of mail still sitting unread ("Reminder: You still have N unread emails
  waiting in your inbox"), fired only while unread mail remains. Options:
  Never, every 15 min, every 30 min, every 1 hr. Clicking this toast opens
  the same Unread Mail list window as a tray left-click.

Click **Save Notification Settings** to apply and persist these — they're
stored in `settings.json` (next to `accounts.json`) and restored
automatically on the next launch.

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
  --hidden-import=keyring.backends.Windows ^
  --collect-submodules win32com ^
  --collect-submodules keyring ^
  main.py
```

Or just run `build.bat` from the project directory.

- `--noconsole` — suppresses the terminal/console window (background tray
  app).
- `--onefile` — bundles everything into a single portable `.exe`.
- The `--hidden-import` / `--collect-submodules` flags are required
  because `pythoncom` / `pywintypes` are compiled DLL-backed modules, and
  `keyring`'s Windows backend is resolved via entry points at runtime —
  both are missed by PyInstaller's static analysis by default.

The compiled executable is created at:

```
dist\MailPulseTray.exe
```

**Note:** `accounts.json` is created automatically the first time you save
an account through the Settings window — there is nothing to copy in
manually. If you move the `.exe` to a new machine/folder, you'll need to
re-add accounts there (Credential Manager entries and `accounts.json` are
both local to the machine/user profile they were created on).

## 5. Run automatically at Windows startup

1. Press `Win + R`, type `shell:startup`, press Enter.
2. Copy `dist\MailPulseTray.exe` to a permanent folder (e.g.
   `C:\Tools\MailPulseTray\`) — `accounts.json` will be created there
   automatically the first time you add an account.
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
