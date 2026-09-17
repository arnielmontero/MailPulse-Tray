# MailPulse Tray

A zero-config Windows system tray utility that tracks unread mail across all
accounts logged into Outlook Classic (desktop), via MAPI/COM automation.

## Requirements

- Windows 10/11
- Outlook Classic (desktop) installed and configured with one or more
  accounts (cPanel/IMAP/Exchange, etc.) — it must be running (or launchable)
  for MAPI automation to work.
- Python 3.9+ (only needed to run from source / build the .exe; end users
  just run the compiled `.exe`).

## 1. Install dependencies

```bash
pip install pywin32 pystray Pillow psutil windows-toasts pyinstaller
```

> `windows-toasts` is preferred for click-to-open notification support. If
> it fails to install on your system, you can substitute `win10toast`
> instead — the script auto-detects whichever is present, but note
> `win10toast` does not support click callbacks reliably.

Or simply:

```bash
pip install -r requirements.txt
```

## 2. Run from source (optional, for testing)

```bash
python main.py
```

Outlook must already be open (or will be launched by COM automation on
first call). The tray icon will appear in the system tray within ~20
seconds showing the current unread badge.

## 3. Build the standalone .exe

From the project directory, run:

```bash
pyinstaller --noconsole --onefile --name MailPulseTray --icon=NONE main.py
```

- `--noconsole` — suppresses the terminal/console window (this is a
  background tray app).
- `--onefile` — bundles everything into a single portable `.exe`.
- `--icon=NONE` — omit if you have a custom `.ico` file; otherwise replace
  with `--icon=app_icon.ico` to brand the built executable and its taskbar
  entry. (The in-tray badge icon is generated dynamically at runtime
  regardless of this flag.)

The compiled executable will be created at:

```
dist\MailPulseTray.exe
```

### Notes on packaging pywin32

PyInstaller usually bundles `pywin32` correctly, but if you hit COM errors
in the frozen `.exe` (e.g. `win32com.client.gencache` issues), add:

```bash
pyinstaller --noconsole --onefile --name MailPulseTray ^
  --hidden-import=win32timezone ^
  --collect-submodules win32com ^
  main.py
```

## 4. Run automatically at Windows startup

1. Press `Win + R`, type `shell:startup`, press Enter. This opens your
   personal Windows Startup folder.
2. Copy `dist\MailPulseTray.exe`, right-click inside the Startup folder,
   and choose **Paste shortcut** (recommended over pasting the raw .exe,
   so you can rename/move the original later without breaking startup).
3. Log off and back on (or reboot) to confirm it launches silently in the
   tray.

To stop it from auto-running, delete the shortcut from the Startup folder.

## How it works

- On each poll cycle (every 20 seconds), the app connects to the running
  Outlook session via `win32com.client.Dispatch("Outlook.Application")`
  and walks `namespace.Stores` — this automatically picks up every account
  configured in Outlook (cPanel/IMAP, POP, Exchange, etc.) with zero
  manual configuration.
- For each store, it recursively walks all folders and sums
  `folder.UnReadItemCount`, so mail filed into subfolders by rules is
  still counted.
- It tracks the set of unread item EntryIDs per account between polls to
  detect genuinely *new* unread mail (vs. a count that simply changed) and
  fires a toast notification with the account name, sender, and subject
  only for those new items.
- Because the badge is always recomputed from Outlook's live
  `UnReadItemCount`, marking an email as read inside Outlook is reflected
  automatically on the very next poll — no separate sync logic required.
- Clicking a notification calls `item.Display()` on that exact mail item
  and brings the Outlook window to the foreground.
