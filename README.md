# Sweep

A small Windows desktop app that finds the newsletters in your **Gmail**, **Outlook / Hotmail**, **Yahoo**, **iCloud**, **Libero**, **Virgilio** and **Tiscali** accounts and unsubscribes you from them, a few clicks for dozens of senders.

![Sweep listing newsletter senders, with two selected for unsubscribing](docs/screenshot.png)

## What it does

- **Scans every folder**, including Spam and Trash, over the last 30 days, 90 days, a year, or all your mail.
- **Groups mail by sender**, with how many emails each one sent and when the latest arrived.
- **Shows Gmail and Libero categories** (Promotions / Offerte, Social, Ordini, Prenotazioni…), with filters by account and category and sorting by category, volume or date.
- **Unsubscribes in the best way each sender allows**, falling back to the next one if a way fails:
  1. *One click*: the sender's official one-click unsubscribe (RFC 8058), done silently.
  2. *Sends an email*: an unsubscribe email from your own account.
  3. *Opens the website*: the sender's unsubscribe page, in your browser.
  4. *Opens the link in the email*: for senders who only put the link at the bottom of their emails.
- **Ignore** senders you want to keep, **dark mode**, and a memory of what you already unsubscribed from.

Sweep never deletes, moves or marks your emails. It only reads them.

## Install

1. Download **`Sweep.exe`** from the [latest release](../../releases/latest).
2. Double-click it. Nothing to install.
   - Windows may show "Windows protected your PC", because the app isn't signed with a paid certificate. Click **More info → Run anyway**.
   - Sweep uses Microsoft Edge WebView2, which Windows 10 and 11 already include.

## Add your accounts

**Gmail** needs an *App Password*, not your normal password:

1. Turn on [2-Step Verification](https://myaccount.google.com/signinoptions/two-step-verification) for your Google account.
2. Create an App Password at [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) and paste it into Sweep.
3. IMAP is on by default; if the login fails, check Gmail settings → *Forwarding and POP/IMAP*.

**Outlook / Hotmail / Live** don't accept passwords from email programs any more: adding the account opens Microsoft's sign-in page in your browser. Sweep keeps the sign-in in `%APPDATA%\Sweep\microsoft_tokens.bin`, encrypted for your Windows user.

**Yahoo** needs an App Password (Account security → Generate app password). **iCloud** needs an app-specific password (account.apple.com → Sign-In and Security). **Virgilio** needs its *Password Sicura*, created in the webmail settings.

**Libero Mail** and **Tiscali** use your normal password. If the login is refused, enable access from email programs (IMAP) in the provider's settings.

Then pick a period and press **Scan inboxes**.

## Privacy

- Passwords are stored in **Windows Credential Manager**, never in a file. Outlook has no password: its Microsoft sign-in is kept in `%APPDATA%\Sweep\microsoft_tokens.bin`, encrypted for your Windows user.
- Your account list, ignored senders and unsubscribe history are kept in `%APPDATA%\Sweep\state.json` on your computer.
- Sweep talks only to your mail provider (IMAP/SMTP, plus Microsoft's sign-in for Outlook) and, when you unsubscribe, to the sender's unsubscribe address. There is no server, account or tracking.

## Run from source

Requires Python 3.10 or newer.

```
git clone https://github.com/aaa58-hub/sweep.git
cd sweep
pip install -r requirements.txt
python app.py
```

Run the self-check with `python test_sweep.py` (prints `ok`).

## Outlook sign-in (for builders)

Outlook needs an app registered with Microsoft Entra, once, and its client ID in `MS_CLIENT_ID` in `app.py`
(or the `SWEEP_MS_CLIENT_ID` environment variable when running from source). Step-by-step guide: [docs/outlook-setup.md](docs/outlook-setup.md).

Without it, every other provider works and Outlook shows an error when added.

## Build the .exe

```
pip install pyinstaller
pyinstaller --onefile --windowed --name Sweep --add-data "index.html;." app.py
```

The app ends up in `dist\Sweep.exe`.

## License

[MIT](LICENSE)
