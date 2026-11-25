"""Sweep — find newsletters in Gmail / Libero Mail and unsubscribe from them."""
import email
import email.policy
import email.utils
import html
import imaplib
import json
import os
import re
import smtplib
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from collections import Counter
from datetime import date, timedelta
from email.message import EmailMessage
from pathlib import Path

import keyring
import keyring.errors
import msal
import msal_extensions

APP = "Sweep"
STATE_FILE = Path(os.environ.get("APPDATA", Path.home())) / APP / "state.json"
HOSTS = {  # provider -> (imap host on 993 with TLS, smtp host, smtp port: 465 implicit TLS, 587 STARTTLS)
    "gmail": ("imap.gmail.com", "smtp.gmail.com", 465),
    "libero": ("imapmail.libero.it", "smtp.libero.it", 465),
    "outlook": ("imap-mail.outlook.com", "smtp-mail.outlook.com", 587),
    "yahoo": ("imap.mail.yahoo.com", "smtp.mail.yahoo.com", 465),
    "icloud": ("imap.mail.me.com", "smtp.mail.me.com", 587),
    "virgilio": ("in.virgilio.it", "out.virgilio.it", 465),
    "tiscali": ("imap.tiscali.it", "smtp.tiscali.it", 465),
}
DOMAINS = {
    "gmail": ("gmail.com", "googlemail.com"),
    "libero": ("libero.it", "inwind.it", "iol.it", "blu.it", "giallo.it"),
    "outlook": ("outlook.com", "outlook.it", "hotmail.com", "hotmail.it", "live.com", "live.it", "msn.com"),
    "yahoo": ("yahoo.com", "yahoo.it", "ymail.com", "rocketmail.com"),
    "icloud": ("icloud.com", "me.com", "mac.com"),
    "virgilio": ("virgilio.it",),
    "tiscali": ("tiscali.it",),
}
LOGIN_HELP = {
    "gmail": "Gmail needs an App Password, not your normal password: turn on 2-Step Verification, "
             "then create one at myaccount.google.com/apppasswords.",
    "yahoo": "Yahoo needs an App Password: create one in Account security → Generate app password.",
    "icloud": "iCloud needs an app-specific password: create one at account.apple.com → Sign-In and Security.",
    "virgilio": "Virgilio needs its 'Password Sicura' for email programs, created in the webmail settings.",
    "outlook": "Check that IMAP is on in Outlook.com → Settings → Mail → Forwarding and IMAP.",
}
# Outlook.com refuses passwords over IMAP/SMTP: it signs in with Microsoft (OAuth2) instead.
# Client ID of the "Sweep" app registered in Microsoft Entra (personal accounts, public client, redirect http://localhost).
MS_CLIENT_ID = os.environ.get("SWEEP_MS_CLIENT_ID", "")
MS_SIGNIN_SECONDS = 300  # give up waiting for the browser sign-in (tab closed) after this
MS_SCOPES = ["https://outlook.office.com/IMAP.AccessAsUser.All", "https://outlook.office.com/SMTP.Send"]
MONTHS = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()  # IMAP dates must be English
FIELDS = ("(UID FLAGS BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE LIST-UNSUBSCRIBE LIST-UNSUBSCRIBE-POST"
          " FEEDBACK-ID LIST-ID PRECEDENCE)])")

# Libero marks its inbox tabs with IMAP keywords. First match wins.
LIBERO_FLAGS = [("$smartnletter", "newsletter"), ("$offerte", "promotions"), ("$dem_offerte", "promotions"),
                ("$dem_inbox", "promotions"), ("$social", "social"), ("$ordini", "orders"),
                ("$prenotazioni", "bookings"), ("$utenze", "utilities")]
# Categories that mean bulk mail even when the sender left out the List-Unsubscribe header.
BULK = {"promotions", "newsletter", "social"}
CATEGORY_RANK = {c: i for i, c in enumerate(
    ["promotions", "newsletter", "social", "forums", "updates", "orders", "bookings", "utilities", "primary", None])}
UNSUB_WORDS = re.compile(
    r"unsub|disiscri|cancellati|(?:cancella|annulla)(?:re)?\s+(?:la\s+|l['’]\s*)?iscrizione|opt[-\s]?out"
    r"|non\s+ricevere|preferenze\s+(?:e-?mail|di\s+comunicazione)|email\s+preferences", re.I)
# Fallback when no real unsubscribe link exists: the sender's preference centre.
PREFS_WORDS = re.compile(r"preferenz|preferences|update\s+(?:your\s+)?profile|aggiorna\s+(?:il\s+(?:tuo\s+)?)?profilo", re.I)


def parse_unsub(header, post_header=""):
    """List-Unsubscribe header -> [(method, target), ...], best first: RFC 8058 one-click, then mailto, then the web link.
    The later ones are fallbacks for when an earlier one fails."""
    links = [re.sub(r"\s+", "", l) for l in re.findall(r"<([^>]+)>", header or "")]
    https = next((l for l in links if l.lower().startswith("https://")), None)
    mailto = next((l for l in links if l.lower().startswith("mailto:")), None)
    web = https or next((l for l in links if l.lower().startswith("http://")), None)
    out = [("one-click", https)] if https and "one-click" in (post_header or "").lower() else []
    if mailto:
        out.append(("mailto", mailto))
    if web:
        out.append(("link", web))
    return out


def safe_url(url):
    """Percent-encode non-ASCII characters and spaces: urllib refuses them ("'ascii' codec can't encode...")."""
    return re.sub(r"[^\x21-\x7e]", lambda c: urllib.parse.quote(c.group()), url)


def describe_error(e):
    """Exception -> short text for the row. Mail servers answer (code, b'text'): show just the text."""
    if isinstance(e, smtplib.SMTPRecipientsRefused):
        code, text = next(iter(e.recipients.values()))
        return f"the mail server refused the unsubscribe address ({text.decode(errors='replace')})"
    if isinstance(e, smtplib.SMTPResponseException):
        return f"the mail server said: {e.smtp_error.decode(errors='replace')}"
    if isinstance(e, urllib.error.HTTPError):
        return f"the sender's website answered {e.code} {e.reason}"
    return str(e)


def parse_mailto(target):
    """mailto:x@y?subject=..&body=.. -> (to, subject, body). Rejects anything but a single plain address."""
    u = urllib.parse.urlsplit(target)
    to = urllib.parse.unquote(u.path).strip()
    if not re.fullmatch(r"[^@\s,;<>]+@[^@\s,;<>]+", to):
        raise ValueError(f"unusable unsubscribe address: {to!r}")
    q = {k.lower(): v[0] for k, v in urllib.parse.parse_qs(u.query).items()}
    return to, q.get("subject", "unsubscribe"), q.get("body", "unsubscribe")


def bulk_marked(msg):
    """Mass-mail markers that some senders set while leaving out List-Unsubscribe (AliExpress: Feedback-ID)."""
    return bool(msg["Feedback-ID"] or msg["List-Id"]) or str(msg["Precedence"] or "").strip().lower() in ("bulk", "list")


def find_unsub_link(raw, prefs=True):
    """Full email bytes -> the unsubscribe URL in its body, or None. For senders with no List-Unsubscribe header.
    prefs=False accepts only a real unsubscribe link, not a preferences page."""
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    htmls, texts = [], []
    for part in msg.walk():
        if part.get_content_type() in ("text/html", "text/plain"):
            try:
                (htmls if part.get_content_type() == "text/html" else texts).append(part.get_content())
            except Exception:  # undecodable part: try the others
                continue
    anchors = [(html.unescape(h).strip(), re.sub(r"<[^>]+>", " ", t))
               for b in htmls for h, t in re.findall(r"""<a\b[^>]*?href\s*=\s*["']([^"']+)["'][^>]*>(.*?)</a>""", b, re.I | re.S)]
    anchors = [(h, t) for h, t in anchors if h.lower().startswith(("http://", "https://"))]
    for words in (UNSUB_WORDS, PREFS_WORDS)[:2 if prefs else 1]:  # a real unsubscribe link anywhere beats a preferences link
        for h, t in anchors:  # link text says "unsubscribe"
            if words.search(t):
                return h
        for h, t in anchors:  # link address says "unsubscribe"
            if words.search(h):
                return h
        for b in htmls:  # "per disiscriverti clicca <a href=...>qui</a>"
            for mt in words.finditer(b):
                near = re.search(r"""href\s*=\s*["'](https?://[^"']+)""", b[mt.end():mt.end() + 600], re.I)
                if near:
                    return html.unescape(near.group(1))
        for b in texts:  # plain text: "Unsubscribe: https://..."
            for mt in words.finditer(b):
                near = re.search(r"https?://[^\s<>\"']+", b[mt.end():mt.end() + 300])
                if near:
                    return near.group(0)
    return None


def link_in(m, wheres, prefs=True):
    """First unsubscribe link found in the emails at [(folder, uid), ...] on a logged-in IMAP session."""
    for folder, uid in wheres:
        if m.select(f'"{folder}"', readonly=True)[0] != "OK":
            continue
        resp = m.uid("FETCH", uid, "(BODY.PEEK[])")[1]
        link = find_unsub_link(next((p[1] for p in resp if isinstance(p, tuple)), b""), prefs)
        if link:
            return link
    return None


def list_folders(m, provider):
    """Folders worth scanning. Gmail: All Mail + Spam + Trash (labels would repeat All Mail). Others: all but Sent/Drafts."""
    out = []
    for line in m.list()[1]:
        mt = re.match(rb'\((?P<f>[^)]*)\) (?:"[^"]*"|NIL) (?P<n>.+)$', line) if isinstance(line, bytes) else None
        if not mt:
            continue
        flags, name = mt["f"].decode().lower(), mt["n"].decode().strip('"')
        if "\\noselect" in flags:
            continue
        if provider == "gmail" and any(f in flags for f in ("\\all", "\\junk", "\\trash")):
            out.append(name)
        elif provider != "gmail" and not any(f in flags for f in ("\\sent", "\\drafts")):
            out.append(name)
    return out or ["INBOX"]


def gmail_categories(m, criteria):
    """UID -> Gmail tab, via Gmail's IMAP search extension. Anything not returned is in Primary."""
    cats = {}
    for cat in ("updates", "forums", "social", "promotions"):  # promotions last so it wins a double label
        for uid in m.uid("SEARCH", "X-GM-RAW", f'"category:{cat}"', *criteria)[1][0].split():
            cats[uid] = cat
    return cats


def libero_category(flags):
    return next((cat for flag, cat in LIBERO_FLAGS if flag in flags), "primary")


def guess_provider(addr):
    domain = addr.rsplit("@", 1)[-1].lower()
    return next((p for p, domains in DOMAINS.items() if domain in domains), None)


def friendly_login_error(provider, err):
    return f"The mail server refused the login ({err}). " + LOGIN_HELP.get(
        provider, "Check the password and that access from email programs (IMAP) is enabled for the mailbox.")


def xoauth2(addr, token):
    """SASL XOAUTH2 string; imaplib and smtplib base64 it themselves."""
    return f"user={addr}\x01auth=Bearer {token}\x01\x01"


def imap_login(m, provider, addr, secret):
    """secret is a password, or a Microsoft access token for Outlook."""
    if provider == "outlook":
        m.authenticate("XOAUTH2", lambda challenge: b"" if challenge else xoauth2(addr, secret).encode())
    else:
        m.login(addr, secret)


def smtp_login(s, provider, addr, secret):
    if provider == "outlook":  # on failure the server sends a JSON challenge: answer empty to get the error
        s.ehlo()  # auth(), unlike login(), doesn't redo the EHLO that STARTTLS wiped
        s.auth("XOAUTH2", lambda challenge=None: "" if challenge else xoauth2(addr, secret))
    else:
        s.login(addr, secret)


class Api:
    """Every public method is callable from the page as window.pywebview.api.<name>()."""

    def __init__(self):
        self._window = None
        self._found = {}  # key -> (account, [(method, target), ...] best first)
        self._imap = {}  # account -> IMAP session reused while unsubscribing
        self._smtp_next = {}  # account -> monotonic time the next unsubscribe email may go out
        self._ms = None  # Microsoft sign-in, created on first use
        try:
            self._state = json.loads(STATE_FILE.read_text("utf-8"))
        except (OSError, ValueError):
            self._state = {}
        for k, v in (("accounts", []), ("handled", {}), ("ignored", []), ("theme", None)):
            self._state.setdefault(k, v)

    def _save(self):
        STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        STATE_FILE.write_text(json.dumps(self._state, indent=2), "utf-8")

    def _account(self, addr):
        """-> (provider, secret): the stored password, or a fresh Microsoft access token for Outlook."""
        provider = next(a for a in self._state["accounts"] if a["email"] == addr)["provider"]
        return provider, self._ms_token(addr) if provider == "outlook" else keyring.get_password(APP, addr)

    def _ms_app(self):
        if not MS_CLIENT_ID:
            raise RuntimeError("this copy of Sweep has no Microsoft client ID, so it cannot sign in to Outlook (see README)")
        if not self._ms:
            # refresh tokens live in a file encrypted for this Windows user (DPAPI): too big for Credential Manager
            cache = msal_extensions.PersistedTokenCache(msal_extensions.build_encrypted_persistence(
                str(STATE_FILE.with_name("microsoft_tokens.bin"))))
            self._ms = msal.PublicClientApplication(MS_CLIENT_ID, authority="https://login.microsoftonline.com/consumers",
                                                    token_cache=cache)
        return self._ms

    def _ms_token(self, addr, interactive=False):
        """Access token for addr, refreshed silently; interactive opens Microsoft's sign-in page in the browser."""
        app = self._ms_app()
        known = [a for a in app.get_accounts() if a.get("username", "").lower() == addr]
        r = app.acquire_token_silent(MS_SCOPES, account=known[0]) if known else None
        if not r and interactive:
            r = app.acquire_token_interactive(MS_SCOPES, login_hint=addr, prompt="select_account", timeout=MS_SIGNIN_SECONDS)
            if "access_token" not in r:
                raise RuntimeError(f"Microsoft sign-in did not finish ({r.get('error_description') or r.get('error')})")
            who = str((r.get("id_token_claims") or {}).get("preferred_username", "")).lower()
            if who != addr:
                for a in app.get_accounts(username=who):
                    app.remove_account(a)
                raise RuntimeError(f"you signed in to Microsoft as {who or 'another account'}, not {addr}")
        if not r or "access_token" not in r:
            detail = (r or {}).get("error_description", "").splitlines()[0:1]
            raise RuntimeError("Microsoft sign-in needed: remove the account and add it again"
                               + (f" ({detail[0]})" if detail else ""))
        return r["access_token"]

    def _connect(self, addr):
        provider, secret = self._account(addr)
        m = imaplib.IMAP4_SSL(HOSTS[provider][0], timeout=60)
        imap_login(m, provider, addr, secret)
        return provider, m

    def _progress(self, payload):
        if self._window:
            self._window.evaluate_js(f"window.onProgress({json.dumps(payload)})")

    # ---- accounts -------------------------------------------------------
    def get_settings(self):
        return {"accounts": self._state["accounts"], "theme": self._state["theme"]}

    def set_theme(self, theme):
        self._state["theme"] = theme
        self._save()

    def ignore(self, key, on):
        ignored = set(self._state["ignored"])
        ignored.add(key) if on else ignored.discard(key)
        self._state["ignored"] = sorted(ignored)
        self._save()

    def guess_provider(self, addr):
        return guess_provider(addr)

    def add_account(self, addr, password, provider):
        addr = addr.strip().lower()
        if provider not in HOSTS:
            return {"error": "Choose your email provider."}
        try:
            secret = self._ms_token(addr, interactive=True) if provider == "outlook" else password
            with imaplib.IMAP4_SSL(HOSTS[provider][0], timeout=30) as m:
                imap_login(m, provider, addr, secret)
        except imaplib.IMAP4.error as e:
            return {"error": friendly_login_error(provider, e)}
        except OSError as e:
            return {"error": f"Could not reach the mail server: {e}"}
        except RuntimeError as e:  # Microsoft sign-in
            return {"error": f"{str(e)[:1].upper()}{str(e)[1:]}."}
        if provider != "outlook":
            keyring.set_password(APP, addr, password)
        self._state["accounts"] = [a for a in self._state["accounts"] if a["email"] != addr]
        self._state["accounts"].append({"email": addr, "provider": provider})
        self._save()
        return {"ok": True, "accounts": self._state["accounts"]}

    def remove_account(self, addr):
        if any(a["email"] == addr and a["provider"] == "outlook" for a in self._state["accounts"]) and MS_CLIENT_ID:
            for a in self._ms_app().get_accounts(username=addr):  # forget its Microsoft tokens too
                self._ms_app().remove_account(a)
        self._state["accounts"] = [a for a in self._state["accounts"] if a["email"] != addr]
        self._save()
        try:
            keyring.delete_password(APP, addr)
        except keyring.errors.PasswordDeleteError:
            pass
        return self._state["accounts"]

    # ---- scan -----------------------------------------------------------
    def scan(self, days):
        days = int(days)
        d = date.today() - timedelta(days=days)
        criteria = ["SINCE", f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year}"] if days else ["ALL"]
        senders, errors = {}, []
        for acct in self._state["accounts"]:
            addr = acct["email"]
            try:
                self._scan_account(addr, criteria, senders, errors)
            except Exception as e:  # one broken account must not sink the others
                errors.append(f"{addr}: {e}")
        for s in senders.values():  # last resort for everyone: the link inside the newest few emails (receipts carry none)
            newest = [w for _, w in sorted(s.pop("_bodies"), reverse=True)[:5]]
            s["_options"] = s["_options"] + ([("body", newest)] if newest else [])
        self._found = {k: (s.pop("_account"), s.pop("_options")) for k, s in senders.items()}
        ignored = set(self._state["ignored"])
        for k, s in senders.items():
            cats, folders = s.pop("_cats"), s.pop("_folders")
            del s["_rank"], s["_sure"]
            s["category"] = cats.most_common(1)[0][0] if cats else None
            s["folder"] = folders.most_common(1)[0][0] if folders else None
            s["handled"] = self._state["handled"].get(k)
            s["ignored"] = k in ignored
        ordered = sorted(senders.values(), key=lambda s: (CATEGORY_RANK.get(s["category"], 99), -s["count"]))
        return {"senders": ordered, "errors": errors}

    def _scan_account(self, addr, criteria, senders, errors):
        provider, m = self._connect(addr)
        with m:
            for folder in list_folders(m, provider):
                try:
                    self._scan_folder(m, provider, addr, folder, criteria, senders)
                except imaplib.IMAP4.error as e:  # one unreadable folder must not sink the account
                    errors.append(f"{addr} › {folder}: {e}")
            self._drop_unproven(m, addr, senders)

    def _drop_unproven(self, m, addr, senders):
        """Senders kept only for a mass-mail marker stay if their newest email holds a real unsubscribe link.
        Drops security alerts and receipts, which carry the same markers."""
        unsure = [k for k, s in senders.items() if s["account"] == addr and not s["_sure"]]
        for i, k in enumerate(unsure):
            self._progress({"account": addr, "folder": "senders without an unsubscribe header", "done": i, "total": len(unsure)})
            try:
                found = link_in(m, [max(senders[k]["_bodies"])[1]], prefs=False)
            except imaplib.IMAP4.error:
                found = None
            if not found:
                del senders[k]

    def _scan_folder(self, m, provider, addr, folder, criteria, senders):
        if m.select(f'"{folder}"', readonly=True)[0] != "OK":
            return
        uids = m.uid("SEARCH", None, *criteria)[1][0].split()
        cats = gmail_categories(m, criteria) if provider == "gmail" and uids else {}
        for i in range(0, len(uids), 250):
            self._progress({"account": addr, "folder": folder, "done": i, "total": len(uids)})
            resp = m.uid("FETCH", b",".join(uids[i:i + 250]), FIELDS)[1]
            for j, part in enumerate(resp):
                if not isinstance(part, tuple):
                    continue
                # UID / FLAGS can come before or after the header literal
                tail = resp[j + 1] if j + 1 < len(resp) and isinstance(resp[j + 1], bytes) else b""
                uid = re.search(rb"UID (\d+)", part[0] + tail)
                uid = uid.group(1) if uid else None
                flags = re.search(rb"FLAGS \(([^)]*)\)", part[0] + tail)
                if provider == "gmail":
                    cat = cats.get(uid, "primary") if uid else None
                else:
                    cat = libero_category(flags.group(1).decode().lower().split() if flags else [])
                self._add_message(addr, part[1], senders, cat, (folder, uid.decode()) if uid else None)

    @staticmethod
    def _add_message(addr, raw, senders, category=None, where=None):
        try:
            msg = email.message_from_bytes(raw, policy=email.policy.default)
            options = parse_unsub(str(msg["List-Unsubscribe"] or ""), str(msg["List-Unsubscribe-Post"] or ""))
            sure = bool(options) or category in BULK
            # no header: keep bulk mail (by tab, or by mass-mail marker, checked later) to use the link inside
            if not options and not (where and (sure or bulk_marked(msg))):
                return
            name, sender = email.utils.parseaddr(str(msg["From"] or ""))
            subject = str(msg["Subject"] or "")
        except Exception:  # malformed headers: skip the message, keep scanning
            return
        try:
            when = email.utils.parsedate_to_datetime(str(msg["Date"])).date().isoformat()
        except (TypeError, ValueError):
            when = ""
        key = f"{addr}|{sender.lower()}"
        s = senders.setdefault(key, {"key": key, "account": addr, "address": sender.lower(), "name": name or sender,
                                     "count": 0, "last": "", "subject": "", "method": "",
                                     "_cats": Counter(), "_folders": Counter(), "_rank": (False, ""), "_bodies": [], "_options": [], "_sure": False})
        s["count"] += 1
        s["_sure"] |= sure
        if category:
            s["_cats"][category] += 1
        if where:
            s["_folders"][where[0]] += 1
            s["_bodies"].append((when, where))
        if when >= s["last"]:
            s.update(last=when, subject=subject)
        # header links beat links found in the body; among equals, the newest wins (old links may have expired)
        rank = (bool(options), when)
        if rank >= s["_rank"]:
            s.update(method=options[0][0] if options else "body", _account=addr, _options=options, _rank=rank)

    # ---- unsubscribe ----------------------------------------------------
    def unsubscribe(self, key):
        account, options = self._found[key]
        problems = []
        for method, target in options:
            try:
                status = self._unsubscribe(account, method, target)
                break
            except Exception as e:  # try the next way, e.g. a dead unsubscribe address, then the web link
                problems.append(describe_error(e))
        else:
            detail = "; ".join(dict.fromkeys(problems)) or "this sender gave no way to unsubscribe"
            return {"key": key, "status": "failed", "detail": detail}
        self._state["handled"][key] = status
        self._save()
        return {"key": key, "status": status}

    def _session(self, account):
        """One IMAP login per account for a whole batch: a fresh login per sender gets rate-limited."""
        m = self._imap.get(account)
        try:
            if m and m.noop()[0] == "OK":
                return m
        except (imaplib.IMAP4.error, OSError):
            pass
        self._imap[account] = m = self._connect(account)[1]
        return m

    def _link_from_bodies(self, account, wheres):
        return link_in(self._session(account), wheres)

    def _send_mail(self, account, target):
        to, subject, body = parse_mailto(target)
        wait = self._smtp_next.get(account, 0) - time.monotonic()
        if wait > 60:
            raise RuntimeError("the mail server asked Sweep to slow down, so unsubscribe emails are paused for a few minutes")
        time.sleep(max(wait, 0))
        provider, secret = self._account(account)
        msg = EmailMessage()
        msg["From"], msg["To"], msg["Subject"] = account, to, subject
        msg.set_content(body)
        host, port, tls = *HOSTS[provider][1:], ssl.create_default_context()
        try:
            with (smtplib.SMTP_SSL(host, port, context=tls, timeout=30) if port == 465
                  else smtplib.SMTP(host, port, timeout=30)) as s:
                if port != 465:
                    s.starttls(context=tls)
                smtp_login(s, provider, account, secret)
                s.send_message(msg)
        except (smtplib.SMTPResponseException, smtplib.SMTPRecipientsRefused) as e:
            if "too many" in str(e).lower():  # "too many requests" / "too many invalid recipients": back off
                self._smtp_next[account] = time.monotonic() + 600
            raise
        finally:
            # ponytail: fixed 5 s gap between sends; make it adaptive if servers still complain
            self._smtp_next[account] = max(self._smtp_next.get(account, 0), time.monotonic() + 5)
        return "done"

    def _unsubscribe(self, account, method, target):
        if method == "body":
            target = self._link_from_bodies(account, target)
            if not target:
                raise ValueError("no unsubscribe link in this sender's recent emails")
            method = "link"
        if method == "one-click":  # on failure the next option opens the same address in the browser
            req = urllib.request.Request(safe_url(target), data=b"List-Unsubscribe=One-Click", method="POST",
                                         headers={"Content-Type": "application/x-www-form-urlencoded",
                                                  "User-Agent": f"{APP}/1.0"})
            with urllib.request.urlopen(req, timeout=20):
                return "done"  # urlopen raises on 4xx/5xx
        if method == "mailto":
            return self._send_mail(account, target)
        if not target.lower().startswith(("https://", "http://")):
            raise ValueError("unsubscribe link is not a web address")
        webbrowser.open(safe_url(target))
        return "browser"


if __name__ == "__main__":
    import webview

    api = Api()
    api._window = webview.create_window(APP, str(Path(__file__).with_name("index.html")), js_api=api,
                                        width=1180, height=800, min_size=(820, 600))
    webview.start(debug="--debug" in sys.argv)
