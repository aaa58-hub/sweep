"""Run: python test_sweep.py"""
import smtplib

from app import (Api, describe_error, find_unsub_link, gmail_categories, guess_provider, libero_category, list_folders,
                 parse_mailto, parse_unsub, safe_url)

# one-click first when the Post header says so, then mailto, then the web link as fallbacks
assert parse_unsub("<mailto:u@x.com>, <https://x.com/u?id=1>", "List-Unsubscribe=One-Click") == [
    ("one-click", "https://x.com/u?id=1"), ("mailto", "mailto:u@x.com"), ("link", "https://x.com/u?id=1")]
assert parse_unsub("<https://x.com/u>, <mailto:u@x.com?subject=stop>") == [("mailto", "mailto:u@x.com?subject=stop"), ("link", "https://x.com/u")]
assert parse_unsub("<http://x.com/u>") == [("link", "http://x.com/u")]
# folded header with whitespace inside the brackets
assert parse_unsub("<https://x.com/\r\n u>", "List-Unsubscribe=One-Click")[0] == ("one-click", "https://x.com/u")
assert parse_unsub("") == [] and parse_unsub("<ftp://x>") == []
# non-ASCII in a link (LinkedIn's "Il lavoro è un viaggio") is percent-encoded instead of crashing
assert safe_url("https://x.com/?t=Illavoroèunviaggio&a=b c") == "https://x.com/?t=Illavoro%C3%A8unviaggio&a=b%20c"
assert describe_error(smtplib.SMTPRecipientsRefused({"u@x.com": (550, b"5.1.1 too many invalid recipients")})) == \
    "the mail server refused the unsubscribe address (5.1.1 too many invalid recipients)"

assert parse_mailto("mailto:u@x.com?Subject=Stop%20it") == ("u@x.com", "Stop it", "unsubscribe")
for bad in ("mailto:a@x.com,b@y.com", "mailto:a@x.com%0ABcc:z@z.com", "mailto:nobody"):
    try:
        parse_mailto(bad)
        raise AssertionError(bad)
    except ValueError:
        pass

# grouping: two messages from one sender -> count 2, newest link kept; message without the header ignored
senders = {}
hdr = "From: News <News@shop.it>\r\nSubject: {s}\r\nDate: {d}\r\nList-Unsubscribe: <https://shop.it/{s}>\r\n\r\n"
Api._add_message("me@gmail.com", hdr.format(s="new", d="Tue, 10 Nov 2026 10:00:00 +0000").encode(), senders)
Api._add_message("me@gmail.com", hdr.format(s="old", d="Mon, 02 Nov 2026 10:00:00 +0000").encode(), senders)
Api._add_message("me@gmail.com", b"From: friend@x.it\r\nSubject: hi\r\n\r\n", senders)
s = senders["me@gmail.com|news@shop.it"]
assert len(senders) == 1 and s["count"] == 2 and s["_options"] == [("link", "https://shop.it/new")] and s["last"] == "2026-11-10", s


# Gmail categories: promotions wins a double label; sender takes its most common tab
class FakeImap:
    def uid(self, cmd, *args):
        hits = {"promotions": b"1 2", "updates": b"2 3", "social": b"", "forums": b""}
        return "OK", [hits[args[1].strip('"').split(":")[1]]]


assert gmail_categories(FakeImap(), ["SINCE", "01-Nov-2026"]) == {b"1": "promotions", b"2": "promotions", b"3": "updates"}
senders = {}
for cat in ("promotions", "promotions", "updates"):
    Api._add_message("me@gmail.com", hdr.format(s="x", d="Tue, 10 Nov 2026 10:00:00 +0000").encode(), senders, cat)
assert senders["me@gmail.com|news@shop.it"]["_cats"].most_common(1)[0][0] == "promotions"

# scan(): ignored flag, category, and Promotions sorted above bigger Updates senders
api = Api.__new__(Api)
api._state = {"accounts": [{"email": "me@gmail.com", "provider": "gmail"}], "handled": {}, "ignored": ["me@gmail.com|big@upd.it"]}


def fake_scan(addr, criteria, senders, errors):
    msg = "From: {f}\r\nDate: Tue, 10 Nov 2026 10:00:00 +0000\r\nList-Unsubscribe: <https://u.it>\r\n\r\n"
    for _ in range(5):
        Api._add_message(addr, msg.format(f="big@upd.it").encode(), senders, "updates")
    Api._add_message(addr, msg.format(f="promo@shop.it").encode(), senders, "promotions")


api._scan_account = fake_scan
out = api.scan(30)["senders"]
assert [(s["address"], s["category"], s["ignored"]) for s in out] == [("promo@shop.it", "promotions", False), ("big@upd.it", "updates", True)], out

# Libero: tab keywords -> category; bulk mail without the header is kept and unsubscribed via the body
assert libero_category(["\\seen", "$offerte"]) == "promotions" and libero_category(["$smartnletter", "$offerte"]) == "newsletter"
assert libero_category(["\\seen"]) == "primary"
senders = {}
bare = b"From: Shop <deals@shop.it>\r\nDate: Tue, 10 Nov 2026 10:00:00 +0000\r\n\r\n"
Api._add_message("me@libero.it", bare, senders, "primary", ("INBOX", "7"))  # personal-looking, no header: skipped
assert not senders
Api._add_message("me@libero.it", bare, senders, "promotions", ("steam", "7"))
s = senders["me@libero.it|deals@shop.it"]
assert s["method"] == "body" and s["_options"] == [] and s["_bodies"] == [("2026-11-10", ("steam", "7"))]
older_with_header = hdr.format(s="h", d="Mon, 02 Nov 2026 10:00:00 +0000").replace("News@shop.it", "deals@shop.it")
Api._add_message("me@libero.it", older_with_header.encode(), senders, "promotions", ("INBOX", "3"))
assert s["method"] == "link" and s["last"] == "2026-11-10", s  # a header link beats a newer body-only message

# unsubscribe(): a failing way falls through to the next; all failing -> every reason shown once
api = Api.__new__(Api)
api._state, api._save, tried = {"handled": {}}, lambda: None, []


def fake_unsub(account, method, target):
    tried.append(method)
    if method != "link":
        raise smtplib.SMTPResponseException(421, b"too many requests")
    return "browser"


api._unsubscribe = fake_unsub
api._found = {"k": ("me@libero.it", [("mailto", "mailto:u@x.it"), ("link", "https://x.it/u"), ("body", [])]),
              "dead": ("me@libero.it", [("mailto", "mailto:u@x.it"), ("body", [])])}
assert api.unsubscribe("k") == {"key": "k", "status": "browser"} and tried == ["mailto", "link"]
assert api.unsubscribe("dead") == {"key": "dead", "status": "failed", "detail": "the mail server said: too many requests"}


# folders: Gmail keeps All Mail/Spam/Trash; Libero keeps everything but Sent/Drafts
class FakeList:
    def __init__(self, lines):
        self.lines = lines

    def list(self):
        return "OK", self.lines


gmail = [b'(\\HasNoChildren) "/" "INBOX"', b'(\\All \\HasNoChildren) "/" "[Gmail]/Tutti i messaggi"',
         b'(\\Junk) "/" "[Gmail]/Spam"', b'(\\Trash) "/" "[Gmail]/Cestino"', b'(\\Noselect) "/" "[Gmail]"']
assert list_folders(FakeList(gmail), "gmail") == ["[Gmail]/Tutti i messaggi", "[Gmail]/Spam", "[Gmail]/Cestino"]
libero = [b'(\\HasNoChildren) "/" INBOX', b'(\\Sent) "/" outbox', b'(\\Drafts) "/" draft',
          b'(\\Junk) "/" Spam', b'(\\HasNoChildren) "/" Work/Projects']
assert list_folders(FakeList(libero), "libero") == ["INBOX", "Spam", "Work/Projects"]


# unsubscribe link inside the body
def mail(html_body):
    head = "From: a@b.it\r\nMIME-Version: 1.0\r\nContent-Type: text/html; charset=utf-8\r\n\r\n"
    return (head + html_body).encode()


assert find_unsub_link(mail('<a href="https://x.it/shop">Shop</a> <a href="https://x.it/u?a=1&amp;b=2">Annulla l’iscrizione</a>')) == "https://x.it/u?a=1&b=2"
assert find_unsub_link(mail('<a href="https://x.it/home">Home</a> <a href="https://x.it/unsubscribe/9">qui</a>')) == "https://x.it/unsubscribe/9"
assert find_unsub_link(mail('Per disiscriverti clicca <a href="https://x.it/r/42">qui</a>.')) == "https://x.it/r/42"
assert find_unsub_link(mail('<a href="https://x.it/">Ciao</a>')) is None
assert find_unsub_link(mail('<a href="https://x.it/p">Gestisci preferenze</a> <a href="https://x.it/u">Unsubscribe</a>')) == "https://x.it/u"
assert find_unsub_link(mail('<a href="https://x.it/web">Web</a> <a href="https://x.it/p">Update Profile</a>')) == "https://x.it/p"


# a "too many" answer from the mail server pauses further unsubscribe emails instead of hammering it
import app  # noqa: E402


class FakeSMTP:
    def __init__(self, *a, **k): pass
    def __enter__(self): return self
    def __exit__(self, *a): pass
    def login(self, *a): pass
    def send_message(self, msg): raise smtplib.SMTPRecipientsRefused({"u@x.it": (452, b"too many invalid recipients")})


app.smtplib.SMTP_SSL, real_smtp = FakeSMTP, app.smtplib.SMTP_SSL
api = Api.__new__(Api)
api._smtp_next, api._account = {}, lambda a: ("libero", "pw")
for expected in (smtplib.SMTPRecipientsRefused, RuntimeError):
    try:
        api._send_mail("me@libero.it", "mailto:u@x.it")
        raise AssertionError("should fail")
    except expected:
        pass
app.smtplib.SMTP_SSL = real_smtp

# providers: domains -> provider; Outlook sends over STARTTLS (587) and signs in with an XOAUTH2 token, not a password
assert [guess_provider(a) for a in ("x@Hotmail.it", "x@live.com", "x@me.com", "x@ymail.com", "x@virgilio.it", "x@ditta.it")] ==     ["outlook", "outlook", "icloud", "yahoo", "virgilio", None]
seen = []


class FakeTLS(FakeSMTP):
    def __init__(self, host, port, **k): seen.append((host, port))
    def starttls(self, **k): seen.append("starttls")
    def ehlo(self): seen.append("ehlo")
    def auth(self, mech, obj): seen.append((mech, obj(), obj(b'{"status":"400"}')))
    def send_message(self, msg): seen.append(msg["From"])


app.smtplib.SMTP, real_plain = FakeTLS, app.smtplib.SMTP
api._smtp_next, api._account = {}, lambda a: ("outlook", "tok")
assert api._send_mail("me@hotmail.it", "mailto:u@x.it") == "done"
assert seen == [("smtp-mail.outlook.com", 587), "starttls", "ehlo",
                ("XOAUTH2", "user=me@hotmail.itauth=Bearer tok", ""), "me@hotmail.it"]
app.smtplib.SMTP = real_plain

# no header, Primary tab, but a mass-mail marker (AliExpress): kept only if its newest email has a real unsubscribe link
senders = {}
marked = "From: AliExpress <ae-picks@deals.aliexpress.com>\r\nFeedback-ID: 1:2:S-1\r\nDate: Tue, 10 Nov 2026 10:00:00 +0000\r\n\r\n"
Api._add_message("me@gmail.com", marked.encode(), senders, "primary", ("Spam", "1"))
Api._add_message("me@gmail.com", marked.replace("ae-picks@deals.aliexpress.com", "no-reply@accounts.google.com").encode(),
                 senders, "primary", ("INBOX", "2"))
Api._add_message("me@gmail.com", b"From: friend@gmail.com\r\nDate: Tue, 10 Nov 2026 10:00:00 +0000\r\n\r\n", senders, "primary", ("INBOX", "3"))
assert sorted(senders) == ["me@gmail.com|ae-picks@deals.aliexpress.com", "me@gmail.com|no-reply@accounts.google.com"]


class FakeBodies:
    bodies = {"1": mail('<a href="https://ae.com/u">Unsubscribe</a>'), "2": mail('<a href="https://g.co/p">Manage preferences</a>')}

    def select(self, *a, **k): return "OK", [b""]
    def uid(self, cmd, uid, what): return "OK", [(b"1 (BODY[] {1}", self.bodies[uid]), b")"]


api = Api.__new__(Api)
api._window = None
api._drop_unproven(FakeBodies(), "me@gmail.com", senders)
assert list(senders) == ["me@gmail.com|ae-picks@deals.aliexpress.com"]  # a preferences link alone is not proof
print("ok")
