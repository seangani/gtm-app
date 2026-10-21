#!/usr/bin/env python3
"""Send personalized outreach emails from your own mailbox, one at a time.

Reads contacts from a CSV, fills {placeholders} in a template, and sends each
email through either:

  mail     - the Apple Mail app already signed in on this Mac (default on Mac).
             Add your school account in System Settings > Internet Accounts.
  outlook  - the Outlook desktop app already signed in on this computer
             (Windows only, needs `pip install pywin32`). Works with
             school/work Microsoft 365 accounts that block SMTP and API apps,
             because it is just Outlook sending mail you could have sent by hand.
  smtp     - a direct SMTP login (smtp.office365.com:587 by default). Many
             Microsoft 365 schools turn this off; use `--test` to check.

Dry run is the default. Nothing is sent unless you pass --send.

Every successful send is appended to sent_log.csv, and anyone already in that
log is skipped, so you can stop and re-run safely without double-emailing.
"""

import argparse
import csv
import getpass
import os
import random
import re
import smtplib
import subprocess
import sys
import time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

HERE = Path(__file__).resolve().parent


def load_template(path):
    text = Path(path).read_text(encoding="utf-8")
    first, _, body = text.partition("\n")
    if not first.lower().startswith("subject:"):
        sys.exit(f"{path}: first line must be 'Subject: ...'")
    return first.split(":", 1)[1].strip(), body.lstrip("\n")


def column_key(header):
    # "First Name" -> first_name, "Email (Apollo)" -> email_apollo
    return re.sub(r"[^a-z0-9]+", "_", (header or "").lower()).strip("_")


def load_contacts(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = [{column_key(k): (v or "").strip() for k, v in r.items() if k} for r in csv.DictReader(f)]
    email_col = next((k for k in (rows[0] if rows else {}) if k == "email" or k.startswith("email_")), None)
    if not email_col:
        sys.exit(f"{path}: needs a header row with an 'email' column")
    for r in rows:
        r["email"] = r[email_col]
    return [r for r in rows if r["email"]]


def load_sent(path):
    if not path.exists():
        return set()
    with open(path, newline="", encoding="utf-8") as f:
        return {r["email"].lower() for r in csv.DictReader(f)}


def log_sent(path, email, subject):
    new = not path.exists()
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["email", "subject", "sent_at"])
        w.writerow([email, subject, datetime.now().isoformat(timespec="seconds")])


def fill(template, contact, where):
    try:
        return template.format(**contact)
    except KeyError as e:
        sys.exit(f"{where} uses {{{e.args[0]}}} but contacts.csv has no '{e.args[0]}' column")


MAIL_SCRIPT = """
on run {toAddr, subj, msgBody, fromAddr}
    tell application "Mail"
        set m to make new outgoing message with properties {subject:subj, content:msgBody, visible:false}
        tell m
            make new to recipient at end of to recipients with properties {address:toAddr}
            if fromAddr is not "" then set sender to fromAddr
        end tell
        send m
    end tell
end run
"""


# Same as MAIL_SCRIPT, but sends formatted text: the whole email in one font,
# and the bold text (the signature name) in bold wherever a line starts with it.
# Mail sends this as HTML with a plain-text copy.
RICH_MAIL_SCRIPT = """
on run {toAddr, subj, msgBody, fromAddr, fontName, boldFont, fontSize, boldText}
    set boldLen to (count of characters of boldText)
    tell application "Mail"
        set m to make new outgoing message with properties {subject:subj, content:msgBody, visible:false}
        tell m
            make new to recipient at end of to recipients with properties {address:toAddr}
            if fromAddr is not "" then set sender to fromAddr
            tell content
                set font to fontName
                set size to (fontSize as integer)
                if boldLen > 0 then
                    repeat with i from 1 to (count of paragraphs)
                        if (paragraph i as string) starts with boldText then
                            set font of characters 1 thru boldLen of paragraph i to boldFont
                        end if
                    end repeat
                end if
            end tell
        end tell
        send m
    end tell
end run
"""


class MailSender:
    def __init__(self, args):
        if sys.platform != "darwin":
            sys.exit("Mail backend only works on a Mac")
        self.sender = args.sender or ""
        self.rich = args.rich
        self.bold = args.bold

    def send(self, to, subject, body):
        if self.rich:
            cmd = ["osascript", "-e", RICH_MAIL_SCRIPT, to, subject, body, self.sender,
                   "Helvetica", "Helvetica-Bold", "14", self.bold]
        else:
            cmd = ["osascript", "-e", MAIL_SCRIPT, to, subject, body, self.sender]
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode:
            raise RuntimeError(r.stderr.strip() or "Mail refused to send")

    def close(self):
        pass


class OutlookSender:
    def __init__(self, _args):
        try:
            import win32com.client
        except ImportError:
            sys.exit("Outlook backend needs Windows + `pip install pywin32`")
        self.app = win32com.client.Dispatch("Outlook.Application")

    def send(self, to, subject, body):
        mail = self.app.CreateItem(0)  # 0 = olMailItem
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        mail.Send()

    def close(self):
        pass


class SmtpSender:
    def __init__(self, args):
        self.user = args.sender or input("Your email address: ").strip()
        password = os.environ.get("OUTREACH_PASSWORD") or getpass.getpass("Password: ")
        try:
            self.server = smtplib.SMTP(args.host, args.port, timeout=30)
        except OSError as e:
            sys.exit(f"Could not reach {args.host}:{args.port}: {e}")
        self.server.starttls()
        try:
            self.server.login(self.user, password)
        except smtplib.SMTPAuthenticationError as e:
            sys.exit(
                f"SMTP login refused ({e.smtp_code}): {e.smtp_error.decode(errors='ignore')}\n"
                "Your school has probably disabled SMTP login. Use --backend outlook instead."
            )

    def send(self, to, subject, body):
        msg = EmailMessage()
        msg["From"] = self.user
        msg["To"] = to
        msg["Subject"] = subject
        msg.set_content(body)
        self.server.send_message(msg)

    def close(self):
        self.server.quit()


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--contacts", default=HERE / "contacts.csv")
    p.add_argument("--template", default=HERE / "template.txt")
    p.add_argument("--log", default=HERE / "sent_log.csv", type=Path)
    p.add_argument("--backend", choices=["mail", "outlook", "smtp"],
                   default="mail" if sys.platform == "darwin" else "outlook")
    p.add_argument("--send", action="store_true", help="actually send (default is a dry run)")
    p.add_argument("--test", metavar="YOUR_EMAIL", help="send only the first contact's email, to yourself")
    p.add_argument("--limit", type=int, default=40, help="max emails this run (default 40)")
    p.add_argument("--delay", type=float, default=60, help="average seconds between emails (default 60)")
    p.add_argument("--sender", help="your email address (picks the account in Mail; required for SMTP)")
    p.add_argument("--rich", action="store_true",
                   help="Mail backend: send formatted (Helvetica 14, signature name in bold) instead of plain text")
    p.add_argument("--bold", default="Sean Alden Gani",
                   help="with --rich, text made bold where a line starts with it (default: Sean Alden Gani)")
    p.add_argument("--host", default="smtp.office365.com")
    p.add_argument("--port", type=int, default=587)
    args = p.parse_args()

    subject_t, body_t = load_template(args.template)
    contacts = load_contacts(args.contacts)
    sent = load_sent(args.log)
    todo = [c for c in contacts if c["email"].lower() not in sent]

    if args.test:
        todo = [dict(todo[0] if todo else contacts[0], email=args.test)]
    todo = todo[: args.limit]

    print(f"{len(contacts)} contacts, {len(sent)} already sent, {len(todo)} in this run\n")
    if not todo:
        return

    if not (args.send or args.test):
        for c in todo:
            print(f"--- To: {c['email']}\nSubject: {fill(subject_t, c, 'subject')}\n\n{fill(body_t, c, 'template')}")
        print("Dry run only. Add --test you@northeastern.edu to email yourself, or --send to send for real.")
        return

    backends = {"mail": MailSender, "outlook": OutlookSender, "smtp": SmtpSender}
    sender = backends[args.backend](args)
    try:
        for i, c in enumerate(todo, 1):
            subject, body = fill(subject_t, c, "subject"), fill(body_t, c, "template")
            try:
                sender.send(c["email"], subject, body)
            except Exception as e:
                print(f"[{i}/{len(todo)}] FAILED {c['email']}: {e}")
                continue
            if not args.test:
                log_sent(args.log, c["email"], subject)
            print(f"[{i}/{len(todo)}] sent to {c['email']}")
            if i < len(todo):
                # Random spacing looks like a person typing, not a blast.
                time.sleep(random.uniform(args.delay * 0.5, args.delay * 1.5))
    except KeyboardInterrupt:
        print("\nStopped. Re-run the same command to pick up where you left off.")
    finally:
        sender.close()


if __name__ == "__main__":
    main()
