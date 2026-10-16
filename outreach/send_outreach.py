#!/usr/bin/env python3
"""Send personalized outreach emails from your own mailbox, one at a time.

Reads contacts from a CSV, fills {placeholders} in a template, and sends each
email through either:

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
import smtplib
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


def load_contacts(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = [{k.strip(): (v or "").strip() for k, v in r.items()} for r in csv.DictReader(f)]
    if not rows or "email" not in rows[0]:
        sys.exit(f"{path}: needs a header row with at least an 'email' column")
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
    p.add_argument("--backend", choices=["outlook", "smtp"], default="outlook")
    p.add_argument("--send", action="store_true", help="actually send (default is a dry run)")
    p.add_argument("--test", metavar="YOUR_EMAIL", help="send only the first contact's email, to yourself")
    p.add_argument("--limit", type=int, default=40, help="max emails this run (default 40)")
    p.add_argument("--delay", type=float, default=60, help="average seconds between emails (default 60)")
    p.add_argument("--sender", help="SMTP: your email address")
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

    sender = (OutlookSender if args.backend == "outlook" else SmtpSender)(args)
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
