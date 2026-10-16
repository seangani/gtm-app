# Handoff: Terms.so email outreach tool

Context for the orchestrator session. It was written by the cloud "coding arm" session that built this tool on 2026-10-16.

## The goal

Sean (gani.s@northeastern.edu) is doing cold outreach for **Terms.so**. Terms is a contract tool for drafting, redlining, e-signing and storing contracts, SOWs and retainer renewals. Sean is partnered with Family Office, a design studio in NYC, on it.

- **Targets:** design studios, agencies, and other businesses that send and receive a lot of contracts.
- **The ask:** a 15-minute feedback call.

**Constraint:** Sean must send from his Northeastern school email, which is Microsoft 365, and can't use third-party sending tools. So the tool is homemade and sends through his own mailbox.

## About the user

- New to coding. Explain things step by step in plain language, with exact commands to copy and paste.
- Uses a **Mac** and Terminal. Normally reads email in the browser.
- The northeastern.edu account has been added to the **Apple Mail app** (System Settings > Internet Accounts > Microsoft Exchange). This works and was confirmed.

## Architecture

```
contacts.csv ──┐
               ├─> send_outreach.py ──> osascript (AppleScript) ──> Apple Mail.app ──> M365 mailbox
template.txt ──┘          │
                          └─> sent_log.csv (dedupe + resume)
```

**Why Apple Mail and not the other options:**
- **SMTP AUTH** is usually disabled on university M365 tenants, and it doesn't work with Duo MFA.
- **The Graph API** needs an Azure app registration plus admin consent, which the school is unlikely to grant.
- **The new Outlook for Mac** has almost no AppleScript support.
- **Apple Mail** signs in with the normal Microsoft OAuth flow and is fully scriptable.

## Code

The code is in repo `seangani/gtm-app`, branch `claude/hopeful-brahmagupta-4gyd1j`, folder `outreach/`.

| File | Purpose |
|---|---|
| `send_outreach.py` | The sender. Python 3 standard library only. |
| `template.txt` | The email. The first line must be `Subject: ...`. Placeholders use Python `str.format` with **single** braces, e.g. `{first_name}` and `{company}`. |
| `contacts.csv` | Sample data only (2 fake rows). The real list is kept on Sean's Mac and is **not committed**. |
| `sent_log.csv` | Created at runtime and gitignored. One row per sent email: `email,subject,sent_at`. |
| `README.md` | Short usage guide. |

### Behaviour of `send_outreach.py`

- **Dry run by default.** Nothing is sent without `--send` or `--test ADDRESS`.
- **`--test ADDRESS`** sends only the first contact's email, redirected to ADDRESS, and doesn't write to the log.
- **Sending backends** (`--backend`):
  - `mail` is the default on macOS. It drives Mail.app through `osascript`, passing the address, subject, body and sender as argv to an `on run` handler, so no escaping is needed.
  - `outlook` is the default on Windows. It uses win32com/pywin32 and is **untested**.
  - `smtp` defaults to smtp.office365.com:587 with STARTTLS. It's probably blocked at Northeastern.
- **`--sender ADDRESS`** picks which Mail.app account sends.
- **Throttling:**
  - `--limit` caps emails per run (default 40).
  - `--delay` sets the average gap in seconds (default 60). The actual gap is random, between 0.5× and 1.5× the delay.
- **Dedupe and resume:** contacts already in `sent_log.csv` are skipped, so Ctrl+C and re-running is safe.
- **CSV headers are normalized.** For example, `First Name` becomes `first_name` and `Email (Apollo)` becomes `email_apollo`. The email column is `email` or the first column starting with `email_`. Extra columns are ignored but can be used as placeholders.
- **Plain text only.** Mail.app's own signature isn't applied to scripted messages, so the signature is written into the template.

### The command Sean uses

```
cd <path to outreach folder>
python3 send_outreach.py                                                 # preview
python3 send_outreach.py --test gani.s@northeastern.edu --sender gani.s@northeastern.edu
caffeinate -i python3 send_outreach.py --send --sender gani.s@northeastern.edu
```

`caffeinate -i` keeps the Mac awake during a run.

## Current template (final, approved by Sean)

```
Subject: Quick question about {company}

Hi {first_name}, hope you're well!

I'm Sean, a student from Northeastern and I’m currently partnered with Family Office, a design studio in NYC. Family Office has recently developed a website called Terms.so to manage and draft contracts, SOWs and retainer renewals all in one place.

We're running a pilot on Terms with a small, hand-picked group of design studios, and I thought {company} would be a great fit to take advantage of the tool. There's nothing to buy, no strings attached - we're really just looking for honest feedback.

Would you be open to a quick 15-minute call next week to share how you handle contracts? I'd love to hear what works and what doesn't.

Thank you for your time and I look forward to hearing from you.

Best,
Sean

Sean Alden Gani | +1 781-815-0191 | gani.s@northeastern.edu | Terms.so
```

**A recurring mistake to watch for:** when Sean pastes new copy, it tends to start with `ubject:` instead of `Subject:`, and to say `Term.so` instead of `Terms.so`. Fix both each time.

Sean chose to drop the "reply no thanks" opt-out line. He was told that CAN-SPAM expects an opt-out and that the line helps deliverability.

## Current state

- **Test sends to himself worked.** Mail automation was confirmed working on his Mac.
- **First batch:** 25 US design-studio founders and owners from his enriched Apollo list.
  - The list is `terms-email-send-ready.csv`, with columns `first_name, last_name, email, company, title, website, opener_hook, time_zone`.
  - I cleaned it and returned it as `contacts.csv`:
    - Removed ", Inc." and ", LLC" from company names (Mance Creative, Think Designs, Kaleidoscopic).
    - Changed three first names: Jonathan Herrera → Jon, Michael Mance → Mike, Renee Dolan → Renée.
- **Whether that batch was sent is unconfirmed.** Sean said "it worked I think" after running `--send`. When he changed the signature and asked to "run the real thing" again, I told him to report the first output line:
  - `25 already sent, 0 in this run` means the first batch went out and nothing new will send.
  - `25 in this run` means the first run hadn't actually sent anything.
  - **Don't delete `sent_log.csv` to resend the same 25 people.**
- **Possible bounce:** Lina Campbell's address `lina@linanorelldesigns.com` doesn't match her site's domain, `norelldesign.com`.
- **Less-verified emails:** rows for Miriam Michel, Renée Dolan, Lina Campbell, Michael Habenicht and Mike Kotler were marked "LinkedIn only" in the source data.

## Ideas for next steps (not built)

1. **Follow-up mode:** a short bump 3–5 days later, sent only to people in `sent_log.csv` who haven't replied. It would need a second template, plus either a manual "replied" list or reading the Mail inbox with AppleScript.
2. **Use `{opener_hook}`:** the CSV already has a personalized first line for each contact. Sean hasn't chosen to use it.
3. **Bounce handling:** scan the inbox for "Undeliverable" messages and remove those addresses.
4. **Reply tracking:** the `Reply` column in his spreadsheet.

## Things to keep telling Sean

- Keep volume low, around 25–40 per day.
- Using school email for commercial outreach may break Northeastern's acceptable-use policy.
- Bounces and spam reports matter more than how fast emails are sent.
