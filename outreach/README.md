# Outreach sender

Sends one personalized email per contact from your own mailbox. No third-party service is involved.

## Setup

1. Put your list in `contacts.csv`. You need an `email` column. Any other column, such as `first_name` or `company`, can be used as `{first_name}` or `{company}` in the template.
2. Edit `template.txt`. The first line is `Subject: ...` and everything below it is the body.
3. **Mac:** add your northeastern.edu account to the Mail app (System Settings > Internet Accounts > Microsoft Exchange). The first time you send, macOS asks to let Terminal control Mail. Click OK.
   **Windows:** sign in to the Outlook desktop app and run `pip install pywin32`.

## Run it

```bash
python send_outreach.py                              # dry run: prints every email, sends nothing
python send_outreach.py --test you@northeastern.edu  # sends the first email to yourself only
python send_outreach.py --send                       # sends for real (40 per run, about 1 per minute)
```

Options:

- `--limit N` changes how many emails go out per run.
- `--delay S` sets the average number of seconds between emails.
- `--sender you@northeastern.edu` chooses which Mail account sends the emails.
- `--backend smtp` logs in with SMTP instead of using Outlook. Many Microsoft 365 schools block this, so use `--test` to check whether yours does.

Every sent address is written to `sent_log.csv`. Running the script again skips those addresses, so you can stop it with Ctrl+C and resume later.
