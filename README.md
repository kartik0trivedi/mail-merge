# Python-Based Mail Merge

Genesis: I couldn't access Office 365 due to a lack of institutional support for Office 365. Then, as it happened, I had to do a mail-merge (which I hadn't done in years). Tried to walk through LibreOffice, but couldn't get it to work on Mac (most likely user error) to send emails. So whipped up this little thing to send emails using Python+Jinja.

This setup sends a simple Markdown-based mail merge from `contact-list.xlsx`
through Gmail SMTP using a Google app password.

The contact file can be `.xlsx`, `.xlsm`, `.ods`, or `.csv`. It must include
this column:

- `Email`

The sample template is `email-template.md`. It can use any column from the
contact file as a merge field:

- `{{ First Name }}`
- `{{ Last Name }}`
- `{{ Email }}`

The placeholder name must match the Excel column header exactly.

The Markdown template front matter supports `subject`, `cc`, and `bcc`:

```md
---
subject: "Hello {{ First Name }} {{ Last Name }}"
cc: "someone@example.com"
bcc: "audit@example.com, archive@example.com"
---
```

Leave `cc` or `bcc` blank when you do not need them.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Copy `.env.example` to `.env` and fill in your Gmail SMTP details.

```bash
cp .env.example .env
```

For Gmail, enable 2-Step Verification on the sending Google account, then create an app password from your Google Account security settings. Use that 16-character app password as `SMTP_PASSWORD`. Do not use your normal Google account password. I didn't test it on Outlook or any other provider. Could work. Or you can use this code and vibe it out.

## Scaffold a new template

Generate a starter template pre-filled with placeholders drawn from a contact file:

```bash
python mail_merge.py --template --contacts contacts.csv
```

This reads the column headers from the contact file and writes a new
`contacts-template.md` (named after the contact file) with:

- YAML front matter (`subject`, `cc`, `bcc`) using the first two non-email
  columns as subject placeholders
- A greeting line using the first non-email column
- A reference list of every available `{{ placeholder }}`

Edit the generated file, then pass it back with
`--template contacts-template.md` to preview or send.

The command refuses to overwrite an existing file; rename or remove it first.

## Preview

Preview without sending:

```bash
python mail_merge.py --preview
```

Preview is the default, so `python mail_merge.py` also previews without sending.

Preview only the first email:

```bash
python mail_merge.py --preview --limit 1
```

Use a CSV contact list:

```bash
python mail_merge.py --contacts contact-list.csv --preview
```

Use an OpenDocument spreadsheet:

```bash
python mail_merge.py --contacts contact-list.ods --preview
```

## Send

Load the environment variables, then send:

```bash
set -a
source .env
set +a
python mail_merge.py --send
```

Start with a limited send first:

```bash
python mail_merge.py --send --limit 1
```

By default, the script waits 3 seconds between sent emails. You can change that with `--delay`:

```bash
python mail_merge.py --send --delay 5
```
