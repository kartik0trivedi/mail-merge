# Python-Based Mail Merge

Genesis: I couldn't access Office 365 due to a lack of institutional support.
Then I had to do a mail-merge and couldn't get LibreOffice working on Mac. So
whipped up this little thing to send emails using Python.

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

For Gmail, enable 2-Step Verification on the sending account, then create an
app password from your Google Account security settings. Use that 16-character
app password as `SMTP_PASSWORD` — not your normal Google password. Untested on
Outlook or other providers, but may work.

## Scaffold a new template

`--template` with **no filename** triggers scaffold mode — it reads the column
headers from the contact file and writes a starter `.md` file:

```bash
python mail_merge.py --template --contacts contacts.csv
```

This writes `contacts-template.md` (named after the contact file) with:

- YAML front matter (`subject`, `cc`, `bcc`) using the first two non-email
  columns as subject placeholders
- A greeting line using the first non-email column
- A reference list of every available `{{ placeholder }}`

The command refuses to overwrite an existing file; rename or remove it first.

## Use a specific template file

`--template filename.md` (with a filename) tells the script which template to
use. The default is `email-template.md` when the flag is omitted entirely.

After scaffolding, edit the generated file and then pass it back:

```bash
python mail_merge.py --template contacts-template.md --preview
python mail_merge.py --template contacts-template.md --send
```

Combine with `--contacts` to pair a custom template with a custom contact list:

```bash
python mail_merge.py --template outreach.md --contacts outreach-list.xlsx --preview
```

## Preview

Preview without sending (uses `email-template.md` by default):

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

## Data quality checks

Before sending (or previewing), the script validates every contact and flags:

- **Duplicate email addresses** — only the first occurrence would be sent to;
  later rows are flagged
- **Invalid email format** — no `@`, missing domain, etc.
- **Name fields** — checked automatically based on column names:
  - `First Name` + `Last Name` columns: empty values, digits, or identical
    first and last name (e.g. `John John`)
  - `Name` column: single-word entries, digits, or repeated first/last word

If issues are found, they are printed in a table and you are prompted:

```text
  1. Exit    — fix the contact file and re-run
  2. Continue — proceed anyway (duplicate/invalid emails will be skipped)
```

Choosing **Continue** automatically drops duplicate and malformed email
addresses; contacts with name warnings (e.g. a real name like "John John")
are still included. In non-interactive (piped) mode the script exits by
default.

## Audit trail

After every `--send` run the script writes two columns back into the contact
file, using the run date in the column name so multiple rounds on the same
list stay separate:

| Column | Value | Example |
| --- | --- | --- |
| `Sent - YYYY-MM-DD` | Time the run started | `14:32` |
| `Template - YYYY-MM-DD` | Template filename used | `outreach.md` |

Only contacts that were successfully sent get a value. Contacts skipped by
`--limit`, validation failures, or send errors are left blank, making it
easy to filter the spreadsheet for "not yet reached" rows.

A second run on the same day updates the time in the existing columns. A run
on a different day adds a new pair of columns, preserving the full send
history for every contact.

Supported for `.xlsx` and `.xlsm` files. `.csv` files are also supported.
`.ods` files are read correctly but write-back is not supported; a note is
printed after sending.

## Send

```bash
python mail_merge.py --send
```

The script loads `.env` automatically.

Start with a limited send first:

```bash
python mail_merge.py --send --limit 1
```

By default, the script waits 3 seconds between sent emails. Change it with
`--delay`:

```bash
python mail_merge.py --send --delay 5
```
