#!/usr/bin/env python3
"""Send a simple mail merge from contact-list.xlsx and a Markdown template."""

from __future__ import annotations

import argparse
import csv
import os
import re
import smtplib
import sys
import time
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook


REQUIRED_COLUMNS = ("Email",)
PLACEHOLDER_RE = re.compile(r"{{\s*([^{}]+?)\s*}}")


@dataclass(frozen=True)
class Contact:
    email: str
    fields: dict[str, str]


@dataclass(frozen=True)
class Template:
    subject: str
    cc: tuple[str, ...]
    bcc: tuple[str, ...]
    body: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Merge Excel columns into a Markdown email template.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Examples:
  python mail_merge.py
  python mail_merge.py --limit 1
  python mail_merge.py --template invite.md
  python mail_merge.py --send --limit 1
  python mail_merge.py --send --delay 5
  python mail_merge.py --template --contacts contacts.csv   (scaffold a new template)

Template front matter:
  ---
  subject: "Hello {{ First Name }} {{ Last Name }}"
  cc: "someone@example.com"
  bcc: "archive@example.com"
  ---

Any Excel column header can be used as a placeholder, for example:
  {{ First Name }}
  {{ Last Name }}
  {{ Email }}
""",
    )
    parser.add_argument(
        "--contacts",
        default="contact-list.xlsx",
        help="Path to the XLSX, ODS, or CSV contact list. Default: contact-list.xlsx",
    )
    parser.add_argument(
        "--template",
        nargs="?",
        default="email-template.md",
        const=None,
        help=(
            "Path to the Markdown email template (default: email-template.md). "
            "Pass the flag with no value to scaffold a new template from --contacts."
        ),
    )
    parser.add_argument(
        "--send",
        action="store_true",
        help="Actually send emails. Omit this for a dry run preview.",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Preview emails without sending. This is the default behavior.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Optional maximum number of contacts to process.",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=3.0,
        help="Seconds to wait between sent emails. Default: 3",
    )
    args = parser.parse_args()
    if args.send and args.preview:
        parser.error("--send and --preview cannot be used together")
    return args


def scaffold_template(contacts_path: Path) -> int:
    output_path = Path(contacts_path.stem + "-template.md")
    if output_path.exists():
        print(
            f"Error: {output_path} already exists. Remove it or rename it before scaffolding.",
            file=sys.stderr,
        )
        return 1

    contacts = load_contacts(contacts_path)
    if not contacts:
        print("No contacts found — cannot infer columns.", file=sys.stderr)
        return 1

    columns = list(contacts[0].fields.keys())
    non_email = [c for c in columns if c.lower() != "email"]

    subject_placeholders = " ".join(f"{{{{ {c} }}}}" for c in non_email[:2]) or "Your Subject"
    greeting_field = non_email[0] if non_email else columns[0]
    placeholders_inline = ", ".join(f"{{{{ {c} }}}}" for c in columns)

    content = (
        "---\n"
        f'subject: "{subject_placeholders}"\n'
        'cc: ""\n'
        'bcc: ""\n'
        "---\n"
        f"Hi {{{{ {greeting_field} }}}},\n"
        "\n"
        "Write your message here.\n"
        "\n"
        f"Available placeholders: {placeholders_inline}\n"
        "\n"
        "Regards,\n"
        "\n"
        "[Your Name]\n"
    )

    output_path.write_text(content, encoding="utf-8")
    print(f"Template scaffolded: {output_path}")
    print(f"Columns: {', '.join(columns)}")
    return 0


def load_contacts(path: Path) -> list[Contact]:
    if path.suffix.lower() == ".csv":
        return load_csv_contacts(path)
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        return load_excel_contacts(path)
    if path.suffix.lower() == ".ods":
        return load_ods_contacts(path)
    raise ValueError(f"{path} must be an .xlsx, .xlsm, .ods, or .csv file")


def load_excel_contacts(path: Path) -> list[Contact]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = sheet.iter_rows(values_only=True)

    try:
        headers = [str(value).strip() if value is not None else "" for value in next(rows)]
    except StopIteration:
        raise ValueError(f"{path} is empty") from None

    missing = [column for column in REQUIRED_COLUMNS if column not in headers]
    if missing:
        raise ValueError(f"{path} is missing required column(s): {', '.join(missing)}")

    email_index = headers.index("Email")
    contacts: list[Contact] = []

    for row_number, row in enumerate(rows, start=2):
        fields = {header: cell_text(row, index) for index, header in enumerate(headers)}
        email = cell_text(row, email_index)

        if not any(fields.values()):
            continue
        if not email:
            print(f"Skipping row {row_number}: missing Email", file=sys.stderr)
            continue

        contacts.append(Contact(email=email, fields=fields))

    return contacts


def load_csv_contacts(path: Path) -> list[Contact]:
    with path.open(newline="", encoding="utf-8-sig") as csv_file:
        reader = csv.DictReader(csv_file)
        if not reader.fieldnames:
            raise ValueError(f"{path} is empty")

        headers = [str(value).strip() if value is not None else "" for value in reader.fieldnames]
        missing = [column for column in REQUIRED_COLUMNS if column not in headers]
        if missing:
            raise ValueError(f"{path} is missing required column(s): {', '.join(missing)}")

        contacts: list[Contact] = []
        for row_number, row in enumerate(reader, start=2):
            fields = {
                header: str(row.get(header, "") or "").strip()
                for header in headers
            }
            email = fields["Email"]

            if not any(fields.values()):
                continue
            if not email:
                print(f"Skipping row {row_number}: missing Email", file=sys.stderr)
                continue

            contacts.append(Contact(email=email, fields=fields))

    return contacts


def load_ods_contacts(path: Path) -> list[Contact]:
    try:
        from odf.opendocument import load
        from odf.table import Table, TableCell, TableRow
        from odf.text import P
    except ImportError as exc:
        raise ValueError(
            "Reading .ods files requires odfpy. Run: pip install -r requirements.txt"
        ) from exc

    document = load(str(path))
    tables = document.spreadsheet.getElementsByType(Table)
    if not tables:
        raise ValueError(f"{path} does not contain any sheets")

    rows = ods_table_rows(tables[0], TableRow, TableCell, P)
    if not rows:
        raise ValueError(f"{path} is empty")

    headers = [value.strip() for value in rows[0]]
    missing = [column for column in REQUIRED_COLUMNS if column not in headers]
    if missing:
        raise ValueError(f"{path} is missing required column(s): {', '.join(missing)}")

    contacts: list[Contact] = []
    for row_number, row in enumerate(rows[1:], start=2):
        fields = {
            header: row[index].strip() if index < len(row) else ""
            for index, header in enumerate(headers)
        }
        email = fields["Email"]

        if not any(fields.values()):
            continue
        if not email:
            print(f"Skipping row {row_number}: missing Email", file=sys.stderr)
            continue

        contacts.append(Contact(email=email, fields=fields))

    return contacts


def ods_table_rows(table: object, table_row_type: object, table_cell_type: object, p_type: object) -> list[list[str]]:
    rows: list[list[str]] = []
    for row in table.getElementsByType(table_row_type):
        row_repeat = int(row.getAttribute("numberrowsrepeated") or 1)
        values: list[str] = []

        for cell in row.getElementsByType(table_cell_type):
            cell_repeat = int(cell.getAttribute("numbercolumnsrepeated") or 1)
            text = " ".join(
                node_text(paragraph).strip()
                for paragraph in cell.getElementsByType(p_type)
                if node_text(paragraph).strip()
            )
            values.extend([text] * cell_repeat)

        for _ in range(row_repeat):
            rows.append(values.copy())

    return rows


def node_text(node: object) -> str:
    text_parts: list[str] = []
    for child in getattr(node, "childNodes", []):
        if hasattr(child, "data"):
            text_parts.append(str(child.data))
        else:
            text_parts.append(node_text(child))
    return "".join(text_parts)


def cell_text(row: tuple[object, ...], index: int) -> str:
    if index >= len(row) or row[index] is None:
        return ""
    return str(row[index]).strip()


def load_template(path: Path) -> Template:
    text = path.read_text(encoding="utf-8")
    front_matter_values: dict[str, str] = {}
    body = text

    if text.startswith("---\n"):
        _, front_matter, body = text.split("---\n", 2)
        front_matter_values = parse_front_matter(front_matter)

    subject = front_matter_values.get("subject", "")
    if not subject:
        raise ValueError(
            f"{path} must start with front matter containing a subject, for example: "
            'subject: "Hello {{ First Name }}"'
        )

    return Template(
        subject=subject,
        cc=parse_recipients(front_matter_values.get("cc", "")),
        bcc=parse_recipients(front_matter_values.get("bcc", "")),
        body=body.strip() + "\n",
    )


def parse_front_matter(front_matter: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for line in front_matter.splitlines():
        key, sep, value = line.partition(":")
        if sep:
            values[key.strip().lower()] = value.strip().strip("\"'")
    return values


def parse_recipients(value: str) -> tuple[str, ...]:
    if not value:
        return ()
    return tuple(item.strip() for item in value.split(",") if item.strip())


def render(text: str, contact: Contact) -> str:
    def replace(match: re.Match[str]) -> str:
        field_name = match.group(1).strip()
        if field_name not in contact.fields:
            raise ValueError(
                f"Template uses {{{{ {field_name} }}}}, but contact-list.xlsx "
                f"does not have a '{field_name}' column"
            )
        return contact.fields[field_name]

    return PLACEHOLDER_RE.sub(replace, text)


def preview(
    contact: Contact,
    subject: str,
    cc: tuple[str, ...],
    bcc: tuple[str, ...],
    body: str,
) -> None:
    print("=" * 72)
    print(f"To: {contact.email}")
    if cc:
        print(f"Cc: {', '.join(cc)}")
    if bcc:
        print(f"Bcc: {', '.join(bcc)}")
    print(f"Subject: {subject}")
    print("-" * 72)
    print(body, end="" if body.endswith("\n") else "\n")


def send_email(
    contact: Contact,
    subject: str,
    cc: tuple[str, ...],
    bcc: tuple[str, ...],
    body: str,
) -> None:
    host = required_env("SMTP_HOST")
    port = int(os.getenv("SMTP_PORT", "587"))
    username = required_env("SMTP_USER")
    password = required_env("SMTP_PASSWORD")
    from_email = os.getenv("FROM_EMAIL", username).strip() or username
    from_name = os.getenv("FROM_NAME", "").strip()
    sender = f"{from_name} <{from_email}>" if from_name else from_email

    message = EmailMessage()
    message["From"] = sender
    message["To"] = contact.email
    if cc:
        message["Cc"] = ", ".join(cc)
    if bcc:
        message["Bcc"] = ", ".join(bcc)
    message["Subject"] = subject
    message.set_content(body)

    with smtplib.SMTP(host, port) as smtp:
        if os.getenv("SMTP_STARTTLS", "true").lower() not in {"0", "false", "no"}:
            smtp.starttls()
        smtp.login(username, password)
        smtp.send_message(message)


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def limited_contacts(contacts: list[Contact], limit: int | None) -> Iterable[Contact]:
    return contacts[:limit] if limit is not None else contacts


def main() -> int:
    args = parse_args()

    if args.template is None:
        return scaffold_template(Path(args.contacts))

    contacts = load_contacts(Path(args.contacts))
    template = load_template(Path(args.template))

    if not contacts:
        print("No contacts found.")
        return 0

    processed = 0
    selected_contacts = list(limited_contacts(contacts, args.limit))

    for index, contact in enumerate(selected_contacts, start=1):
        subject = render(template.subject, contact)
        cc = tuple(render(value, contact) for value in template.cc)
        bcc = tuple(render(value, contact) for value in template.bcc)
        body = render(template.body, contact)

        if args.send:
            send_email(contact, subject, cc, bcc, body)
            print(f"Sent: {contact.email}")
            if args.delay > 0 and index < len(selected_contacts):
                print(f"Waiting {args.delay:g}s before next email...")
                time.sleep(args.delay)
        else:
            preview(contact, subject, cc, bcc, body)

        processed += 1

    mode = "sent" if args.send else "previewed"
    print(f"\nDone: {mode} {processed} email(s).")
    if not args.send:
        print("Dry run only. Add --send after Gmail SMTP is configured.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
