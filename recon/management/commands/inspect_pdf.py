"""Diagnose how a bank statement PDF is read:  python manage.py inspect_pdf statement.pdf [--password XXX]
Prints the first rows each reading method sees. Blank out names/numbers before sharing the output."""
import io

from django.core.files.uploadedfile import SimpleUploadedFile
from django.core.management.base import BaseCommand

from recon.services import StatementError, _parse_sheets, _pdf_sheets


class Command(BaseCommand):
    help = "Show how a statement PDF is read (ruled tables vs word positions)."

    def add_arguments(self, parser):
        parser.add_argument("path")
        parser.add_argument("--password", default="")
        parser.add_argument("--rows", type=int, default=12)

    def handle(self, *a, **o):
        data = open(o["path"], "rb").read()
        try:
            sheets = _pdf_sheets(data, o["password"])
        except StatementError as e:
            self.stdout.write(f"ERROR: {e}")
            return
        self.stdout.write(f"{len(sheets)} reading(s) found")
        for i, sh in enumerate(sheets, 1):
            rows, opening, closing, found = _parse_sheets([sh])
            self.stdout.write(f"\n=== Reading {i}: {len(sh)} raw rows | header found: {found} | transactions: {len(rows)} | opening {opening} closing {closing}")
            for r in sh[: o["rows"]]:
                self.stdout.write("  " + " | ".join((c or "")[:28].replace("\n", "\\n") for c in r))
