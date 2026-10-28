"""Enter the expense figures printed in the audited accounts for the year ended 30 Sep 2025 as comparatives:
    python manage.py load_fy2025_comparatives
Safe to run again; it replaces the figures it entered before."""
import datetime
from decimal import Decimal

from django.core.management.base import BaseCommand

from recon.models import FinancialYear, ManualEntry

PRINTED_2025 = [
    ("Transport", 39540), ("Stationery, Photocopying & typesetting", 31497), ("Speakers appreciation", 163500), ("Correspondences", 73385),
    ("Refreshments", 239526), ("Depreciation", 224924), ("AGM expenses", 34830), ("KUCU Care", 61500), ("FOCUS STEM Support", 66000),
    ("FOCUS Tithe", 129599), ("Repairs and Maintenance", 88555), ("Hospitality", 50660), ("Audit fees", 45000),
    ("Annual mission and mission support", 719500), ("Executive Programming", 55900), ("Gifts", 20100), ("Bank service charges", 24404),
    ("Training and Retreats", 123107), ("Computer & internet expenses", 4100), ("Chapel Project Remittances", 782000), ("Condolences", 10000)]
SPECIAL_EVENTS_2025 = 480097


class Command(BaseCommand):
    help = "Load the printed 2025 expense figures as comparatives."

    def handle(self, *a, **o):
        fy, _ = FinancialYear.objects.get_or_create(start=datetime.date(2024, 10, 1), end=datetime.date(2025, 9, 30), defaults={"name": "2024/25"})
        ManualEntry.objects.filter(period=fy, note="printed 2025 accounts").delete()
        for label, amount in PRINTED_2025:
            ManualEntry.objects.create(period=fy, year=2025, section="out_1_expenses", label=label, amount=Decimal(amount), note="printed 2025 accounts")
        ManualEntry.objects.create(period=fy, year=2025, section="out_2_special_events", label="Special events", amount=Decimal(SPECIAL_EVENTS_2025), note="printed 2025 accounts")
        total = sum(a for _, a in PRINTED_2025) + SPECIAL_EVENTS_2025
        self.stdout.write(f"Entered {len(PRINTED_2025) + 1} lines for {fy.label}; total {total:,} (printed total 3,467,724).")
