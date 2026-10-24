"""Reads the treasurer's budget workbook (monthly Actuals / Budget / Difference per expense line)."""
import re
from decimal import Decimal

import openpyxl

from .models import BudgetLine, Category, ItemMap
from .services import StatementError, _compile, find_category, to_dec

MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
STOP = re.compile(r"monthly totals?|total expenditure|grand total|^totals?$", re.I)


def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()[:120]


def _month_of(label, year):
    if label is None:
        return None
    t = str(label).strip().lower()
    if t[:3] in MONTHS and not t.startswith("ytd"):
        return year, MONTHS[t[:3]]
    m = re.fullmatch(r"(20\d\d)(\d\d)", t)  # period codes such as 202510
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))
    return None


def parse_budget(upload, year=None):
    try:
        wb = openpyxl.load_workbook(upload, data_only=True)
    except Exception:
        raise StatementError("This file could not be read as an Excel workbook.")
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400), values_only=True))
        hr = next((i for i, r in enumerate(rows[:12]) if any(str(c or "").strip().lower() == "actuals" for c in r)
                   and any(str(c or "").strip().lower() == "budget" for c in r)), None)
        if hr is None:
            continue
        y = year or next((int(m) for m in re.findall(r"(20\d\d)", ws.title)), None) or __import__("datetime").date.today().year
        head = rows[hr]
        blocks, ignored = [], []
        for c, v in enumerate(head):
            if str(v or "").strip().lower() != "actuals" or c + 1 >= len(head):
                continue
            label = next((rows[r][c] for r in range(hr - 1, -1, -1) if rows[r][c] not in (None, "")), None)
            ym = _month_of(label, y)
            col = openpyxl.utils.get_column_letter(c + 1)
            if ym:
                blocks.append({"col": c, "year": ym[0], "month": ym[1], "label": str(label).strip(), "letter": col})
            else:
                ignored.append(f"column {col} ({str(label).strip() if label else 'no month label'})")
        if not blocks:
            continue
        lab = next((c for r in rows[:hr + 1] for c, v in enumerate(r) if "expenditure" in str(v or "").lower()), 2)
        lines, group, want_header = [], "", True
        for r in rows[hr + 1:]:
            name = re.sub(r"\s+", " ", str(r[lab] or "")).strip() if lab < len(r) else ""
            if not name or name.lower() in ("actuals", "budget"):
                continue
            if STOP.search(name):
                break
            if name.lower().startswith("subtotal"):
                want_header = True
                continue
            if want_header:
                group, want_header = name, False
                continue
            for b in blocks:
                a = to_dec(r[b["col"]]) if b["col"] < len(r) else Decimal(0)
                bu = to_dec(r[b["col"] + 1]) if b["col"] + 1 < len(r) else Decimal(0)
                if a or bu:
                    lines.append({"year": b["year"], "month": b["month"], "group": group, "item": name, "actual": a, "budget": bu})
        for b in blocks:
            mine = [l for l in lines if (l["year"], l["month"]) == (b["year"], b["month"])]
            b["actual"] = sum((l["actual"] for l in mine), Decimal(0))
            b["budget"] = sum((l["budget"] for l in mine), Decimal(0))
        return {"sheet": ws.title, "blocks": blocks, "lines": lines, "ignored": ignored}
    raise StatementError("Could not find the Actuals / Budget columns in this workbook.")


def import_budget(parsed):
    """Store the months that have actual spending; re-uploading a month replaces it."""
    expense = _compile([c for c in Category.objects.filter(is_active=True) if not c.is_income])
    saved = {i.key: i.category for i in ItemMap.objects.select_related("category")}
    months = sorted({(b["year"], b["month"]) for b in parsed["blocks"] if b["actual"] > 0})
    done = set(months)
    cache = {}
    out = []
    for l in parsed["lines"]:
        if (l["year"], l["month"]) not in done:
            continue
        cat = saved.get(norm(l["item"]))
        if cat is None:
            cat, _ = find_category(l["item"], expense, True, cache)
        out.append(BudgetLine(category=cat, **l))
    for y, m in months:
        BudgetLine.objects.filter(year=y, month=m).delete()
    BudgetLine.objects.bulk_create(out, batch_size=500)
    return months
