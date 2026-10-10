"""Statement parsing and categorisation (replaces the Excel SUMPRODUCT/SEARCH formulas)."""
import csv
import io
import re
from difflib import SequenceMatcher
from functools import lru_cache
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

import openpyxl

from .models import Category, Transaction


class StatementError(Exception):
    pass


def to_dec(v):
    if v is None or v == "":
        return Decimal("0")
    if isinstance(v, (int, float)):
        return Decimal(str(v)).quantize(Decimal("0.01"))
    try:
        return Decimal(re.sub(r"[,\s]", "", str(v))).quantize(Decimal("0.01"))
    except InvalidOperation:
        return Decimal("0")


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(v).strip()[:10], fmt).date()
        except (ValueError, TypeError):
            pass
    return None


def _sheets(upload):
    name = upload.name.lower()
    data = upload.read()
    if name.endswith(".csv"):
        text = data.decode("utf-8-sig", errors="replace")
        yield [list(r) for r in csv.reader(io.StringIO(text))]
    elif name.endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        for ws in wb.worksheets:
            yield [list(r) for r in ws.iter_rows(values_only=True)]
    else:
        raise StatementError("Please upload an .xlsx or .csv file (save .xls files as .xlsx first).")


def parse_statement(upload):
    """Return (rows, opening, closing). rows: dicts of date/description/debit/credit."""
    rows, opening, closing, found = [], None, None, False
    for sheet in _sheets(upload):
        hi = next((i for i, r in enumerate(sheet)
                   if any(re.search(r"details|description|narrat|particular", str(c or ""), re.I) for c in r)
                   and any(re.search(r"credit", str(c or ""), re.I) for c in r)), None)
        if hi is None:
            continue
        head = [str(c or "").lower() for c in sheet[hi]]
        col = lambda rx: next((i for i, h in enumerate(head) if re.search(rx, h)), None)
        c_desc, c_deb, c_cr = col(r"details|description|narrat|particular"), col(r"debit"), col(r"credit")
        c_dt, c_bal = col(r"transaction date|^date"), col(r"balance")
        found = True
        get = lambda r, i: r[i] if i is not None and i < len(r) else None
        for r in sheet[hi + 1:]:
            desc = re.sub(r"\s+", " ", str(get(r, c_desc) or "")).strip()
            if not desc or re.match(r"total", desc, re.I):
                continue
            if re.search(r"b/?fwd|brought forward|opening balance", desc, re.I):
                bal = to_dec(get(r, c_bal)) or to_dec(get(r, c_cr))
                opening = bal
                continue
            cr, db = to_dec(get(r, c_cr)), to_dec(get(r, c_deb))
            if not cr and not db:
                continue
            if c_bal is not None and get(r, c_bal) not in (None, ""):
                closing = to_dec(get(r, c_bal))
            rows.append({"date": to_date(get(r, c_dt)), "description": desc, "debit": db, "credit": cr})
    if not found:
        raise StatementError("Could not find Transaction Details / Debit / Credit columns in this file.")
    if not rows:
        raise StatementError("No transactions were found in this file.")
    return rows, opening, closing


# ---- Smart matching -------------------------------------------------------
# A keyword matches a word in the description when it is: identical; the start of the
# word (plurals: offerings, tithes); a close misspelling (ofering, titthe); or two
# glued words (TitheOfferings). When a payment matches several categories the lowest
# `order` wins, so "Tithe/Offering" counts as Offering.

@lru_cache(maxsize=100000)
def _sim(a, b):
    return SequenceMatcher(None, a, b).ratio()


def _words(description):
    return [t for t in re.findall(r"[a-z0-9]+", description.lower()) if t.isalpha()]


def _candidates(words):
    out = [(w, w) for w in words]
    out += [(a + b, f"{a} {b}") for a, b in zip(words, words[1:])]
    return out


def _kw_match(kw, words, cands):
    """Return None, or (kind, original text) where kind is exact | fuzzy."""
    for c, orig in cands:
        if c == kw:
            return "exact", orig
    if len(kw) < 4:  # short keywords like BS or SLS must match exactly
        return None
    for c, orig in cands:
        if c.startswith(kw):
            return "exact", orig
    for c, orig in cands:
        if abs(len(c) - len(kw)) <= 2 and _sim(kw, c) >= 0.8:
            return "fuzzy", orig
    for w in words:  # glued words such as "titheofferings"
        for i in range(4, len(w) - 3):
            if w[:i].startswith(kw) or w[i:].startswith(kw):
                return "fuzzy", w
    return None


def _compile(categories):
    return [(c, [k.lower().replace(" ", "") for k in c.keyword_list()]) for c in categories]


def find_category(description, compiled):
    """Return (category or None, note)."""
    words = _words(description)
    cands = _candidates(words)
    hits, notes = [], []
    for cat, kws in compiled:
        for kw in kws:
            m = _kw_match(kw, words, cands)
            if m:
                hits.append(cat)
                if m[0] == "fuzzy":
                    notes.append(f"\u201c{m[1]}\u201d read as {cat.name}")
                break
    if not hits:
        return None, ""
    note = "; ".join(dict.fromkeys(notes))
    if len(hits) > 1:
        extra = f"Matched {' and '.join(c.name for c in hits)}; counted as {hits[0].name}"
        note = f"{note}; {extra}" if note else extra
    return hits[0], note[:200]


def categorise(statement):
    """Apply the keyword rules to every transaction whose category was not set by hand."""
    cats = list(Category.objects.filter(is_active=True))
    income = _compile([c for c in cats if c.is_income])
    expense = _compile([c for c in cats if not c.is_income])
    changed = []
    for t in statement.transactions.filter(manual=False):
        cat, note = find_category(t.description, income if t.credit else expense)
        if t.category_id != (cat.id if cat else None) or t.note != note:
            t.category, t.note = cat, note
            changed.append(t)
    Transaction.objects.bulk_update(changed, ["category", "note"])


def import_statement(statement, rows):
    Transaction.objects.bulk_create([Transaction(statement=statement, **r) for r in rows])
    categorise(statement)
