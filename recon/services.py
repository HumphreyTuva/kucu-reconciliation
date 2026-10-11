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
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d %b %Y", "%d-%b-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(str(v).strip()[:11].strip(), fmt).date()
        except (ValueError, TypeError):
            pass
    return None


def _sheets(upload, password=""):
    name = upload.name.lower()
    data = upload.read()
    if name.endswith(".csv"):
        text = data.decode("utf-8-sig", errors="replace")
        yield [list(r) for r in csv.reader(io.StringIO(text))]
    elif name.endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        for ws in wb.worksheets:
            yield [list(r) for r in ws.iter_rows(values_only=True)]
    elif name.endswith(".pdf"):
        yield _pdf_rows(data, password)
    else:
        raise StatementError("Please upload an .xlsx, .csv or .pdf file (save .xls files as .xlsx first).")



HEAD_DESC = r"details|description|narrat|particular"
HEAD_DATE = r"transaction date|trans\.? ?date|posting date|^date"


def _pdf_rows(data, password):
    """Read the tables of a (possibly password-protected) bank statement PDF into sheet-style rows."""
    try:
        import pdfplumber
    except ImportError:
        raise StatementError("PDF support is not installed on the server (pip install pdfplumber cryptography).")
    try:
        pdf = pdfplumber.open(io.BytesIO(data), password=password or "")
    except Exception as e:
        inner = e.args[0] if e.args else e  # pdfplumber wraps pdfminer's error
        if type(inner).__name__ == "PDFPasswordIncorrect" or type(e).__name__ == "PDFPasswordIncorrect":
            raise StatementError("The PDF password was not accepted." if password else
                                 "This PDF is password-protected. Enter its password in the password box and upload again.")
        raise StatementError("This PDF could not be read. Is it a scanned image or a damaged file?")
    rows = []
    with pdf:
        for page in pdf.pages:
            for t in page.extract_tables():  # tables drawn with ruled lines
                rows.extend([[(c or "").strip() for c in r] for r in t])
        if not any(_is_header(r) for r in rows):  # no ruled lines: rebuild the table from word positions
            rows = _words_rows(pdf)
    return _merge_wrapped(rows)


def _is_header(r):
    return any(re.search(HEAD_DESC, c, re.I) for c in r) and any(re.search("credit", c, re.I) for c in r)


NUM = re.compile(r"^-?[\d,]+\.\d{2}$")
DATE = re.compile(r"^\d{1,2}[-/ ][A-Za-z0-9]{1,3}[-/ ]\d{2,4}$")
AMT_HEAD = re.compile("debit|credit|balance", re.I)


def _lines(words):
    out = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if out and abs(out[-1][0]["top"] - w["top"]) <= 3:
            out[-1].append(w)
        else:
            out.append([w])
    return [sorted(l, key=lambda w: w["x0"]) for l in out]


def _words_rows(pdf):
    """Fallback for PDFs without ruled lines: place each word in a column using the header's x-positions,
    start a transaction at every line that begins with a date, and attach wrapped lines to the nearest one."""
    cols, rows = None, []
    for page in pdf.pages:
        lines = _lines(page.extract_words())
        top = -1
        if cols is None:
            for l in lines:
                t = [w["text"].lower() for w in l]
                if "debit" in t and "credit" in t:
                    cols = []
                    for w in l:
                        if cols and w["x0"] - cols[-1]["x1"] <= 9:
                            cols[-1].update(text=cols[-1]["text"] + " " + w["text"], x1=w["x1"])
                        else:
                            cols.append({"text": w["text"], "x0": w["x0"], "x1": w["x1"]})
                    rows.append([c["text"] for c in cols])
                    top = l[0]["top"]
                    break
            if cols is None:
                continue

        def col_of(w):
            if NUM.match(w["text"]):
                nums = [i for i, c in enumerate(cols) if AMT_HEAD.search(c["text"])]
                if nums:
                    return min(nums, key=lambda i: min(abs(w["x1"] - cols[i]["x1"]),
                                                       abs((w["x0"] + w["x1"] - cols[i]["x0"] - cols[i]["x1"]) / 2)))
            idx = 0
            for i, c in enumerate(cols):
                if w["x0"] >= c["x0"] - 6:
                    idx = i
            return idx

        anchors = []
        loose = []
        for l in lines:
            t = [w["text"].lower() for w in l]
            if l[0]["top"] <= top or ("debit" in t and "credit" in t):
                continue
            if DATE.match(l[0]["text"]):
                anchors.append({"top": l[0]["top"], "words": list(l)})
            else:
                loose.append(l)
        for l in loose:
            near = min(anchors, key=lambda a: abs(a["top"] - l[0]["top"]), default=None)
            if near is not None and abs(near["top"] - l[0]["top"]) <= 22:
                near["words"].extend(l)
        for a in anchors:
            cells = [[] for _ in cols]
            for w in sorted(a["words"], key=lambda w: (round(w["top"] / 3), w["x0"])):
                cells[col_of(w)].append(w["text"])
            rows.append([" ".join(c) for c in cells])
    return rows


def _merge_wrapped(rows):
    """Long descriptions wrap onto extra lines in PDFs; glue those lines back onto the transaction above."""
    hi = next((i for i, r in enumerate(rows) if any(re.search(HEAD_DESC, c, re.I) for c in r)
               and any(re.search("credit", c, re.I) for c in r)), None)
    if hi is None:
        return rows
    head = [c.lower() for c in rows[hi]]
    col = lambda rx: next((i for i, h in enumerate(head) if re.search(rx, h)), None)
    c_desc, c_deb, c_cr, c_dt = col(HEAD_DESC), col("debit"), col("credit"), col(HEAD_DATE)
    out, prev = rows[:hi + 1], None
    cell = lambda r, i: r[i] if i is not None and i < len(r) else ""
    for r in rows[hi + 1:]:
        if not cell(r, c_dt) and not cell(r, c_deb) and not cell(r, c_cr) and cell(r, c_desc) and prev is not None:
            prev[c_desc] = f"{prev[c_desc]} {cell(r, c_desc)}".strip()
            continue
        out.append(r)
        prev = r
    return out


def parse_statement(upload, password=""):
    """Return (rows, opening, closing). rows: dicts of date/description/debit/credit."""
    rows, opening, closing, found = [], None, None, False
    for sheet in _sheets(upload, password):
        hi = next((i for i, r in enumerate(sheet)
                   if any(re.search(HEAD_DESC, str(c or ""), re.I) for c in r)
                   and any(re.search(r"credit", str(c or ""), re.I) for c in r)), None)
        if hi is None:
            continue
        head = [str(c or "").lower() for c in sheet[hi]]
        col = lambda rx: next((i for i, h in enumerate(head) if re.search(rx, h)), None)
        c_desc, c_deb, c_cr = col(HEAD_DESC), col(r"debit"), col(r"credit")
        c_dt, c_bal = col(HEAD_DATE), col(r"balance")
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
