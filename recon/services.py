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


OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # password-protected .xlsx files are stored in this container


def _open_protected_office(data, password):
    if not data.startswith(OLE_MAGIC):
        return data  # not encrypted (sheet protection, which only blocks editing, opens fine)
    try:
        import msoffcrypto
    except ImportError:
        raise StatementError("Protected Excel support is not installed on the server (pip install msoffcrypto-tool).")
    if not password:
        raise StatementError("This Excel file is password-protected. Enter its password in the password box and upload again.")
    try:
        office = msoffcrypto.OfficeFile(io.BytesIO(data))
        office.load_key(password=password)
        out = io.BytesIO()
        office.decrypt(out)
        return out.getvalue()
    except Exception:
        raise StatementError("The Excel password was not accepted.")


def _sheets(name, data, password=""):
    if name.endswith(".csv"):
        text = data.decode("utf-8-sig", errors="replace")
        yield [list(r) for r in csv.reader(io.StringIO(text))]
    elif name.endswith((".xlsx", ".xlsm")):
        data = _open_protected_office(data, password)
        try:
            wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        except Exception:
            raise StatementError("This Excel file could not be read. Is it damaged, or an old .xls saved with the wrong extension?")
        for ws in wb.worksheets:
            yield [list(r) for r in ws.iter_rows(values_only=True)]
    else:
        raise StatementError("Please upload an .xlsx, .csv or .pdf file (save .xls files as .xlsx first).")



HEAD_DESC = r"details|description|narrat|particular"
HEAD_DATE = r"transaction date|trans\.? ?date|posting date|^date"


def _pdf_sheets(data, password):
    """Yield readings of a (possibly password-protected) bank statement PDF: first from word positions (fast and
    exact for wrapped descriptions), then from ruled tables. parse_statement stops at the first that reconciles."""
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
    with pdf:
        words = _words_rows(pdf)
        if words:
            yield words
        tables = []
        for page in pdf.pages:
            for t in page.extract_tables():  # tables drawn with ruled lines
                tables.extend([[(c or "").strip() for c in r] for r in t])
        if tables:
            yield tables


def _is_header(r):
    return any(re.search(HEAD_DESC, c, re.I) for c in r) and any(re.search("credit", c, re.I) for c in r)


NUM = re.compile(r"^-?[\d,]+\.\d{2}$")
DATE = re.compile(r"\b\d{1,2}[-/. ](?:\d{1,2}|[A-Za-z]{3,9})[-/. ]\d{2,4}\b|\b\d{4}-\d{2}-\d{2}\b")
AMT_HEAD = re.compile("debit|credit|balance", re.I)


def _lines(words):
    out = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if out and abs(out[-1][0]["top"] - w["top"]) <= 3:
            out[-1].append(w)
        else:
            out.append([w])
    return [sorted(l, key=lambda w: w["x0"]) for l in out]


def _is_head_line(l):
    t = [w["text"].lower() for w in l]
    return "debit" in t and "credit" in t


def _words_rows(pdf):
    """Rebuild the transaction table from word positions: place each word in a column using the header's
    x-positions, start a transaction at every line that begins with a date, and attach the wrapped
    description lines. The layout is learned from the statement itself:
      * "below": the first row sits right under the header, so wrapped lines lie under their dated line
        (they continue the transaction above, even across a page break);
      * "nearest": wrapped lines sit around the dated line (vertically centred cells)."""
    cols, anchors, policy, last = None, [], None, None
    for page in pdf.pages:
        lines = _lines(page.extract_words())
        hi = next((k for k, l in enumerate(lines) if _is_head_line(l)), None)
        if cols is None:
            if hi is None:
                continue
            l = lines[hi]
            cols = []
            for w in l:
                if cols and w["x0"] - cols[-1]["x1"] <= 9:
                    cols[-1].update(text=cols[-1]["text"] + " " + w["text"], x1=w["x1"])
                else:
                    cols.append({"text": w["text"], "x0": w["x0"], "x1": w["x1"]})
            if not any(re.search(HEAD_DESC, c["text"], re.I) for c in cols):  # header wrapped onto two lines
                for w in (w for pl in lines for w in pl if l[0]["top"] - 16 <= w["top"] < l[0]["top"] - 1):
                    if re.fullmatch(HEAD_DESC, w["text"].lower()):
                        cols.append({"text": w["text"], "x0": w["x0"], "x1": w["x1"]})
                cols.sort(key=lambda c: c["x0"])

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

        dcol = next((i for i, c in enumerate(cols) if re.search(HEAD_DATE, c["text"], re.I)), 0)
        is_anchor = lambda l: bool(DATE.search(" ".join(w["text"] for w in l if col_of(w) == dcol)))
        body = [l for k, l in enumerate(lines) if (hi is None or k > hi) and not _is_head_line(l)]
        head_bottom = lines[hi][0]["top"] if hi is not None else None
        while body and head_bottom is not None and all(re.fullmatch(r"(date|dates|value|transaction|ref|no\.?|\(?kes\)?|amount)", w["text"].lower()) for w in body[0]):
            head_bottom = body.pop(0)[0]["top"]  # second line of a wrapped header, e.g. "Date  Date"
        if policy is None:
            first = next((k for k, l in enumerate(body) if is_anchor(l)), None)
            if first is None:
                continue
            policy = "below" if first == 0 else "nearest"
        if policy == "below":
            cur = last
            prev_top = head_bottom
            for l in body:
                if is_anchor(l):
                    last = cur = {"top": l[0]["top"], "words": list(l)}
                    anchors.append(cur)
                    prev_top = l[0]["top"]
                elif (cur is not None and prev_top is not None and l[0]["top"] - prev_top <= 16
                      and not any(NUM.match(w["text"]) and AMT_HEAD.search(cols[col_of(w)]["text"]) for w in l)
                      and not re.match(r"total|cleared|effects|book balance|kindly", l[0]["text"], re.I)):
                    cur["words"].extend(l)  # still part of the transaction above
                    prev_top = l[0]["top"]
                else:
                    cur = None  # a gap: footer or page text, not part of any transaction
        else:
            page_anchors = [{"top": l[0]["top"], "words": list(l)} for l in body if is_anchor(l)]
            for l in body:
                if is_anchor(l):
                    continue
                near = min(page_anchors, key=lambda a: abs(a["top"] - l[0]["top"]), default=None)
                if near is not None and abs(near["top"] - l[0]["top"]) <= 22:
                    near["words"].extend(l)
            anchors.extend(page_anchors)
    if cols is None:
        return []
    rows = [[c["text"] for c in cols]]
    for a in anchors:
        cells = [[] for _ in cols]
        for w in sorted(a["words"], key=lambda w: (round(w["top"] / 3), w["x0"])):
            # col_of needs the page-independent header positions, which are fixed for the whole statement
            idx = 0
            if NUM.match(w["text"]):
                nums = [i for i, c in enumerate(cols) if AMT_HEAD.search(c["text"])]
                idx = min(nums, key=lambda i: min(abs(w["x1"] - cols[i]["x1"]),
                                                  abs((w["x0"] + w["x1"] - cols[i]["x0"] - cols[i]["x1"]) / 2))) if nums else 0
            else:
                for i, c in enumerate(cols):
                    if w["x0"] >= c["x0"] - 6:
                        idx = i
            cells[idx].append(w["text"])
        rows.append([" ".join(c) for c in cells])
    return rows


PENNY = Decimal("0.01")


def _parse_sheets(sheets):
    """Read transactions from sheet-style rows.

    * Money in or out is decided from the running balance, not from which column the amount sits in:
      converters such as iLovePDF sometimes put credits under Debit on alternate pages, but the
      balance column stays right.
    * Rows with only a description are wrapped text; they are joined to the transaction above
      (text at the very top of a page belongs to the last transaction of the page before).
    """
    rows, opening, closing, found, prev_bal = [], None, None, False, None
    for sheet in sheets:
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
        first_on_page = True
        for r in sheet[hi + 1:]:
            if not any(str(c or "").strip() for c in r):  # blank row = page break
                first_on_page = True
                continue
            raw_desc = str(get(r, c_desc) or "").strip()
            desc = re.sub(r"\s+", " ", raw_desc).strip()
            if not desc or re.match(r"total|cleared balance|effects not cleared|(book |ledger |available )?balance as at|kindly|page \d", desc, re.I):
                continue
            raw_bal = get(r, c_bal)
            has_bal = raw_bal not in (None, "") and re.search(r"\d", str(raw_bal)) is not None
            bal = to_dec(raw_bal) if has_bal else None
            if re.search(r"b/?fwd|brought forward|opening balance", desc, re.I):
                opening = bal if bal is not None else to_dec(get(r, c_cr))
                prev_bal = opening
                continue
            cr, db = to_dec(get(r, c_cr)), to_dec(get(r, c_deb))
            if not cr and not db:
                if not has_bal and rows and not to_date(get(r, 0)) and not str(get(r, 0) or "").strip():
                    rows[-1]["description"] += " " + desc  # wrapped text of the transaction above
                continue
            lines = [l.strip() for l in raw_desc.splitlines() if l.strip()]
            if first_on_page and rows and len(lines) > 1:
                # the last line is this row's own first line; earlier ones spilled over from the page before
                k = next((i for i, l in enumerate(lines) if re.match(r"transfer in from acc", l, re.I)), 0)
                if k > 0:
                    rows[-1]["description"] += " " + " ".join(lines[:k])
                    desc = re.sub(r"\s+", " ", " ".join(lines[k:])).strip()
            first_on_page = False
            amount = cr or db
            if bal is not None and prev_bal is not None:
                delta = bal - prev_bal
                if abs(abs(delta) - amount) <= PENNY:  # direction from the balance, whatever column it was in
                    cr, db = (amount, Decimal("0")) if delta > 0 else (Decimal("0"), amount)
            if bal is not None:
                prev_bal = closing = bal
            rows.append({"date": to_date(get(r, c_dt)), "description": desc, "debit": db, "credit": cr})
    return rows, opening, closing, found


def parse_statement(upload, password=""):
    """Return (rows, opening, closing). rows: dicts of date/description/debit/credit."""
    name, data = upload.name.lower(), upload.read()
    if name.endswith(".pdf"):
        best, found = None, False
        for sheet in _pdf_sheets(data, password):
            rows, opening, closing, f = _parse_sheets([sheet])
            found = found or f
            if rows:
                balanced = (opening is not None and closing is not None and
                            opening + sum(r["credit"] for r in rows) - sum(r["debit"] for r in rows) == closing)
                if best is None or (balanced, len(rows)) > best[0]:
                    best = ((balanced, len(rows)), rows, opening, closing)
                if balanced:
                    break  # this reading reconciles with the bank's own balances
        if best:
            return best[1], best[2], best[3]
        if found:
            raise StatementError("The PDF's columns were found but no transactions could be read from it. Its layout may not be supported yet.")
        raise StatementError("Could not find Transaction Details / Debit / Credit columns in this PDF.")
    rows, opening, closing, found = _parse_sheets(_sheets(name, data, password))
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
    out = []
    for t in re.findall(r"[a-z0-9]+", description.lower()):
        if t.isalpha():
            out.append(t)
            continue
        out += [x for x in re.findall(r"[a-z]+", t) if len(x) >= 5]  # Tithe500, 50Offering
        if t.replace("0", "o").isalpha():  # 0ffering typed with a zero
            out.append(t.replace("0", "o"))
    return out


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
