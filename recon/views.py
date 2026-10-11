import io
from collections import defaultdict
from datetime import date
from decimal import Decimal

import openpyxl
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Count, Max, Min, Sum
from django.http import HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.views.decorators.http import require_POST

from .forms import CategoryFormSet, UploadForm
from .models import SECTIONS, Category, ManualEntry, Statement, Transaction
from .services import StatementError, categorise, import_statement, parse_statement

SEC_LABEL = dict(SECTIONS)
ZERO = Decimal("0")


def groups(cats):
    out = defaultdict(list)
    for c in cats:
        out[c.section].append(c)
    return [(SEC_LABEL[k], v) for k, v in sorted(out.items())]


# ---------------------------------------------------------------- statements
@login_required
def home(request):
    form = UploadForm()
    if request.method == "POST":
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            made = []
            for f in form.cleaned_data["file"]:
                try:
                    rows, opening, closing = parse_statement(f, form.cleaned_data.get("pdf_password", ""))
                except StatementError as e:
                    messages.error(request, f"{f.name}: {e}")
                    continue
                st = Statement.objects.create(name=f.name, opening=opening, closing=closing)
                import_statement(st, rows)
                made.append(st)
                if opening is not None and closing is not None:
                    gap = opening + sum(r["credit"] for r in rows) - sum(r["debit"] for r in rows) - closing
                    if gap != 0:
                        messages.warning(request, f"{f.name}: the totals do not balance (off by {gap:,.2f}), so some transactions may have been misread. "
                                                  "If this is a PDF, try the Excel version of the statement and compare.")
                seen = set(Transaction.objects.exclude(statement=st).values_list("date", "description", "debit", "credit"))
                dups = sum(1 for t in st.transactions.all() if (t.date, t.description, t.debit, t.credit) in seen)
                if dups:
                    messages.warning(request, f"{f.name}: {dups} of {len(rows)} transactions already exist in another statement. "
                                              "The yearly report would count them twice, so delete one of the statements if this was a repeat upload.")
            if len(made) == 1:
                return redirect("statement", pk=made[0].pk)
            if made:
                messages.success(request, f"{len(made)} statements reconciled: " + ", ".join(m.name for m in made) + ". Open any of them below, or see the Annual report for the combined totals.")
                return redirect("home")
    statements = Statement.objects.annotate(first=Min("transactions__date"), last=Max("transactions__date")).order_by("-first", "-uploaded_at")[:36]
    return render(request, "recon/home.html", {"form": form, "statements": statements})


def _lines(cats, agg, key, name):
    out = [{"key": str(c.id), "name": c.name, "n": agg[c.id]["n"], "total": agg[c.id]["total"]} for c in cats if c.id in agg]
    if None in agg:
        out.append({"key": key, "name": name, "n": agg[None]["n"], "total": agg[None]["total"], "warn": True})
    tot = sum((l["total"] for l in out), ZERO)
    for l in out:
        l["pct"] = float(l["total"] / tot * 100) if tot else 0
    return out, tot


def build_summary(st):
    cats = list(Category.objects.all())
    agg = lambda f, fld: {r["category"]: r for r in st.transactions.filter(**{f + "__gt": 0}).values("category")
                          .annotate(n=Count("id"), total=Sum(fld))}
    income, cr = _lines([c for c in cats if c.is_income], agg("credit", "credit"), "nocr", "Uncategorised credits")
    expense, db = _lines([c for c in cats if not c.is_income], agg("debit", "debit"), "nodb", "Uncategorised debits")
    return income, cr, expense, db


@login_required
def statement(request, pk):
    st = get_object_or_404(Statement, pk=pk)
    income, cr, expense, db = build_summary(st)
    check = st.opening + cr - db - st.closing if st.opening is not None and st.closing is not None else None
    flt = request.GET.get("cat", "")
    qs = st.transactions.select_related("category")
    if flt == "nocr":
        qs = qs.filter(credit__gt=0, category__isnull=True)
    elif flt == "nodb":
        qs = qs.filter(debit__gt=0, category__isnull=True)
    elif flt.isdigit():
        qs = qs.filter(category_id=int(flt))
    dates = [d for d in st.transactions.values_list("date", flat=True) if d]
    active = list(Category.objects.filter(is_active=True))
    return render(request, "recon/statement.html", {
        "st": st, "income": income, "expense": expense, "total_credit": cr, "total_debit": db,
        "net": cr - db, "check": check, "txs": qs, "flt": flt,
        "income_groups": groups([c for c in active if c.is_income]),
        "expense_groups": groups([c for c in active if not c.is_income]),
        "start": min(dates) if dates else None, "end": max(dates) if dates else None,
    })


@login_required
@require_POST
def set_category(request, pk):
    t = get_object_or_404(Transaction, pk=pk)
    cid = request.POST.get("category")
    t.category = Category.objects.filter(pk=cid).first() if cid else None
    t.manual, t.note = True, ""
    t.save(update_fields=["category", "manual", "note"])
    return redirect(f"/statement/{t.statement_id}/?cat={request.POST.get('cat', '')}#tx")


@login_required
@require_POST
def recategorise(request, pk):
    st = get_object_or_404(Statement, pk=pk)
    if request.POST.get("reset_manual"):
        st.transactions.update(manual=False)
    categorise(st)
    messages.success(request, "Category rules re-applied.")
    return redirect("statement", pk=pk)


@login_required
@require_POST
def delete_statement(request, pk):
    get_object_or_404(Statement, pk=pk).delete()
    return redirect("home")


def _xlsx_response(wb, name):
    buf = io.BytesIO()
    wb.save(buf)
    resp = HttpResponse(buf.getvalue(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp["Content-Disposition"] = f'attachment; filename="{name}.xlsx"'
    return resp


@login_required
def export(request, pk):
    st = get_object_or_404(Statement, pk=pk)
    income, cr, expense, db = build_summary(st)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append([f"KUCU reconciliation: {st.name}"]); ws.append([])
    ws.append(["Income (credits)", "Transactions", "Amount (KES)"])
    for l in income:
        ws.append([l["name"], l["n"], float(l["total"])])
    ws.append(["Total credits", sum(l["n"] for l in income), float(cr)]); ws.append([])
    ws.append(["Expenses (debits)", "Transactions", "Amount (KES)"])
    for l in expense:
        ws.append([l["name"], l["n"], float(l["total"])])
    ws.append(["Total debits", sum(l["n"] for l in expense), float(db)])
    if st.opening is not None:
        ws.append([]); ws.append(["Opening balance", None, float(st.opening)])
    if st.closing is not None:
        ws.append(["Closing balance", None, float(st.closing)])
    ws.column_dimensions["A"].width = 46; ws.column_dimensions["C"].width = 18
    for r in ws.iter_rows(min_col=3, max_col=3):
        r[0].number_format = "#,##0.00"
    tx = wb.create_sheet("Transactions")
    tx.append(["Date", "Details", "Debit", "Credit", "Category"])
    for t in st.transactions.select_related("category"):
        tx.append([t.date, t.description, float(t.debit) or None, float(t.credit) or None,
                   t.category.name if t.category else "Uncategorised"])
    tx.column_dimensions["B"].width = 80
    return _xlsx_response(wb, f"reconciliation-{st.pk}")


@login_required
def categories(request):
    qs = Category.objects.order_by("section", "order", "name")
    if request.method == "POST":
        fs = CategoryFormSet(request.POST, queryset=qs)
        if fs.is_valid():
            fs.save()
            messages.success(request, "Categories saved. Open a statement and press Re-run rules to apply them.")
            return redirect("categories")
    else:
        fs = CategoryFormSet(queryset=qs)
    return render(request, "recon/categories.html", {"formset": fs})


# -------------------------------------------------------------------- report
def fy_of(d):
    """Financial year ends 30 September."""
    return d.year + (1 if d.month >= 10 else 0)


def year_data(y):
    qs = Transaction.objects.filter(date__range=(date(y - 1, 10, 1), date(y, 9, 30)))
    amt, uncat = defaultdict(lambda: ZERO), 0
    for fld, prefix, fallback in (("credit", "in_", ("in_3_other", "Other incomes (uncategorised)")),
                                  ("debit", "out_", ("out_1_expenses", "Uncategorised expenses"))):
        for r in qs.filter(**{fld + "__gt": 0}).values("category__section", "category__name").annotate(t=Sum(fld), n=Count("id")):
            sec = r["category__section"]
            if sec and sec.startswith(prefix):
                amt[(sec, r["category__name"])] += r["t"]
            else:
                amt[fallback] += r["t"]; uncat += r["n"]
    manual = list(ManualEntry.objects.filter(year=y))
    for m in manual:
        amt[(m.section, m.label)] += m.amount
    return amt, uncat, qs.exists() or bool(manual)


def build_report(y):
    cur, uncat, has_cur = year_data(y)
    prev, _, has_prev = year_data(y - 1)

    def rows(sec):
        names = [c.name for c in Category.objects.filter(section=sec)]
        extra = sorted({l for (s, l) in list(cur) + list(prev) if s == sec and l not in names})
        out = [{"label": n, "cur": cur.get((sec, n), ZERO), "prev": prev.get((sec, n), ZERO)} for n in names + extra]
        out = [r for r in out if r["cur"] or r["prev"]]
        return out, sum((r["cur"] for r in out), ZERO), sum((r["prev"] for r in out), ZERO)

    S = {k: rows(k) for k, _ in SECTIONS}
    row = lambda label, k, note=None: {"label": label, "note": note, "cur": S[k][1], "prev": S[k][2]}
    income = [row("Tithes and Offering", "in_1_tithes_offering", 5), row("Income from projects", "in_2_projects", 5), row("Other income", "in_3_other", 5)]
    expenses = [{"label": r["label"], "cur": r["cur"], "prev": r["prev"]} for r in S["out_1_expenses"][0]]
    expenses.append(row("Special events", "out_2_special_events", 6))
    tot = lambda rs, f: sum((r[f] for r in rs), ZERO)
    ti, te = {f: tot(income, f) for f in ("cur", "prev")}, {f: tot(expenses, f) for f in ("cur", "prev")}
    return {
        "year": y, "income": income, "expenses": expenses, "total_income": ti, "total_expenses": te,
        "surplus": {f: ti[f] - te[f] for f in ("cur", "prev")},
        "note5": [(title, S[k]) for title, k in (("Tithes and Offering", "in_1_tithes_offering"),
                                                  ("Net proceeds from projects", "in_2_projects"), ("Other income", "in_3_other"))],
        "note6": S["out_2_special_events"], "uncat": uncat, "has_data": has_cur, "has_prev": has_prev,
    }


def _years():
    ys = {fy_of(d) for d in Transaction.objects.exclude(date=None).dates("date", "month")}
    ys |= set(ManualEntry.objects.values_list("year", flat=True))
    return sorted(ys, reverse=True)


def _pick_year(request):
    ys = _years()
    try:
        y = int(request.GET.get("year", ""))
    except ValueError:
        y = ys[0] if ys else None
    return y, ys


@login_required
def report(request):
    y, ys = _pick_year(request)
    ctx = {"years": ys, "year": y}
    if y:
        ctx.update(build_report(y))
    return render(request, "recon/report.html", ctx)


@login_required
def report_export(request):
    y, _ = _pick_year(request)
    if not y:
        return redirect("report")
    r = build_report(y)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Income statement"
    for t in ("KENYATTA UNIVERSITY CHRISTIAN UNION (KUCU) - MAIN CAMPUS", "REPORT AND FINANCIAL STATEMENTS",
              f"FOR THE YEAR ENDED 30 SEPTEMBER {y}", "STATEMENT OF COMPREHENSIVE INCOME", ""):
        ws.append([t])
    ws.append(["", "Notes", f"{y} Kshs", f"{y - 1} Kshs"])

    def put(label, note, c, p, bold=False):
        ws.append([label, note, float(c), float(p)])
        for cell in ws[ws.max_row]:
            cell.font = openpyxl.styles.Font(bold=bold)
        for cell in ws[ws.max_row][2:]:
            cell.number_format = '#,##0;(#,##0);"-"'

    ws.append(["INCOME"])
    for x in r["income"]:
        put(x["label"], x["note"], x["cur"], x["prev"])
    put("Total Income", None, r["total_income"]["cur"], r["total_income"]["prev"], True)
    ws.append([]); ws.append(["EXPENSES"])
    for x in r["expenses"]:
        put(x["label"], x.get("note"), x["cur"], x["prev"])
    put("Total expenses", None, r["total_expenses"]["cur"], r["total_expenses"]["prev"], True)
    ws.append([])
    put("Surplus/(Deficit) for the year", None, r["surplus"]["cur"], r["surplus"]["prev"], True)
    ws.column_dimensions["A"].width = 52
    for c in "CD":
        ws.column_dimensions[c].width = 16
    notes = wb.create_sheet("Notes")
    notes.append(["NOTE 5 - INCOME", "", f"{y} Kshs", f"{y - 1} Kshs"])
    for title, (rows, tc, tp) in r["note5"]:
        notes.append([title])
        for x in rows:
            notes.append([x["label"], "", float(x["cur"]), float(x["prev"])])
        notes.append([f"Total {title.lower()}", "", float(tc), float(tp)])
    notes.append([]); notes.append(["NOTE 6 - SPECIAL EVENTS"])
    for x in r["note6"][0]:
        notes.append([x["label"], "", float(x["cur"]), float(x["prev"])])
    notes.append(["Total special events", "", float(r["note6"][1]), float(r["note6"][2])])
    notes.column_dimensions["A"].width = 52
    for rw in notes.iter_rows(min_col=3, max_col=4):
        for cell in rw:
            cell.number_format = '#,##0;(#,##0);"-"'
    return _xlsx_response(wb, f"KUCU-financial-report-{y}")


# ----------------------------------------------------------------------- PWA
import json

from django.templatetags.static import static
from django.views.decorators.cache import never_cache


def manifest(request):
    data = {
        "name": "KUCU Treasury Reconciliation", "short_name": "KUCU Treasury",
        "description": "Reconcile bank statements and build the annual financial report.",
        "start_url": "/", "scope": "/", "display": "standalone",
        "background_color": "#ffffff", "theme_color": "#28309a",
        "icons": [
            {"src": static("img/icon-192.png") + "?v=2", "sizes": "192x192", "type": "image/png"},
            {"src": static("img/icon-512.png") + "?v=2", "sizes": "512x512", "type": "image/png"},
            {"src": static("img/icon-maskable-512.png") + "?v=2", "sizes": "512x512", "type": "image/png", "purpose": "maskable"},
        ],
    }
    return HttpResponse(json.dumps(data), content_type="application/manifest+json")


@never_cache
def service_worker(request):
    js = render(request, "sw.js", {"css": static("css/app.css"), "logo": static("img/logo-wide.png"),
                                   "icon": static("img/icon-192.png") + "?v=2"}).content
    resp = HttpResponse(js, content_type="application/javascript")
    resp["Service-Worker-Allowed"] = "/"
    return resp


def offline(request):
    return render(request, "offline.html")
