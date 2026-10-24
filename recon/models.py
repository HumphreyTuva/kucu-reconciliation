import re

from django.db import models

SECTIONS = [
    ("in_1_tithes_offering", "Income: Tithes and Offering"),
    ("in_2_projects", "Income: Net proceeds from projects"),
    ("in_3_other", "Income: Other income"),
    ("out_1_expenses", "Expense: General expenses"),
    ("out_2_special_events", "Expense: Special events"),
    ("ex_1_transfers", "Not counted: transfers between own accounts"),
]


class Category(models.Model):
    name = models.CharField(max_length=80)
    section = models.CharField(max_length=30, choices=SECTIONS, default="in_3_other",
                               help_text="Where it appears in the financial report. Income sections apply to credits, expense sections to debits.")
    keywords = models.CharField(max_length=300, help_text="Comma-separated. Misspellings and plurals are recognised automatically.")
    order = models.PositiveSmallIntegerField(default=0, help_text="Lower numbers are checked first and win when several categories match.")
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ["order", "name"]
        verbose_name_plural = "categories"
        constraints = [models.UniqueConstraint(fields=["name", "section"], name="uniq_cat_name_section")]

    def __str__(self):
        return f"{self.name} ({self.get_section_display()})"

    @property
    def is_income(self):
        return self.section.startswith("in_")

    @property
    def is_excluded(self):
        return self.section.startswith("ex_")

    def keyword_list(self):
        return [k.strip() for k in self.keywords.split(",") if k.strip()]


class Statement(models.Model):
    name = models.CharField(max_length=200)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    opening = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    closing = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)

    class Meta:
        ordering = ["-uploaded_at"]

    def __str__(self):
        return self.name


class Transaction(models.Model):
    statement = models.ForeignKey(Statement, on_delete=models.CASCADE, related_name="transactions")
    date = models.DateField(null=True, blank=True)
    description = models.TextField()
    debit = models.DecimalField(max_digits=14, decimal_places=2, default=0)
    credit = models.DecimalField(max_digits=14, decimal_places=2, default=0)
    category = models.ForeignKey(Category, null=True, blank=True, on_delete=models.SET_NULL)
    manual = models.BooleanField(default=False, help_text="Category set by hand; kept when rules are re-run.")
    note = models.CharField(max_length=200, blank=True, help_text="How the category was decided, e.g. a misspelling that was corrected.")

    class Meta:
        ordering = ["date", "id"]

    @property
    def short_description(self):
        return re.sub(r"^Transfer In From Acc\.No\. - Kes\d+ ", "", self.description, flags=re.I)


class ManualEntry(models.Model):
    """Report items that never pass through the bank, e.g. depreciation or tuck shop net proceeds."""
    year = models.PositiveSmallIntegerField(help_text="Financial year ending 30 September of this year, e.g. 2025.")
    section = models.CharField(max_length=30, choices=SECTIONS)
    label = models.CharField(max_length=80, help_text="Line name, e.g. Depreciation")
    amount = models.DecimalField(max_digits=14, decimal_places=2)
    note = models.CharField(max_length=200, blank=True)

    class Meta:
        ordering = ["-year", "section", "label"]
        verbose_name_plural = "manual report entries"

    def __str__(self):
        return f"{self.year} {self.label}: {self.amount}"


class BudgetLine(models.Model):
    """One expense line for one month from the budget workbook: what was budgeted and what was actually spent."""
    year = models.PositiveSmallIntegerField()
    month = models.PositiveSmallIntegerField()
    group = models.CharField(max_length=80, blank=True, help_text="Department, e.g. Discipleship or Worship Comm")
    item = models.CharField(max_length=120)
    actual = models.DecimalField(max_digits=14, decimal_places=2, default=0)
    budget = models.DecimalField(max_digits=14, decimal_places=2, default=0)
    category = models.ForeignKey(Category, null=True, blank=True, on_delete=models.SET_NULL, related_name="budget_lines")

    class Meta:
        ordering = ["year", "month", "id"]
        indexes = [models.Index(fields=["year", "month"])]

    def __str__(self):
        return f"{self.year}-{self.month:02d} {self.group}: {self.item}"


class ItemMap(models.Model):
    """Remembers which report line a budget item belongs to, so later uploads are mapped automatically."""
    key = models.CharField(max_length=120, unique=True)
    category = models.ForeignKey(Category, on_delete=models.CASCADE)
