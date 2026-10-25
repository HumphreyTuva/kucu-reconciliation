from django.contrib import admin

from .models import BudgetLine, Category, FinancialYear, ItemMap, ManualEntry, Statement, Transaction

admin.site.register(Category, list_display=("name", "section", "keywords", "order", "is_active"), list_filter=("section",))
admin.site.register(Statement)
admin.site.register(Transaction, list_display=("date", "description", "credit", "debit", "category"), list_filter=("category",))
admin.site.register(ManualEntry, list_display=("period", "year", "section", "label", "amount"), list_filter=("period", "year"))
admin.site.register(FinancialYear, list_display=("name", "start", "end"))


@admin.register(BudgetLine)
class BudgetLineAdmin(admin.ModelAdmin):
    """Imported budget workbook lines. Select rows and choose "Delete selected" to clear old data before a fresh import."""
    list_display = ("year", "month", "group", "item", "actual", "budget", "category")
    list_filter = ("year", "month", "group", ("category", admin.EmptyFieldListFilter))
    search_fields = ("item", "group")
    list_per_page = 100
    list_select_related = ("category",)


@admin.register(ItemMap)
class ItemMapAdmin(admin.ModelAdmin):
    """Remembered choices of report line for each budget item."""
    list_display = ("key", "category")
    search_fields = ("key",)
