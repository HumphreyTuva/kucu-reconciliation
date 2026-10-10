from django.contrib import admin

from .models import Category, ManualEntry, Statement, Transaction

admin.site.register(Category, list_display=("name", "section", "keywords", "order", "is_active"), list_filter=("section",))
admin.site.register(Statement)
admin.site.register(Transaction, list_display=("date", "description", "credit", "debit", "category"), list_filter=("category",))
admin.site.register(ManualEntry, list_display=("year", "section", "label", "amount"), list_filter=("year",))
