from django.urls import path

from . import views

urlpatterns = [
    path("support/", views.support, name="support"),
    path("manifest.webmanifest", views.manifest, name="manifest"),
    path("sw.js", views.service_worker, name="sw"),
    path("offline/", views.offline, name="offline"),
    path("", views.home, name="home"),
    path("statement/<int:pk>/", views.statement, name="statement"),
    path("statement/<int:pk>/export/", views.export, name="export"),
    path("statement/<int:pk>/recategorise/", views.recategorise, name="recategorise"),
    path("statement/<int:pk>/delete/", views.delete_statement, name="delete_statement"),
    path("transaction/<int:pk>/category/", views.set_category, name="set_category"),
    path("categories/", views.categories, name="categories"),
    path("budget/", views.budget, name="budget"),
    path("budget/delete/", views.budget_delete, name="budget_delete"),
    path("budget/mapping/", views.budget_mapping, name="budget_mapping"),
    path("budget/mapping/auto/", views.budget_automap, name="budget_automap"),
    path("budget/<int:year>/<int:month>/", views.budget_month, name="budget_month"),
    path("periods/", views.periods, name="periods"),
    path("report/", views.report, name="report"),
    path("report/export/", views.report_export, name="report_export"),
]
