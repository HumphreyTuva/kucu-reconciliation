from django.urls import path

from . import views

urlpatterns = [
    path("", views.home, name="home"),
    path("statement/<int:pk>/", views.statement, name="statement"),
    path("statement/<int:pk>/export/", views.export, name="export"),
    path("statement/<int:pk>/recategorise/", views.recategorise, name="recategorise"),
    path("statement/<int:pk>/delete/", views.delete_statement, name="delete_statement"),
    path("transaction/<int:pk>/category/", views.set_category, name="set_category"),
    path("categories/", views.categories, name="categories"),
    path("report/", views.report, name="report"),
    path("report/export/", views.report_export, name="report_export"),
]
