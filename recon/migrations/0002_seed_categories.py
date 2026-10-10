from django.db import migrations

DEFAULTS = [("Tithe", "tithe"), ("Offering", "offering,offerings"), ("Kucu care", "kucu care,kucucare"),
            ("BS", "bs"), ("Bible", "bible"), ("Commission", "commission"), ("Library card", "library card")]


def seed(apps, schema_editor):
    Category = apps.get_model("recon", "Category")
    for i, (n, k) in enumerate(DEFAULTS):
        Category.objects.get_or_create(name=n, defaults={"keywords": k, "order": i})


class Migration(migrations.Migration):
    dependencies = [("recon", "0001_initial")]
    operations = [migrations.RunPython(seed, migrations.RunPython.noop)]
