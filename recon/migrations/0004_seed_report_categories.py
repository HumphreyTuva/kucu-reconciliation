from django.db import migrations

from recon.defaults import DEFAULTS, OI, PR, TO


def seed(apps, schema_editor):
    C = apps.get_model("recon", "Category")
    for old, sec in (("Tithe", TO), ("Offering", TO), ("Library card", PR)):
        C.objects.filter(name=old, section="in_3_other").update(section=sec)
    C.objects.filter(name="Library card", section=PR).update(name="Library desk", keywords="library desk,library card,library")
    C.objects.filter(name="Kucu care", section="in_3_other").update(name="KUCU Care")
    tithe, off = C.objects.filter(name="Tithe", section=TO).first(), C.objects.filter(name="Offering", section=TO).first()
    if tithe and off and tithe.order < off.order:  # Offering must win over Tithe
        tithe.order, off.order = off.order, tithe.order
        tithe.save(); off.save()
    for i, (sec, name, kw) in enumerate(DEFAULTS):
        C.objects.get_or_create(name=name, section=sec, defaults={"keywords": kw, "order": 10 + i})


class Migration(migrations.Migration):
    dependencies = [("recon", "0003_sections_notes_manual")]
    operations = [migrations.RunPython(seed, migrations.RunPython.noop)]
