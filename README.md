# KUCU Reconciliation (Django)

    python -m venv venv && source venv/bin/activate
    pip install -r requirements.txt
    python manage.py migrate          # also seeds Tithe, Offering, Kucu care, BS, Bible, Commission, Library card
    python manage.py createsuperuser  # treasurer/admin login
    python manage.py runserver

Log in, upload each month's bank statement, read the totals. Open **Annual report** for the
year-end financial statements (financial year ends 30 September), built from all uploaded statements.

- Matching is forgiving: plurals, misspellings (ofering, titthe) and glued words are recognised. A payment that
  mentions both Tithe and Offering (e.g. "Tithe/Offerings") counts as Offering; "Tithe"/"Tithes" count as Tithe.
  This works because Offering has a lower *order* than Tithe under Categories.
- Each category belongs to a report section (Tithes and Offering, Projects, Other income, General expenses, Special events).
- Non-bank items such as Depreciation: Admin > Manual report entries.

Production: set DJANGO_SECRET_KEY, DJANGO_DEBUG=0, DJANGO_ALLOWED_HOSTS=yourdomain,
run `python manage.py collectstatic` and serve /static/ (e.g. WhiteNoise or your host's static mapping).
