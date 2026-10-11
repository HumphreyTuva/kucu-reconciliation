from django import forms
from django.forms import modelformset_factory

from .models import Category


class MultipleFileInput(forms.ClearableFileInput):
    allow_multiple_selected = True


class MultipleFileField(forms.FileField):
    widget = MultipleFileInput

    def clean(self, data, initial=None):
        files = data if isinstance(data, (list, tuple)) else [data]
        return [super(MultipleFileField, self).clean(f, initial) for f in files if f]


class UploadForm(forms.Form):
    file = MultipleFileField(label="Bank statements (.xlsx, .csv or .pdf)")
    pdf_password = forms.CharField(required=False, label="PDF password (if protected)",
                                   widget=forms.PasswordInput(render_value=False, attrs={"autocomplete": "off", "placeholder": "Only needed for protected PDFs"}))


CategoryFormSet = modelformset_factory(
    Category, fields=["section", "order", "name", "keywords", "is_active"], extra=2, can_delete=True)
