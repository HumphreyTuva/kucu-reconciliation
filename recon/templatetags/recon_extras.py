from decimal import Decimal

from django import template

register = template.Library()


@register.filter
def kes(v):
    """Whole shillings, negatives in brackets, zero as a dash (as in the printed accounts)."""
    if v in (None, "") or v == 0:
        return "–"
    v = Decimal(v)
    s = f"{abs(v):,.0f}"
    return f"({s})" if v < 0 else s
