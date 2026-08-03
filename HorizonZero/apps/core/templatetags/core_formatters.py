from django import template

from apps.core.formatters import format_currency_value


register = template.Library()


@register.filter(name="crc_currency")
def crc_currency(value):
    return format_currency_value(value)
