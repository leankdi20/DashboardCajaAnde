from decimal import Decimal, InvalidOperation


def format_currency_value(value, symbol="₡", empty_value="—"):
    if value in (None, ""):
        return empty_value

    try:
        amount = Decimal(str(value).strip())
    except (InvalidOperation, ValueError, TypeError, AttributeError):
        return str(value)

    negative = amount < 0
    amount = abs(amount)
    quantized = amount.quantize(Decimal("0.01"))
    integer_part = int(quantized)
    decimal_part = int((quantized - integer_part) * 100)

    integer_formatted = f"{integer_part:,}".replace(",", ".")
    if decimal_part:
        formatted = f"{symbol}{integer_formatted},{decimal_part:02d}"
    else:
        formatted = f"{symbol}{integer_formatted}"

    if negative:
        return f"-{formatted}"
    return formatted
