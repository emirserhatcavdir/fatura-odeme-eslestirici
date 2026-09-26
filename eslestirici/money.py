"""Para hesapları yalnızca Decimal doğrulaması ve tam sayı kuruş kullanır."""

import re
from decimal import Decimal

# 50.000 satırlık dosyalarda da toplamlar int64 sınırının içinde kalır.
MAX_AMOUNT_CENTS = 999_999_999_999


def parse_money(value: object) -> int:
    text = "" if value is None else str(value).strip()
    if not re.fullmatch(r"[0-9]+(?:\.[0-9]{1,2})?", text):
        raise ValueError("Tutar 1250.50 biçiminde olmalı: binlik ayırıcı yok, ondalık nokta ve en fazla iki ondalık basamak.")
    if len(text) > 16:
        raise ValueError("Tutar sınırı aşıldı: en fazla 9.999.999.999,99 TL.")
    amount = Decimal(text)
    if amount <= 0:
        raise ValueError("Tutar sıfırdan büyük olmalı; iade ve alacak dekontları kapsam dışıdır.")
    cents = int(amount * 100)
    if cents > MAX_AMOUNT_CENTS:
        raise ValueError("Tutar sınırı aşıldı: en fazla 9.999.999.999,99 TL.")
    return cents


def decimal_amount(cents: int) -> str:
    """Dışa aktarımda kuruş kaybı olmadan noktalı ondalık metin."""
    sign = "-" if cents < 0 else ""
    whole, fraction = divmod(abs(cents), 100)
    return f"{sign}{whole}.{fraction:02d}"


def format_tl(cents: int) -> str:
    sign = "-" if cents < 0 else ""
    whole, fraction = divmod(abs(cents), 100)
    grouped = f"{whole:,}".replace(",", ".")
    return f"{sign}{grouped},{fraction:02d} TL"
