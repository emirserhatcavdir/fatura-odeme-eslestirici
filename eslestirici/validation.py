from collections import defaultdict
from datetime import date, datetime
import re

from .models import FATURA_COLUMNS, ODEME_COLUMNS, Fatura, Issue, Odeme, Table
from .money import parse_money


def parse_date(value: object) -> date:
    if isinstance(value, datetime):
        if value.time().isoformat() != "00:00:00":
            raise ValueError("Tarih saat içermemeli; YYYY-MM-DD biçimini kullanın.")
        return value.date()
    if isinstance(value, date):
        return value
    text = "" if value is None else str(value).strip()
    if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", text):
        raise ValueError("Tarih YYYY-MM-DD biçiminde olmalı (örnek: 2026-06-30).")
    try:
        return date.fromisoformat(text)
    except ValueError:
        raise ValueError("Geçersiz takvim tarihi; gün ve ayı kontrol edin.") from None


def _required_text(value: object) -> str:
    text = "" if value is None else str(value).strip()
    if not text:
        raise ValueError("Bu alan boş bırakılamaz.")
    if len(text) > 500:
        raise ValueError("Bu alan en fazla 500 karakter olabilir.")
    if any(ord(char) < 32 for char in text):
        raise ValueError("Bu alan satır sonu veya kontrol karakteri içeremez.")
    return text


def _validate(table: Table, invoice: bool) -> tuple[list, list[Issue]]:
    required = FATURA_COLUMNS if invoice else ODEME_COLUMNS
    id_field = "fatura_no" if invoice else "odeme_id"
    errors = list(table.issues)
    missing = [column for column in required if column not in table.columns]
    if missing:
        errors.append(Issue(table.filename, 1, ", ".join(missing), "Eksik sütunlar: " + ", ".join(missing)))
        return [], errors
    identifiers = defaultdict(list)
    records = []
    for number, row in table.rows:
        parsed = {}
        for column in required:
            value = row.get(column)
            parser = parse_money if column == "tutar" else parse_date if column.endswith("_tarihi") else _required_text
            try:
                parsed[column] = parser(value)
            except ValueError as exc:
                errors.append(Issue(table.filename, number, column, str(exc)))
        if id_field in parsed:
            identifiers[parsed[id_field]].append(number)
        if len(parsed) != len(required):
            continue
        if invoice:
            customer_id = None
            if row.get("musteri_id") is not None and str(row["musteri_id"]).strip():
                try:
                    customer_id = _required_text(row["musteri_id"])
                except ValueError as exc:
                    errors.append(Issue(table.filename, number, "musteri_id", str(exc)))
                    continue
            if parsed["vade_tarihi"] < parsed["fatura_tarihi"]:
                errors.append(Issue(table.filename, number, "vade_tarihi", "Vade tarihi fatura tarihinden önce olamaz."))
                continue
            records.append(Fatura(parsed["fatura_no"], parsed["musteri"], parsed["fatura_tarihi"], parsed["vade_tarihi"], parsed["tutar"], customer_id))
        else:
            records.append(Odeme(parsed["odeme_id"], parsed["fatura_no"], parsed["odeme_tarihi"], parsed["tutar"]))
    for identifier, row_numbers in identifiers.items():
        if len(row_numbers) > 1:
            locations = ", ".join(map(str, row_numbers[:10])) + ("…" if len(row_numbers) > 10 else "")
            for number in row_numbers:
                errors.append(Issue(table.filename, number, id_field, f"Yinelenen {id_field}: {identifier}. Satırlar: {locations}. Kayıtlar otomatik birleştirilmez."))
    # Hatalı dosyadan kısmi sonuç sızdırma: rapor yalnızca tüm kontrollerden sonra üretilir.
    return ([] if errors else records), errors


def validate_invoices(table: Table) -> tuple[list[Fatura], list[Issue]]:
    return _validate(table, invoice=True)


def validate_payments(table: Table) -> tuple[list[Odeme], list[Issue]]:
    return _validate(table, invoice=False)
