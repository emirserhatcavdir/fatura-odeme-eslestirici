from collections import defaultdict
from dataclasses import asdict
from datetime import date

from .data_io import read_table
from .models import Fatura, Issue, Odeme, Report
from .validation import validate_invoices, validate_payments

STATUSES = ("Ödendi", "Kısmen ödendi", "Ödenmedi", "Fazla ödeme")
AGING_BUCKETS = ("1–30 gün", "31–60 gün", "61–90 gün", "90 üzeri gün")
INVOICE_RESULT_COLUMNS = ["fatura_no", "musteri", "fatura_tarihi", "vade_tarihi", "tutar_kurus", "toplam_odeme_kurus", "kalan_borc_kurus", "fazla_odeme_kurus", "durum", "gecikme_gun", "gecikme_grubu"]
PAYMENT_COLUMNS = ["odeme_id", "fatura_no", "odeme_tarihi", "tutar_kurus"]
INVOICE_COLUMNS = ["fatura_no", "musteri", "fatura_tarihi", "vade_tarihi", "tutar_kurus"]


def invoice_record(invoice: Fatura) -> dict:
    row = asdict(invoice)
    if row["musteri_id"] is None:
        row.pop("musteri_id")
    return row


def invoice_columns(rows: list[dict], *, results: bool = True) -> list[str]:
    """Kimlik verilmeyen eski dosyaların CSV/Excel sütunlarını koru."""
    columns = INVOICE_RESULT_COLUMNS if results else INVOICE_COLUMNS
    return columns + (["musteri_id"] if any(row.get("musteri_id") is not None for row in rows) else [])


def aging_bucket(days: int) -> str:
    if days <= 0:
        return "Gecikme yok"
    if days <= 30:
        return AGING_BUCKETS[0]
    if days <= 60:
        return AGING_BUCKETS[1]
    if days <= 90:
        return AGING_BUCKETS[2]
    return AGING_BUCKETS[3]


def reconcile(invoices: list[Fatura], payments: list[Odeme], as_of: date) -> Report:
    # Bu fonksiyon doğrudan çağrılsa da yinelenen kimlikleri toplayamaz.
    if len({item.fatura_no for item in invoices}) != len(invoices):
        raise ValueError("Yinelenen fatura_no ile rapor üretilemez.")
    if len({item.odeme_id for item in payments}) != len(payments):
        raise ValueError("Yinelenen odeme_id ile rapor üretilemez.")
    eligible = {item.fatura_no: item for item in invoices if item.fatura_tarihi <= as_of}
    all_invoice_ids = {item.fatura_no for item in invoices}
    future_invoices = [invoice_record(item) for item in invoices if item.fatura_tarihi > as_of]
    future_payments, unmatched, matched = [], [], []
    paid = defaultdict(int)
    for payment in payments:
        row = asdict(payment)
        if payment.odeme_tarihi > as_of:
            future_payments.append(row)
        elif payment.fatura_no not in eligible:
            row["neden"] = "Fatura raporlama tarihinden sonra düzenlenmiş" if payment.fatura_no in all_invoice_ids else "Fatura bulunamadı"
            unmatched.append(row)
        else:
            paid[payment.fatura_no] += payment.tutar_kurus
            matched.append(row)
    results = []
    for invoice in eligible.values():
        total = paid[invoice.fatura_no]
        debt = max(invoice.tutar_kurus - total, 0)
        excess = max(total - invoice.tutar_kurus, 0)
        if excess:
            status = "Fazla ödeme"
        elif total == invoice.tutar_kurus:
            status = "Ödendi"
        elif total:
            status = "Kısmen ödendi"
        else:
            status = "Ödenmedi"
        days = max((as_of - invoice.vade_tarihi).days, 0) if debt else 0
        results.append({**invoice_record(invoice), "toplam_odeme_kurus": total, "kalan_borc_kurus": debt, "fazla_odeme_kurus": excess, "durum": status, "gecikme_gun": days, "gecikme_grubu": aging_bucket(days)})
    return Report(as_of, results, matched, unmatched, future_invoices, future_payments)


def summarize(rows: list[dict]) -> dict[str, int]:
    return {
        "Toplam fatura": sum(row["tutar_kurus"] for row in rows),
        "Eşleşen ödeme": sum(row["toplam_odeme_kurus"] for row in rows),
        "Kalan borç": sum(row["kalan_borc_kurus"] for row in rows),
        "Gecikmiş borç": sum(row["kalan_borc_kurus"] for row in rows if row["gecikme_gun"] > 0),
        "Fazla ödeme": sum(row["fazla_odeme_kurus"] for row in rows),
    }


def aging_totals(rows: list[dict]) -> list[dict]:
    return [{"gecikme_grubu": bucket, "borc_kurus": sum(row["kalan_borc_kurus"] for row in rows if row["gecikme_grubu"] == bucket)} for bucket in AGING_BUCKETS]


def process_files(invoice_bytes: bytes, invoice_name: str, payment_bytes: bytes, payment_name: str, as_of: date) -> tuple[Report | None, list[Issue]]:
    invoices, invoice_errors = validate_invoices(read_table(invoice_bytes, invoice_name))
    payments, payment_errors = validate_payments(read_table(payment_bytes, payment_name))
    errors = invoice_errors + payment_errors
    return (None, errors) if errors else (reconcile(invoices, payments, as_of), [])
