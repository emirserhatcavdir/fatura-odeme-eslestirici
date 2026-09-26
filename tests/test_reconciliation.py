from datetime import date, timedelta

import pytest

from eslestirici.demo import DEMO_DATE, demo_bytes
from eslestirici.models import Fatura, Odeme
from eslestirici.reconciliation import aging_bucket, aging_totals, process_files, reconcile, summarize

AS_OF = date(2026, 6, 30)


def invoice(identifier="0001", amount=500000, due=AS_OF, issued=date(2026, 1, 1)):
    return Fatura(identifier, "Hayalî Müşteri", issued, due, amount)


def payment(identifier="P01", invoice_id="0001", amount=300000, occurred=AS_OF):
    return Odeme(identifier, invoice_id, occurred, amount)


@pytest.mark.parametrize("paid, status, debt, excess", [(0, "Ödenmedi", 500000, 0), (300000, "Kısmen ödendi", 200000, 0), (500000, "Ödendi", 0, 0), (600000, "Fazla ödeme", 0, 100000)])
def test_all_statuses(paid, status, debt, excess):
    report = reconcile([invoice()], [payment(amount=paid)] if paid else [], AS_OF)
    row = report.faturalar[0]
    assert (row["durum"], row["kalan_borc_kurus"], row["fazla_odeme_kurus"]) == (status, debt, excess)


def test_multiple_payments_and_penny_precision():
    report = reconcile([invoice(amount=30)], [payment(amount=10), payment("P02", amount=20)], AS_OF)
    assert report.faturalar[0]["toplam_odeme_kurus"] == 30
    assert report.faturalar[0]["durum"] == "Ödendi"
    assert len(report.eslesen_odemeler) == 2


def test_excess_is_never_net_against_other_invoice():
    report = reconcile([invoice("0001", 100), invoice("0002", 100)], [payment(amount=200)], AS_OF)
    totals = summarize(report.faturalar)
    assert totals["Kalan borç"] == 100
    assert totals["Fazla ödeme"] == 100


def test_exact_matching_no_fuzzy_or_case_folding():
    report = reconcile([invoice("0001"), invoice("AB")], [payment(invoice_id="1"), payment("P02", "ab")], AS_OF)
    assert len(report.eslesmeyen_odemeler) == 2
    assert not report.eslesen_odemeler


@pytest.mark.parametrize("days, bucket", [(0, "Gecikme yok"), (1, "1–30 gün"), (30, "1–30 gün"), (31, "31–60 gün"), (60, "31–60 gün"), (61, "61–90 gün"), (90, "61–90 gün"), (91, "90 üzeri gün")])
def test_aging_boundaries(days, bucket):
    report = reconcile([invoice(due=AS_OF - timedelta(days=days))], [], AS_OF)
    row = report.faturalar[0]
    assert row["gecikme_gun"] == days
    assert row["gecikme_grubu"] == bucket
    assert aging_bucket(days) == bucket


def test_paid_or_not_due_invoices_are_not_overdue():
    invoices = [invoice("a", 100, AS_OF - timedelta(days=10)), invoice("b", 100, AS_OF + timedelta(days=10))]
    report = reconcile(invoices, [payment(invoice_id="a", amount=100)], AS_OF)
    assert summarize(report.faturalar)["Gecikmiş borç"] == 0
    assert all(row["gecikme_gun"] == 0 for row in report.faturalar)


def test_future_records_and_reporting_date_inclusive():
    tomorrow = AS_OF + timedelta(days=1)
    invoices = [invoice("a", 100, issued=AS_OF), invoice("b", 200, issued=tomorrow, due=tomorrow)]
    payments = [payment("p1", "a", 30, AS_OF), payment("p2", "a", 70, tomorrow), payment("p3", "b", 50, AS_OF), payment("p4", "missing", 10, tomorrow)]
    report = reconcile(invoices, payments, AS_OF)
    assert len(report.faturalar) == 1
    assert report.faturalar[0]["kalan_borc_kurus"] == 70
    assert len(report.gelecek_faturalar) == 1
    assert len(report.gelecek_odemeler) == 2
    assert report.eslesmeyen_odemeler[0]["neden"] == "Fatura raporlama tarihinden sonra düzenlenmiş"
    later = reconcile(invoices, payments, tomorrow)
    assert len(later.faturalar) == 2
    assert later.faturalar[0]["durum"] == "Ödendi"
    assert later.eslesmeyen_odemeler[0]["fatura_no"] == "missing"


def test_duplicate_ids_cannot_bypass_validation():
    with pytest.raises(ValueError, match="fatura_no"):
        reconcile([invoice(), invoice()], [], AS_OF)
    with pytest.raises(ValueError, match="odeme_id"):
        reconcile([invoice()], [payment(), payment()], AS_OF)


def test_errors_block_entire_report_including_future_duplicates():
    invoices, payments = demo_bytes()
    payments += b"P0001,0001,2099-01-01,10.00\n"
    report, errors = process_files(invoices, "a.csv", payments, "b.csv", DEMO_DATE)
    assert report is None
    assert any(error.alan == "odeme_id" for error in errors)


def test_reproducible_demo_totals_and_partitions():
    invoices, payments = demo_bytes()
    report, errors = process_files(invoices, "a.csv", payments, "b.csv", DEMO_DATE)
    assert not errors
    assert len(report.faturalar) == 9
    assert summarize(report.faturalar) == {"Toplam fatura": 1905075, "Eşleşen ödeme": 795050, "Kalan borç": 1130025, "Gecikmiş borç": 990000, "Fazla ödeme": 20000}
    assert len(report.eslesen_odemeler) == 5
    assert len(report.eslesmeyen_odemeler) == 2
    assert len(report.gelecek_faturalar) == len(report.gelecek_odemeler) == 1
    assert sum(row["borc_kurus"] for row in aging_totals(report.faturalar)) == 990000


def test_empty_report():
    report = reconcile([], [], AS_OF)
    assert report.faturalar == []
    assert all(value == 0 for value in summarize(report.faturalar).values())
