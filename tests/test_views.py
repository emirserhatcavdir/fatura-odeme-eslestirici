from datetime import date
from io import StringIO

import pandas as pd
import pytest

from eslestirici.demo import DEMO_DATE, demo_bytes
from eslestirici.exports import csv_bytes, frame
from eslestirici.reconciliation import INVOICE_RESULT_COLUMNS, process_files, summarize
from eslestirici.views import filter_invoices, invoice_detail, invoice_table


@pytest.fixture
def report():
    invoices, payments = demo_bytes()
    result, errors = process_files(invoices, "faturalar.csv", payments, "odemeler.csv", DEMO_DATE)
    assert not errors
    return result


@pytest.mark.parametrize("query, expected", [
    (" 0001 ", ["0001"]), ("ÇINAR", ["0002", "0005"]),
    ("cinar", ["0002", "0005"]), ("MAVİ", ["0001", "0006"]),
    ("[", []), ("bulunmayan müşteri", []),
])
def test_search_identifiers_customers_unicode_and_literal_symbols(report, query, expected):
    data = frame(report.faturalar, INVOICE_RESULT_COLUMNS)
    result = filter_invoices(data, query=query)
    assert result["fatura_no"].tolist() == expected
    assert len(data) == 9


def test_search_intersects_other_filters_and_csv_keeps_all_columns(report):
    data = frame(report.faturalar, INVOICE_RESULT_COLUMNS)
    filtered = filter_invoices(data, statuses=["Kısmen ödendi"], buckets=["1–30 gün"], query="mavi")
    assert filtered["fatura_no"].tolist() == ["0001"]
    exported = pd.read_csv(StringIO(csv_bytes(filtered).decode("utf-8-sig")), dtype=str)
    assert exported["fatura_no"].tolist() == ["0001"]
    assert set(INVOICE_RESULT_COLUMNS).issubset(exported.columns)
    assert exported.iloc[0]["kalan_borc_tl"] == "2000.00"
    assert exported.iloc[0]["fazla_odeme_tl"] == "0.00"


def test_no_results_and_empty_source_preserve_export_schema(report):
    for data in (frame([], INVOICE_RESULT_COLUMNS), frame(report.faturalar, INVOICE_RESULT_COLUMNS)):
        filtered = filter_invoices(data, query="bulunmayan müşteri")
        assert filtered.empty
        assert filtered.columns.tolist() == INVOICE_RESULT_COLUMNS
        assert csv_bytes(filtered).decode("utf-8-sig").startswith("fatura_no,musteri,")


def test_money_sort_stays_numeric_with_turkish_display(report):
    data = frame(report.faturalar, INVOICE_RESULT_COLUMNS)
    styled = invoice_table(data)
    assert styled.data.columns.tolist() == ["Fatura no", "Müşteri", "Vade", "Fatura tutarı", "Ödenen", "Kalan", "Durum"]
    assert pd.api.types.is_integer_dtype(styled.data["Fatura tutarı"])
    assert styled.data.sort_values("Fatura tutarı")["Fatura no"].tolist()[:3] == ["0007", "0006", "0004"]
    assert len(invoice_table(data, detailed=True).data.columns) == 11
    assert data.columns.tolist() == INVOICE_RESULT_COLUMNS


def test_detail_lists_installments_and_never_mutates_report(report):
    detail = invoice_detail(report, "0005")
    assert detail["fatura"]["toplam_odeme_kurus"] == 250000
    assert [row["odeme_id"] for row in detail["odemeler"]] == ["P0004", "P0005"]
    assert sum(row["tutar_kurus"] for row in detail["odemeler"]) == 250000
    detail["odemeler"][0]["tutar_kurus"] = 1
    assert invoice_detail(report, "0005")["odemeler"][0]["tutar_kurus"] == 100000


def test_future_payment_is_visible_but_excluded_until_reporting_date(report):
    detail = invoice_detail(report, "0007")
    assert detail["odemeler"] == []
    assert [row["odeme_id"] for row in detail["gelecek_odemeler"]] == ["P0007"]
    assert detail["fatura"]["toplam_odeme_kurus"] == 0
    assert detail["fatura"]["kalan_borc_kurus"] == 60025
    assert summarize(report.faturalar)["Eşleşen ödeme"] == 795050
    invoices, payments = demo_bytes()
    later, errors = process_files(invoices, "a.csv", payments, "b.csv", date(2026, 7, 1))
    assert not errors
    later_detail = invoice_detail(later, "0007")
    assert later_detail["gelecek_odemeler"] == []
    assert later_detail["fatura"]["durum"] == "Ödendi"
    assert [row["odeme_id"] for row in later_detail["odemeler"]] == ["P0007"]


def test_detail_lookup_is_exact_and_handles_unpaid_and_missing(report):
    assert invoice_detail(report, "1") is None
    assert invoice_detail(report, "0008") is None  # Bu tarihte gelecek faturadır.
    assert invoice_detail(report, "olmayan") is None
    detail = invoice_detail(report, "0003")
    assert detail["odemeler"] == detail["gelecek_odemeler"] == []
    assert detail["fatura"]["durum"] == "Ödenmedi"
