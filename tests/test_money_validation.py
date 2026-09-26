from datetime import date, datetime
from io import BytesIO

from openpyxl import Workbook
import pytest

from eslestirici.data_io import read_table
from eslestirici.models import FATURA_COLUMNS, ODEME_COLUMNS
from eslestirici.money import decimal_amount, format_tl, parse_money
from eslestirici.validation import parse_date, validate_invoices, validate_payments


@pytest.mark.parametrize("value, expected", [("0.01", 1), ("0.10", 10), ("1250.50", 125050), ("5000", 500000), (" 2.5 ", 250), ("9999999999.99", 999999999999)])
def test_money_exact_cents(value, expected):
    actual = parse_money(value)
    assert actual == expected
    assert type(actual) is int
    assert parse_money(decimal_amount(actual)) == actual


@pytest.mark.parametrize("value", ["", None, "0", "0.00", "-1.00", "+1", "1,50", "1.000", "1.234,56", "1,234.56", "1e3", "1.", ".50", "NaN", "Infinity", "1 000.00", "10000000000", True])
def test_invalid_money_is_not_rounded(value):
    with pytest.raises(ValueError):
        parse_money(value)


def test_money_display_retains_pennies():
    assert format_tl(125050) == "1.250,50 TL"
    assert format_tl(1) == "0,01 TL"
    assert format_tl(0) == "0,00 TL"


@pytest.mark.parametrize("value", ["2026-02-30", "2026-2-01", "01/02/2026", "", None, "2026-01-01 00:00:00", "2026-13-01", datetime(2026, 1, 1, 12)])
def test_invalid_dates(value):
    with pytest.raises(ValueError):
        parse_date(value)


def test_excel_dates_and_leap_day():
    assert parse_date(datetime(2024, 2, 29)) == date(2024, 2, 29)
    assert parse_date(date(2026, 1, 1)) == date(2026, 1, 1)
    assert parse_date("2024-02-29") == date(2024, 2, 29)


def invoice_csv(body=""):
    return (",".join(FATURA_COLUMNS) + "\n" + body).encode("utf-8-sig")


def payment_csv(body=""):
    return (",".join(ODEME_COLUMNS) + "\n" + body).encode("utf-8-sig")


def test_missing_columns():
    records, errors = validate_invoices(read_table(b"fatura_no,tutar\n1,20", "faturalar.csv"))
    assert records == []
    assert "musteri" in errors[0].aciklama
    assert errors[0].satir == 1


def test_empty_identity_customer_amount_and_row_locations():
    content = invoice_csv("001,A,2026-01-01,2026-01-02,10\n\n, ,2026-01-01,2026-01-02,\n")
    records, errors = validate_invoices(read_table(content, "faturalar.csv"))
    assert records == []
    assert {(item.satir, item.alan) for item in errors} == {(4, "fatura_no"), (4, "musteri"), (4, "tutar")}


@pytest.mark.parametrize("filename, content", [("bos.csv", b""), ("bos.csv", b"  \n"), ("bozuk.xlsx", b"not an excel file"), ("dosya.txt", b"abc"), ("dosya.csv", b"\xff\xff"), ("dosya.csv", b'fatura_no,tutar\n"unterminated')])
def test_bad_files_do_not_crash(filename, content):
    assert read_table(content, filename).issues


def test_header_only_files_are_valid_empty_data():
    assert validate_invoices(read_table(invoice_csv(), "bos.csv")) == ([], [])
    assert validate_payments(read_table(payment_csv(), "bos.csv")) == ([], [])


def test_duplicate_headers_and_wrong_column_count():
    assert read_table(b"fatura_no,fatura_no\n1,2", "a.csv").issues
    errors = read_table(invoice_csv("1,A,2026-01-01,2026-01-02,10,extra"), "a.csv").issues
    assert errors[0].satir == 2


def test_multiline_csv_uses_physical_source_row():
    content = invoice_csv('1,"Hayali\nMusteri",2026-01-01,2026-01-02,10\n2,A,2026-01-01,2026-01-02,bozuk\n')
    _, errors = validate_invoices(read_table(content, "a.csv"))
    assert [(error.satir, error.alan) for error in errors] == [(2, "musteri"), (4, "tutar")]


def test_whitespace_trim_and_leading_zeros():
    invoices, errors = validate_invoices(read_table(invoice_csv(" 0001 , Hayalî Müşteri ,2026-01-01,2026-01-02,1.01\n"), "a.csv"))
    assert not errors
    assert invoices[0].fatura_no == "0001"
    assert invoices[0].musteri == "Hayalî Müşteri"


def test_duplicate_invoice_ids_flag_every_occurrence():
    body = "001,A,2026-01-01,2026-01-02,10\n 001 ,B,2026-01-01,2026-01-02,20\n"
    records, errors = validate_invoices(read_table(invoice_csv(body), "a.csv"))
    assert records == []
    assert {error.satir for error in errors} == {2, 3}
    assert all(error.alan == "fatura_no" for error in errors)


def test_duplicate_payment_ids_but_not_same_invoice():
    valid = "P01,001,2026-01-01,10\nP02,001,2026-01-02,20\n"
    payments, errors = validate_payments(read_table(payment_csv(valid), "a.csv"))
    assert not errors
    assert len(payments) == 2
    records, errors = validate_payments(read_table(payment_csv(valid.replace("P02", " P01 ")), "a.csv"))
    assert records == []
    assert {error.satir for error in errors} == {2, 3}


def test_due_date_before_invoice_is_an_error():
    records, errors = validate_invoices(read_table(invoice_csv("1,A,2026-02-01,2026-01-31,10\n"), "a.csv"))
    assert records == []
    assert errors[0].alan == "vade_tarihi"


def workbook_bytes(rows):
    workbook = Workbook()
    for row in rows:
        workbook.active.append(row)
    buffer = BytesIO()
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


def test_xlsx_text_ids_native_dates_and_numeric_amounts():
    content = workbook_bytes([FATURA_COLUMNS, ["0001", "Hayalî", datetime(2026, 1, 1), datetime(2026, 1, 2), 1250.5]])
    records, errors = validate_invoices(read_table(content, "a.xlsx"))
    assert not errors
    assert records[0].fatura_no == "0001"
    assert records[0].tutar_kurus == 125050


def test_xlsx_formula_rejected_with_location():
    content = workbook_bytes([FATURA_COLUMNS, ["0001", "Hayalî", "2026-01-01", "2026-01-02", "=100+20"]])
    records, errors = validate_invoices(read_table(content, "a.xlsx"))
    assert not records
    assert any(item.satir == 2 and item.alan == "tutar" and "formülleri" in item.aciklama for item in errors)


def test_blank_xlsx_and_header_only_xlsx():
    assert read_table(workbook_bytes([]), "a.xlsx").issues
    assert validate_invoices(read_table(workbook_bytes([FATURA_COLUMNS]), "a.xlsx")) == ([], [])
