"""Yeni sabit kabul verisi; genel doğrulama/güvenlik testlerini tekrar etmez."""

from scripts.verify_acceptance import verify_case
import pytest


@pytest.mark.parametrize("report_date", ["2026-06-30", "2026-09-26"])
def test_csv_and_xlsx_acceptance_and_actual_export_round_trip(tmp_path, report_date):
    csv_result = verify_case("csv", tmp_path / "csv_input", report_date)
    xlsx_result = verify_case("xlsx", tmp_path / "xlsx_input", report_date)
    assert {key: value for key, value in csv_result.items() if key != "input_format"} == {
        key: value for key, value in xlsx_result.items() if key != "input_format"
    }
