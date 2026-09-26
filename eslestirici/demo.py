"""Sabit, tamamen hayalî veriler. Sistem tarihine göre değişmez."""
from datetime import date
from pathlib import Path

import pandas as pd

from .exports import csv_bytes, xlsx_bytes
from .models import FATURA_COLUMNS, ODEME_COLUMNS

DEMO_DATE = date(2026, 6, 30)
DEMO_INVOICES = [
    ("0001", "Hayalî Mavi Kırtasiye", "2026-06-01", "2026-06-10", "5000.00"),
    ("0002", "Hayalî Çınar Atölye", "2026-06-05", "2026-06-20", "1250.50"),
    ("0003", "Hayalî Lale Tasarım", "2026-03-01", "2026-03-15", "4200.00"),
    ("0004", "Hayalî Ada Kitap", "2026-06-02", "2026-06-15", "1000.00"),
    ("0005", "Hayalî Çınar Atölye", "2026-05-01", "2026-05-30", "2500.00"),
    ("0006", "Hayalî Mavi Kırtasiye", "2026-06-20", "2026-06-30", "800.00"),
    ("0007", "Hayalî Ada Kitap", "2026-06-25", "2026-07-15", "600.25"),
    ("0008", "Hayalî Lale Tasarım", "2026-07-01", "2026-07-31", "2000.00"),
    ("0009", "Hayalî Ufuk Seramik", "2026-04-01", "2026-05-15", "1500.00"),
    ("0010", "Hayalî Ufuk Seramik", "2026-03-10", "2026-04-15", "2200.00"),
]
DEMO_PAYMENTS = [
    ("P0001", "0001", "2026-06-10", "3000.00"),
    ("P0002", "0002", "2026-06-20", "1250.50"),
    ("P0003", "0004", "2026-06-15", "1200.00"),
    ("P0004", "0005", "2026-05-10", "1000.00"),
    ("P0005", "0005", "2026-05-30", "1500.00"),
    ("P0006", "0099", "2026-06-21", "350.75"),
    ("P0007", "0007", "2026-07-01", "600.25"),
    ("P0008", "0008", "2026-06-30", "500.00"),
]


def sample_frame(kind: str, blank: bool = False) -> pd.DataFrame:
    columns = FATURA_COLUMNS if kind == "faturalar" else ODEME_COLUMNS
    rows = DEMO_INVOICES if kind == "faturalar" else DEMO_PAYMENTS
    return pd.DataFrame([] if blank else rows, columns=columns, dtype=object)


def demo_bytes() -> tuple[bytes, bytes]:
    return csv_bytes(sample_frame("faturalar")), csv_bytes(sample_frame("odemeler"))


def write_assets(folder: Path) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    for kind in ("faturalar", "odemeler"):
        for blank, prefix in ((False, "ornek"), (True, "sablon")):
            data = sample_frame(kind, blank)
            (folder / f"{prefix}_{kind}.csv").write_bytes(csv_bytes(data))
            (folder / f"{prefix}_{kind}.xlsx").write_bytes(xlsx_bytes({kind: data}))
    invalid_folder = folder / "hatali"
    invalid_folder.mkdir(exist_ok=True)
    invoices = sample_frame("faturalar").iloc[:3].copy()
    invoices.loc[1, "fatura_no"] = "0001"
    invoices.loc[2, "tutar"] = "1250,50"
    invoices.loc[2, "vade_tarihi"] = "2026-02-30"
    payments = sample_frame("odemeler").iloc[:3].copy()
    payments.loc[1, "odeme_id"] = "P0001"
    payments.loc[2, "fatura_no"] = ""
    (invalid_folder / "hatali_faturalar.csv").write_bytes(csv_bytes(invoices))
    (invalid_folder / "hatali_odemeler.csv").write_bytes(csv_bytes(payments))


if __name__ == "__main__":
    write_assets(Path(__file__).resolve().parent.parent / "veriler")
