"""Arayüz için filtre ve detay görünümleri; hesaplanan raporu değiştirmez."""

import unicodedata

import pandas as pd

from .exports import safe_text
from .customers import customer_key
from .models import Report
from .money import format_tl

BASE_COLUMNS = ["fatura_no", "musteri", "vade_tarihi", "tutar_kurus", "toplam_odeme_kurus", "kalan_borc_kurus", "durum"]
DETAIL_COLUMNS = ["fatura_tarihi", "fazla_odeme_kurus", "gecikme_gun", "gecikme_grubu"]
INVOICE_LABELS = {
    "fatura_no": "Fatura no", "musteri": "Müşteri", "vade_tarihi": "Vade",
    "tutar_kurus": "Fatura tutarı", "toplam_odeme_kurus": "Ödenen", "kalan_borc_kurus": "Kalan",
    "durum": "Durum", "fatura_tarihi": "Fatura tarihi", "fazla_odeme_kurus": "Fazla ödeme",
    "gecikme_gun": "Gecikme günü", "gecikme_grubu": "Gecikme grubu",
}
STATUS_COLORS = {
    "Ödendi": ("#DCF4EB", "#14523D"),
    "Kısmen ödendi": ("#FFF0CC", "#714600"),
    "Ödenmedi": ("#FCE4E5", "#8F2331"),
    "Fazla ödeme": ("#E5EDFF", "#244578"),
}


def search_text(value: str) -> str:
    value = value.strip().translate(str.maketrans({"I": "i", "İ": "i", "ı": "i"}))
    return "".join(char for char in unicodedata.normalize("NFKD", value.casefold()) if not unicodedata.combining(char))


def filter_invoices(data: pd.DataFrame, customers=(), statuses=(), buckets=(), query: str = "", *, customer_keys=()) -> pd.DataFrame:
    result = data
    for column, values in (("musteri", customers), ("durum", statuses), ("gecikme_grubu", buckets)):
        if values:
            result = result[result[column].isin(values)]
    term = search_text(query)
    if term:
        matches = result["fatura_no"].map(lambda value: term in search_text(value)) | result["musteri"].map(lambda value: term in search_text(value))
        result = result.loc[matches.astype(bool)]
    if customer_keys:
        keys = set(customer_keys)
        mask = pd.Series([customer_key(row) in keys for row in result.to_dict("records")], index=result.index, dtype=bool)
        result = result.loc[mask]
    return result.copy()


def invoice_detail(report: Report, invoice_id: str) -> dict | None:
    invoice = next((row for row in report.faturalar if row["fatura_no"] == invoice_id), None)
    if invoice is None:
        return None
    sort_key = lambda row: (row["odeme_tarihi"], row["odeme_id"])
    return {
        "fatura": invoice.copy(),
        "odemeler": sorted((row.copy() for row in report.eslesen_odemeler if row["fatura_no"] == invoice_id), key=sort_key),
        "gelecek_odemeler": sorted((row.copy() for row in report.gelecek_odemeler if row["fatura_no"] == invoice_id), key=sort_key),
    }


def invoice_table(data: pd.DataFrame, detailed: bool = False):
    columns = BASE_COLUMNS + (DETAIL_COLUMNS if detailed else [])
    if "musteri_id" in data.columns:
        columns = columns[:2] + ["musteri_id"] + columns[2:]
    view = data.loc[:, columns].reset_index(drop=True).map(safe_text)
    # Sayısal veri kuruş olarak kalır. Styler yalnızca görünen yazıyı değiştirir.
    money_columns = [column for column in columns if column.endswith("_kurus")]
    for column in money_columns:
        view[column] = view[column].astype("int64")
    view = view.rename(columns={**INVOICE_LABELS, "musteri_id": "Müşteri kimliği"})
    formats = {INVOICE_LABELS[column]: format_tl for column in money_columns}
    formats.update({INVOICE_LABELS[column]: lambda value: value.strftime("%d.%m.%Y") for column in columns if column.endswith("_tarihi")})

    def status_style(status):
        background, foreground = STATUS_COLORS[status]
        return f"background-color: {background}; color: {foreground}; font-weight: 600"

    return view.style.format(formats).map(status_style, subset=["Durum"])
