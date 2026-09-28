"""Müşteri toplamları yalnızca hesaplanmış, aynı kapsamdaki faturalardan üretilir."""

from collections import defaultdict

from .reconciliation import summarize

CUSTOMER_COLUMNS = [
    "musteri", "musteri_id", "tutar_kurus", "toplam_odeme_kurus",
    "kalan_borc_kurus", "gecikmis_alacak_kurus", "fazla_odeme_kurus",
    "acik_fatura_sayisi", "en_eski_gecikme_gun",
]


def customer_key(row: dict) -> tuple[str, str]:
    """Kimlik ve ad ayrı ad alanlarıdır; esnek arama kimlikleri birleştirmez."""
    identifier = row.get("musteri_id")
    return ("id", identifier) if identifier is not None and identifier != "" else ("name", row["musteri"])


def customer_summary(rows: list[dict]) -> list[dict]:
    groups = defaultdict(list)
    for row in rows:
        groups[customer_key(row)].append(row)
    result = []
    for (kind, identifier), invoices in groups.items():
        totals = summarize(invoices)
        open_invoices = [row for row in invoices if row["kalan_borc_kurus"] > 0]
        overdue = [row for row in open_invoices if row["gecikme_gun"] > 0]
        result.append({
            # Aynı kimlikle farklı adlar gelirse hepsini göster; bir adı tahmin etme.
            "musteri": " / ".join(sorted({row["musteri"] for row in invoices})),
            "musteri_id": identifier if kind == "id" else None,
            "tutar_kurus": totals["Toplam fatura"],
            "toplam_odeme_kurus": totals["Eşleşen ödeme"],
            "kalan_borc_kurus": totals["Kalan borç"],
            "gecikmis_alacak_kurus": sum(row["kalan_borc_kurus"] for row in overdue),
            "fazla_odeme_kurus": totals["Fazla ödeme"],
            "acik_fatura_sayisi": len(open_invoices),
            "en_eski_gecikme_gun": max((row["gecikme_gun"] for row in overdue), default=0),
        })
    return sorted(result, key=lambda row: (-row["gecikmis_alacak_kurus"], row["musteri"], row["musteri_id"] or ""))
