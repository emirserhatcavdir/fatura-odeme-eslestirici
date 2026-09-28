from dataclasses import dataclass, field
from datetime import date

FATURA_COLUMNS = ("fatura_no", "musteri", "fatura_tarihi", "vade_tarihi", "tutar")
ODEME_COLUMNS = ("odeme_id", "fatura_no", "odeme_tarihi", "tutar")


@dataclass(frozen=True)
class Issue:
    dosya: str
    satir: int | None
    alan: str
    aciklama: str


@dataclass
class Table:
    filename: str
    columns: list[str] = field(default_factory=list)
    rows: list[tuple[int, dict[str, object]]] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)


@dataclass(frozen=True)
class Fatura:
    fatura_no: str
    musteri: str
    fatura_tarihi: date
    vade_tarihi: date
    tutar_kurus: int
    musteri_id: str | None = None


@dataclass(frozen=True)
class Odeme:
    odeme_id: str
    fatura_no: str
    odeme_tarihi: date
    tutar_kurus: int


@dataclass
class Report:
    rapor_tarihi: date
    faturalar: list[dict]
    eslesen_odemeler: list[dict]
    eslesmeyen_odemeler: list[dict]
    gelecek_faturalar: list[dict]
    gelecek_odemeler: list[dict]
