"""Yerel Streamlit arayüzü. Çalıştırma komutları README.md dosyasındadır."""

from datetime import date
from hashlib import sha256
from pathlib import Path

import streamlit as st

from eslestirici.demo import DEMO_DATE, demo_bytes, sample_frame
from eslestirici.exports import csv_bytes, frame, issues_frame, report_sheets, safe_text, xlsx_bytes
from eslestirici.exports import EXCEL_EXPORT_VERSION, ExcelPrecisionError
from eslestirici.money import format_tl
from eslestirici.reconciliation import AGING_BUCKETS, INVOICE_COLUMNS, INVOICE_RESULT_COLUMNS, PAYMENT_COLUMNS, STATUSES, process_files, summarize
from eslestirici.ui import clear_filters, show_aging_chart, show_invoice_detail, show_invoice_table, show_rules, show_table
from eslestirici.views import filter_invoices

st.set_page_config(page_title="Fatura–Ödeme Eşleştirici", page_icon="₺", layout="wide")
stylesheet = (Path(__file__).resolve().parent / "assets" / "app.css").read_text(encoding="utf-8")
st.html(f"<style>{stylesheet}</style>")
st.session_state.setdefault("report_date", date.today())


@st.cache_data(show_spinner=False)
def download_asset(kind: str, blank: bool, extension: str) -> bytes:
    data = sample_frame(kind, blank)
    return csv_bytes(data) if extension == "csv" else xlsx_bytes({kind: data})


def activate_demo() -> None:
    st.session_state["source"] = "Örnek veriler"
    st.session_state["report_date"] = DEMO_DATE


with st.sidebar:
    st.markdown("### Çalışma alanı")
    st.caption("Python portföy projesi · Sürüm 2 · TL")
    st.button("Örnek verilerle dene", on_click=activate_demo, type="primary", width="stretch", key="sidebar_demo")
    source = st.radio("Veri kaynağı", ["Dosya yükle", "Örnek veriler"], key="source")
    as_of = st.date_input("Raporlama tarihi", value=None, min_value=date(1900, 1, 1), max_value=date(2100, 12, 31), format="YYYY-MM-DD", key="report_date")
    st.caption("Bu tarih dahil düzenlenen faturalar ve gerçekleşen ödemeler hesaba katılır.")
    st.divider()
    st.markdown("### Şablonlar ve örnekler")
    for blank, label in ((True, "Boş şablonlar"), (False, "Hayalî örnek veriler")):
        with st.expander(label):
            for kind, kind_label in (("faturalar", "Faturalar"), ("odemeler", "Ödemeler")):
                st.caption(kind_label)
                for column, extension in zip(st.columns(2), ("csv", "xlsx")):
                    prefix = "sablon" if blank else "ornek"
                    with column:
                        st.download_button(extension.upper(), download_asset(kind, blank, extension), file_name=f"{prefix}_{kind}.{extension}", mime="text/csv" if extension == "csv" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"{prefix}_{kind}_{extension}", width="stretch")
    st.divider()
    st.caption("Hayalî verilerle dene. Dosyalar uygulamanın çalıştığı sunucunun belleğinde işlenir; bulut demosuna gerçek veya kişisel veri yükleme.")

st.markdown('<div class="eyebrow">TAHSİLAT TAKİBİ / TL</div>', unsafe_allow_html=True)
st.title("Fatura–Ödeme Eşleştirici")
st.markdown('<p class="intro">Tahsilatını izle, açık bakiyeleri ve ödeme detaylarını incele.</p>', unsafe_allow_html=True)
welcome = st.container()
show_rules()
if as_of is None:
    st.info("Raporu görmek için bir raporlama tarihi seç.")
    st.session_state.pop("selected_invoice", None)
    st.stop()

if source == "Örnek veriler":
    st.caption(f"Hayalî örnek veri seti · Önerilen raporlama tarihi: {DEMO_DATE:%d.%m.%Y}")
    invoice_bytes, payment_bytes = demo_bytes()
    invoice_name, payment_name = "ornek_faturalar.csv", "ornek_odemeler.csv"
else:
    with st.container(border=True, key="upload-grid"):
        st.subheader("Dosyalarını yükle")
        left, right = st.columns(2)
        with left:
            invoice_upload = st.file_uploader("Fatura dosyası", type=["csv", "xlsx"], key="invoice_upload", help="fatura_no, musteri, fatura_tarihi, vade_tarihi, tutar")
        with right:
            payment_upload = st.file_uploader("Ödeme dosyası", type=["csv", "xlsx"], key="payment_upload", help="odeme_id, fatura_no, odeme_tarihi, tutar")
        st.caption("İki dosya hazır olduğunda kontroller ve eşleştirme otomatik çalışır.")
    if invoice_upload is None or payment_upload is None:
        with welcome:
            with st.container(border=True, key="demo-callout"):
                action, purpose = st.columns([1, 2], vertical_alignment="center")
                action.button("Örnek raporu incele", on_click=activate_demo, type="primary", width="stretch", key="main_demo")
                purpose.markdown("Faturaları ödemelerle eşleştir, kalan borcu ve gecikmeleri hayalî verilerle incele.")
        st.session_state.pop("report_bundle", None)
        st.session_state.pop("selected_invoice", None)
        st.session_state.pop("detail_context", None)
        st.caption("Kendi raporun için fatura ve ödeme dosyalarını birlikte yükle.")
        st.stop()
    invoice_bytes, payment_bytes = invoice_upload.getvalue(), payment_upload.getvalue()
    invoice_name, payment_name = invoice_upload.name, payment_upload.name

signature = (source, sha256(invoice_bytes).hexdigest(), invoice_name, sha256(payment_bytes).hexdigest(), payment_name, as_of.isoformat())
bundle = st.session_state.get("report_bundle")
if bundle is None or bundle["signature"] != signature:
    with st.spinner("Dosyalar kontrol ediliyor ve rapor hazırlanıyor…"):
        report, issues = process_files(invoice_bytes, invoice_name, payment_bytes, payment_name, as_of)
        bundle = {"signature": signature, "report": report, "issues": issues}
        st.session_state["report_bundle"] = bundle
    clear_filters()
    st.session_state.pop("detail_context", None)

report, issues = bundle["report"], bundle["issues"]
if issues:
    st.error(f"{len(issues)} veri hatası bulundu. Nihai rapor oluşturulmadı; kayıtları düzeltip dosyaları yeniden yükle.")
    st.subheader("Veri hataları")
    errors = issues_frame(issues)
    st.dataframe(errors.map(safe_text), hide_index=True, width="stretch")
    st.download_button("Hataları CSV indir", csv_bytes(errors), "veri_hatalari.csv", "text/csv")
    st.stop()

st.caption(f"Raporlama: {as_of:%d.%m.%Y} · {len(report.faturalar)} fatura · {len(report.eslesen_odemeler)} eşleşen ödeme")
invoices = frame(report.faturalar, INVOICE_RESULT_COLUMNS)
# Yerini önce ayır, içeriğini güncel filtreler okunduktan sonra doldur.
summary = st.container(key="summary-panel")
with st.container(border=True, key="filter-grid"):
    search, reset = st.columns([4, 1], vertical_alignment="bottom")
    search.text_input("Fatura no veya müşteri ara", placeholder="Örneğin: 0001 veya Çınar", key="invoice_search")
    reset.button("Filtreleri temizle", on_click=clear_filters, key="clear_filters", width="stretch")
    filter_columns = st.columns(3)
    with filter_columns[0]:
        customers = st.multiselect("Müşteri", sorted(invoices["musteri"].unique()), placeholder="Tüm müşteriler", key="customer_filter")
    with filter_columns[1]:
        statuses = st.multiselect("Durum", STATUSES, placeholder="Tüm durumlar", key="status_filter")
    with filter_columns[2]:
        buckets = st.multiselect("Gecikme", ("Gecikme yok", *AGING_BUCKETS), placeholder="Tüm gecikme grupları", key="aging_filter")
    st.caption("Arama ve filtreler kartları, grafiği, faturaları ve CSV'yi etkiler. Excel ve diğer ödeme tabloları tüm raporu kapsar. Boş seçim tüm kayıtları gösterir.")

query = st.session_state["invoice_search"]
filtered = filter_invoices(invoices, customers, statuses, buckets, query)
rows = filtered.to_dict("records")
with summary:
    st.subheader("Rapor özeti")
    if customers or statuses or buckets or query.strip():
        st.caption(f"Filtre uygulanıyor · {len(filtered)} / {len(invoices)} fatura · Kartlar aşağıdaki seçimleri yansıtır.")
    else:
        st.caption(f"Tüm faturalar · {len(invoices)} fatura")
    with st.container(key="summary"):
        for column, (label, amount) in zip(st.columns(5), summarize(rows).items()):
            help_text = "Fazla ödeme dahildir. Eşleşmeyen ödemeler dahil değildir." if label == "Eşleşen ödeme" else None
            column.metric(label, format_tl(amount), help=help_text)
    st.caption("Eşleşen ödeme fazla ödemeyi içerir. Fazla ödemeler başka faturaların borcundan düşülmez.")
if report.eslesmeyen_odemeler:
    unmatched_total = sum(item["tutar_kurus"] for item in report.eslesmeyen_odemeler)
    st.warning(f"{len(report.eslesmeyen_odemeler)} eşleşmeyen ödeme · {format_tl(unmatched_total)}. Tahsilat kartlarına dahil değildir. [Ödemeleri incele ↓](#eslesmeyen-odemeler)")

st.subheader("Fatura sonuçları")
st.caption(f"{len(filtered)} / {len(invoices)} fatura · Tutarlar TL cinsindedir. Tutar başlıklarıyla sayısal sıralayabilirsin.")
detailed = st.checkbox("Detaylı sütunları göster", key="detailed_columns")
if filtered.empty:
    st.info("Bu tarih ve filtrelerde gösterilecek fatura yok.")
else:
    show_invoice_table(filtered, detailed)

# Seçim satır sırasına değil kimliğe bağlıdır; veri/tarih/filtre değişince temizlenir.
detail_context = (signature, tuple(customers), tuple(statuses), tuple(buckets), query)
if st.session_state.get("detail_context") != detail_context:
    st.session_state["selected_invoice"] = None
    st.session_state["detail_context"] = detail_context
if not filtered.empty:
    labels = {row["fatura_no"]: f"{row['fatura_no']} · {row['musteri']}" for row in rows}
    selected = st.selectbox("Ödeme detayını incelemek için fatura seç", list(labels), index=None,
                            format_func=labels.get, placeholder="Listeden bir fatura seç", key="selected_invoice")
    if selected is not None:
        show_invoice_detail(report, selected)

show_aging_chart(rows)

st.subheader("Eşleşmeyen ödemeler", anchor="eslesmeyen-odemeler")
st.caption("Tüm rapor kapsamı · Arama ve filtrelerden bağımsızdır. Bu ödemeler fatura tahsilatına eklenmez.")
unmatched = frame(report.eslesmeyen_odemeler, PAYMENT_COLUMNS + ["neden"])
if unmatched.empty:
    st.success("Eşleşmeyen ödeme yok.")
else:
    show_table(unmatched)
    st.download_button("Eşleşmeyen ödemeleri CSV indir", csv_bytes(unmatched), "eslesmeyen_odemeler.csv", "text/csv")
with st.expander(f"Gelecek tarihli kayıtlar ({len(report.gelecek_faturalar) + len(report.gelecek_odemeler)})"):
    st.caption("Tüm rapor kapsamı · Raporlama tarihinden sonraki kayıtlar hesaplara dahil edilmez.")
    st.markdown("**Gelecek faturalar**")
    show_table(frame(report.gelecek_faturalar, INVOICE_COLUMNS))
    st.markdown("**Gelecek ödemeler**")
    show_table(frame(report.gelecek_odemeler, PAYMENT_COLUMNS))
with st.expander(f"Eşleşen ödemeler ({len(report.eslesen_odemeler)})"):
    st.caption("Tüm rapor kapsamı · Arama ve filtrelerden bağımsızdır.")
    show_table(frame(report.eslesen_odemeler, PAYMENT_COLUMNS))

st.divider()
st.subheader("Sonuçları indir")
st.caption("CSV arama ve filtrelere uyan faturaların tüm alanlarını içerir. Excel tüm raporu kapsar. Detaylı sütun seçimi indirme kapsamını değiştirmez.")
left, right = st.columns(2)
with left:
    st.download_button("Filtreli faturaları CSV indir", csv_bytes(filtered), f"fatura_raporu_{as_of.isoformat()}.csv", "text/csv", width="stretch", key="filtered_csv")
with right:
    try:
        if "excel" not in bundle or bundle.get("excel_version") != EXCEL_EXPORT_VERSION:
            with st.spinner("Excel dosyası hazırlanıyor…"):
                bundle["excel"] = xlsx_bytes(report_sheets(report))
                bundle["excel_version"] = EXCEL_EXPORT_VERSION
    except ExcelPrecisionError as exc:
        st.error(str(exc))
    else:
        st.download_button("Tüm raporu Excel indir", bundle["excel"], f"eslestirme_raporu_{as_of.isoformat()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", width="stretch", key="full_excel")
st.caption("Fatura–Ödeme Eşleştirici · Sürüm 2 · TL")
