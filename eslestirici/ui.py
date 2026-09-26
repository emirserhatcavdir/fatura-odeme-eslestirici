"""Streamlit görünümü; finans hesapları reconciliation modülündedir."""

import pandas as pd
import streamlit as st

from .exports import display_frame, frame, safe_text
from .money import format_tl
from .reconciliation import AGING_BUCKETS, aging_totals
from .views import invoice_detail, invoice_table

FILTER_KEYS = ("customer_filter", "status_filter", "aging_filter")


def clear_filters() -> None:
    for key in FILTER_KEYS:
        st.session_state[key] = []
    st.session_state["invoice_search"] = ""
    st.session_state["selected_invoice"] = None


def show_table(data: pd.DataFrame, key: str | None = None) -> None:
    st.dataframe(display_frame(data).map(safe_text), hide_index=True, width="stretch", key=key)


def show_invoice_table(data: pd.DataFrame, detailed: bool) -> None:
    styled = invoice_table(data, detailed)
    # 50.000 satırlık geçerli bir dosya da Styler'ın varsayılan hücre sınırına takılmaz.
    with pd.option_context("styler.render.max_elements", max(262144, styled.data.size + 1)):
        st.dataframe(styled, hide_index=True, width="stretch", key="invoice_grid", lazy=False,
                     column_config={"Müşteri": st.column_config.TextColumn(width="medium")})


def show_rules() -> None:
    with st.expander("Dosya biçimi ve iş kuralları"):
        st.markdown("""
**Faturalar:** `fatura_no, musteri, fatura_tarihi, vade_tarihi, tutar`

**Ödemeler:** `odeme_id, fatura_no, odeme_tarihi, tutar`

- CSV: UTF-8, virgülle ayrılmış. XLSX: ilk çalışma sayfası okunur. Her veri türü için ayrı dosya yükle.
- Tutar: `1250.50`; binlik ayırıcı yok, en fazla iki ondalık basamak. Tarih: `YYYY-MM-DD`.
- Kimlikler metindir. Excel'de baştaki sıfırları korumak için **Metin** hücre biçimini kullan.
- Fatura numaraları birebir eşleşir; baştaki sıfırlar ve harf büyüklükleri korunur. Arama esnekliği eşleştirme kuralını değiştirmez.
- Farklı ödeme kimlikleriyle yapılan kısmi ödemeler toplanır. Yinelenen ödeme veya fatura kimliği hatadır. Herhangi bir veri hatasında rapor hazırlanmaz.
- Yalnızca pozitif TL tutarları geçerlidir. Vadesi rapor günü dolan fatura gecikmiş sayılmaz. Gecikme yalnızca kalan borç içindir.
- Dosya başına 10 MB / 50.000 satır; kayıt başına en fazla 9.999.999.999,99 TL. Excel formülleri kabul edilmez.
- Para hesapları tam sayı kuruşla yapılır. Tablo Türkçe TL gösterimini sayısal veriden ayrı tutar; tutar başlığına tıklayarak sayısal sıralayabilirsin.
- CSV'de `*_kurus` tam sayı, `*_tl` noktalı ondalık metindir. Excel'de TL alanları sayısal, teknik kuruş sütunları gizlidir. Formül başlangıcı taşıyan kullanıcı metinleri güvenli hâle getirilir. Raporları sayfadaki indirme düğmelerinden al.
- Dosyalar uygulamanın çalıştığı sunucunun belleğinde işlenir. Bulut demosuna yalnızca hayalî veri yükle. İade, alacak dekontu, kur dönüşümü ve resmî muhasebe kaydı kapsam dışıdır. Gerçek SAP/banka bağlantısı ve SAP onayı yoktur.
""")


def show_invoice_detail(report, invoice_id: str) -> None:
    detail = invoice_detail(report, invoice_id)
    if detail is None:
        return
    invoice = detail["fatura"]
    with st.container(border=True, key="invoice-detail"):
        st.subheader("Seçili fatura")
        st.text(f"Fatura no: {invoice['fatura_no']}\nMüşteri: {invoice['musteri']}")
        st.caption(f"Fatura tarihi: {invoice['fatura_tarihi']:%d.%m.%Y} · Vade: {invoice['vade_tarihi']:%d.%m.%Y} · Durum: {invoice['durum']}")
        with st.container(horizontal=True, wrap=True):
            for field, label in (("tutar_kurus", "Fatura tutarı"), ("toplam_odeme_kurus", "Toplam ödeme"), ("kalan_borc_kurus", "Kalan borç"), ("fazla_odeme_kurus", "Fazla ödeme")):
                with st.container(width=190):
                    st.caption(label)
                    st.write(format_tl(invoice[field]))
        st.markdown("**Rapora dahil ödemeler**")
        st.caption(f"{report.rapor_tarihi:%d.%m.%Y} dahil gerçekleşen hareketler.")
        columns = ["odeme_id", "odeme_tarihi", "tutar_kurus"]
        if detail["odemeler"]:
            show_table(frame(detail["odemeler"], columns), key="detail_payments")
        else:
            st.info("Bu faturanın raporlama tarihine kadar gerçekleşmiş ödemesi yok.")
        st.markdown("**Raporlama tarihinden sonraki ödemeler**")
        if detail["gelecek_odemeler"]:
            st.info("Aşağıdaki gelecek tarihli ödemeler toplam ödemeye, kartlara ve rapor bakiyelerine dahil değildir.")
            show_table(frame(detail["gelecek_odemeler"], columns), key="detail_future_payments")
        else:
            st.caption("Bu faturaya bağlı gelecek tarihli ödeme yok.")


def show_aging_chart(rows: list[dict]) -> None:
    st.subheader("Gecikmiş borç dağılımı")
    aging = aging_totals(rows)
    if not any(item["borc_kurus"] for item in aging):
        st.success("Seçili faturalarda gecikmiş borç yok.")
        return
    chart = pd.DataFrame([{**item, "tutar": format_tl(item["borc_kurus"])} for item in aging])
    st.vega_lite_chart(chart, {
        "height": 200,
        "encoding": {"y": {"field": "gecikme_grubu", "type": "nominal", "sort": list(AGING_BUCKETS), "title": None}},
        "layer": [
            {"mark": {"type": "bar", "color": "#9CD8CF", "cornerRadiusEnd": 4}, "encoding": {
                "x": {"field": "borc_kurus", "type": "quantitative", "title": "Gecikmiş borç (TL)", "axis": {"labelExpr": "format(datum.value / 100, ',.0f') + ' TL'", "tickCount": 3}},
                "tooltip": [{"field": "gecikme_grubu", "title": "Gecikme"}, {"field": "tutar", "title": "Kalan borç"}],
            }},
            {"mark": {"type": "text", "align": "left", "baseline": "middle", "fontSize": 12, "fontWeight": 600, "color": "#172B45"},
             "encoding": {"x": {"value": 8}, "text": {"field": "tutar"}}},
        ],
        "config": {"view": {"stroke": None}},
    }, width="stretch")
