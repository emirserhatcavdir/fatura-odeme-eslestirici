# Fatura–Ödeme Eşleştirici

Fatura ve ödeme dosyalarını fatura numarasıyla eşleştiren, tahsilat durumunu ve gecikmiş borcu gösteren Türkçe bir **Python + Streamlit portföy projesi**. İktisat ve finans verilerini programlamayla inceleme amacıyla geliştirilmiştir.

Bu bağımsız bir uygulamadır. SAP veya banka bağlantısı, SAP onayı ve resmî muhasebe sistemi niteliği yoktur. `sap.py` yalnızca giriş dosyasının adıdır. Repodaki tüm örnekler ve kabul verileri tamamen hayalîdir.

**Canlı demo:** Yayın ve çalışma doğrulaması tamamlandığında bağlantı burada yer alacak.

## Bir dakikada dene

1. Uygulamayı açıp yan panelde **Örnek verilerle dene** düğmesine bas.
2. Raporlama tarihi otomatik olarak **30.06.2026** olur; dosya yüklemek gerekmez.
3. Fatura numarası veya müşteriyle ara, durum/gecikme filtrelerini kullan. Fatura seçerek parçalı ve gelecek tarihli ödemeleri incele.
4. **Filtreli faturaları CSV indir** veya **Tüm raporu Excel indir** düğmesiyle raporu al.

Yerleşik örnekte beklenen sonuçlar:

| Gösterge | Tutar |
| --- | ---: |
| Toplam fatura | 19.050,75 TL |
| Eşleşen ödeme (fazla ödeme dahil) | 7.950,50 TL |
| Kalan borç | 11.300,25 TL |
| Gecikmiş borç | 9.900,00 TL |
| Fazla ödeme | 200,00 TL |

`0005` iki parçalı ödemeyi, `0007` rapor tarihinden sonraki ödemeyi, `0003` ödemesiz faturayı gösterir. Boş şablonlar ve hayalî örnekler uygulamadan indirilebilir; dosyalar `veriler/` klasöründedir.

## Özellikler ve iş kuralları

- CSV/XLSX okuma; fatura numarasıyla **kesin eşleştirme**. `0001` ile `1`, `AB` ile `ab` farklıdır. Kimliklerin baştaki sıfırları korunur.
- Para girişleri `Decimal` ile doğrulanır, hesaplar **tam sayı kuruşla** yapılır. Aynı faturanın farklı ödeme kimlikleriyle gelen parçalı ödemeleri toplanır.
- Ödendi, kısmen ödendi, ödenmedi ve fazla ödeme durumları; kalan borç, gecikme günü ve gecikme grupları hesaplanır. Bir faturanın fazla ödemesi başka bir faturanın borcundan düşülmez.
- Raporlama tarihi dahildir. Sonraki tarihli faturalar/ödemeler ayrı gösterilir ve toplamlara katılmaz. Eşleşmeyen ödemeler de tahsilat toplamından ayrıdır.
- Yalnızca kalan borç gecikebilir. Vadesi rapor günü dolan fatura gecikmiş sayılmaz; takvim günü esas alınır.
- Eksik/bozuk veri, yinelenen fatura numarası veya ödeme kimliği raporu engeller. Aynı faturaya ait farklı ödeme kayıtları mükerrer sayılmaz.
- Arama ve filtreler kartları, grafiği, fatura tablosunu ve filtreli CSV'yi birlikte etkiler. **Excel her zaman tüm raporu içerir.** Filtreleri temizlemek yüklü verileri ve raporlama tarihini değiştirmez.

## Dosya biçimi

Her veri türü için ayrı dosya yüklenir. CSV: UTF-8, virgülle ayrılmış. XLSX: ilk çalışma sayfası. Kimlikleri Excel'de **Metin** biçiminde hazırlayın; sayısal girişte önceden kaybolmuş sıfırlar geri getirilemez.

Faturalar:

```csv
fatura_no,musteri,fatura_tarihi,vade_tarihi,tutar
0001,Hayalî Mavi Kırtasiye,2026-06-01,2026-06-10,5000.00
```

Ödemeler:

```csv
odeme_id,fatura_no,odeme_tarihi,tutar
P0001,0001,2026-06-10,3000.00
```

Tutarlar pozitif TL, ondalık ayırıcı nokta ve en fazla iki basamaktır; binlik ayırıcı kullanılmaz. Tarihler `YYYY-MM-DD` veya gerçek Excel tarih hücresidir. Vade, fatura tarihinden önce olamaz. Formüller, hata hücreleri, saat içeren tarihler ve sessiz yuvarlama kabul edilmez. Dosya başına 10 MB / 50.000 satır; kayıt başına en fazla 9.999.999.999,99 TL desteklenir.

## Excel ve CSV raporları

Excel'in ilk sayfası **Özet**; üstünde belirgin raporlama tarihi bulunur. Diğer sayfalar: Rapor Bilgisi, Faturalar, Eşleşen Ödemeler, Eşleşmeyen Ödemeler, Gelecek Faturalar, Gelecek Ödemeler, Gecikme Dağılımı.

Excel'deki TL tutarları iki ondalıklı **gerçek sayısal hücrelerdir**; `SUM`/`TOPLA` ve sayısal sıralama için kullanılabilir. Türkçe başlıklar, sütun filtreleri ve sabit başlıklar vardır. Teknik `*_kurus` sütunları silinmez, varsayılan olarak gizlenir. Kimlikler metin, tarihler `gg.aa.yyyy` görünümünde gerçek tarihtir.

CSV'de mevcut teknik alanlar, tam sayı `*_kurus` ve kesin noktalı ondalık `*_tl` metinleri korunur. CSV UTF-8 BOM içerir. Formül başlangıcı taşıyan kullanıcı metinleri güvenli hâle getirilir. Excel'in sayısal hassasiyetini aşarak kuruş kaybına yol açacak çok büyük toplamlar sessizce yuvarlanmaz; açıklayıcı hata gösterilir ve CSV kullanılabilir. Sonuç raporları giriş şablonu değildir.

## Gizlilik ve kapsam

**Herkese açık demoya gerçek müşteri verisi, kişisel bilgi veya gizli dosya yüklemeyin.** Dosyalar uygulamanın çalıştığı sunucunun belleğinde işlenir. Yerel çalışmada bu sunucu kendi bilgisayarınızdır; Community Cloud'da dosyalar bulut sunucusuna iletilir. Uygulama yüklemeleri kendi koduyla diske veya veritabanına kaydetmez. Sonuçlar Streamlit oturumunda tutulur, rapor indirmesini tarayıcı kaydeder.

Bu projede kullanıcı hesabı, veritabanı, SAP/banka API bağlantısı, gerçek tahsilat, iade/alacak dekontu veya döviz dönüşümü yoktur. API anahtarı ve `secrets.toml` gerektirmez.

GitHub'a yalnızca denetlenen kaynaklar, testler, şablonlar ve hayalî veriler alınır. `.venv`, önbellekler, yerel ayarlar, yüklemeler ve üretilmiş/indirilmiş raporlar `.gitignore` ile dışlanır.

## Yerel kurulum

Python **3.12** ile:

```powershell
cd fatura-odeme-eslestirici
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\baslat.bat
```

Kurulumdan sonra Windows'ta `baslat.bat` dosyasına çift tıklamak yeterlidir. Başlatıcı kendi klasörüne geçer, `.venv` Python'unu kullanır ve yalnızca `127.0.0.1:8501` üzerinde çalışır. Eksik sanal ortam veya dolu port için anlaşılır hata verir; paket kurmaz, süreç sonlandırmaz, hata penceresini hemen kapatmaz. Uygulama zaten çalışıyorsa aynı adresi kullanın. Alternatif port: `baslat.bat 8502`.

Linux/macOS:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run sap.py --server.address 127.0.0.1
```

Çalışma bağımlılıkları `requirements.txt`, test bağımlılıkları `requirements-dev.txt` içindedir. `requirements-lock.txt` önceki Windows doğrulama ortamının kaydıdır; bulut kurulumuna zorlanmaz. Sunucu adresini/portunu bulut platformu belirler; yerel başlatıcı ayrıca localhost sınırı uygular.

## Testler ve bağımsız kabul verileri

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m pytest -q
```

Testler para hassasiyeti, kesin eşleştirme, parçalı/fazla ödeme, tarih sınırları, veri doğrulama, arama/filtre/detay akışları ve güvenli CSV/Excel dışa aktarımını kapsar. Streamlit AppTest otomasyonu gerçek tarayıcı dosya diyaloğu testi değildir.

Yayın hazırlığında Windows / Python 3.12 ortamında **133 test geçti**. Linux bağımlılık çözümlemesi ve gerçek bulut açılışı ayrı kontrollerdir; bu sonuç bulut yayınının tamamlandığı anlamına gelmez.

`acceptance_data/faturalar.csv`, `faturalar.xlsx`, `odemeler.csv` ve `odemeler.xlsx` yerleşik demodan ayrı, tamamen hayalî 4 fatura / 6 ödeme içerir. `expected_results.json` bağımsız sabit beklentilerdir.

| Raporlama tarihi | Fatura | Eşleşen ödeme | Kalan borç | Gecikmiş borç | Fazla ödeme |
| --- | ---: | ---: | ---: | ---: | ---: |
| 30.06.2026 | 5.000,00 TL | 2.350,00 TL | 2.750,00 TL | 1.250,00 TL | 100,00 TL |
| 26.09.2026 | 5.000,00 TL | 3.850,00 TL | 1.250,00 TL | 1.250,00 TL | 100,00 TL |

30.06 tarihinde 0001 iki ödemeyle kapanır; 0002 kısmen ödenmiş ve 10 gün gecikmiştir; 0003 ödenmemiş ve gecikmemiştir; 0004 için 100 TL fazla ödeme vardır. Eşleşmeyen P005 iki tarihte de 250 TL'dir. P006 (1.500 TL), 30.06 tarihinde gelecek ödeme olarak ayrılır; 26.09 tarihinde 0003'ü kapatır. Bu tarihte 0002'nin gecikmesi 98 gündür.

```powershell
.\.venv\Scripts\python.exe -m scripts.verify_acceptance --date 2026-06-30 --output acceptance_data/reports/2026-06-30
.\.venv\Scripts\python.exe -m scripts.verify_acceptance --date 2026-09-26 --output acceptance_data/reports/2026-09-26
```

Gerçek dışa aktarma işlevleriyle üretilen raporlar yeniden okunur; hücre tipleri, baştaki sıfırlar, tarihler, Türkçe karakterler, kuruş karşılıkları, filtreli CSV ve tüm Excel kapsamı denetlenir. Üretilen raporlar GitHub'a gönderilmez.

## Streamlit Community Cloud yayını

[Community Cloud](https://share.streamlit.io/) üzerinde **Create app → Yup, I have an app** yoluyla bu repoyu seçin. Branch: `main`; main file path: **`sap.py`**; Advanced settings → Python: **3.12**. Secrets alanı boş kalır. [Resmî yayın rehberi](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy).

Kök dizindeki `requirements.txt` bulutta kurulur. [Bağımlılık rehberi](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies). Demo adresi ancak yayın, örnek veri akışı ve Excel indirmesi doğrulandıktan sonra yukarıya eklenir.

## Proje yapısı

```text
sap.py                    Streamlit giriş dosyası
eslestirici/              Okuma, doğrulama, eşleştirme, rapor ve görünüm modülleri
assets/app.css            Arayüz stilleri
tests/                    Otomatik testler
veriler/                  Hayalî örnekler, boş şablonlar ve hatalı örnekler
acceptance_data/           Bağımsız hayalî kabul girdileri
scripts/verify_acceptance.py  Dosya okuma/yazma kabul kontrolü
.streamlit/config.toml    Tema ve platformdan bağımsız ayarlar
baslat.bat / baslat.py    Windows yerel başlatıcı
```
