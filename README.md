# Logo Tiger – e-Fatura Aracı (Excel / VBA)

Logo Tiger ERP'ye kaydedilmiş faturalar ile Logo'nun e-Fatura tablosuna (`APPROVAL`) düşen gelen/giden e-faturaları tek ekranda yan yana listeleyen, fatura numarası üzerinden eşleştiren ve tutar farklarını, işlenmemiş faturaları ve reddedilen belgeleri gösteren makrolu Excel çalışma kitabı.

NOT: Tiger'da "e-fatura al" çalıştırdıktan sonra raporu kullanabilirsiniz. Aksi takdirde güncel veri gelmez. E-Fatura bilgilerini Tiger veritabanından okur.

Ay sonu kapanışlarında "e-Fatura portalına düşen her fatura Logo'ya işlendi mi, tutarlar tutuyor mu?" sorusunu birkaç saniyede yanıtlamak için hazırlanmıştır.

---

## Özellikler

- **İki kaynak yan yana:** Solda Logo'daki faturalar (`LG_XXX_01_INVOICE`), sağda e-Fatura kayıtları (`LG_XXX_01_APPROVAL`).
- **Otomatik eşleştirme:** Her e-fatura, fatura numarasıyla Logo kaydına bağlanır.
  - **Fark** sütununda Logo tutarı ile e-fatura tutarı arasındaki fark,
  - **Durum** sütununda faturanın Logo'daki muhasebe durumu (`Muhasebeleşmiş` / `Muhasebeleşmemiş`) ya da Logo'da hiç bulunamadıysa `İşlenmemiş` yazar.
- **Özet göstergeler:** Logo toplamı, e-Fatura toplamı (reddedilenler hariç), toplam işlenen fark, işlenmemiş fatura sayısı ve RET sayısı.
- **Filtreler:** Tarih aralığı, fatura türü (alış / satış / hepsi), muhasebe durumu, cari unvan, fatura no ve işyeri (BRANCH).
- **Birden fazla firma / dönem:** Ayar sayfasındaki firma tablosundan seçilen firma numarasına göre ilgili Logo tabloları sorgulanır.
- **Kurulum gerektirmez:** Ek eklenti yoktur; Windows'ta yerleşik ADO/OLE DB sürücüsüyle doğrudan SQL Server'a bağlanır.

---

## Gereksinimler

| Gereksinim | Açıklama |
|---|---|
| İşletim sistemi | Windows (VBA + ADO kullanıldığı için macOS Excel'de çalışmaz) |
| Excel | Microsoft Excel 2010 ve üzeri (masaüstü sürüm, makro desteği) |
| Veritabanı | Logo Tiger / Tiger 3 / Go3 vb. Logo ürününün SQL Server veritabanı |
| Ağ | Kullanıcı bilgisayarından SQL Server'a (varsayılan port 1433) erişim |
| Yetki | İlgili tablolarda **yalnızca okuma (SELECT)** yetkisi olan bir SQL kullanıcısı (bkz. [Güvenlik Notları](#güvenlik-notları)) |

---

## Kurulum

1. [Releases](../../releases) bölümünden son sürüm `vk_faturasorgulama_v1.0.xlsm` dosyasını indirin.
2. Dosyaya sağ tıklayın → **Özellikler** → alttaki **Engellemeyi kaldır** kutusunu işaretleyin → **Tamam**.
   (İnternetten indirilen dosyalarda Windows makroları varsayılan olarak engeller.)
3. Dosyayı açın ve sarı çubukta **İçeriği Etkinleştir**'e tıklayın.
4. **Ayar** sayfasında bağlantı bilgilerini girin:

   | Hücre | Alan | Örnek |
   |---|---|---|
   | `K2` | Sunucu IP Adresi | `10.0.0.10` veya `SUNUCU\SQLEXPRESS` |
   | `K3` | Veritabanı Adı | `TIGER_DB` |
   | `K4` | Kullanıcı Adı | `rapor_okuma` |
   | `K5` | Şifre | `••••••••` |

5. Yine **Ayar** sayfasındaki **Firma** tablosunu (`D12:E15`) kendi firmalarınıza göre düzenleyin:

   | Firma (Rapor'da görünen ad) | Firma No (Logo firma numarası, 3 hane) |
   |---|---|
   | ABC A.Ş. 2025 | 025 |
   | ABC A.Ş. 2026 | 026 |

   > Firma numarası, Logo tablo adlarındaki numaradır (`LG_026_01_INVOICE` → `026`). Başındaki sıfırların kaybolmaması için hücre biçimi **Metin** olmalıdır.

---

## Kullanım

**Rapor** sayfasında üst bölümdeki alanları doldurun ve **Listele** butonuna basın.

| Alan | Hücre | Açıklama |
|---|---|---|
| Başlangıç Tarihi | `C2` | Sorgu başlangıç tarihi |
| Bitiş Tarihi | `C3` | Sorgu bitiş tarihi |
| Cari Unvan | `E2` | Unvanın bir kısmı (boş = hepsi) |
| Fatura No | `E3` | Fatura numarasının bir kısmı (boş = hepsi) |
| Fatura Türü | `G2` | `1: Alış Faturası`, `2: Satış Faturası` veya `Hepsi` |
| Muhasebe Durumu | `G3` | `Muhasebeleşmiş`, `Muhasebeleşmemiş` veya `Hepsi` |
| Şirket | `J2` | Ayar sayfasındaki firma tablosundan seçim |
| İşyeri No | `J3` | Logo işyeri (BRANCH) numarası; **boş bırakılırsa tüm işyerleri** listelenir |

### Sonuç sütunları

**LOGO (B–G):** Tarih · Fatura Türü · Fatura No · Cari Unvan · Tutar (net toplam) · Muhasebe Durumu

**E-FATURA (I–N):** Tarih · Fatura No · Cari Unvan · Tutar · Fark · Durum

### Özet hücreleri

| Hücre | Gösterge |
|---|---|
| `L2` | Logo toplamı |
| `L3` | e-Fatura toplamı (RET durumundakiler hariç) |
| `L1` | Logo − e-Fatura farkı |
| `M3` | Eşleşen faturalardaki toplam tutar farkı |
| `N3` | Logo'da bulunamayan (işlenmemiş) e-fatura sayısı |
| `O3` | Reddedilen (RET) e-fatura sayısı |

### Fatura türü filtresinin karşılığı

| Seçim | Logo `TRCODE` değerleri |
|---|---|
| 1: Alış Faturası | 1, 2, 3, 4 |
| 2: Satış Faturası | 6, 7, 8, 9, 13, 14 |
| Hepsi | Filtre uygulanmaz |

---

## Nasıl çalışır?

1. **Logo tarafı:** `LG_XXX_01_INVOICE` tablosu `LG_XXX_CLCARD` ile birleştirilerek seçili tarih aralığındaki faturalar çekilir. Banka kaynaklı kayıtlar (`FROMBANK = 0`) dışarıda tutulur. Fatura numarasındaki `/` karakteri `0`'a çevrilir; Logo'da aynı numaralı faturalar ayırt edilirken araya `/` eklenebildiği için bu dönüşüm e-fatura numarasıyla eşleşmeyi sağlar.
2. **e-Fatura tarafı:** `LG_XXX_01_APPROVAL` tablosundan aynı tarih aralığındaki, belge numarası dolu kayıtlar çekilir. `RECTYPE` 114, 115, 125 ve 127 olan kayıtlar hariç tutulur. Döviz cinsinden faturalarda yerel para karşılığı (`DOCTOTALLOCAL`) kullanılır.
3. **Eşleştirme:** Her e-fatura numarası Logo listesinde aranır. Bulunursa tutar farkı ve muhasebe durumu yazılır; bulunamazsa `İşlenmemiş` olarak işaretlenir.

VBA kaynak kodu, makrolu dosyayı açmadan incelenebilsin diye [`src/`](src/) klasöründe düz metin olarak da yer alır.

---

## Güvenlik Notları

Bu araç doğrudan muhasebe veritabanınıza bağlanır. Lütfen aşağıdaki maddeleri okumadan canlı ortamda kullanmayın.

### 1. `sa` veya yönetici hesabı kullanmayın

Araç yalnızca `SELECT` sorgusu çalıştırır. Bu yüzden sadece gerekli tablolarda okuma yetkisi olan ayrı bir SQL kullanıcısı tanımlayın. Örnek (SQL Server Management Studio'da, sistem yöneticisi tarafından çalıştırılmalıdır):

```sql
-- Sunucu düzeyinde oturum
CREATE LOGIN rapor_okuma WITH PASSWORD = 'GüçlüBirŞifre!2026', CHECK_POLICY = ON;
GO

-- Logo veritabanında kullanıcı
USE [TIGER_DB];
CREATE USER rapor_okuma FOR LOGIN rapor_okuma;
GO

-- Yalnızca ihtiyaç duyulan tablolara okuma yetkisi (her firma/dönem için tekrarlayın)
GRANT SELECT ON dbo.LG_026_01_INVOICE  TO rapor_okuma;
GRANT SELECT ON dbo.LG_026_CLCARD      TO rapor_okuma;
GRANT SELECT ON dbo.LG_026_01_APPROVAL TO rapor_okuma;
GO
```

> Tablo bazında yetki vermek, veritabanının tamamına okuma yetkisi (`db_datareader`) vermekten daha güvenlidir. Bordro, personel ve banka tabloları bu kullanıcıya açılmamış olur.

### 2. Şifre hücrede düz metin olarak durur

- Ayar sayfasındaki şifre (`K5`) **şifrelenmeden** saklanır. Dosyayı açabilen herkes görebilir.
- Bağlantı bilgileri doldurulmuş dosyayı **e-posta, bulut klasörü veya repo üzerinden kimseyle paylaşmayın.** Paylaşacaksanız önce `K2:K5` hücrelerini temizleyin.
- Sayfa gizleme veya sayfa koruması güvenlik önlemi **değildir**; kolayca aşılır. Gerçek koruma için dosyayı parolayla şifreleyin: **Dosya → Bilgi → Çalışma Kitabını Koru → Parolayla Şifrele**.
- Bu repoya **asla** doldurulmuş bir dosya göndermeyin (commit/push etmeyin).

### 3. Filtre alanları doğrudan SQL'e eklenir

Cari Unvan, Fatura No ve tarih alanlarındaki değerler sorguya metin olarak eklenir (parametreli sorgu kullanılmaz). Bunun iki sonucu vardır:

- Değerlerde tek tırnak (`'`) gibi karakterler varsa sorgu hata verebilir.
- Kötü niyetli bir kullanıcı bu alanlara SQL ifadesi yazabilir (SQL injection).

Bu nedenle aracı yalnızca güvendiğiniz kişilere verin ve 1. maddedeki **yalnızca okuma yetkili** kullanıcıyla çalıştırın. Böylece en kötü senaryoda bile veri değiştirilemez veya silinemez. İşyeri No alanı sayısal kontrol edildiği için bu risk o alanda yoktur.

### 4. Ağ güvenliği

- SQL Server'ı **internete açmayın**. Araç yalnızca şirket iç ağında veya VPN üzerinden kullanılmalıdır.
- Varsayılan `SQLOLEDB` sağlayıcısı bağlantıyı şifrelemez. Bağlantı şifrelemesi istiyorsanız Microsoft'un güncel **OLE DB Driver for SQL Server** (`MSOLEDBSQL`) sürücüsünü kurup VBA'daki bağlantı cümlesini şu şekilde değiştirebilirsiniz:

  ```vba
  conn.Open "Provider=MSOLEDBSQL;Data Source=" & sunucu & ";Initial Catalog=" & veritabani & _
            ";User Id=" & kullanici & ";Password=" & sifre & ";Encrypt=yes;TrustServerCertificate=yes;"
  ```

### 5. Makro güvenliği

- Dosyayı yalnızca bu reponun **Releases** bölümünden indirin. Başka kaynaklardan gelen kopyalar değiştirilmiş olabilir.
- Açmadan önce isterseniz [`src/`](src/) klasöründeki kodu inceleyin. Makro yalnızca SQL'den veri okur ve Rapor sayfasına yazar; dosya sistemine yazmaz, internete veri göndermez, başka program çalıştırmaz.
- Makroları genel olarak açmak yerine dosyayı **Güvenilir Konum** olarak tanımlanmış bir klasörde tutmanız önerilir (**Dosya → Seçenekler → Güven Merkezi → Güvenilir Konumlar**).

### 6. Kişisel veri ve ticari bilgi (KVKK)

- Rapor çıktısı cari unvanlar, fatura numaraları ve tutarlar içerir. Gerçek kişi tacirlerde (şahıs firmalarında) unvan **kişisel veri** niteliğindedir.
- Rapor alınmış dosyayı, ekran görüntülerini veya örnek çıktıları paylaşmadan önce `B7:O10000` aralığını temizleyin. README veya issue'lara görsel eklerken **örnek/sahte veri** kullanın.
- Hata bildirirken (issue açarken) gerçek firma, cari, sunucu veya veritabanı bilgisi paylaşmayın.

---

## Bilinen Kısıtlar

- En fazla **10.000** satır listelenir (`B7:N10000`).
- Tutar farkı hesaplaması Logo listesinin ilk **1.000** satırında arama yapar (`D7:F1000`); daha büyük listelerde bazı farklar boş kalabilir. Gerekirse VBA'daki aralığı `D7:F10000` olarak genişletin.
- Eşleştirme yalnızca fatura numarasına göre yapılır. Logo'da fatura numarası e-fatura numarasından farklı girilmişse kayıt `İşlenmemiş` görünür.
- Firma tablosu 4 satırlıktır (`Ayar!D12:E15`). Daha fazla firma eklemek için Rapor sayfasındaki `J2` veri doğrulama listesini ve `H5` hücresindeki `DÜŞEYARA` aralığını da genişletmeniz gerekir.
- Logo sürümleri arasında `APPROVAL` tablosundaki `RECTYPE` / `STATUS` kodları farklılık gösterebilir. Kendi sürümünüzde sonuçları birkaç fatura üzerinde elle kontrol edin.

---

## Repo yapısı

```
├── vk_faturasorgulama_v1.0.xlsm   # Kullanıma hazır çalışma kitabı (Releases'ta da yer alır)
├── src/
│   └── vk_faturasorgulama_v1.0.cls             # Rapor sayfasının VBA kodu (düz metin)
├── README.md
└── LICENSE
```

---

## Sorumluluk Reddi

Bu araç muhasebe kontrolünü kolaylaştırmak için hazırlanmış yardımcı bir yazılımdır. Resmî bir Logo ürünü değildir ve Logo Yazılım ile bir bağlantısı yoktur. Sonuçların doğruluğu kullanılan Logo sürümüne ve veri girişine bağlıdır. Muhasebe ve vergi kararlarında tek başına dayanak olarak kullanılmamalıdır. Yazılım "olduğu gibi" sunulur; ayrıntılar için [LICENSE](LICENSE) dosyasına bakın.

---

## Lisans

[MIT Lisansı](LICENSE) ile dağıtılmaktadır. Ticari kullanım dahil serbestçe kullanabilir, değiştirebilir ve dağıtabilirsiniz; tek şart telif ve lisans metninin korunmasıdır.

---

**Geliştirici:** [github.com/kocvolkan38](https://github.com/kocvolkan38)
