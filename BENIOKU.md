# Ekspa İş Takibi — Satış + Teknik Servis (Sürüm 4)

Ekspa Fırsat Takibi ve Çağrı Karşılama artık **tek uygulama**: tek adres, tek giriş, tek kullanıcı listesi.
Üstteki **Satış / Teknik Servis** düğmesiyle bölümler arasında geçilir.

- **Satış:** Pano, Liste, Özet, Projeler (fırsat ve kazanılan proje takibi)
- **Teknik Servis:** Yeni çağrı, Çağrılar, Envanter, Firmalar, Özet, E-postalar, Ayarlar, Kullanıcılar

Gereksinim: **Python 3.8+**, ek kurulum yok. Varsayılan adres: `http://SUNUCU:8080`.

## Mevcut kurulumdan geçiş (verileriniz otomatik taşınır)

1. Çalışan uygulamaları durdurun (Ekspa Fırsat Takibi ve kurduysanız Çağrı Karşılama).
2. **Yedek alın:** `firsatlar.db` (ve varsa `cagri.db`) dosyalarını yanlarındaki `-wal`, `-shm` ile birlikte kopyalayın.
3. Bu paketi yeni bir klasöre açın (örn. `C:\ekspa`).
4. Eski Ekspa klasöründeki `data\firsatlar.db` dosyasını (ve `-wal`, `-shm` varsa onları) yeni klasörün `data\` içine kopyalayın.
   Çağrı Karşılama'yı da kurup kullandıysanız onun `data\cagri.db` dosyasını da aynı yere kopyalayın.
5. `baslat.bat` ile başlatın. Ekranda şunları görmelisiniz:
   - "Ekspa Fırsat Takibi veritabanı (firsatlar.db) kopyalandı"
   - "Fırsat Takibi kullanıcıları yeni yetki yapısına taşındı."
   - (cagri.db varsa) "Çağrı Karşılama verileri ve kullanıcıları taşındı."
6. Kullanıcılar aynı kullanıcı adı ve şifreyle girer. Tarayıcıda bir kez yeniden giriş yapmaları gerekir.

Yeni veritabanı `data\ekspa.db`'dir. Eski `firsatlar.db` / `cagri.db` dosyalarına dokunulmaz; yedek olarak durur.

**Kullanıcı taşıma kuralları**
| Eski | Yeni |
|---|---|
| Fırsat Takibi Yönetici | Yönetici (iki bölüm) |
| Fırsat Takibi Düzenleyici | Satış: Düzenleyici · Servis: erişim yok |
| Fırsat Takibi Görüntüleyici | Satış: Görüntüleyici · Servis: erişim yok |
| Çağrı Karşılama Uzman / Görüntüleyici | Servis: Uzman / Görüntüleyici (Satış'ta erişim yok) |
| İki uygulamada aynı kullanıcı adı | Tek hesap; şifre Fırsat Takibi'ndeki kalır, servis yetkisi ve e-posta eklenir |

Taşımadan sonra **Teknik Servis → menü → Kullanıcıları yönet** ekranından herkesin Satış ve Servis yetkilerini kontrol edin.

**Cloudflare Tunnel:** Uygulama aynı port (8080) üzerinden çalıştığı için mevcut tünel ayarınız değişmez.
Ayrıca Çağrı Karşılama için ayrı bir adres açtıysanız onu kaldırabilirsiniz. `EKSPA_HTTPS=1` ve `--host 127.0.0.1` önerisi geçerlidir:
```
set EKSPA_HTTPS=1
py -3 app.py --host 127.0.0.1
```
Ayarlar'daki "Uygulama adresi"ni tünel adresiyle (örn. `https://is.ekspa.com.tr/servis`) güncelleyin.

## Sıfırdan kurulum

Paketi açıp `baslat.bat` çalıştırın; tarayıcıdan ilk yönetici hesabını oluşturun. Başlangıç verisi olarak
143 fırsat, 2.262 envanter ürünü ve 68 yardım masası kaydı yüklenir.

## Yetkiler

| | Satış | Teknik Servis |
|---|---|---|
| **Yönetici** | Tam yetki | Tam yetki + garanti dışı çağrı, Ayarlar, Kullanıcılar |
| **Düzenleyici** (Satış) | Fırsat/proje ekler, düzenler, siler | — |
| **Uzman** (Servis) | — | Çağrı açar ve günceller, envanteri düzenler, firmalara e-posta girer; çağrı atanabilir |
| **Görüntüleyici** | Yalnızca görür | Yalnızca görür |

Her kullanıcıya iki bölüm için ayrı yetki verilir (biri "Erişim yok" olabilir). Yetkisi olmayan bölüm menüde görünmez.

## Bu sürümdeki değişiklikler (Teknik Servis)

- **"Cihaz müşteriye iade edildi"** aşaması: onarılan cihaz müşteriye teslim edildi. Sıra: … Çözüldü → Cihaz müşteriye iade edildi → Kapandı.
  Kapalı aşamalar arasında sayılır ve çözüm açıklaması gerektirir. Müşteriye aşama değişikliği e-postası gider.
- **Garanti tarihi girilmemiş ürünlerde çağrı açılabilir.** Çağrıda "Garanti bilgisi yok" etiketi ve geçmişte
  "Garanti doğrulanamadı" notu yer alır. Garantisi **bitmiş** ya da **henüz başlamamış** ürünlerde çağrıyı yine yalnızca
  yönetici, gerekçe yazarak açabilir.
- **Envanter alanları uzmanlar tarafından düzenlenebilir** (seri no, firma, marka, parça no, tanım, yetkili servis,
  garanti başlangıç/bitiş, garanti bilgisi). Yeni çağrı ekranında ürün seçildikten sonra
  **"Ürün ve garanti bilgilerini düzenle"** ile tarihleri girip garanti durumunu anında yeniden kontrol edebilirsiniz.

## Teknik Servis kullanımı (özet)

1. **Yeni çağrı** → seri numarasını kontrol edin → ürünü seçin → konu, uzman, müşteri e-postası → **Çağrıyı aç**.
2. Uzmana ve müşteriye e-posta gider. Aşama ve bilgi değişikliklerinde müşteriye bildirim gider; iç notlar gitmez.
3. **Ayarlar:** SMTP (Microsoft 365: `smtp.office365.com`, 587, STARTTLS), gönderen adres, hedef çözüm süresi (iş günü),
   bildirim seçenekleri, test e-postası. Ayar yapılmadan oluşan bildirimler saklanır, sonradan gönderilir.
4. **Firmalar:** firma başına bildirim e-postaları ve telefon; yeni çağrıda otomatik dolar.

Aşamalar: Çağrı açıldı · İnceleniyor · Parça sipariş edildi · Parça teslim alındı · Müdahale / onarım ·
Müşteri yanıtı bekleniyor · Çözüldü · Cihaz müşteriye iade edildi · Kapandı · İptal edildi.

## Satış kullanımı (özet)

Pano (sürükle-bırak aşama), Liste, Özet, **Projeler** (kazanılan fırsatlar: sözleşme tarihleri ve tutarı, maliyet,
faturalar/tahsilat, tedarikçi ödemeleri), Excel ve CSV indirme. Ayrıntılar önceki sürümlerle aynıdır.

## Komut satırı

```
python app.py --kullanicilar
python app.py --sifre-sifirla levent
python app.py --kullanici-ekle ali --satis duzenleyici --servis yok
python app.py --kullanici-ekle dogan --servis uzman
python app.py --kullanici-ekle ayse --yonetici

pip install openpyxl     (yalnızca Excel içe aktarma için)
python app.py --import Ekspa-Proje.xlsx                       (Odoo CRM fırsatları)
python app.py --import-envanter "Teknik Servis Envanter.xlsx"
python app.py --import-talepler "Yardım Masası Talebi.xlsx"
```

## Yedekleme

Tüm veriler `data\ekspa.db` dosyasındadır (yanındaki `-wal`, `-shm` ile birlikte yedekleyin).
Şifreler geri döndürülemez özet olarak saklanır; SMTP şifresi e-posta gönderebilmek için veritabanında tutulur.

## Sürekli çalıştırma

- **Windows:** Görev Zamanlayıcı → "Bilgisayar başladığında" → `baslat.bat`; "Kullanıcı oturum açmış olsun ya da olmasın çalıştır".
- **Linux:** `ekspa.service` dosyasını `/etc/systemd/system/` altına kopyalayıp `systemctl enable --now ekspa`.
- **Docker:** `docker build -t ekspa .` ve `docker run -d -p 8080:8080 -v ekspa-veri:/veri --restart unless-stopped ekspa`
  (geçiş için eski `firsatlar.db`/`cagri.db` dosyalarını birime `/veri` altına koyun).

## Railway'de yayınlama

Uygulama Railway'i kendiliğinden tanır: `PORT` değişkenini dinler, veriyi bağlı Volume'a yazar, oturum çerezini
`Secure` işaretler (TLS'i Railway sonlandırır), hatalı giriş kilidinde gerçek istemci IP'sini (`X-Real-IP`) kullanır
ve saat dilimi olarak Türkiye saatini kullanır (Dockerfile'da `TZ=Europe/Istanbul`).

1. **Kodu Railway'e gönderin** (ikisinden biri):
   - GitHub'da **özel (private)** bir depo açıp bu klasörü yükleyin → Railway'de *New Project → Deploy from GitHub repo*.
     `.gitignore` canlı veritabanını (`data/ekspa.db*`) depoya almaz.
   - Ya da Railway CLI ile bu klasörden: `railway login`, `railway init`, `railway up`.
   Railway kökteki `Dockerfile`'ı otomatik kullanır.
2. **Volume ekleyin (zorunlu):** proje tuvalinde sağ tık ya da Ctrl/⌘+K → *Volume* → servise bağlayın, **Mount path: `/veri`**.
   Volume yoksa veritabanı her deploy'da silinir; bu durumda loglarda `!! UYARI: ... Volume bağlı değil` yazar.
3. **Adres alın:** servis → *Settings → Networking → Generate Domain*. Port sorulursa deploy logundaki
   `Ekspa İş Takibi çalışıyor: ... (port XXXX ...)` satırındaki değeri yazın. İsterseniz kendi alan adınızı da (örn. `is.ekspa.com.tr`) buradan bağlayabilirsiniz.
4. (İsteğe bağlı) *Settings → Deploy → Healthcheck Path:* `/api/kurulum`
5. **İlk yönetici hesabı:** internete açık kurulumda adresi ilk açan kişi yönetici olamasın diye ilk kurulum bir
   **kurulum kodu** ister. Kod deploy loglarında yazar: `>> Kurulum kodu: XXXX-XXXX-XXXX` (servis → *Deployments → View logs*).
   Adresi açın, kodu ve yönetici bilgilerini girin. Kod her yeniden başlatmada değişir; sabit kod isterseniz
   *Variables*'a `EKSPA_KURULUM_KODU` ekleyin. Yönetici oluşturulduktan sonra kod bir daha sorulmaz.
6. **Teknik Servis → Ayarlar → Uygulama adresi:** `https://<railway-adresi>/servis` (boş bırakılırsa e-postalardaki
   bağlantılar için Railway adresi otomatik kullanılır).

**E-posta uyarısı:** Railway, **Free/Trial/Hobby** planlarda dışarıya SMTP bağlantısını (25, 465, 587, 2525) engeller;
SMTP yalnızca **Pro** ve üstünde açıktır. Hobby planda çağrı bildirim e-postaları gönderilemez (E-postalar ekranında "hata"
olarak görünür). Çözüm: Pro plan ya da HTTPS üzerinden e-posta gönderen bir servis (örn. Microsoft Graph) entegrasyonu.

**Notlar**
- SQLite kullanıldığı için servis **tek kopya (1 replica)** çalışmalıdır; Volume bağlı serviste Railway zaten böyle yapar.
- Yeniden deploy sırasında birkaç saniyelik kesinti olur (Volume bağlı servislerde Railway'in davranışı).
- Veritabanı Volume içinde `/veri/ekspa.db` dosyasındadır; düzenli yedek almayı unutmayın (Volume'un *Backups* sekmesi ya da CLI).
- Komut satırı işlemleri (`--kullanicilar`, `--sifre-sifirla` vb.) için Railway CLI ile `railway ssh` açıp
  `python app.py --sifre-sifirla levent` gibi çalıştırabilirsiniz.

## Klasör yapısı

```
app.py                          Sunucu (Satış + Teknik Servis API, oturum, e-posta kuyruğu, taşıma)
static/satis.html               Satış arayüzü (Pano, Liste, Özet, Projeler)
static/servis.html              Teknik Servis arayüzü (Çağrılar, Envanter, Firmalar, Ayarlar, Kullanıcılar)
static/giris.html               Giriş ve ilk kurulum
static/xlsx.full.min.js         Excel indirme kütüphanesi (SheetJS 0.18.5, Apache-2.0)
data/baslangic_*.json           Sıfırdan kurulum için başlangıç verileri
data/ekspa.db                   Canlı veritabanı (ilk çalıştırmada oluşur)
baslat.bat / baslat.sh, ekspa.service   Windows/Linux başlatıcıları
Dockerfile, .dockerignore, .gitignore Docker / Railway için
```
