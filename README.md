# parasut-converter

Şirketimiz Uyumsoft'tan Paraşüt'e geçerken fatura aktarımında ciddi bir zaman kaybı yaşıyordu. İki platformun şablonları birbiriyle uyumlu olmadığı için 4.700+ faturayı manuel girmek gerekiyordu. Bunu otomatikleştirmek için yazdım.

Aynı sorunu yaşayan varsa işine yarar.

---

## Ne yapıyor?

Uyumsoft'un **Giden** ve **Gelen Faturalar** Excel çıktılarını alıp Paraşüt'ün içe aktarma şablonlarına dönüştürüyor.

Basit  detayları hallediyor:
- Sütun isimlerini eşliyor (`Alıcı` → `MÜŞTERİ ÜNVANI`, `Gönderici` → `Tedarikçi` vs.)
- Tarihleri string'den gerçek Excel datetime nesnesine çeviriyor (Paraşüt metin kabul etmiyor)
- TRL faturalarda döviz kuru alanını boş bırakıyor (dolu olursa hata veriyor)
- Şablonun içindeki örnek satırları temizliyor
- VKN/TCKN ve Fatura No'daki baştaki `'` karakterini siliyor

## Kurulum

```bash
pip install pandas openpyxl
```

## Kullanım

### Giden Faturalar → Paraşüt Satış Faturası

**1. Dosyaları hazırla:**
- Uyumsoft → Raporlar → **Giden Faturalar** → Excel olarak indir
- Paraşüt → Satış Faturaları → İçeri Aktar → "Excel Şablonunu İndir"
- İkisini de bu klasöre koy

**2. Çalıştır:**
```bash
python convert.py
# veya özel dosya adlarıyla:
python convert.py --kaynak "giden.xlsx" --sablon "parasut_satis_sablonu.xlsx" --cikti "yuklenecek.xlsx"
```

**3. Yükle:**  
Paraşüt → Satış Faturaları → İçeri Aktar → **Şablonu Geri Yükle**

---

### Gelen Faturalar → Paraşüt Alış Faturası

**1. Dosyaları hazırla:**
- Uyumsoft → Raporlar → **Gelen Faturalar** → Excel olarak indir
- Paraşüt → Alış Faturaları → İçeri Aktar → "Excel Şablonunu İndir"
- İkisini de bu klasöre koy

**2. Çalıştır:**
```bash
python convert_gelen.py
# veya özel dosya adlarıyla:
python convert_gelen.py --kaynak "gelen.xlsx" --sablon "parasut_alis_sablonu.xlsx" --cikti "gelen_yuklenecek.xlsx"
```

**3. Yükle:**  
Paraşüt → Alış Faturaları → İçeri Aktar → **Şablonu Geri Yükle**

---

## Notlar

- Tüm faturalar TRL olarak işleniyor
- Giden faturalar: KDV `%20` sabit — Paraşüt birim fiyatı otomatik KDV'li yapmasın diye `Tutar / 1.20` şeklinde hesaplanarak (KDV hariç) aktarılır. Farklıysa `convert.py` içinde değiştirin.
- Vade tarihi boş bırakılıyor → açık fatura olarak giriyor
- Fatura sıra no boş → Paraşüt otomatik atıyor
- Gelen faturalarda "Toplam KDV" sütunu otomatik hesaplanır, "Toplam Tutar" KDV dahil tutar olarak gönderilir.

---

## Ekstra: Yanlış Yüklenen Faturaları Toplu Silme Aracı (`delete_invoices.py`)

Paraşüt web arayüzünde "Tümünü Seç" ile 90 sayfa faturayı tek tek silmek tam bir eziyettir. Yanlış bir yükleme yaparsanız Paraşüt API'sini kullanarak hepsini tek tuşla silebilirsiniz.

**Kurulum:**
1. Paraşüt destek ekibinden API anahtarlarınızı isteyin.
2. `.env.example` dosyasının kopyasını oluşturup adını `.env` yapın.
3. İçindeki bilgileri kendi hesap bilgilerinize göre doldurun.

**Kullanım:**
```bash
python delete_invoices.py
```
Program size hangi tür faturaları silmek istediğinizi soracak ve işlem bittikten sonra silecektir.

---

Developed by [vr0cks](https://www.vr0cks.com/tr)
