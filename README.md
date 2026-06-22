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
- Giden faturalar: KDV `%20` sabit — farklıysa `convert.py` içinde değiştir
- Vade tarihi boş bırakılıyor → açık fatura olarak giriyor
- Fatura sıra no boş → Paraşüt otomatik atıyor
- Gelen faturalarda "Toplam KDV" sütunu Uyumsoft export'unda boş geliyorsa Paraşüt kendi hesaplıyor

---

Developed by [vr0cks](https://github.com/Vr0cks)
