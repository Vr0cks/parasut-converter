# Paraşüt Göç ve Yönetim Asistanı 🚀

Şirketiniz Uyumsoft (veya başka bir ön muhasebe programından) Paraşüt'e geçerken fatura aktarımında ve bakiye eşitlemede ciddi zaman kayıpları yaşıyorsa, bu açık kaynak kodlu araç tam size göre!

Başlangıçta sadece şablon çevirici olarak tasarlanan bu araç, artık **API entegrasyonuna sahip, tam teşekküllü bir Paraşüt Yönetim Asistanı** haline geldi. 

Bu araç sayesinde saatler sürecek işlemleri saniyeler içinde halledebilir, muhasebecinizin hayatını kurtarabilirsiniz. 😎

---

## 🛠️ Özellikler (Menü İçeriği)

Programı çalıştırdığınızda karşınıza çıkan interaktif menü ile şu işlemleri tek tuşla yapabilirsiniz:

1. **Uyumsoft Faturalarını Dönüştür:** Uyumsoft "Gelen" ve "Giden" fatura Excel çıktılarını alır, Paraşüt'ün milimetrik şablon formatlarına (tarihler, sütunlar vb.) anında çevirir.
2. **Tüm Faturaları Toplu Sil (Sıfırlama):** Yanlış bir aktarım mı yaptınız? 90 sayfa faturayı Paraşüt panelinden elle silmek işkencedir. Bu modül API kullanarak tüm satış/alış faturalarınızı tek tuşla temizler.
3. **Mükerrer (Çift) Faturaları Temizle:** İki kez yüklediğiniz faturaları API üzerinden tespit eder (Tutar ve Tarih eşleşmesiyle) ve mükerrer kopyaları otomatik temizler.
4. **Eksik VKN ve Adresleri Güncelle:** Uyumsoft Cari Excel'indeki Vergi Numarası, TC Kimlik No ve Vergi Dairesi bilgilerini okur, Paraşüt'teki müşterilerinizi API ile bularak eksik bilgilerini günceller.
5. **Bakiyeleri Karşılaştır ve Düzeltme Faturası Üret:** Paraşüt geçmiş tahsilatları bilemeyeceği için devir bakiyelerini eksik gösterir. Bu araç, Uyumsoft bakiyeleri ile Paraşüt bakiyelerini kuruşu kuruşuna karşılaştırır ve aradaki farkı kapatacak bir "Geçmiş Dönem Bakiye Düzeltmesi" Excel dosyası üretir.

---

## ⚙️ Kurulum

Python yüklü olduğundan emin olduktan sonra kütüphaneleri kurun:
```bash
pip install pandas openpyxl requests python-dotenv
```

API gerektiren özellikler (2, 3, 4, 5 numaralı işlemler) için bir `.env` dosyası oluşturmanız gereklidir. Proje dizininde bulunan `.env.example` dosyasını kopyalayıp adını `.env` yapın ve içini Paraşüt bilgilerinizle doldurun:

```env
PARASUT_CLIENT_ID=sizin_client_id
PARASUT_CLIENT_SECRET=sizin_secret_id
PARASUT_USERNAME=parasut_mailiniz@sirket.com
PARASUT_PASSWORD=parasut_sifreniz
PARASUT_COMPANY_ID=firma_id_numaraniz
```
*(Şifrenizi ve bilgilerinizi ASLA GitHub vb. yerlere commit etmeyin, `.gitignore` dosyasında `.env` satırının bulunduğundan emin olun.)*

---

## 🚀 Kullanım

Uçbirimi (Terminal) veya Komut Satırını açıp şu komutu çalıştırın:

```bash
python parasut_toolkit.py
```

Karşınıza şöyle şık bir menü çıkacak:
```text
==================================================
    PARAŞÜT GÖÇ VE YÖNETİM ASİSTANI V1.0
==================================================

[ MENÜ ]
1. Uyumsoft Faturalarını Dönüştür (Yakında)
2. Tüm Faturaları Toplu Sil (Sıfırlama)
3. Mükerrer (Çift) Faturaları Temizle
4. Uyumsoft'tan Eksik VKN ve Adresleri Çekip Güncelle
5. Bakiyeleri Karşılaştır ve Düzeltme Faturası Üret
0. Çıkış
Lütfen bir işlem seçin (0-5): _
```
İhtiyacınız olan işlem numarasını tuşlayarak asistanı kullanmaya başlayabilirsiniz!

---

## 📝 Notlar
- Bakiye Düzeltme İşlemi (**5. Menü**) için klasörde Uyumsoft'tan aldığınız `cari uyumsoft.xlsx` dosyası, ayrıca `parasut_satis_faturalari (2).xlsx` ve `parasut_fis_faturalari.xlsx` şablonları bulunmalıdır.
- VKN Güncelleme işlemi (**4. Menü**) de `cari uyumsoft.xlsx` dosyasını baz alır.

---

**Geliştirici:** [vr0cks](https://www.vr0cks.com/tr)
Açık kaynak olarak yayınlanmıştır. Herhangi bir hata veya öneriniz varsa Issue / Pull Request göndermekten çekinmeyin!
