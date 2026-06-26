[🇹🇷 Türkçe](#türkçe) | [🇬🇧 English](#english)

---

<h1 id="türkçe">🇹🇷 Paraşüt Göç ve Yönetim Asistanı 🚀</h1>

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

## 📝 Lisans / Ticari Kullanım İhlali

**Ticari Amacla Kullanimi Yasaktir (Non-Commercial Use Only)**

Bu proje tamamen **ücretsiz** ve açık kaynak olarak, muhasebecilerin, stajyerlerin ve entegrasyon süreçlerinde boğulan personellerin hayatını kolaylaştırmak amacıyla paylaşılmıştır. Kodları istediğiniz gibi alıp inceleyebilir, kendi şirket içi projelerinizde ücretsiz kullanabilirsiniz. 

Ancak bu aracın veya kodlarının **başka kişilere, kurumlara veya şirketlere satılması, ticari bir ürün veya SaaS içerisine entegre edilerek üzerinden para kazanılması KESİNLİKLE YASAKTIR.** Yoksa ben de bilirdim alıp millete parayla satmayı. Lütfen emeğe saygı gösterin ve bu aracı tamamen ücretsiz bir araç olarak yaşamaya bırakın.

**Geliştirici:** [vr0cks](https://yigit.vr0cks.com)

---
---

<h1 id="english">🇬🇧 Paraşüt Migration & Management Assistant 🚀</h1>

If your company is experiencing significant time losses in invoice transfers and balance synchronizations while migrating from Uyumsoft (or another pre-accounting software) to Paraşüt, this open-source tool is exactly what you need!

Initially designed simply as a template converter, this tool has now evolved into a **fully-fledged Paraşüt Management Assistant with complete API integration.**

With this tool, you can handle operations that would normally take hours in mere seconds, effectively saving your accountant's life. 😎

---

## 🛠️ Features (Menu Contents)

When you run the program, you can perform the following operations with a single click via the interactive menu:

1. **Convert Uyumsoft Invoices:** Takes Uyumsoft "Incoming" and "Outgoing" invoice Excel outputs and instantly converts them to Paraşüt's millimeter-perfect template formats (dates, columns, etc.).
2. **Bulk Delete All Invoices (Reset):** Made a wrong transfer? Manually deleting 90 pages of invoices from the Paraşüt panel is torture. This module uses the API to clean all your sales/purchase invoices with a single click.
3. **Clean Duplicate Invoices:** Detects invoices you've uploaded twice via the API (by matching Amount and Date) and automatically cleans the duplicate copies.
4. **Update Missing Tax IDs and Addresses:** Reads Tax Number, ID Number, and Tax Office information from the Uyumsoft Customer Excel, finds your customers in Paraşüt via API, and updates their missing information.
5. **Compare Balances and Generate Correction Invoice:** Since Paraşüt cannot know past collections, it shows transfer balances incompletely. This tool compares Uyumsoft balances with Paraşüt balances down to the penny and generates a "Past Period Balance Correction" Excel file to close the difference.

---

## ⚙️ Installation

After ensuring Python is installed, install the required libraries:
```bash
pip install pandas openpyxl requests python-dotenv
```

For features requiring the API (operations 2, 3, 4, 5), you need to create a `.env` file. Copy the `.env.example` file in the project directory, rename it to `.env`, and fill it with your Paraşüt credentials:

```env
PARASUT_CLIENT_ID=your_client_id
PARASUT_CLIENT_SECRET=your_secret_id
PARASUT_USERNAME=your_parasut_email@company.com
PARASUT_PASSWORD=your_parasut_password
PARASUT_COMPANY_ID=your_company_id
```
*(NEVER commit your password and information to GitHub etc., make sure the `.env` line is in your `.gitignore` file.)*

---

## 🚀 Usage

Open the Terminal or Command Prompt and run the following command:

```bash
python parasut_toolkit.py
```

You will see a stylish menu like this:
```text
==================================================
    PARAŞÜT MIGRATION AND MANAGEMENT ASST. V1.0
==================================================

[ MENU ]
1. Convert Uyumsoft Invoices (Coming Soon)
2. Bulk Delete All Invoices (Reset)
3. Clean Duplicate Invoices
4. Pull & Update Missing Tax IDs and Addresses from Uyumsoft
5. Compare Balances and Generate Correction Invoice
0. Exit
Please select an operation (0-5): _
```
You can start using the assistant by typing the number of the operation you need!

---

## 📝 License / Commercial Use Restriction

**Non-Commercial Use Only**

This project is shared completely **free** and open-source to make life easier for accountants, interns, and personnel drowning in integration processes. You can take the codes, examine them, and use them for free in your internal company projects.

However, **selling this tool or its codes to other individuals, institutions, or companies, or integrating it into a commercial product or SaaS to make money is STRICTLY PROHIBITED.** Otherwise, I would have known how to take it and sell it to people for money. Please respect the effort and let this tool live as a completely free resource.

**Developer:** [vr0cks](https://yigit.vr0cks.com)
