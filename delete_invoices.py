import os
import sys
import time
import requests
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.getenv('PARASUT_CLIENT_ID')
CLIENT_SECRET = os.getenv('PARASUT_CLIENT_SECRET')
USERNAME = os.getenv('PARASUT_USERNAME')
PASSWORD = os.getenv('PARASUT_PASSWORD')
COMPANY_ID = os.getenv('PARASUT_COMPANY_ID')

if not all([CLIENT_ID, CLIENT_SECRET, USERNAME, PASSWORD, COMPANY_ID]):
    print("HATA: .env dosyasindaki API bilgileri eksik.")
    print("Lutfen .env.example dosyasini kopyalayip .env olarak kaydedin ve icini doldurun.")
    sys.exit(1)

def get_token():
    url = "https://api.parasut.com/oauth/token"
    payload = {
        "username": USERNAME,
        "password": PASSWORD,
        "grant_type": "password",
        "redirect_uri": "urn:ietf:wg:oauth:2.0:oob"
    }
    response = requests.post(url, data=payload, auth=(CLIENT_ID, CLIENT_SECRET))
    if response.status_code == 200:
        return response.json()['access_token']
    else:
        print(f"Token alinamadi: {response.text}")
        sys.exit(1)

def fetch_invoices(token, endpoint):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://api.parasut.com/v4/{COMPANY_ID}/{endpoint}?page[size]=25&page[number]=1"
    invoices = []
    
    while url:
        res = requests.get(url, headers=headers)
        
        if res.status_code == 429:
            print("Sayfa cekilirken limit asildi, 2 sn bekleniyor...")
            time.sleep(2)
            continue
            
        if res.status_code == 200:
            data = res.json()
            invoices.extend(data.get('data', []))
            
            # Pagination
            links = data.get('links', {})
            next_url = links.get('next')
            if next_url:
                url = next_url
            else:
                break
        else:
            print(f"Faturalar cekilemedi: {res.status_code} - {res.text}")
            break
            
    return invoices

def delete_invoices(token, endpoint, invoices):
    headers = {"Authorization": f"Bearer {token}"}
    total = len(invoices)
    print(f"\nToplam {total} adet fatura bulundu. Silme islemi basliyor...\n")
    
    deleted = 0
    failed = 0
    
    for i, inv in enumerate(invoices, 1):
        inv_id = inv['id']
        url = f"https://api.parasut.com/v4/{COMPANY_ID}/{endpoint}/{inv_id}"
        
        while True:
            res = requests.delete(url, headers=headers)
            
            if res.status_code in [204, 200]:
                print(f"[{i}/{total}] ID {inv_id} basariyla silindi.")
                deleted += 1
                break
            elif res.status_code == 429:
                print(f"[{i}/{total}] Limit asildi (429), 2 saniye bekleyip tekrar deneniyor...")
                time.sleep(2)
                # Dongu basa sarip tekrar delete yapacak
            else:
                print(f"[{i}/{total}] ID {inv_id} silinemedi! Hata: {res.text}")
                failed += 1
                break
            
        time.sleep(0.5) # API limitlerine takilmamak icin kisa bekleme
        
    print(f"\nIslem bitti! Basarili: {deleted}, Basarisiz: {failed}")

if __name__ == "__main__":
    print("Parasut Toplu Fatura Silme Araci (API)")
    print("---------------------------------------")
    print("1) Satis Faturalarini (Giden) Sil")
    print("2) Alis Faturalarini (Gelen) Sil")
    print("3) Cikis")
    
    secim = input("\nSeciminiz (1/2/3): ").strip()
    
    if secim == "1":
        endpoint = "sales_invoices"
    elif secim == "2":
        endpoint = "purchase_bills"
    else:
        print("Cikiliyor...")
        sys.exit(0)
        
    onay = input(f"\nDIKKAT: Hesaptaki tum '{endpoint}' verileri SILINECEKTIR. Onayliyor musunuz? (Evet/Hayir): ").strip().lower()
    if onay not in ['evet', 'e']:
        print("Iptal edildi.")
        sys.exit(0)
        
    print("\nAPI Token aliniyor...")
    token = get_token()
    print("Token basariyla alindi. Faturalar listeleniyor (bu islem biraz surebilir)...")
    
    invoices = fetch_invoices(token, endpoint)
    
    if not invoices:
        print("Silinecek fatura bulunamadi.")
        sys.exit(0)
        
    delete_invoices(token, endpoint, invoices)
