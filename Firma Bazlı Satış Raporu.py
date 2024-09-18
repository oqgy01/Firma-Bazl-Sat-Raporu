#Doğrulama Kodu
import requests
from bs4 import BeautifulSoup
url = "https://docs.google.com/spreadsheets/d/1AP9EFAOthh5gsHjBCDHoUMhpef4MSxYg6wBN0ndTcnA/edit#gid=0"
response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, "html.parser")
first_cell = soup.find("td", {"class": "s2"}).text.strip()
if first_cell != "Aktif":
    exit()
first_cell = soup.find("td", {"class": "s1"}).text.strip()
print(first_cell)


import pandas as pd
import requests
from io import BytesIO
import re

# İndirilecek linkler
links = [
    "https://task.haydigiy.com/FaprikaXls/ZIMVGV/1/",
    "https://task.haydigiy.com/FaprikaXls/ZIMVGV/2/",
    "https://task.haydigiy.com/FaprikaXls/ZIMVGV/3/"
]

# Kullanıcıdan Firma Kodu alınması
firma_kodu = input("Firma Kodu (Ör: .1234.): ")


print(" ")
print("Oturum Açma Başarılı Oldu")
print(" /﹋\ ")
print("(҂`_´)")
print("<,︻╦╤─ ҉ - -")
print("/﹋\\")
print("Mustafa ARI")
print(" ")



# Excel dosyalarını indirip birleştirme
dfs = []
for link in links:
    response = requests.get(link)
    if response.status_code == 200:
        # BytesIO kullanarak indirilen veriyi DataFrame'e dönüştürme
        df = pd.read_excel(BytesIO(response.content))
        
        # Firma Kodu'nu içeren satırları seçme
        selected_rows = df[df['UrunAdi'].astype(str).str.contains(firma_kodu.replace(".", "\."), case=False, na=False)]
        dfs.append(selected_rows)
    else:
        print(f"Hata: {response.status_code} - {link}")

# Firma Kodu'nu içeren satırları birleştirme
merged_df = pd.concat(dfs, ignore_index=True)




# Belirli başlıklar dışındaki sütunları silme
selected_columns = ["UrunAdi", "StokAdedi", "AlisFiyati", "SatisFiyati", "Resim", "AramaTerimleri", "MorhipoKodu", "VaryasyonMorhipoKodu", "HepsiBuradaKodu"]
filtered_df = merged_df[selected_columns]


# Sonuç DataFrame'i tek bir Excel dosyasına yazma
filtered_df.to_excel("sonuc_excel.xlsx", index=False)





# "sonuc_excel.xlsx" Excel dosyasını oku
df_calisma_alani = pd.read_excel('sonuc_excel.xlsx')

# Aynı "UrunAdi" hücrelerinin "StokAdedi" sayılarını toplama
df_calisma_alani.loc[:, "StokAdedi"] = df_calisma_alani.groupby("UrunAdi")["StokAdedi"].transform("sum")

# "MorhipoKodu" sütununun adını değiştirme
df_calisma_alani = df_calisma_alani.rename(columns={"VaryasyonMorhipoKodu": "N11 & Zimmet"})

# Veri tiplerini uyumlu hale getirme
df_calisma_alani["StokAdedi"] = pd.to_numeric(df_calisma_alani["StokAdedi"], errors="coerce")
df_calisma_alani["N11 & Zimmet"] = pd.to_numeric(df_calisma_alani["N11 & Zimmet"], errors="coerce")

df_calisma_alani['N11 & Zimmet'].fillna(0, inplace=True)

# "Toplam Stok Adedi" sütununu oluşturma ve "StokAdedi" ile "N11 & Zimmet" sütunlarındaki verileri toplama
df_calisma_alani["Toplam Stok Adedi"] = df_calisma_alani["StokAdedi"] + df_calisma_alani["N11 & Zimmet"]

# "VaryasyonMorhipoKodu" sütununun adını değiştirme
df_calisma_alani = df_calisma_alani.rename(columns={"MorhipoKodu": "Günlük Satış Adedi"})
df_calisma_alani['Günlük Satış Adedi'].fillna(0, inplace=True)

# "Kaç Güne Biter" sütununu oluşturma ve "Toplam Stok Adedi" sütunundaki verileri "Günlük Satış Adedi" sütunundaki verilere bölme işlemi
df_calisma_alani["Kaç Güne Biter"] = "Satış Adedi Yok"  # Varsayılan değer olarak "Satış Adedi Yok" atanır

non_zero_mask = df_calisma_alani["Günlük Satış Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Kaç Güne Biter"] = round(df_calisma_alani["Toplam Stok Adedi"] / df_calisma_alani["Günlük Satış Adedi"])

# "Resim" sütunundaki ".jpeg" ifadesinden sonrasını temizleme
df_calisma_alani["Resim"] = df_calisma_alani["Resim"].str.replace(r"\.jpeg.*$", "", regex=True)

# Kalan verilere ".jpeg" eklenmesi
df_calisma_alani["Resim"] = df_calisma_alani["Resim"] + ".jpeg"

# Sütun sıralamasını ayarlama
column_order = ["UrunAdi", "StokAdedi", "N11 & Zimmet", "Toplam Stok Adedi", "Günlük Satış Adedi", "Kaç Güne Biter", "AlisFiyati", "SatisFiyati", "AramaTerimleri", "Resim"]
df_calisma_alani = df_calisma_alani[column_order]

# Tekrarlanan satırları silme
df_calisma_alani = df_calisma_alani.drop_duplicates(subset=["UrunAdi"])

# Sonuç DataFrame'ini tekrar "sonuc_excel.xlsx" adlı bir Excel dosyasına yazma
df_calisma_alani.to_excel("sonuc_excel.xlsx", index=False)






# "CalismaAlani" Excel dosyasını oku
df_calisma_alani = pd.read_excel('sonuc_excel.xlsx')

# Tarihleri çıkarmak için regex deseni
date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})'

# "AramaTerimleri" sütunundaki tarihleri temizle
df_calisma_alani['AramaTerimleri'] = df_calisma_alani['AramaTerimleri'].apply(lambda x: re.search(date_pattern, str(x)).group(1) if re.search(date_pattern, str(x)) else None)



# Güncellenmiş DataFrame'i aynı Excel dosyasının üzerine yaz
with pd.ExcelWriter('sonuc_excel.xlsx', engine='xlsxwriter') as writer:
    df_calisma_alani.to_excel(writer, index=False, sheet_name='Sheet1')




    





# Excel dosyasını güncelleyerek genişlikleri ve ortalamayı ayarla
with pd.ExcelWriter('sonuc_excel.xlsx', engine='xlsxwriter') as writer:
    df_calisma_alani.to_excel(writer, index=False, sheet_name='Sheet1')

    # ExcelWriter objesinden workbook ve worksheet'e eriş
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # DataFrame sütun genişliklerini al
    column_widths = [max(df_calisma_alani[col].astype(str).apply(len).max(), len(col)) + 2 for col in df_calisma_alani.columns]

    # Sütun genişliklerini Excel worksheet'e ayarla
    for i, width in enumerate(column_widths):
        worksheet.set_column(i, i, width)

    # Tabloyu ortala
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    for i, col in enumerate(df_calisma_alani.columns):
        worksheet.write(0, i, col, center_format)
        
    # Sütun başlıklarının rengini gri yap
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'align': 'center', 'valign': 'vcenter'})
    for col_num, value in enumerate(df_calisma_alani.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Verileri tabloya yazarken ortala
    for i, col in enumerate(df_calisma_alani.columns):
        for j, value in enumerate(df_calisma_alani[col]):
            # "Resim" sütunundaki linkleri "Tıkla" adlı bağlantıya çevir
            if col == 'Resim':
                worksheet.write_url(j + 1, i, value, string='Tıkla', cell_format=center_format, tip='url')
            else:
                worksheet.write(j + 1, i, value, center_format)

    # "Resim" sütununun genişliğini 20 piksel olarak ayarla
    worksheet.set_column('J:J', 20)



import os

# Dosyanın adını değiştirme
excel_file_name = "sonuc_excel.xlsx"
new_excel_file_name = f"{firma_kodu} Satış Raporu.xlsx"
os.rename(excel_file_name, new_excel_file_name)


