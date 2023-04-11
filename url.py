import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

input_file = 'final_data.xlsx'  # Mevcut Excel dosyanızın adı
output_file = 'output_hyperlink.xlsx'  # Yeni oluşturulacak Excel dosyasının adı

# Excel dosyasını oku
df = pd.read_excel(input_file)

# Yeni bir Excel çalışma kitabı oluştur
wb = Workbook()
ws = wb.active

# Dataframe'i satırlara dönüştür ve Excel çalışma sayfasına yaz
for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
    for c_idx, value in enumerate(row):
        cell = ws.cell(row=r_idx + 1, column=c_idx + 1, value=value)
        # İlk satırda (başlıkta) değilsek ve ikinci sütunda (URL sütunu) ise
        if r_idx > 0 and c_idx == 1:
            # HYPERLINK formülünü kullanarak bağlantıları tıklanabilir hale getir
            cell.value = f'=HYPERLINK("{value}", "{value}")'
            cell.font = Font(color="0000FF", underline="single")

# Excel dosyasını kaydet
wb.save(output_file)
print(f"{output_file} dosyası başarıyla oluşturuldu.")
