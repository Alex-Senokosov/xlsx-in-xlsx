import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook

# Чтение данных из файла ostatki.xlsx
df = pd.read_excel('ostatki.xlsx')

# Группировка данных и расчет средней прибыли
grouped = df.groupby(['BRAND', 'Diametr', 'Color', 'SUPPLIER']).agg({
    'QUANTITY': 'sum',
    'PURCHASE_PRICE': 'mean',
    'PRICE': 'mean'
}).reset_index()
grouped['Средняя прибыль'] = grouped['PRICE'] - grouped['PURCHASE_PRICE']

# Создание нового файла Excel
wb = Workbook()
ws = wb.active

# Запись заголовков
headers = ['Бренд', 'Диаметр', 'Цвет', 'Количество', 'Средняя цена закупки', 'Средняя цена продажи', 'Средняя прибыль', 'Поставщик']
ws.append(headers)

# Запись данных
for _, row in df.iterrows():
    data = [row['BRAND'], row['Diametr'], row['Color'], row['QUANTITY'], row['PURCHASE_PRICE'], row['PRICE'], row['PRICE'] - row['PURCHASE_PRICE'], row['SUPPLIER']]
    ws.append(data)

# Сохранение файла результат.xlsx
wb.save('результат.xlsx')