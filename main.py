from openpyxl import load_workbook


book = load_workbook(filename="База данных.xlsx")
sheet1 = book['Движение товаров']
sheet2 = book['Товар']
sheet3 = book['Магазин']

shop_ids = [] # идентификаторы магазинов Заречного района
article = 0 # артикул товара "Яйцо диетическое"
packages = 0 # количество упаковок товара "Яйцо диетическое"

for i in range(1, 18):
    if sheet3['B' + str(i)].value == 'Заречный':
        shop_ids.append(sheet3['A' + str(i)].value)

for i in range(1, 66):
    if sheet2['C' + str(i)].value == 'Яйцо диетическое':
        article = sheet2['A' + str(i)].value
        break

for i in range(1, 2274):
    if sheet1['C' + str(i)].value in shop_ids:
        if sheet1['D' + str(i)].value == article and sheet1['F' + str(i)].value == 'Поступление':
            packages += int(sheet1['E' + str(i)].value)
        elif sheet1['D' + str(i)].value == article and sheet1['F' + str(i)].value == 'Продажа':
            packages -= int(sheet1['E' + str(i)].value)

print(packages) # 966
