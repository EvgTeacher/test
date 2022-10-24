import openpyxl

book = openpyxl.Workbook()
sheet = book.active

for i in range(1, 5):
    for j in range(1, 8):
        sheet.cell(row=i, column=j).value = f'Строка {i}, Столбец {j}'


# book.close()
# i, j = 0, 0


book.save('t.xlsx')



# book.close()

for i in range(1, 5):
    for j in range(1, 8):
        print(sheet[i][j-1].value)