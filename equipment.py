import xlrd

Note = open("D:/桌面/PCR AutoScript/equipment.txt", mode='w')
workbook = xlrd.open_workbook("D:/桌面/PCR AutoScript/xls/15715631699.xlsx")
worksheet = workbook.sheet_by_name("装备")
rows = worksheet.nrows
print(rows)
for i in range(rows):
    Note.write(str(worksheet.cell_value(i, 0))[:-2])
    Note.write(' ')
    Note.write(str(worksheet.cell_value(i, 2))[:-2])
    Note.write('\n')
