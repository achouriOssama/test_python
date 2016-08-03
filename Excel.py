import xlrd
print('***')
wb = xlrd.open_workbook('C:/FichierSource/Reaction Quad-197-215.xlsx')
print('1', wb.sheet_names())
a= wb.sheet_names()
print('a',a)
print('a1',a[0])
sh = wb.sheet_by_name(a[0])
for rownum in range(sh.nrows):
    print(sh.row_values(rownum))
colonne1 = sh.col_values(6)
for iti in colonne1:
    print(iti)
