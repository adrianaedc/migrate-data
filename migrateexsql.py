import xlrd
#import mysqlclient

book = xlrd.open_workbook("/home/adrianaedc/Documentos/Compra y Ventas Enero 2011.xls")
print("Numero de hojas de trabajo: {0}".format(book.nsheets))
print("Nombre de hoja de trabajo: {0}".format(book.sheet_names()))
sh = book.sheet_by_name('Ventas Enero 2011')
print("Hoja: {0}, Filas: {1}, Columnas: {2}".format(sh.name, sh.nrows, sh.ncols))
print("Celda H30 es {0}".format(sh.cell_value(rowx=29, colx=7)))
for r in range(15, sh.nrows-16):
    print(sh.row(r))