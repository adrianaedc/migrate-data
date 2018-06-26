import datetime, xlrd
#import mysqlclient

book = xlrd.open_workbook("/home/adrianaedc/Documentos/Compra y Ventas Enero 2011.xls")
print("Numero de hojas de trabajo: {0}".format(book.nsheets))
print("Nombre de hoja de trabajo: {0}".format(book.sheet_names()))
sh = book.sheet_by_name('Ventas Enero 2011')
print("\nHoja de trabajo actual: {0}, \nNumero de Filas: {1},\tNumero de Columnas: {2}\n".format(sh.name, sh.nrows, sh.ncols))
print("Celda N55 es {0}\n".format(sh.cell_value(rowx=54, colx=13)))
for r in range(15, sh.nrows-16):
    if sh.cell(r,1).value != "":
        #Fecha
        fdate = datetime.datetime(*xlrd.xldate_as_tuple(sh.cell(r,1).value, book.datemode))
        #Maquina fiscal
        if sh.cell(r,6).value == "":
            maq = "Sin maquina"
        else:
            maq = sh.cell(r,6).value[3:13]
        #Numero de reporte z
        if sh.cell(r,2).value == "":
            repz = 0
        else:
            repz = int(sh.cell(r,2).value)
        #Numero Inicial
        if sh.cell(r,4).value == "":
            inicial = 0
        else:
            inicial = int(sh.cell(r,4).value)
        #Numero Final
        if sh.cell(r,5).value == "":
            final = 0
        else:
            final = int(sh.cell(r,5).value)
        #Base No Contribuyente
        if sh.cell(r,22).value == "":
            base_nc = 0
        else:
            base_nc = float("{0:.2f}".format(sh.cell(r,22).value,2))
        #Alicuota No Contribuyente
        if sh.cell(r,23).value == "":
            alicuota_nc = 0
        else:
            alicuota_nc = int(sh.cell(r,23).value)        
        #IVA No Contribuyente
        if sh.cell(r,24).value == "":
            iva_nc = 0
        else:
            iva_nc = float("{0:.2f}".format(sh.cell(r,24).value,2))
        #Base Contribuyente
        if sh.cell(r,19).value == "":
            base_c = 0
        else:
            base_c = float("{0:.2f}".format(sh.cell(r,19).value,2))
        #Alicuota Contribuyente
        if sh.cell(r,20).value == "":
            alicuota_c = 0
        else:
            alicuota_c = int(sh.cell(r,20).value)
        #IVA Contribuyente
        if sh.cell(r,21).value == "":
            iva_c = 0
        else:
            iva_c = float("{0:.2f}".format(sh.cell(r,21).value,2))
        #Total 
        if sh.cell(r,16).value == "":
            total = 0
        else:
            total = float("{0:.2f}".format(sh.cell(r,16).value,2))