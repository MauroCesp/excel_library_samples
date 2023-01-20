# Trabajaremos con Pillow para poder manipular imagenes

# Primero importo algunas funciones para tablas e imagenes
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.drawing.image import Image

from openpyxl import load_workbook

#---------------------#
#       WORK BOOK     #
#---------------------#
wb = load_workbook('docs/graphs.xlsx')

#----------------------#
#       WORK SHEET     #
#----------------------#
# grab the active worksheet
# This will create the active sheet on this work BOOK
ws = wb.active
#Provide le sheet with a name
ws.title = "Titulo que queramos"

# Ahora necesitamos ver como están formateados nuestros datos
# El el caso de este Excel van de la A1 hasta B5
# Hay que tener eso siempre en cuenta

# Creamos un objeto de data que guarde los datos de la tabla
# el parametro 'ref' indica el rango de los datos
tab = Table(displayName='Table1',ref = 'A1:B5')

# Ahora le damos estilo a la tabla
# Creo una variable que guarde el objeto de estilo
style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False, showLastColumn= False,
                                                    showRowStripes=True, showColumnStripes=True)

# ahora asigno la variable
tab.tableStyleInfo = style

# Por ultimo agregamo la tabla al worksheet
ws.add_table(tab)
wb.save('docs/table.xlsx')

# imagenes
img = Image('img/excel.jpg')

# El objeto de Image tiene propiedad de height and width
# pero no logré verlo
# Buscar en internet


# Cunado tenemos la imagen la agregamos al worksheet
######################
#    IMPORTANTE
######################

# Cunado metemos algo en el worksheet el primer valor (access cell) va a
# ser siempre el TOP-LEFT_CORNER
ws.add_image(img,'C1')

wb.save('docs/image.xlsx')
