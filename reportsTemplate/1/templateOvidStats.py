from datetime import datetime, date
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

# Esto es para poder utilizar formulas
from openpyxl.utils import FORMULAE

# Ayuda a los estilo de la celdas
from openpyxl.styles import NamedStyle

from openpyxl.drawing.image import Image


############################
#Importo mis librerías
from models.excel_formattings import Format as misFormulas

##########################
#  NEW COLUMNS MONTH-YEAR
##########################

#Recibo como parametro el documento ya trabajado y el tamaño de la tableStyleInfo
# Paso como parametro el sheet
misFormulas.new_column()

##########################
#  Convert to TABLE
##########################
misFormulas.convert_excel()

misFormulas.set_table()
##########################
#  CREATE GRAPH
##########################

misFormulas.set_graphic()

#misFormulas.set_pivot()
