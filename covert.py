import pandas as pd
import openpyxl
from openpyxl import Workbook

# En este caso importamos esta librria para que nos ayude apasasr la info de Pandas a rows en excel
from openpyxl.utils.dataframe import dataframe_to_rows

#---------------------#
#       WORK BOOK     #
#---------------------#
wb = Workbook()

ws = wb.active

# Ahora paso la info de Excel a pandas
# Pandas funciona para hacer todo el analisisi de la Data y despues presentarlo en Excel.
# abro la informacion de los reportes combinados y analizo esa data con Pandas
df = pd.read_excel('docs/stats/all_docs.xlsx')

# Quiero analizar solo algunas columnas entonces creo otr dataframa para guardar la info.
# Busco las columnas que necesito
df_1 = df[['Total_Item_Requests','Unique_Item_Requests','Platform','UsedByCustomer']]

############################
# Pasarlo de PANDAS a EXCEL
############################

# Cunado terminamos de analizar la data y tla tenemos organizada
# La pasamos de nuevo a EXCEL# Para ello necesitamos de la libreria  dataframe_to_rows
# El primer parametro es el DF que queremos utilizar
# El INDEX lo ponemos en FLASE proque no queremos pegar indices cuando Excel ya los tiene
# Row es una variable que guarda varia informacion y podemos interactuar con ella para analizar lo que guarda
rows = dataframe_to_rows(df_1,index=False)

############################
# NESTED FOR LOOP
############################

#for row in rows:
    # Podemos quer que toda se imprime en forma de rows tipo lista
    # Ahora necesito ingresar a cada value para asignarlo a una celda en EXCEL# Para ello utilizo un NESTED FOR LOOP
    #for col in row:
        #print(col)
# Ahora que tenemos la data de una forma que pueda ser accesidad por openpyxl# Necesitamos ver la manera de acceder a la dataframe
#------------------------------------------------------------------------------
# Openpyxl ofrece la posibilida de buscar por Index de ROW COL
# Guarda los index en las variables '# r _idx, c_idx'

for r_idx, row in enumerate(rows,1):
    # Podemos quer que toda se imprime en forma de rows tipo lista
    # Ahora necesito ingresar a cada value para asignarlo a una celda en EXCEL# Para ello utilizo un NESTED FOR LOOP
    # El '1' solo significa la columna donde queremos comenzar
    # Podemos comenzar donde queramos
    for c_idx, col in enumerate(row,1):
        # Con los indices que creamos vamos a usarlos para ubicar cada celdas
        ws.cell(row=r_idx, column=c_idx, value= col)
# Ahora que tenemos la data de una forma que pueda ser accesidad por openpyxl# Necesitamos ver la manera de acceder a la dataframe

# Save the file
wb.save('docs/stats/regions.xlsx')
