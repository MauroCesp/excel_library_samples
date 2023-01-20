import pandas as pd

# Vamos a trabajar con varios documentos workbooks
# Crearemos un data frama para cada documento
# Creo 3 variables para cada dataframe
# Cunado se quiere cojer un sheet especifico dentro de un workbook hay que señalar cual.
df_1 = pd.read_excel('docs/stats/1/1_jorunals.xlsx', sheet_name='Sheet1')
df_2 = pd.read_excel('docs/stats/2/2_journals.xlsx', sheet_name='Sheet1')
df_3 = pd.read_excel('docs/stats/3/3_journals.xlsx', sheet_name='Jornals')

# Cuando lo DF tienen las mismas colomnas los podemos CONCATENAR
# Para ello creamos un DF nuevo que guarde la informacion de todos lso demas.
# Asegurarnos que 'sort' sea FALSE para que no cambie el orden de las columnas
df_all = pd.concat([df_1,df_2,df_3], sort = False)

print(df_all)

# Toda la informacion de los 3 ha sido unida en un dataframe
# Pero cada sheet conserva su ubicacion en la hoja
# Es decir que tenemo 3 de cada ROW y COLUMN
# Como ejemplo ubicamos todo lo que tengamos en el con el index 50
print(df_all.loc[50])

# También podemos  analizar data agrupando por columnas
# 'mean'---> significa lo que quiero analizar al agruparlas
# En este caso estoy comparando el              Total_Item_Requests----Unique_Item_Requests

# Es este caso lo analizo por customer porque son diferente OwnedByCustomer# Podriá realizar una grfica comparativa
# Si fuese el mimo cliente lo agrupo por 'Platform'
print(df_all.groupby(['OwnedByCustomer']).mean()['Total_Item_Requests'])

# Ahora vamos a gruardar el reporte combinado a un excel
to_excel = df_all.to_excel('docs/stats/all_docs.xlsx', index=None)
