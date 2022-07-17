import openpyxl

from openpyxl.chart import PieChart, Reference, Series,PieChart3D
wb = openpyxl.Workbook()
ws = wb.active

#creamos nustro data set
data = [
        ['Flavor','Sold'],
        ['Vanila',1500],
        ['Chocolate', 1700],
        ['Strawberry',600],
        ['Pumkin Spice',950]
]

# Ahora interactuamos la la data y la adjuntamos al archivo de excel
for rows in data:
    ws.append(rows)

# Ahora que tenemos informacion tenemos que decidir como se presenta en la grafica
chart = PieChart()

# Especificamos desde donde, hasta donde queremos la data
labels = Reference(ws,min_col=1, min_row=2, max_row=5)
data = Reference(ws,min_col=1, min_row=2, max_row=5)

chart.add_data(data,titles_from_data=True)
chart.set_categories(labels)
chart.title = 'Ice crema Flavor'

# Como las columnas terminan en B ponemos el gráfico en C

ws.add_chart(chart,'C1')

# Cunado tenemos todo lo que queremos guardamos el archivo

wb.save('graphs.xlsx')
