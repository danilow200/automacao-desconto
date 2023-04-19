import openpyxl
from openpyxl.chart import BarChart, Reference, Series

pagina = 'Comparação Total'
 
workbook = openpyxl.load_workbook('teste.xlsx')

worksheet = workbook[pagina]

chart = BarChart()
chart.type = 'col'
chart.title = 'Porcetagem PADTEC'
chart.x_axis.title = 'Categorias'
chart.y_axis.title = 'Porcetagem'

categorias = Reference(worksheet, min_col=2, min_row=4, max_row=11)
valores = Reference(worksheet, min_col=7, min_row=4, max_row=11)
serie = Series(values=valores, xvalues=categorias)
chart.add_data(valores, titles_from_data=True)
chart.set_categories(categorias)

worksheet.add_chart(chart, 'B14')

workbook.save('teste.xlsx')