import openpyxl
from openpyxl.chart import BarChart, Reference

# abrir o arquivo xlsx
wb = openpyxl.load_workbook('teste.xlsx')

# selecionar a planilha desejada
ws = wb['Codigos Geral']

# contar as repetições de cada string na coluna 8
count_dict = {}
for cell in ws['H']:
    if cell.value in count_dict:
        count_dict[cell.value] += 1
    else:
        count_dict[cell.value] = 1

# criar a lista de dados para o gráfico
data = []
for key, value in count_dict.items():
    data.append([key, value])

# definir a faixa de dados para o gráfico
chart_data = Reference(ws, min_col=8, min_row=4, max_col=8, max_row=len(data))

# criar o gráfico de barras
chart = BarChart()
chart.add_data(chart_data)

# adicionar o gráfico à planilha
ws.add_chart(chart, "E2")

# salvar o arquivo com o gráfico
wb.save("arquivo_com_grafico.xlsx")