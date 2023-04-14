import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook

# Lê o arquivo xlsx
df = pd.read_excel('teste.xlsx')

# Agrupa os dados por tipo de categoria e conta o número de tickets em cada categoria
counts = df.groupby('Categoria').size()

# Gera o gráfico de barras
plt.bar(counts.index, counts.values)
plt.title('Quantidade de Tickets por Categoria')
plt.xlabel('Categoria')
plt.ylabel('Quantidade')

# Carrega o teste xlsx e insere o gráfico na planilha 'Gráfico'
book = load_workbook('teste.xlsx')
writer = pd.ExcelWriter('teste.xlsx', engine='openpyxl') 
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

# Insere o gráfico na planilha 'Gráfico'
fig = plt.gcf()
ws = book['Gráfico']
img = plt.imshow(fig)
img.set_extent([0, 10, 0, 10])
ws.add_image(img, 'A1')

# Salva as alterações no arquivo xlsx
writer.save()