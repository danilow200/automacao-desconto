#------------------------------------------------------------------------------------------------------------------------
                                                #IMPORTAÇÕES
import pandas as pd
import openpyxl
from EstilizandoExcel import estilizar_excel
from gerarGrafico import gerar_grafico
from datetime import datetime
import shutil
#------------------------------------------------------------------------------------------------------------------------
                                                #FUNÇÕES PARA FAZER OS CALCULOS DAS DATAS
def strfdelta(tdelta, fmt):
    d = {"days": tdelta.days}
    d["hours"], rem = divmod(tdelta.seconds, 3600)
    d["hours"] += 24 * tdelta.days # adicionar horas correspondentes aos dias
    d["minutes"], d["seconds"] = divmod(rem, 60)
    return fmt.format(**d)

def soma_tempos(tempo1, tempo2):
    h1, m1, s1 = map(int, tempo1.split(':'))
    h2, m2, s2 = map(int, tempo2.split(':'))
    total = (h1+h2)*3600 + (m1+m2)*60 + s1+s2
    horas, resto = divmod(total, 3600)
    minutos, segundos = divmod(resto, 60)
    return '{:02d}:{:02d}:{:02d}'.format(horas, minutos, segundos)

def subtract_times(time1, time2):
    h1, m1, s1 = map(int, time1.split(':'))
    h2, m2, s2 = map(int, time2.split(':'))
    total_seconds1 = h1 * 3600 + m1 * 60 + s1
    total_seconds2 = h2 * 3600 + m2 * 60 + s2
    if total_seconds1 >= total_seconds2:
        sign = ''
        diff_seconds = total_seconds1 - total_seconds2
    else:
        diff_seconds = total_seconds2 - total_seconds1
        sign = '-'
    diff_hours = diff_seconds // 3600
    diff_minutes = (diff_seconds % 3600) // 60
    diff_seconds = diff_seconds % 60
    return '{}{:02d}:{:02d}:{:02d}'.format(sign, diff_hours, diff_minutes, diff_seconds)

def calculate_time_percentage(time1, time2):
    if time1 == '00:00:00':
        return 0.0
    if time2 == '00:00:00':
        return 0.0
    h1, m1, s1 = map(int, time1.split(':'))
    h2, m2, s2 = map(int, time2.split(':'))
    total_seconds1 = h1 * 3600 + m1 * 60 + s1
    total_seconds2 = h2 * 3600 + m2 * 60 + s2
    percentage = total_seconds2 / total_seconds1
    return percentage
#------------------------------------------------------------------------------------------------------------------------
                                            #DICIONARIOS
categoria_dicionario = {
                    'IPFA': 'Acesso', 
                    'IPAC': 'Aguardando CIGR',
                    'IPAR': 'Área de Risco',
                    'IPAA':'Atividade Agendada',
                    'IPFR': 'Falha Restabelecida',
                    'IPFE': 'Falta de Energia',
                    'IPOS': 'Outros',
                    'IPTS': 'Terceiros'
                }

float_format = {'Porcetagem PADTEC': '%.2f%%', 'Porcetagem RADIANTE': '%.2f%%'}
#------------------------------------------------------------------------------------------------------------------------
                                            #ARRAYS E VARIÁVEIS
desconto_dado = []
desconto_auto = []
desconto_total = []
diferenca_desconto = []
desconto_auto_final = []
empresa = []
tickets_corretos = []
categoria_corretos = []

ticket_anterior = 0

categorias = ["Falha Restabelecida", "Acesso", "Aguardando CIGR", "Terceiros", "Área de Risco", "Falta de Energia", "Outros", "Atividade Agendada"]
categoria_soma_dada_pad = {categoria: '00:00:00' for categoria in categorias}
categoria_soma_auto_pad = {categoria: '00:00:00' for categoria in categorias}
categoria_soma_dada_rad = {categoria: '00:00:00' for categoria in categorias}
categoria_soma_auto_rad = {categoria: '00:00:00' for categoria in categorias}
categoria_porcetagem_pad = {categoria: '00:00:00' for categoria in categorias}
categoria_porcetagem_rad = {categoria: '00:00:00' for categoria in categorias}

arquivo_desconto_auto = "descontos abril.xlsx"
arquivo_desconto_dado = "Indicadores - Abril.xlsx"
#------------------------------------------------------------------------------------------------------------------------
                                            #LENDO ARQUIVOS EXCEL
desconto_auto_planilha = pd.read_excel(arquivo_desconto_auto,sheet_name='Codigos com desconto automatico')
desconto_dado_planilha = pd.read_excel(arquivo_desconto_dado,sheet_name='Descontos')
desconto_dado_planilha = desconto_dado_planilha.drop(index=range(0,21),axis=0)

tickets_data_planilha = pd.read_excel(arquivo_desconto_auto,sheet_name='Tickets dentro da data')
tickets_dentro_data = tickets_data_planilha['Tickets dentro da data']
tickets_dentro_data_str = list(map(str, tickets_dentro_data))
#------------------------------------------------------------------------------------------------------------------------
                                            #COMPARAÇÃO DE DESCONTOS PEDIDOS COM DESCONTOS DADOS
valor_ja_utilizado_auto = [True] * len(desconto_auto_planilha['Tickets'])

cont2 = 0

for index_dado, row_dado in desconto_dado_planilha.iterrows():
    if str(row_dado['Unnamed: 0']) in tickets_dentro_data_str:
        if str(row_dado['Unnamed: 4'])[4:5] == '-':
            data1 = datetime.strptime(str(row_dado['Unnamed: 4']), '%Y-%m-%d %H:%M:%S')
        else:
            data1 = datetime.strptime(str(row_dado['Unnamed: 4']), '%d/%m/%Y %H:%M:%S') #datetime faz tratamento para que as strings que vem em formato de dado para um novo formato capaz de fazer seus calculos
        data2 = datetime.strptime(str(row_dado['Unnamed: 5']), '%d/%m/%Y %H:%M:%S')
        diferenca = data2 - data1
        diferenca_str = strfdelta(diferenca, "{hours:02d}:{minutes:02d}:{seconds:02d}")
        desconto_dado.append(diferenca_str)
        
        tickets_corretos.append(row_dado['Unnamed: 0'])
        categoria_corretos.append(row_dado['Unnamed: 3'])
        
        if 'padtec' in row_dado['Unnamed: 7']:
            empresa.append(str('PADTEC'))
        else:
            empresa.append(str('RADIANTE'))

        ticket_anterior = row_dado['Unnamed: 0']
        valida = True
        
        if row_dado['Unnamed: 3'] in categorias:
            if empresa[cont2] == 'PADTEC':
                categoria_soma_dada_pad[row_dado['Unnamed: 3']] = soma_tempos(categoria_soma_dada_pad[row_dado['Unnamed: 3']], desconto_dado[cont2])
            else:
                categoria_soma_dada_rad[row_dado['Unnamed: 3']] = soma_tempos(categoria_soma_dada_rad[row_dado['Unnamed: 3']], desconto_dado[cont2])    
        
        for index_auto, row_auto in desconto_auto_planilha.iterrows():
            if row_dado['Unnamed: 0'] == row_auto['Tickets'] and valor_ja_utilizado_auto[index_auto]:
                categoria_auto = categoria_dicionario[row_auto['Abertura'][6:10]]
                if row_dado['Unnamed: 3'] == categoria_auto:
                    valida = False
                    valor_ja_utilizado_auto[index_auto] = False
                    desconto_auto_final.append(row_auto['Desconto'])
                    diferenca_desconto.append(subtract_times(desconto_dado[cont2], row_auto['Desconto']))
                    break
        if valida:
            desconto_auto_final.append(str(''))
            diferenca_desconto.append(subtract_times(desconto_dado[cont2], str('00:00:00')))
        cont2 += 1
#------------------------------------------------------------------------------------------------------------------------
                                            #SOMANDO AS HORAS TOTAIS
                                            
for index_auto, row_auto in desconto_auto_planilha.iterrows():
    if row_auto['Abertura'][6:10] in categoria_dicionario:
        categoria = categoria_dicionario[row_auto['Abertura'][6:10]]
        if row_auto['Empresa'] == 'PADTEC':
            categoria_soma_auto_pad[categoria] = soma_tempos(categoria_soma_auto_pad[categoria], row_auto['Desconto'])
        else:
            categoria_soma_auto_rad[categoria] = soma_tempos(categoria_soma_auto_rad[categoria], row_auto['Desconto'])
            
#------------------------------------------------------------------------------------------------------------------------
                                            #SALVANDO EM PORCETAGEM
            
for index in categorias:
    categoria_porcetagem_pad[index] = calculate_time_percentage(categoria_soma_dada_pad[index], categoria_soma_auto_pad[index] )
    categoria_porcetagem_rad[index] = calculate_time_percentage(categoria_soma_dada_rad[index], categoria_soma_auto_rad[index] )
#------------------------------------------------------------------------------------------------------------------------
                                            # DATA FRAMES
data = {
    'Ticket': tickets_corretos,
    'Categoria': categoria_corretos,
    'Empresa': empresa,
    'Desconto Dado': desconto_dado,
    'Desconto Automatico': desconto_auto_final,
    'Diferença': diferenca_desconto
}

dataf = {
    'Desconto Dado Total PADTEC': categoria_soma_dada_pad,
    'Desconto Automatico Total PADTEC': categoria_soma_auto_pad,
    'Desconto Dado Total RADIANTE': categoria_soma_dada_rad,
    'Desconto Automatico Total RADIANTE': categoria_soma_auto_rad,
    'Porcetagem PADTEC': categoria_porcetagem_pad,
    'Porcetagem RADIANTE': categoria_porcetagem_rad
}

df2 = pd.DataFrame(dataf)
df = pd.DataFrame(data)
#------------------------------------------------------------------------------------------------------------------------
                                    #SALVANDO NO EXCEL E ESTILIZANDO

novo_arquivo = input("Digite o nome para a planilha\n")
novo_arquivo += ".xlsx"

shutil.copy(arquivo_desconto_auto, novo_arquivo)

book = openpyxl.load_workbook(novo_arquivo)

with pd.ExcelWriter(novo_arquivo, engine='openpyxl', mode='a') as writer:
    df.to_excel(writer, sheet_name='Comparação', index=False)
    df2.to_excel(writer, sheet_name='Comparação Total', index_label='Categorias')
    
    writer.book.remove(writer.book['Tickets dentro da data'])
    
estilizar_excel (novo_arquivo)
gerar_grafico (novo_arquivo)
#------------------------------------------------------------------------------------------------------------------------