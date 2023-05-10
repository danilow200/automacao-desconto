import pandas as pd
from EstilizandoExcel import estilizar_excel
from gerarGrafico import gerar_grafico
import shutil

#------------------------------------------------------------------------------------------------------------------------
                                            #FUNÇÕES

def mescla(x, y):
    x = x.drop(columns=['Unnamed: 0'])
    x = x.drop(index=range(0,1),axis=0)

    y = y.drop(columns=['Unnamed: 0'])
    y = y.drop(index=range(0,2),axis=0)

    # concatenar as duas planilhas
    df_concatenado = pd.concat([x, y])

    # definir a primeira linha como novo header e remover a primeira linha
    df_concatenado.columns = df_concatenado.iloc[0]
    df_concatenado = df_concatenado.iloc[1:]

    return df_concatenado

def soma_tempos(tempo1, tempo2):
    h1, m1, s1 = map(int, tempo1.split(':'))
    h2, m2, s2 = map(int, tempo2.split(':'))
    total = (h1+h2)*3600 + (m1+m2)*60 + s1+s2
    horas, resto = divmod(total, 3600)
    minutos, segundos = divmod(resto, 60)
    return '{:02d}:{:02d}:{:02d}'.format(horas, minutos, segundos)

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
                    '0': 'Acesso', 
                    '1': 'Aguardando CIGR',
                    '2': 'Área de Risco',
                    '3':'Atividade Agendada',
                    '4': 'Falha Restabelecida',
                    '5': 'Falta de Energia',
                    '6': 'Outros',
                    '7': 'Terceiros'
                }

#------------------------------------------------------------------------------------------------------------------------
                                            #VARIAVEIS

categorias = ["Falha Restabelecida", "Acesso", "Aguardando CIGR", "Terceiros", "Área de Risco", "Falta de Energia", "Outros", "Atividade Agendada"]
categoria_soma_dada_pad = {categoria: '00:00:00' for categoria in categorias}
categoria_soma_auto_pad = {categoria: '00:00:00' for categoria in categorias}
categoria_soma_dada_rad = {categoria: '00:00:00' for categoria in categorias}
categoria_soma_auto_rad = {categoria: '00:00:00' for categoria in categorias}
categoria_porcetagem_pad = {categoria: '00:00:00' for categoria in categorias}
categoria_porcetagem_rad = {categoria: '00:00:00' for categoria in categorias}

lista_paginas = [
                    'Codigos Geral', 
                    'Codigos com desconto automatico', 
                    'Comparação',
                    'Comparação Total'
                ]
mes = []
nova_entrada = []

mes_arquivo = '.\\planilhas\\maio.xlsx'
mes_arquivo_copia = mes_arquivo[:-5]
mes_arquivo_copia += ' - copia.xlsx'

arquivo_entrada = '.\\planilhas\\'
arquivo_entrada += input('insira o nome do arquivo\n')
arquivo_entrada += '.xlsx'

#------------------------------------------------------------------------------------------------------------------------
                                            #MESCLA AS PLANILHAS

for nome in lista_paginas: 
    mes.append(pd.read_excel(mes_arquivo, sheet_name=nome))
    nova_entrada.append(pd.read_excel(arquivo_entrada, sheet_name=nome))

nova_df_p1 = mescla(mes[0], nova_entrada[0])
nova_df_p2 = mescla(mes[1], nova_entrada[1])
nova_df_p3 = mescla(mes[2], nova_entrada[2])

# Retira cabeçalho e linha em branco da planilha
mes[3] = mes[3].drop(columns=['Unnamed: 0'])
mes[3] = mes[3].drop(index=range(0,1),axis=0)

nova_entrada[3] = nova_entrada[3].drop(columns=['Unnamed: 0'])
nova_entrada[3] = nova_entrada[3].drop(index=range(0,1),axis=0)

#------------------------------------------------------------------------------------------------------------------------
                                    #GERA NOVA TABELA COM O TOTAL DE HORAS

cont = 2

for categoria in categorias:
    categoria_soma_dada_pad[categoria] = soma_tempos(mes[3]['Unnamed: 2'][cont], nova_entrada[3]['Unnamed: 2'][cont])
    categoria_soma_auto_pad[categoria] = soma_tempos(mes[3]['Unnamed: 3'][cont], nova_entrada[3]['Unnamed: 3'][cont])
    categoria_soma_dada_rad[categoria] = soma_tempos(mes[3]['Unnamed: 4'][cont], nova_entrada[3]['Unnamed: 4'][cont])
    categoria_soma_auto_rad[categoria] = soma_tempos(mes[3]['Unnamed: 5'][cont], nova_entrada[3]['Unnamed: 5'][cont])
    cont += 1

for index in categorias:
    categoria_porcetagem_pad[index] = calculate_time_percentage(categoria_soma_dada_pad[index], categoria_soma_auto_pad[index] )
    categoria_porcetagem_rad[index] = calculate_time_percentage(categoria_soma_dada_rad[index], categoria_soma_auto_rad[index] )

#------------------------------------------------------------------------------------------------------------------------
                                            #CRIA ARQUIVO EXCEL

dataf = {
    'Desconto Dado Total PADTEC': categoria_soma_dada_pad,
    'Desconto Automatico Total PADTEC': categoria_soma_auto_pad,
    'Desconto Dado Total RADIANTE': categoria_soma_dada_rad,
    'Desconto Automatico Total RADIANTE': categoria_soma_auto_rad,
    'Porcetagem PADTEC': categoria_porcetagem_pad,
    'Porcetagem RADIANTE': categoria_porcetagem_rad
}

nova_df_p4 = pd.DataFrame(dataf)

shutil.copy(mes_arquivo, mes_arquivo_copia)

with pd.ExcelWriter(mes_arquivo, engine='xlsxwriter') as writer: #utilizando do Writer para fazer um arquivo com mais de duas páginas
    nova_df_p1.to_excel(writer, sheet_name='Codigos Geral', index = False)
    nova_df_p2.to_excel(writer, sheet_name='Codigos com desconto automatico', index = False)
    nova_df_p3.to_excel(writer, sheet_name='Comparação', index=False)
    nova_df_p4.to_excel(writer, sheet_name='Comparação Total', index_label='Categorias')

estilizar_excel (mes_arquivo)
gerar_grafico (mes_arquivo)