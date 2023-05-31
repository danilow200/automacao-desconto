import requests
import pandas as pd

# Cria classe com todas as informações necessarias para criar um desconto

def converter_data(data):
    partes = data.split(" ")
    data_partes = partes[0].split("-")
    nova_data = data_partes[2] + "/" + data_partes[1] + "/" + data_partes[0] + " " + partes[1]
    return nova_data

class Desconto:
    def __init__(self, insert, ticket, observacao, inicio, fim, categoria):
        self.insert = insert
        self.ticket = ticket
        self.observacao = observacao
        self.inicio = inicio
        self.fim = fim
        self.categoria = categoria

def manu_desconto(entrada, linha):

    descontos = []

    # Abri a planilha solicitada
    if entrada == 1:
        empresa = 'PADTEC'
    else:
        empresa = 'RADIANTE'

    nome_do_arquivo = f'.\\planilhas\\Solicitação de Descontos {empresa}.xlsx'

    planilha = pd.read_excel(nome_do_arquivo,sheet_name='Planilha1')
    linha = int(linha) - 2
    # Exclui linhas e colunas extras
    planilha = planilha.drop(index=range(0,linha),axis=0)
    # print(planilha)

    # Adiciona os tickets com descontos para um array de descontos utilizando a classe Desconto para estrutura-lo correntamente
    for index,row in planilha.iterrows():
        if pd.isna(row[empresa]):
            break
        descontos.append(Desconto("descontos", row[empresa], row['Unnamed: 10'], converter_data(str(row["Unnamed: 4"])), converter_data(str(row["Unnamed: 5"])), row["Unnamed: 3"]))

    # Roda o array de descontos e faz a solicitação de desconto
    for i in descontos:
        url = "https://report.telebras.com.br/pages/tickets/control.php?insert=" + i.insert + '&ticket=' + str(i.ticket) + '&observacao=' + i.observacao + '&inicio=' + i.inicio + '&fim=' + i.fim + '&categoria=' + i.categoria
        payload={} 
        headers = { 'Cookie': 'PHPSESSID=578a030e15574ea6c89b3b77590e9353'} 
        response = requests.request("POST", url, headers=headers, data=payload) 
        print(response.text)