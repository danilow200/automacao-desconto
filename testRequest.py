import requests
import pandas as pd

# Cria classe com todas as informações necessarias para criar um desconto
class Desconto:
    def __init__(self, insert, ticket, observacao, inicio, fim, categoria):
        self.insert = insert
        self.ticket = ticket
        self.observacao = observacao
        self.inicio = inicio
        self.fim = fim
        self.categoria = categoria

def auto_request(nome_arquivo, cookie):

    descontos = []

    # Abri a planilha solicitada
    nome_do_arquivo = '.\\planilhas\\' + nome_arquivo + '.xlsx'

    planilha = pd.read_excel(nome_do_arquivo,sheet_name='Codigos com desconto automatico')
    # Exclui linhas e colunas extras
    planilha = planilha.drop(columns=['Unnamed: 0'])
    planilha = planilha.drop(index=range(0,1),axis=0)

    cookie_atu = 'PHPSESSID=' + cookie

    # Adiciona os tickets com descontos para um array de descontos utilizando a classe Desconto para estrutura-lo correntamente
    for index,row in planilha[1:].iterrows():
        descontos.append(Desconto("descontos", row["Unnamed: 1"], f"{row['Unnamed: 8']}: Desconto Automático", row["Unnamed: 5"] + ':00', row["Unnamed: 6"] + ':00', row["Unnamed: 4"]))

    # Roda o array de descontos e faz a solicitação de desconto
    for i in descontos:
        url = "https://report.telebras.com.br/pages/tickets/control.php?insert=" + i.insert + '&ticket=' + str(i.ticket) + '&observacao=' + i.observacao + '&inicio=' + i.inicio + '&fim=' + i.fim + '&categoria=' + i.categoria
        payload={} 
        headers = { 'Cookie': cookie_atu} 
        response = requests.request("POST", url, headers=headers, data=payload) 
        print(response.text)