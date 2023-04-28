import requests
import json
import pandas as pd

# class Desconto:
#     def __init__(self, insert, ticket, observacao, inicio, fim, categoria):
#         self.insert = insert
#         self.ticket = ticket
#         self.observacao = observacao
#         self.inicio = inicio
#         self.fim = fim
#         self.categoria = categoria

# descontos = []
# dados = []

# nome_do_arquivo = '.\\planilhas\\abril.xlsx'

# planilha = pd.read_excel(nome_do_arquivo,sheet_name='Codigos com desconto automatico')
# planilha = planilha.drop(columns=['Unnamed: 0'])
# planilha = planilha.drop(index=range(0,1),axis=0)


# for index,row in planilha[1:].iterrows():
#     descontos.append(Desconto("descontos", row["Unnamed: 1"], "Desconto Automatico", row["Unnamed: 5"], row["Unnamed: 6"], row["Unnamed: 4"]))

# for i in descontos:
#     dados.append({
#         "insert": i.insert, 
#         "ticket": i.ticket, 
#         "observacao": i.observacao, 
#         "inicio": i.inicio, 
#         "fim": i.fim, 
#         "categoria": i.categoria
#     })

# dados_json = json.dumps(dados, indent=4)
# print(dados_json)

#_____________________________________________________________________________________________________________________

data = { 
            "insert":"descontos", 
            "ticket":3297906, 
            "observacao":"teste", 
            "inicio":"03-04-2023 12:12:00", 
            "fim":"03-04-2023 14:26:00", 
            "categoria":"Terceiros"
        }

# Fazer uma solicitação GET para obter o cookie da sessão
response = requests.get('https://report.telebras.com.br/index.php')
cookie = {'PHPSESSID': '578a030e15574ea6c89b3b77590e9353'}

# Fazer outra solicitação GET com o cookie da sessão
url = 'https://report.telebras.com.br/pages/tickets/control.php'
response = requests.post(url, cookies=cookie, data=data)

print(response.status_code)