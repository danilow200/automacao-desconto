##importações para automatizar na web
from selenium import webdriver #importando o nevagador
from selenium.webdriver.common.by import By #serve para achar os elementos no navegador
from selenium.webdriver.common.keys import Keys #importa o teclado para digitar na web
from datetime import datetime
from bs4 import BeautifulSoup
import re
import time
import pandas as pd 
import xlsxwriter

def strfdelta(tdelta, fmt):
    d = {"days": tdelta.days}
    d["hours"], rem = divmod(tdelta.seconds, 3600)
    d["hours"] += 24 * tdelta.days # adicionar horas correspondentes aos dias
    d["minutes"], d["seconds"] = divmod(rem, 60)
    return fmt.format(**d)

def nome_empresa(x):
    if x == "[PAD]": #salva qual foi a empresa responsavel
        return "PADTEC"
    else:
        return "RADIANTE"
    
def tipo_de_codigo(x):
    if x == "$":
        return "Abertura"
    else:
        return "Fechamento"
    
def par_correto(x):
    if x:
        return "Possui"
    else:
        return "Não possui"

#lista de textos
lista_textos = ['[PAD]$IPFA#', '[PAD]#IPFA$', '[PAD]$IPAC#', '[PAD]#IPAC$',
                '[PAD]$IPAR#', '[PAD]#IPAR$', '[PAD]$IPAA#', '[PAD]#IPAA$',
                '[PAD]$IPFR#', '[PAD]#IPFR$', '[PAD]$IPFE#', '[PAD]#IPFE$',
                '[PAD]$IPOS#', '[PAD]#IPOS$', '[PAD]$IPTS#', '[PAD]#IPTS$',
                '[RAD]$IPFA#', '[RAD]#IPFA$', '[RAD]$IPAC#', '[RAD]#IPAC$',
                '[RAD]$IPAR#', '[RAD]#IPAR$', '[RAD]$IPAA#', '[RAD]#IPAA$',
                '[RAD]$IPFR#', '[RAD]#IPFR$', '[RAD]$IPFE#', '[RAD]#IPFE$',
                '[RAD]$IPOS#', '[RAD]#IPOS$', '[RAD]$IPTS#', '[RAD]#IPTS$',]

#Array para todos os tickets que possuem algum dos códigos da lista de textos
tickets_codigo = []
codigo_codigos = []
data_codigos = []
estacao_codigos = []
empresa_codigos = []
estado_codigos = []
categoria_codigos = []
tipo_codigos = []
possui_par = []

tickets_detro_data = []

#arrays utilizado para a tabela de desconto
desconto_auto = []
tickets_auto = []
desconto_abertura = []
desconto_fechamento = []
data_abertura = []
data_fechamento = []
empresa_auto = []
codigo_auto = []

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

#Site da Telebras e Planilha que será analisada
url_logs = "https://report.telebras.com.br/scripts/get_incidentes.php" # variável que armazena o link do site que vamos pesquisar
nome_do_arquivo = 'Indicadores - Abril.xlsx' #armazenando o nome da planilha em uma variável


#Ler Folha de Descontos dentro da Planilha Indicadores - Março
numero_tickets = pd.read_excel(nome_do_arquivo,sheet_name='Incidentes') #lendo a planilha do excel
numero_tickets = numero_tickets.drop(index=range(0,2),axis=0)
numero_tickets = numero_tickets.drop_duplicates(subset=['Unnamed: 0']) #exclui as linhas com números de tickets duplicados
#após isso, a tabela passa a ter o número de linhas e colunas que sobraram.

cont2 = 0 #contador da posição data para uso do calculo do desconto

while True:
    entrada = input('Digite a data de entrada no formato dia/mes/ano:\n')
    pattern = re.compile(r"\d{2}/\d{2}/\d{4}")
    if pattern.match(entrada):
        break
    
while True:
    entrada2 = input('Digite a ultima data de entrada no formato dia/mes/ano:\n')
    pattern = re.compile(r"\d{2}/\d{2}/\d{4}")
    if pattern.match(entrada2):
        break
    

entrada2_data = datetime.strptime(entrada2, '%d/%m/%Y')

data_validada = False

for index,row in numero_tickets.iterrows():  #Loop que indica o número de repetições que o navegador deve ser aberto e fechado
    data_compara = datetime.strptime(row['Unnamed: 5'][0:10], '%d/%m/%Y')
    
    if data_validada == False:
        if row['Unnamed: 5'][0:10] == entrada:
            data_validada = True
            
    elif data_compara > entrada2_data:
        break

    else:
        tickets_detro_data.append(row['Unnamed: 0'])
        options = webdriver.ChromeOptions()
        options.add_argument("--headless") #define para o chrome abrir em segundo plano
        chrome = webdriver.Chrome(executable_path='chromedriver.exe', chrome_options=options) #cria uma instância do chrome
        chrome.get(url_logs)#navega para essa url do chrome    
        time.sleep(1) #Delay 
        
        elemento_ticket = chrome.find_element(By.XPATH,'//*[@id="filter-box"]')  #"encontra" o campo de preenchimento de ticket
        elemento_ticket.send_keys(row['Unnamed: 0']) #cola o número do ticket no campo
        elemento_botao = chrome.find_element(By.XPATH,'//*[@id="filter-clear"]').click() #encontra e depois clica no botão "enviar"
        
        print(row['Unnamed: 0']) #prita no terminal o ticket atual
        
        tabela_existe = chrome.find_elements(By.XPATH, '//*[@id="manual"]/table/tbody') #cria um array de elementos tabela para saber se a tabela foi rederinzada na página
        
        elemento_estacao = chrome.find_element(By.XPATH,'//*[@id="content"]/table/tbody/tr[1]/td[2]' ) #busca estacao no navegador
        estacao = elemento_estacao.get_attribute('innerHTML')
        
        if len(tabela_existe) == 0 or estacao == " ": #checa se a tabela e a estação foram carregados na página
            while True: #roda loop até que ela seja carregada
                print(f"salve ticket {row['Unnamed: 0']}")
                chrome.refresh() #recarregar página do navegador
                time.sleep(1)
                tabela_existe = chrome.find_elements(By.XPATH, '//*[@id="manual"]/table/tbody')
                elemento_estacao = chrome.find_element(By.XPATH,'//*[@id="content"]/table/tbody/tr[1]/td[2]' ) #busca estacao no navegador
                estacao = elemento_estacao.get_attribute('innerHTML')
                if len(tabela_existe) > 0 and estacao != " ": #para o loop quando a tabela for carregada
                    break
            
        
        elemento_tabela = chrome.find_element(By.XPATH,'//*[@id="manual"]/table' ) #buscando a tabela/pega as informações
        html_content = elemento_tabela.get_attribute('outerHTML') #trazendo o HTML do elemento tabela/ transforma a tabela em variável a partir dos dados HTML

        #Parsear o conteúdo HTML utilizando a BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        table = soup.find(name='table')
            
        pd_tabela = pd.read_html(str(table))[0] #passando a tabela para o pandas
#------------------------------------------------------------------------------------------------------------------------
                                            #LENDO SEGUNDA TABELA
        tabela_existe = chrome.find_elements(By.XPATH, '//*[@id="automatico"]/table/tbody') #cria um array de elementos tabela para saber se a tabela foi rederinzada na página
        
        if len(tabela_existe) == 0: #checa se a tabela e a estação foram carregados na página
            while True: #roda loop até que ela seja carregada
                print(f"salve ticket de novo {row['Unnamed: 0']}")
                chrome.refresh() #recarregar página do navegador
                time.sleep(1)
                tabela_existe = chrome.find_elements(By.XPATH, '//*[@id="automatico"]/table/tbody')
                if len(tabela_existe) > 0: #para o loop quando a tabela for carregada
                    break
            
        
        elemento_tabela2 = chrome.find_element(By.XPATH,'//*[@id="automatico"]/table' ) #buscando a tabela/pega as informações
        html_content2 = elemento_tabela2.get_attribute('outerHTML') #trazendo o HTML do elemento tabela/ transforma a tabela em variável a partir dos dados HTML

        #Parsear o conteúdo HTML utilizando a BeautifulSoup
        soup2 = BeautifulSoup(html_content2, 'html.parser')
        table2 = soup2.find(name='table')
            
        pd_tabela_2 = pd.read_html(str(table2))[0] #passando a tabela para o pandas
        
        ultima_entrada = str('')
        
        for index_tabela2,row_tabela2 in reversed(list(pd_tabela.iterrows())):
            
            if type(row_tabela2[5]) == float: #caso tenha uma celula vazia na tabela, ela será convertida para string
                row_tabela2[5] = str(row_tabela2[5])
            
            for texto in lista_textos:
                #verifica se os códigos estão presente na coluna 5
                if texto in row_tabela2[5]:
                    #se for true, excuta:
                    tickets_codigo.append(row['Unnamed: 0']) #jogar no array o número do ticket
                    codigo_codigos.append(texto) #se o string codigo receber algo diferente de vazio, ou seja, receber texto. Ai o array recebe a string
                    data_codigos.append(row_tabela2[0]) #salva data da ocorrencia
                    estacao_codigos.append(estacao) #salva em qual estação ocorreu
                    categoria_codigos.append(categoria_dicionario[texto[6:10]])
                    tipo_codigos.append(tipo_de_codigo(texto[5:6]))
                    if estacao[0] == ' ':
                        estado_codigos.append(estacao[1:3])
                    else:
                        estado_codigos.append(estacao[0:2]) #salva qual estado pertence a estacao
                    empresa_codigos.append(nome_empresa(texto[0:5]))
                    
                    par = False
                    
                    if texto[6:10] == ultima_entrada[6:10] and texto[-1] != ultima_entrada[-1] and texto[-1] != "#":
                        data1 = datetime.strptime(data_codigos[cont2 - 1], '%d/%m/%Y %H:%M') #datetime faz tratamento para que as strings que vem em formato de dado para um novo formato capaz de fazer seus calculos
                        data2 = datetime.strptime(data_codigos[cont2], '%d/%m/%Y %H:%M')
                        data_abertura.append(data_codigos[cont2 - 1])
                        data_fechamento.append(data_codigos[cont2])
                        diferenca = data2 - data1 #calcula o desconto
                        desconto_abertura.append(ultima_entrada)
                        desconto_fechamento.append(texto)
                        # Converter a diferença em uma string com o formato horas:minutos:segundos
                        diferenca_str = strfdelta(diferenca, "{hours:02d}:{minutes:02d}:{seconds:02d}")
                        desconto_auto.append(diferenca_str)
                        tickets_auto.append(row['Unnamed: 0'])
                        codigo_auto.append(categoria_dicionario[texto[6:10]])
                        empresa_auto.append(nome_empresa(texto[0:5]))
                        par = True
                        possui_par[cont2 - 1] = "Possui"
                    
                    possui_par.append(par_correto(par))
                        
                    ultima_entrada = texto
                    cont2 += 1
                    
        if ultima_entrada[5:11]  == '$IPFE#' or ultima_entrada[5:11] =='$IPFR#':
            tickets_codigo.append(row['Unnamed: 0']) #jogar no array o número do ticket
            codigo_codigos.append('Fechamento junto com a ocorrencia') #se o string codigo receber algo diferente de vazio, ou seja, receber texto. Ai o array recebe a string
            estacao_codigos.append(estacao) #salva em qual estação ocorreu
            categoria_codigos.append(categoria_dicionario[ultima_entrada[6:10]])
            tipo_codigos.append('Fechamento')
            if estacao[0] == ' ':
                estado_codigos.append(estacao[1:3])
            else:
                estado_codigos.append(estacao[0:2]) #salva qual estado pertence a estacao
            empresa_codigos.append(nome_empresa(ultima_entrada[0:5]))
            data1 = datetime.strptime(data_codigos[cont2 - 1], '%d/%m/%Y %H:%M')
            valida = True
            if ultima_entrada[5:11]  == '$IPFE#':
                for index_tabela3,row_tabela3 in reversed(list(pd_tabela_2.iterrows())):
                    if row_tabela3['Categoria'] == 'Direcionamento da Solicitação':
                        data2 = datetime.strptime(pd_tabela_2['Informações da ocorrência'][index_tabela3][0:16], '%d/%m/%Y %H:%M')
                        if data2 > data1:
                            data_codigos.append(pd_tabela_2['Informações da ocorrência'][index_tabela3][0:16])
                            valida = False
            if valida:
                data_codigos.append(pd_tabela_2['Informações da ocorrência'][0][0:16]) #salva data da ocorrencia
                data2 = datetime.strptime(data_codigos[cont2], '%d/%m/%Y %H:%M')
            data_abertura.append(data_codigos[cont2 - 1])
            data_fechamento.append(data_codigos[cont2])
            diferenca = data2 - data1 #calcula o desconto
            desconto_abertura.append(ultima_entrada)
            desconto_fechamento.append('Fechamento junto com a ocorrencia')
            # Converter a diferença em uma string com o formato horas:minutos:segundos
            diferenca_str = strfdelta(diferenca, "{hours:02d}:{minutes:02d}:{seconds:02d}")
            desconto_auto.append(diferenca_str)
            tickets_auto.append(row['Unnamed: 0'])
            codigo_auto.append(categoria_dicionario[ultima_entrada[6:10]])
            empresa_auto.append(nome_empresa(ultima_entrada[0:5]))
            par = True
            possui_par[cont2 - 1] = "Possui"
                    
            possui_par.append(par_correto(par))
            cont2 += 1
                        
        chrome.quit #fecha o chrome após terminar a operação desejada
        
#Fazer a tabela no Pandas
data = {
        'Tickets': tickets_codigo, 
        'Códigos': codigo_codigos, 
        'Data': data_codigos, 
        'Estação': estacao_codigos, 
        'Empresa': empresa_codigos, 
        'Estado': estado_codigos,
        'Categoria': categoria_codigos,
        'Tipo': tipo_codigos,
        'Par': possui_par
    } # Criando uma variavel data para a tabela ficar na ordem correta
df=pd.DataFrame(data) #dessa forma o data frame é printado com uma coluna contendo os tickets e uma coluna contendo todos os códigos

data2 = {
    'Tickets': tickets_auto,
    'Abertura': desconto_abertura,
    'Fechamento': desconto_fechamento,
    'Categoria': codigo_auto,
    'Data de Abertura': data_abertura,
    'Data de Fechamento': data_fechamento,
    'Desconto': desconto_auto,
    'Empresa': empresa_auto
}
df2 = pd.DataFrame(data2)

data3 = {
    'Tickets dentro da data': tickets_detro_data
}
df3 = pd.DataFrame(data3)

with pd.ExcelWriter('descontos abril.xlsx', engine='xlsxwriter') as writer: #utilizando do Writer para fazer um arquivo com mais de duas páginas
    df.to_excel(writer, sheet_name='Codigos Geral', index = False)
    df2.to_excel(writer, sheet_name='Codigos com desconto automatico', index = False)
    df3.to_excel(writer, sheet_name='Tickets dentro da data', index = False)