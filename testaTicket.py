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

#------------------------------------------------------------------------------------------------------------------------
                                            #FUNÇÕES

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

def ler_indicadores(mes, data_inicio, data_fim):
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

    #------------------------------------------------------------------------------------------------------------------------
                                                #DECLARAÇÃO DE VARIAVEIS

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



    #Site da Telebras e Planilha que será analisada
    url_logs = "https://report.telebras.com.br/scripts/get_incidentes.php" # variável que armazena o link do site que vamos pesquisar
    # mes_arquivo = input('Informe o mês da planilha de Indiacadores\n')
    nome_do_arquivo = '.\\planilhas\\Indicadores - ' + mes.capitalize() + '.xlsx' #armazenando o nome da planilha em uma variável


    def insere_data_desconto(str1, str2, abertura, fechamento):
        data1 = datetime.strptime(str1, '%d/%m/%Y %H:%M') #datetime faz tratamento para que as strings que vem em formato de dado para um novo formato capaz de fazer seus calculos
        data2 = datetime.strptime(str2, '%d/%m/%Y %H:%M')
        data_abertura.append(str1)
        data_fechamento.append(str2)
        diferenca = data2 - data1 #calcula o desconto
        desconto_abertura.append(abertura)
        desconto_fechamento.append(fechamento)
        diferenca_str = strfdelta(diferenca, "{hours:02d}:{minutes:02d}:{seconds:02d}")
        desconto_auto.append(diferenca_str)

    def insere_codigo(ticket, codigo, data, est, cat, tipo_cod):
        tickets_codigo.append(ticket) #jogar no array o número do ticket
        codigo_codigos.append(codigo) #se o string codigo receber algo diferente de vazio, ou seja, receber texto. Ai o array recebe a string
        data_codigos.append(data) #salva data da ocorrencia
        estacao_codigos.append(est) #salva em qual estação ocorreu
        categoria_codigos.append(cat)
        tipo_codigos.append(tipo_cod)

    #------------------------------------------------------------------------------------------------------------------------
                                                #LEITURA DA PLANILHA

    #Ler Folha de Descontos dentro da Planilha Indicadores - Março
    numero_tickets = pd.read_excel(nome_do_arquivo,sheet_name='Incidentes') #lendo a planilha do excel
    numero_tickets = numero_tickets.drop(index=range(0,2),axis=0)
    numero_tickets = numero_tickets.drop_duplicates(subset=['Unnamed: 0']) #exclui as linhas com números de tickets duplicados
    #após isso, a tabela passa a ter o número de linhas e colunas que sobraram.

    cont2 = 0 #contador da posição data para uso do calculo do desconto

    #------------------------------------------------------------------------------------------------------------------------
                                                #ENTRADA DO CÓDIGO

    # data de inicio
    # while True:
    #     entrada = input('Digite a data de entrada no formato dia/mes/ano:\n')
    #     pattern = re.compile(r"\d{2}/\d{2}/\d{4}")
    #     if pattern.match(entrada):
    #         break

    # # ultima data para leituta    
    # while True:
    #     entrada2 = input('Digite a ultima data de entrada no formato dia/mes/ano:\n')
    #     pattern = re.compile(r"\d{2}/\d{2}/\d{4}")
    #     if pattern.match(entrada2):
    #         break
        

    entrada2_data = datetime.strptime(data_fim, '%d/%m/%Y')

    data_validada = False

    #------------------------------------------------------------------------------------------------------------------------
                                        #LEITURA DOS TICKETS DA PLANILHA PARA E USO DO SELENIUM

    for index,row in numero_tickets.iterrows():  #Loop que indica o número de repetições que o navegador deve ser aberto e fechado
        data_compara = datetime.strptime(row['Unnamed: 5'][0:10], '%d/%m/%Y')
        
        if data_validada == False:
            if row['Unnamed: 5'][0:10] == data_inicio:
                data_validada = True
                
        elif data_compara > entrada2_data:
            break

        if data_validada:
            tickets_detro_data.append(row['Unnamed: 0'])
            options = webdriver.ChromeOptions()
            options.add_argument("--headless") #define para o chrome abrir em segundo plano
            # options.add_argument("--force-device-scale-factor=0.75")  # Define o zoom em 75%
            chrome = webdriver.Chrome(executable_path='chromedriver.exe', chrome_options=options) #cria uma instância do chrome
            chrome.get(url_logs)#navega para essa url do chrome    
            time.sleep(1) #Delay 
            
            elemento_ticket = chrome.find_element(By.XPATH,'//*[@id="filter-box"]')  #"encontra" o campo de preenchimento de ticket
            elemento_ticket.send_keys(row['Unnamed: 0']) #cola o número do ticket no campo
            elemento_botao = chrome.find_element(By.XPATH,'//*[@id="filter-box"]')
            # Simula o pressionamento da tecla "Enter"
            elemento_botao.send_keys(Keys.RETURN)
            
            print(row['Unnamed: 0']) #prita no terminal o ticket atual

    #------------------------------------------------------------------------------------------------------------------------
                                        #LEITURA DA PRIMEIRA TABELA DO TICKET  
                                            
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
                                                #LENDO SEGUNDA TABELA DO TICKET
                                                
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
            
    #------------------------------------------------------------------------------------------------------------------------
                                #ANALISA PRIMEIRA TABELA EM BUSCA DE CODIGOS PARA DESCONTO
            
            for index_tabela2,row_tabela2 in reversed(list(pd_tabela.iterrows())):
                
                if type(row_tabela2[5]) == float: #caso tenha uma celula vazia na tabela, ela será convertida para string
                    row_tabela2[5] = str(row_tabela2[5])
                
                for texto in lista_textos:
                    #verifica se os códigos estão presente na coluna 5
                    if texto in row_tabela2[5]:
                        #se for true, excuta:
                        # if texto != '$IPFR#':
                        insere_codigo(row['Unnamed: 0'], texto, row_tabela2[0], estacao, categoria_dicionario[texto[6:10]], tipo_de_codigo(texto[5:6]))
                        if estacao[0] == ' ':
                            estado_codigos.append(estacao[1:3])
                        else:
                            estado_codigos.append(estacao[0:2]) #salva qual estado pertence a estacao
                        empresa_codigos.append(nome_empresa(texto[0:5]))
                        
                        par = False
                        
                        if texto[6:10] == ultima_entrada[6:10] and texto[-1] != ultima_entrada[-1] and texto[-1] != "#":
                            insere_data_desconto(data_codigos[cont2 - 1], data_codigos[cont2], ultima_entrada, texto)
                            tickets_auto.append(row['Unnamed: 0'])
                            codigo_auto.append(categoria_dicionario[texto[6:10]])
                            empresa_auto.append(nome_empresa(texto[0:5]))
                            par = True
                            possui_par[cont2 - 1] = "Possui"
                        
                        possui_par.append(par_correto(par))
                            
                        ultima_entrada = texto
                        cont2 += 1
                        
    #------------------------------------------------------------------------------------------------------------------------
                #ANALISA SEGUNDA TABELA EM BUSCA PARA APLICAR DESCONTOS QUE NÃO FOI POSSIVEL APENAS COM A PRIMEIRA TABELA
                                
            if ultima_entrada[5:11]  == '$IPFE#':
                if estacao[0] == ' ':
                    estado_codigos.append(estacao[1:3])
                else:
                    estado_codigos.append(estacao[0:2]) #salva qual estado pertence a estacao
                empresa_codigos.append(nome_empresa(ultima_entrada[0:5]))
                valida = True
                aux = 0
                for index_tabela3,row_tabela3 in reversed(list(pd_tabela_2.iterrows())):
                    if row_tabela3['Categoria'] == 'Ocorrências: Direcionamento da tarefa Diagnosticar para o grupo N1':
                        data2 = datetime.strptime(pd_tabela_2['Informações da ocorrência'][index_tabela3 - 2][0:16], '%d/%m/%Y %H:%M')
                        aux = index_tabela3 - 2
                        break
                if valida:
                    insere_codigo(row['Unnamed: 0'], 'Direcionamento da tarefa Diagnosticar para o grupo N1', pd_tabela_2['Informações da ocorrência'][aux][0:16], estacao, categoria_dicionario[texto[6:10]], 'Fechamento')
                    insere_data_desconto(data_codigos[cont2 - 1], pd_tabela_2['Informações da ocorrência'][aux][0:16], ultima_entrada, 'Fechamento junto com a ocorrencia')
                tickets_auto.append(row['Unnamed: 0'])
                codigo_auto.append(categoria_dicionario[ultima_entrada[6:10]])
                empresa_auto.append(nome_empresa(ultima_entrada[0:5]))
                par = True
                possui_par[cont2 - 1] = "Possui"
                        
                possui_par.append(par_correto(par))
                cont2 += 1

            valida_categoria = False
            if ultima_entrada[5:11]  == '$IPFR#':
                for index_tabela3,row_tabela3 in reversed(list(pd_tabela_2.iterrows())):
                    if row_tabela3['Categoria'] == 'Solicitação Restaurada' and valida_categoria == False:
                        for i in range(2):
                            possui_par.append('Possui')
                            if estacao[0] == ' ':
                                estado_codigos.append(estacao[1:3])
                            else:
                                estado_codigos.append(estacao[0:2]) #salva qual estado pertence a estacao
                            empresa_codigos.append(nome_empresa(ultima_entrada[0:5])) #achar outra solução
                        insere_codigo(row['Unnamed: 0'], 'Solicitação Restaurada', row_tabela3['Informações da ocorrência'][0:16], estacao, 'Falha Restabelecida', 'Abertura')
                        
                        cont2 += 1
                        # insere_data_desconto(row_tabela3['Informações da ocorrência'][0:16], pd_tabela_2['Informações da ocorrência'][0][0:16], 'Solicitação Restaurada Abertura', 'Solicitação Restaurada Fechamento')
                        valida_categoria = True
                        salva_info = row_tabela3['Informações da ocorrência'][0:16]

                    if valida_categoria:
                        if row_tabela3['Categoria'] == 'Ocorrências: Direcionamento da tarefa Fechar para o grupo N1' or row_tabela3['Categoria'] == 'Ocorrências: Direcionamento da tarefa Diagnosticar para o grupo N1':
                            insere_codigo(row['Unnamed: 0'], 'Direcionamento da tarefa Fechar para o grupo N1', pd_tabela_2['Informações da ocorrência'][index_tabela3 - 2][0:16], estacao, 'Falha Restabelecida', 'Fechamento')
                            insere_data_desconto(salva_info, pd_tabela_2['Informações da ocorrência'][index_tabela3 - 2][0:16], 'Solicitação Restaurada', 'Direcionamento da tarefa Fechar para o grupo N1')
                            tickets_auto.append(row['Unnamed: 0'])
                            codigo_auto.append('Falha Restabelecida')
                            empresa_auto.append(nome_empresa(ultima_entrada[0:5]))
                            cont2 += 1
                            break
                            
            chrome.quit #fecha o chrome após terminar a operação desejada
            
            
    #------------------------------------------------------------------------------------------------------------------------
                                    #CRIAÇÃO DOS DATAFRAMES E DO DOCUMENTO EXCEL
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

    with pd.ExcelWriter('.\\planilhas\\descontos.xlsx', engine='xlsxwriter') as writer: #utilizando do Writer para fazer um arquivo com mais de duas páginas
        df.to_excel(writer, sheet_name='Codigos Geral', index = False)
        df2.to_excel(writer, sheet_name='Codigos com desconto automatico', index = False)
        df3.to_excel(writer, sheet_name='Tickets dentro da data', index = False)