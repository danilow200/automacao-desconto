import requests
import pandas as pd
from datetime import datetime
from selenium import webdriver #importando o nevagador
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time

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

def manu_desconto(entrada, linha, cookie):

    chrome_options = Options()
    servico = Service(ChromeDriverManager().install())
    chrome_options.add_argument("--start-maximized")  # Maximiza a janela do navegador
    chrome_options.add_argument("--disable-extensions")  # Desativa as extensões do Chrome
    chrome_options.add_argument("--disable-gpu")  # Desativa a aceleração de hardware
    chrome_options.add_argument("--disable-dev-shm-usage")  # Desativa o uso compartilhado de memória /tmp
    chrome_options.add_argument("--no-sandbox")  # Desativa o sandbox do Chrome
    chrome_options.add_argument("--force-device-scale-factor=0.75")  # Define o zoom em 25%
    #chrome_options.add_argument('--headless')

    driver = webdriver.Chrome(options=chrome_options, service=servico)

    descontos = []
    expurgo = []
    neutraliza = []

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

    cookie_atu = 'PHPSESSID=' + cookie

    # Adiciona os tickets com descontos para um array de descontos utilizando a classe Desconto para estrutura-lo correntamente
    for index,row in planilha.iterrows():
        if pd.isna(row[empresa]):
            break
        if row['Unnamed: 9'] == 'Aprovado' or row['Unnamed: 9'] == 'APROVADO':
            descontos.append(Desconto("descontos", row[empresa], f"{empresa}: {row['Unnamed: 10']}", converter_data(str(row["Unnamed: 4"])), converter_data(str(row["Unnamed: 5"])), row["Unnamed: 3"]))
        elif str(row['Unnamed: 9']).capitalize() == 'Expurgado':
            expurgo.append(row[empresa])
        elif str(row['Unnamed: 9']).capitalize() == 'Neutralizado':
            neutraliza.append(row[empresa])

    # Roda o array de descontos e faz a solicitação de desconto
    for i in descontos:
        url = "https://report.telebras.com.br/pages/tickets/control.php?insert=" + i.insert + '&ticket=' + str(i.ticket) + '&observacao=' + i.observacao + '&inicio=' + i.inicio + '&fim=' + i.fim + '&categoria=' + i.categoria
        payload={} 
        headers = { 'Cookie': cookie_atu} 
        response = requests.request("POST", url, headers=headers, data=payload) 
        print(response.text)

    print(expurgo)
    print(neutraliza)

    if len(expurgo) != 0 or len(neutraliza) != 0:
        driver.get('https://report.telebras.com.br/pages/tickets/tickets.php')
        driver.add_cookie({'name': 'PHPSESSID', 'value': cookie})
        driver.refresh()
        time.sleep(1)

        elemento_ticket = driver.find_element(By.XPATH,'/html/body/div[5]/form/div/div[1]/label/span/span[2]/input')

        for i in expurgo:
            elemento_ticket.send_keys(str(i))
            time.sleep(1)
            driver.find_element(By.XPATH, '/html/body/div[5]/form/div/table/tbody/tr/td[1]/a').click()
            time.sleep(2)
            driver.find_element(By.ID, 'expurgar').click()
            alert = driver.switch_to.alert
            alert.accept()
            time.sleep(2)
            elemento_ticket.clear()
        
        for i in neutraliza:
            elemento_ticket.send_keys(str(i))
            time.sleep(1)
            driver.find_element(By.XPATH, '/html/body/div[5]/form/div/table/tbody/tr/td[1]/a').click()
            time.sleep(2)
            driver.find_element(By.ID, 'tma').click()
            time.sleep(1)
            driver.find_element(By.ID, 'tma').click()
            # driver.find_element(By.ID, 'confirmar_tma').click()
            driver.find_element(By.CSS_SELECTOR, 'button.ui-dialog-titlebar-close').click()
            time.sleep(2)
            elemento_ticket.clear()