from datetime import datetime
from selenium import webdriver #importando o nevagador
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup
import pandas as pd
import time
import re

def converte_data(data):
    dia, mes, ano = data.split('/')
    data_formatada = ano + '-' + mes + '-' + dia
    return data_formatada

chrome_options = Options()
chrome_options.add_argument("--start-maximized")  # Maximiza a janela do navegador
chrome_options.add_argument("--disable-extensions")  # Desativa as extensões do Chrome
chrome_options.add_argument("--disable-gpu")  # Desativa a aceleração de hardware
chrome_options.add_argument("--disable-dev-shm-usage")  # Desativa o uso compartilhado de memória /tmp
chrome_options.add_argument("--no-sandbox")  # Desativa o sandbox do Chrome
chrome_options.add_argument("--force-device-scale-factor=0.75")  # Define o zoom em 25%
#chrome_options.add_argument('--headless')

while True:
    entrada = input('Digite a data de entrada no formato dia/mes/ano:\n')
    pattern = re.compile(r"\d{2}/\d{2}/\d{4}")
    if pattern.match(entrada):
        break

entra_convertida = datetime.strptime(converte_data(entrada), '%Y-%m-%d')

driver = webdriver.Chrome(options=chrome_options)

driver.get('https://report.telebras.com.br/pages/tickets/tickets.php')
driver.add_cookie({'name': 'PHPSESSID', 'value': '578a030e15574ea6c89b3b77590e9353'})
driver.refresh()

time.sleep(1)

driver.find_element(By.XPATH,'//*[@id="container"]/a').click()
time.sleep(1)


elemento_tabela = driver.find_element(By.XPATH,'//*[@id="descontos_panel"]' ) #buscando a tabela/pega as informações
html_content = elemento_tabela.get_attribute('outerHTML') #trazendo o HTML do elemento tabela/ transforma a tabela em variável a partir dos dados HTML

soup = BeautifulSoup(html_content, 'html.parser')
tabela = soup.find(name='table')
pd_tabela = pd.read_html(str(tabela))[0]
# print(pd_tabela)
time.sleep(1)

wait = WebDriverWait(driver, 10)

cont = 1

for index,row in pd_tabela.iterrows(): 
    data_descont_atual = datetime.strptime(row['Solicitação'][0:10], '%Y-%m-%d')
    if(data_descont_atual <= entra_convertida and row['Status'] == 'pendente' and row['Solicitante'] == 'danilo.silva'):
        driver.find_element(By.XPATH, f'//*[@id="descontos_panel"]/tbody/tr[{cont}]/td[1]/a').click()
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, 'button.aprovar.ui-button.ui-corner-all.ui-widget.ui-button-icon-only').click()
        #driver.find_element(By.CSS_SELECTOR, 'button.lixeira.ui-button.ui-corner-all.ui-widget.ui-button-icon-only').click()
        wait.until(expected_conditions.alert_is_present())
        alert = driver.switch_to.alert
        text = alert.text
        alert.accept()
        driver.find_element(By.CSS_SELECTOR, 'button.ui-dialog-titlebar-close').click()
        time.sleep(1)
    else:
        cont += 1


time.sleep(3)