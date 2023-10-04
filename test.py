from datetime import datetime
from selenium import webdriver #importando o nevagador
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time

def converte_data(data):
    dia, mes, ano = data.split('/')
    data_formatada = ano + '-' + mes + '-' + dia
    return data_formatada


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

driver.get('https://report.telebras.com.br/pages/tickets/tickets.php')
driver.add_cookie({'name': 'PHPSESSID', 'value': '578a030e15574ea6c89b3b77590e9353'})
driver.refresh()

time.sleep(1)

elemento_ticket = driver.find_element(By.XPATH,'/html/body/div[5]/form/div/div[1]/label/span/span[2]/input')  #"encontra" o campo de preenchimento de ticket
elemento_ticket.send_keys('3863007')

time.sleep(1)

# ---------------------LIMPAR INPUT

# elemento_ticket.clear()

# elemento_ticket.send_keys('3862007')

# time.sleep(1)

# -------------------DELETAR

driver.find_element(By.XPATH, '/html/body/div[5]/form/div/table/tbody/tr/td[1]/a').click()
time.sleep(2)
driver.find_element(By.ID, 'expurgar').click()
alert = driver.switch_to.alert
alert.dismiss()
driver.find_element(By.CSS_SELECTOR, 'button.ui-dialog-titlebar-close').click()
time.sleep(2)
elemento_ticket.clear()

driver.find_element(By.XPATH, '/html/body/div[5]/form/div/table/tbody/tr/td[1]/a').click()
time.sleep(2)
driver.find_element(By.ID, 'tma').click()
time.sleep(1)
driver.find_element(By.ID, 'tma').click()
driver.find_element(By.ID, 'confirmar_tma').click()

# time.sleep(10)