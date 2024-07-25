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

def aceita_auto(cookie):

    chrome_options = Options()
    servico = Service(ChromeDriverManager().install())
    servico2 = Service(executable_path='./chromedriver.exe')
    chrome_options.add_argument("--start-maximized")  # Maximiza a janela do navegador
    chrome_options.add_argument("--disable-extensions")  # Desativa as extensões do Chrome
    chrome_options.add_argument("--disable-gpu")  # Desativa a aceleração de hardware
    chrome_options.add_argument("--disable-dev-shm-usage")  # Desativa o uso compartilhado de memória /tmp
    chrome_options.add_argument("--no-sandbox")  # Desativa o sandbox do Chrome
    chrome_options.add_argument("--force-device-scale-factor=0.75")  # Define o zoom em 25%
    #chrome_options.add_argument('--headless')

    driver = webdriver.Chrome(options=chrome_options, service=servico2)

    driver.get('https://report.telebras.com.br/pages/tickets/tickets.php')
    driver.add_cookie({'name': 'PHPSESSID', 'value': cookie})
    driver.refresh()

    time.sleep(1)

    driver.find_element(By.XPATH,'//*[@id="container"]/a').click()
    time.sleep(1)
    driver.find_element(By.XPATH,'//*[@id="descontos_panel"]/thead/tr/th[6]').click()
    time.sleep(1)

    elemento_tabela = driver.find_element(By.XPATH,'//*[@id="descontos_panel"]' ) #buscando a tabela/pega as informações
    html_content = elemento_tabela.get_attribute('outerHTML') #trazendo o HTML do elemento tabela/ transforma a tabela em variável a partir dos dados HTML

    soup = BeautifulSoup(html_content, 'html.parser')
    tabela = soup.find(name='table')
    pd_tabela = pd.read_html(str(tabela))[0]
    time.sleep(1)

    wait = WebDriverWait(driver, 10)

    cont = 1

    for index,row in pd_tabela.iterrows(): 
        data_descont_atual = datetime.strptime(row['Solicitação'][0:10], '%Y-%m-%d')
        if(row['Status'] == 'pendente'):
            driver.find_element(By.XPATH, f'//*[@id="descontos_panel"]/tbody/tr[{cont}]/td[1]/a').click()
            time.sleep(4)
            atendimento = driver.find_element(By.XPATH, '/html/body/div[12]/div[2]/table/tbody/tr[11]/td[3]/span/span/font')
            if(atendimento != 'desconto maior que período'):
                butao_existe = driver.find_elements(By.CSS_SELECTOR, 'button.aprovar.ui-button.ui-corner-all.ui-widget.ui-button-icon-only')
                if len(butao_existe) == 0:
                    while True:
                        print('não carregou')
                        time.sleep(1)
                        butao_existe = driver.find_elements(By.CSS_SELECTOR, 'button.aprovar.ui-button.ui-corner-all.ui-widget.ui-button-icon-only')
                        if len(butao_existe) > 0:
                            break;
                driver.find_element(By.CSS_SELECTOR, 'button.aprovar.ui-button.ui-corner-all.ui-widget.ui-button-icon-only').click()
                wait.until(expected_conditions.alert_is_present())
                alert = driver.switch_to.alert
                text = alert.text
                alert.accept()

            else:
                butao_existe = driver.find_elements(By.CSS_SELECTOR, 'button.reprovar.ui-button.ui-corner-all.ui-widget.ui-button-icon-only')
                if len(butao_existe) == 0:
                    while True:
                        print('não carregou')
                        time.sleep(1)
                        butao_existe = driver.find_elements(By.CSS_SELECTOR, 'button.reprovar.ui-button.ui-corner-all.ui-widget.ui-button-icon-only')
                        if len(butao_existe) > 0:
                            break;
                driver.find_element(By.CSS_SELECTOR, 'button.reprovar.ui-button.ui-corner-all.ui-widget.ui-button-icon-only').click()
                wait.until(expected_conditions.alert_is_present())
                alert = driver.switch_to.alert
                text = alert.text
                alert.accept()

            driver.find_element(By.CSS_SELECTOR, 'button.ui-dialog-titlebar-close').click()
            time.sleep(2)
        else:
            cont += 1