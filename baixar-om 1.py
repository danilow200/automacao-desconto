import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

file1 = 'Lista_modelos_bilhete.xls'
file2 = 'Lista_modelos_bilhete (1).xls'
file3 = 'Lista_modelos_bilhete (2).xls'
file4 = 'Lista_modelos_bilhete (3).xls'
arquivo_renomeado = 'OM_combinado.xlsx' 

def is_download_complete(download_dir, file_name):
    files = os.listdir(download_dir)
    for file in files:
        if file == file_name:
            return False
    return True

timeout = 300  

url = "https://oss.telebras.com.br/cpqdom-web/login.xhtml"
url_tabela = "https://oss.telebras.com.br/cpqdom-web/operation/OrderQueryList.xhtml"
url_tabela_2 = "https://oss.telebras.com.br/cpqdom-web/operation/ActivityQueryList.xhtml"
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-extensions")
options.add_argument("--disable-gpu")
options.add_argument("--force-device-scale-factor=0.75")

download_dir = os.getcwd()

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

chrome = webdriver.Chrome(options=options)
chrome.get(url)

wait = WebDriverWait(chrome, 300)
wait.until(EC.url_changes(url))

chrome.get(url_tabela)

elemento_filtro = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[2]/div/table/thead/tr/th[7]/input')
elemento_filtro.send_keys('GMP')
time.sleep(5)

elemento_download = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[1]/a[1]')
elemento_download.click()

download_complete = False
start_time = time.time()

while not download_complete and (time.time() - start_time) < timeout:
    if is_download_complete(download_dir, file1):
        download_complete = True
    else:
        time.sleep(1) 

time.sleep(3)

elemento_filtro.clear()
elemento_filtro.send_keys('Radiante')
time.sleep(5)

elemento_download.click()

download_complete = False
start_time = time.time()

while not download_complete and (time.time() - start_time) < timeout:
    if is_download_complete(download_dir, file2):
        download_complete = True
    else:
        time.sleep(1) 

time.sleep(3)
elemento_filtro.clear()
time.sleep(3)

chrome.get(url_tabela_2)

elemento_filtro = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[2]/div/table/thead/tr/th[8]/input')
elemento_filtro.send_keys('GMP')
time.sleep(10)

elemento_download = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[1]/a[1]')
elemento_download.click()

download_complete = False
start_time = time.time()

while not download_complete and (time.time() - start_time) < timeout:
    if is_download_complete(download_dir, file3):
        download_complete = True
    else:
        time.sleep(1) 

time.sleep(3)

chrome.get(url_tabela)

time.sleep(2)

elemento_filtro = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[2]/div/table/thead/tr/th[7]/input')
elemento_filtro.clear()

elemento_filtro = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[1]/div/button[2]')
elemento_filtro.click()
time.sleep(1)
elemento_filtro = chrome.find_element(By.XPATH, '/html/body/div[8]/div/ul/li[32]/a')
elemento_filtro.click()
time.sleep(5)

elemento_filtro = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[2]/div/table/thead/tr/th[1]/input')
elemento_filtro.send_keys('VDS')
time.sleep(3)

elemento_filtro = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[2]/div/table/thead/tr/th[6]/span[3]')
elemento_filtro.click()
time.sleep(3)

elemento_download = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[1]/a[1]')
elemento_download.click()

download_complete = False
start_time = time.time()

while not download_complete and (time.time() - start_time) < timeout:
    if is_download_complete(download_dir, file4):
        download_complete = True
    else:
        time.sleep(1) 

time.sleep(3)

chrome.quit()

df1 = pd.read_excel(file1)

df2 = pd.read_excel(file2)

df_combined = pd.concat([df1, df2])

df_combined.to_excel(arquivo_renomeado, index=False, engine='openpyxl')

os.remove(file1)
os.remove(file2)
os.rename(file3, 'litas_atividades.xls')
os.rename(file4, 'litas_vds.xls')