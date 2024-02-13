# -*- coding: latin-1 -*-
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from time import sleep
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import NoSuchElementException
from os import remove
import pandas as pd
import urllib.request
import shutil
import time
import os
import subprocess
import pyautogui as py
from dotenv import load_dotenv

op = webdriver.ChromeOptions()
#op.binary_location = "ruta del ejecutable de Chrome"
op.add_argument("--headless")
op.add_argument("--disable-gpu")
op.add_experimental_option('excludeSwitches', ['enable-logging'])
#driver = webdriver.Chrome(chrome_options=op)
#driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver = webdriver.Chrome(executable_path='C:/Users/ROBOTRPASI/Desktop/chromedriver.exe')

driver.maximize_window()

load_dotenv()
input_email_id = os.getenv('NAME2')
input_pwd = os.getenv('PASSWORD2')

def executeBot():

        # Get the page. 
        driver.get('http://10.181.3.106/consultas/web/')
        time.sleep(5)
        print("SAC ESTA ABIERTO...")
        email = driver.find_element(By.XPATH,"//*[@id='root']/div/div/div[1]/div/div[2]/form/div[1]/div/input")
        email.send_keys(input_email_id)
        print("Email enviado.")
        time.sleep(5)
        
        password = driver.find_element(By.XPATH,"//*[@id='root']/div/div/div[1]/div/div[2]/form/div[2]/div/input")
        password.send_keys(input_pwd)
        print("Password enviado.")
        time.sleep(5)
        
        button = driver.find_element(By.XPATH,"//*[@id='root']/div/div/div[1]/div/div[2]/form/button")
        button.click()
        time.sleep(5)
        
        print("SAC esta abierto.")
       
        button = driver.find_element(By.XPATH,"//*[@id='pestanahome']")
        button.click()
        time.sleep(5)
        
        button = driver.find_element(By.XPATH,"//*[@id='pestanaadminUser']/div/div[1]/div[2]/div/button[1]")
        button.click()
        time.sleep(5)
        print("descarga.")
        
        shutil.move("C:/Users/ROBOTRPASI/Downloads/UsuariosConsultas.xlsx", "C:/Users/ROBOTRPASI/Pictures/Carpeta de Planos/TMP_MODULOCONSULTAS_USUARIOS.xlsx")
        
        driver.close()
        #proc.terminate()
executeBot()

archivo_excel = 'C:/Users/ROBOTRPASI/Pictures/Carpeta de Planos/TMP_MODULOCONSULTAS_USUARIOS.xlsx'

df = pd.read_excel(archivo_excel)

archivo_txt = 'S:/PLANOS/TMP_MODULOCONSULTAS_USUARIOS.txt'

df.to_csv(archivo_txt, sep='\t', index=False)

print('Terminado')

