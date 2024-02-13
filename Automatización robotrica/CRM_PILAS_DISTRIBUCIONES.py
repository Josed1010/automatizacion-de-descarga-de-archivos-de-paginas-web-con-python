# -*- coding: latin-1 -*-
from datetime import datetime
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

op = webdriver.ChromeOptions()
#op.binary_location = "ruta del ejecutable de Chrome"
op.add_argument("--headless")
op.add_argument("--disable-gpu")
op.add_experimental_option('excludeSwitches', ['enable-logging'])
#driver = webdriver.Chrome(chrome_options=op)
#driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver = webdriver.Chrome(executable_path='C:/Users/ROBOTRPASI/Desktop/chromedriver.exe')

driver.maximize_window()


#Funcion para el descargue de archivo
def executeBot():
        input_email_id = "ROBOTRPASI"
        input_pwd = "Cancun8*"
        # Get the page. 
        driver.get('https://crm.corbeta.com.co/distribuciones_pilas/')
        time.sleep(5)
        print("SAC ESTA ABIERTO...")
        
        # We find the xpath for insert our email address.
        email = driver.find_element(By.XPATH,"//*[@id='usuario']")
        # Put mail.
        email.send_keys(input_email_id)
        print("Email enviado.")
        # We find the xpath for insert our password.
        time.sleep(2)
        # We find xpath of button Login.     
        password = driver.find_element(By.XPATH,"/html/body/center/form/table/tbody/tr[3]/td[2]/input")
        # Put password.
        password.send_keys(input_pwd)
        print("Password enviado.")
        # We find xpath of button Login.
        button = driver.find_element(By.XPATH,"//*[@id='ingresa']")
        # We click on button of login.
        button.click()
        # Open session.
        print("SAC esta abierto.")
        x = 0
        # View Issues
        time.sleep(5)
        
        driver.get('https://crm.corbeta.com.co/distribuciones_pilas//revisa.php?modulo=16&tablase=0&nomlibro=Lista_usuarios&sqlprocesa=x%25DAM%258E%25C1%250A%25C20%2510D%257FeoQX%2502%257E%2540ORAP%250B%25D6%257BI%25EAFV%25DA%25AC%2524Y%25BF_ks%25F08%25F3%2586%2599%25E9%25DBS%25BB%25BF%25C1%2528%25F7A%25B3%25BA%25C4%2582%259A%2529%25A1%25DA%2528%25B3O%25844%253B%259E%25D0%253D%2528%2596%25E1G%257C%2525%25E02%25BC%2528%2585%252F%25E5%25B0Q%25EB%25C6%25C2oivh%25FA%25A3As%25E9%25CCv%2589%25AC6%251C%25AE%25DD%2519%25EA%2502%2528L%2514%250A%253C%2585c%25AD%2500%250F%2512A%25ED%25F2cu%251A%25FF%2527%253E%2508%2527%253AR')
        # Download report
        time.sleep(5)
                
        #Mover Archivo
        shutil.move("C:/Users/ROBOTRPASI/Downloads/Lista_usuarios.xls", "C:/Users/ROBOTRPASI/Pictures/Carpeta de Planos/TMP_CRMPILASDISTRIBUCIONES_USUARIOS.xls")
        
        
        driver.close()
        #proc.terminate()
executeBot()

print('Terminado')