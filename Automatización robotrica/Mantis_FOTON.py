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
input_email_id = os.getenv('NAME1')
input_pwd = os.getenv('PASSWORD1')

def executeBot():

        # Get the page. 
        driver.get('https://mantis.colcomercio.com.co:81/foton/login_page.php')
        time.sleep(5)
        print("SAC ESTA ABIERTO...")
        
        # We find the xpath for insert our email address.
        email = driver.find_element(By.XPATH,"//input[@name='username']")
        # Put mail.
        email.send_keys(input_email_id)
        print("Email enviado.")
        # We find the xpath for insert our password.
        time.sleep(2)
        # We find xpath of button Login.
        button = driver.find_element(By.XPATH,"//*[@id='login-form']/fieldset/input[2]")
        button.click()        
        password = driver.find_element(By.XPATH,"//input[@name='password']")
        # Put password.
        password.send_keys(input_pwd)
        print("Password enviado.")
        # We find xpath of button Login.
        button = driver.find_element(By.XPATH,"//input[@type='submit']")
        # We click on button of login.
        button.click()
        # Open session.
        print("SAC esta abierto.")
        x = 0
        # View Issues
        time.sleep(5)
        driver.get('https://mantis.colcomercio.com.co:81/foton/accessadmin_reportes_acceso.php')
        # Download report
        time.sleep(5)
        button = driver.find_element(By.XPATH,"//*[@id='seccion1']/div/div/div/table/tbody/tr[4]/td/input").click()
        time.sleep(5)
        
        now=datetime.now()

        year=now.strftime("%Y")       
        print('año: ', year)
        
        mes= now.strftime("%m")
        print('mes: ', mes)
        
        dia= now.strftime("%d")
        print('dia: ', dia)
        
        fecha=now.strftime("%Y_%m_%d")
        print('fecha', fecha)
        
        archivo=("C:/Users/ROBOTRPASI/Downloads/Reporte_general_usuarios_" + fecha + ".xlsx")
        
        print("Reporte_general_usuarios_" + fecha + ".xlsx")
        
        #Mover Archivo
        shutil.move(archivo, "C:/Users/ROBOTRPASI/Pictures/Carpeta de Planos/TMP_MANTISCORAUTOS_USUARIOS.xlsx")
        
        driver.close()
        #proc.terminate()
executeBot()

archivo_excel = 'C:/Users/ROBOTRPASI/Pictures/Carpeta de Planos/TMP_MANTISCORAUTOS_USUARIOS.xlsx'

df = pd.read_excel(archivo_excel)

archivo_txt = 'S:/PLANOS/TMP_MANTISCORAUTOS_USUARIOS.txt'

df.to_csv(archivo_txt, sep='\t', index=False)


print('Terminado')