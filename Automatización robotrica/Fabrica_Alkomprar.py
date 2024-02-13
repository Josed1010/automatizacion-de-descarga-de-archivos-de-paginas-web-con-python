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
        driver.get('https://fabricadecreditos.colcomercio.com.co/alkomprar/index.php')
        time.sleep(5)
        print("SAC ESTA ABIERTO...")
        
        # We find the xpath for insert our email address.
        email = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/div[2]/form/fieldset/div[2]/div/div[1]/div/input")
        # Put mail.
        email.send_keys(input_email_id)
        print("Email enviado.")
        # We find the xpath for insert our password.
        time.sleep(2)
        # We find xpath of button Login.
           
        password = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/div[2]/form/fieldset/div[2]/div/div[2]/div/input")
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
        time.sleep(2)
        
        driver.get("https://fabricadecreditos.colcomercio.com.co/alkomprar/revisa.php?rand=7088195579945480539&modulo=15")
        # Download report
        time.sleep(5)
        
        button = driver.find_element(By.XPATH,"//*[@id='idprincipalcentro']/div/div[1]/div/a[1]").click()
        time.sleep(5)
        
        shutil.move("C:/Users/ROBOTRPASI/Downloads/Lista_usuarios.xlsx", "C:/Users/ROBOTRPASI/Pictures/Carpeta de Planos/TMP_FABRICALKOMPRAR_USUARIOS.xlsx")
        
        driver.close()
        #proc.terminate()
executeBot()

archivo_excel = 'C:/Users/ROBOTRPASI/Pictures/Carpeta de Planos/TMP_FABRICALKOMPRAR_USUARIOS.xlsx'

df = pd.read_excel(archivo_excel)

archivo_txt = 'S:/PLANOS/TMP_FABRICALKOMPRAR_USUARIOS.txt'

df.to_csv(archivo_txt, sep='\t', index=False)

print('Terminado')