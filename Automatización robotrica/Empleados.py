# -*- coding: latin-1 -*
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
from pathlib import Path
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
#input_email_id = os.getenv('C:/Users/ROBOTRPASI/Pictures/.env/.env/NAME1')
#input_pwd = os.getenv('C:/Users/ROBOTRPASI/Pictures/.env/.env/PASSWORD1')
input_email_id = os.getenv('NAME1')
input_pwd = os.getenv('PASSWORD1')
def executeBot():
        
        
        # Get the page. 
        driver.get('https://aks4uc4ph.accounts.ondemand.com/saml2/idp/sso?sp=https://www.successfactors.com/colombiana&idp=https://aks4uc4ph.accounts.ondemand.com')
        time.sleep(5)
        print("SAC ESTA ABIERTO...")
        
        # We find the xpath for insert our email address.
        email = driver.find_element(By.XPATH,"//*[@id='j_username']")
        # Put mail.
        email.send_keys(input_email_id)
        print("Email enviado.")
        # We find the xpath for insert our password.
        time.sleep(2)     
        password = driver.find_element(By.XPATH,"//*[@id='j_password']")
        # Put password.
        password.send_keys(input_pwd)
        print("Password enviado.")
        # We find xpath of button Login.
        button = driver.find_element(By.XPATH,"//*[@id='logOnFormSubmit']/div")
        # We click on button of login.
        button.click()
        # Open session.
        print("SAC esta abierto.")
        x = 0
        # View Issues
        time.sleep(15)
        
        time.sleep(5)
        
        #Boton informes
        driver.get('https://hcm19.sapsf.com/xi/ui/reportcenter/pages/reportCenter.xhtml?bplte_company=colombiana&_s.crb=jOfHnKYFd6hz1ZznksmqTLHnfFfLSZ56btUWXNDkSHM%253d')
        time.sleep(10)
        
        #Boton informes
        button = driver.find_element(By.XPATH,"//*[@id='__link3']")
        button.click()
        #driver.get('https://hcm19.sapsf.com/xi/ui/reportcenter/pages/reportCenter.xhtml?bplte_company=colombiana&_s.crb=jOfHnKYFd6hz1ZznksmqTLHnfFfLSZ56btUWXNDkSHM%253d#')
        time.sleep(15)
        
        
        #Boton exportar
        button = driver.find_element(By.XPATH,"//*[@id='84:_radiolabel']")
        button.click()
        time.sleep(5)
        
        '''
        #Boton excel
        #button = driver.find_element(By.XPATH,"//*[@id='79:']")
        #button.click()
        
        #Boton excel
        button = driver.find_element(By.XPATH,"//*[@id='79:']/option[2]")
        button.click()
        time.sleep(5)
        
        #Boton generar
        button = driver.find_element(By.XPATH,"//*[@id='dlgButton_126:']")
        button.click()
        time.sleep(35)
        
        #Boton generar
        button = driver.find_element(By.XPATH,"//*[@id='134:_link']")
        button.click()
        time.sleep(55)
        '''
               
        #Mover Archivo
        #shutil.move(archivo, "C:/Users/ROBOTRPASI/Pictures/Carpeta de Planos/TMP_MANTISAKT_USUARIOS.xlsx")
        
        driver.close()

executeBot()

