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
import openpyxl
import pandas as pd
import pyexcel as pe
import urllib.request
import shutil
import time
import os
import subprocess
import pyautogui as py
from openpyxl import workbook
from openpyxl import load_workbook
from dotenv import load_dotenv
import warnings
warnings.simplefilter("ignore")

op = webdriver.ChromeOptions()
#op.binary_location = "ruta del ejecutable de Chrome"
op.add_argument("--headless")
op.add_argument("--disable-gpu")
op.add_experimental_option('excludeSwitches', ['enable-logging'])
#driver = webdriver.Chrome(chrome_options=op)
#driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver = webdriver.Chrome(executable_path=r'C:/Users/ROBOTRPASI/Desktop/chromedriver.exe')

driver.maximize_window()

load_dotenv()
input_email_id = os.getenv('NAME1')
input_pwd = os.getenv('PASSWORD1')

def executeBot():
        

        # Get the page. 
        driver.get('https://multicrm.colcomercio.com.co/alkosto_vozcliente/')
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
       
        password = driver.find_element(By.XPATH,"/html/body/div[1]/div/form/div[2]/input")
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
        
        button = driver.find_element(By.XPATH,"//*[@id='menu-toggle']")
        button.click()
        time.sleep(5)
        
                
        
        driver.get("https://multicrm.colcomercio.com.co/alkosto_vozcliente/revisa.php?rand=6533&modulo=14")
        # Download report
        time.sleep(5)
        
        button = driver.find_element(By.XPATH,"//*[@id='cod_reporte']")
        button.click()
        time.sleep(5)
        
        
        button = driver.find_element(By.XPATH,"//*[@id='iddatospersonalizados']/div/div/div[1]/span/div/button")
        button.click()
        time.sleep(5)
        
        
        button = driver.find_element(By.XPATH,"//*[@id='iddatospersonalizados']/div/div/div[1]/span/div/ul/li[9]/a/label/input")
        button.click()
        time.sleep(5)
        
        
        
        button = driver.find_element(By.XPATH,"//*[@id='iddatospersonalizados']/div/div/div[2]/span/div/button/span")
        button.click()
        time.sleep(5)
        
        
        
        button = driver.find_element(By.XPATH,"//*[@id='iddatospersonalizados']/div/div/div[2]/span/div/ul/li[2]/a/label")
        button.click()
        time.sleep(5)
        
        
        
        button = driver.find_element(By.XPATH,"//*[@id='generar']")
        button.click()
        time.sleep(15)
        
        
        button = driver.find_element(By.XPATH,"//*[@id='table_usuarios_wrapper']/div[1]/div[2]/div/button")
        button.click()
        time.sleep(5)
        
        
        shutil.move("C:/Users/ROBOTRPASI/Downloads/Reporte Usuarios por Aplicacion.xlsx", "C:/Users/ROBOTRPASI/Pictures/Planos CRM/CRM_FOTON_PREVENTA.xlsx")
        time.sleep(8)
        driver.close()
        #proc.terminate()
executeBot()


print('Terminado')