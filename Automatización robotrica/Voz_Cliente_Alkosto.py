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
import pyexcel as pe
import urllib.request
import shutil
import time
import os
import subprocess
import pyautogui as py
from openpyxl import workbook
from openpyxl import load_workbook
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



def executeBot():
        input_email_id = "1010148970"
        input_pwd = "Corbeta23.C"
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
        
        
        
        
        
        
        driver.get("https://multicrm.colcomercio.com.co/alkosto_vozcliente//revisa.php?modulo=16&tablase=0&nomlibro=Lista_usuarios&sqlprocesa=x%25DAmP%25D1%250A%25830%250C%257C%25DFW%25E4%25AD%251B%2514a%251F%25E0%25C3%2518%250A%250E%25D7%25C2%25F4%255DZ%25CD%25B6%2582%251A%25A9%25D6%25EF%251F%25B2%25C2%25AC%25F3%25DEr%2597%255Cr%2529%2592%253C%25B9%2596p%2580%2515jj%252A7%253Ae%250D%25F1%2540p%2523%25DA%250D%2513%25F5%25D4i%258B%2521%258B%259D2mH%25A9%2517%25F6S%25B5g%25B0l%251B%25D0%253E%25B7%2513%25DA%255B%25C3%25A5%2580%253D%253DK%258F.R%25F5df%258A%25CF%259C%2515%2519%25E3LHvZ%25FA%25BF4%25FFK5X%259A%2511%251B%25B2U%2583%25E3%25A0%25EA7%25FD%2592%25A7%250Fy%2507%251F%251B%251C%2584%252F%25C9%2593%25B4%2584%259B%25CC%2584%253F%2505t%25A8%252F%2590%2522H%2513%25EBU%2501%251F%25E0%25F3Y-")
        # Download report
        time.sleep(5)
        
        
        
        
        
        shutil.move("C:/Users/ROBOTRPASI/Downloads/Lista_usuarios.xls", "C:/Users/ROBOTRPASI/Pictures/Carpeta de Planos/TMP_VOZCLIENTEALKOSTO_USUARIOS.xls")
        
        driver.close()
        #proc.terminate()
executeBot()
time.sleep(5)
