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
import urllib.request
import shutil
import time
import os
import subprocess
import pyautogui as py
from dotenv import load_dotenv
from datetime import datetime

op = webdriver.ChromeOptions()
op.add_argument("--headless")
op.add_argument("--disable-gpu")
op.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(executable_path='C:/Users/ROBOTRPASI/Desktop/chromedriver.exe')

driver.maximize_window()



def executeBot():
        input_email_id = "robotrpasi@colcomercio.com.co"
        input_pwd = "Cancun8*"
        # Get the page. 
        driver.get('https://colcomerciocom.sharepoint.com/sites/GESTIONDEACCESOS/Documentos%20compartidos/Forms/AllItems.aspx?id=%2Fsites%2FGESTIONDEACCESOS%2FDocumentos%20compartidos%2FReporte%20de%20Ausentismos&p=true&ct=1651583888926&or=Teams%2DHL&ga=1')

        #driver.execute_script("document.body.style.zoom='90%'")
        time.sleep(5)
        print("SAC ESTA ABIERTO...")
        
        
        email = driver.find_element(By.XPATH,"//*[@id='i0116']")
        email.send_keys(input_email_id)
        time.sleep(5) 
        button = driver.find_element(By.XPATH,"//*[@id='idSIButton9']")
        # We click on button of login.
        button.click()
        time.sleep(4)
        
         
        password = driver.find_element(By.XPATH,"//*[@id='i0118']")
        # Put password.
        password.send_keys(input_pwd)
        print("Password enviado.")
        # We find xpath of button Login.
        button = driver.find_element(By.XPATH,"//*[@id='idSIButton9']")
        # We click on button of login.
        button.click()
        # Open session.
        time.sleep(5)
        print("SAC esta abierto.")
        x = 0
        
        
        
        button = driver.find_element(By.XPATH,"//*[@id='KmsiCheckboxField']").click()
        # We click on button of login.
        time.sleep(5)
        
        
        button = driver.find_element(By.XPATH,"//*[@id='idBtn_Back']")
        button.click()
        time.sleep(7)
        
        
        
        button = driver.find_element(By.XPATH,"//*[@id='header55-dateModifiedColumn_707']")
        button.click()
        time.sleep(6)
        
        
        
        button = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/div/div/div/div/ul/li[2]/button")
        button.click()
        time.sleep(5)
        
        
        button = driver.find_element(By.XPATH,"//*[@id='appRoot']/div[1]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[1]/div[1]")
        button.click()
        time.sleep(8)
        
       
        
        #//*[@id="appRoot"]/div[1]/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div/div[1]/div[6]/button/span/span
        button = driver.find_element(By.XPATH,"//*[@id='appRoot']/div[1]/div[2]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div/div[1]/div[6]/button/span/span")
        button.click()
        time.sleep(25)
       
        driver.close()
        #proc.terminate()
executeBot()
print('Terminado plano ausentismos en descargas')