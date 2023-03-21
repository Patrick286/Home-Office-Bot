import pandas as pd
import pyautogui
import time
import requests
import openpyxl

from PIL import ImageGrab
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

nome_do_arquivo = "teste.xlsx"
url_do_regrid = "https://app.regrid.com/us#b=admin&base=satellite"
url_do_appraisal = "https://www.google.com.br/maps/@-12.9452397,-38.4534359,3669m/data=!3m1!1e3"
url_do_maps = "https://www.google.com.br/maps/@-12.9452397,-38.4534359,3669m/data=!3m1!1e3?hl=pt_BR"
df = pd.read_excel(nome_do_arquivo)

for index, row in df.iterrows():
    print("Index: " + str(index) + "O parcel da casa é " + row["Parcel Number"])
    chrome = webdriver.Chrome(executable_path= 'chromedriver.exe')

    chrome.set_window_size(1024, 600)
    chrome.maximize_window()
    chrome.get(url_do_regrid)

    time.sleep(5)

    login1 = chrome.find_element(By.XPATH, '//*[@id="signup-box"]/div/div/div/div/div/div[2]/a')
    login1.click()

    time.sleep(1)

    login2 = chrome.find_element(By.XPATH, '//*[@id="signup-box"]/div/div/div/div/div/div[3]/div/div[2]/div/a')
    login2.click()

    time.sleep(1)

    email = chrome.find_element(By.XPATH, '//*[@id="map_signin_email"]')
    email.send_keys(row["Email"])

    senha = chrome.find_element(By.XPATH, '//*[@id="map_signin_password"]')
    senha.send_keys(row["Senha"])

    logar = chrome.find_element(By.XPATH, '//*[@id="signInCard-signIn"]/div[2]/input')
    logar.click()

    time.sleep(2)

    sat1 = chrome.find_element(By.XPATH, '//*[@id="mapbar"]/div[3]/div/a/i')
    sat1.click()

    time.sleep(2)

    sat2 = chrome.find_element(By.XPATH, '//*[@id="satellite-layer"]/a/i')
    sat2.click()

    elemento_texto_parcel = chrome.find_element(By.XPATH, "//*[@id='nav-search-query']")
    elemento_texto_parcel.send_keys(row["Parcel Number"])

    time.sleep(2)

    enter = chrome.find_element(By.XPATH, "//*[@id='nav-search']/span/div/div/div/div[1]/div[1]/a")
    enter.click()

    time.sleep(4)

    region = (565, 206, 1599, 860)
    img = ImageGrab.grab(region)
    img.save(r'C:\Users\Patrick\Pictures\Saved Pictures\screen.png')

    time.sleep(2)

    link_regrid1 = chrome.find_element(By.XPATH, "//*[@id='property']/div[2]/div/ul/li[1]/div/a")
    link_regrid1.click()

    time.sleep(2)

    link_regrid1 = chrome.find_element(By.XPATH, "//*[@id='property']/div[2]/div/ul/li[1]/div/ul/li[4]/a")
    link_regrid1.click()
    link = link_regrid1.click()

    time.sleep(2)

    chrome.quit

