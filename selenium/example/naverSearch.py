from selenium import webdriver
from selenium.webdriver.common.by import By
import time

def naverSearch():
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get("http://www.naver.com")

    time.sleep(1)

    elem = driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/form/fieldset/div/input")

    elem.click()

    elem.send_keys("나이스피앤아이")

    elemSearch = driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/form/fieldset/button/span[1]")

    elemSearch.click()

    time.sleep(1)
    
    
    
try:
    naverSearch()
except Exception as ex:
    print(ex)    