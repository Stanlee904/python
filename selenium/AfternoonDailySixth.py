from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from openpyxl import load_workbook
import time



curTime = datetime.today()
year = curTime.strftime("%Y")
month = curTime.strftime("%m")
day = curTime.strftime("%d")




def dailySixthCheck(dailyUrl,dailySaveFile):
    try:
                
        #엑셀 파일 로드 하기
        wbDaily = load_workbook(dailySaveFile)
        workSheetDaily = wbDaily["Sheet1"]        
        
        chrome = webdriver.Chrome()
        chrome.maximize_window()
        chrome.get(dailyUrl)
        
        #금투협 Matrix, CDCP Matrix 데이터 전송 확인
        isKofiaMatrixAndCDCPMatrix = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[42]/td[1]/span").text
        
        # 3 / 33번 서버 동작 확인
        isServerCheck3 = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[43]/td[1]/span").text
        isServerCheck33 = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[44]/td[1]/span").text
        
        
        #국민연금 보유종목수신 (SW_01 / FX_02)
        isNpsHoldingsReceive = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[30]/td[1]/span").text
                
        #BEFORE 6 확인
        isBefore6Chk = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[45]/td[1]/span").text
        
        
        if isKofiaMatrixAndCDCPMatrix == "정상":
            workSheetDaily['M32'] = 'O'
            
        if isServerCheck3 == "정상" and isServerCheck33 == "정상":
            workSheetDaily['M33'] = 'O'
                    
        if isNpsHoldingsReceive == "정상":
            workSheetDaily['M34'] = 'O'
            
        if isBefore6Chk == "정상":
            workSheetDaily['M36'] = 'O'
        
        wbDaily.save(dailySaveFile)
        
    except Exception as ex:
        print("dailySixthCheck 오류 발생!")
        print(ex)