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




def dailyFifthCheck(dailyUrl,dailySaveFile):
    try:

        
        #엑셀 파일 로드 하기
        wbDaily = load_workbook(dailySaveFile)
        workSheetDaily = wbDaily["Sheet1"]        
        
        chrome = webdriver.Chrome()
        chrome.maximize_window()
        chrome.get(dailyUrl)
        
        #현대증권 BPMS 실시간 데이터 전송(45번 실시간 데이터 전송)
        isRealDataSend = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[34]/td[1]/span").text
        
        #Forward Rate 생성 확인
        isForwardRateChk = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[35]/td[1]/span").text
        
        #현대증권 부도율 처리
        isHyundaiDefaultRate = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[31]/td[1]/span").text
        
        
        if isRealDataSend == "정상":
            workSheetDaily['M30'] = 'O'
            
        if isForwardRateChk == "정상":
            workSheetDaily['M31'] = 'O'
                    
        if isHyundaiDefaultRate == "정상":
            workSheetDaily['M35'] = 'O'
        
        wbDaily.save(dailySaveFile)
        
    except Exception as ex:
        print("dailyFifthCheck 오류 발생!")
        print(ex)