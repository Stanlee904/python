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



def dailyThirdCheck(dailyUrl,dailySaveFile):
    try:
        
        #엑셀 파일 로드 하기
        wbDaily = load_workbook(dailySaveFile)
        workSheetDaily = wbDaily["Sheet1"]        
        
        chrome = webdriver.Chrome()
        chrome.maximize_window()
        chrome.get(dailyUrl)
        
        #MMF 금리 입력
        isMMFRate = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[7]/td[1]/span").text
        
        #한국은행 기준금리 입력
        isKorBankStandardRate = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[6]/td[1]/span").text
        
        #CDCP 전송프로그램 확인
        isCDCPSend = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[29]/td[1]/span").text
        
        #KOSCOM - NICE - CITI은행 전송
        isKosNiceCitiSend = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[26]/td[1]/span").text
        
        
        
        #세금계산
        
                
        if isMMFRate == "정상":
            workSheetDaily['M17'] = 'O'
            
        if isKorBankStandardRate == "정상":
            workSheetDaily['M18'] = 'O'            
            
        if isCDCPSend == "정상":
            workSheetDaily['M19'] = 'O'            
        
        if isKosNiceCitiSend == "정상":
            workSheetDaily['M20'] = 'O'
        
        wbDaily.save(dailySaveFile)
        
    except Exception as ex:
        print("dailyThirdCheck 오류 발생!")
        print(ex)