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



def dailySecondCheck(dailyUrl,dailySaveFile):
    try:
        
        futureCompleteCount = 0
        
        #엑셀 파일 로드 하기
        wbDaily = load_workbook(dailySaveFile)
        workSheetDaily = wbDaily["Sheet1"]        
        
        chrome = webdriver.Chrome()
        chrome.maximize_window()
        chrome.get(dailyUrl)
        
        #NCO 컷오프 오전전송 확인
        isCutOffMorningSend = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[3]/td[1]/span").text
        
        #전일가격테이블 백업 / 삭제
        isPriceBackUpDelete = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[4]/td[1]/span").text
        
        #국민연금 외화채권 전송 확인
        isNpsForeignCurBondSend = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[2]/td[1]/span").text
        
        
        #선물수신프로그램 동작 확인
        for index in range(1,16):
            isFuturesOpencheck =  chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr["+str(index)+"]/td[6]/span").text
            if isFuturesOpencheck == "정상":
                futureCompleteCount +=1
                break
        
        if isCutOffMorningSend == "정상":
            workSheetDaily['M12'] = 'O'
            
        if isPriceBackUpDelete == "정상":
            workSheetDaily['M13'] = 'O'            
            
        if isNpsForeignCurBondSend == "정상":
            workSheetDaily['M14'] = 'O'            
        
        if futureCompleteCount == 1:
            workSheetDaily['M15'] = 'O'
        
        wbDaily.save(dailySaveFile)
        
    except Exception as ex:
        print("dailySecondCheck 오류 발생!")
        print(ex)