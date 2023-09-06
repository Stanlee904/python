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

def dailyFirstCheck(dailyUrl,dailyFile,dailySaveFile):
    try:
        
        #엑셀 파일 로드 하기
        wbDaily = load_workbook(dailyFile)
        workSheetDaily = wbDaily["Sheet1"]        
        
        chrome = webdriver.Chrome()
        chrome.maximize_window()
        chrome.get(dailyUrl)
        
        #국민연금 보유종목 수신(14:05)
        isNpsHoldingsReceive = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[1]/td[1]/span").text
        
        #실시간 발행정보 결과 확인
        isRealTimeIssueInform = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[2]/td[1]/span").text
        
        #(월초 작업) 정기예금
        isFixedDepositAndMMDA = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[2]/td[1]/span").text
        
        #(월말 작업) 로이터 지수
        isReuterCheck = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[46]/td[1]/span").text
        
        #전이행렬, LifetimePD, 국민연금
        isTransitionMatrix = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[48]/td[1]/span").text
        isLifeTimePD = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[49]/td[1]/span").text
        isLastDayNPS = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[50]/td[1]/span").text
        
        
        
        if isNpsHoldingsReceive == "정상":
            workSheetDaily['M4'] = 'O'
            
        if isRealTimeIssueInform == "정상":
            workSheetDaily['M6'] = 'O'            
            
        if isFixedDepositAndMMDA == "정상":
            workSheetDaily['M7'] = 'O'            
        
        if isReuterCheck == "정상":
            workSheetDaily['M8'] = 'O'

        if isTransitionMatrix == "정상" and isLifeTimePD == "정상" and isLastDayNPS == "정상":
            workSheetDaily['M8'] = 'O'
            workSheetDaily['M9'] = 'O'
            workSheetDaily['M10'] = 'O'
        elif isTransitionMatrix == "정상" and isLifeTimePD == "정상" and isLastDayNPS == "미입력":
            workSheetDaily['M8'] = 'O'
            workSheetDaily['M9'] = 'O'
        elif isTransitionMatrix == "미입력" and isLifeTimePD == "정상" and isLastDayNPS == "정상":
            workSheetDaily['M9'] = 'O'
            workSheetDaily['M10'] = 'O'
        
        wbDaily.save(dailySaveFile)
        
    except Exception as ex:
        print("dailyFirstCheck 오류 발생!")
        print(ex)