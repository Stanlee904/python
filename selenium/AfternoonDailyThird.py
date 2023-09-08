from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from openpyxl import load_workbook
from pywinauto import application
from pywinauto import findwindows
import time



curTime = datetime.today()
year = curTime.strftime("%Y")
month = curTime.strftime("%m")
day = curTime.strftime("%d")



def dailyThirdCheck(dailyUrl,dailySaveFile):
    try:
        
        irsRateCount = 0
        
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
        
        #IRS RATE
        for index in range(19,24):
            irsCorrect = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr["+str(index)+"]/td[1]/span").text
            if irsCorrect == "정상":
                irsRateCount += 1
            
        
        #세금계산        
        # app = application.Application(backend='uia').start("C:\\Users\\user\\AppData\\Roaming\\NICE P&I\\NICE P&I\\NICE피앤아이.exe")
        # dlg = app['NICE피앤아이 V 2.81']
        
        # dlg.child_window(title="통합시스템", auto_id="8", control_type="Button").click_input()
        
        # time.sleep(5)
        
        # procs = findwindows.find_elements()

        # for proc in procs:
        #     tempProc = f"{proc}"        
        #     if '통합 System' in tempProc:
        #         tempProcessId = proc.process_id
            
        # app2 = application.Application(backend='uia').connect(process=tempProcessId)
        
        # dlg2 = app2['Dialog']
        
        # dlg2.child_window(title="기타", control_type="MenuItem").select()
                
        # dlg2['세금계산MenuItem2'].select()
        
        # dlg2['생성Button'].click()
        
                
        if isMMFRate == "정상":
            workSheetDaily['M17'] = 'O'
            
        if isKorBankStandardRate == "정상":
            workSheetDaily['M18'] = 'O'            
            
        if isCDCPSend == "정상":
            workSheetDaily['M19'] = 'O'            
        
        if isKosNiceCitiSend == "정상":
            workSheetDaily['M20'] = 'O'
            
        if irsRateCount == 5:
            workSheetDaily['M21'] = 'O'
        
        wbDaily.save(dailySaveFile)
        
    except Exception as ex:
        print("dailyThirdCheck 오류 발생!")
        print(ex)