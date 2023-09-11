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
hour = curTime.strftime("%H")
minute = curTime.strftime("%M")

futureCheckList = ["종목정보", "30 종가단일가","종목마감", "M2 당일 확정","정산가격", "현물정보결제기준채권",]
gbFutureDataList = []

def dailyFourthCheck(dailyUrl,dailySaveFile):
    try:
        
        dailyFutureTotal = ""
        kdbHoldingsCount = 0
        
        #엑셀 파일 로드 하기
        wbDaily = load_workbook(dailySaveFile)
        workSheetDaily = wbDaily["Sheet1"]        
        
        chrome = webdriver.Chrome()
        chrome.maximize_window()
        chrome.get(dailyUrl)
        
        # #6,13,14,24번 서버별 확인
        isEachBondSendOpenCheck = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[32]/td[1]/span").text
        
        # #예탁원 데이터 수신
        isKsdReceive = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[33]/td[1]/span").text
        
        # #Daily 선물 수신 확인 (이 부분 다시 로직 짜기 -> 20230907)
        for index in range(1,16):
            for futureIndex in futureCheckList:
                # 병합 되어 있으면 값을 못 읽음. 
                # chromeFuture = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr["+str(index)+"]/th/span").text
                if chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr["+str(index)+"]/td[1]/span").is_displayed() == True:
                    chromeFuture = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr["+str(index)+"]/td[1]/span").text
                    if "현물정보" in chromeFuture and "기준채권" in chromeFuture:
                        gbFutureData = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr[15]/td[2]/span").text
                        gbFutureDataList.append(gbFutureData)
                        break
                    elif futureIndex in chromeFuture:
                        gbFutureData = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr["+str(index)+"]/td[4]/span").text
                        gbFutureDataList.append(gbFutureData)
                            
        for index2 in range(0,len(futureCheckList)):
            if index2 == len(futureCheckList)-1:
                dailyFutureTotal += futureCheckList[index2] + " : " + gbFutureDataList[index2]
            else:
                dailyFutureTotal += futureCheckList[index2] + " : " + gbFutureDataList[index2] + " / "
        
        
        #산업은행 보유종목 수신 확인
        # isKosNiceCitiSend = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[26]/td[1]/span").text
        for index3 in range(36,41):
            isKdbHoldingsCheck = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr["+str(index3)+"]/td[1]/span").text
            if isKdbHoldingsCheck == "정상":
                kdbHoldingsCount += 1
                    
        #기업신용정보 생성
        
        if int(hour) >= 17 and int(minute) >= 25:
            app = application.Application(backend='uia').start("C:\\Users\\user\\AppData\\Roaming\\NICE P&I\\NICE P&I\\NICE피앤아이.exe")
            dlg = app['NICE피앤아이 V 2.81']
            
            dlg.child_window(title="통합시스템", auto_id="8", control_type="Button").click_input()
            
            time.sleep(5)
            
            procs = findwindows.find_elements()

            for proc in procs:

                tempProc = f"{proc}"        
                if '통합 System' in tempProc:
                    tempProcessId = proc.process_id
                
            app2 = application.Application(backend='uia').connect(process=tempProcessId)
            
            dlg2 = app2['Dialog']
            
            dlg2.child_window(title="기타", control_type="MenuItem").select()
                    
            dlg2['데이터검수후일괄작업MenuItem2'].select()
            
            dlg2['기업별신용정보생성MenuItem'].select()
            
            dlg2['생성Button'].click()
            
            time.sleep(60000)
            
            isCreateEnterPriseCreditInForm = "정상"
            
            app2.kill() # app종료
        else:
            isCreateEnterPriseCreditInForm = "비정상"
        
        #해외지수 입력
        
        if isEachBondSendOpenCheck == "정상":
            workSheetDaily['M24'] = 'O'
        
        if isKsdReceive == "정상":
            workSheetDaily['M25'] = 'O'
        
        workSheetDaily['E26'] = dailyFutureTotal
        
        if isCreateEnterPriseCreditInForm == "정상":
            workSheetDaily['M27'] = 'O'
        
        
        if kdbHoldingsCount == 5:
            workSheetDaily['M29'] = 'O'
        
        wbDaily.save(dailySaveFile)
        
    except Exception as ex:
        print("dailyFourthCheck 오류 발생!")
        print(ex)
        