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

futureCheckList = ["종목정보", "종목마감", "정산가격", "현물정보결제기준채권","30 종가단일가","M2 당일 확정"]
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
        
        # #Daily 선물 수신 확인
        for index in range(1,16):
            for futureIndex in futureCheckList:
                # 병합 되어 있으면 값을 못 읽음. 
                # chromeFuture = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr["+str(index)+"]/th/span").text
                if chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr["+str(index)+"]/td[1]/span").is_displayed() == True:
                    chromeFuture = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr["+str(index)+"]/td[1]/span").text
                    if futureIndex in chromeFuture and "현물정보" in futureIndex:
                        gbFutureData = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr[15]/td[2]/span").text
                        print(gbFutureData + "현물정보 +++++ ")
                    else:
                        gbFutureData = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[2]/article/div/table/tbody/tr["+str(index)+"]/td[4]/span").text
                        print(gbFutureData)
                            
                gbFutureDataList.append(gbFutureData)        
        
        
        for index2 in range(0,len(futureCheckList)):
            dailyFutureTotal += futureCheckList[index2] + " : " + gbFutureDataList[index2] + " / "
        
        
        #산업은행 보유종목 수신 확인
        # isKosNiceCitiSend = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr[26]/td[1]/span").text
        for index3 in range(36,41):
            isKdbHoldingsCheck = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[2]/section[1]/article/div/table/tbody/tr["+str(index3)+"]/td[1]/span").text
            if isKdbHoldingsCheck == "정상":
                kdbHoldingsCount += 1
                    
        #기업신용정보 생성
        
        
        
        
        #해외지수 입력
        

        
        if isEachBondSendOpenCheck == "정상":
            workSheetDaily['M24'] = 'O'
            
        if isKsdReceive == "정상":
            workSheetDaily['M25'] = 'O'
            
        workSheetDaily['E26'] = dailyFutureTotal
        
        if kdbHoldingsCount == 5:
            workSheetDaily['M29'] = 'O'
        
        wbDaily.save(dailySaveFile)
        
    except Exception as ex:
        print("dailyFourthCheck 오류 발생!")
        print(ex)