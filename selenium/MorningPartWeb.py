from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from openpyxl import load_workbook
import time


dailyUrl = "http://222.111.237.40:11110/main"
jobNameList  = [] # a4 ~ a13
siteList = ["https://data.koscom.co.kr/kor/main.do", 
            "https://www.nicebir.co.kr/", 
            "https://www.niceport.co.kr/", 
            "http://www.nicepni.com/main",
            "http://www.nicecni.com/"] # a17 ~ 21
chkElementList = ["/html/body/div[3]/div[1]/h1/a/img",
                  "/html/frameset",
                  "/html/frameset",
                  "/html/body/div/header/hgroup/div/a",
                  "/html/frameset"]
chkCompleteOpenWebPageCount = 0



def dailyCheckAction(dailyFile,dailySaveFile):
    try:
        #엑셀 파일 로드 하기
        wb = load_workbook(dailyFile)
        workSheet = wb["Sheet1"]
        
        chrome = webdriver.Chrome()
        chrome.maximize_window()
        chrome.get(dailyUrl)
        
        dailyNameAndResult = {}
        
        for index in range(1,9):
            jobName = chrome.find_element(By.XPATH, "/html/body/div/div/section/section[1]/section[1]/article/div/table/tbody/tr["+str(index)+"]/th/span").text
            dailyNameAndResult[jobName] = chrome.find_element(By.XPATH,"/html/body/div/div/section/section[1]/section[1]/article/div/table/tbody/tr["+str(index)+"]/td[1]/span").text
            
        time.sleep(1)
        
        # 오전 데일리 자동화 파일에 확인여부 값 넣기
        # CUT OFF 백업 / 홈페이지 확인 / 배치 확인 등등
        for kndex in range(4,12):
            if dailyNameAndResult[workSheet["A"+str(kndex)].value] == '정상':
                workSheet["M"+str(kndex)] = 'O' # 값 입력시 '' 으로 사용
        
        if dailyNameAndResult['홈페이지 구동 확인'] == '-':
            for index2 in range(0,len(siteList)):
                chrome.execute_script('window.open("'+siteList[index2]+'");')
                tabs = chrome.window_handles
                chrome.switch_to.window(tabs[index2+1])
                chrome.get(siteList[index2])
                time.sleep(1)
                if chrome.find_element(By.XPATH,chkElementList[index2]).is_displayed():
                    global chkCompleteOpenWebPageCount
                    chkCompleteOpenWebPageCount += 1
                
        if chkCompleteOpenWebPageCount == 5:
            workSheet["M10"] = 'O' # 값 입력시 '' 으로 사용
        # C&I 페이지 로그인 페이지 내에 frame 확인    
        frameMainChk = chrome.find_element(By.XPATH, "/html/frameset/frame[2]")
        
        time.sleep(1)
        chrome.switch_to.frame(frameMainChk) # frame으로 변경
        
        # ID에 GUEST 글자 제거 (백스페이스 5번 누름)
        chrome.find_element(By.XPATH, "/html/body/div[1]/div/div/section/div[2]/form/fieldset/input[1]").send_keys(5*Keys.BACK_SPACE)
        time.sleep(1) 
        # PW에 글자 제거 (백스페이스 4번 누름)
        chrome.find_element(By.XPATH, "/html/body/div[1]/div/div/section/div[2]/form/fieldset/input[2]").send_keys(4*Keys.BACK_SPACE)
        time.sleep(1)
        
        # ID 작성
        chrome.find_element(By.XPATH, "/html/body/div[1]/div/div/section/div[2]/form/fieldset/input[1]").send_keys("cniadmin")
        time.sleep(1)
        # PW 작성
        chrome.find_element(By.XPATH, "/html/body/div[1]/div/div/section/div[2]/form/fieldset/input[2]").send_keys("nice1111")
        time.sleep(1)
        
        # 로그인 버튼 클릭
        chrome.find_element(By.XPATH, "/html/body/div[1]/div/div/section/div[2]/form/fieldset/button").click()
        time.sleep(1)
        
        # frame Default 진행
        chrome.switch_to.default_content()
        
        # C&I 위에 헤더부분 frame 확인
        frameChk = chrome.find_element(By.XPATH, "/html/frameset/frame[2]")
        # C&I 위에 헤더부분 frame 변경
        chrome.switch_to.frame(frameChk)
        
        # C&I 유통시장 클릭 
        chrome.find_element(By.XPATH, "/html/body/div[1]/header/nav[2]/dl/dd[3]/a").click()
        time.sleep(1)
        # C&I 투자자별 매매현황 클릭
        chrome.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[3]/nav[3]/dl/dd[3]/ul/li/a").click()
        time.sleep(1)
        
        
        # frame Default 진행
        chrome.switch_to.default_content()
        
        # C&I 페이지 투자자별 매매동향 내부 frame 확인
        frameTradingChk = chrome.find_element(By.XPATH, "/html/frameset/frame[2]")
        
        # C&I 투자자별 매매동향 내부 frame 변경
        chrome.switch_to.frame(frameTradingChk)
        
        # C&I 투자자별 매매동향 글자 가져오기(페이지 정확하게 나왔는지 확인용)
        tradingText1 = chrome.find_element(By.XPATH, "/html/body/div[1]/div[1]/section[1]/section/h3/span").text
        
        
        if tradingText1 == "투자자별 매매동향" : 
            workSheet["M13"] = 'O'        
            
        # frame Default 진행    
        chrome.switch_to.default_content()
        
        
        frameChk = chrome.find_element(By.XPATH, "/html/frameset/frame[2]")
        
        chrome.switch_to.frame(frameChk)
        
        
        chrome.find_element(By.XPATH, "/html/body/div[1]/header/nav[2]/dl/dd[4]/a").click()
        
        time.sleep(1)
        chrome.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[3]/nav[4]/dl/dd[1]/ul/li[5]/a").click()
        
        chrome.switch_to.default_content()
        
        frameCarryEarningRateChk = chrome.find_element(By.XPATH, "/html/frameset/frame[2]")
        
        chrome.switch_to.frame(frameCarryEarningRateChk)
        
        tradingText2 = chrome.find_element(By.XPATH, "/html/body/div[1]/div[1]/section[1]/section/h3/span").text
        
        if tradingText2 == "캐리수익률" :
            workSheet["M14"] = 'O'
            
        chrome.switch_to.default_content()
        
        frameEdfChk = chrome.find_element(By.XPATH, "/html/frameset/frame[2]")
        
        chrome.switch_to.frame(frameEdfChk)
        
        chrome.find_element(By.XPATH, "/html/body/div[1]/header/nav[2]/dl/dd[7]/a").click()
        
        time.sleep(1)
        chrome.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[3]/nav[7]/dl/dd[3]/ul/li[3]/a").click()
        
        if chrome.find_element(By.XPATH, "/html/body/div[1]/div[1]/section[1]/section/h3/span").is_displayed() == True:
            workSheet["M15"] = 'O'
        
        
        time.sleep(5)        
                    
        wb.save(dailySaveFile)
        
    except Exception as ex:
        print(ex)