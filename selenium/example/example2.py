from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from selenium.webdriver.support.select import Select
import time
import pyautogui


curTime = datetime.today()
year = curTime.strftime("%Y")
month = curTime.strftime("%m")
day = curTime.strftime("%d")

moveList = ["양준혁","최동호","이태희"]


try:
    
    chrome = webdriver.Chrome()
    chrome.maximize_window()
    chrome.get("http://gw.nice.co.kr/")
    time.sleep(1)
    
    # ID 작성
    chrome.find_element(By.XPATH, "/html/body/form/div/div/div/p[1]/input").send_keys("thlee1")
    time.sleep(1)
    
    select = Select(chrome.find_element(By.XPATH, "/html/body/form/div/div/div/p[1]/span/select"))
    
    select.select_by_visible_text('nicepni.co.kr')
    time.sleep(1)
    
    chrome.find_element(By.XPATH, "/html/body/form/div/div/div/p[2]/input").send_keys("1q2w3e4r!!")
    time.sleep(1)
    
    chrome.find_element(By.XPATH, "/html/body/form/div/div/div/div[3]/button").click()
    
    time.sleep(5)
    
    #화면 전환
    chrome.switch_to.window(chrome.window_handles[0])
    
    time.sleep(1)
    for mainIndex in moveList:
        chrome.find_element(By.XPATH, "/html/body/div/div[1]/div[2]/div/ul/li[3]/a").click()
        
        time.sleep(1)
        
        iframeSideFirstChk = chrome.find_element(By.XPATH, "/html/body/div/div[2]/div/div[2]/iframe")
        chrome.switch_to.frame(iframeSideFirstChk) # frame으로 변경
            
        iframeSideSecondChk = chrome.find_element(By.XPATH, "/html/body/div/div[1]/iframe")
        chrome.switch_to.frame(iframeSideSecondChk) # frame으로 변경
        
        
        #메일분류함 클릭
        chrome.find_element(By.XPATH, "/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[1]/span").click()
        time.sleep(1)
        
        
        # frame Default 진행
        chrome.switch_to.default_content()
        
        iframeFirstChk = chrome.find_element(By.XPATH, "/html/body/div/div[2]/div/div[2]/iframe")
        chrome.switch_to.frame(iframeFirstChk) # frame으로 변경
            
        iframeSecondChk = chrome.find_element(By.XPATH, "/html/body/div/div[3]/div[3]/div/iframe")
        chrome.switch_to.frame(iframeSecondChk) # frame으로 변경
        
        #읽지 않음 ->more 클릭
        chrome.find_element(By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/table[3]/tbody/tr[1]/td/div/form/table[1]/tbody/tr/td[2]/span[1]/span").click()
        
        chrome.switch_to.default_content()
        
        iframeFirstChk = chrome.find_element(By.XPATH, "/html/body/div/div[2]/div/div[2]/iframe")
        chrome.switch_to.frame(iframeFirstChk) # frame으로 변경
            
        iframeSecondChk = chrome.find_element(By.XPATH, "/html/body/div/div[3]/div[3]/div/iframe")
        chrome.switch_to.frame(iframeSecondChk) # frame으로 변경    
        
        # table의 row 개수 읽어오기
        rows = chrome.find_elements(By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/ok/form/table[2]/tbody/tr[1]/td/table/tbody/tr")
                
        #안읽은 메일 개수    
        tempTotalCount = chrome.find_element(By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/ok/form/table[2]/tbody/tr[1]/td/table/tbody/tr["+str(len(rows))+"]/td/table/tbody/tr/td[3]/font[2]").text
        
        if int(tempTotalCount) > 15:
            # 목록개수 select 가져오기
            select2 = Select(chrome.find_element(By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/ok/form/table[2]/tbody/tr[1]/td/table/tbody/tr[12]/td/table/tbody/tr/td[1]/select"))
            
            # 목록개수 100개로 변경
            select2.select_by_visible_text('100')
            time.sleep(3)                    
        
        #안읽은 메일 개수 만큼 반복실행하여 체크박스 클릭하기     
        for index in range(2,int(tempTotalCount)+2):
            mailTo = chrome.find_element(By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/ok/form/table[2]/tbody/tr[1]/td/table/tbody/tr["+str(index)+"]/td[6]/div").text
            if mailTo == mainIndex:
                time.sleep(1)
                chrome.find_element(By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/ok/form/table[2]/tbody/tr[1]/td/table/tbody/tr["+str(index)+"]/td[2]").click()
                
        chrome.switch_to.default_content()
        
        iframeFirstChk = chrome.find_element(By.XPATH, "/html/body/div/div[2]/div/div[2]/iframe")
        
        chrome.switch_to.frame(iframeFirstChk) # frame으로 변경    
        
        #이동 아이콘 클릭!
        chrome.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/table/tbody/tr/td[2]/table/tbody/tr/td[10]/div[1]").click()
        
        #생성한 폴더명의 개수 가져오기
        totalMailBox = chrome.find_elements(By.XPATH, "/html/body/div/div[3]/div[2]/table/tbody/tr/td[2]/table/tbody/tr/td[10]/span/ok/table/tbody/tr[4]/td[3]/table[1]/tbody/tr[1]/td/div/table")

        #생성한 폴더명의 개수를 이용하여 이동할 폴더명 위치를 알아내기 
        for index2 in range(2,len(totalMailBox)):
            tempText = chrome.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/table/tbody/tr/td[2]/table/tbody/tr/td[10]/span/ok/table/tbody/tr[4]/td[3]/table[1]/tbody/tr[1]/td/div/table["+str(index2)+"]/tbody/tr/td/table/tbody/tr/td[3]/a/font").text
            
            if tempText == mainIndex:
                chrome.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/table/tbody/tr/td[2]/table/tbody/tr/td[10]/span/ok/table/tbody/tr[4]/td[3]/table[1]/tbody/tr[1]/td/div/table["+str(index2)+"]/tbody/tr/td/table/tbody/tr/td[3]/a/font").click()
                break
        
        # 이동 버튼 클릭!
        chrome.find_element(By.XPATH,"/html/body/div/div[3]/div[2]/table/tbody/tr/td[2]/table/tbody/tr/td[10]/span/ok/table/tbody/tr[4]/td[3]/table[2]/tbody/tr/td/span[2]").click()        
        
        time.sleep(2)
        
        # 모달 창을 따로 컨트롤하는게 어려워서 ENTER로 변경
        pyautogui.press('ENTER')
        
        time.sleep(3)
        
        chrome.switch_to.default_content()
    
    
except Exception as ex:
    time.sleep(30)
    print(ex)
    