from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime
import time


url = "https://search.naver.com/search.naver?where=news&sm=tab_jum&query=나이스피앤아이"
newsTitleList = []
curTime = datetime.today()
year = curTime.strftime("%Y")
month = curTime.strftime("%m")
day = curTime.strftime("%d")

excelName = "D:\새 폴더\dd_" + year + month + day + ".xlsx"


print(excelName)


try:
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(url)
    
    
    
    for index in range(1,11):
        newsTitle = driver.find_element(By.XPATH,"/html/body/div[3]/div[2]/div/div[1]/section/div/div[2]/ul/li["+str(index)+"]/div[1]/div/a").text
        newsTitleList.append(newsTitle)
        
    
    for index2 in range(0,len(newsTitleList)):
        print(newsTitleList[index2] + "\n")
    
    time.sleep(5)
    
    driver.quit()
    
    workBook = Workbook()
    
    workSheet = workBook.active
    
    for index3 in range(0,len(newsTitleList)):
        workSheet["A"+str(index3+1)] = newsTitleList[index3]
        
        
    
    workBook.save(excelName)
    
    
    
    


    
except Exception as ex:
    print(ex)
    