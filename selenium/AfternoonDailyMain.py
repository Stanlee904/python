from datetime import datetime
from openpyxl import load_workbook
import os
import AfternoonDailyETN
import AfternoonDailyFirst
import AfternoonDailySecond
import AfternoonDailyThird
import AfternoonDailyFourth
import AfternoonDailyFifth
import AfternoonDailySixth

try:
    curTime = datetime.today()
    year = curTime.strftime("%Y")
    month = curTime.strftime("%m")
    day = curTime.strftime("%d")
    # hour = curTime.strftime("%H")
    # minute = curTime.strftime("%M")


    dailyFile = "D:\새 폴더\오후데일리_자동화.xlsx"
    dailySaveFile = "D:\새 폴더\오후데일리_자동화_"+year+month+day+".xlsx"
    dailyUrl = "http://222.111.237.40:11110/main"
    


    def fileCheck():
        try:
            if os.path.isfile(dailySaveFile):
                os.remove(dailySaveFile)
                                        
        except Exception as ex:
            print(ex)
            
    fileCheck()
    AfternoonDailyFirst.dailyFirstCheck(dailyUrl,dailyFile,dailySaveFile)
    #outlook 실거래 데이터 수신확인 Check
    AfternoonDailyETN.dailyETNCheck(dailySaveFile)
    AfternoonDailySecond.dailySecondCheck(dailyUrl,dailySaveFile)
    AfternoonDailyThird.dailyThirdCheck(dailyUrl,dailySaveFile)
    AfternoonDailyFourth.dailyFourthCheck(dailyUrl,dailySaveFile)
    AfternoonDailyFifth.dailyFifthCheck(dailyUrl,dailySaveFile)
    AfternoonDailySixth.dailySixthCheck(dailyUrl,dailySaveFile)
    
    
    
    
            
except Exception as ex:
    print(ex)
    
    





