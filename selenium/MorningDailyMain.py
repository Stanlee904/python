from datetime import datetime
from openpyxl import load_workbook
import os
import MorningPartWeb
import MorningDailyNPSystem
import MorningDailyDetailPart1AM

try:
    curTime = datetime.today()
    year = curTime.strftime("%Y")
    month = curTime.strftime("%m")
    day = curTime.strftime("%d")
    # hour = curTime.strftime("%H")
    # minute = curTime.strftime("%M")


    dailyFile = "D:\새 폴더\오전데일리_자동화.xlsx"
    dailySaveFile = "D:\새 폴더\오전데일리_자동화_"+year+month+day+".xlsx"


    def fileCheck():
        try:
            if os.path.isfile(dailySaveFile):
                os.remove(dailySaveFile)
                                        
        except Exception as ex:
            print(ex)
            
            
    fileCheck()
    MorningPartWeb.dailyCheckAction(dailyFile,dailySaveFile)
    MorningDailyNPSystem.NPS_WorkCompleteCheck(dailySaveFile)
    MorningDailyDetailPart1AM.ChkPart1AM(dailySaveFile)
    
    
except Exception as ex:
    print(ex)
    
    





