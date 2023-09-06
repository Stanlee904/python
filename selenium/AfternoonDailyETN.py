from openpyxl import Workbook
from datetime import datetime
from openpyxl import load_workbook
import win32com.client



curTime = datetime.today()
year = curTime.strftime("%Y")
month = curTime.strftime("%m")
day = curTime.strftime("%d")

etnCount = 0

def dailyETNCheck(dailySaveFile):
    try:
        #엑셀 파일 로드 하기
        wbDaily = load_workbook(dailySaveFile)
        workSheetDaily = wbDaily["Sheet1"]    
        
        #Outlook Application에 대한 객체 생성하기
        outlook = win32com.client.Dispatch("Outlook.Application")

        #MAPI만 지원함.
        rxOutlook = outlook.GetNamespace("MAPI")

        #ETN 메일을 받고 있는 rmsend폴더 access
        inbox = rxOutlook.GetDefaultFolder(6).Folders("rmsend")

        messages = inbox.Items
        
        for message in messages:
            if (str(message.ReceivedTime)[0:10] == year+"-"+month+"-"+day) and "[ETN]" in message.Subject:
                etnCount += 1
                
        if etnCount == 3:
            workSheetDaily['M5'] = 'O'
        
        wbDaily.save(dailySaveFile)
        
    except Exception as ex:
        print(ex)