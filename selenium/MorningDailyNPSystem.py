from datetime import datetime
from openpyxl import load_workbook
from pywinauto import application
import time
import pyautogui


curTime = datetime.today()
year = curTime.strftime("%Y")
month = curTime.strftime("%m")
day = curTime.strftime("%d")
hour = curTime.strftime("%H")
minute = curTime.strftime("%M")

def NPS_WorkCompleteCheck(dailySaveFile):
    try:
        ksdCompleteCheck = 'D:\\workspace\\python\\selenium\\IMG\\NiceIntegrated2.png'
        dmsCompleteCheck = 'D:\\workspace\\python\\selenium\\IMG\\NiceIntegrated3.png'
        elnCompleteCheck = 'D:\\workspace\\python\\selenium\\IMG\\NiceIntegrated4.png'
        

        app = application.Application(backend='uia').start("C:\\Users\\user\\AppData\\Roaming\\NICE P&I\\NICE P&I\\NICE피앤아이.exe")
        dlg = app['NICE피앤아이 V 2.81']

        dlg.child_window(title="통합시스템", auto_id="8", control_type="Button").click_input()
        
        time.sleep(5)
                
        # 데일리 엑셀 불러오기
        wbDaily = load_workbook(dailySaveFile)        
        #엑셀 시트 접근
        workSheetDaily = wbDaily["Sheet1"]        
        
        # 해당 이미지가 존재 한다면 
        if (pyautogui.locateOnScreen(ksdCompleteCheck) is not None) and workSheetDaily["M12"].value is None :
            workSheetDaily["M12"] = 'O'


        if (pyautogui.locateOnScreen(dmsCompleteCheck) is not None) and workSheetDaily["M12"].value is None  and workSheetDaily["M16"].value is None :
            workSheetDaily["M12"] = 'O'
            workSheetDaily["M16"] = 'O'
            
        elif (pyautogui.locateOnScreen(dmsCompleteCheck) is not None) and workSheetDaily["M12"].value is not None  and workSheetDaily["M16"].value is None :
            workSheetDaily["M16"] = 'O'
                        
        if int(hour) >= 10 and int(minute) >= 00:
            if (pyautogui.locateOnScreen(elnCompleteCheck) is not None) and workSheetDaily["M12"].value is None and workSheetDaily["M16"].value is None and workSheetDaily["M17"].value is None:
                workSheetDaily["M12"] = 'O'
                workSheetDaily["M16"] = 'O'            
                workSheetDaily["M17"] = 'O'
                
            elif (pyautogui.locateOnScreen(elnCompleteCheck) is not None) and workSheetDaily["M12"].value is not None and workSheetDaily["M16"].value is None and workSheetDaily["M17"].value is None:
                workSheetDaily["M16"] = 'O'
                workSheetDaily["M17"] = 'O'
                
            elif (pyautogui.locateOnScreen(elnCompleteCheck) is not None) and workSheetDaily["M12"].value is None and workSheetDaily["M16"].value is not None and workSheetDaily["M17"].value is None:
                workSheetDaily["M12"] = 'O'
                workSheetDaily["M17"] = 'O'            
                
            elif (pyautogui.locateOnScreen(elnCompleteCheck) is not None) and workSheetDaily["M12"].value is not None and workSheetDaily["M16"].value is not None and workSheetDaily["M17"].value is None:
                workSheetDaily["M17"] = 'O'
                            
        wbDaily.save(dailySaveFile)

    except Exception as ex:
        print(ex)