from pywinauto import application

import time
import pyautogui

# 현재 윈도우 화면에 있는 프로세스 목록 리스트를 반환한다. 
# 리스트의 각 요소는 element 객체로 프로세스 id, 핸들값, 이름 등의 정보를 보유한다.  
# procs = findwindows.find_elements()

# for proc in procs:
#      print(f"{proc}  / 프로세스: {proc.process_id}")


# 위에서 찾은 프로세스의 정보나 프로그램 경로를 기재하면 간단히 연결된다. 테스트할 프로그램을 연결하는 방법은 크게 두가지가 있다.

# Application().start() : 프로그램 경로를 넣어서 실행
# Application().connect() : 이미 실행되고 있는 프로그램을 연결
# 둘중 원하는 방법으로 연결 해주면 된다.

# 이때  Application() 함수의 인수로 backend 값을 지정해주어야 하는데, 값은 win32 와 uia 가 있다.

# 이 요소값은 어떤 종류의 프로그램을 제어하려는지 알려주는 것인데, 프로그램 개발 시 사용된 GUI 프레임워크에 따라 다르다.

# Application(backend="win32") : 메모장과 같은 old한 프로그램을 실행할때
# Application(backend="uia") : 최신 기술이 사용된 프로그램 (왠만한 요즘 프로그램들이 속함)


# app = application.Application(backend='uia').start("D:\\igate\\inzent - 운영\\iTools Standard.exe")
application.Application(backend='uia').start("C:\\Users\\user\\AppData\\Roaming\\NICE P&I\\NICE P&I\\NICE피앤아이.exe")
# app = application.Application(backend='uia').connect(process='23280')
#procname = "iTools Standard.exe"

# apploaded = False

#프로세스의 경로를 넣어 실행해준다. 
#app.start("D:\\igate\\inzent - 운영\\iTools Standard.exe")
# app.start("C:\\Program Files\\Notepad++\\notepad++.exe")
#app.connect(process="7708")



# 컨트롤 요소 출력
#dlg_spect = app['iTools Standard.exe'] # 변수에 노트패드 윈도우 어플리케이션 객체를 할당
# dlg_spect = app.window(title='ITools Standard')
#dlg_spect.print_control_identifiers()
#dlg.print_control_identifiers() # 노트패드의 컨트롤 요소를 트리로 모두 출력


time.sleep(3)

img_cap = pyautogui.locateOnScreen("D:\\workspace\\python\\selenium\\IMG\\NiceIntegrated1.png")
pyautogui.moveTo(img_cap)
pyautogui.click(img_cap)