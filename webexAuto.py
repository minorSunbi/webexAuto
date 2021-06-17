# pip install pywinauto
# pip install apscheduler
from pywinauto.application import Application
from apscheduler.schedulers.background import BackgroundScheduler
import win32gui
import win32com.client
import time

# 처음에 웹엑스 실행
app = Application(backend="uia").start("C:\Program Files (x86)\Webex\Webex\Applications\ptoneclk.exe")
is_first = True

def register_class(meet_number):
    global is_first

    print("called")
    app = Application(backend="uia").connect(path="ptoneclk.exe")
    #print(app.Pane.print_control_identifiers())
    meet_box = app.Pane.Edit

    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(app.top_window().handle)

    # button 없을 때(처음 입력 시)만 클릭
    if is_first:
        meet_box.click_input()
        is_first = False
    
    meet_box.type_keys(meet_number, with_spaces=True)
    meet_box.Button.click()


# 이와 같은 형식으로 수업들을 작성
# def class*():
#     register_class("미팅룸 번호")
def class1():
    register_class("333 333 333")

def class2():
    register_class("000 000 000")

####
scheduler = BackgroundScheduler()
scheduler.start()
####

# 아래에 job들(수업들) 추가
# id 다르게 설정
# 각 job별로 매개변수 설정
# month="3-6" 설정 시 3월부터 6월까지만 작동, "9-12" 설정 시  9월부터 12월까지만 작동
# start_date="2021-09-01", end_date="2021-11-25" 설정 시 9월 1일부터 11월 25일까지만 작동 
# day_of_week="mon, tue, thu" 설정 시 월요일, 화요일 ,목요일마다 작동
# hour="15" 설정 시 오후 3시마다 작동
# minute="55-59/1" 설정 시 55분부터 59분까지 1분마다 작동

# 예시: 2021.03.01부터 2021.06.15까지 월, 화, 목요일 16시 수업인 class1을 위해 15시 55분부터 59분까지 1분마다 작동
scheduler.add_job(class1, "cron", id="class1", start_date="2021-03-01", end_date="2021-06-15",\
                    day_of_week="mon, tue, thu", hour="15", minute="55-59/1")

# 예시2: 2021.09.01부터 2021.11.25까지 금요일 10시 수업이 class2를 위해 9시 55분에 작동
scheduler.add_job(class2, "cron", id="class2", start_date="2021-09-01", end_date="2021-11-25",\
                    day_of_week="fri", hour="9", minute="55")

####
while True:
    time.sleep(60000)
####