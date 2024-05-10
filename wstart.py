from pywinauto.application import Application
import pywinauto
from selenium.webdriver import Chrome
from subprocess import Popen
import time, os, sys

class WookStart:
    def __init__(self):
        self.taskbar_dlg = None
        self.hanaro_password = 's321321321'
        self.start_process()

    def start_process(self):
        print('Wook start-up process started.\n')

        # # self.start_pointnix()
        # # self.start_explorer()
        self.start_ADT()
        # self.start_pycharm()
        # self.start_excel()
        # # self.start_millie()
        # # self.start_kiwoom()
        # # self.start_efriend()
        # self.start_hanaro()
        # self.start_chrome()

        print('Start-up whole process done successfully!\n')

    def wait_for_app(self, window_title, app_name):
        taskbar_app = Application(backend='uia')
        taskbar_app.connect(title='작업 표시줄')
        self.taskbar_dlg = taskbar_app.window(title='작업 표시줄')

        dlg = self.taskbar_dlg.window(title_re=window_title)

        for count in range(1, 20):
            if dlg.exists():
                print('{} is running.'.format(app_name))
                return dlg
            else:
                print('{} is not running. Waiting...trial({})'.format(app_name, count))
                time.sleep(3)

    def start_hanaro(self):
        print('Hanaro start-up procedure begins')

        app = Application('uia')
        app.start('C:/OSSTEM/HANAROOK/hanaro.exe')
        hanaro_dlg = app.window(title_re='^하나로 OK.*')
        time.sleep(1)
        hanaro_login_dlg = hanaro_dlg.window(title='로그인')

        for count in range(10):
            if hanaro_login_dlg.exists():
                break
            time.sleep(2)
            hanaro_login_dlg = hanaro_dlg.window(title='로그인')

        if not hanaro_login_dlg.exists():
            print('Login failed : Maybe hanaro server is not turned on')
            return

        hanaro_password_dlg = hanaro_login_dlg.window(title='비밀 번호:', auto_id='1001', control_type="Edit")
        hanaro_password_dlg.type_keys(self.hanaro_password)
        pywinauto.keyboard.send_keys('{ENTER}')
        time.sleep(0.5)

        hanaro_today_dlg = hanaro_dlg.window(title_re='하나로 투데이')
        if hanaro_today_dlg.exists():
            hanaro_today_checkbox = hanaro_today_dlg.window(title='오늘은 열지 않음')
            hanaro_today_checkbox.click_input()

    def start_pointnix(self):
        print('CDX-ViewADT start-up procedure begins')
        # app = Application('uia')
        # app.start('C:/PointNix/CDX-View/cdx_pointnix.exe')
        Popen('C:/PointNix/CDX-View/cdx_pointnix.exe')

    def start_ADT(self):
        print('ADT start-up procedure begins')
        app = Application('uia')
        app.start("C:/Program Files/ADT EYE3/EYE3.exe")

        # wait ADT window ready
        for count in range(1, 20):
            try:
                app.connect(title_re='ADT-EYE 3.0', control_type='Window')
                break
            except:
                print('ADT window is not ready... trial({})'.format(count))
                time.sleep(3)

        adt_dlg = app.window(title_re='ADT-EYE 3.0')
        adt_dlg.double_click_input(coords=(110, 250))
        adt_dlg.click_input(coords=(150, 270), button_up=False)
        adt_dlg.click_input(coords=(400, 275), button_down=False)

    def start_excel(self):
        print('Excel start-up procedure begins')
        # taskbar_excel_dlg = self.taskbar_dlg.window(title_re='Excel .*개의 실행 중인 창')
        # if taskbar_excel_dlg.exists():
        #     print('Excel is already running')
        #     return

        try:
            app = Application('uia')
            # app.start('C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXE "d:/개원/경영/0 차오름치과 경영.xlsx"')
            # Popen(('C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXE', "d:/개원/경영/차오름치과 경영.xlsx"))
            Popen(('C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE', "d:/개원/경영/차오름치과 경영.xlsx"))
            # time.sleep(12)
            # app.connect(title_re='.*Excel')
            # excel_dlg = app.window(title_re='.*Excel')
            # excel_dlg.maximize()
            #
            # if excel_dlg.exists():
            #     print('it exists')
            # else:
            #     print('it does not exists')
        except Exception as e:
            print('Excel error', e)

    def start_explorer(self):
        print('Explore start-up procedure begins')
        # app = Application('uia')
        # app.start('explorer d:\\')
        Popen(('explorer', 'd:\\'))

    def start_pycharm(self):
        # "C:\Program Files\JetBrains\PyCharm Community Edition 2020.1.2\bin\pycharm64.exe"
        app = Application('uia')
        # app.start('"C:\Program Files\JetBrains\PyCharm Community Edition 2020.1.2//bin\pycharm64.exe"')
        # app.start('"C:\Program Files\JetBrains\PyCharm Community Edition 2021.3.2//bin\pycharm64.exe"')
        app.start('C:\\Program Files\\JetBrains\\PyCharm Community Edition 2022.1.1\\bin\\pycharm64.exe')

        # taskbar_pycharm_dlg = self.taskbar_dlg.window(title_re='^PyCharm Community Edition.*')
        # taskbar_pycharm_dlg = self.taskbar_dlg.window(title_re='^PyCharm Community Edition.*')
        # taskbar_pycharm_dlg.click_input()

    def start_chrome(self):
        print('Chrome start-up procedure begins')
        app = Application('uia')
        # app.start('C:/Program Files (x86)/Google/Chrome/Application/chrome.exe')
        app.start('C:/Program Files/Google/Chrome/Application/chrome.exe')

        chrome_dlg = app.top_window()
        # chrome_dlg.maximize()

        chrome_search_dlg = chrome_dlg.window(title='주소창 및 검색창', control_type='Edit')
        chrome_search_dlg.type_keys('http://naver.com/')
        pywinauto.keyboard.send_keys('{ENTER}')
        chrome_mail_dlg = chrome_dlg.window(title='메일')
        chrome_mail_dlg.click_input()

        # try:
        #     chrome_newtag_dlg = chrome_dlg.window(title='새 탭', control_type='Button')
        #     chrome_newtag_dlg.click_input()
        #     chrome_search_dlg.type_keys('https://kebhana.com')
        #     pywinauto.keyboard.send_keys('{ENTER}')
        #     # # chrome_inquiry_dlg = chrome_dlg.window(title='조회')
        #     # # chrome_inquiry_dlg.click_input()
        #     # # chrome_total_inquiry_dlg = chrome_dlg.window(title='전체계좌조회')
        #     # # chrome_total_inquiry_dlg.click_input()
        #
        #     chrome_newtag_dlg.click_input()
        #     chrome_search_dlg.type_keys('https://remotedesktop.google.com')
        #     pywinauto.keyboard.send_keys('{ENTER}')
        #
        #     chrome_newtag_dlg.click_input()
        #     chrome_search_dlg.type_keys('https://remotedesktop.google.com')
        #     pywinauto.keyboard.send_keys('{ENTER}')

        # except Exception as e:
        #     print('Something is wrong when open KEB bank', e)

    def start_kiwoom(self):
        Popen('C:/KiwoomHero4/Bin/NKStarter.exe')

    def start_millie(self):
        Popen('C:/Program Files/millie/millie.exe')

    def start_efriend(self):
        Popen('C:/eFriend Expert/efriendexpert/efriendexpert.exe')

if __name__ == '__main__':
    ws = WookStart()