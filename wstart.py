from pywinauto.application import Application
from pywinauto import findwindows
import pywinauto, pyautogui
from selenium.webdriver import Chrome
from subprocess import Popen
import time, os, sys

class WookStart:
    def __init__(self):
        self.hanaro_password = 's321321321'
        self.fail_message = ''

        self.start_process()

    def start_process(self):
        print('Wook start-up process started.\n')

        # processes = findwindows.find_elements()
        # for process in processes:
        #     print(f'{process} : {process.process_id}')

        # self.start_ADT()
        # self.start_explorer()
        # self.start_hanaro()
        # self.start_pointnix()
        # self.start_pycharm()
        # self.start_excel()
        # self.start_millie()
        # self.start_chrome()
        # self.start_kakaotalk()
        self.start_line()

        # self.start_kiwoom()
        # self.start_efriend()

        print('\nWhole start-up process done successfully!' + self.fail_message)

    def report_fail(self, message):
        if self.fail_message:
            self.fail_message += '\n' + message
        else:
            self.fail_message = '\n====== Except ======\n' + message

    def get_dlg(self, app, window_title):
        maximum_trial = 20
        waiting_time = 1

        for count in range(1, maximum_trial):
            try:
                app.connect(title_re=window_title)
                print('{} app is now connected.'.format(window_title))
                break
            except:
                print('{} app is not ready. Waiting {}s...trial({})'.format(window_title, waiting_time, count))
                time.sleep(waiting_time)

        dlg = app.window(title_re=window_title)

        for count in range(1, maximum_trial):
            if dlg.exists():
                print('{} dlg is now running.'.format(window_title))
                return dlg
            else:
                print('{} dlg is not running. Waiting {}s...trial({})'.format(window_title, waiting_time, count))
                time.sleep(waiting_time)

        print('Getting {} dlg failed!!!'.format(window_title))
        self.report_fail(window_title)

        return None

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
        pyautogui.hotkey('ctrl', 'win', 'left')
        pyautogui.hotkey('ctrl', 'win', 'left')

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

        pyautogui.hotkey('ctrl', 'win', 'right')

    def start_excel(self):
        print('Excel start-up procedure begins')

        app = Application('uia')

        # Popen(('C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE', "d:/개원/경영/차오름치과 경영.xlsx"))
        app.start('C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE "d:/개원/경영/차오름치과 경영.xlsx"')
        excel_dlg = app.window(title_re='.*Excel')
        excel_dlg.maximize()

    def start_explorer(self):
        print('Explore start-up procedure begins')
        # app = Application('uia')
        # app.start('explorer d:\\')
        Popen(('explorer', 'd:\\'))

    def start_pycharm(self):
        print('PyCharm start-up procedure begins')
        app = Application('uia')
        app.start('C:\\Program Files\\JetBrains\\PyCharm Community Edition 2023.3.5\\bin\\pycharm64.exe')

    def start_chrome(self):
        print('Chrome start-up procedure begins')
        app = Application('uia')
        app.start('C:/Program Files/Google/Chrome/Application/chrome.exe')
        # app.connect(title_re='새 탭', control_type='Window')
        chrome_dlg = app.top_window()
        search_dlg = chrome_dlg.window(title_re='작업 영역')
        search_dlg.click_input()

        # chrome_dlg.print_control_identifiers()

        # chrome_dlg = app.top_window()
        # # chrome_dlg.maximize()
        #
        # chrome_search_dlg = chrome_dlg.window(title='주소창 및 검색창', control_type='Edit')
        # chrome_search_dlg.type_keys('http://naver.com/')
        # pywinauto.keyboard.send_keys('{ENTER}')
        # chrome_mail_dlg = chrome_dlg.window(title='메일')
        # chrome_mail_dlg.click_input()

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

    def start_kakaotalk(self):
        print('KakaoTalk start-up procedure begins.')
        app = Application('uia')
        app.start('C:/Program Files (x86)/Kakao/KakaoTalk/KakaoTalk.exe')
        kakaotalk_dlg = self.get_dlg(app, '카카오톡')
        if not kakaotalk_dlg: return

        password_dlg = kakaotalk_dlg['Edit2']
        if password_dlg.exists():
            password_dlg.type_keys('^a''{DELETE}''{ENTER}')
        else:
            self.report_fail('KakaoTalk')

    def start_line(self):
        print('Line start-up procedure begins.')
        app = Application('uia')
        app.start('C:/Users/neuroris/AppData/Local/LINE/bin/LineLauncher.exe')
        line_dlg = self.get_dlg(app, 'LINE')
        if not line_dlg: return

        password_dlg = line_dlg['Edit2']
        if password_dlg.exists():
            password_dlg.type_keys('^a''{DELETE}''{ENTER}')
        else:
            self.report_fail('LINE')

    def start_millie(self):
        print('millie start-up procedure begins')
        # Popen('C:/Program Files/millie/millie.exe')

        app = Application('uia')
        app.start('C:/Program Files/millie/millie.exe')

        millie_dlg = app.top_window()
        study_dlg = millie_dlg.window(title_re='내 서재')
        time.sleep(4)
        study_dlg.set_focus()
        study_dlg.click_input()

    def start_kiwoom(self):
        Popen('C:/KiwoomHero4/Bin/NKStarter.exe')

    def start_efriend(self):
        Popen('C:/eFriend Expert/efriendexpert/efriendexpert.exe')

if __name__ == '__main__':
    ws = WookStart()