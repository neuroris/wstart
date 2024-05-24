from pywinauto.application import Application
from pywinauto import findwindows
import pywinauto, pyautogui
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import Remote
from subprocess import Popen
import time, os, sys
from wookutil import WookCipher

class WookStart:
    def __init__(self):
        self.failed = False
        self.fail_message = '--------- Failed app list ---------'
        self.wc = WookCipher()
        self.hanaro_password = self.wc.decrypt('c:/data/hanaro_key', 'c:/data/hanaro_encrypt.bin')
        self.message_password = self.wc.decrypt('c:/data/start_key', 'c:/data/start_encrypt.bin')

        self.start_process()

    def start_process(self):
        print('Wook start-up process started.\n')

        self.start_ADT()
        self.start_explorer()
        self.start_hanaro()
        self.start_pointnix()
        self.start_pycharm()
        self.start_excel()
        self.start_millie()
        self.start_kakaotalk()
        self.start_line()
        self.start_chrome()

        # self.start_kiwoom()
        # self.start_efriend()

        if self.failed:
            print('\nSome app failed to launch.\n')
            print(self.fail_message)
        else:
            print('\nWhole start-up process done successfully!')

    def report_failure(self, message, e=None):
        self.failed = True
        self.fail_message += '\n' + message

        if e is not None:
            print('========= Exception Occur =========')
            print(e)
            print('===================================')

    def get_top_dlg(self, app, window_title):
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
                print('{} dlg is now active.'.format(window_title))
                return dlg
            else:
                print('{} dlg is not active. Waiting {}s...trial({})'.format(window_title, waiting_time, count))
                time.sleep(waiting_time)

        print('Getting {} dlg failed!!!'.format(window_title))
        self.report_failure(window_title)

        return None

    def get_dlg(self, parent_dlg, window_title):
        maximum_trial = 10
        waiting_time = 1

        dlg = parent_dlg.window(title_re=window_title)
        for count in range(1, maximum_trial):
            if dlg.exists():
                print('{} dlg is now active.'.format(window_title))
                return dlg
            else:
                print('{} dlg is not active. Waiting {}s...trial({})'.format(window_title, waiting_time, count))
                time.sleep(waiting_time)

        print('Getting {} dlg failed!!!'.format(window_title))
        self.report_failure(window_title)

        return None

    def get_dlg_title(self, parent_dlg, window_title):
        maximum_trial = 10
        waiting_time = 1

        dlg = parent_dlg.window(title=window_title)
        for count in range(1, maximum_trial):
            if dlg.exists():
                print('{} dlg is now active.'.format(window_title))
                return dlg
            else:
                print('{} dlg is not active. Waiting {}s...trial({})'.format(window_title, waiting_time, count))
                time.sleep(waiting_time)

        print('Getting {} dlg failed!!!'.format(window_title))
        self.report_failure(window_title)

        return None

    def wait_dlg(self, dlg):
        maximum_trial = 10
        waiting_time = 1
        for count in range(1, maximum_trial):
            if dlg.exists():
                return dlg
            else:
                time.sleep(waiting_time)
        print('Getting dlg failed!!!')

        return None

    def start_hanaro_deprecated(self):
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

    def start_hanaro(self):
        try:
            print('\nHanaro start-up procedure begins')
            app = Application('uia')
            app.start('C:/OSSTEM/HANAROOK/hanaro.exe')
            hanaro_dlg = self.get_top_dlg(app, '하나로')
            hanaro_login_dlg = self.get_dlg(hanaro_dlg, '로그인')
            hanaro_password_dlg = hanaro_login_dlg.window(title='비밀 번호:', auto_id='1001', control_type="Edit")
            hanaro_password_dlg.type_keys(self.hanaro_password + '{ENTER}')
            time.sleep(0.5)
            hanaro_today_dlg = hanaro_dlg.window(title_re='하나로 투데이')
            if hanaro_today_dlg.exists():
                hanaro_today_checkbox = hanaro_today_dlg.window(title='오늘은 열지 않음')
                hanaro_today_checkbox.click_input()
        except Exception as e:
            self.report_failure('하나로', e)

    def start_pointnix(self):
        print('\nCDX-ViewADT start-up procedure begins')
        app = Application('uia')
        app.start('C:/PointNix/CDX-View/cdx_pointnix.exe')
        # Popen('C:/PointNix/CDX-View/cdx_pointnix.exe')

    def start_ADT_deprecated(self):
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

    def start_ADT(self):
        try:
            print('\nADT start-up procedure begins')
            pyautogui.hotkey('ctrl', 'win', 'left')
            pyautogui.hotkey('ctrl', 'win', 'left')
            app = Application('uia')
            app.start("C:/Program Files/ADT EYE3/EYE3.exe")
            adt_dlg = self.get_top_dlg(app, 'ADT')
            adt_dlg.double_click_input(coords=(110, 250))
            adt_dlg.click_input(coords=(150, 270), button_up=False)
            adt_dlg.click_input(coords=(400, 275), button_down=False)
            # maximize_dlg = adt_dlg['Button5']
            # maximize_dlg.click_input()
            pyautogui.hotkey('ctrl', 'win', 'right')
        except Exception as e:
            self.report_failure('ADT', e)

    def start_excel(self):
        try:
            print('\nExcel start-up procedure begins')
            app = Application('uia')
            app.start('C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE "d:/개원/경영/차오름치과 경영.xlsx"')
            excel_dlg = self.get_top_dlg(app, '차오름치과 경영.xlsx - Excel')
            excel_dlg.maximize()
        except Exception as e:
            self.report_failure('Excel', e)

    def start_explorer(self):
        print('\nExplore start-up procedure begins')
        app = Application('uia')
        app.start('explorer d:\\')
        # Popen(('explorer', 'd:\\'))

    def start_pycharm(self):
        print('\nPyCharm start-up procedure begins')
        app = Application('uia')
        app.start('C:/Program Files/JetBrains/PyCharm Community Edition 2023.3.5/bin/pycharm64.exe')

    def start_chrome_deprecated(self):
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

    def start_chrome(self):
        try:
            print('\nChrome start-up procedure begins')
            app = Application('uia')
            app.start('C:/Program Files/Google/Chrome/Application/chrome.exe')
            chrome_dlg = self.get_top_dlg(app, '새 탭')
            search_dlg = chrome_dlg['Edit']

            search_dlg.type_keys('https://mail.naver.com/v2/folders/0/all''{ENTER}')
            # # search_dlg.type_keys('https://naver.com''{ENTER}')
            # # naver_dlg = self.get_top_dlg(app, 'NAVER')
            # # naver_mail_dlg = self.get_dlg_title(naver_dlg, '메일')
            # # naver_mail_dlg.click_input()

            pyautogui.hotkey('ctrl', 't')
            search_dlg.type_keys('https://kebhana.com''{ENTER}')
            hana_dlg = self.get_top_dlg(app, '하나은행')
            hana_check_dlg = hana_dlg['조회2']
            hana_check_dlg.click_input()
            hana_cert_dlg = hana_dlg['공동인증서 로그인(구 공인인증서)']
            self.wait_dlg(hana_cert_dlg)
            hana_cert_dlg.click_input()

            pyautogui.hotkey('ctrl', 't')
            search_dlg.type_keys('https://remotedesktop.google.com''{ENTER}')

            #     chrome_newtag_dlg = chrome_dlg.window(title='새 탭', control_type='Button')
            #     chrome_newtag_dlg.click_input()
        except Exception as e:
            self.report_failure('Chrome', e)

    def start_kakaotalk(self):
        try:
            print('\nKakaoTalk start-up procedure begins.')
            app = Application('uia')
            app.start('C:/Program Files (x86)/Kakao/KakaoTalk/KakaoTalk.exe')
            kakaotalk_dlg = self.get_top_dlg(app, '카카오톡')
            password_dlg = kakaotalk_dlg['Edit2']
            typing_message = '^a''{DELETE}' + self.message_password + '{ENTER}'
            password_dlg.type_keys(typing_message)
            self.wait_dlg(kakaotalk_dlg)
            kakaotalk_dlg.minimize()
        except Exception as e:
            self.report_failure('KakaoTalk', e)

    def start_line(self):
        try:
            print('\nLine start-up procedure begins.')
            app = Application('uia')
            app.start('C:/Users/neuroris/AppData/Local/LINE/bin/LineLauncher.exe')
            line_dlg = self.get_top_dlg(app, 'LINE')
            password_dlg = line_dlg['Edit2']
            typing_message = '^a''{DELETE}' + self.message_password + '{ENTER}'
            password_dlg.type_keys(typing_message)
            self.wait_dlg(line_dlg)
            line_dlg.minimize()
        except Exception as e:
            self.report_failure('LINE', e)

    def start_millie_deprecated(self):
        try:
            print('millie start-up procedure begins')

            app = Application('uia')
            app.start('C:/Program Files/millie/millie.exe')
            millie_dlg = app.top_window()
            study_dlg = millie_dlg.window(title_re='내 서재')
            time.sleep(4)
            study_dlg.set_focus()
            study_dlg.click_input()
        except Exception as e:
            self.report_failure('Millie', e)

    def start_millie(self):
        try:
            print('\nmillie start-up procedure begins')
            app = Application('uia')
            app.start('C:/Program Files/millie/millie.exe')
            millie_dlg = self.get_top_dlg(app, '밀리의 서재')
            study_dlg = self.get_dlg(millie_dlg, '내 서재')
            study_dlg.set_focus()
            study_dlg.click_input()
        except Exception as e:
            self.report_failure('Millie', e)

    def start_kiwoom(self):
        Popen('C:/KiwoomHero4/Bin/NKStarter.exe')

    def start_efriend(self):
        Popen('C:/eFriend Expert/efriendexpert/efriendexpert.exe')

if __name__ == '__main__':
    ws = WookStart()