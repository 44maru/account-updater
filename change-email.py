import sys
import threading
import queue
import random
import logging.config
import subprocess
import base64
import os

from selenium.webdriver.common.keys import Keys

import licencemanager
from datetime import datetime as dt
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException

ACCOUNT_SETTING_URL = "https://www.nike.com/jp/ja_jp/p/settings"
CHROME_DRIVER_PATH = "./chromedriver.exe"
PHANTOMJS_DRIVER_PATH = "./phantomjs.exe"
LOG_CONF = "./logging.conf"
INPUT_CSV = "./input.csv"
CONFIG_TXT = "./config.txt"
PROXY_TXT = "./proxy.txt"
CHROME_PROXY_EXTENTION = "proxy_mng.crx"
KEY_THREAD_NUM = "THREAD_NUM"
CONFIG_DICT = {}
PROXY_LIST = []

HTML_LOGIN_PATH = "/html/body/div[6]/nav/div[1]/ul[2]/li[2]/button/span"
HTML_LOGIN_EMAIL_PATH = """//*[@id='nike-unite-loginForm']/div[2]/input"""
HTML_LOGIN_PASS_PATH = """//*[@id="nike-unite-loginForm"]/div[3]/input"""
HTML_KEEP_LOING_CHKBOX_PATH = """//*[@id="keepMeLoggedIn"]/label"""
HTML_LOGIN_BUTTON_PATH = """//*[@id="nike-unite-loginForm"]/div[6]/input"""
HTML_LOGIN_ERR_MSG_PATH = """//*[@id="nike-unite-loginForm"]/div[1]/ul/li"""
HTML_LOGIN_BLOCK_MSG_PATH = """//*[@id="nike-unite-error-view"]/div/ul/li"""
HTML_ACCOUNT_SETTING_PATH = """/html/body/div[7]/nav/div[1]/ul[2]/li[1]/a/span[2]"""
HTML_ACCOUNT_SETTING_EMAIL_PATH = """//*[@id="email"]"""
HTML_ACCOUNT_SETTING_LAST_NAME_KANA_PATH = """//*[@id="kana-last-name"]"""
HTML_ACCOUNT_SETTING_FIRST_NAME_KANA_PATH = """//*[@id="kana-first-name"]"""
HTML_ACCOUNT_SETTING_FEET_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[12]/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/a/span"""
HTML_ACCOUNT_SETTING_FEET_3_0_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[12]/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/ul/li[2]"""
HTML_ACCOUNT_SETTING_POND_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[12]/div[4]/div[1]/div[1]/div[2]/div[2]/div[1]/a/span"""
HTML_ACCOUNT_SETTING_POND_30_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[12]/div[4]/div[1]/div[1]/div[2]/div[2]/div[1]/ul/li[2]"""
HTML_ACCOUNT_SETTING_COUNTRY_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[15]/div[2]/div[1]/div/a/span"""
HTML_ACCOUNT_SETTING_COUNTRY_JAPAN_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[15]/div[2]/div[1]/div/ul/li[1]"""
HTML_ACCOUNT_SETTING_STATE_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[15]/div[2]/div[2]/div/a/span"""
HTML_ACCOUNT_SETTING_STATE_HOKKAIDO_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[15]/div[2]/div[2]/div/ul/li[1]"""
HTML_ACCOUNT_SETTING_ZIP_CODE_PATH = """//*[@id="postalCode"]"""
HTML_ACCOUNT_SETTING_SAVE_BUTTON_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[19]/button[2]"""
HTML_ACCOUNT_SETTING_SAVE_RESULT_PATH = """//*[@id="content"]/div[1]/div[2]/div[1]/form/div[2]/span"""
HTML_LOGOUT_PATH = """//*[@id="exp-profile-dropdown"]/ul/li[5]/a"""
HTML_MY_ADD_DIV_PATH_TEMPL = """//*[@id="nike-unite-loginForm"]/div[{}]"""

SUCCESS = "成功"
ERROR = "失敗"
OUT_DIR = "result"

HOKKAIDO = "北海道"
AKAHIRA_ZIP_CODE = "079-1143"
LAST_NAME = "タナカ"
FIRST_NAME = "タイチ"

LOGIN_ERR_MSG = "入力されたEメールアドレスまたはパスワードに誤りがあります。"
LOGIN_BLOCK_MSG = "申し訳ありませんが、現在、サーバーに接続できません。 あとでもう一度お試しください。"
ACCOUNT_SAVE_SUCCESS_MSG = "変更が正常に保存されました。"
WAIT_SEC = 20

JAVA_SCRIPT_TO_RESIZE_WINDOW = """
  function getRandomInt(max) {
    return Math.floor(Math.random() * Math.floor(max));
  }
  
  for(x=0; x<200;x++) {
      var mybnbdiv = document.createElement('div');
      var bnb_parent = document.querySelector('form#nike-unite-loginForm');
      mybnbdiv.className = 'bnbnike' + getRandomInt(10000);
      mybnbdiv.style.width = getRandomInt(50) + 'px';
      mybnbdiv.style.height = getRandomInt(50) + 'px';
      mybnbdiv.style.display = 'inline-block';
      bnb_parent.appendChild(mybnbdiv);
  }
"""


class LoginError(Exception):
    pass
    # def __init__(self):
    #    super(LoginError, self).__init__()


class SitePasteThread(threading.Thread):
    def __init__(self):
        super(SitePasteThread, self).__init__()
        self.b64_json = None
        self.driver = None
        self.element = None
        self.action = None

    def run(self):
        global input_q
        global output_q
        global lock
        options = webdriver.ChromeOptions()
        options.add_argument("--disk-cache=false")
        options.add_extension(CHROME_PROXY_EXTENTION)
        # options.add_argument('--incognito')
        # options.add_argument("--load-images=false")
        # options.add_argument(
        #    "--user-agent=Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36")
        options.add_argument("--disable-infobars")
        # options.add_argument("--window-position=-3200,-3200")
        # options.add_argument("--headless")
        # options.add_argument("--proxy-server=203.104.207.178:3128")
        # options.add_argument("--proxy-server=125.6.65.74:3128")
        # week5/12JKiuy11
        # options.add_argument("--remote-debugging-port=9222")
        # options.add_argument('--disable-gpu')
        # options.add_argument('--ignore-certificate-errors')
        # options.add_argument('--allow-running-insecure-content')

        """
        self.driver = webdriver.Remote(
            command_executor='http://localhost:4444/wd/hub',
            desired_capabilities={
            'browserName': 'chrome',
            'chromeOptions': {
                'args': [
                    '--headless',
                    '--ignore-certificate-errors',
                    '--proxy-server=http://localhost:8080'
                ]
            }
        })
        """
        success_cnt = 0
        err_cnt = 0
        while True:
            try:
                item = input_q.get(timeout=1)

                if self.driver is None:
                    lock.acquire()
                    self.driver = webdriver.Chrome(
                        executable_path=CHROME_DRIVER_PATH, chrome_options=options, service_args=["hide_console"])
                    lock.release()
                self.driver.delete_all_cookies()
                self.driver.set_window_size(random.randint(700, 1000), random.randint(700, 1000))
                self.exec_selenium(item[0], item[1], item[2])
                log.info("Succeeded to update account %s", item[0])
                output_q.put([item[0], item[1], SUCCESS, ""])
                success_cnt += 1

            except queue.Empty:
                log.debug("Nothing job in input_q. Finished process...")
                break;

            except LoginError as e:
                log.info("Failed to update account %s", item[0])
                output_q.put([item[0], item[1], ERROR, e])
                err_cnt += 1

            except Exception as e:
                log.info("Failed to update account %s", item[0])
                log.exception("Unknown Exception : %s.", e)
                output_q.put([item[0], item[1], ERROR, "Unknown error. Send log file to developer."])
                self.driver.save_screenshot("%s/%s.err.png" % (OUT_DIR, item[0]))
                err_cnt += 1
            finally:
                log.debug("Current Success:{} Current Error:{}".format(success_cnt, err_cnt))
                if self.driver is not None:
                    self.driver.quit()
                    self.driver = None

    def send_keys(self, xpath, target_val):
        self.wait_display(xpath)
        self.driver.find_element_by_xpath(xpath).clear()
        self.driver.find_element_by_xpath(xpath).send_keys(target_val)

    def send_keys2(self, xpath, target_val):
        self.wait_display(xpath)
        self.driver.find_element_by_xpath(xpath).clear()
        for ch in target_val:
            self.driver.find_element_by_xpath(xpath).send_keys(ch)
            sleep(random.randint(10, 100) / 1000.0)

    def wait_display(self, xpath):
        cnt = 0
        while not self.is_displayed(xpath):
            if cnt > WAIT_SEC:
                raise Exception

            cnt += 1
            log.debug("%02d : No display : %s" % (cnt, xpath))
            sleep(1)

    def wait_clickable(self, xpath):
        cnt = 0
        while not EC.element_to_be_clickable((By.XPATH, xpath)):
            if cnt > WAIT_SEC:
                raise Exception

            cnt += 1
            log.debug("%02d : No clickable : %s" % (cnt, xpath))
            sleep(1)

    def is_displayed(self, xpath):
        while True:
            try:
                return self.driver.find_element_by_xpath(xpath).is_displayed()
            except (StaleElementReferenceException, NoSuchElementException):
                return False

    def click(self, xpath):
        self.click_move_if_need(xpath, False)

    def click_with_move(self, xpath):
        self.click_move_if_need(xpath, True)

    def click_move_if_need(self, xpath, need_move):
        self.wait_clickable(xpath)
        self.wait_display(xpath)

        if need_move:
            self.action = ActionChains(self.driver)
            self.element = self.driver.find_element_by_xpath(xpath)
            self.action.move_to_element(self.element).perform()
            sleep(0.1)

        self.driver.find_element_by_xpath(xpath).click()

    def exec_selenium(self, email, new_email, passwd):
        global need_mouse_moving
        global mouse_lock
        log.debug("Start updating account %s", email)
        # mouse_lock.acquire()
        # need_mouse_moving = True
        # mouse_lock.release()
        sleep(random.randint(500, 2000) / 1000.0)
        if 0 < len(PROXY_LIST):
            self.b64_json = base64.b64encode(("""{"url":"%s", "proxy":"%s"}""" % (
                ACCOUNT_SETTING_URL, PROXY_LIST[random.randint(0, len(PROXY_LIST) - 1)])).encode("utf-8"))
            self.driver.get("https://configure.bnb/" + self.b64_json.decode("utf-8"))
        else:
            self.driver.get(ACCOUNT_SETTING_URL)

        # self.do_random_operate_1()

        self.type_login_info(email, passwd)
        self.resize_and_scroll_window()

        # if random.randint(0, 1) == 0:
        #    self.random_operate_2()

        # if random.randint(0, 1) == 0:
        #    if self.driver.find_element_by_xpath(HTML_KEEP_LOING_CHKBOX_PATH).is_enabled():
        #        self.click(HTML_KEEP_LOING_CHKBOX_PATH)

        # if random.randint(0, 1) == 0:
        #    self.send_keys(HTML_LOGIN_EMAIL_PATH, email)

        # self.element = self.driver.find_element_by_xpath(HTML_LOGIN_BUTTON_PATH)
        # ActionChains(self.driver).move_to_element(self.element).perform()
        # self.random_operate_mouse_mv()

        sleep(random.randint(500, 4000) / 1000)
        log.debug("Click login button %s", email)
        self.click(HTML_LOGIN_BUTTON_PATH)

        cnt = 0
        while not self.is_displayed(HTML_ACCOUNT_SETTING_EMAIL_PATH):
            if self.is_displayed(HTML_LOGIN_ERR_MSG_PATH) \
                    and self.driver.find_element_by_xpath(HTML_LOGIN_ERR_MSG_PATH).text == LOGIN_ERR_MSG:
                log.debug("Login error. account(%s)", email)
                raise LoginError("Login error")

            if self.is_displayed(HTML_LOGIN_BLOCK_MSG_PATH) \
                    and self.driver.find_element_by_xpath(HTML_LOGIN_BLOCK_MSG_PATH).text == LOGIN_BLOCK_MSG:
                log.debug("Login blocked. account(%s)", email)
                raise LoginError("Login block")

            if cnt > WAIT_SEC:
                log.debug("Unknown error when try to login. account(%s)", email)
                raise Exception

            sleep(1)
            cnt += 1

        mouse_lock.acquire()
        need_mouse_moving = False
        mouse_lock.release()

        self.update_account_setting(email, new_email)

        # self.click(HTML_ACCOUNT_SETTING_FEET_PATH)

    def resize_and_scroll_window(self):
        self.driver.execute_script(JAVA_SCRIPT_TO_RESIZE_WINDOW)
        for i in range(2):
            self.element = self.driver.find_element_by_xpath(HTML_MY_ADD_DIV_PATH_TEMPL.format(random.randint(9, 208)))
            self.action = ActionChains(self.driver)
            self.action.move_to_element_with_offset(self.element, random.randint(1, 10), random.randint(1, 10))
            self.action.click()
            self.action.perform()
            sleep(random.randint(400, 1000) / 1000)

    def random_operate_mouse_mv(self):
        for i in range(100):
            ActionChains(self.driver).move_by_offset(random.randint(-1, 1), random.randint(-1, 1)).perform()
        ActionChains(self.driver).move_to_element(self.element).perform()
        ActionChains(self.driver).move_by_offset(random.randint(-10, 10), random.randint(-10, 10)).perform()

    def random_operate_2(self):
        self.element = self.driver.find_element_by_xpath(HTML_LOGIN_EMAIL_PATH)
        for i in range(3):
            ActionChains(self.driver).click_and_hold(self.element). \
                move_by_offset(random.randint(10, 10), random.randint(10, 10)).perform()

    def type_login_info(self, email, passwd):
        for i in range(random.randint(1, 5)):
            if random.randint(0, 1) == 0:
                self.send_keys(HTML_LOGIN_PASS_PATH, Keys.ENTER)
            else:
                self.send_keys(HTML_LOGIN_EMAIL_PATH, Keys.ENTER)
            sleep(0.5)

        self.send_keys(HTML_LOGIN_EMAIL_PATH, email)
        self.send_keys(HTML_LOGIN_PASS_PATH, passwd)

    def do_random_operate_1(self):
        for i in range(random.randint(1, 10)):
            if random.randint(0, 1) == 0:
                self.click(HTML_LOGIN_PASS_PATH)
            if random.randint(0, 1) == 0:
                self.click(HTML_LOGIN_EMAIL_PATH)
            if random.randint(0, 1) == 0:
                self.wait_clickable(HTML_LOGIN_PASS_PATH)
                self.wait_display(HTML_LOGIN_PASS_PATH)
                self.element = self.driver.find_element_by_xpath(HTML_LOGIN_PASS_PATH)
                ActionChains(self.driver).move_to_element(self.element).perform()
            if random.randint(0, 1) == 0:
                self.wait_clickable(HTML_LOGIN_EMAIL_PATH)
                self.wait_display(HTML_LOGIN_EMAIL_PATH)
                self.element = self.driver.find_element_by_xpath(HTML_LOGIN_EMAIL_PATH)
                ActionChains(self.driver).move_to_element(self.element).perform()

    def update_account_setting(self, email, new_email):
        self.send_keys(HTML_ACCOUNT_SETTING_EMAIL_PATH, new_email)
        self.send_keys(HTML_ACCOUNT_SETTING_LAST_NAME_KANA_PATH, LAST_NAME)
        self.send_keys(HTML_ACCOUNT_SETTING_FIRST_NAME_KANA_PATH, FIRST_NAME)

        self.click_from_drop_down_list(
            HTML_ACCOUNT_SETTING_FEET_3_0_PATH, HTML_ACCOUNT_SETTING_FEET_PATH, "FEET")
        self.click_from_drop_down_list(
            HTML_ACCOUNT_SETTING_POND_30_PATH, HTML_ACCOUNT_SETTING_POND_PATH, "POND")
        self.click_from_drop_down_list(
            HTML_ACCOUNT_SETTING_COUNTRY_JAPAN_PATH, HTML_ACCOUNT_SETTING_COUNTRY_PATH, "COUNTRY")
        self.click_from_drop_down_list(
            HTML_ACCOUNT_SETTING_STATE_HOKKAIDO_PATH, HTML_ACCOUNT_SETTING_STATE_PATH, "STATE")
        self.send_keys(HTML_ACCOUNT_SETTING_ZIP_CODE_PATH, AKAHIRA_ZIP_CODE)

        self.driver.find_element_by_xpath(HTML_ACCOUNT_SETTING_SAVE_BUTTON_PATH).submit()

        cnt = 0
        while True:
            if self.is_displayed(HTML_ACCOUNT_SETTING_SAVE_RESULT_PATH) \
                    and self.driver.find_element_by_xpath(
                    HTML_ACCOUNT_SETTING_SAVE_RESULT_PATH).text == ACCOUNT_SAVE_SUCCESS_MSG:
                log.debug("Success account changed (%s)", email)
                break

            if cnt > WAIT_SEC:
                log.debug("Save button was clicked but didn't show success message (%s)", email)
                raise Exception

            sleep(1)
            cnt += 1

    def click_from_drop_down_list(self, drop_down_list_path, click_path, drop_down_name):
        cnt = 0
        while not self.is_displayed(drop_down_list_path):
            self.click_with_move(click_path)
            if cnt > WAIT_SEC:
                log.debug(
                    "Clicked {} times {} drop down but not displayed drop down list.".format(
                        WAIT_SEC / 2, drop_down_name))
                raise Exception
            sleep(0.5)
            cnt += 1

        self.click(drop_down_list_path)


def load_config():
    for line in open(CONFIG_TXT, "r"):
        items = line.replace("\n", "").split("=")

        if len(items) != 2:
            continue

        if items[0] == KEY_THREAD_NUM:
            CONFIG_DICT[KEY_THREAD_NUM] = int(items[1])


def load_proxy():
    if not os.path.exists(PROXY_TXT):
        return

    for line in open(PROXY_TXT, "r"):
        PROXY_LIST.append(line.replace("\n", ""))


def read_input_csv():
    global input_q
    line_num = 0
    for line in open(INPUT_CSV, "r"):
        line_num += 1
        if line_num == 1:
            continue

        items = line.replace("\n", "").split(",")

        if len(items) != 3:
            continue

        input_q.put(items)


def change_email():
    thread_num = CONFIG_DICT[KEY_THREAD_NUM]
    thread_list = []

    log.info("Start processing.\n")

    for i in range(thread_num):
        thread = SitePasteThread()
        thread.start()
        thread_list.append(thread)

    for i in range(thread_num):
        thread_list[i].join()


def write_result_csv():
    success_cnt = 0
    error_cnt = 0
    f = open("%s\\%s" % (OUT_DIR, dt.now().strftime('result-%Y%m%d-%H%M%S.csv')), "w")
    f.write("旧アドレス,新アドレス,結果,失敗理由\n")
    for i in range(output_q.qsize()):
        items = output_q.get()
        f.write("%s,%s,%s,%s\n" % (items[0], items[1], items[2], items[3]))

        if items[2] == SUCCESS:
            success_cnt += 1
        else:
            error_cnt += 1
    f.close()
    log.info("")
    log.info("Success:%d Fail:%d" % (success_cnt, error_cnt))
    log.info("")
    log.info("Refer %s for more detail." % f.name)


def mach_license():
    return licencemanager.match_license()


def mv_mouse():
    global is_changing_account
    global need_mouse_moving
    global mouse_lock

    while is_changing_account:
        mouse_lock.acquire()
        if need_mouse_moving:
            subprocess.call("UWSC.exe mv_mouse.UWS")
        mouse_lock.release()
        sleep(1)


if __name__ == "__main__":
    logging.config.fileConfig(LOG_CONF)
    log = logging.getLogger()

    if not mach_license():
        input("Please push Enter to exit.")
        sys.exit(-1)

    lock = threading.Lock()
    mouse_lock = threading.Lock()
    input_q = queue.Queue()
    output_q = queue.Queue()
    load_config()
    load_proxy()
    read_input_csv()
    is_changing_account = True
    need_mouse_moving = False
    # threading.Thread(target=mv_mouse).start()
    change_email()
    write_result_csv()
    is_changing_account = False
    input("Please push Enter to exit.")
