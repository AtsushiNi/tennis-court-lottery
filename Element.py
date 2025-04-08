import random
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import TimeoutException

# ======================================
# 定数
# ======================================
WAIT_LIMIT_TIME=5

class Ele :
    # ======================================
    # メンバの定義
    # ======================================
    def __init__(self, driver, name='', xpath=''):
        self.__driver = driver
        self.__name = name
        self.__xpath = xpath
        try :
            self.__element = WebDriverWait(self.__driver, WAIT_LIMIT_TIME).until(
                EC.element_to_be_clickable((By.XPATH, self.__xpath))
            )
        except UnexpectedAlertPresentException as e :
            print('アラート表示により要素「{self.__name}」を見つけられませんでした')
            self.__element = None
        except TimeoutException as e :
            print(f'読み込み後{str(WAIT_LIMIT_TIME)}秒間要素「{self.__name}」を見つけられませんでした')
            self.__element = None

    # ======================================
    # メソッドの定義
    # ======================================
    # エレメントを返します
    def getElement(self) :
        return self.__element
    
    # この要素をクリックします
    def click(self) :
        if self.__element :
            try :
                self.__element.click()
                print(f"「{self.__name}」をクリックしました")
            except ElementNotInteractableException as e :
                print(f"「{self.__name}」の要素が非活性状態なのでクリックできませんでした")
                self.__element = None
            except ElementClickInterceptedException as e :
                print(f"「{self.__name}」のクリックは遮られました")
        else :
            print(f"「{self.__name}」要素が読み込まれていないため、クリックできませんでした")

    # この要素に文字を送信します
    def sendKeys(self, key) :
        if self.__element :
            try :
                self.__element.send_keys(key)
                print(f"「{self.__name}」に「{str(key)}」と入力しました")
            except ElementNotInteractableException as e :
                print(f"「{self.__name}」の要素が非活性状態なので文字を送信できませんでした")
                self.__element = None
        else :
            print(f"「{self.__name}」要素が読み込まれていないため、文字を送信できませんでした")

    # この要素にエンターを送信します
    def sendEnter(self) :
        self.sendKeys(Keys.ENTER)

    # この要素にJavaScriptを実行します
    def executeScript(self, script) :
        self.__driver.execute_script(script, self.__element)

    # この要素のAlt属性を返却します
    def getAlt(self) :
        if self.__element :
            return self.__element.get_attribute('alt')
        else :
            return None

    # 指定した要素の属性を返却します
    def getAttribute(self, key) :
        if self.__element :
            return self.__element.get_attribute(key)
        else :
            return None

    # この要素が存在するかどうかを返却します
    def exist(self) :
        return self.__element != None

    # 画面外の要素をクリックします
    def click_out_of_screen(self) :
        if self.__element :
            try :
                self.__driver.execute_script("arguments[0].click();", self.__element)
                print(f"「{self.__name}」をクリックしました")
            except ElementNotInteractableException as e :
                print(f"「{self.__name}」の要素が非活性状態なのでクリックできませんでした")
                self.__element = None
            except ElementClickInterceptedException as e :
                print(f"「{self.__name}」のクリックは遮られました")
        else :
            print(f"「{self.__name}」要素が読み込まれていないため、クリックできませんでした")
    
     # ロボット対策のスリープ処理
    def randomSleep(self) :
        sleepTime = random.randrange(15) / 10 + 1
        time.sleep(sleepTime)