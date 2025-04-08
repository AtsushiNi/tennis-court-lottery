#ライブラリ読み込み
import sys
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from oauth2client.service_account import ServiceAccountCredentials
import time
import gspread
from Element import Ele

URL = "https://kouen.sports.metro.tokyo.lg.jp/web/"

#rangeの初期値 スタートは2から
start_index = 2
end_index = 125

#jsonファイルを使って認証情報を取得
try:
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]
    c = ServiceAccountCredentials.from_json_keyfile_name(
        './tennis-court-lottery-c005eff2651c.json',
        scope
    )
except FileNotFoundError:
    print("認証情報のファイルが見つかりませんでした。ファイルパスを確認してください。")
    sys.exit()

#認証情報を使ってスプレッドシートの操作権を取得
try:
    gs = gspread.authorize(c)
    SPREADSHEET_KEY = '17bmqAs-pH30KM7ux9bcatyTVsEI2Svr0YOBBet84g3c'
    worksheet = gs.open_by_key(SPREADSHEET_KEY).worksheet("新カード")
except Exception as e:
    print("スプレッドシートの操作権を取得できませんでした。認証情報が正しいか、スプレッドシートのキーとワークシート名が正しいか確認してください。")
    sys.exit()

#クロームの立ち上げ
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

#ページ接続
driver.get(URL)


#パソコンからのご利用はコチラ（多機能版）
# driver.find_element(By.CSS_SELECTOR, 'a[href="./user/view/user/homeIndex.html"]').click()
# driver.find_element(By.CSS_SELECTOR, 'body > a:nth-child').click()

# スプレッドシートからデータを一括で取得
data = worksheet.get("B{}:H{}".format(start_index, end_index))

for row in data:

    #ログイン画面に移動する
    ele = Ele(driver, "ログイン画面へ", "/html/body/div[2]/div/form/header/div/div[1]/div/button[1]")
    ele.click()

    #スプレッドシートの情報を取得
    print(row)
    place = row[1]
    placepath = f'/html/body/div[2]/div/form/main/div/article/div[2]/div[1]/span/select/option[{place}]'
    day = row[2]
    daypath = f'//*[contains(text(), "{day}日")]'
    id = row[5]
    pw = row[6]

    #ログイン情報を入力する
    ele = Ele(driver, "ユーザーID", "/html/body/div[2]/div/form/main/div/article/table/tbody/tr[1]/td/input")
    ele.sendKeys(id)
    ele = Ele(driver, "パスワード", "/html/body/div[2]/div/form/main/div/article/table/tbody/tr[2]/td/input")
    ele.sendKeys(pw)

    #ログインする
    ele = Ele(driver, "ログインする", "/html/body/div[2]/div/form/main/div/article/div[2]/button[1]")
    ele.click()
    
    #抽選がかかってたらキャンセルする処理
    #抽選をかける
    ele = Ele(driver, "抽選画面へ", "/html/body/div[2]/div/form/header/div/div[2]/div/nav/div[2]/a")
    ele.click()
    
    ele = Ele(driver, "抽選画面へ", "/html/body/div[2]/div/form/div[3]/div/div/div/table/tbody/tr[1]/td/a")
    ele.click()
    ele = Ele(driver, "テニス(人工芝)", "/html/body/div[2]/div/form/main/div/article/div[2]/table/tbody/tr[4]/td[5]/button")
    ele.click()
    
    
    #公園選択
    #二回抽選をかける
    for j in range(0):
    # for j in range(2):
        #日付指定
        facilityEle = Ele(driver, "施設選択", "/html/body/div[2]/div/form/main/div/article/div[2]/div[2]/span/select/option[2]")
        facilityEle.click()
        facilityEle = Ele(driver, "コート選択", placepath)
        facilityEle.click()
        time.sleep(2)
        facilityEle = Ele(driver, "施設選択", "/html/body/div[2]/div/form/main/div/article/div[2]/div[2]/span/select/option[2]")
        facilityEle.click()

        dayElement = Ele(driver, "日付確認", daypath)
        time.sleep(1)
        
        while dayElement.exist() == False:
            ele = Ele(driver, "施設選択", "/html/body/div[2]/div/form/main/div/article/div[4]/div[1]/div[2]/div/span[2]/button[1]")
            ele.click()
            time.sleep(1)
            dayElement = Ele(driver, "日付確認", daypath)
            #アラート許可
            # alert = driver.switch_to.alert
        
        print(dayElement.exist())
        print(dayElement.getElement().parent)
        time.sleep(1)
        dayParent = dayElement.getElement().find_element(By.XPATH, '../..')
        # element = driver.find_element(By.XPATH, '/html/body/div[2]/div/form/main/div/article/div[4]/div[1]/table/thead/tr/th[7]/div[1]/span[2]')
        lastchar = dayParent.get_attribute('id')[-1]
        getdate = int(lastchar) + 1

        # driver.execute_script('window.scrollTo(0, 500);')

        print("------------------------------------------------------")
        datepath = f'/html/body/div[2]/div/form/main/div/article/div[4]/div[1]/table/tbody/tr[5]/td[{getdate}]'
        time.sleep(2)
        ele = Ele(driver, "日付選択", datepath)
        ele.click()
        ele = Ele(driver, "申込みへ", "/html/body/div[2]/div/form/main/div/article/div[4]/div[3]/button[1]")
        ele.click()
        
        #申し込み画面
        ele = Ele(driver, "申込み件数選択", "/html/body/div[2]/div/form/main/div/article/div[2]/table/tbody/tr/td/span/select/option[2]")
        ele.click()
        # driver.find_element(By.CSS_SELECTOR, '#info-area > table > tbody > tr > th').click()
        ele = Ele(driver, "申し込む", "/html/body/div[2]/div/form/main/div/article/div[3]/button[1]")
        ele.click()
        #アラート許可
        alert = driver.switch_to.alert
        print(alert.text)
        alert.accept()
        # ele = Ele(driver, "ロボットチェック", '//*[@id="recaptcha-anchor"]/div[4]')
        # ele.click()
        time.sleep(12)
        # ロボット対策が出てきたら手動で操作する。
        try :
            element = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/form/main/div/article/div[2]/button"))
            )
            ele = Ele(driver, "続けて申し込む", "/html/body/div[2]/div/form/main/div/article/div[2]/button")
            ele.click()
        except TimeoutException as e :
            ele = Ele(driver, "申込み件数選択", "/html/body/div[2]/div/form/main/div/article/div[2]/table/tbody/tr/td/span/select/option[2]")
            ele.click()
            ele = Ele(driver, "申し込む", "/html/body/div[2]/div/form/main/div/article/div[3]/button[1]")
            ele.click()
            #アラート許可
            alert = driver.switch_to.alert
            print(alert.text)
            alert.accept()
            ele = Ele(driver, "続けて申し込む", "/html/body/div[2]/div/form/main/div/article/div[2]/button")
            ele.click()
        
        

        #申し込み画面
        # applypath = f"//input[@value='{j}-1']"
        # element6 = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, "apply")))
        # print(element6.get_attribute('class'))
        # element7 = driver.find_element(By.XPATH, '//*[contains(text(), "選択してください。")]')
        # print(parent.get_attribute('value'))
        #apply > option:nth-child(2)
        # print(driver.find_element(By.XPATH, "/html/body/div[2]/div/form/main/div/article/div[2]/table/tbody/tr/td/span/select").get_attribute('class'))
        # driver.find_element(By.XPATH, applypath).click()
        # try:
        #     driver.find_element(By.CSS_SELECTOR, '#nav-home').click()
        #     wait = WebDriverWait(driver, 30)
        #     wait.until(EC.alert_is_present())
        #     alert = driver.switch_to.alert
        #     print(alert.text)
        #     print("アラートは発生しませんでした")
        #     alert.accept()
        #     time.sleep(3)
        #     driver.find_element(By.CSS_SELECTOR, '#subtitle').click()
        # except TimeoutException:
        #     print("アラートは発生しませんでした")
        # except Exception as e:
        #     Alert(driver).accept()
        #     time.sleep(3)
        # driver.find_element(By.CSS_SELECTOR, '#btn-go').click()
        # Alert(driver).accept()

        # #続けて申し込み
        # driver.find_element(By.CSS_SELECTOR, '#btn-light').click()
        # driver.implicitly_wait(20)


    #ログアウトする
    ele = Ele(driver, "メニュー", "/html/body/div[2]/div/form/header/div/div[1]/div/div[2]/div/a")
    ele.click()
    ele = Ele(driver, "ログアウト", "/html/body/div[2]/div/form/header/div/div[1]/div/div[2]/div/div/a[6]")
    ele.click()
    



#10秒終了を待つ
time.sleep(5)

#クロームの終了処理
driver.close()

