#ライブラリ読み込み
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from oauth2client.service_account import ServiceAccountCredentials
import time
import gspread

#rangeの初期値2
range1 = 47
range2 = 125

#jsonファイルを使って認証情報を取得
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
c = ServiceAccountCredentials.from_json_keyfile_name('C:\\Users\\tenal\\OneDrive\\デスクトップ\\python\\eightknot-4437ab21367f.json', scope)

#認証情報を使ってスプレッドシートの操作権を取得
gs = gspread.authorize(c)
SPREADSHEET_KEY = '17bmqAs-pH30KM7ux9bcatyTVsEI2Svr0YOBBet84g3c'
worksheet = gs.open_by_key(SPREADSHEET_KEY).worksheet("新カード")

#クロームの立ち上げ
chop = Options()
chop.set_capability('pageLoadStrategy', 'normal')
driver=webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chop)
url = "https://kouen.sports.metro.tokyo.lg.jp/web/"

#ページ接続
driver.get(url)


for i in range(range1, range2):

    #ログイン画面に移動する
    wait = WebDriverWait(driver, 3)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#btn-login')))
    driver.find_element(By.CSS_SELECTOR, '#btn-login').click()

    #スプレッドシートの情報を取得
    print(i)
    place = worksheet.acell("C"+str(i)).value
    placepath = "//input[@name=\"layoutChildBody:childForm:igcdListItems:"+ place +":doAreaSet\"]"
    date = worksheet.acell("D"+str(i)).value
    datepath = "//a[@onclick=\"javascript:selectCalendarDate(2023,"+ date +");return false;\"]"
    id = worksheet.acell("G"+str(i)).value
    pw = worksheet.acell("H"+str(i)).value

    #ログイン情報を入力する
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#userId')))
    driver.find_element(By.CSS_SELECTOR, '#userId').send_keys(id)
    driver.find_element(By.CSS_SELECTOR, '#password').send_keys(pw)

    #ログインする
    driver.find_element(By.CSS_SELECTOR, '#btn-go').click()

    #当選したか判定する
    reservation = True #判定式リセット
    count = 0 #カウンターリセット
    #btn-go
    while reservation :
        # 抽選確認画面に移動する
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#nav-menu > nav > div:nth-child(3) > a')))
        driver.find_element(By.CSS_SELECTOR, '#nav-menu > nav > div:nth-child(3) > a').click()
        # wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#modal-menus > div > div > div > table > tbody > tr:nth-child(3) > td')))
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, '#modal-menus > div > div > div > table > tbody > tr:nth-child(3) > td > a').click()

        try: 
            # 当選しなかった場合
            driver.find_element(By.CSS_SELECTOR, '#text-area')
            reservation = False
        except Exception:
            # 当選した場合
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#refine-checkbox0 > label')))
            driver.find_element(By.CSS_SELECTOR, '#refine-checkbox0 > label').click()
            driver.find_element(By.CSS_SELECTOR, '#btn-go').click()
            # 利用人数を指定
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#applynum0')))
            driver.find_element(By.CSS_SELECTOR, '#applynum0').send_keys(4)
            # 当選確定する
            driver.find_element(By.CSS_SELECTOR, '#btn-go').click()
            alert = driver.switch_to.alert
            alert.accept()
            #当選していたらスプレッドシートに記入
            worksheet.update_cell(i, 9 + count, '当選')
            count = count + 1 #カウンタ更新
        
        # #当選していなければ、ループから抜け出す
        # if element.text == "【当選】確定する" :
        #     #当選していたら確定する
        #     element.click()
        #     driver.implicitly_wait(20)
        #     driver.find_element(By.XPATH,'//*[@id="doOnceLockFix"]').click()

        #     #当選していたらスプレッドシートに記入
        #     worksheet.update_cell(i, 9 + count, '当選')

        #     #マイページに戻る
        #     driver.implicitly_wait(20)
        #     driver.find_element(By.XPATH,'/html/body/div/form[1]/table[1]/tbody/tr/td[3]/a/img').click()
        #     count = count + 1 #カウンタ更新
        # else :
        #     reservation = False

    #ログアウトする
    driver.find_element(By.CSS_SELECTOR, '#userName').click()
    driver.find_element(By.CSS_SELECTOR, '#userMenu > div > a:nth-child(9)').click()


#10秒終了を待つ
time.sleep(5)

#クロームの終了処理
driver.close()