#selenium起動
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# 指定時間待機
import time
import re

#gmail
from gmail import sendGmail

# マウスやキーボード操作に利用
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.by import By

# セレクトボックスの選択に利用
from selenium.webdriver.support.ui import Select

DISP_MODE = "ON"   # "ON" or "OFF"
USER_ID = "316-0389811"
USER_PWD = "3r8mZYN5HX"
ORD_PWD = "fAGgL9vWzJ"
MAIL_ADR = 'mtake88@gmail.com'
MAIL_PWD = 'jnfzzdwkghwmrgkm'

RETRY = 3
orderList = []  # 注文内容をメールで送信

#-----------------------------
# gmail送信
#-----------------------------
def sendIpoMail(type):
    if type == -1:      #重要なお知らせ
        sub = "重要なお知らせがあります"
        hon = "重要なお知らせがあります。\nログインして内容をご確認下さい\n（IPO申込は行なっていません）"
    elif type == -2:    #ログインエラー
        sub = "ログインエラーが発生しました"
        hon = "ログインエラーが発生しました。\n内容を確認して下さい"
    elif type == -3:  # HTMLエラー
        sub = "HTMLエラーが発生しました"
        hon = "HTMLエラーが発生しました。\n内容を確認して下さい"
    elif type == 0:   #申込OK
        sub = f'本日のIPO申込：{len(orderList)}件'
        hon = '本日のIPO申し込み状況\n'
        for ll in orderList:
            hon += f"{ll[0]}  ■単価：{ll[1]} ■株数：{ll[2]}\n"
    else:
        return
    sendGmail(MAIL_ADR, MAIL_ADR, MAIL_ADR, MAIL_PWD, "【SBI IPO】"+sub, hon)

#-----------------------------
#ログアウトする
#-----------------------------
def sbiLogOut():
    try:
        tmp = driver.find_element(by=By.XPATH, value="//img[@title='ログアウト']")
    except NoSuchElementException:
        print("err")
        sendIpoMail(-3)
        return
    tmp.click()
    time.sleep(2)

#-----------------------------
#SBI証券の口座でIPOのBB申込を行なう
#-----------------------------
def sbiIpoOrder():
    # サイトを開く
    driver.get("https://www.sbisec.co.jp/ETGate")
    time.sleep(3)

    # ユーザIDを入力
    userID = driver.find_element(by=By.NAME, value="user_id")
    userID.send_keys(USER_ID)

    # パスワードを入力
    userpass = driver.find_element(by=By.NAME, value="user_password")
    userpass.send_keys(USER_PWD)

    # ログインをクリック
    login = driver.find_element(by=By.NAME, value="ACT_login")
    login.click()

    time.sleep(3)

    # name属性で指定
    try:
        moneyTag = driver.find_element(by=By.XPATH,
                                       value="/html/body/table/tbody/tr[1]/td[1]/div[2]/div[1]/div/div/div/div/table/tbody/tr/td[1]/span")
    except NoSuchElementException:
        tmp = driver.find_elements(by=By.XPATH, value="//b[contains(text(),'重要なお知らせ')]")
        if len(tmp) >= 1:
            ii = -1
        else:
            ii = -2
        return ii

    money = int(moneyTag.text.replace(",", ""))
    print(money)

    # IPOページに入る
    driver.find_element(by=By.LINK_TEXT, value="IPO・PO").click()
    time.sleep(3)
    driver.find_element(by=By.XPATH, value="//img[@alt='新規上場株式ブックビルディング / 購入意思表示']").click()
    time.sleep(3)

    # 申込できる銘柄をチェック
    while True:
        reqs = driver.find_elements(by=By.XPATH, value="//img[@alt='申込']")

        if len(reqs) > 0:
            order_one = []
            print(f"ipo = {len(reqs)}")
            reqs[0].click()
            time.sleep(3)

            # ===== 個別の申し込み画面 =====
            # 銘柄名
            tag = driver.find_element(by=By.XPATH,
                                      value="/html/body/table/tbody/tr/td/table[1]/tbody/tr/td/table[1]/tbody/tr[1]/td/form/table[4]/tbody/tr/td")
            print(tag.text)
            order_one.append(tag.text)

            kari_tag = driver.find_element(by=By.XPATH,
                                           value="/html/body/table/tbody/tr/td/table[1]/tbody/tr/td/table[1]/tbody/tr[1]/td/form/table[5]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/div")

            if 'いずれか' in kari_tag.text :
                ss = kari_tag.text.split('.')
                kari = int(ss[0].replace(",", ""))

            else:
                ss = kari_tag.text.split(' ')
                kari = int(ss[-2].replace(",", ""))
                # print(kari)
            order_one.append(kari)

            unit_tag = driver.find_element(by=By.XPATH, value="//td[contains(text(),'売買単位')]")
            unit = int(re.findall(r'/(.*)株', unit_tag.text)[0])
            # print(unit)
            cnt = int(money / (kari * unit)) * unit  # 申込株数
            # print(cnt)

            # name属性で数量指定
            suryo = driver.find_element(by=By.NAME, value="suryo")
            suryo.send_keys(str(cnt))
            order_one.append(cnt)

            # id属性で「ストライクプライス」にチェック指定
            strike = driver.find_element(by=By.ID, value="strPriceRadio")
            strike.click()

            # print(order_one)

            # name属性で数量指定
            toripass = driver.find_element(by=By.NAME, value="tr_pass")
            toripass.send_keys(ORD_PWD)

            # name属性で申込確認ボタン指定・クリック
            driver.find_element(by=By.NAME, value="order_kakunin").click()
            time.sleep(2)

            # ===== 申し込み確認画面 =====
            # name属性で申込ボタン指定・クリック
            driver.find_element(by=By.NAME, value="order_btn").click()
            time.sleep(2)

            # ===== 申し込み完了画面 =====
            # IPOトップへ戻る・クリック
            orderList.append(order_one)
            driver.find_element(by=By.XPATH, value="//a[contains(text(),'購入意思表示画面へ戻る')]").click()
            time.sleep(3)
        else:
            break
    return 0

if __name__ == "__main__":

    if DISP_MODE == "OFF":
        options = Options()
        options.add_argument('--headless')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    else:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    for retry in range(RETRY):
        ret = sbiIpoOrder()
        sbiLogOut()
        if ret == 0:
            break

    sendIpoMail(ret)

    driver.quit()

    print(orderList)
    print("complete")

