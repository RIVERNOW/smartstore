import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import tkinter as tk
from tkinter import filedialog
import openpyxl
import requests
import bcrypt
import pybase64
import urllib.request
import urllib.parse
from datetime import datetime, timedelta
from selenium.webdriver.common.keys import Keys
import pyperclip
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText






# Get webdriver
co = Options()
driver = webdriver.Chrome(options=co)
driver.implicitly_wait(10)


driver.get('https://sell.smartstore.naver.com/#/home/about')
#새탭 안 열리게 방지
# driver.execute_script("window.open = function(url, name, features) { window.location.href = url; }")

#########################################################################################

#프로그램 시작 전 프로그램 실행에 필요한 데이터가 담겨 있는 excel 파일 선택 및 엑세스
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

wb = openpyxl.load_workbook(file_path, data_only=True)
load_data_ws = wb['Sheet1']
load_mail_ws = wb['Sheet2']

# 필요한 데이터 엑셀 파일에서 가져와 변수 선언.
client_id = load_data_ws['A2'].value
client_secret = load_data_ws['B2'].value
naver_id = load_data_ws['C2'].value
naver_pwd = load_data_ws['D2'].value
naver_store = load_data_ws['E2'].value
my_email = load_data_ws['F2'].value
my_email_pwd = load_data_ws['G2'].value

# 스마트스토어 API 코드, 토큰값 가져오는 함수
def get_token(client_id, client_secret, type_="SELF") -> str:
    timestamp = str(int((time.time()-3) * 1000))
    pwd = f'{client_id}_{timestamp}'
    hashed = bcrypt.hashpw(pwd.encode('utf-8'), client_secret.encode('utf-8'))
    client_secret_sign = pybase64.standard_b64encode(hashed).decode('utf-8')

    headers = {"content-type": "application/x-www-form-urlencoded"}
    data_ = {
        "client_id": client_id,
        "timestamp": timestamp,
        "client_secret_sign": client_secret_sign,
        "grant_type": "client_credentials",
        "type": type_
    }

    query = urllib.parse.urlencode(data_)
    url = 'https://api.commerce.naver.com/external/v1/oauth2/token?' + query
    res = requests.post(url=url, headers=headers)
    res_data = res.json()

    while True:
        if 'access_token' in res_data:
            token = res_data['access_token']
            return token
        else:
            print(f'[{res_data}] 토큰 요청 실패')
            time.sleep(1)

# 스마트스토어 API 코드, 주문 목록 가져오는 함수
def get_new_order_list():

    headers = {'Authorization': token}
    url = 'https://api.commerce.naver.com/external/v1/pay-order/seller/product-orders/last-changed-statuses'

    now = datetime.now()
    # before_date = now - timedelta(hours=3) #3시간전
    # before_date = now - timedelta(seconds=10) #10초전
    # before_date = now - timedelta(minutes=10) #10분전
    before_date = now - timedelta(days=1)  # 이틀전
    iosFormat = before_date.astimezone().isoformat()
    # iosFormat_to = now.astimezone().isoformat()


    #하루 전 to now로 세팅하기 (24시간이 최대)
    params = {
        'lastChangedFrom': iosFormat,  # 조회시작일시
        # 'lastChangedTo' : iosFormat_to,
        'lastChangedType': 'PAYED',  # 최종변경구분(PAYED : 결제완료, DISPATCHED : 발송처리)
    }

    res = requests.get(url=url, headers=headers, params=params)
    res_data = res.json()

    if 'data' not in res_data:  # 조회된 정보가 없을 경우 data키 없음
        print('주문 내역 없음')
        return False

    data_list = res_data['data']['lastChangeStatuses']

    productOrderIds = []
    for data in data_list:
        productOrderIds.append(data['productOrderId'])
    print(f'주문 조회 성공: {productOrderIds}')
    return productOrderIds

# 스마트스토어 API 코드, 상품주문번호를 매개변수로 주문상세정보 가져오는 함수.
# 상품ID와 옵션코드를 리턴해 줌.
def get_order_detail(productOrderId):
    import datetime

    headers = {'Authorization': token}
    url = 'https://api.commerce.naver.com/external/v1/pay-order/seller/product-orders/query'

    params = {
        'productOrderIds': [productOrderId]
    }

    res = requests.post(url=url, headers=headers, json=params)
    res_data = res.json()

    # print(res_data)
    if 'data' not in res_data:
        return False

    # 상품ID와 옵션코드를 리턴해줌
    productId = ""
    optionManageCode = ""
    for data in res_data['data']:
        for d in data.keys():
            for d2 in data[d].keys():
                if d2 == 'productId':
                    # print(f'productId : {data[d][d2]}')
                    productId = data[d][d2]
                if d2 == 'optionManageCode':
                    # print(f'optionManageCode : {data[d][d2]}')
                    optionManageCode = data[d][d2]
                # print(f'{d2} : {data[d][d2]}')
            return productId, optionManageCode

# get_order_detail 함수가 리턴해 준 상품ID와 옵션코드로 엑셀에서 상응하는 메일 내용 및 제목 가져오기.
def get_mail_details():
    for i in range(2, load_mail_ws.max_row + 1):
        if str(load_mail_ws['A' + str(i)].value) == str(productId):
            if str(load_mail_ws['B' + str(i)].value) == str(optionManageCode):
                title = load_mail_ws['C' + str(i)].value
                contents = load_mail_ws['D' + str(i)].value
                return title, contents

# 스마트스토어 API 코드, 발송처리하는 코드
def item_sending(productOrderIds, dispatchDate):
    headers = {
        'Authorization': token,
        'content-type': "application/json"
    }
    url = 'https://api.commerce.naver.com/external/v1/pay-order/seller/product-orders/dispatch'

    params = {
        'dispatchProductOrders': [{
            'productOrderId': str(productOrderIds[0]),
            'deliveryMethod': 'NOTHING',
            'dispatchDate': dispatchDate,  # 배송일
        }]}

    res = requests.post(url=url, headers=headers, json=params)
    res_data = res.json()

    # print(res_data)
    if 'data' not in res_data:
        return False

# SMTP로 구매자에게 이메일 보내기
# get_mail_details 함수로 가져온 메일 제목과 내용, 그리고 구매자email을 매개변수로함.
def email(title, contents, email):
    # 세션 생성
    s = smtplib.SMTP('smtp.gmail.com', 587)
    # TLS 보안 시작
    s.starttls()
    # 로그인 인증
    s.login(my_email, my_email_pwd)
    # 보낼 메시지 설정
    msg = MIMEMultipart('alternative')
    msg['Subject'] = title
    msg['From'] = my_email
    msg['To'] = email

    #html 메일도 보낼 수 있게 세팅
    html = MIMEText(contents, 'html')
    msg.attach(html)

    text = MIMEText(contents, 'plain')
    msg.attach(text)

    s.sendmail(my_email , email, msg.as_string())

# 구매자에게 메일을 보내고 발송처리까지 마무리 후에 프로그램 사용자에게 확인 메일을 보내는 함수,
def confirm_email(productOrderId):
    # 세션 생성
    s = smtplib.SMTP('smtp.gmail.com', 587)
    # TLS 보안 시작
    s.starttls()
    # 로그인 인증
    s.login(my_email, my_email_pwd)
    # 보낼 메시지 설정
    msg = MIMEMultipart('alternative')
    msg['Subject'] = f'{productOrderId} / 메일이 정상적으로 발송되었습니다.'
    msg['From'] = my_email
    msg['To'] = my_email

    # html = MIMEText(f'{productOrderId} / 메일이 정상적으로 발송되었습니다.', 'html')
    # msg.attach(html)

    text = MIMEText(f'{productOrderId} / 메일이 정상적으로 발송되었습니다.', 'plain')
    msg.attach(text)

    s.sendmail(my_email, my_email, msg.as_string())

# 스마트 스토어 로그인 함수.
def login():
    to_login_btn = driver.find_element(By.CLASS_NAME, 'btn-login')
    to_login_btn.click()

    id_input = driver.find_elements(By.CLASS_NAME, 'Login_ipt__cPqIR')[0]
    id_input.click()
    pyperclip.copy(naver_id)
    id_input.send_keys(Keys.CONTROL, 'v')
    time.sleep(1)

    pwd_input = driver.find_elements(By.CLASS_NAME, 'Login_ipt__cPqIR')[1]
    pwd_input.click()
    pyperclip.copy(naver_pwd)
    pwd_input.send_keys(Keys.CONTROL, 'v')
    time.sleep(1)

    login_btn = driver.find_element(By.CLASS_NAME, 'Button_btn__enzXE')
    login_btn.click()

# 스마트스토어에 뜨는 팝업을 없애는 함수.
# JS를 건드려 DOM을 조작해 팝업 및 백그라운드를 없앰.
def clean_up_popup():
    driver.execute_script("""
       var elms = document.getElementsByClassName("modal fade seller-layer-modal modal-transparent has-close-check-box in");
    Array.from(elms).forEach(function(element) {
        element.parentNode.removeChild(element);
    });
    """)

    time.sleep(1)

    driver.execute_script("""
       var elms = document.getElementsByClassName("modal-backdrop fade in");
    Array.from(elms).forEach(function(element) {
        element.parentNode.removeChild(element);
    });
    """)

    time.sleep(1)

    driver.execute_script("""
       var elms = document.getElementsByClassName("modal-content");
    Array.from(elms).forEach(function(element) {
        element.parentNode.removeChild(element);
    });
    """)

    time.sleep(1)

    driver.execute_script("""
       var elms = document.getElementsByClassName("modal-content");
    Array.from(elms).forEach(function(element) {
        element.parentNode.removeChild(element);
    });
    """)


#######################################################################################

# 메인 룹
while True:
    #토큰 3시간 정도 유효하므로 2시간 반마다 새로.
    token = get_token(client_id=client_id, client_secret=client_secret)
    #새로 로그인?############################

    max_time_end = time.time() + (60 * 150)
    while True:
        if time.time() > max_time_end:
            break
        productOrderIds = get_new_order_list()
        #주문 내역이 있을 시에만
        if productOrderIds != False:
            #로그인
            driver.get('https://sell.smartstore.naver.com/#/home/about')
            time.sleep(2)
            login()
            time.sleep(3)

            #대시보드페이지 및 팝업제거
            driver.get('https://sell.smartstore.naver.com/#/home/dashboard')
            time.sleep(6)
            clean_up_popup()

            #스토어 이동
            to_store = driver.find_element(By.XPATH, '//*[@id="_gnb_nav"]/ul/li[2]/a')
            to_store.click()
            time.sleep(3)

            store_text = driver.find_element(By.XPATH, f"//span[contains(text(), '{naver_store}')]")
            store_text.click()
            driver.execute_script('arguments[0].click()', store_text)
            time.sleep(3)

            # 발주주문페이지
            driver.get('https://sell.smartstore.naver.com/#/naverpay/sale/delivery')
            time.sleep(1)
            driver.get('https://sell.smartstore.naver.com/#/naverpay/sale/delivery')
            time.sleep(1)
            driver.get('https://sell.smartstore.naver.com/#/naverpay/sale/delivery')

            for productOrderId in productOrderIds:
                # 조회버튼 클릭
                productId = get_order_detail(productOrderId)[0]
                optionManageCode = get_order_detail(productOrderId)[1]

                content = driver.find_element(By.TAG_NAME, "iframe")
                driver.switch_to.frame(content)
                time.sleep(1)
                # 조회버튼 클릭
                checking_btn = driver.find_element(By.XPATH,'//*[@id="__app_root__"]/div/div[2]/div[4]/div[2]/button[2]')
                checking_btn.click()
                time.sleep(1)

                # 상품 디테일 클릭 및 구매자아이디 크롤링
                detail_a = driver.find_elements(By.XPATH, f"//a[contains(text(), '{productOrderId}')]")
                detail_a[0].click()
                time.sleep(2)
                driver.switch_to.window(driver.window_handles[-1])
                time.sleep(1)
                id = driver.find_element(By.XPATH, f"//th[contains(text(), '구매자 ID')]/following-sibling::td").text
                client_email = f'{id}@naver.com'

                # 이메일 발송
                title = get_mail_details()[0]
                contents = get_mail_details()[1]
                email(title, contents, client_email)
                confirm_email(productOrderId)

                #발송처리
                now = datetime.now()
                iosFormat = now.astimezone().isoformat()
                timestamp = str(int((time.time() - 3) * 1000))

                item_sending(productOrderIds=[productOrderId], dispatchDate=timestamp)

                driver.close()
                time.sleep(1)
                driver.switch_to.window(driver.window_handles[0])


        driver.refresh()
        time.sleep(30)





#
# 애플리케이션 ID
# 4vFMKHfU0srLmvW9UAqoSr
# 시크릿
# $2a$04$Ir43wxNaJX4nNwVVIeFBt.

