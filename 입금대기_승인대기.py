from dotenv import load_dotenv
import os

load_dotenv()

# -*- coding: utf-8 -*-
import locale
import sys
import time
import os
from datetime import datetime, timedelta
from telnetlib import EC

from selenium.common import TimeoutException, NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
from selenium.webdriver.common.keys import Keys
import xlrd
import os, shutil
import pandas as pd
from datetime import datetime
from datetime import datetime, timedelta

from selenium.webdriver.support.wait import WebDriverWait

def get_configured_driver(download_directory):
    # Get today's date
    today_date = datetime.now().strftime("%Y%m%d")


    # 다운로드 디렉토리 설정
    home_directory = os.path.expanduser('~')
    download_directory = os.path.join(home_directory, 'Downloads', 'python')

    # Check if the directory exists
    if not os.path.exists(download_directory):
        # Create the directory if it doesn't exist
        os.makedirs(download_directory)

    # Now download_directory points to the 'python' directory inside 'Downloads'

    print(download_directory)

    # Configure Chrome options to change the download directory and set the desired file name
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        'download.default_directory': download_directory,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'detach': True,
        'download.mime_types': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    })

    # Initialize Chrome WebDriver with configured options
    driver = webdriver.Chrome(options=chrome_options)
    return driver

# Set the download directory and specify desired file name
home_directory = os.path.expanduser('~')
download_directory = os.path.join(home_directory, 'Downloads', 'python')


# Get the configured WebDriver instance

driver = get_configured_driver(download_directory)

files = os.listdir(download_directory)


# Navigate to the URL
url = 'https://www.con.or.kr/'
driver.get(url)
driver.maximize_window()

# 로그인 함수 정의 (한 번만 실행)
def login(driver):
    # URL로 이동
    url = 'https://www.con.or.kr/'
    driver.get(url)
    driver.maximize_window()
    time.sleep(2)
    
    # 팝업 처리
    try:
        driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div/label/span').click()
        driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[2]/button').click()
    except:
        print("팝업이 없거나 이미 처리되었습니다.")
    
    # 환경 변수에서 로그인 정보 가져오기
    USERNAME = os.getenv('LOGIN_USERNAME')
    PASSWORD = os.getenv('LOGIN_PASSWORD')
    
    if not USERNAME or not PASSWORD:
        raise ValueError("로그인 정보가 환경 변수에 설정되지 않았습니다.")
    
    try:
        # 로그인 시도
        id_field = driver.find_element(By.XPATH, '//*[@id="id"]')
        id_field.click()
        id_field.send_keys(USERNAME)
        time.sleep(1)
        
        pw_field = driver.find_element(By.XPATH, '//*[@id="pw"]')
        pw_field.click()
        pw_field.send_keys(PASSWORD)
        time.sleep(1)
        
        # 로그인 버튼 클릭
        driver.find_element(By.XPATH, '/html/body/div[3]/div/div[1]/div/div[2]/div[1]/button').click()
        time.sleep(2)

        time.sleep(2)
        driver.find_element(By.XPATH,'//*[@id="popup_layout_list"]/div/div[2]/div[3]/div[2]').click()
        time.sleep(2)
        
    except Exception as e:
        print(f"로그인 중 오류 발생: {str(e)}")
        raise

# 로그인 한 번만 실행
login(driver)

driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[5]/div[1]').click()     #결제내역 클릭
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[5]/div[2]/div[1]/div[1]/a').click() #주문내역 조회 클릭
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div/table/tbody/tr[2]/td[4]/div[1]/div[2]').click() #결제대기클릭
time.sleep(2)
selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div/table/tbody/tr[5]/td[2]/div[1]/select')) #기간 - 교육종료일 선택
selectbox.select_by_value('5')
time.sleep(2)


# Function to add suffix to day
def add_suffix(day):
    if 4 <= day <= 20 or 24 <= day <= 30:
        return f"{day}th"
    else:
        return f"{day}st" if day == 1 else f"{day}nd" if day == 2 else f"{day}rd"

# Function to return the next weekday excluding Saturday and Sunday
def get_next_weekday(current_date):
    while True:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5: # Monday to Friday
            return current_date

# Get the current date in Korean time zone (UTC+9).
current_date_korean_time = datetime.utcnow() + timedelta(hours=9)

# Get the date after adding 3 days, excluding Saturday and Sunday.
next_weekday = get_next_weekday(current_date_korean_time)
days_added = 0
while days_added < 2: # Add 2 days excluding weekends.
    next_weekday += timedelta(days=1)
    if next_weekday.weekday() < 5: # Check if it is a weekday.
        days_added += 1

# Extract the month and day.
month = str(next_weekday.month).zfill(2)
day = str(next_weekday.day).zfill(2)

# Format the date in "year-month-day" format.
formatted_date = f"{next_weekday.year}-{month}-{day}"


start_date_field = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div/table/tbody/tr[5]/td[2]/div[2]/div[1]/div/input')
start_date_field.click()

print("교육 종료일 시작일자 : " + formatted_date)
time.sleep(1)
start_date_field.send_keys(formatted_date)


end_date_field = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div/table/tbody/tr[5]/td[2]/div[2]/div[3]/div/input')
end_date_field.click()

print("교육 종료일 종료일자 : " + formatted_date)
time.sleep(1)
end_date_field.send_keys(formatted_date) # 교육 종료일 입력

driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div/div/div[2]/div[2]').click() #검색 클릭
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[3]/div/div[2]/div[1]').click() #다운로드 버튼 클릭
#다운받을 항목 검사
existence_exceldownload_xpath = '//*[@id="popup_layout_list"]/div/div[2]/div[5]/div[1]'
time.sleep(2)
try:
    element = driver.find_element(By.XPATH, existence_exceldownload_xpath)
    element.click()
    print("다운로드 받을 항목이 없습니다.")
    print("종료합니다.")
    sys.exit()
except NoSuchElementException:
    print("다운로드 받겠습니다.")
time.sleep(2)
time.sleep(5)
driver.find_element(By.XPATH,'//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[3]/div[3]').click() #엑셀 다운로드 버튼 클릭
time.sleep(5)

print(formatted_date + "엑셀 다운로드 완료")
downloaded_file_path = os.path.join(download_directory, '주문내역 조회.xls')

# 현재 날짜 가져오기
today_date = datetime.now().strftime("%Y%m%d")

df = pd.read_excel(downloaded_file_path)
new_file_path = downloaded_file_path.replace('주문내역 조회.xls', f'{today_date}_주문내역.xlsx')
df.to_excel(new_file_path, index=False)

# 변경 전 파일 경로
old_file_path = os.path.join(download_directory, '주문내역 조회.xls')

# 변경 후 파일 경로
new_file_path = os.path.join(download_directory, f'{today_date}_주문내역.xlsx')

# 변경 후 파일이 존재하는지 확인
if os.path.exists(new_file_path):
    # 파일이 존재한다면, 변경 전 파일 삭제
    os.remove(old_file_path)
else:
    # 파일이 존재하지 않는다면, 에러 메시지 출력
    print(f"Error: {new_file_path} 파일이 존재하지 않습니다.")

print("주문내역 파일 다운로드 완료")


################################### 뿌리오 형식으로 변경###################################
today_date = datetime.now().strftime("%Y%m%d")
# 엑셀 파일 경로
excel_file = os.path.join(download_directory, f'{today_date}_주문내역.xlsx')

# 엑셀 파일에서 읽어올 열 이름 지정
columns_to_read = ['휴대폰번호', '이름', '패키지명', '교육종료일', '주문금액', '결제상태' ,'관련기관']





# Filter data based on payment status for 'Awaiting Deposit'
df_deposit = df[df['결제상태'] == '입금대기']

# Check if there are records in 'Awaiting Deposit'
if df_deposit.empty:
    print("입금대기건이 없습니다.")
else:
    # 필터링된 데이터프레임에서 필요한 열만 선택
    df_deposit = df_deposit[['휴대폰번호', '이름', '패키지명', '교육종료일', '주문금액']]

    # 열 이름 변경
    df_deposit.columns = ['수신자 번호(숫자, 공백, 하이픈(-)만)', '[*1*]', '[*2*]', '[*3*]', '[*4*]']

    # 새로운 엑셀 파일 경로와 이름 지정
    new_excel_file_deposit = os.path.join(download_directory, f"{today_date}_입금대기.xlsx")

    # 필터링된 데이터프레임을 엑셀 파일로 저장
    df_deposit.to_excel(new_excel_file_deposit, index=False)
    print("입금대기 엑셀 파일이 생성되었습니다.")
    print("엑셀 파일 경로:", new_excel_file_deposit)

# df_deposit = df_deposit[['휴대폰번호', '이름', '패키지명', '교육종료일', '주문금액']]
#
# # Rename columns
# df_deposit.columns = ['수신자 번호(숫자, 공백, 하이픈(-)만)', '[*1*]', '[*2*]', '[*3*]', '[*4*]']
#
# # New Excel file path and file name for 'Awaiting Deposit'
# new_excel_file_deposit = os.path.join(download_directory, f"{today_date}_입금대기.xlsx")


# Save DataFrame to Excel file for 'Awaiting Deposit'
# df_deposit.to_excel(new_excel_file_deposit, index=False)


df_approval = df[df['결제상태'] == '승인대기']
if df_approval.empty:
    print("승인대기건이 없습니다.")
else:
    df_approval = df_approval[['휴대폰번호', '이름', '패키지명', '교육종료일', '관련기관']]
    df_approval.columns = ['수신자 번호(숫자, 공백, 하이픈(-)만)', '[*1*]', '[*2*]', '[*3*]', '[*4*]']
    new_excel_file_approval = os.path.join(download_directory, f"{today_date}_승인대기.xlsx")
    # 필터링된 데이터프레임을 엑셀 파일로 저장
    df_approval.to_excel(new_excel_file_approval, index=False)
    print("승인대기 엑셀 파일이 생성되었습니다.")
    print("엑셀 파일 경로:", new_excel_file_approval)

#Filter data based on payment status for 'Awaiting Approval'
# df_approval = df[df['결제상태'] == '승인대기']
# df_approval = df_approval[['휴대폰번호', '이름', '패키지명', '교육종료일', '관련기관']]
#
# # Rename columns for 'Awaiting Approval'
# df_approval.columns = ['수신자 번호(숫자, 공백, 하이픈(-)만)', '[*1*]', '[*2*]', '[*3*]', '[*4*]']
#
# # New Excel file path and file name for 'Awaiting Approval'
# new_excel_file_approval = os.path.join(download_directory, f"{today_date}_승인대기.xlsx")
#
# # Save DataFrame to Excel file for 'Awaiting Approval'
# df_approval.to_excel(new_excel_file_approval, index=False)
# print("승인대기 엑셀")
# print(df_approval.columns)

####################################비즈 뿌리오 로그인 및 문자 전송 #####################################

new_excel_file_deposit = os.path.join(download_directory, f"{today_date}_입금대기.xlsx")

# Check if the file exists
if os.path.exists(new_excel_file_deposit):
    today_date = datetime.now().strftime("%Y%m%d")
    print("파일이 존재합니다. 추가 처리를 진행합니다.")
    print("입금대기 파일이 존재합니다.")
    print("입금대기 비즈뿌리오를 실행합니다.")

    # ChromeOptions 객체 생성
    chrome_options = webdriver.ChromeOptions()

    # detach 옵션을 True로 설정
    chrome_options.add_experimental_option('detach', True)

    # 설정된 옵션으로 Chrome WebDriver 초기화
    driver = webdriver.Chrome(options=chrome_options)

    # URL로 이동
    new_url = 'https://www.bizppurio.com/'
    driver.get(new_url)
    driver.maximize_window()

    # 예시: 로그인
    # 로그인 요소를 찾아서 클릭하거나 입력
    PPURIO_USERNAME = os.getenv('PPURIO_LOGIN_USERNAME')
    PPURIO_PASSWORD = os.getenv('PPURIO_LOGIN_PASSWORD')
    username_input = driver.find_element(By.ID, 'bizwebHeaderUserId')
    username_input.send_keys(PPURIO_USERNAME)  # 아이디 입력

    password_input = driver.find_element(By.ID, 'bizwebHeaderUserPwd')
    password_input.send_keys(PPURIO_PASSWORD)  # 비밀번호 입력

    login_button = driver.find_element(By.XPATH, '//*[@id="bizwebHeaderBtnLogin"]')
    login_button.click()
    session_cookies = driver.get_cookies()

    # 페이지 로딩을 위해 충분한 시간을 주거나 필요한 요소에 대한 대기 조건을 추가
    time.sleep(2)  # 필요한 경우 시간 조정

    # 비밀번호 만료 연장 버튼이 존재하면 클릭, 존재하지 않으면 패스
    xpath = '//*[@id="bizwebBtnWebPasswdExpiredateDelay"]'

    try:
        element = driver.find_element(By.XPATH, xpath)
        element.click()
        print("비밀번호 연장 처리")
    except NoSuchElementException:
        print("비밀번호 연장 처리 없음")

    time.sleep(2)


    # 로그인 후에는 Selenium을 사용하여 웹 사이트에서 추가 작업을 계속할 수 있습니다.

    # 메세지 전송 버튼 클릭
    driver.find_element(By.XPATH, '//*[@id="header"]/div[1]/div[1]/ul/li[2]/a').click()
    time.sleep(1)

    #일반 클릭
    driver.find_element(By.XPATH, '//*[@id="container"]/div[2]/div[1]/fieldset/div[4]/div[1]/div[2]/table/tbody/tr[1]/td/ul/li[2]/label[1]').click()
    time.sleep(1)

    # 수신번호별 문구 클릭
    driver.find_element(By.XPATH, '//*[@id="container"]/div[2]/div[1]/fieldset/div[4]/div[1]/div[2]/table/tbody/tr[3]/td/ul/li[2]/label').click()
    time.sleep(1)

    # 메세지 내용 입력란 선택
    text_area = driver.find_element(By.XPATH, '//*[@id="messageContentArea"]')
    text_area.click()
    time.sleep(1)

    # 입력할 텍스트 설정
    text_to_add = """[건설산업교육원] 입금대기 안내
    
    안녕하세요, [*1*]님!
    
    - 과정명 : [*2*]
    - 교육종료일 : [*3*] (변경 가능)
    
    신청하신 교육과정이 입금대기 상태입니다.
    미임급 시 교육을 진행할 수 없습니다.
    
    - 입금 금액 : [*4*]원
    
    ▷ 가상계좌번호 확인
    나의 강의실 → 대기과정 → [신청정보]
    
    ▷ 신청 취소
    나의 강의실 → 결제내역 → [취소]
    
    con.or.kr
    1522-2938"""

    # 텍스트 입력란에 텍스트 추가
    text_area.send_keys(text_to_add)
    time.sleep(1)

    # 현재 날짜 가져오기
    today_date = datetime.today().strftime("%Y%m%d")

    # 파일 입력란 선택
    file_input = driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
    time.sleep(5)

    # 파일 경로 설정
    file_path = os.path.join(download_directory, "{}_입금대기.xlsx".format(today_date))
    print(file_input)
    print(file_path)

    # 파일 업로드
    file_input.send_keys(file_path)
    time.sleep(3)

    # 메시지 전송 버튼에 포커스 설정 후 탭 키 입력
    element = driver.find_element(By.XPATH, '//*[@id="messageSendSubmit"]')
    driver.execute_script("arguments[0].focus();", element)
    element.send_keys(Keys.TAB)

    print("입금대기 파일 업로드 완료.")
else:
    print("입금대기 파일 없음.")
    print("입금대기 문자전송 종료.")

##################################### 승인대기 문자 ########################################

new_excel_file_approval = os.path.join(download_directory, f"{today_date}_승인대기.xlsx")
if os.path.exists(new_excel_file_approval):
    today_date = datetime.now().strftime("%Y%m%d")
    print("파일이 존재합니다. 추가 처리를 진행합니다.")
    print("입금대기 파일이 존재합니다.")
    print("입금대기 비즈뿌리오를 실행합니다.")
    # ChromeOptions 객체 생성
    chrome_options = webdriver.ChromeOptions()

    # detach 옵션을 True로 설정
    chrome_options.add_experimental_option('detach', True)

    # 설정된 옵션으로 Chrome WebDriver 초기화
    driver = webdriver.Chrome(options=chrome_options)

    # URL로 이동
    new_url = 'https://www.bizppurio.com/'
    driver.get(new_url)
    driver.maximize_window()

    # 예시: 로그인
    # 로그인 요소를 찾아서 클릭하거나 입력
    PPURIO_USERNAME = os.getenv('PPURIO_LOGIN_USERNAME')
    PPURIO_PASSWORD = os.getenv('PPURIO_LOGIN_PASSWORD')
    username_input = driver.find_element(By.ID, 'bizwebHeaderUserId')
    username_input.send_keys(PPURIO_USERNAME)  # 아이디 입력

    password_input = driver.find_element(By.ID, 'bizwebHeaderUserPwd')
    password_input.send_keys(PPURIO_PASSWORD)  # 비밀번호 입력

    login_button = driver.find_element(By.XPATH, '//*[@id="bizwebHeaderBtnLogin"]')
    login_button.click()
    session_cookies = driver.get_cookies()

    # 페이지 로딩을 위해 충분한 시간을 주거나 필요한 요소에 대한 대기 조건을 추가
    time.sleep(2)  # 필요한 경우 시간 조정

    # 비밀번호 만료 연장 버튼이 존재하면 클릭, 존재하지 않으면 패스
    xpath = '//*[@id="bizwebBtnWebPasswdExpiredateDelay"]'

    try:
        element = driver.find_element(By.XPATH, xpath)
        element.click()
        print("비밀번호 연장 처리")
    except NoSuchElementException:
        print("비밀번호 연장 처리 없음")

    time.sleep(2)

    # 로그인 후에는 Selenium을 사용하여 웹 사이트에서 추가 작업을 계속할 수 있습니다.

    # 메세지 전송 버튼 클릭
    driver.find_element(By.XPATH, '//*[@id="header"]/div[1]/div[1]/ul/li[2]/a').click()
    time.sleep(1)

    # 수신번호별 문구 클릭
    driver.find_element(By.XPATH, '//*[@id="container"]/div[2]/div[1]/fieldset/div[4]/div[1]/div[2]/table/tbody/tr[3]/td/ul/li[2]/label').click()
    time.sleep(1)

    # 메세지 내용 입력란 선택
    text_area = driver.find_element(By.XPATH, '//*[@id="messageContentArea"]')
    text_area.click()
    time.sleep(1)

    # 입력할 텍스트 설정
    text_to_add = """[건설산업교육원] 승인대기 안내
    
    안녕하세요, [*1*]님!
    
    - 과정명 : [*2*]
    - 교육종료일 : [*3*] (변경 가능)
    - 소속명 : [*4*]
    
    신청하신 교육과정이 '승인대기' 상태임을 알려드립니다.
    
    교육을 진행하시는 경우,
    소속 담당자님께 문의하여 '승인 요청'하시기 바랍니다.
    미승인 시 교육을 진행하실 수 없습니다.
    
    ▷ 신청내역 확인
    나의 강의실 → 결제내역
    
    con.or.kr
    1522-2938"""

    # 텍스트 입력란에 텍스트 추가
    text_area.send_keys(text_to_add)
    time.sleep(1)

    # 현재 날짜 가져오기
    today_date = datetime.today().strftime("%Y%m%d")

    # 파일 입력란 선택
    file_input = driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
    time.sleep(5)

    # 파일 경로 설정
    file_path = os.path.join(download_directory, "{}_승인대기.xlsx".format(today_date))
    print(file_input)
    print(file_path)

    # 파일 업로드
    file_input.send_keys(file_path)
    time.sleep(3)

    # 메시지 전송 버튼에 포커스 설정 후 탭 키 입력
    element = driver.find_element(By.XPATH, '//*[@id="messageSendSubmit"]')
    driver.execute_script("arguments[0].focus();", element)
    element.send_keys(Keys.TAB)

    print("승인대기 파일 업로드 완료.")
else:
    print("승인대기 파일 없음.")
    print("승인대기 문자전송 종료.")