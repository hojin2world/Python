# -*- coding: utf-8 -*-
import locale
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
    # 오늘 날짜 가져오기
    today_date = datetime.now().strftime("%Y%m%d")
    file_name = f"{today_date}_진도율미만.xlsx"

    print("파일 이름:", file_name)

    # 다운로드 디렉토리 설정
    home_directory = os.path.expanduser('~')
    download_directory = os.path.join(home_directory, 'Downloads', 'python')

    # 디렉토리가 있는지 확인
    if not os.path.exists(download_directory):
        # 디렉토리가 없다면 생성
        os.makedirs(download_directory)

    # 이제 다운로드 디렉토리는 'Downloads' 내부의 'python' 디렉토리를 가리킴

    print(download_directory)

    # Chrome 옵션 설정하여 다운로드 디렉토리 및 원하는 파일 이름 설정
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        'download.default_directory': download_directory,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'download.default_filename': file_name,
        'detach': True,
        'download.mime_types': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    })

    # 구성된 옵션으로 Chrome WebDriver 초기화
    driver = webdriver.Chrome(options=chrome_options)
    return driver

# 다운로드 디렉토리 설정 및 원하는 파일 이름 지정
home_directory = os.path.expanduser('~')
download_directory = os.path.join(home_directory, 'Downloads', 'python')

# 구성된 WebDriver 인스턴스 가져오기
driver = get_configured_driver(download_directory)

files = os.listdir(download_directory)


# URL로 이동
url = 'https://www.con.or.kr/'
driver.get(url)
driver.maximize_window()
time.sleep(2)


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
driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[1]').click()
time.sleep(1)
driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[2]/div[2]/div[1]/a').click()
time.sleep(0.5)

selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[2]/div/select'))
selectbox.select_by_value('2025')
time.sleep(1)

input_field = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[4]/input')
input_field.click()
# 로케일 한국어로 설정
#locale

# 현재 날짜 가져오기
current_date = datetime.now()

# 월 숫자를 한국어 월 이름으로 매핑하는 딕셔너리
month_names_korean = {
    1: "01월",
    2: "02월",
    3: "03월",
    4: "04월",
    5: "05월",
    6: "06월",
    7: "07월",
    8: "08월",
    9: "09월",
    10: "10월",
    11: "11월",
    12: "12월"
}

# 월과 일 추출
month_number = current_date.month
month = month_names_korean[month_number]  # 한국어 월 이름
day = current_date.day

# 날짜에 접미사를 추가하는 함수
def add_suffix(day):
    if 4 <= day <= 20 or 24 <= day <= 30:
        return f"{day:02d}일"
    else:
        return f"{day:02d}일"

# 토요일과 일요일을 제외한 다음 평일을 반환하는 함수
def get_next_weekday(current_date):
    while True:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5:  # 월요일부터 금요일까지
            return current_date

# 한국 시간대(UTC+9)의 현재 날짜를 가져옵니다.
current_date_korean_time = datetime.utcnow() + timedelta(hours=9)

# 토요일과 일요일을 제외하고 3일을 추가한 이후의 날짜를 가져옵니다.
next_weekday = get_next_weekday(current_date_korean_time)
days_added = 0
while days_added < 2:  # 주말을 제외하고 2일을 추가합니다.
    next_weekday += timedelta(days=1)
    if next_weekday.weekday() < 5:  # 평일인지 확인합니다.
        days_added += 1

# 월과 일을 추출합니다.
month = next_weekday.strftime("%m월")  # 한국어로 월 이름 표시
day = next_weekday.day

# 날짜를 "월 일[접미사]" 형식으로 포맷합니다.
formatted_date = f"{month} {add_suffix(day)}"

print(formatted_date)
time.sleep(1)
input_field.send_keys(formatted_date)


selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[6]/td[2]/div/div[2]/select'))
selectbox.select_by_value('90')
time.sleep(1)

selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[1]/td[6]/div/select'))
selectbox.select_by_value('2,3,5') #혼합
time.sleep(1)

driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div[2]/div[2]').click()
time.sleep(5)



driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[2]/div/div[2]/div[4]').click()
time.sleep(5)
driver.find_element(By.XPATH,'//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[3]/div[3]').click()
time.sleep(5)

# 파일이 다운로드될 때까지 기다림
wait = WebDriverWait(driver, 10)

# 다운로드된 파일이 실제로 저장된 경로를 얻어옴
downloaded_file_path = os.path.join(download_directory, 'download.xls')

print(downloaded_file_path)

# 현재 날짜 가져오기
today_date = datetime.now().strftime("%Y%m%d")

df = pd.read_excel(downloaded_file_path)
new_file_path = downloaded_file_path.replace('download.xls', f'{today_date}_진도율90미만.xlsx')
df.to_excel(new_file_path, index=False)

# 변경 전 파일 경로
old_file_path = os.path.join(download_directory, 'download.xls')

# 변경 후 파일 경로
new_file_path = os.path.join(download_directory, f'{today_date}_진도율90미만.xlsx')

# 변경 후 파일이 존재하는지 확인
if os.path.exists(new_file_path):
    # 파일이 존재한다면, 변경 전 파일 삭제
    os.remove(old_file_path)
else:
    # 파일이 존재하지 않는다면, 에러 메시지 출력
    print(f"Error: {new_file_path} 파일이 존재하지 않습니다.")

# 파일이 다운로드될 때까지 기다림
# timeout = 10
# while not os.path.exists(downloaded_file_path) and timeout > 0:
#    time.sleep(1)
#    timeout -= 1

# 파일이 존재하지 않으면 TimeoutException 발생
# if not os.path.exists(downloaded_file_path):
#    raise TimeoutException(f"File '{downloaded_file_path}' was not downloaded within {timeout} seconds")

print("진도율미달 파일 다운로드 및 형식변환 완료")

################################### 뿌리오 형식으로 변경###################################


# 엑셀 파일 경로
excel_file = os.path.join(download_directory, f'{today_date}_진도율90미만.xlsx')
# excel_file = rf'C:\Users\user\Downloads\python\{today_date}_진도율90미만(row).xlsx'




# 엑셀 파일에서 읽어올 열 이름 지정
columns_to_read = ['휴대전화번호', '이름', '패키지명', '학습 종료일']

# 주말을 제외하고 3일 후의 날짜를 가져오는 함수입니다.
def get_next_weekday(current_date):
    next_day = current_date
    for _ in range(3):
        next_day += timedelta(days=1)
        while next_day.weekday() >= 5:  # If it is Saturday or Sunday
            next_day += timedelta(days=1)  # Skip Saturday and Sunday
    return next_day

# 현재 시간을 한국 시간으로 설정합니다
current_date_korean_time = datetime.utcnow() + timedelta(hours=9)

# 주말을 제외하고 3일 후의 날짜를 가져옵니다.
next_weekday = get_next_weekday(current_date_korean_time)

# Convert to Korean day of the week
weekday_dict = {
    0: '월요일',
    1: '화요일',
    2: '수요일',
    3: '목요일',
    4: '금요일',
    5: '토요일',
    6: '일요일'
}
next_weekday_korean = weekday_dict[next_weekday.weekday()]

print("주말 제외 3일 뒤 요일:", next_weekday_korean)
# 엑셀 파일에서 열 읽기
df = pd.read_excel(excel_file, usecols=columns_to_read, dtype="str")

# 열의 순서 변경
df = df[['휴대전화번호', '이름', '패키지명', '학습 종료일']]

# '요일' 열 추가
df.insert(df.columns.get_loc('학습 종료일') + 1, '요일', next_weekday_korean)

# 열 이름 변경
df.columns = ['수신자 번호(숫자, 공백, 하이픈(-)만)', '[*1*]', '[*2*]', '[*3*]','[*4*]']
print(df.columns)
# 데이터프레임 확인
print(df)
print(df.columns)

# 오늘의 날짜를 문자열로 생성 (형식: 년도-월-일)
today_date = datetime.now().strftime("%Y-%m-%d")

# 새로운 엑셀 파일에 저장
#new_excel_file = r'C:\Users\user\Downloads\python\{}_뿌리오_진도율미달.xlsx'.format(today_date)
new_excel_file = rf'{download_directory}\{today_date}_뿌리오_진도율미달.xlsx'
# 새로운 엑셀 파일 경로 및 파일 이름
df.to_excel(new_excel_file, index=False)  # index=False로 설정하여 인덱스 열을 포함하지 않음

print("진도율미달 파일 뿌리오 변환 완료")

#########################################비즈뿌리오 로그인#########################################


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


# 로그인 후에는 Selenium을 사용하여 웹 사이트에서 추가 작업을 계속할 수 있습니다.

# 메세지 전송 버튼 클릭
driver.find_element(By.XPATH, '//*[@id="header"]/div[1]/div[1]/ul/li[2]/a').click()
time.sleep(1)

#일반 클릭
driver.find_element(By.XPATH, '//*[@id="container"]/div[2]/div[1]/fieldset/div[4]/div[1]/div[2]/table/tbody/tr[1]/td/ul/li[2]/label[1]').click()
time.sleep(1)

# 수신번호별 문구 클릭
driver.find_element(By.XPATH, '//*[@id="container"]/div[2]/div[1]/fieldset/div[4]/div[1]/div[2]/table/tbody/tr[3]/td/ul/li[2]').click()
time.sleep(1)

# 메세지 내용 입력란 선택
text_area = driver.find_element(By.XPATH, '//*[@id="messageContentArea"]')
text_area.click()
time.sleep(1)

# 입력할 텍스트 설정
text_to_add = """[건설산업교육원] 온라인강의 수강 독려 안내

안녕하세요, [*1*]님!
온라인강의 수강 독려 안내 드립니다.

- 과정명 : [*2*]
- 집체교육일 : [*3*] ([*4*])

현재 온라인강의 진도율이 수료기준(90%) 미달입니다.
(*90% 미달 시 수료 불가)

현 시점 기준 집체교육 참석일까지 수료조건을 모두 만족할 수 있는지 일정 확인하시기 바랍니다.

※ 수료조건을 만족하지 않는 경우, 해당 일정으로 수료 불가합니다. (교육일정 변경하여 이어서 진행)


▷교육일정 변경 방법
나의 강의실 → [교육일정 변경] 클릭

con.or.kr
1522-2938"""

# 텍스트 입력란에 텍스트 추가
text_area.send_keys(text_to_add)
time.sleep(1)

# 현재 날짜 가져오기
today_date = datetime.today().strftime("%Y-%m-%d")

# 파일 입력란 선택
file_input = driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
time.sleep(5)

# 파일 경로 설정
file_path = rf"C:\Users\user\Downloads\python\{today_date}_뿌리오_진도율미달.xlsx"
print(file_input)
print(file_path)

# 파일 업로드
file_input.send_keys(file_path)
time.sleep(3)

# 메시지 전송 버튼에 포커스 설정 후 탭 키 입력
element = driver.find_element(By.XPATH, '//*[@id="messageSendSubmit"]')
driver.execute_script("arguments[0].focus();", element)
element.send_keys(Keys.TAB)

print("진도율미달 파일 업로드 완료")