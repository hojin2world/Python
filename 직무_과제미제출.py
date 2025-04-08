from dotenv import load_dotenv
import os
from login_module import get_login_credentials
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
    today_date = datetime.now().strftime("%Y%m%d")
    file_name = f"{today_date}_과제미제출.xlsx"

    print("File name:", file_name)

    # 다운로드 디렉토리 설정
    home_directory = os.path.expanduser('~')
    download_directory = os.path.join(home_directory, 'Downloads', 'python')

    if not os.path.exists(download_directory):
        os.makedirs(download_directory)

    print(download_directory)

    # Chrome 옵션 설정
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        'download.default_directory': download_directory,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'download.default_filename': file_name,
        'detach': True,
        'download.mime_types': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    })

    driver = webdriver.Chrome(options=chrome_options)
    return driver

def get_next_weekday(current_date):
    while True:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5:  # 월요일부터 금요일까지
            return current_date

# 다운로드 디렉토리 설정
home_directory = os.path.expanduser('~')
download_directory = os.path.join(home_directory, 'Downloads', 'python')

# WebDriver 인스턴스 가져오기
driver = get_configured_driver(download_directory)

# 로그인 정보 받아오기
login_credentials = get_login_credentials()

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
    
    try:
        # 로그인 시도
        id_field = driver.find_element(By.XPATH, '//*[@id="id"]')
        id_field.click()
        id_field.send_keys(login_credentials['con_username'])
        time.sleep(1)
        
        pw_field = driver.find_element(By.XPATH, '//*[@id="pw"]')
        pw_field.click()
        pw_field.send_keys(login_credentials['con_password'])
        time.sleep(1)
        
        # 로그인 버튼 클릭
        driver.find_element(By.XPATH, '/html/body/div[3]/div/div[1]/div/div[2]/div[1]/button').click()
        time.sleep(2)
        
    except Exception as e:
        print(f"로그인 중 오류 발생: {str(e)}")
        raise

# 로그인 한 번만 실행
login(driver)
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="popup_layout_list"]/div/div[2]/div[3]/div[2]').click()
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[1]').click()  #학습자 평가 관리 클릭
time.sleep(1)
driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[2]/div[2]/div[1]/a').click() #학습자 모니터링 클릭
time.sleep(0.5)

selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[2]/div/select'))
selectbox.select_by_value('2025')
time.sleep(1)

learningST = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[1]/td[6]/div/select'))
learningST.select_by_value('2,3,5') #혼합
time.sleep(1)

input_field = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[4]/input')
input_field.click()

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

print("과제 미제출"+formatted_date)
time.sleep(1)
input_field.send_keys(formatted_date)


selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[6]/td[4]/div/div[2]/select'))  #과제 미제출
selectbox.select_by_value('0')
time.sleep(1)
driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div[2]/div[2]').click()
time.sleep(5)
driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[2]/div/div[2]/div[4]').click()
time.sleep(10)
driver.find_element(By.XPATH,'//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[3]/div[3]').click()
time.sleep(5)

print(formatted_date + "엑셀 다운로드 완료" )
downloaded_file_path = os.path.join(download_directory, 'download.xls')

# 현재 날짜 가져오기
today_date = datetime.now().strftime("%Y%m%d")

df = pd.read_excel(downloaded_file_path)
new_file_path = downloaded_file_path.replace('download.xls', f'{today_date}_과제미제출.xlsx')
df.to_excel(new_file_path, index=False)

# 변경 전 파일 경로
old_file_path = os.path.join(download_directory, 'download.xls')

# 변경 후 파일 경로
new_file_path = os.path.join(download_directory, f'{today_date}_과제미제출.xlsx')

# 변경 후 파일이 존재하는지 확인
if os.path.exists(new_file_path):
    # 파일이 존재한다면, 변경 전 파일 삭제
    os.remove(old_file_path)
else:
    # 파일이 존재하지 않는다면, 에러 메시지 출력
    print(f"Error: {new_file_path} 파일이 존재하지 않습니다.")
    

print("과제미제출 파일 다운로드 완료")

################################### 뿌리오 형식으로 변경###################################

# 엑셀 파일 경로
excel_file = os.path.join(download_directory, f'{today_date}_과제미제출.xlsx')


# 엑셀 파일에서 읽어올 열 이름 지정
columns_to_read = ['휴대전화번호', '이름', '패키지명', '학습 종료일']

# 엑셀 파일에서 열 읽기
df = pd.read_excel(excel_file, usecols=columns_to_read, dtype="str")

# 열의 순서 변경
df = df[['휴대전화번호', '이름', '패키지명', '학습 종료일']]

df['학습 종료일'] = pd.to_datetime(df['학습 종료일'])
df['학습 종료일'] -= timedelta(days=2)
df['학습 종료일'] = df['학습 종료일'].dt.strftime("%Y.%m.%d")

# 열 이름 변경
df.columns = ['수신자 번호(숫자, 공백, 하이픈(-)만)', '[*1*]', '[*2*]', '[*3*]']
print(df.columns)
# 데이터프레임 확인
print(df)
print(df.columns)

# 오늘의 날짜를 문자열로 생성 (형식: 년도-월-일)
today_date = datetime.now().strftime("%Y-%m-%d")

# 새로운 엑셀 파일에 저장
new_excel_file = r'C:\Users\user\Downloads\python\{}_뿌리오_과제미제출.xlsx'.format(today_date)

# 새로운 엑셀 파일 경로 및 파일 이름
df.to_excel(new_excel_file, index=False)  # index=False로 설정하여 인덱스 열을 포함하지 않음



# 엑셀 파일에서 읽어올 열 이름 지정
columns_to_read = ['휴대전화번호', '이름', '패키지명', '학습 종료일']


print("과제미제출 파일 변환 완료")

####################################비즈 뿌리오 로그인 및 문자 전송 #####################################

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

# PPURIO 로그인 부분 수정
def ppurio_login(driver):
    username_input = driver.find_element(By.ID, 'bizwebHeaderUserId')
    username_input.send_keys(login_credentials['ppurio_username'])

    password_input = driver.find_element(By.ID, 'bizwebHeaderUserPwd')
    password_input.send_keys(login_credentials['ppurio_password'])

    login_button = driver.find_element(By.XPATH, '//*[@id="bizwebHeaderBtnLogin"]')
    login_button.click()

    # 비밀번호 만료 연장 처리
    try:
        element = driver.find_element(By.XPATH, '//*[@id="bizwebBtnWebPasswdExpiredateDelay"]')
        element.click()
        print("비밀번호 연장 처리")
    except NoSuchElementException:
        print("비밀번호 연장 처리 없음")

    time.sleep(2)

ppurio_login(driver)

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
text_to_add = """[건설산업교육원] 과제미제출 안내

안녕하세요, [*1*]님!
과제제출 독려 안내 드립니다.

- 과정명 : [*2*]
- 제출마감일 : [*3*] (집체교육일 2일 전)

현재 과제 미제출 상태입니다. 기한 내 과제를 제출해주세요.
(*미제출 시 수료 불가)

※ 수료조건을 만족하지 않는 경우, 해당 일정으로 수료 불가합니다. (교육일정 변경하여 이어서 진행)


▷과제제출 방법
나의 강의실 → 진행과정 → 강의실 입장 → 학습활동 및 평가 → 과제(리포트) → [제출하기] 클릭

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
file_path = rf"C:\Users\user\Downloads\python\{today_date}_뿌리오_과제미제출.xlsx"
print(file_input)
print(file_path)

# 파일 업로드
file_input.send_keys(file_path)
time.sleep(3)

# 메시지 전송 버튼에 포커스 설정 후 탭 키 입력
element = driver.find_element(By.XPATH, '//*[@id="messageSendSubmit"]')
driver.execute_script("arguments[0].focus();", element)
element.send_keys(Keys.TAB)

print("과제미제출 파일 업로드 완료.")