from dotenv import load_dotenv
import os

load_dotenv()

# -*- coding: utf-8 -*-
import locale
import time
import os
from datetime import datetime, timedelta
from telnetlib import EC

from selenium.common import TimeoutException
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
from datetime import datetime, timedelta

from selenium.webdriver.support.wait import WebDriverWait

def get_configured_driver(download_directory):
    # 오늘 날짜 가져오기
    today_date = datetime.now().strftime("%Y%m%d")
    file_name = f"{today_date}_과제미제출_연장.xlsx"

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
        
    except Exception as e:
        print(f"로그인 중 오류 발생: {str(e)}")
        raise

# 로그인 한 번만 실행
login(driver)





time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="popup_layout_list"]/div/div[2]/div[3]/div[2]').click()
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[1]').click()
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[2]/div[2]/div[1]/a').click()
time.sleep(2)

selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[2]/div/select'))
selectbox.select_by_value('2025')
time.sleep(2)

input_field = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[4]/input')
input_field.click()


current_date = datetime.now()

# 월 번호를 한국어 월 이름으로 매핑하는 사전
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
        if current_date.weekday() < 5: # 월요일부터 금요일까지
            return current_date

# 오늘 날짜를 가져옴
current_date = datetime.now()

# 다음 평일의 날짜를 가져옴 (토요일과 일요일을 제외)
next_day = get_next_weekday(current_date)

# 만약 내일이 토요일인 경우, 월요일로 넘어가도록 하기 위해 하루를 추가
if next_day.weekday() == 5: # 토요일
    next_day += timedelta(days=2) # 토요일과 일요일을 건너뛰기 위해 2일 추가
elif next_day.weekday() == 6: # 일요일
    next_day += timedelta(days=1) # 일요일을 건너뛰기 위해 1일 추가

# 월과 일을 추출
month_number = next_day.month
month = month_names_korean[month_number] # 한국어로 월 이름 표시
day = next_day.day

# "월 일[접미사]" 형식으로 날짜를 포맷
formatted_date = f"{month} {add_suffix(day)}"
print(formatted_date)


time.sleep(1)
input_field.send_keys(formatted_date)

selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[6]/td[4]/div/div[2]/select'))  #과제 미제출
selectbox.select_by_value('0')
time.sleep(1)
learningST = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[1]/td[6]/div/select'))
learningST.select_by_value('2,3,5') #혼합
time.sleep(1)
driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div[2]/div[2]').click()
time.sleep(5)
driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[2]/div/div[2]/div[4]').click()
time.sleep(10)
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
new_file_path = downloaded_file_path.replace('download.xls', f'{today_date}_과제미제출_연장.xlsx')
df.to_excel(new_file_path, index=False)

# 변경 전 파일 경로
old_file_path = os.path.join(download_directory, 'download.xls')

# 변경 후 파일 경로
new_file_path = os.path.join(download_directory, f'{today_date}_과제미제출_연장.xlsx')

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

##########################################엑셀 다운 후 패키지


# 엑셀 파일 경로
excel_file = os.path.join(download_directory, f'{today_date}_과제미제출_연장.xlsx')
# excel_file = rf'C:\Users\user\Downloads\python\{today_date}_진도율90미만(row).xlsx'


# 엑셀 파일에서 읽어올 열 이름 지정
columns_to_read = ['기수번호', '패키지명']

df = pd.read_excel(excel_file, usecols=columns_to_read, dtype="str")

print(df.columns)
print('중복 제거 전 데이터프레임 확인')
print(df)
# 패키지명 중복제거
df.drop_duplicates(subset=['패키지명'], inplace=True)
print('중복 제거 후 데이터프레임 확인')
print(df.columns)
# 데이터프레임 확인
print(df)

#엑셀의 열 확인
num_rows = df.shape[1]

# 오늘의 날짜를 문자열로 생성 (형식: 년도-월-일)
today_date = datetime.now().strftime("%Y-%m-%d")

# 새로운 엑셀 파일에 저장
#new_excel_file = r'C:\Users\user\Downloads\python\{}_뿌리오_진도율미달.xlsx'.format(today_date)
new_excel_file = rf'{download_directory}\{today_date}_과제미제출_연장_변환.xlsx'
# 새로운 엑셀 파일 경로 및 파일 이름
df.to_excel(new_excel_file, index=False)  # index=False로 설정하여 인덱스 열을 포함하지 않음

print("진도율미달 파일 변환 완료")

driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]').click()
time.sleep(1)

for index, row in df.iterrows():



    current_date = datetime.now()

    # 월 번호를 한국어 월 이름으로 매핑하는 사전
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


    # 오늘 날짜를 가져옴
    current_date = datetime.now()

    # 다음 평일의 날짜를 가져옴 (토요일과 일요일을 제외)
    next_day = get_next_weekday(current_date)

    # 만약 내일이 토요일인 경우, 월요일로 넘어가도록 하기 위해 하루를 추가
    if next_day.weekday() == 5:  # 토요일
        next_day += timedelta(days=2)  # 토요일과 일요일을 건너뛰기 위해 2일 추가
    elif next_day.weekday() == 6:  # 일요일
        next_day += timedelta(days=1)  # 일요일을 건너뛰기 위해 1일 추가

    # 월과 일을 추출
    month_number = next_day.month
    month = month_names_korean[month_number]  # 한국어로 월 이름 표시
    day = next_day.day

    # "월 일[접미사]" 형식으로 날짜를 포맷
    formatted_date = f"{month} {add_suffix(day)}"
    print(formatted_date)

    time.sleep(2)



    # driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[3]/div[2]').click()
    # time.sleep(1)
    # driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]').click()
    # time.sleep(1)
    # driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]').click()
    # time.sleep(1)
    # 클릭 전에 WebDriverWait을 사용하여 해당 엘리먼트가 다시 나타날 때까지 기다림
    try:
        # 고정된 요소를 매번 클릭
        fixed_element = driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]')
        fixed_element.click()
        time.sleep(1)  # 클릭 후 잠깐 대기

        fixed_element = driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]')
        fixed_element.click()
        time.sleep(1)  # 클릭 후 잠깐 대기

        driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[2]/div[2]/div[1]').click()
        time.sleep(2)  # 클릭 후 잠깐 대기

    except Exception as e:
        print(f"An error occurred: {e}")
        break  # 에러 발생 시 루프 중지

    # 엑셀 파일 경로
    excel_file = os.path.join(download_directory, f'{today_date}_과제미제출_연장_변환.xlsx')

    # 엑셀 파일을 pandas DataFrame으로 읽어오기
    df = pd.read_excel(excel_file)

    # Get the second value in column A
    # Get the value from the first column
    value_column_A = row.iloc[0]  # Using iloc for positional indexing

    # Get the value from the second column
    value_column_B = row.iloc[1]  # Using iloc for positional indexing


    print("Value from column A:", value_column_A)
    print("Value from column B:", value_column_B)

    selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[2]/div/select')) #연도 선택
    selectbox.select_by_value('2025')
    time.sleep(1)
    learningST = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[1]/td[4]/div/select')) #학습방식
    learningST.select_by_value('2')  # 블렌디드러닝
    time.sleep(1)

    input_field  = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[5]/td[4]/input') #기수번호
    input_field.clear()
    time.sleep(2)
    input_field.send_keys(value_column_A)
    time.sleep(2)

    gisu_input = driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[5]/td[2]/input')  # 기수이름
    time.sleep(2)
    gisu_input.send_keys(formatted_date)
    time.sleep(2)

    input_field2  = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[6]/td[2]/input') #패키지명
    input_field2.clear()
    time.sleep(2)
    input_field2.send_keys(value_column_B)
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div[2]/div[2]').click() #검색 버튼
    time.sleep(2)


    div_element = driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[3]/div[1]/div/table/tbody/tr[2]/td[15]')
    # Get the text value of the div element
    div_value = div_element.text

    # Parse the date string into a datetime object
    date_time_object = datetime.strptime(div_value, "%Y-%m-%d %H:%M:%S")
    date_time_object = date_time_object.replace(hour=23, minute=59, second=59)
    # Subtract one day from the datetime object
    previous_date = date_time_object - timedelta(days=1)
    date_str = previous_date.strftime("%Y-%m-%d %H:%M:%S")
    print(date_str)

    driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[3]/div[1]/div/table/tbody/tr[2]/td[9]').click()
    time.sleep(2)


    driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[2]/div[9]/div[1]/div').click() #평가 출제 및 채점 클릭
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[2]/div[9]/div[2]/div[2]').click() #평가출제 클릭
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[2]/div/div/table/tbody/tr[2]/td[3]').click() #평가명 클릭
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[1]/div[2]').click()  # 고급설정 클릭
    time.sleep(2)

    checkbox_xpath = '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]/td[2]/div[1]/img'
    checkbox_element = driver.find_element(By.XPATH, checkbox_xpath)

    # 이미지의 src 속성 가져오기
    src_value = checkbox_element.get_attribute('src')

    # 체크박스 상태 확인
    if 'checkbox_pressed.png' in src_value:  # 체크된 상태를 나타내는 이미지 파일명
        print("체크박스가 이미 체크되어 있습니다.")
    else:
        print("체크박스가 체크되어 있지 않습니다. 체크를 시도합니다.")
        print("결과리뷰 체크박스로 이동합니다.")
        # 체크박스의 부모 요소나 상위 div 클릭
        checkbox_parent_xpath = '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]/td[2]/div[1]'
        checkbox_parent_element = driver.find_element(By.XPATH, checkbox_parent_xpath)
        checkbox_parent_element.click()  # 체크박스 클릭 시도

    submit_field = driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[2]/td[2]/div[1]/input')  # 연장제출 반영비율
    submit_field.clear()
    #driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[2]/td[2]/div[1]/input') #연장제출 시 전수반영비율
    submit_field.send_keys(100)
    time.sleep(2)

    #결과 리뷰 체크박스
    checkbox_xpath = '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[3]/td[2]/div[1]/img'
    checkbox_element = driver.find_element(By.XPATH, checkbox_xpath)

    # 이미지의 src 속성 가져오기
    src_value = checkbox_element.get_attribute('src')

    # 체크박스 상태 확인
    if 'checkbox_pressed.png' in src_value:  # 체크된 상태를 나타내는 이미지 파일명
        print("체크박스가 이미 체크되어 있습니다.")
    else:
        print("체크박스가 체크되어 있지 않습니다. 체크를 시도합니다.")
        print("결과리뷰 체크박스로 이동합니다.")
        # 체크박스의 부모 요소나 상위 div 클릭
        checkbox_parent_xpath = '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[3]/td[2]/div[1]'
        checkbox_parent_element = driver.find_element(By.XPATH, checkbox_parent_xpath)
        checkbox_parent_element.click()  # 체크박스 클릭 시도

    # #결과리뷰 체크박스 확인
    # checkbox_image_xpath = '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[3]/td[2]/div[1]/img'
    #
    # try:
    #     # 주어진 XPath로 이미지 요소 찾기
    #     checkbox_image = WebDriverWait(driver, 10).until(
    #         EC.presence_of_element_located((By.XPATH, checkbox_image_xpath)))
    #
    #     # 이미지 요소의 'src' 속성 값 확인
    #     src_value = checkbox_image.get_attribute('src')
    #
    #     # 'src' 속성이 'checkbox_pressed.png'인지 확인
    #     if 'checkbox_pressed.png' in src_value:
    #         print("체크박스가 체크되어 있습니다.")
    #     else:
    #         print("체크박스가 체크되어 있지 않습니다. 다음 작업을 진행합니다.")
    #         # 체크박스를 체크하는 추가 작업
    # except Exception as e:
    #     print(f"체크박스 상태를 확인할 수 없거나 오류가 발생했습니다: {e}")


    #current_date_time = datetime.now()
    #formatted_date_time = current_date_time.strftime("%Y-%m-%d %H:%M:%S")
    deadline_field = driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]/td[4]/div/div/div[3]/div/input')
    deadline_field.clear()
    deadline_field.send_keys(date_str)  # 연장제출기한 클릭
    #input_field.send_keys(date_str)
    time.sleep(2)
    print("연장제출기한:" + date_str)



    review_end_day = driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[3]/td[4]/div/div/div[3]/div/input')  # 결과 리뷰 기간
    review_end_day.clear()
    review_end_day.send_keys(date_str)
    #driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[3]/td[4]/div/div/div[3]/div/input').send_keys(date_str)  # 결과 리뷰 기간
    #input_field.send_keys(date_str)
    time.sleep(2)
    print("결과리뷰기간 종료일자:" + date_str)

    #연장제출기간
    date_time_object = datetime.strptime(div_value, "%Y-%m-%d %H:%M:%S")
    date_time_object = date_time_object.replace(hour=00, minute=00, second=00)
    # Subtract one day from the datetime object
    previous_zero_date = date_time_object - timedelta(days=1)
    date_zero_date = previous_zero_date.strftime("%Y-%m-%d %H:%M:%S")
    review_start_day = driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]/td[4]/div/div/div[1]/div/input')  # 연장제출기간
    review_start_day.clear()
    review_start_day.send_keys(date_zero_date)
    #input_field.send_keys(date_zero_date)
    time.sleep(2)
    print("연장제출기간 시작일자:" + date_zero_date)

    driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div/div[2]/div[2]/div[2]').click() #저장버튼
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[5]/div[1]').click() # 저장확인
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[5]/div[2]').click() #미리보기 취소
    time.sleep(2)

print("과제_연장_완료")

