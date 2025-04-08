from dotenv import load_dotenv
import os
import configparser
import tkinter as tk
from tkinter import messagebox

load_dotenv()

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

from login_module import get_login_credentials

def get_configured_driver(download_directory):
    # 오늘 날짜 가져오기
    today_date = datetime.now().strftime("%Y%m%d")
    file_name = f"{today_date}_28시간_수료처리.xlsx"

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

def calculate_gisu():
    # 기준일: 2025년 1월 3일 (금요일)
    base_date = datetime(2025, 1, 3)
    
    # 현재 날짜
    current_date = datetime.now()
    
    # 현재 날짜의 요일 (0:월요일, 1:화요일, ..., 4:금요일, ..., 6:일요일)
    current_weekday = current_date.weekday()
    
    # 이번 주 금요일 날짜 계산
    # 현재 요일이 금요일(4)보다 크면 다음 주 금요일, 작으면 이번 주 금요일
    days_until_friday = (4 - current_weekday) % 7
    this_friday = current_date + timedelta(days=days_until_friday)
    
    # 전 주의 금요일 계산 (현재 금요일에서 7일 빼기)
    previous_friday = this_friday - timedelta(days=7)
    
    # 기준일부터 전 주 금요일까지의 주차 차이 계산
    weeks_diff = (previous_friday - base_date).days // 7
    
    # 기수 계산 (1기수부터 시작)
    gisu = weeks_diff + 1
    
    return gisu

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

# 로그인 정보 받아오기
login_credentials = get_login_credentials()

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
driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[1]').click()
time.sleep(1)
driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[2]/div[2]/div[1]/a').click()
time.sleep(0.5)

selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[2]/div/select'))
selectbox.select_by_value('2025')
time.sleep(1)

gisu_name_input_field = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[4]/input')
gisu_name_input_field.click()
time.sleep(1)
gisu_name_input_field.send_keys("28시간")
time.sleep(1)

gisu_number_field = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[6]/input')
gisu_number_field.click()
time.sleep(1)
gisu_number = calculate_gisu()
print(f"학습자 평가 관리 현재 기수: {gisu_number}")
gisu_number_field.send_keys(gisu_number)
time.sleep(1)

driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div[2]/div[2]').click()
time.sleep(5)



driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[2]/div/div[2]/div[4]').click()
time.sleep(7)
driver.find_element(By.XPATH,'//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[3]/div[3]').click()
time.sleep(7)

# 파일이 다운로드될 때까지 기다림
wait = WebDriverWait(driver, 10)

# 다운로드된 파일이 실제로 저장된 경로를 얻어옴
downloaded_file_path = os.path.join(download_directory, 'download.xls')

print(downloaded_file_path)

# 현재 날짜 가져오기
today_date = datetime.now().strftime("%Y%m%d")

df = pd.read_excel(downloaded_file_path)

# 'AE' 열의 값이 '수료'인 행만 추출하여 새로운 데이터프레임 생성
completed_df = df[df['수료 여부'] == '수료']

new_file_path = downloaded_file_path.replace('download.xls', f'{today_date}_28시간_수료처리.xlsx')
df.to_excel(new_file_path, index=False)

# 변경 전 파일 경로
old_file_path = os.path.join(download_directory, 'download.xls')

# 변경 후 파일 경로
new_file_path = os.path.join(download_directory, f'{today_date}_28시간_수료처리.xlsx')

# 변경 후 파일이 존재하는지 확인
if os.path.exists(new_file_path):
    # 파일이 존재한다면, 변경 전 파일 삭제
    os.remove(old_file_path)
else:
    # 파일이 존재하지 않는다면, 에러 메시지 출력
    print(f"Error: {new_file_path} 파일이 존재하지 않습니다.")
    
    
# 엑셀 파일 경로
excel_file = os.path.join(download_directory, f'{today_date}_28시간_수료처리.xlsx')
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
today_date = datetime.now().strftime("%Y%m%d")

# 새로운 엑셀 파일에 저장
#new_excel_file = r'C:\Users\user\Downloads\python\{}_뿌리오_진도율미달.xlsx'.format(today_date)
new_excel_file = rf'{download_directory}\{today_date}_28시간_수료처리_변환.xlsx'
# 새로운 엑셀 파일 경로 및 파일 이름
df.to_excel(new_excel_file, index=False)  # index=False로 설정하여 인덱스 열을 포함하지 않음

print("수료 파일 변환 완료")

########################
driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]').click() #과정개설관리 클릭
time.sleep(0.5)
driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[2]/div[2]/div[1]').click() #패키지개설관리 클릭
time.sleep(0.5)

year_selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[2]/div/select'))  # 연도 선택
year_selectbox.select_by_value('2025')
time.sleep(1)
# learningSTY = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[1]/td[4]/div/select'))  # 학습방식
# learningSTY.select_by_value('2')  # 블렌디드러닝
# time.sleep(1)

gisu_name_input = driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[5]/td[2]/input')  # 기수이름
time.sleep(2)
gisu_name_input.send_keys("28시간")
time.sleep(2)

# 기존 코드에서 기수 관련 부분을 수정
gisu_number = calculate_gisu()
print(f"패키지 개설 현재 기수: {gisu_number}")

# 기수번호 입력 부분 수정
gisu_num_field = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[5]/td[4]/input')
gisu_num_field.clear()
time.sleep(2)
gisu_num_field.send_keys(f"{gisu_number}기")
time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div[2]/div[2]').click()  # 검색 버튼
time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[2]/div[2]/div[3]').click()  # 엑셀 다운로드 버튼 클릭
time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[3]/div[3]').click()  #  다운로드 버튼 클릭
time.sleep(5)

# 파일이 다운로드될 때까지 기다림
wait = WebDriverWait(driver, 10)

# 다운로드된 파일이 실제로 저장된 경로를 얻어옴
downloaded_file_path2 = os.path.join(download_directory, 'download.xls')

print(downloaded_file_path2)

# 현재 날짜 가져오기
today_date = datetime.now().strftime("%Y%m%d")

df = pd.read_excel(downloaded_file_path2)


new_file_path = downloaded_file_path2.replace('download.xls', f'{today_date}_28시간_패키지.xlsx')
df.to_excel(new_file_path, index=False)

# 변경 전 파일 경로
old_file_path = os.path.join(download_directory, 'download.xls')

# 변경 후 파일 경로
new_file_path = os.path.join(download_directory, f'{today_date}_28시간_패키지.xlsx')

# 변경 후 파일이 존재하는지 확인
if os.path.exists(new_file_path):
    # 파일이 존재한다면, 변경 전 파일 삭제
    os.remove(old_file_path)
else:
    # 파일이 존재하지 않는다면, 에러 메시지 출력
    print(f"Error: {new_file_path} 파일이 존재하지 않습니다.")

# 엑셀 파일 경로
excel_file = os.path.join(download_directory, f'{today_date}_28시간_패키지.xlsx')



######################## 패키지 다운로드 끝 ######################








######################## 수료처리 파일과 대조 시작#####################

# 수료처리변환 파일 경로
evaluation_file = os.path.join(download_directory, f'{today_date}_28시간_수료처리_변환.xlsx')
# 패키지 파일 경로
package_file = os.path.join(download_directory, f'{today_date}_28시간_패키지.xlsx')    

# 원본과 사본 엑셀 파일을 읽습니다.
evaluation_df = pd.read_excel(evaluation_file)
package_df = pd.read_excel(package_file)

# 원본의 패키지명과 사본의 과목명 컬럼을 비교하여 일치하는 경우 사본의 기수명 컬럼을 원본에 추가
# 여기서 '패키지명'은 원본의 컬럼명, '과목명'과 '기수명'은 사본의 컬럼명이라고 가정
merged_df = pd.merge(evaluation_df, package_df[['과목명', '기수명']], how='left', left_on='패키지명', right_on='과목명')

# 필요 없는 과목명 컬럼을 제거합니다.
merged_df.drop(columns=['과목명'], inplace=True)

# 결과를 download_directory에 저장합니다.
output_path = os.path.join(download_directory, f'{today_date}_28시간_검색변환.xlsx')
merged_df.to_excel(output_path, index=False)

print(f"File saved to {output_path}")


######################  수료처리 파일과 대조 끝########################


# 엑셀 파일 경로
excel_file = os.path.join(download_directory, f'{today_date}_28시간_검색변환.xlsx')

# 엑셀 파일을 pandas DataFrame으로 읽어오기
df = pd.read_excel(excel_file)

# "기수번호" 열의 첫 번째 값 가져오기
first_gisu_number = df.loc[0, "기수번호"]


for index, row in df.iterrows():

    # URL로 이동
    
    time.sleep(2)
    driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]').click()
    time.sleep(2)
    driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]').click()
    time.sleep(2)
    driver.find_element(By.XPATH,'//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[2]/div[2]/div[1]').click()
    time.sleep(2)

    # 한국 시간대(UTC+9)의 현재 날짜를 가져옵니다.
    current_date_korean_time = datetime.utcnow() + timedelta(hours=9)


    # 주말을 제외하고 어제 날짜를 반환하는 함수
    def get_previous_weekday(current_date):
        while True:
            current_date -= timedelta(days=1)
            if current_date.weekday() < 5:  # 월요일부터 금요일까지
                return current_date


    # 어제 날짜를 가져옵니다.
    previous_weekday = get_previous_weekday(current_date_korean_time)

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


    # 날짜에 접미사를 추가하는 함수
    def add_suffix(day):
        return f"{day:02d}일"


    # 월과 일을 추출합니다.
    month_number = previous_weekday.month
    month = month_names_korean[month_number]  # 한국어 월 이름
    day = previous_weekday.day

    # 날짜를 "월 일[접미사]" 형식으로 포맷합니다.
    formatted_date = f"{month} {add_suffix(day)}"

    print(formatted_date)

    time.sleep(1)



    # 엑셀 파일 경로
    excel_file = os.path.join(download_directory, f'{today_date}_28시간_검색변환.xlsx')

    # 엑셀 파일을 pandas DataFrame으로 읽어오기
    df = pd.read_excel(excel_file)

    # Get the second value in column A
    # Get the value from the first column
    value_column_A = row.iloc[0]  # Using iloc for positional indexing

    # Get the value from the second column
    value_column_B = row.iloc[1]  # Using iloc for positional indexing

    # Get the value from the second column
    value_column_C = row.iloc[2]  # Using iloc for positional indexing


    print("Value from column A:", value_column_A)
    print("Value from column B:", value_column_B)
    print("Value from column C:", value_column_C)

    selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[4]/td[2]/div/select')) #연도 선택
    selectbox.select_by_value('2025')
    time.sleep(1)
    # learningST = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[1]/td[4]/div/select')) #학습방식
    # learningST.select_by_value('2')  # 블렌디드러닝
    # time.sleep(1)

    gisu_num_field  = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[5]/td[4]/input') #기수번호
    gisu_num_field.clear()
    time.sleep(2)
    gisu_num_field.send_keys(value_column_A)
    time.sleep(2)

    gisu_name_field  = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[5]/td[2]/input')  # 기수이름
    gisu_name_field.clear()
    time.sleep(2)
    gisu_name_field.send_keys(value_column_C)
    time.sleep(2)

    input_field2  = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[6]/td[2]/input') #패키지명
    input_field2.clear()
    time.sleep(2)
    input_field2.send_keys(value_column_B)
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div[2]/div[2]').click() #검색 버튼
    time.sleep(2)

    driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[3]/div[1]/div/table/tbody/tr[2]/td[15]').click()
    time.sleep(2) #패키지 클릭

    driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[2]/div[13]/div[1]/div').click()
    time.sleep(2) #성적관리 클릭

    selectbox = Select(driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div[2]/div[3]/div/div[2]/div[8]/select'))
    selectbox.select_by_value('100')
    time.sleep(2) #100개씩 보기 클릭

    # 해당 요소의 텍스트 추출
    element_text = driver.find_element(By.CSS_SELECTOR, '.triton_flag_radio_button_text.triton_content').text
    # 숫자 부분만 추출
    number_text = ''.join(filter(str.isdigit, element_text))
    # 숫자로 변환
    number = int(number_text)
    print("전체 건수 :" + str(number))


    # 수료 상태 확인 함수
    def check_completion_status(driver, row_number):
        try:
            # XPath를 사용하여 각 행의 '수료' 값 가져오기
            xpath = f'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div[2]/div[4]/div[1]/div/table/tbody/tr[{row_number}]/td[14]/div/div/span'
            element = driver.find_element(By.XPATH, xpath)
            return element.text == "수료"
        except NoSuchElementException:
            # 요소를 찾지 못한 경우 False 반환
            return None


    # 미수료자 정보를 저장할 리스트
    incomplete_learners = []

    # 패키지명 가져오기
    package_name = value_column_B

    # 해당 패키지의 총 수료처리 완료된 건 수를 저장할 변수
    package_completed = 0

    # 모든 행을 확인하며 미수료자 정보 수집


    # XPath에서 행 번호를 변수 i로 사용하도록 수정
    # 요소를 찾지 못할 경우 break를 사용하여 반복 중단
    # 디버깅용 출력 메시지 추가
    # 파일 저장 로직 유지
    #미수료자 정보 저장 구조 변경 (기수번호, 패키지명, 기수명, 사용자명)
    # DataFrame 생성 시 열 순서 지정
    # 엑셀 파일 저장 시 지정된 열 순서 유지
    # 기존 파일에 추가할 때도 열 순서 유지
    for i in range(2, number+2):
        try:
            # 수료 상태 확인
            status_xpath = f'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div[2]/div[4]/div[1]/div/table/tbody/tr[{i}]/td[14]/div/div/span'
            status = driver.find_element(By.XPATH, status_xpath).text
            
            # 수료 상태가 '수료'가 아닌 경우
            if status != "수료":
                # 사용자명 가져오기 (td[3]에서)
                name_xpath = f'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div[2]/div[4]/div[1]/div/table/tbody/tr[{i}]/td[3]'
                name = driver.find_element(By.XPATH, name_xpath).text
                
                print(f"미수료자 발견 - 행 번호: {i}, 이름: {name}")  # 디버깅용 출력
                
                incomplete_learners.append({
                    '기수번호': value_column_A,
                    '패키지명': value_column_B,
                    '기수명': value_column_C,
                    '사용자명': name
                })
            else:
                # 수료된 행의 체크박스 클릭
                checkbox_xpath = f'//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div[2]/div[4]/div[1]/div/table/tbody/tr[{i}]/td[1]/div/img'
                driver.find_element(By.XPATH, checkbox_xpath).click()
                time.sleep(0.1)
                package_completed += 1
            
        except NoSuchElementException as e:
            print(f"Row {i}: 요소를 찾을 수 없음 - {e}")
            break  # 더 이상 행이 없으면 반복 중단
        except Exception as e:
            print(f"Row {i}: 처리 중 오류 발생 - {e}")
            continue

    # 미수료자 정보 출력 및 저장
    if incomplete_learners:
        print(f"\n{package_name} 패키지의 미수료자 목록:")
        for learner in incomplete_learners:
            print(f"사용자명: {learner['사용자명']}")
        
        try:
            # 저장 경로 지정
            save_dir = r'C:\Users\user\Downloads\python'
            
            # 디렉토리가 없으면 생성
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
                print(f"디렉토리 생성됨: {save_dir}")
            
            # 현재 날짜를 파일명에 포함
            today_date = datetime.now().strftime("%Y%m%d")
            file_name = f'{today_date}_28시간_미수료자목록.xlsx'
            incomplete_file_path = os.path.join(save_dir, file_name)
            
            print(f"저장할 파일 경로: {incomplete_file_path}")
            
            # 미수료자 정보를 DataFrame으로 변환
            incomplete_df = pd.DataFrame(incomplete_learners)
            
            # 열 순서 지정
            column_order = ['기수번호', '패키지명', '기수명', '사용자명']
            incomplete_df = incomplete_df[column_order]
            
            # 파일 저장
            if os.path.exists(incomplete_file_path):
                print("기존 파일이 존재합니다. 데이터를 추가합니다.")
                existing_df = pd.read_excel(incomplete_file_path)
                updated_df = pd.concat([existing_df, incomplete_df], ignore_index=True)
                updated_df.to_excel(incomplete_file_path, index=False)
            else:
                print("새 파일을 생성합니다.")
                incomplete_df.to_excel(incomplete_file_path, index=False)
            
            print(f"파일이 성공적으로 저장되었습니다: {incomplete_file_path}")
            
        except Exception as e:
            print(f"파일 저장 중 오류 발생: {e}")

    print(f"\n{package_name} 패키지의 총 {package_completed}건의 수료처리가 완료되었습니다.")

    # 수료처리 버튼 클릭 및 확인
    driver.find_element(By.XPATH,
                        '//*[@id="wrapper"]/div[1]/div/div/div/div[2]/div[1]/div/div[2]/div[3]/div/div[2]/div[1]').click()
    time.sleep(2)
    driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[5]/div[1]').click()
    time.sleep(2)



################################### 뿌리오 형식으로 변경###################################


# 엑셀 파일 경로 (뿌리오 변환용)
excel_file = os.path.join(download_directory, f'{today_date}_28시간_수료처리.xlsx')

# 파일 존재 여부 확인
if not os.path.exists(excel_file):
    print(f"오류: 파일을 찾을 수 없습니다: {excel_file}")
    print("다운로드 디렉토리의 파일 목록:")
    for file in os.listdir(download_directory):
        print(f"- {file}")
    raise FileNotFoundError(f"파일을 찾을 수 없습니다: {excel_file}")

# 원본 파일에서 필요한 열 읽기
columns_to_read = ['휴대전화번호', '이름', '패키지명', '수료 여부']
df = pd.read_excel(excel_file, usecols=columns_to_read, dtype="str")

# 수료자만 필터링
df = df[df['수료 여부'] == '수료']

# 필요한 열만 선택하고 순서 변경
df = df[['휴대전화번호', '이름', '패키지명']]

# 열 이름 변경
df.columns = ['수신자 번호(숫자, 공백, 하이픈(-)만)', '[*1*]', '[*2*]']

# 데이터프레임 확인
print("변환된 데이터프레임 구조:")
print(df.columns)
print(df.head())

# 오늘의 날짜를 문자열로 생성 (형식: 년도-월-일)
today_date = datetime.now().strftime("%Y%m%d")

# 새로운 엑셀 파일에 저장
new_excel_file = rf'{download_directory}\{today_date}_{first_gisu_number}_수료증발행안내(28시간).xlsx'
df.to_excel(new_excel_file, index=False)

print("수료증 파일 뿌리오 변환 완료")

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

# PPURIO 로그인 부분 수정
username_input = driver.find_element(By.ID, 'bizwebHeaderUserId')
username_input.send_keys(login_credentials['ppurio_username'])

password_input = driver.find_element(By.ID, 'bizwebHeaderUserPwd')
password_input.send_keys(login_credentials['ppurio_password'])

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

# 메세지 유형 클릭
driver.find_element(By.XPATH,
                    '//*[@id="container"]/div[2]/div[1]/fieldset/div[4]/div[1]/div[2]/table/tbody/tr[1]/td/ul/li[2]/label[1]').click()
time.sleep(1)

# 수신번호별 문구 클릭
driver.find_element(By.XPATH,
                    '//*[@id="container"]/div[2]/div[1]/fieldset/div[4]/div[1]/div[2]/table/tbody/tr[3]/td/ul/li[2]/label').click()
time.sleep(1)

# 메세지 내용 입력란 선택
text_area = driver.find_element(By.XPATH, '//*[@id="messageContentArea"]')
text_area.click()
time.sleep(1)

# 입력할 텍스트 설정
text_to_add = """[건설산업교육원] 수료증 발행 안내

안녕하세요, [*1*]님!
수료증이 발행되었습니다.

- 과정명 : [*2*]

※ 협회통보 : 1주일 이내 교육원에서 통보합니다.

※ 수료증은 온라인 교육을 신청한 내용으로 발행됩니다.
(협회에 등록된 이후 변경 불가)


▷ 수료증 확인 방법
나의 강의실 → 완료과정 → [수료증 출력] (다운로드 가능)

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
file_path = os.path.join(download_directory, rf"{today_date}_{first_gisu_number}_수료증발행안내(28시간).xlsx")
print(file_input)
print(file_path)

# 파일 업로드
file_input.send_keys(file_path)
time.sleep(3)

# 메시지 전송 버튼에 포커스 설정 후 탭 키 입력
element = driver.find_element(By.XPATH, '//*[@id="messageSendSubmit"]')
driver.execute_script("arguments[0].focus();", element)
element.send_keys(Keys.TAB)

print("수료증발행 파일 업로드 완료.")
print("수료증발행 종료.")