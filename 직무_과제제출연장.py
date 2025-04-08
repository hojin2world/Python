from dotenv import load_dotenv
import os
import configparser
import tkinter as tk
from tkinter import messagebox

load_dotenv()

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
from datetime import datetime, timedelta

from selenium.webdriver.support.wait import WebDriverWait

from login_module import get_login_credentials

def create_config():
    config = configparser.ConfigParser()
    
    def save_config():
        config['CON'] = {
            'username': con_username_entry.get(),
            'password': con_password_entry.get()
        }
        config['PPURIO'] = {
            'username': ppurio_username_entry.get(),
            'password': ppurio_password_entry.get()
        }
        
        with open('config.ini', 'w', encoding='utf-8') as configfile:
            config.write(configfile)
        messagebox.showinfo("알림", "설정이 저장되었습니다.")
        root.destroy()
    
    root = tk.Tk()
    root.title("로그인 정보 설정")
    
    # CON 로그인 정보
    tk.Label(root, text="CON 로그인 정보", font=('Helvetica', 10, 'bold')).grid(row=0, column=0, columnspan=2, pady=5)
    tk.Label(root, text="아이디:").grid(row=1, column=0, padx=5, pady=2)
    tk.Label(root, text="비밀번호:").grid(row=2, column=0, padx=5, pady=2)
    
    con_username_entry = tk.Entry(root)
    con_password_entry = tk.Entry(root, show="*")
    con_username_entry.grid(row=1, column=1, padx=5, pady=2)
    con_password_entry.grid(row=2, column=1, padx=5, pady=2)
    
    # PPURIO 로그인 정보
    tk.Label(root, text="PPURIO 로그인 정보", font=('Helvetica', 10, 'bold')).grid(row=3, column=0, columnspan=2, pady=5)
    tk.Label(root, text="아이디:").grid(row=4, column=0, padx=5, pady=2)
    tk.Label(root, text="비밀번호:").grid(row=5, column=0, padx=5, pady=2)
    
    ppurio_username_entry = tk.Entry(root)
    ppurio_password_entry = tk.Entry(root, show="*")
    ppurio_username_entry.grid(row=4, column=1, padx=5, pady=2)
    ppurio_password_entry.grid(row=5, column=1, padx=5, pady=2)
    
    # 기존 설정 불러오기
    if os.path.exists('config.ini'):
        config.read('config.ini', encoding='utf-8')
        if 'CON' in config:
            con_username_entry.insert(0, config['CON'].get('username', ''))
            con_password_entry.insert(0, config['CON'].get('password', ''))
        if 'PPURIO' in config:
            ppurio_username_entry.insert(0, config['PPURIO'].get('username', ''))
            ppurio_password_entry.insert(0, config['PPURIO'].get('password', ''))
    
    # 저장 버튼
    tk.Button(root, text="저장", command=save_config).grid(row=6, column=0, columnspan=2, pady=10)
    
    # 창을 화면 중앙에 위치
    root.eval('tk::PlaceWindow . center')
    root.mainloop()

def get_configured_driver(download_directory):
    today_date = datetime.now().strftime("%Y%m%d")
    file_name = f"{today_date}_과제미제출_연장.xlsx"

    print("파일 이름:", file_name)

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

# 로그인 후 팝업 처리
try:
    time.sleep(3)  # 페이지 로드를 위해 대기 시간 증가
    
    # 페이지가 완전히 로드될 때까지 대기
    wait = WebDriverWait(driver, 20)
    wait.until(
        lambda driver: driver.execute_script('return document.readyState') == 'complete'
    )
    
    # 팝업 닫기 버튼을 명시적으로 기다림
    popup_close = wait.until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[3]/div[2]'))
    )
    popup_close.click()
    time.sleep(2)
    
    # 사이드 메뉴 요소가 나타날 때까지 명시적으로 기다림
    side_menu = wait.until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[1]'))
    )
    # JavaScript로 클릭 실행
    driver.execute_script("arguments[0].click();", side_menu)
    time.sleep(2)
    
except TimeoutException as e:
    print(f"요소를 찾을 수 없습니다: {str(e)}")
    print("페이지를 새로고침하고 다시 시도합니다.")
    driver.refresh()
    time.sleep(3)
    # 새로고침 후 다시 시도
    side_menu = wait.until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[9]/div[1]'))
    )
    driver.execute_script("arguments[0].click();", side_menu)
except Exception as e:
    print(f"오류 발생: {str(e)}")

time.sleep(2)

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

gisu_num_input = driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[5]/td[2]/input')  # 기수이름
time.sleep(2)
gisu_num_input.send_keys(formatted_date)
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

# 새 파일 경로 설정
new_file_path2 = downloaded_file_path2.replace('download.xls', f'{today_date}_패키지.xlsx')
df.to_excel(new_file_path2, index=False)

# 원본 파일 삭제
if os.path.exists(new_file_path2):
    os.remove(downloaded_file_path2)
else:
    print(f"Error: {new_file_path2} 파일이 존재하지 않습니다.")

######################## 수료처리 파일과 대조 시작#####################

# 수료처리변환 파일 경로
evaluation_file = os.path.join(download_directory, f'{today_date}_과제미제출_연장.xlsx')
# 패키지 파일 경로
package_file = os.path.join(download_directory, f'{today_date}_패키지.xlsx')

# 원본과 사본 엑셀 파일을 읽습니다.
evaluation_df = pd.read_excel(evaluation_file)
package_df = pd.read_excel(package_file)

# 원본의 패키지명과 사본의 과목명 컬럼을 비교하여 일치하는 경우 사본의 기수명 컬럼을 원본에 추가합니다.
# 여기서 '패키지명'은 원본의 컬럼명, '과목명'과 '기수명'은 사본의 컬럼명이라고 가정합니다.
merged_df = pd.merge(evaluation_df, package_df[['과목명', '기수명']], how='left', left_on='패키지명', right_on='과목명')

# 필요 없는 과목명 컬럼을 제거합니다.
merged_df.drop(columns=['과목명'], inplace=True)

# 결과를 download_directory에 저장합니다.
output_path = os.path.join(download_directory, f'{today_date}_검색변환.xlsx')
merged_df.to_excel(output_path, index=False)

print(f"File saved to {output_path}")


print("과제미제출 파일 다운로드 및 형식변환 완료")

##########################################엑셀 다운 후 패키지

# 엑셀 파일 경로
excel_file = os.path.join(download_directory, f'{today_date}_검색변환.xlsx')
# excel_file = rf'C:\Users\user\Downloads\python\{today_date}_진도율90미만(row).xlsx'

# 엑셀 파일에서 읽어올 열 이름 지정
columns_to_read = ['기수번호', '패키지명','기수명']

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

print("과제미제출 연장 파일 변환 완료")

driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]').click()
time.sleep(1)

fixed_element = driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]')
fixed_element.click()
time.sleep(1)  # 클릭 후 잠깐 대기

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
        time.sleep(1)

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

    # Get the value from the second column
    value_column_C = row.iloc[2]  # Using iloc for positional indexing

    print("Value from column A:", value_column_A)
    #print("Value from column B:", value_column_B)
    #print("Value from column C:", value_column_C)

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

    # gisu_input = driver.find_element(By.XPATH,'//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[5]/td[2]/input')  # 기수이름
    # time.sleep(2)
    # gisu_input.send_keys(formatted_date)
    # time.sleep(2)

    input_field2  = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[6]/td[2]/input') #패키지명
    input_field2.clear()
    time.sleep(2)
    input_field2.send_keys(value_column_B)
    time.sleep(2)

    gisu_name_field  = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[5]/td[2]/input')  # 기수이름
    gisu_name_field.clear()
    time.sleep(2)
    gisu_name_field.send_keys(value_column_C)
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

