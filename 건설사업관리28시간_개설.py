from dotenv import load_dotenv
import os
import configparser
import tkinter as tk
from tkinter import messagebox
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
    # 오늘 날짜 가져오기
    today_date = datetime.now().strftime("%Y%m%d")
    file_name = f"{today_date}_건설사업관리 일반 계속교육(28시간)_개설.xlsx"

    print("파일 이름:", file_name)

    # 다운로드 디렉토리 설정
    home_directory = os.path.expanduser('~')
    download_directory = os.path.join(home_directory, 'Downloads', 'python')

    # 디렉토리가 있는지 확인
    if not os.path.exists(download_directory):
        # 디렉토리가 없다면 생성
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

    # WebDriver 초기화
    driver = webdriver.Chrome(options=chrome_options)
    return driver

# 다운로드 디렉토리 설정
home_directory = os.path.expanduser('~')
download_directory = os.path.join(home_directory, 'Downloads', 'python')

# WebDriver 인스턴스 가져오기
driver = get_configured_driver(download_directory)

# 로그인 정보 받아오기
login_credentials = get_login_credentials()

# URL로 이동
time.sleep(2)
url = 'https://www.con.or.kr/'
driver.get(url)
driver.maximize_window()

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

def get_login_credentials():
    config = configparser.ConfigParser()
    
    # config.ini 파일이 없으면 생성
    if not os.path.exists('config.ini'):
        create_config()
    
    # config.ini 파일 읽기
    config.read('config.ini', encoding='utf-8')
    
    # 설정이 없거나 불완전한 경우 설정 창 표시
    if not ('CON' in config and 'PPURIO' in config and
            all(config['CON'].get(key) for key in ['username', 'password']) and
            all(config['PPURIO'].get(key) for key in ['username', 'password'])):
        create_config()
        config.read('config.ini', encoding='utf-8')
    
    return {
        'con_username': config['CON']['username'],
        'con_password': config['CON']['password'],
        'ppurio_username': config['PPURIO']['username'],
        'ppurio_password': config['PPURIO']['password']
    }

# 로그인 정보 받아오기
login_credentials = get_login_credentials()
# 로그인 함수 수정
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
# 테스트
# 팝업
# driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div/label/span').click()
# driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[2]/button').click()
# xpath를 이용해 클릭 ##

start_date = datetime(2025, 3, 28)  # 2025년 2월 다섯 번째 금요일
current_date = start_date
current_end_friday = start_date

current_date_time = datetime.now()
print("current_date_time", current_date_time)
count = 12
while current_date.year == 2025:
    # 현재 날짜가 금요일인 경우 count 증가
    if current_date.weekday() == 4:  # 4는 금요일
        count += 1
    print("시작 기수:", count)
    print(current_date.weekday())
    formatted_count = f"2025_{count}_건설사업관리 일반 계속교육(28시간) [기계 · 전기전자 · 환경 · 안전관리] 특급 패키지 1"

    # 과정개설관리 클릭
    driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]').click()
    time.sleep(2)

    # 기수개설관리 클릭
    driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[2]/div[1]/div[1]/a').click()
    time.sleep(2)

    # 추가 버튼 클릭
    driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[1]/div/div/div[2]/div/div[4]').click()
    time.sleep(2)

    # 기존항목복사 라디오버튼 클릭
    driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[1]/div[2]/span').click()
    time.sleep(2)

    # 확인버튼 클릭 (기수 추가에 따른 확인)
    driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[2]').click()
    time.sleep(2)

    driver.find_element(By.XPATH,
                        '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[1]/td[2]/div/div[1]/div[4]/div[2]').click()
    time.sleep(2)

    # 검색 버튼 클릭
    input_field = driver.find_element(By.XPATH,
                                      '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div/div[1]/div/table/tbody/tr/td[2]/input')  # 검색 input
    input_field.click()
    time.sleep(2)

    # 16시간 패키지 입력
    input_field.send_keys("2025-0기_건설사업관리 일반 계속교육(28시간) [기계 · 전기전자 · 환경 · 안전관리] 특급 패키지 1")
    driver.find_element(By.XPATH,
                        '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div/div[2]/div/table/tbody/tr[2]/td').click()
    time.sleep(5)

    # 선택 버튼 클릭
    driver.find_element(By.XPATH,
                        '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[2]/div/table/tbody/tr[1]/td[2]/div/div[1]/div[5]').click()
    time.sleep(5)

    input_field = driver.find_element(By.XPATH,
                                      '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[1]/div/table/tbody/tr[4]/td[2]/input')
    input_field.click()
    time.sleep(20)

    input_field.send_keys("2025-0기_건설사업관리 일반 계속교육(28시간) [기계 · 전기전자 · 환경 · 안전관리] 특급 패키지 1")
    time.sleep(30)

    driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[2]/div[2]/div').click()
    time.sleep(2)

    # 체크박스 클릭
    driver.find_element(By.XPATH,
                        '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[3]/div/table/tbody/tr[2]/td[1]').click()
    time.sleep(3)

    # 확인 버튼
    driver.find_element(By.XPATH, '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[4]/div[1]').click()
    time.sleep(3)

    # 연도 선택
    selectbox = Select(driver.find_element(By.XPATH,
                                           '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[4]/td[2]/div/select'))  # 기간 - 교육종료일 선택
    selectbox.select_by_value('2025')


    def count_fridays(year, month):
        first_day_of_month = datetime(year, month, 1)
        first_friday = first_day_of_month + timedelta(days=(4 - first_day_of_month.weekday() + 7) % 7)

        fridays = []
        current_friday = first_friday
        count = 1
        while current_friday.month == month:
            fridays.append((count, current_friday))
            count += 2
            current_friday += timedelta(days=7)

        return fridays


    def count_from_first_friday(start_date, current_date_time):
        start_friday = start_date + timedelta((4 - start_date.weekday() + 7) % 7)
        current_friday = current_date_time + timedelta((4 - current_date_time.weekday() + 7) % 7)

        if current_friday < start_friday:
            friday_count = 1
        else:
            friday_count = (current_friday - start_friday).days // 7 + 1

        return friday_count, current_friday  # 현재 금요일도 반환


    def count_from_end_friday(start_date, current_date):
        start_friday = start_date
        # current_friday 계산
        current_end_friday = start_friday + timedelta((current_date - start_friday).days // 7 * 7)

        count = (current_end_friday - start_friday).days // 7 + 1
        return count, current_end_friday  # 기수와 현재 금요일 반환


    friday_count, current_end_friday = count_from_first_friday(start_date, current_date)

    # 사용 예시
    start_date = datetime(2025, 2, 28)  # 2025년 1월의 첫 번째 금요일
    current_date_time = datetime.now()
    print("current_date_time", current_date_time)
    friday_count, current_friday = count_from_first_friday(start_date, current_date_time)
    print("기수:", friday_count)

    # 기수 입력
    # driver = webdriver.Chrome()  # 드라이버 초기화 (필요에 따라 적절한 경로로 수정)
    gisu = driver.find_element(By.XPATH,
                               '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[4]/td[4]/input')
    gisu.click()
    time.sleep(3)
    gisu.clear()
    gisu.send_keys(count)
    time.sleep(3)

    # 기수 이름 입력
    gisu_name = driver.find_element(By.XPATH,
                                    '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[5]/td[2]/input')
    gisu_name.click()
    time.sleep(3)
    formatted_count = f"2025_{count}기_건설사업관리 일반 계속교육(28시간) [기계 · 전기전자 · 환경 · 안전관리] 특급 패키지 1"
    gisu_name.clear()
    gisu_name.send_keys(formatted_count)
    time.sleep(3)
    # 수강신청기간 입력
    regi_start_date = driver.find_element(By.XPATH,
                                          '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[6]/td[2]/div/div[1]/div/input')
    regi_start_date.click()
    time.sleep(3)
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("현재 날짜:", current_datetime)
    regi_start_date.send_keys(current_datetime)
    time.sleep(3)
    # 수강신청기간 마감일 입력
    regi_end_date = driver.find_element(By.XPATH,
                                        '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[6]/td[2]/div/div[3]/div/input')
    regi_end_date.click()
    result_date = current_end_friday - timedelta(days=3)
    print("이번 주 금요일에서 3일 뺀 날짜:", result_date.strftime("%Y-%m-%d 23:59:59"))
    regi_end_date.send_keys(result_date.strftime("%Y-%m-%d 23:59:59"))
    time.sleep(3)
    # 학습기간 시작일 입력
    course_start_date = driver.find_element(By.XPATH,
                                            '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[8]/td[2]/div/div[1]/div/input')
    course_start_date.click()
    course_start_date.send_keys('2023-12-28 00:00:00')
    time.sleep(3)
    print('학습시작일', course_start_date)
    # 학습기간 종료일 입력
    course_end_date = driver.find_element(By.XPATH,
                                          '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[8]/td[2]/div/div[3]/div/input')
    course_end_date.click()
    course_end_date.send_keys(current_end_friday.strftime("%Y-%m-%d 23:59:59"))
    print("학습 종료일:", result_date.strftime("%Y-%m-%d 23:59:59"))
    course_end_date.send_keys(Keys.RETURN)
    time.sleep(3)

    # 시험기간 시작일 입력
    course_end_date = driver.find_element(By.XPATH,
                                          '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[10]/td[2]/div/div[1]/div/div[1]/div/input')
    course_end_date.click()
    course_end_date.clear()
    exam_start_date = current_end_friday - timedelta(days=2)
    print("시험 시작일:", result_date.strftime("%Y-%m-%d 00:00:00"))
    course_end_date.send_keys(exam_start_date.strftime("%Y-%m-%d 00:00:00"))
    course_end_date.send_keys(Keys.RETURN)
    time.sleep(3)

    # 과제 종료일 입력
    report_end_date = driver.find_element(By.XPATH,
                                          '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[10]/td[2]/div/div[3]/div/div[3]/div/input')
    report_end_date.click()
    report_end_date.clear()
    report_end_date_before = current_end_friday - timedelta(days=2)
    print("과제 종료일:", report_end_date_before.strftime("%Y-%m-%d 23:59:59"))
    report_end_date.send_keys(report_end_date_before.strftime("%Y-%m-%d 23:59:59"))
    report_end_date.send_keys(Keys.RETURN)
    time.sleep(3)

    # 복습기간 설정
    selectbox = Select(driver.find_element(By.XPATH,
                                           '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[9]/td[2]/div/div[1]/select'))  # 기간 - 교육종료일 선택
    # 일수로 설정
    selectbox.select_by_value('2')
    time.sleep(3)

    # 기간설정 364일
    review_date = driver.find_element(By.XPATH,
                                      '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[9]/td[2]/div/input')
    review_date.click()
    review_date.send_keys('364')
    time.sleep(3)

    # 일일 학습 제한 10시간
    learning_limit = driver.find_element(By.XPATH,
                                         '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[11]/td[4]/input')
    learning_limit.click()
    learning_limit.send_keys('10')
    time.sleep(3)

    # 비게시 체크박스 클릭
    # checkbox_xpath = '//*[@id="wrapper"]/div[1]/div/div/div[1]/div[3]/div/table/tbody/tr[14]/td[2]/div/img'
    # checkbox_element = driver.find_element(By.XPATH, checkbox_xpath)
    # checkbox_element.click()

    # 저장 버튼 클릭
    save_btn = driver.find_element(By.XPATH,
                                   '//*[@id="wrapper"]/div[1]/div/div/div[2]/div[2]/div[1]')
    save_btn.click()
    time.sleep(3)

    close_btn = driver.find_element(By.XPATH,
                                    '//*[@id="popup_layout_list"]/div/div[2]/div[2]/div[2]/div[3]')
    close_btn.click()
    time.sleep(3)
    print("현재 기수 생성 : ", count)
    print(formatted_count, "가 개설되었습니다.")
    current_date += timedelta(days=7)  # 다음 금요일로 이동

    # 과정개설관리 클릭
    driver.find_element(By.XPATH, '//*[@id="side_drop_down_menu"]/div/div[4]/div[7]/div[1]').click()
    time.sleep(3)


