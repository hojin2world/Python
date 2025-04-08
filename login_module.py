import configparser
import tkinter as tk
from tkinter import messagebox
import os

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