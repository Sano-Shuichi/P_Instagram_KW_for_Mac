from datetime import datetime
import requests
from bs4 import BeautifulSoup
import tkinter as tk
import tkinter.messagebox as mb
import tkinter.simpledialog as sd
import openpyxl as excel, time, random, os, sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# 初期設定
path = os.path.dirname(sys.argv[0])+'/'
ran = random.randint(3,5)
t = datetime.now()
fmt = t.strftime('%Y-%m-%d')
fmt_m = t.strftime('%m')
tk.Tk().withdraw() # tkinterのウィンドウを隠す
option = Options()
option.add_argument('--incognito') # シークレットモードで起動
option.add_experimental_option('excludeSwitches', ['enable-logging']) # エラーメッセージを消す

book = excel.load_workbook(path+'tmp_Instagram.xlsx')
sheet = book.active

user_id = sheet['B2'].value
password = sheet['B3'].value

# Instagramにログイン
driver = webdriver.Chrome(path+'chromedriver',options=option)
driver.set_window_size(700, 900)
driver.get('https://www.instagram.com/')
time.sleep(ran)
usr = driver.find_element_by_name('username')
usr.send_keys(user_id)
time.sleep(ran)
pwd = driver.find_element_by_name('password')
pwd.send_keys(password)
time.sleep(10)
pwd.submit()
time.sleep(10)
driver.find_element_by_class_name('sqdOP.yWX7d.y3zKF').click()
time.sleep(7)
driver.find_element_by_class_name('aOOlW.HoLwm').click()
time.sleep(7)


for row in sheet.iter_rows(min_row=6):
    values = [v.value for v in row]
    if values[0] is None: break # 空白セルであれば読み取りを終わる
    
    name = values[0]
    
    for i, v in enumerate(values):
        if i == 0: continue
        if v is None: break
        kw = v

        # Instagramでキーワードのサジェスト内容をキャプチャ
        kw_form = driver.find_element_by_class_name('XTCLo.d_djL.DljaH')
        kw_form.send_keys(kw)
        time.sleep(10)
        
        save_dir_1st = os.path.join(os.path.dirname(sys.argv[0]), values[0])
        if not os.path.exists(save_dir_1st):
            os.mkdir(save_dir_1st)
        save_dir_2nd = os.path.join(save_dir_1st, fmt_m)
        if not os.path.exists(save_dir_2nd):
            os.mkdir(save_dir_2nd)
        
        driver.save_screenshot(f'{save_dir_2nd}/{kw}_{fmt}.png')
        
        driver.get('https://www.instagram.com/')
        time.sleep(5)

driver.quit()