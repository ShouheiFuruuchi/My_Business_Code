

#----------------------------------------------------------------------------------------------


import time
import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
import datetime
import os
import glob
import shutil
from operator import itemgetter
import datetime
import pandas as pd
import re
import numpy as np
import collections
import requests
import openpyxl as xlpy
from webdriver_manager.chrome import ChromeDriverManager


#このプログラムは店別品番別実績を自動ダウンロードを行う

#ーーーーーーー販売NETスクレイピングーーーーーーーーーーー

kasiwa = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[3]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[3]','柏','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[3]',"01001008 FUN柏","01001008",]
chiba = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[4]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[4]', '千葉','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[4]',"01001009 FUN千葉C-one","01001009",]
isesaki = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[9]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[9]','伊勢崎','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[9]',"01001028 FUNスマーク伊勢崎","01001028",]
nagamachi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[11]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[11]','長町','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[11]',"01001032 FUNララガーデン長町","01001032",]
hunabashi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[12]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[12]','船橋','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[12]',"01001033 FUNららぽーとTOKYO-BAY","01001033",]
hujimi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[13]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[13]','富士見','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[13]',"01001034 FUNららぽーと富士見","01001034",]
reiku = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[15]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[15]','レイク','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[15]',"01001036 FUNイオンレイクタウン","01001036",]
ebina = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[17]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[17]','海老名','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[17]',"01001038 FUNららぽーと海老名","01001038",]
musashi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[18]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[18]','むさし','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[18]',"01001039 FUNイオンモールむさし村山","01001039",]
hiratuka = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[19]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[19]','平塚','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[19]',"01001040 FUNららぽーと湘南平塚","01001040",]
natori = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[20]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[20]','名取','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[20]',"01001041 FUNイオンモール名取","01001041",]
otaka = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[21]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[21]','大高','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[21]',"01001042 FUNイオンモール大高","01001042",]
togocyo = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[22]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[22]','東郷町','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[22]',"01001043 FUNららぽーと愛知東郷","01001043",]
ota = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[23]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[23]','太田','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[23]',"01001044 FUNイオンモール太田","01001044",]
mito = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[24]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[24]','水戸','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[24]',"01001045 FUNイオンモール水戸内原","01001045",]
expo = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[25]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[25]','EXPO','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[25]',"01001046 FUNららぽーとEXPOCITY","01001046",]
kawasaki = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[26]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[26]','川崎','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[26]',"01001047 FUNラゾーナ川崎プラザ","01001047",]
sinmisato = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[27]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[27]','新三郷','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[27]',"01001048 FUNららぽーと新三郷","01001048",]
makuhari = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[28]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[28]','幕張','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[28]',"01001049 FUNイオンモール幕張新都心","01001049",]
kagamihara = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[29]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[29]','各務原','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[29]',"01001050 FUNイオンモール各務原","01001050",]
sakai = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[30]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[30]','堺','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[30]',"01001051 FUNららぽーと堺","01001051",]



tenpo_list = [
  kasiwa,
  chiba,
  isesaki,
  # nagamachi,
  # hunabashi,
  hujimi,
  reiku,
  ebina,
  musashi,
  hiratuka,
  natori,
  otaka,
  togocyo,
  ota,
  mito,
  expo,
  kawasaki,
  sinmisato,
  makuhari,
  kagamihara,
  sakai,
  ]


#ーーーーーーーーーー前回データの削除ーーーーーーーーーーーーー

#ーーーーーーーー前回ダウンロードファイル削除ーーーーーーーーーー
#dr_files = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/Desktop/myfile/dataf'
dr_files = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory'
dr_read = os.listdir(dr_files)

print(dr_read)

for file_name in dr_read:
  del_f_path = dr_files + '/' + file_name#削除ファイルパスの設定
  os.remove(del_f_path)#dataf内のファイルの削除
  
folders = [0,1,2,3,4,5,6]
no = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]

w_day = datetime.datetime.today()

wd_no = w_day.weekday()#曜日Noを指定

main_dr = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/Desktop/myfile'

print(wd_no)

to_file_path = str(main_dr) + '/' + str(wd_no)#drpathの指定

#---------------------------------------------------

url = 'http://tri.hanbai-net.com/system/Login.aspx'
#driver = webdriver.Chrome('C:/Users/fun-f/Downloads/chromedriver.exe')#旧
#driver = webdriver.Chrome('C:/Users/fun-f/Desktop/myfile/chromedriver.exe')
#driver = webdriver.Chrome("C:/Users/古内翔平/chromedriver.exe")#2021 0724
driver = webdriver.Chrome()#ChromeDriverManager().install())

driver.get(url)

id_1 = 'trinityadmin'
id_2 = 'AdminTrinity'


loginid_1 = driver.find_element(By.ID, "ContentPlaceHolder1_txtUserCode")
loginid_2 = driver.find_element(By.ID, "ContentPlaceHolder1_txtPassword")

loginid_1.send_keys(id_1)#ユーザーIDを入力
loginid_2.send_keys(id_2)#パスワードを入力


#ログインボタンをクリック
driver.find_element(By.ID, "ContentPlaceHolder1_btnLogin").click()


time.sleep(2)

driver.get('http://tri.hanbai-net.com/system/21026001.aspx?id=010199')#在庫一覧

#CSV
driver.find_element(By.ID, "ContentPlaceHolder1_btnCSV").click()
time.sleep(5)

filelists = []

for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
    base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
    if ext == '.csv':#拡張子csvが一致した場合…
        if base == '在庫一覧_':
            filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
            #print("file:{},csv:{}" .format(file,csv))
            filelists.sort(key=itemgetter(0), reverse=True)#
            MAX_CNT = 0
            for i, file in enumerate(filelists):
                if i > MAX_CNT-1:
                    print(file[0])
                    #file_1 = os.rename(i[0], 'kasi.csv')
                    os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" + str("全店") + '.csv')
                    shutil.move("C:/Users/古内翔平/Downloads/" + str("全店") + '.csv' ,'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory')                        
time.sleep(1)            


for shop in tenpo_list :
    Select_ele = driver.find_element(By.ID, "ContentPlaceHolder1_DropDownListCond01")
    Select_value = Select(Select_ele)
    Select_value.select_by_value(shop[5])
    
    
    #日付入力
    #driver.find_element(By.ID, "ContentPlaceHolder1_txtCond02").send_keys(str(tod))
    #----------全店------------

    #CSV
    driver.find_element(By.ID, "ContentPlaceHolder1_btnCSV").click()
    time.sleep(3)

    filelists = []

    for file in os.listdir("C:/Users/古内翔平/Downloads"):#ディレクトリ内をfor文で取り出す
        base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
        if ext == '.csv':#拡張子csvが一致した場合…
            if base == '在庫一覧_':
                filelists.append([file, os.path.getctime("C:/Users/古内翔平/Downloads/" + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
                #print("file:{},csv:{}" .format(file,csv))
                filelists.sort(key=itemgetter(0), reverse=True)#
                MAX_CNT = 0
                for i, file in enumerate(filelists):
                    if i > MAX_CNT-1:
                        print(file[0])
                        #file_1 = os.rename(i[0], 'kasi.csv')
                        os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" + str(shop[2]) + '.csv')
                        shutil.move("C:/Users/古内翔平/Downloads/" + str(shop[2]) + '.csv' ,'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory')                        
    time.sleep(1)                    

#--------店別---------
