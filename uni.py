#デニムの類似品番を代表品番ひ集計

import openpyxl as pyxl
import pandas as pd


USER = "古内翔平"

FILE_PATH = "C:/Users/{}/Desktop/品番一覧.xlsx".format(USER)

df = pd.DataFrame(pd.read_excel(FILE_PATH))
print(df.columns)

DENIM_LIST = []

for col in df.columns:
    key_data = df[col].values
    
    for i in key_data:
        
        try:
            
            ItemCD = str(int(i)).zfill(10)
            
            create_data = pd.DataFrame({"商品名":[col],"品番":[int(ItemCD)]})
            DENIM_LIST.append(create_data)
            
            
        except:
            print("No")
        


CONCAT_DENIM_LIST = pd.concat(DENIM_LIST)


print(CONCAT_DENIM_LIST)

#売れ筋在庫一覧
#このプログラムは店別品番別実績を自動ダウンロードを行う

#----------------------------------------------------------------------------------------------

import openpyxl as pyxl
from openpyxl.worksheet.datavalidation import DataValidation
import time
import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service as ChromeServeice
import numpy as np
import datetime
import os
import glob
import shutil
from operator import itemgetter

import datetime
import pandas as pd
from webdriver_manager.chrome import ChromeDriverManager

import requests
import win32com.client

#ーーーーーーーーーーーーーーーーーーーーーーー

#OUPPUTファイル
OUTPUTFILE = "C:/Users/{}/Desktop/店舗在庫/SKU別売上在庫一覧.xlsx".format(USER)
OUT_WB = pyxl.load_workbook(OUTPUTFILE)

#削除ファイルの設定

del_folder_path = "C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/myfile/basket-analysis/data-folder".format(USER)#削除対象フォルダー
del_folder_path2 = 'C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory'.format(USER)
del_read = os.listdir(del_folder_path)#削除対象のディレクトリ内のファイル名を取得
del_read2 = os.listdir(del_folder_path2)

#売上データの削除
# for file_name in del_read:
#   del_faile_path = del_folder_path + '/' + file_name#削除ファイルパスの設定
#   os.remove(del_faile_path)#dataf内のファイルの削除
  
# #在庫データ削除
# for file_name2 in del_read2:
#   del_faile_path2 = del_folder_path2 + '/' + file_name2#削除ファイルパスの設定
#   os.remove(del_faile_path2)#dataf内のファイルの削除  
  
# print('削除完了')  
  
#ーーーーーーー今日の日付設定ーーーーーーーーー

fold = 'C:/Users/{}/Downloads'.format(USER)


todaytime = datetime.date.today()
print(todaytime)
tod = '{0:20%y}'.format(todaytime)#今日の日付(西暦)
y = todaytime.year
m = todaytime.month
d = todaytime.day

SELECTDAY = str(y) + str(m).zfill(2) + str(d).zfill(2)
print(tod)
print("期間のSTART日を入力して下さい！")
print("例 1101(11月1日の場合)")
#period1 = tod + str(input())
period1 = SELECTDAY#str(input())
print("期間のEND日を入力して下さい！")
print("例 1107(11月7日の場合)")
# = tod + str(input())
period2 = SELECTDAY#str(input())
print("店舗指定するかを設定して下さい！")
print("0 = 全店　1 = 店舗指定")

switch = 0#input()

if switch == str(1) :
    print("店舗リストNoを入力して下さい！")
    print("0 = 01001008 FUN柏","\n"
        "1 = 01001009 FUN千葉C-one","\n",
        "2 = 01001028 FUNスマーク伊勢崎","\n"
        "3 = 01001032 FUNララガーデン長町","\n"
        "4 = 01001033 FUNららぽーとTOKYO-BAY","\n"
        "5 = 01001034 FUNららぽーと富士見","\n"
        "6 = 01001036 FUNイオンレイクタウン","\n"
        "7 = 01001038 FUNららぽーと海老名","\n"
        "8 = 01001039 FUNイオンモールむさし村山","\n"
        "9 = 01001040 FUNららぽーと湘南平塚","\n"
        "10 = 01001041 FUNイオンモール名取","\n"
        "11 = 01001042 FUNイオンモール大高","\n"
        "12 = 01001043 FUNららぽーと愛知東郷","\n"
        "13 = 01001044 FUNイオンモール太田","\n"
        "14 = 01001045 FUNイオンモール水戸内原","\n"
        "15 = 01001046 FUNららぽーとEXPOCITY","\n"
        "16 = 01001047 FUNラゾーナ川崎プラザ","\n"
        "17 = 01001048 FUNららぽーと新三郷","\n"
        "18 = 01001049 FUNイオンモール幕張新都心","\n"
        "19 = 01001050 FUNイオンモール各務原","\n"
        )
    t_no = input()
    
else:
    
    print("全店実績データをダウンロード")    


print("検索開始……")


#ーーーーーーー販売NETスクレイピングーーーーーーーーーーー

tenpo = [
    "01001008 FUN柏",
    "01001009 FUN千葉C-one",
    "01001028 FUNスマーク伊勢崎",
    "01001032 FUNララガーデン長町",
    "01001033 FUNららぽーとTOKYO-BAY",
    "01001034 FUNららぽーと富士見",
    "01001036 FUNイオンレイクタウン",
    "01001038 FUNららぽーと海老名",
    "01001039 FUNイオンモールむさし村山",
    "01001040 FUNららぽーと湘南平塚",
    "01001041 FUNイオンモール名取",
    "01001042 FUNイオンモール大高",
    "01001043 FUNららぽーと愛知東郷",
    "01001044 FUNイオンモール太田",
    "01001045 FUNイオンモール水戸内原",
    "01001046 FUNららぽーとEXPOCITY",
    "01001047 FUNラゾーナ川崎プラザ",
    "01001048 FUNららぽーと新三郷",
    "01001049 FUNイオンモール幕張新都心",
    "01001050 FUNイオンモール各務原",
]
def DownLoad():
    del_read = os.listdir(del_folder_path)#削除対象のディレクトリ内のファイル名を取得
    del_read2 = os.listdir(del_folder_path2)

    #売上データの削除
    for file_name in del_read:
        del_faile_path = del_folder_path + '/' + file_name#削除ファイルパスの設定
        os.remove(del_faile_path)#dataf内のファイルの削除
        
    #在庫データ削除
    for file_name2 in del_read2:
        del_faile_path2 = del_folder_path2 + '/' + file_name2#削除ファイルパスの設定
        os.remove(del_faile_path2)#dataf内のファイルの削除  
        
    print('削除完了')  

    url = 'http://tri.hanbai-net.com/system/Login.aspx'
    #driver = webdriver.Chrome("C:/Users/古内翔平/chromedriver.exe")#旧
    driver = webdriver.Chrome(service=ChromeServeice(ChromeDriverManager().install()))#ChromeDriverManager().install())
    #driver = webdriver.Chrome('C:/Users/fun-f/Desktop/myfile/chromedriver.exe')

    driver.get(url)

    #id_1 = 'tenpo'
    #id_2 = 'tenpo'
    #id_2 = 'Tenpo'

    id_1 = 'trinityadmin'
    id_2 = 'AdminTrinity'

    loginid_1 = driver.find_element(By.ID,"ContentPlaceHolder1_txtUserCode")
    loginid_2 = driver.find_element(By.ID,"ContentPlaceHolder1_txtPassword")

    loginid_1.send_keys(id_1)#ユーザーIDを入力
    loginid_2.send_keys(id_2)#パスワードを入力

    driver.find_element(By.ID,"ContentPlaceHolder1_btnLogin").click() 
    #ログインボタンをクリック

    driver.get('http://tri.hanbai-net.com/system/00000000.aspx')

    #.find_element_by_xpath('//*[@id="Menu1"]/ul/li[9]/a').click()

    #driver.find_element_by_xpath('//*[@id="Menu1:submenu:88"]/li[2]/a').click()
    #'//*[@id="Menu1:submenu:58"]/li[9]/a'#変更前

    driver.get('http://tri.hanbai-net.com/system/50010201.aspx?id=010199')

    # if switch == str(1) :
    driver.find_element(By.XPATH,'//*[@id="ContentPlaceHolder1_DropDownList10"]').send_keys("0 販売")
    # driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_DropDownList9"]').send_keys(tenpo[int(t_no)])#店舗指定

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond07").clear()#日付クリア

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond07").send_keys(period1)#日付入力(前)

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond10").clear()#日付クリア

    driver.find_element(By.ID,"ContentPlaceHolder1_txtCond10").send_keys(period2)#日付入力(後)

    # driver.find_element(By.ID,"ContentPlaceHolder1_btnCondRun").click()#検索
    # time.sleep(5)

    driver.find_element(By.ID,"ContentPlaceHolder1_btnCSV").click()#CSV出力

    time.sleep(10)#一時待機

    filelists = []
    for file in os.listdir("C:/Users/{}/Downloads".format(USER)):#ディレクトリ内をfor文で取り出す
        base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
        if ext == '.csv':#拡張子csvが一致した場合…
            if base == '販売分析ログ':
                filelists.append([file, os.path.getctime("C:/Users/{}/Downloads/".format(USER) + file)])#filelistsに取り出したfileにダウンロード時間を追加
                #print("file:{},csv:{}" .format(file,csv))
                filelists.sort(key=itemgetter(0), reverse=True)#
                MAX_CNT = 0
                for i, file in enumerate(filelists):
                    if i > MAX_CNT-1:
                        print(file[0])
                        time.sleep(2)
                        #file_1 = os.rename(i[0], 'kasi.csv')
                        #os.rename("C:/Users/古内翔平/Downloads/販売分析ログ.csv", '顧客データ.csv')
                        shutil.move('C:/Users/{}/Downloads/販売分析ログ.csv'.format(USER),'C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/myfile/basket-analysis/data-folder'.format(USER))                        
    time.sleep(2)                    

    #ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

    driver.get('http://tri.hanbai-net.com/system/21026001.aspx?id=010199')#在庫一覧

    #CSV
    driver.find_element(By.ID, "ContentPlaceHolder1_btnCSV").click()
    time.sleep(10)

    filelists = []

    for file in os.listdir("C:/Users/{}/Downloads".format(USER)):#ディレクトリ内をfor文で取り出す
        base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
        if ext == '.csv':#拡張子csvが一致した場合…
            if base == '在庫一覧_':
                filelists.append([file, os.path.getctime("C:/Users/{}/Downloads/".format(USER) + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
                #print("file:{},csv:{}" .format(file,csv))
                filelists.sort(key=itemgetter(0), reverse=True)#
                MAX_CNT = 0
                for i, file in enumerate(filelists):
                    if i > MAX_CNT-1:
                        print(file[0])
                        #file_1 = os.rename(i[0], 'kasi.csv')
                        os.rename("C:/Users/{}/Downloads/".format(USER) + str(file[0]), "C:/Users/{}/Downloads/".format(USER) + str("全店") + '.csv')
                        shutil.move("C:/Users/{}/Downloads/".format(USER) + str("全店") + '.csv' ,'C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory'.format(USER))                        
    time.sleep(5)      

    #商品マスタ
    driver.get("http://tri.hanbai-net.com/system/10017501.aspx?id=010199")   

    #CSV
    driver.find_element(By.ID, "ContentPlaceHolder1_btnCSV").click()
    time.sleep(10)

    filelists = []

    for file in os.listdir("C:/Users/{}/Downloads".format(USER)):#ディレクトリ内をfor文で取り出す
        base, ext = os.path.splitext(file)#splitextは拡張子を取得する関数
        if ext == '.csv':#拡張子csvが一致した場合…
            if base == '商品マスタ':
                filelists.append([file, os.path.getctime("C:/Users/{}/Downloads/".format(USER) + str(file))])#filelistsに取り出したfileにダウンロード時間を追加
                #print("file:{},csv:{}" .format(file,csv))
                filelists.sort(key=itemgetter(0), reverse=True)#
                MAX_CNT = 0
                for i, file in enumerate(filelists):
                    if i > MAX_CNT-1:
                        print(file[0])
                        #file_1 = os.rename(i[0], 'kasi.csv')
                        #os.rename("C:/Users/古内翔平/Downloads/" + str(file[0]), "C:/Users/古内翔平/Downloads/" + str("全店") + '.csv')
                        shutil.move("C:/Users/{}/Downloads/".format(USER) + str("商品マスタ") + '.csv' ,'C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory'.format(USER))                        
    time.sleep(1)   

    


    driver.close()   

def totaling():
    
    #TenpoName = [
#     "FUN柏",
#     "FUN千葉C-one",
#     "FUNスマーク伊勢崎",
#     "FUNららぽーと富士見",
#     "FUNイオンレイクタウン",
#     "FUNららぽーと海老名",
#     "FUNイオンモールむさし村山",
#     "FUNららぽーと湘南平塚",
#     "FUNイオンモール名取",
#     "FUNイオンモール大高",
#     "FUNららぽーと愛知東郷",
#     "FUNイオンモール太田",
#     "FUNイオンモール水戸内原",
#     "FUNららぽーとEXPOCITY",
#     "FUNラゾーナ川崎プラザ",
#     "FUNららぽーと新三郷",
#     "FUNイオンモール幕張新都心",
#     "FUNイオンモール各務原",
# ]
    TenpoName ={
    "FUN柏":"柏",
    "FUN千葉C-one":"千葉",
    "FUNスマーク伊勢崎":"伊勢崎",
    "FUNららぽーと富士見":"富士見",
    "FUNイオンレイクタウン":"レイクタウン",
    "FUNららぽーと海老名":"海老名",
    "FUNイオンモールむさし村山":"むさし",
    "FUNららぽーと湘南平塚":"平塚",
    "FUNイオンモール名取":"名取",
    "FUNイオンモール大高":"大高",
    "FUNららぽーと愛知東郷":"愛知東郷",
    "FUNイオンモール太田":"太田",
    "FUNイオンモール水戸内原":"水戸",
    "FUNららぽーとEXPOCITY":"EXPO",
    "FUNラゾーナ川崎プラザ":"川崎",
    "FUNららぽーと新三郷":"新三郷",
    "FUNイオンモール幕張新都心":"幕張",
    "FUNイオンモール各務原":"各務原",
    "FUNららぽーと堺":"堺",
    }
    #売上
    File1 = "C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/myfile/basket-analysis/data-folder/販売分析ログ.csv".format(USER)
    DF1 = pd.DataFrame(pd.read_csv(File1,encoding="cp932"))
    shops = pd.DataFrame(DF1["店舗名"].values,columns=["店舗名"])
    item_cd = pd.DataFrame(DF1["商品コード"].astype('str').str.zfill(10).str[:10].values,columns=["商品CD"])
    item_name = pd.DataFrame(DF1["商品名"],columns=["商品名"])
    category_cd = pd.DataFrame(DF1["商品コード"].astype('str').str.zfill(10).str[2:4].values,columns=["アイテムCD"])
    color = pd.DataFrame(DF1["分類１名"].values,columns=["カラー"])
    size = pd.DataFrame(DF1["分類２名"].values,columns=["サイズ"])
    quantity = pd.DataFrame(DF1["数量"].values,columns=["数量"])
    amount = pd.DataFrame(DF1["小計金額"].values,columns=["金額"])
    data1 = pd.concat([shops,item_cd,item_name,category_cd,color,size,quantity,amount],axis=1)

    
    #在庫一覧
    File2 = "C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory/全店.csv".format(USER)
    DF2 = pd.DataFrame(pd.read_csv(File2,encoding="cp932"))
    
    shops2 = pd.DataFrame(DF2["拠点名"].values,columns=["店舗名"])
    item_cd2 = pd.DataFrame(DF2["商品コード"].astype('str').str.zfill(10).str[:10].values,columns=["商品CD"])
    item_name2 = pd.DataFrame(DF2["商品名"],columns=["商品名"])
    category_cd2 = pd.DataFrame(DF2["商品コード"].astype('str').str.zfill(10).str[2:4].values,columns=["アイテムCD"])
    color2 = pd.DataFrame(DF2["色名"].values,columns=["カラー"])
    size2 = pd.DataFrame(DF2["サイズ名"].values,columns=["サイズ"])
    quantity2 = pd.DataFrame(DF2["現在数量"].values,columns=["現在数量"])

    data2 = pd.concat([shops2,item_cd2,item_name2,category_cd2,color2,size2,quantity2],axis=1)

    #商品マスタ
    File3 = "C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory/商品マスタ.csv".format(USER)
    DF3 = pd.DataFrame(pd.read_csv(File3,encoding="cp932"))
    
    item_cd3 = pd.DataFrame(DF3["商品コード"].astype('str').str.zfill(10).str[:10].values,columns=["商品CD"])
    item_name3 = pd.DataFrame(DF3["商品名"],columns=["商品名"])
    category_cd3 = pd.DataFrame(DF3["商品コード"].astype('str').str.zfill(10).str[2:4].values,columns=["アイテムCD"])
    color3 = pd.DataFrame(DF3["色名"].values,columns=["カラー"])
    size3 = pd.DataFrame(DF3["サイズ名"].values,columns=["サイズ"])

    data3 = pd.concat([item_cd3,item_name3,category_cd3,color3,size3],axis=1)

    
    print(data1)
    print(data2)
    print(data3)
    counter = 1
    for shop,shop_ele in zip(TenpoName,TenpoName.values()) :
        
        #SalesInventoryList
        SalesInventoryList = []
        
        #売れ筋ランキング表を作成
        RankingList = []
        
        #売上データ
        select_data1 = data1[data1["店舗名"] == shop ]
        #在庫データ
        select_data2 = data2[data2["店舗名"] == shop ]
        #UniqueItem = np.unique(select_data1["商品名"].values)
        
        #デニムリスと
        UniqueItem = np.unique(CONCAT_DENIM_LIST["商品名"].values)
        
        
        for ItemName in UniqueItem :
            
            Key_Item = CONCAT_DENIM_LIST[CONCAT_DENIM_LIST["商品名"] == ItemName]["品番"].values
            
            for ItemCD in Key_Item:
            
                HitItem = select_data1[select_data1["商品CD"] == ItemCD ]
                
                Sum_Quantity = sum(HitItem["数量"].values)
                Sum_Amount = sum(HitItem["金額"].values)
                
                SalesData = pd.DataFrame([{"商品名":ItemName,"商品CD":ItemCD,"数量":Sum_Quantity,"金額":Sum_Amount}])
                
                RankingList.append(SalesData)
            
        ConcatRankingList = pd.concat(RankingList).sort_values("金額",ascending=False)#.head(15)        
        print(ConcatRankingList)
        
        for i in ConcatRankingList.values:
            print(i[1])
        
            
            M_Data = data3[data3["商品CD"] == str(i[1])]
            print("一致データ",M_Data)
            for i2 in M_Data.values :
                try :
                
                    Sales_Match_Data = select_data1[(select_data1["商品名"] == i2[1]) &(select_data1["カラー"] == i2[3]) & (select_data1["サイズ"] == i2[4])]
                    Sales_Data_Q = Sales_Match_Data["数量"].values[0]
                    Sales_Data_A = Sales_Match_Data["金額"].values[0]
                except :
                    Sales_Data_Q = 0
                    Sales_Data_A = 0

                try:
                    Inv_Match_Data = select_data2[(select_data2["商品名"] == i2[1]) &(select_data2["カラー"] == i2[3]) & (select_data2["サイズ"] == i2[4])]
                    Inv_Data = Inv_Match_Data["現在数量"].values[0]
                except :
                    Inv_Data = 0    
                    
                
                #CreateData = pd.DataFrame([{"店舗名":shop,"商品CD":i2[0],"商品名":i2[1],"カラー":i2[3],"サイズ":i2[4],"点数":Sales_Data_Q,"金額":Sales_Data_A,"在庫数":Inv_Data}])   
                #カラー表記修正
                if i2[3] == "BL":
                    dec_color = "BLUE"
                    
                elif i2[3] == "BK":
                    dec_color = "BLACK"
                    
                else:
                    dec_color = i2[3]
                        
                CreateData = pd.DataFrame([{"店舗名":shop,"商品CD":i2[0],"商品名":i[0],"カラー":dec_color,"サイズ":i2[4],"点数":Sales_Data_Q,"金額":Sales_Data_A,"在庫数":Inv_Data}])   
                 
                #print("CreateData",CreateData)
                SalesInventoryList.append(CreateData)  
        #print("作成リスト",SalesInventoryList)
        ConcatSalesInventoryList = pd.concat(SalesInventoryList)       
        print(ConcatSalesInventoryList)
        
        size_no = {
            "XS":1,
            "S":2,
            "M":3,
            "L":4,
            "XL":5,
            "XS-丈短め":6,
            "XS-丈長め":7,
            "S-丈短め":8,
            "S-丈長め":9,
            "M-丈短め":10,
            "M-丈長め":11,
            "L-丈短め":12,
            "L-丈長め":13,
            "XL-丈短め":14,
            "XL-丈長め":15,
        }
        
        size_no2 = {
            1:"XS",
            2:"S",
            3:"M",
            4:"L",
            5:"XL",
            6:"XS-丈短め",
            7:"XS-丈長め",
            8:"S-丈短め",
            9:"S-丈長め",
            10:"M-丈短め",
            11:"M-丈長め",
            12:"L-丈短め",
            13:"L-丈長め",
            14:"XL-丈短め",
            15:"XL-丈長め",
            
            
        }
        
        ReConcatSalesInventoryList = []
        for itemname in UniqueItem:
            select_data3 =  ConcatSalesInventoryList[ConcatSalesInventoryList["商品名"] == itemname]
            shop = np.unique(select_data3["店舗名"].values)
            colors = np.unique(select_data3["カラー"].values)
            sizes = np.unique(select_data3["サイズ"].values)
            
            for c in colors:
                for s in sizes:
                    try :
                        select = select_data3[(select_data3["カラー"] == c) & (select_data3["サイズ"] == s) ]
                        #CreateData = pd.DataFrame([{"店舗名":shop,"商品CD":i2[0],"商品名":i[0],"カラー":dec_color,"サイズ":i2[4],"点数":Sales_Data_Q,"金額":Sales_Data_A,"在庫数":Inv_Data}])   
                    
                        #print(select)
                        #shop = np.unique(select["店舗名"].values)
                        sales_quant = sum(select["点数"].values)
                        inv_quant = sum(select["在庫数"].values)
                        
                    except :
                        inv_quant = 0
                            
                    reCreateData = pd.DataFrame([{"店舗名":shop[0],"商品CD":"代表CDあれば代入","商品名":itemname,"カラー":c,"サイズ":s,"サイズNo":size_no[s],"在庫数":inv_quant}])  
                    ReConcatSalesInventoryList.append(reCreateData) 
                    
        Last_ReConcatSalesInventoryList = pd.concat(ReConcatSalesInventoryList)  
        
        print(Last_ReConcatSalesInventoryList)
        
        
        OUT_WS = OUT_WB["集計"]
        #counter = 1
        
        MST_Sheet = {
            "F-0147 ﾎﾞﾀﾝﾌﾚｱｰPT":"ボタンフレアーPT",
            "F-0101 ﾍﾞﾙﾄ付ｽﾄﾚｰﾄPT":"ベルト付ストレートPT",
            "F-0058-1美脚ｽｷﾆｰ":"美脚スキニー",
            "F-0036 ﾎﾞﾀﾝﾊｲｳｴｽﾄPT":"ボタンハイウエスト",
            
        }
        for itemname in UniqueItem:
            MST_SH_C = OUT_WB[itemname]
            MST_SH_S = OUT_WB[MST_Sheet[itemname]]
            
            HITDATA = Last_ReConcatSalesInventoryList[Last_ReConcatSalesInventoryList["商品名"] == itemname].sort_values("サイズNo")  
            colors2 = np.unique(HITDATA["カラー"].values)
            sizes2 = np.unique(HITDATA["サイズNo"].values)
            # dv = DataValidation(
            #     type="list",
            #     formula1=colors2,
            #     allow_blank=True,
            #     showErrorMessage=True,
            #     errorStyle="warning",
            #     errorTitle="選択リストにない場合のみ、入力してください",
            #     error="続けますか？"
            #     )
            
            print(sizes2)
            #dv.add(f"B3")
            # dv.add(f"B3:D3")
            # MST_SH.add_data_validation(dv)
            
            #カラーデータを出力
            c_counter = 0
            for c_mst in colors2:
                MST_SH_C.cell(2,6  + c_counter).value = c_mst
                c_counter += 1
                
            #サイズデータを出力
            s_counter = 0
            for s_mst in sizes2:
                MST_SH_S["E" + str(3 + s_counter)].value = size_no2[s_mst]
                s_counter += 1
                
            for c2 in colors2:
            
                HITDATA2 = HITDATA[HITDATA["カラー"] == c2]
        
                for i3 in HITDATA2.values :
                    #OUT_WS["B" + str(3)].value = dv[0]
                    OUT_WS["A" + str(1 + counter)].value = i3[0]
                    OUT_WS["B" + str(1 + counter)].value = i3[1]
                    OUT_WS["C" + str(1 + counter)].value = i3[2]
                    OUT_WS["D" + str(1 + counter)].value = i3[3]
                    OUT_WS["E" + str(1 + counter)].value = i3[4]
                    OUT_WS["F" + str(1 + counter)].value = i3[5]
                    OUT_WS["G" + str(1 + counter)].value = i3[6]
                # OUT_WS["H" + str(1 + counter)].value = i3[7]
                
                    counter += 1
                
    OUT_WB.save("C:/Users/{}/Desktop/店舗在庫/{}_SKU別売上在庫一覧.xlsx".format(USER,SELECTDAY))   
    #"C:\Users\FUN-PC132\Desktop\店舗在庫\SKU別売上在庫一覧.xlsx"         

        

    #     OUT_WS = OUT_WB[shop_ele]
    #     counter = 1
    #     for i3 in ConcatSalesInventoryList.values :
    #         OUT_WS["A" + str(1 + counter)].value = i3[0]
    #         OUT_WS["B" + str(1 + counter)].value = i3[1]
    #         OUT_WS["C" + str(1 + counter)].value = i3[2]
    #         OUT_WS["D" + str(1 + counter)].value = i3[3]
    #         OUT_WS["E" + str(1 + counter)].value = i3[4]
    #         OUT_WS["F" + str(1 + counter)].value = i3[5]
    #         OUT_WS["G" + str(1 + counter)].value = i3[6]
    #         OUT_WS["H" + str(1 + counter)].value = i3[7]
            
    #         counter += 1
            
    # OUT_WB.save("C:/Users/{}/Desktop/店舗在庫/{}_SKU別売上在庫一覧.xlsx".format(USER,SELECTDAY))   
    # #"C:\Users\FUN-PC132\Desktop\店舗在庫\SKU別売上在庫一覧.xlsx"         



excel = win32com.client.Dispatch("Excel.Application")
#path = r'C:/abc/def/ghi'
#path = "C:/Users/古内翔平/Downloads"
path = "C:/Users/{}/Desktop/{}_SKU別売上在庫一覧.xlsx".format(USER,SELECTDAY)


#r削除
#inputDir = COPYFILE_PATH#"C:/Users/古内翔平/Downloads/【20220703】 2023 10月シフト 【販売部】 ver 17.xlsm"#path + r'\Excel'
inputDir = "C:/Users/{}/Desktop/{}_SKU別売上在庫一覧.xlsx".format(USER,SELECTDAY)
outputDir = path + r'\PDF'

shop_name_list = [
  "柏",
  "千葉",
  "伊勢崎",
  "富士見",
  "レイクタウン",
  "海老名",
  "むさし",
  "平塚",
  "名取",
  "大高",
  "愛知東郷",
  "太田",
  "水戸",
  "EXPO",
  "川崎",
  "新三郷",
  "幕張",
  "各務原",
  "堺",
  
]

def PDF_FILE(file_path):

    base, ext = os.path.splitext(file_path)
    if ext == '.xlsx' and '~$' not in base:
        wb = excel.Workbooks.Open(os.path.join(inputDir,file_path))
        for shop_name in shop_name_list:
            SHEET_NAME = shop_name 
            wb.WorkSheets(SHEET_NAME).Select()
            #"C:/Users/古内翔平・OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/シフト管理 - 遠藤 孝道 さんのファイル/在庫管理表
            wb.ActiveSheet.ExportAsFixedFormat(0,"C:/Users/{}/Desktop/{}.pdf".format(USER,SHEET_NAME))# outputDir + '/' + base + '.pdf")
        wb.Close()
DownLoad()        
totaling()    
#PDF_FILE(inputDir)   

shutil.copy("C:/Users/{}/Desktop/店舗在庫/{}_SKU別売上在庫一覧.xlsx".format(USER,SELECTDAY),"C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/シフト管理 - 遠藤 孝道 さんのファイル/在庫管理表/{}_SKU別売上在庫一覧.xlsx".format(USER,SELECTDAY)) 



TOKEN = 'NxPDQg0tpI4oY6BZHo7vkZ3gxPtTIijpWFyN85xL2q1'#テストの部屋トークン
#TOKEN = 'TNKXBcpEMmK4JAmRPaOyVABkA5GWIJkIHQOnsyfu4MD'#FUNの部屋トークン
api_url = 'https://notify-api.line.me/api/notify'
headers = {'Authorization' : 'Bearer ' + TOKEN}

message_1 = ('\nお疲れ様です。\n\n在庫一覧を作成完了しました')
payload = {'message': message_1}
requests.post(api_url, headers=headers, params=payload)   
print("SUCCESSFULL!!")    


    
