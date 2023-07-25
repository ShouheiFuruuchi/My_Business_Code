from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
import openpyxl as pyxl

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#所属リスト
belong_list = [
  '500005040',
  '本社',
  '//*[@id="timerecorder_id"]/option[2]',#本部
  '//*[@id="timerecorder_id"]/option[4]'
]

d_list = [
  '500011173',
  #1001:'500011173::正社員(本部)',
  #1001:'//*[@id="working_type_id"]',#正社員(本部)
]

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

url = 'https://s3.kingtime.jp/admin/hsXkygNOCWhads8DiqIitHyU294L8qT2?page_id=/login/do_logout'
driver = webdriver.Chrome("C:/Users/fun-f/chromedriver.exe")#旧
#driver = webdriver.Chrome('C:/Users/fun-f/Desktop/myfile/chromedriver.exe')

driver.get(url)


id = "8k61kanri"

password = "Trinity0130"


driver.find_element_by_xpath('//*[@id="login_id"]').send_keys(id)
driver.find_element_by_xpath('//*[@id="login_password"]').send_keys(password)


driver.find_element_by_xpath('//*[@id="login_button"]').click()

#driver.get('https://s3.kingtime.jp/admi')

time.sleep(2)

driver.find_element_by_xpath('//*[@id="intro_start"]').click()

time.sleep(1)

driver.find_element_by_xpath('//*[@id="test"]/div[7]/div/div[5]/a[3]').click()

time.sleep(1)
#driver.execute_script('//*[@id="test"]/div[7]/div/div[5]/a[3]').click()

driver.find_element_by_xpath('//*[@id="test"]/div[7]/div/div[5]/a[3]').click()

time.sleep(1)

driver.find_element_by_xpath('//*[@id="test"]/div[7]/div/div[5]/a[3]').click()

time.sleep(1)

driver.find_element_by_xpath('//*[@id="test"]/div[7]/div/div[5]/a[3]').click()

time.sleep(2)

#従業員をクリック

driver.find_element_by_xpath('//*[@id="step3"]/div/ul/li[3]/a').click()

#ページを以降
#従業員をクリック
driver.find_element_by_xpath('//*[@id="employee_row"]/span[2]/button').click()

time.sleep(2)

driver.find_element_by_xpath('//*[@id="button_01"]/span').click()


#■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■　STAFF DATA 取得　　■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

staff_datafile = "C:/Users/fun-f/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/シフト管理/STAFF DataBase.xlsx"

col_list = [
  "B",#STAFF No
  "C",#氏名
  "D",#雇用形態
  "E",#役職
  "F",#店舗
  "G",#入社日
  "H",#登録(済)
  "I"#ふりがな
  
]

data_file_wb = pyxl.load_workbook(staff_datafile)

data_file_ws = data_file_wb["STAFF_DB"]

no = 606
staff_no = data_file_ws["B" + str(no)].value
staff_name = data_file_ws["C" + str(no)].value
staff_name_last = data_file_ws["C" + str(no)].value[0:3]
staff_name_first = data_file_ws["C" + str(no)].value[3:]

working_status = data_file_ws["D" + str(no)].value
ruby_1 = data_file_ws["I" + str(no)].value[0:3]
ruby_2 = data_file_ws["I" + str(no)].value[3:]


print(staff_no,staff_name,working_status)

#詳細をクリック
driver.find_element_by_xpath('//*[@id="employee_edit_form"]/div[2]/div[1]/h3/span/button[2]').click()

driver.find_element_by_xpath('//*[@id="employee_code"]').send_keys(str(0) + str(staff_no))

#名前を登録
driver.find_element_by_xpath('//*[@id="last_name"]').send_keys(staff_name_last)
driver.find_element_by_xpath('//*[@id="first_name"]').send_keys(staff_name_first)

#フリガナを登録
driver.find_element_by_xpath('//*[@id="last_name_kana"]').send_keys(ruby_1)
driver.find_element_by_xpath('//*[@id="first_name_kana"]').send_keys(ruby_2)

#性別を指定
sex = 0

if sex == 0:

  driver.find_element_by_xpath('//*[@id="sex_code_man"]').click()

else:
  driver.find_element_by_xpath('//*[@id="sex_code_woman"]').click()
  
#生年月日を入力　※全て 2022/1/1

driver.find_element_by_xpath('//*[@id="birth_date_y"]').send_keys(2022)

driver.find_element_by_xpath('//*[@id="birth_date_m"]').send_keys(1)

driver.find_element_by_xpath('//*[@id="birth_date_d"]').send_keys(1)

#所属登録
tag_name1 = driver.find_element_by_name('timerecorder_id')

target1 = Select(tag_name1)

target1.select_by_value(belong_list[0])


#雇用区分登録
tag_name2 = driver.find_element_by_name('working_type_id')
target2 = Select(tag_name2)

target2.select_by_value(d_list[0])
#driver.get('https://s3.kingtime.jp/admin/TyZ6dI0ozkzylFOuYI4x3x1jJTdsrzNH?page_id=/setup/top&rYqwN2uC=zi4mUMuz8ym548Uk#tab4')


#driver.find_element_by_xpath('//*[@id="button_1"]').click()
