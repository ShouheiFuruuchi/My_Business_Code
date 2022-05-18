import pandas as pd
import os
import shutil
import datetime
import requests
import schedule
import time

#店舗リスト

kasiwa = 'C:/Users/fun-f/Desktop/myfile/dataf/柏.csv'
tiba = 'C:/Users/fun-f/Desktop/myfile/dataf/千葉.csv'
yokohama = 'C:/Users/fun-f/Desktop/myfile/dataf/横浜.csv'
isesaki = 'C:/Users/fun-f/Desktop/myfile/dataf/伊勢崎.csv'
gihu = 'C:/Users/fun-f/Desktop/myfile/dataf/岐阜.csv'
nagamachi = 'C:/Users/fun-f/Desktop/myfile/dataf/長町.csv'
hunabasi = 'C:/Users/fun-f/Desktop/myfile/dataf/船橋.csv'
hujimi = 'C:/Users/fun-f/Desktop/myfile/dataf/富士見.csv'
reiku = 'C:/Users/fun-f/Desktop/myfile/dataf/レイク.csv'
ebina = 'C:/Users/fun-f/Desktop/myfile/dataf/海老名.csv'
musasi = 'C:/Users/fun-f/Desktop/myfile/dataf/むさし.csv'
hiratuka = 'C:/Users/fun-f/Desktop/myfile/dataf/平塚.csv'
natori = 'C:/Users/fun-f/Desktop/myfile/dataf/名取.csv'
otaka = 'C:/Users/fun-f/Desktop/myfile/dataf/大高.csv'
togocyo = 'C:/Users/fun-f/Desktop/myfile/dataf/東郷町.csv'
ota = 'C:/Users/fun-f/Desktop/myfile/dataf/太田.csv'
mito = 'C:/Users/fun-f/Desktop/myfile/dataf/水戸.csv'
expo = 'C:/Users/fun-f/Desktop/myfile/dataf/EXPO.csv'
kawasaki = 'C:/Users/fun-f/Desktop/myfile/dataf/川崎.csv'
sinmisato = 'C:/Users/fun-f/Desktop/myfile/dataf/新三郷.csv'
all_sp = 'C:/Users/fun-f/Desktop/myfile/dataf/全店.csv'

no = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21]

shops_d = {1:kasiwa,2:tiba,3:yokohama,4:isesaki,5:gihu,6:nagamachi,7:hunabasi,8:hujimi,9:reiku,10:ebina,11:musasi,12:hiratuka,13:natori,14:otaka,15:togocyo,16:ota,17:mito,18:expo,19:kawasaki,20:sinmisato}

shops_l = [kasiwa,tiba,yokohama,isesaki,gihu,nagamachi,hunabasi,hujimi,reiku,ebina,musasi,hiratuka,natori,otaka,togocyo,ota,mito,expo,kawasaki,sinmisato,all_sp]

print("条件指数を入力して下さい！")
print("0 = 本日実績")
print("1 = 期間指定実績")

switch = input()

if switch == str(0):
  path1 = 'C:/Users/fun-f/Desktop/myfile/dataf/全店.csv'
  
else:
  
  path1 = 'C:/Users/fun-f/Desktop/myfile/basket-analysis/data-folder/品番売上集計データ.csv'

#dt_r = pd.read_csv('C:/Users/fun-f/Desktop/myfile/dataf/川崎.csv',encoding='SHIFT-JIS')

dt_r = pd.read_csv(path1,encoding='SHIFT-JIS')

df_pv = pd.DataFrame(dt_r)

df = pd.DataFrame(columns=['商品CD','商品名','アイテムCD','数量','金額'])


itm_dt = pd.DataFrame(df_pv['商品コード'].astype('str').str.zfill(10).values,columns=['商品CD'])
itm_name = pd.DataFrame(df_pv['商品名'].values,columns=['商品名'])
itm_cd = pd.DataFrame(df_pv['商品コード'].astype('str').str.zfill(10).str[2:4].values,columns=['アイテムCD'])
itm_amt = pd.DataFrame(df_pv['合計金額'].values,columns=['金額'])
itm_qyt = pd.DataFrame(df_pv['合計数量'].values,columns=['数量'])

df_1 = pd.concat([itm_dt,itm_name,itm_cd,itm_qyt,itm_amt],axis=1)

print(df_1)

print(itm_dt)
print(itm_cd)

df_1.to_excel('C:/Users/fun-f/Desktop/myfile/テスト.xlsx')

sort_v = df_1.sort_values(by='金額',ascending=False).head(10)#全アイテムベスト
op = df_1[df_1['アイテムCD'] == '01']
op_best = pd.DataFrame(op.sort_values(by='金額',ascending=False))

op1 = pd.DataFrame(op_best.head(5)[0:1])#1位
op2 = pd.DataFrame(op_best.head(5)[1:2])#2位
op3 = pd.DataFrame(op_best.head(5)[2:3])#3位
op4 = pd.DataFrame(op_best.head(5)[3:4])#4位
op5 = pd.DataFrame(op_best.head(5)[4:5])#5位

item_list1 = [op1,op2,op3,op4,op5]

r_list = []
for item in item_list1:
  print(item)
  
  
print(r_list)  

for itm in op1['商品名']:
  itm_1 = str(itm)

for q in op1['数量']:
  q_1 = str(q)   

for v in op1['金額']:
  v_1 = str(v)

com1 = '1位 '+ itm_1 + '\n'+ '点数 ' + q_1 +' 金額 ¥'+ v_1
op_1 = pd.DataFrame({"商品名":[itm_1] ,"点数":[q_1],"金額":[v_1]})

for itm in op2['商品名']:
  itm_2 = str(itm)

for q in op2['数量']:
  q_2 = str(q)   

for v in op2['金額']:
  v_2 = str(v)

com2 = '2位 '+ itm_2 + '\n'+ '点数 ' + q_2  +' 金額 ¥'+ v_2
op_2 = pd.DataFrame({"商品名":[itm_2] ,"点数":[q_2],"金額":[v_2]})

for itm in op3['商品名']:
  itm_3 = str(itm)

for q in op3['数量']:
  q_3 = str(q)   

for v in op3['金額']:
  v_3 = str(v)

com3 = '3位 '+ itm_3 + '\n'+ '点数 ' + q_3 +' 金額 ¥'+ v_3
op_3 = pd.DataFrame({"商品名":[itm_3] ,"点数":[q_3],"金額":[v_3]})
for itm in op4['商品名']:
  itm_4 = str(itm)

for q in op4['数量']:
  q_4 = str(q)   

for v in op4['金額']:
  v_4 = str(v)

com4 = '4位 '+ itm_4 + '\n'+ '点数 ' + q_4  +' 金額 ¥'+ v_4
op_4 = pd.DataFrame({"商品名":[itm_4] ,"点数":[q_4],"金額":[v_4]})
for itm in op5['商品名']:
  itm_5 = str(itm)

for q in op5['数量']:
  q_5 = str(q)   

for v in op5['金額']:
  v_5 = str(v)

com5 = '5位 '+ itm_5 + '\n'+ '点数 ' + q_5 + ' 金額 ¥'+ v_5
op_5 = pd.DataFrame({"商品名":[itm_5] ,"点数":[q_5],"金額":[v_5]})


#print(str(itm))
msg1 = ('\n'+'【OP】ワンピース'+'\n'+(com1)+'\n'+'\n'+(com2)+'\n'+'\n'+(com3)+'\n'+'\n'+(com4)+'\n'+'\n'+(com5))


op_list = pd.concat([op_1,op_2,op_3,op_4,op_5],axis=0)

op_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_1.xlsx')


#ーーーーーーーーーー　アイテム　ーーーーーーーーーー

cd = df_1[df_1['アイテムCD'] == '02']
cd_best = pd.DataFrame(cd.sort_values(by='金額',ascending=False))

cd1 = pd.DataFrame(cd_best.head(5)[0:1])#1位
cd2 = pd.DataFrame(cd_best.head(5)[1:2])#2位
cd3 = pd.DataFrame(cd_best.head(5)[2:3])#3位
cd4 = pd.DataFrame(cd_best.head(5)[3:4])#4位
cd5 = pd.DataFrame(cd_best.head(5)[4:5])#5位

item_list2 = [cd1,cd2,cd3,cd4,cd5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if cd1['アイテムCD'].values == ['02']:
  print(cd1)
  
  for i in cd1['商品名'].values:
    cd1_1 = (i)
 
  for i in cd1['数量'].values:
    cd1_2 = str(i)
  
  for i in cd1['金額'].values:
    cd1_3 = str(i)
  
  cd_cm1 = '1位 '+ cd1_1 + '\n'+ '点数 ' + cd1_2 + ' 金額 ¥'+ cd1_3
else:
  print('non item')
  cd_cm1 = str('1位 '+'NO ITEM')

cd_1 = pd.DataFrame({"商品名":[cd1_1] ,"点数":[cd1_2],"金額":[cd1_3]})  
#ーーーーーーーーー　２位　ーーーーーーーーーーー

if cd2['アイテムCD'].values == ['02']:
  print(cd2)
    
  for i in cd2['商品名'].values:
    cd2_1 = (i)
 
  for i in cd2['数量'].values:
    cd2_2 = str(i)
  
  for i in cd2['金額'].values:
    cd2_3 = str(i)
    cd_cm2 = '2位 '+ cd2_1 + '\n'+ '点数 ' + cd2_2 + ' 金額 ¥'+ cd2_3
  
else:
  print('non item')
  cd_cm2 = str('2位 '+'NO ITEM')

cd_2 = pd.DataFrame({"商品名":[cd2_1] ,"点数":[cd2_2],"金額":[cd2_3]})    
#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if cd3['アイテムCD'].values == ['02']:
    
  for i in cd3['商品名'].values:
    cd3_1 = (i)
 
  for i in cd3['数量'].values:
    cd3_2 = str(i)
  
  for i in cd3['金額'].values:
    cd3_3 = str(i)
    cd_cm3 = '3位 '+ cd3_1 + '\n'+ '点数 ' + cd3_2 + ' 金額 ¥'+ cd3_3

else:
  print('non item')
  cd_cm3 = str('3位 '+'NO ITEM')

cd_3 = pd.DataFrame({"商品名":[cd3_1] ,"点数":[cd3_2],"金額":[cd3_3]})  
#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if cd4['アイテムCD'].values == ['02']:
    
  for i in cd4['商品名'].values:
    cd4_1 = (i)
 
  for i in cd4['数量'].values:
    cd4_2 = str(i)
  
  for i in cd4['金額'].values:
    cd4_3 = str(i)
    cd_cm4 = '4位 '+ cd4_1 + '\n'+ '点数 ' + cd4_2 + ' 金額 ¥'+ cd4_3
else:
  print('non item')
  cd_cm4 = str('4位 '+'NO ITEM')

cd_4 = pd.DataFrame({"商品名":[cd4_1] ,"点数":[cd4_2],"金額":[cd4_3]})
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if cd5['アイテムCD'].values == ['02']:
    
  for i in cd5['商品名'].values:
    cd5_1 = (i)
 
  for i in cd5['数量'].values:
    cd5_2 = str(i)
  
  for i in cd5['金額'].values:
    cd5_3 = str(i)
    cd_cm5 = '5位 '+ cd5_1 + '\n'+ '点数 ' + cd5_2 + ' 金額 ¥'+ cd5_3

else:
  print('non item')  
  cd_cm5 = str('5位 '+'NO ITEM')
  cd5_1 = "Non Item"
  cd5_2 = "Non Item"
  cd5_3 = "Non Item"
  
cd_5 = pd.DataFrame({"商品名":[cd5_1] ,"点数":[cd5_2],"金額":[cd5_3]})

#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg2 = ('\n'+'【CD】カーデ'+'\n'+(cd_cm1)+'\n'+'\n'+(cd_cm2)+'\n'+'\n'+(cd_cm3)+'\n'+'\n'+(cd_cm4)+'\n'+'\n'+(cd_cm5))


cd_list = pd.concat([cd_1,cd_2,cd_3,cd_4,cd_5],axis=0)

cd_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_2.xlsx')  

#ーーーーーーーー　END　ーーーーーーーーーーー


#ーーーーーーーーーー　アイテム　ーーーーーーーーーーjk

jk = df_1[df_1['アイテムCD'] == '03']
jk_best = pd.DataFrame(jk.sort_values(by='金額',ascending=False))

jk1 = pd.DataFrame(jk_best.head(5)[0:1])#1位
jk2 = pd.DataFrame(jk_best.head(5)[1:2])#2位
jk3 = pd.DataFrame(jk_best.head(5)[2:3])#3位
jk4 = pd.DataFrame(jk_best.head(5)[3:4])#4位
jk5 = pd.DataFrame(jk_best.head(5)[4:5])#5位

item_list2 = [cd1,cd2,cd3,cd4,cd5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if jk1['アイテムCD'].values == ['03']:
  print(cd1)
  
  for i in jk1['商品名'].values:
    jk1_1 = (i)
 
  for i in jk1['数量'].values:
    jk1_2 = str(i)
  
  for i in jk1['金額'].values:
    jk1_3 = str(i)
  
  jk_cm1 = '1位 '+ jk1_1 + '\n'+ '点数 ' + jk1_2 + ' 金額 ¥'+ jk1_3
  
else:
  
  print('non item')
  jk_cm1 = str('1位 '+'NO ITEM')
  jk1_1 = "Non Item"
  jk1_2 = "Non Item"
  jk1_3 = "Non Item"   

jk_1 = pd.DataFrame({"商品名":[jk1_1] ,"点数":[jk1_2],"金額":[jk1_3]})  

#ーーーーーーーーー　２位　ーーーーーーーーーーー

if jk2['アイテムCD'].values == ['03']:
  
  print(jk2)
    
  for i in jk2['商品名'].values:
    jk2_1 = (i)
 
  for i in jk2['数量'].values:
    jk2_2 = str(i)
  
  for i in jk2['金額'].values:
    jk2_3 = str(i)
    jk_cm2 = '2位 '+ jk2_1 + '\n'+ '点数 ' + jk2_2 + ' 金額 ¥'+ jk2_3
  
else:
  print('non item')
  jk_cm2 = str('2位 '+'NO ITEM')
  jk2_1 = "Non Item"
  jk2_2 = "Non Item"
  jk2_3 = "Non Item"  

jk_2 = pd.DataFrame({"商品名":[jk2_1] ,"点数":[jk2_2],"金額":[jk2_3]})    
#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if jk3['アイテムCD'].values == ['03']:
  print(jk3)
    
  for i in jk3['商品名'].values:
    jk3_1 = (i)
 
  for i in jk3['数量'].values:
    jk3_2 = str(i)
  
  for i in jk3['金額'].values:
    jk3_3 = str(i)
    jk_cm3 = '3位 '+ jk3_1 + '\n'+ '点数 ' + jk3_2 + ' 金額 ¥'+ jk3_3
  
else:
  print('non item')
  jk_cm3 = str('3位 '+'NO ITEM')
  jk3_1 = "Non Item"
  jk3_2 = "Non Item"
  jk3_3 = "Non Item"

jk_3 = pd.DataFrame({"商品名":[jk3_1] ,"点数":[jk3_2],"金額":[jk3_3]})    
#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if jk4['アイテムCD'].values == ['03']:
    
  for i in jk4['商品名'].values:
    jk4_1 = (i)
 
  for i in jk4['数量'].values:
    jk4_2 = str(i)
  
  for i in jk4['金額'].values:
    jk4_3 = str(i)
    jk_cm4 = '4位 '+ jk4_1 + '\n'+ '点数 ' + jk4_2 + ' 金額 ¥'+ jk4_3
else:
  print('non item')
  jk_cm4 = str('4位 '+'NO ITEM')
  
  jk4_1 = "Non Item"
  jk4_2 = "Non Item"
  jk4_3 = "Non Item"


jk_4 = pd.DataFrame({"商品名":[jk4_1] ,"点数":[jk4_2],"金額":[jk4_3]})  
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if jk5['アイテムCD'].values == ['03']:
    
  for i in jk5['商品名'].values:
    jk5_1 = (i)
 
  for i in jk5['数量'].values:
    jk5_2 = str(i)
  
  for i in jk5['金額'].values:
    jk5_3 = str(i)
    jk_cm5 = '5位 '+ jk5_1 + '\n'+ '点数 ' + jk5_2 + ' 金額 ¥'+ jk5_3

else:
  print('non item')  
  jk_cm5 = str('5位 '+'NO ITEM')
  jk5_1 = "Non Item"
  jk5_2 = "Non Item"
  jk5_3 = "Non Item"
 
  
  

jk_5 = pd.DataFrame({"商品名":[jk5_1] ,"点数":[jk5_2],"金額":[jk5_3]})  

#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg3 = ('\n'+'【JK】ジャケット'+'\n'+(jk_cm1)+'\n'+'\n'+(jk_cm2)+'\n'+'\n'+(jk_cm3)+'\n'+'\n'+(jk_cm4)+'\n'+'\n'+(jk_cm5))


jk_list = pd.concat([jk_1,jk_2,jk_3,jk_4,jk_5],axis=0)

jk_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_3.xlsx')  

#ーーーーーーーー　END　ーーーーーーーーーーー

#ーーーーーーーーーー　アイテム　ーーーーーーーーーーKT

kt = df_1[df_1['アイテムCD'] == '04']
kt_best = pd.DataFrame(kt.sort_values(by='金額',ascending=False))

kt1 = pd.DataFrame(kt_best.head(5)[0:1])#1位
kt2 = pd.DataFrame(kt_best.head(5)[1:2])#2位
kt3 = pd.DataFrame(kt_best.head(5)[2:3])#3位
kt4 = pd.DataFrame(kt_best.head(5)[3:4])#4位
kt5 = pd.DataFrame(kt_best.head(5)[4:5])#5位

item_list2 = [cd1,cd2,cd3,cd4,cd5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if kt1['アイテムCD'].values == ['04']:
  print(cd1)
  
  for i in kt1['商品名'].values:
    kt1_1 = (i)
 
  for i in kt1['数量'].values:
    kt1_2 = str(i)
  
  for i in kt1['金額'].values:
    kt1_3 = str(i)
  
  kt_cm1 = '1位 '+ kt1_1 + '\n'+ '点数 ' + kt1_2 + ' 金額 ¥'+ kt1_3
else:
  print('non item')
  kt_cm1 = str('1位 '+'NO ITEM')

kt_1 = pd.DataFrame({"商品名":[kt1_1] ,"点数":[kt1_2],"金額":[kt1_3]})  

#ーーーーーーーーー　２位　ーーーーーーーーーーー

if kt2['アイテムCD'].values == ['04']:
  print(kt2)
    
  for i in kt2['商品名'].values:
    kt2_1 = (i)
 
  for i in kt2['数量'].values:
    kt2_2 = str(i)
  
  for i in kt2['金額'].values:
    kt2_3 = str(i)
    kt_cm2 = '2位 '+ kt2_1 + '\n'+ '点数 ' + kt2_2 + ' 金額 ¥'+ kt2_3
  
else:
  print('non item')
  kt_cm2 = str('2位 '+'NO ITEM')

kt_2 = pd.DataFrame({"商品名":[kt2_1] ,"点数":[kt2_2],"金額":[kt2_3]})  

#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if kt3['アイテムCD'].values == ['04']:
    
  for i in kt3['商品名'].values:
    kt3_1 = (i)
 
  for i in kt3['数量'].values:
    kt3_2 = str(i)
  
  for i in kt3['金額'].values:
    kt3_3 = str(i)
    kt_cm3 = '3位 '+ kt3_1 + '\n'+ '点数 ' + kt3_2 + ' 金額 ¥'+ kt3_3

else:
  print('non item')
  kt_cm3 = str('3位 '+'NO ITEM')

kt_3 = pd.DataFrame({"商品名":[kt3_1] ,"点数":[kt3_2],"金額":[kt3_3]})
#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if kt4['アイテムCD'].values == ['04']:
    
  for i in kt4['商品名'].values:
    kt4_1 = (i)
 
  for i in kt4['数量'].values:
    kt4_2 = str(i)
  
  for i in kt4['金額'].values:
    kt4_3 = str(i)
    kt_cm4 = '4位 '+ kt4_1 + '\n'+ '点数 ' + kt4_2 + ' 金額 ¥'+ kt4_3
else:
  print('non item')
  kt_cm4 = str('4位 '+'NO ITEM')

kt_4 = pd.DataFrame({"商品名":[kt4_1] ,"点数":[kt4_2],"金額":[kt4_3]})
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if kt5['アイテムCD'].values == ['04']:
    
  for i in kt5['商品名'].values:
    kt5_1 = (i)
 
  for i in kt5['数量'].values:
    kt5_2 = str(i)
  
  for i in kt5['金額'].values:
    kt5_3 = str(i)
    kt_cm5 = '5位 '+ kt5_1 + '\n'+ '点数 ' + kt5_2 + ' 金額 ¥'+ kt5_3

else:
  print('non item')  
  kt_cm5 = str('5位 '+'NO ITEM')

kt_5 = pd.DataFrame({"商品名":[kt5_1] ,"点数":[kt5_2],"金額":[kt5_3]})  

#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg4 = ('\n'+'【KT】ニット'+'\n'+(kt_cm1)+'\n'+'\n'+(kt_cm2)+'\n'+'\n'+(kt_cm3)+'\n'+'\n'+(kt_cm4)+'\n'+'\n'+(kt_cm5))

kt_list = pd.concat([kt_1,kt_2,kt_3,kt_4,kt_5],axis=0)

kt_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_4.xlsx')  

#ーーーーーーーー　END　ーーーーーーーーーーー

#ーーーーーーーーーー　アイテム　ーーーーーーーーーーCS

cs = df_1[df_1['アイテムCD'] == '05']
cs_best = pd.DataFrame(cs.sort_values(by='金額',ascending=False))

cs1 = pd.DataFrame(cs_best.head(5)[0:1])#1位
cs2 = pd.DataFrame(cs_best.head(5)[1:2])#2位
cs3 = pd.DataFrame(cs_best.head(5)[2:3])#3位
cs4 = pd.DataFrame(cs_best.head(5)[3:4])#4位
cs5 = pd.DataFrame(cs_best.head(5)[4:5])#5位

item_list2 = [cd1,cd2,cd3,cd4,cd5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if cs1['アイテムCD'].values == ['05']:
  print(cd1)
  
  for i in cs1['商品名'].values:
    cs1_1 = (i)
 
  for i in cs1['数量'].values:
    cs1_2 = str(i)
  
  for i in cs1['金額'].values:
    cs1_3 = str(i)
  
  cs_cm1 = '1位 '+ cs1_1 + '\n'+ '点数 ' + cs1_2 + ' 金額 ¥'+ cs1_3
else:
  print('non item')
  cs_cm1 = str('1位 '+'NO ITEM')

cs_1 = pd.DataFrame({"商品名":[cs1_1] ,"点数":[cs1_2],"金額":[cs1_3]})
#ーーーーーーーーー　２位　ーーーーーーーーーーー

if cs2['アイテムCD'].values == ['05']:
  print(kt2)
    
  for i in cs2['商品名'].values:
    cs2_1 = (i)
 
  for i in cs2['数量'].values:
    cs2_2 = str(i)
  
  for i in cs2['金額'].values:
    cs2_3 = str(i)
    cs_cm2 = '2位 '+ cs2_1 + '\n'+ '点数 ' + cs2_2 + ' 金額 ¥'+ cs2_3
  
else:
  print('non item')
  cs_cm2 = str('2位 '+'NO ITEM')

cs_2 = pd.DataFrame({"商品名":[cs2_1] ,"点数":[cs2_2],"金額":[cs2_3]})  
#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if cs3['アイテムCD'].values == ['05']:
    
  for i in cs3['商品名'].values:
    cs3_1 = (i)
 
  for i in cs3['数量'].values:
    cs3_2 = str(i)
  
  for i in cs3['金額'].values:
    cs3_3 = str(i)
    cs_cm3 = '3位 '+ cs3_1 + '\n'+ '点数 ' + cs3_2 + ' 金額 ¥'+ cs3_3

else:
  print('non item')
  cs_cm3 = str('3位 '+'NO ITEM')
  

cs_3 = pd.DataFrame({"商品名":[cs3_1] ,"点数":[cs3_2],"金額":[cs3_3]})  

#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if cs4['アイテムCD'].values == ['05']:
    
  for i in cs4['商品名'].values:
    cs4_1 = (i)
 
  for i in cs4['数量'].values:
    cs4_2 = str(i)
  
  for i in cs4['金額'].values:
    cs4_3 = str(i)
    cs_cm4 = '4位 '+ cs4_1 + '\n'+ '点数 ' + cs4_2 + ' 金額 ¥'+ cs4_3
else:
  print('non item')
  cs_cm4 = str('4位 '+'NO ITEM')

cs_4 = pd.DataFrame({"商品名":[cs4_1] ,"点数":[cs4_2],"金額":[cs4_3]})
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if cs5['アイテムCD'].values == ['05']:
    
  for i in cs5['商品名'].values:
    cs5_1 = (i)
 
  for i in cs5['数量'].values:
    cs5_2 = str(i)
  
  for i in cs5['金額'].values:
    cs5_3 = str(i)
    cs_cm5 = '5位 '+ cs5_1 + '\n'+ '点数 ' + cs5_2 + ' 金額 ¥'+ cs5_3

else:
  print('non item')  
  cs_cm5 = str('5位 '+'NO ITEM')
  

cs_5 = pd.DataFrame({"商品名":[cs5_1] ,"点数":[cs5_2],"金額":[cs5_3]})  

#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg5 = ('\n'+'【CS】カットソー'+'\n'+(cs_cm1)+'\n'+'\n'+(cs_cm2)+'\n'+'\n'+(cs_cm3)+'\n'+'\n'+(cs_cm4)+'\n'+'\n'+(cs_cm5))

cs_list = pd.concat([cs_1,cs_2,cs_3,cs_4,cs_5],axis=0)

cs_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_5.xlsx')  

#ーーーーーーーー　END　ーーーーーーーーーーー
#CTがあります【06】
#★★

ct = df_1[df_1['アイテムCD'] == '06']
ct_best = pd.DataFrame(ct.sort_values(by='金額',ascending=False))

ct1 = pd.DataFrame(ct_best.head(5)[0:1])#1位
ct2 = pd.DataFrame(ct_best.head(5)[1:2])#2位
ct3 = pd.DataFrame(ct_best.head(5)[2:3])#3位
ct4 = pd.DataFrame(ct_best.head(5)[3:4])#4位
ct5 = pd.DataFrame(ct_best.head(5)[4:5])#5位

item_list2 = [ct1,ct2,ct3,ct4,ct5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if ct1['アイテムCD'].values == ['06']:
  print(ct1)
  
  for i in ct1['商品名'].values:
    ct1_1 = (i)
  
 
  for i in ct1['数量'].values:
    ct1_2 = str(i)
  
  for i in ct1['金額'].values:
    ct1_3 = str(i)
  
  ct_cm1 = '1位 '+ ct1_1 + '\n'+ '点数 ' + ct1_2 + ' 金額 ¥'+ ct1_3
else:
  print('non item')
  ct_cm1 = str('1位 '+'NO ITEM')
  ct1_1 = "Non Item"
  ct1_2 = "Non Item"
  ct1_3 = "Non Item"

ct_1 = pd.DataFrame({"商品名":[ct1_1] ,"点数":[ct1_2],"金額":[ct1_3]})
print(ct_1)
#ーーーーーーーーー　２位　ーーーーーーーーーーー

if ct2['アイテムCD'].values == ['06']:
  print(kt2)
    
  for i in ct2['商品名'].values:
    ct2_1 = (i)
 
  for i in ct2['数量'].values:
    ct2_2 = str(i)
  
  for i in ct2['金額'].values:
    ct2_3 = str(i)
    ct_cm2 = '2位 '+ ct2_1 + '\n'+ '点数 ' + ct2_2 + ' 金額 ¥'+ ct2_3
  
else:
  print('non item')
  ct_cm2 = str('2位 '+'NO ITEM')
  
  ct2_1 = "Non Item"
  ct2_2 = "Non Item"
  ct2_3 = "Non Item"
  
  

ct_2 = pd.DataFrame({"商品名":[ct2_1] ,"点数":[ct2_2],"金額":[ct2_3]})  
#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if ct3['アイテムCD'].values == ['06']:
    
  for i in ct3['商品名'].values:
    ct3_1 = (i)
 
  for i in ct3['数量'].values:
    ct3_2 = str(i)
  
  for i in ct3['金額'].values:
    ct3_3 = str(i)
    ct_cm3 = '3位 '+ ct3_1 + '\n'+ '点数 ' + ct3_2 + ' 金額 ¥'+ ct3_3

else:
  print('non item')
  ct_cm3 = str('3位 '+'NO ITEM')
  
  ct3_1 = "Non Item"
  ct3_2 = "Non Item"
  ct3_3 = "Non Item"
  

ct_3 = pd.DataFrame({"商品名":[ct3_1] ,"点数":[ct3_2],"金額":[ct3_3]})  

#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if ct4['アイテムCD'].values == ['06']:
    
  for i in ct4['商品名'].values:
    ct4_1 = (i)
 
  for i in ct4['数量'].values:
    ct4_2 = str(i)
  
  for i in ct4['金額'].values:
    ct4_3 = str(i)
    ct_cm4 = '4位 '+ ct4_1 + '\n'+ '点数 ' + ct4_2 + ' 金額 ¥'+ ct4_3
else:
  print('non item')
  ct_cm4 = str('4位 '+'NO ITEM')
  
  ct4_1 = "Non Item"
  ct4_2 = "Non Item"
  ct4_3 = "Non Item"

ct_4 = pd.DataFrame({"商品名":[ct4_1] ,"点数":[ct4_2],"金額":[ct4_3]})
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if ct5['アイテムCD'].values == ['06']:
    
  for i in ct5['商品名'].values:
    ct5_1 = (i)
 
  for i in ct5['数量'].values:
    ct5_2 = str(i)
  
  for i in ct5['金額'].values:
    ct5_3 = str(i)
    ct_cm5 = '5位 '+ ct5_1 + '\n'+ '点数 ' + ct5_2 + ' 金額 ¥'+ ct5_3

else:
  print('non item')  
  ct_cm5 = str('5位 '+'NO ITEM')
  
  ct5_1 = "Non Item"
  ct5_2 = "Non Item"
  ct5_3 = "Non Item"
  

ct_5 = pd.DataFrame({"商品名":[ct5_1] ,"点数":[ct5_2],"金額":[ct5_3]})  

#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg6 = ('\n'+'【CT】コート'+'\n'+(ct_cm1)+'\n'+'\n'+(ct_cm2)+'\n'+'\n'+(ct_cm3)+'\n'+'\n'+(ct_cm4)+'\n'+'\n'+(ct_cm5))


ct_list = pd.concat([ct_1,ct_2,ct_3,ct_4,ct_5],axis=0)

ct_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_6.xlsx')  

#ーーーーーーーーーー　アイテム　ーーーーーーーーーーCS

bl = df_1[df_1['アイテムCD'] == '07']
bl_best = pd.DataFrame(bl.sort_values(by='金額',ascending=False))

bl1 = pd.DataFrame(bl_best.head(5)[0:1])#1位
bl2 = pd.DataFrame(bl_best.head(5)[1:2])#2位
bl3 = pd.DataFrame(bl_best.head(5)[2:3])#3位
bl4 = pd.DataFrame(bl_best.head(5)[3:4])#4位
bl5 = pd.DataFrame(bl_best.head(5)[4:5])#5位

item_list2 = [cd1,cd2,cd3,cd4,cd5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if bl1['アイテムCD'].values == ['07']:
  print(bl1)
  
  for i in bl1['商品名'].values:
    bl1_1 = (i)
 
  for i in bl1['数量'].values:
    bl1_2 = str(i)
  
  for i in bl1['金額'].values:
    bl1_3 = str(i)
  
  bl_cm1 = '1位 '+ bl1_1 + '\n'+ '点数 ' + bl1_2 + ' 金額 ¥'+ bl1_3
else:
  print('non item')
  bl_cm1 = str('1位 '+'NO ITEM')


bl_1 = pd.DataFrame({"商品名":[bl1_1] ,"点数":[bl1_2],"金額":[bl1_3]})  
#ーーーーーーーーー　２位　ーーーーーーーーーーー

if bl2['アイテムCD'].values == ['07']:
  print(kt2)
    
  for i in bl2['商品名'].values:
    bl2_1 = (i)
 
  for i in bl2['数量'].values:
    bl2_2 = str(i)
  
  for i in bl2['金額'].values:
    bl2_3 = str(i)
    bl_cm2 = '2位 '+ bl2_1 + '\n'+ '点数 ' + bl2_2 + ' 金額 ¥'+ bl2_3
  
else:
  print('non item')
  bl_cm2 = str('2位 '+'NO ITEM')

bl_2 = pd.DataFrame({"商品名":[bl2_1] ,"点数":[bl2_2],"金額":[bl2_3]})  
#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if bl3['アイテムCD'].values == ['07']:
    
  for i in bl3['商品名'].values:
    bl3_1 = (i)
 
  for i in bl3['数量'].values:
    bl3_2 = str(i)
  
  for i in bl3['金額'].values:
    bl3_3 = str(i)
    bl_cm3 = '3位 '+ bl3_1 + '\n'+ '点数 ' + bl3_2 + ' 金額 ¥'+ bl3_3

else:
  print('non item')
  bl_cm3 = str('3位 '+'NO ITEM')


bl_3 = pd.DataFrame({"商品名":[bl3_1] ,"点数":[bl3_2],"金額":[bl3_3]})
#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if bl4['アイテムCD'].values == ['07']:
    
  for i in bl4['商品名'].values:
    bl4_1 = (i)
 
  for i in bl4['数量'].values:
    bl4_2 = str(i)
  
  for i in bl4['金額'].values:
    bl4_3 = str(i)
    bl_cm4 = '4位 '+ bl4_1 + '\n'+ '点数 ' + bl4_2 + ' 金額 ¥'+ bl4_3
else:
  print('non item')
  bl_cm4 = str('4位 '+'NO ITEM')

bl_4 = pd.DataFrame({"商品名":[bl4_1] ,"点数":[bl4_2],"金額":[bl4_3]})
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if bl5['アイテムCD'].values == ['07']:
    
  for i in bl5['商品名'].values:
    bl5_1 = (i)
 
  for i in bl5['数量'].values:
    bl5_2 = str(i)
  
  for i in bl5['金額'].values:
    bl5_3 = str(i)
    
    bl_cm5 = '5位 '+ bl5_1 + '\n'+ '点数 ' + bl5_2 + ' 金額 ¥'+ bl5_3

else:
  print('non item')  
  bl_cm5 = str('5位 '+'NO ITEM')

bl_5 = pd.DataFrame({"商品名":[bl5_1] ,"点数":[bl5_2],"金額":[bl5_3]})
#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg7 = ('\n'+'【BL】ブラウス'+'\n'+(bl_cm1)+'\n'+'\n'+(bl_cm2)+'\n'+'\n'+(bl_cm3)+'\n'+'\n'+(bl_cm4)+'\n'+'\n'+(bl_cm5))


bl_list = pd.concat([bl_1,bl_2,bl_3,bl_4,bl_5],axis=0)

bl_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_7.xlsx')  

#ーーーーーーーー　END　ーーーーーーーーーーー

#ーーーーーーーーーー　アイテム　ーーーーーーーーーーCS

sk = df_1[df_1['アイテムCD'] == '08']
sk_best = pd.DataFrame(sk.sort_values(by='金額',ascending=False))

sk1 = pd.DataFrame(sk_best.head(5)[0:1])#1位
sk2 = pd.DataFrame(sk_best.head(5)[1:2])#2位
sk3 = pd.DataFrame(sk_best.head(5)[2:3])#3位
sk4 = pd.DataFrame(sk_best.head(5)[3:4])#4位
sk5 = pd.DataFrame(sk_best.head(5)[4:5])#5位

item_list2 = [cd1,cd2,cd3,cd4,cd5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if sk1['アイテムCD'].values == ['08']:
  print(sk1)
  
  for i in sk1['商品名'].values:
    sk1_1 = (i)
 
  for i in sk1['数量'].values:
    sk1_2 = str(i)
  
  for i in sk1['金額'].values:
    sk1_3 = str(i)
  
  sk_cm1 = '1位 '+ sk1_1 + '\n'+ '点数 ' + sk1_2 + ' 金額 ¥'+ sk1_3
else:
  print('non item')
  sk_cm1 = str('1位 '+'NO ITEM')

sk_1 = pd.DataFrame({"商品名":[sk1_1] ,"点数":[sk1_2],"金額":[sk1_3]})

#ーーーーーーーーー　２位　ーーーーーーーーーーー

if sk2['アイテムCD'].values == ['08']:
  print(kt2)
    
  for i in sk2['商品名'].values:
    sk2_1 = (i)
 
  for i in sk2['数量'].values:
    sk2_2 = str(i)
  
  for i in sk2['金額'].values:
    sk2_3 = str(i)
    sk_cm2 = '2位 '+ sk2_1 + '\n'+ '点数 ' + sk2_2 + ' 金額 ¥'+ sk2_3
  
else:
  print('non item')
  sk_cm2 = str('2位 '+'NO ITEM')

sk_2 = pd.DataFrame({"商品名":[sk2_1] ,"点数":[sk2_2],"金額":[sk2_3]})
  
#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if sk3['アイテムCD'].values == ['08']:
    
  for i in sk3['商品名'].values:
    sk3_1 = (i)
 
  for i in sk3['数量'].values:
    sk3_2 = str(i)
  
  for i in sk3['金額'].values:
    sk3_3 = str(i)
    sk_cm3 = '3位 '+ sk3_1 + '\n'+ '点数 ' + sk3_2 + ' 金額 ¥'+ sk3_3

else:
  print('non item')
  sk_cm3 = str('3位 '+'NO ITEM')

sk_3 = pd.DataFrame({"商品名":[sk3_1] ,"点数":[sk3_2],"金額":[sk3_3]})
#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if sk4['アイテムCD'].values == ['08']:
    
  for i in sk4['商品名'].values:
    sk4_1 = (i)
 
  for i in sk4['数量'].values:
    sk4_2 = str(i)
  
  for i in sk4['金額'].values:
    sk4_3 = str(i)
    sk_cm4 = '4位 '+ sk4_1 + '\n'+ '点数 ' + sk4_2 + ' 金額 ¥'+ sk4_3
else:
  print('non item')
  sk_cm4 = str('4位 '+'NO ITEM')

sk_4 = pd.DataFrame({"商品名":[sk4_1] ,"点数":[sk4_2],"金額":[sk4_3]})
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if sk5['アイテムCD'].values == ['08']:
    
  for i in sk5['商品名'].values:
    sk5_1 = (i)
 
  for i in sk5['数量'].values:
    sk5_2 = str(i)
  
  for i in sk5['金額'].values:
    sk5_3 = str(i)
    sk_cm5 = '5位 '+ sk5_1 + '\n'+ '点数 ' + sk5_2 + ' 金額 ¥'+ sk5_3

else:
  print('non item')  
  sk_cm5 = str('5位 '+'NO ITEM')


sk_5 = pd.DataFrame({"商品名":[sk5_1] ,"点数":[sk5_2],"金額":[sk5_3]})  

#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg8 = ('\n'+'【SK】スカート'+'\n'+(sk_cm1)+'\n'+'\n'+(sk_cm2)+'\n'+'\n'+(sk_cm3)+'\n'+'\n'+(sk_cm4)+'\n'+'\n'+(sk_cm5))


sk_list = pd.concat([sk_1,sk_2,sk_3,sk_4,sk_5],axis=0)

sk_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_8.xlsx') 

#ーーーーーーーーーー　アイテム　ーーーーーーーーーー

pt = df_1[df_1['アイテムCD'] == '09']
pt_best = pd.DataFrame(pt.sort_values(by='金額',ascending=False))

pt1 = pd.DataFrame(pt_best.head(5)[0:1])#1位
pt2 = pd.DataFrame(pt_best.head(5)[1:2])#2位
pt3 = pd.DataFrame(pt_best.head(5)[2:3])#3位
pt4 = pd.DataFrame(pt_best.head(5)[3:4])#4位
pt5 = pd.DataFrame(pt_best.head(5)[4:5])#5位

item_list2 = [cd1,cd2,cd3,cd4,cd5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if pt1['アイテムCD'].values == ['09']:
  print(sk1)
  
  for i in pt1['商品名'].values:
    pt1_1 = (i)
 
  for i in pt1['数量'].values:
    pt1_2 = str(i)
  
  for i in pt1['金額'].values:
    pt1_3 = str(i)
  
  pt_cm1 = '1位 '+ pt1_1 + '\n'+ '点数 ' + pt1_2 + ' 金額 ¥'+ pt1_3
else:
  print('non item')
  pt_cm1 = str('1位 '+'NO ITEM')

pt_1 = pd.DataFrame({"商品名":[pt1_1] ,"点数":[pt1_2],"金額":[pt1_3]})

#ーーーーーーーーー　２位　ーーーーーーーーーーー

if pt2['アイテムCD'].values == ['09']:
  print(kt2)
    
  for i in pt2['商品名'].values:
    pt2_1 = (i)
 
  for i in pt2['数量'].values:
    pt2_2 = str(i)
  
  for i in pt2['金額'].values:
    pt2_3 = str(i)
    pt_cm2 = '2位 '+ pt2_1 + '\n'+ '点数 ' + pt2_2 + ' 金額 ¥'+ pt2_3
  
else:
  print('non item')
  pt_cm2 = str('2位 '+'NO ITEM')

pt_2 = pd.DataFrame({"商品名":[pt2_1] ,"点数":[pt2_2],"金額":[pt2_3]})  
#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if pt3['アイテムCD'].values == ['09']:
    
  for i in pt3['商品名'].values:
    
    pt3_1 = (i)
 
  for i in pt3['数量'].values:
    
    pt3_2 = str(i)
  
  for i in pt3['金額'].values:
    
    pt3_3 = str(i)
    
    pt_cm3 = '3位 '+ pt3_1 + '\n'+ '点数 ' + pt3_2 + ' 金額 ¥'+ pt3_3

else:
  print('non item')
  pt_cm3 = str('3位 '+'NO ITEM')

pt_3 = pd.DataFrame({"商品名":[pt3_1] ,"点数":[pt3_2],"金額":[pt3_3]})
#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if pt4['アイテムCD'].values == ['09']:
    
  for i in pt4['商品名'].values:
    pt4_1 = (i)
 
  for i in pt4['数量'].values:
    pt4_2 = str(i)
  
  for i in pt4['金額'].values:
    pt4_3 = str(i)
    pt_cm4 = '4位 '+ pt4_1 + '\n'+ '点数 ' + pt4_2 + ' 金額 ¥'+ pt4_3
else:
  print('non item')
  pt_cm4 = str('4位 '+'NO ITEM')

pt_4 = pd.DataFrame({"商品名":[pt4_1] ,"点数":[pt4_2],"金額":[pt4_3]})
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if pt5['アイテムCD'].values == ['09']:
    
  for i in pt5['商品名'].values:
    pt5_1 = (i)
 
  for i in pt5['数量'].values:
    pt5_2 = str(i)
  
  for i in pt5['金額'].values:
    pt5_3 = str(i)
    pt_cm5 = '5位 '+ pt5_1 + '\n'+ '点数 ' + pt5_2 + ' 金額 ¥'+ pt5_3

else:
  print('non item')  
  pt_cm5 = str('5位 '+'NO ITEM')
  

pt_5 = pd.DataFrame({"商品名":[pt5_1] ,"点数":[pt5_2],"金額":[pt5_3]})  

#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg9 = ('\n'+'【PT】パンツ'+'\n'+(pt_cm1)+'\n'+'\n'+(pt_cm2)+'\n'+'\n'+(pt_cm3)+'\n'+'\n'+(pt_cm4)+'\n'+'\n'+(pt_cm5))


pt_5 = pd.DataFrame({"商品名":[pt5_1] ,"点数":[pt5_2],"金額":[pt5_3]})

pt_list = pd.concat([pt_1,pt_2,pt_3,pt_4,pt_5],axis=0)

pt_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_9.xlsx') 

#ーーーーーーーーーー　アイテム　ーーーーーーーーーー

tr = df_1[df_1['アイテムCD'] == '10']
tr_best = pd.DataFrame(tr.sort_values(by='金額',ascending=False))

tr1 = pd.DataFrame(tr_best.head(5)[0:1])#1位
tr2 = pd.DataFrame(tr_best.head(5)[1:2])#2位
tr3 = pd.DataFrame(tr_best.head(5)[2:3])#3位
tr4 = pd.DataFrame(tr_best.head(5)[3:4])#4位
tr5 = pd.DataFrame(tr_best.head(5)[4:5])#5位

item_list2 = [cd1,cd2,cd3,cd4,cd5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if tr1['アイテムCD'].values == ['10']:
  print(tr1)
  
  for i in tr1['商品名'].values:
    tr1_1 = (i)
 
  for i in tr1['数量'].values:
    tr1_2 = str(i)
  
  for i in tr1['金額'].values:
    tr1_3 = str(i)
  
  tr_cm1 = '1位 '+ tr1_1 + '\n'+ '点数 ' + tr1_2 + ' 金額 ¥'+ tr1_3
else:
  print('non item')
  tr_cm1 = str('1位 '+'NO ITEM')

tr_1 = pd.DataFrame({"商品名":[tr1_1] ,"点数":[tr1_2],"金額":[tr1_3]})

#ーーーーーーーーー　２位　ーーーーーーーーーーー

if tr2['アイテムCD'].values == ['10']:
  print(tr2)
    
  for i in tr2['商品名'].values:
    tr2_1 = (i)
 
  for i in tr2['数量'].values:
    tr2_2 = str(i)
  
  for i in tr2['金額'].values:
    tr2_3 = str(i)
    tr_cm2 = '2位 '+ tr2_1 + '\n'+ '点数 ' + tr2_2 + ' 金額 ¥'+ tr2_3
  
else:
  print('non item')
  tr_cm2 = str('2位 '+'NO ITEM')
  tr2_1 = "Non Item"
  tr2_2 = "0"
  tr2_3 = "0"

tr_2 = pd.DataFrame({"商品名":[tr2_1] ,"点数":[tr2_2],"金額":[tr2_3]})  
#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if tr3['アイテムCD'].values == ['10']:
    
  for i in tr3['商品名'].values:
    tr3_1 = (i)
 
  for i in tr3['数量'].values:
    tr3_2 = str(i)
  
  for i in tr3['金額'].values:
    tr3_3 = str(i)
    tr_cm3 = '3位 '+ tr3_1 + '\n'+ '点数 ' + tr3_2 + ' 金額 ¥'+ tr3_3

else:
  print('non item')
  tr_cm3 = str('3位 '+'NO ITEM')
  tr3_1 = "Non Item"
  tr3_2 = "0"
  tr3_3 = "0"

tr_3 = pd.DataFrame({"商品名":[tr3_1] ,"点数":[tr3_2],"金額":[tr3_3]})
#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if tr4['アイテムCD'].values == ['10']:
    
  for i in tr4['商品名'].values:
    tr4_1 = (i)
 
  for i in tr4['数量'].values:
    tr4_2 = str(i)
  
  for i in tr4['金額'].values:
    tr4_3 = str(i)
    tr_cm4 = '4位 '+ tr4_1 + '\n'+ '点数 ' + tr4_2 + ' 金額 ¥'+ tr4_3
else:
  print('non item')
  tr_cm4 = str('4位 '+'NO ITEM')
  tr4_1 = "Non Item"
  tr4_2 = "0"
  tr4_3 = "0"

tr_4 = pd.DataFrame({"商品名":[tr4_1] ,"点数":[tr4_2],"金額":[tr4_3]})
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if tr5['アイテムCD'].values == ['10']:
    
  for i in tr5['商品名'].values:
    tr5_1 = (i)
 
  for i in tr5['数量'].values:
    tr5_2 = str(i)
  
  for i in tr5['金額'].values:
    tr5_3 = str(i)
    tr_cm5 = '5位 '+ tr5_1 + '\n'+ '点数 ' + tr5_2 + ' 金額 ¥'+ tr5_3

else:
  
  print('non item')  
  tr_cm5 = str('5位 '+'NO ITEM')
  tr5_1 = "Non Item"
  tr5_2 = "0"
  tr5_3 = "0"
  
tr_5 = pd.DataFrame({"商品名":[tr5_1] ,"点数":[tr5_2],"金額":[tr5_3]})  

#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg10 = ('\n'+'【TR】トレーナー'+'\n'+(tr_cm1)+'\n'+'\n'+(tr_cm2)+'\n'+'\n'+(tr_cm3)+'\n'+'\n'+(tr_cm4)+'\n'+'\n'+(tr_cm5))


tr_5 = pd.DataFrame({"商品名":[tr5_1] ,"点数":[tr5_2],"金額":[tr5_3]})

tr_list = pd.concat([tr_1,tr_2,tr_3,tr_4,tr_5],axis=0)

tr_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_10.xlsx') 



#ーーーーーーーーーー　アイテム　ーーーーーーーーーー
#セットアップ
print(df_1)
setup = df_1[df_1['アイテムCD'] == '12']
print(setup)
setup_best = pd.DataFrame(setup.sort_values(by='金額',ascending=False))
print(setup_best)

setup1 = pd.DataFrame(setup_best.head(5)[0:1])#1位
setup2 = pd.DataFrame(setup_best.head(5)[1:2])#2位
setup3 = pd.DataFrame(setup_best.head(5)[2:3])#3位
setup4 = pd.DataFrame(setup_best.head(5)[3:4])#4位
setup5 = pd.DataFrame(setup_best.head(5)[4:5])#5位

item_list2 = [setup1,setup2,setup3,setup4,setup5]

#ーーーーーーーーー　1位　ーーーーーーーーーーー

if setup1['アイテムCD'].values == ['12']:
  print(setup1)
  
  for i in setup1['商品名'].values:
    setup1_1 = (i)
 
  for i in setup1['数量'].values:
    setup1_2 = str(i)
  
  for i in setup1['金額'].values:
    setup1_3 = str(i)
  
  setup_cm1 = '1位 '+ setup1_1 + '\n'+ '点数 ' + setup1_2 + ' 金額 ¥'+ setup1_3
else:
  print('non item')
  setup_cm1 = str('1位 '+'NO ITEM')
  setup1_1 = "Non Item"
  setup1_2 = "0"
  setup1_3 = "0"


setup_1 = pd.DataFrame({"商品名":[setup1_1] ,"点数":[setup1_2],"金額":[setup1_3]})

#ーーーーーーーーー　２位　ーーーーーーーーーーー

if setup2['アイテムCD'].values == ['12']:
  print(setup2)
    
  for i in setup2['商品名'].values:
    setup2_1 = (i)
 
  for i in setup2['数量'].values:
    setup2_2 = str(i)
  
  for i in setup2['金額'].values:
    setup2_3 = str(i)
    setup_cm2 = '2位 '+ setup2_1 + '\n'+ '点数 ' + setup2_2 + ' 金額 ¥'+ setup2_3
  
else:
  print('non item')
  setup_cm2 = str('2位 '+'NO ITEM')
  setup2_1 = "Non Item"
  setup2_2 = "0"
  setup2_3 = "0"

setup_2 = pd.DataFrame({"商品名":[setup2_1] ,"点数":[setup2_2],"金額":[setup2_3]})  

#ーーーーーーーーー　３位　ーーーーーーーーーーー  

if setup3['アイテムCD'].values == ['12']:
    
  for i in setup3['商品名'].values:
    setup3_1 = (i)
 
  for i in setup3['数量'].values:
    setup3_2 = str(i)
  
  for i in setup3['金額'].values:
    setup3_3 = str(i)
    setup_cm3 = '3位 '+ setup3_1 + '\n'+ '点数 ' + setup3_2 + ' 金額 ¥'+ setup3_3

else:
  print('non item')
  setup_cm3 = str('3位 '+'NO ITEM')
  setup3_1 = "Non Item"
  setup3_2 = "0"
  setup3_3 = "0"
  
  

setup_3 = pd.DataFrame({"商品名":[setup3_1] ,"点数":[setup3_2],"金額":[setup3_3]})
#ーーーーーーーーー　４位　ーーーーーーーーーーー    
  
  
if setup4['アイテムCD'].values == ['12']:
    
  for i in setup4['商品名'].values:
    setup4_1 = (i)
 
  for i in setup4['数量'].values:
    setup4_2 = str(i)
  
  for i in setup4['金額'].values:
    setup4_3 = str(i)
    setup_cm4 = '4位 '+ setup4_1 + '\n'+ '点数 ' + setup4_2 + ' 金額 ¥'+ setup4_3
else:
  print('non item')
  setup_cm4 = str('4位 '+'NO ITEM')
  setup4_1 = "Non Item"
  setup4_2 = "0"
  setup4_3 = "0"

setup_4 = pd.DataFrame({"商品名":[setup4_1] ,"点数":[setup4_2],"金額":[setup4_3]})
    
#ーーーーーーーーー　５位　ーーーーーーーーーーー    

if setup5['アイテムCD'].values == ['12']:
    
  for i in setup5['商品名'].values:
    setup5_1 = (i)
 
  for i in setup5['数量'].values:
    setup5_2 = str(i)
  
  for i in setup5['金額'].values:
    setup5_3 = str(i)
    setup_cm5 = '5位 '+ setup5_1 + '\n'+ '点数 ' + setup5_2 + ' 金額 ¥'+ setup5_3

else:
  print('non item')  
  setup_cm5 = str('5位 '+'NO ITEM')
  setup5_1 = "Non Item"
  setup5_2 = "0"
  setup5_3 = "0"
  

setup_5 = pd.DataFrame({"商品名":[setup5_1] ,"点数":[setup5_2],"金額":[setup5_3]})  


#ーーーーーーー　LINE通知メッセージ　ーーーーーーーーーー

msg12 = ('\n'+'【SET UP】セットアップ'+'\n'+(setup_cm1)+'\n'+'\n'+(setup_cm2)+'\n'+'\n'+(setup_cm3)+'\n'+'\n'+(setup_cm4)+'\n'+'\n'+(setup_cm5))


setup_5 = pd.DataFrame({"商品名":[setup5_1] ,"点数":[setup5_2],"金額":[setup5_3]})

setup_list = pd.concat([setup_1,setup_2,setup_3,setup_4,setup_5],axis=0)

setup_list.to_excel('C:/Users/fun-f/Desktop/myfile/basket-analysis/key_item_12.xlsx') 

#--------------------------------------------------------------------------

sh = df_1[df_1['アイテムCD'] == '15']
sp = df_1[df_1['アイテムCD'] == '98']

#sort_v.to_excel('C:/Users/fun-f/Desktop/myfile/テスト.xlsx')i 

comm = '今日の全店売上実績になります！'
rv = pd.read_excel('C:/Users/fun-f/Desktop/myfile/テスト.xlsx')
if switch == str(0):
  TOKEN = 'TNKXBcpEMmK4JAmRPaOyVABkA5GWIJkIHQOnsyfu4MD'#FUNの部屋トークン
  #TOKEN = 'NxPDQg0tpI4oY6BZHo7vkZ3gxPtTIijpWFyN85xL2q1'#テストの部屋トークン

else:  
  
  TOKEN = 'NxPDQg0tpI4oY6BZHo7vkZ3gxPtTIijpWFyN85xL2q1'#テストの部屋トークン
  
#TOKEN = 'NxPDQg0tpI4oY6BZHo7vkZ3gxPtTIijpWFyN85xL2q1'#テストの部屋トークン
#TOKEN = 'TNKXBcpEMmK4JAmRPaOyVABkA5GWIJkIHQOnsyfu4MD'#FUNの部屋トークン
api_url = 'https://notify-api.line.me/api/notify'
headers = {'Authorization' : 'Bearer ' + TOKEN}
#message = ('\n'+'柏'+'\n'+'【売上予算/実績】'+'\n' + str(mg1) +'\n' +'【P率】' +str(p1) +'\n'+ '【客数】'+ str(noc_2) +str(p2)+str(p3))
#トップスカテゴリに変更【CS、KT、BL、TR】
#message = ('\n'+'[アイテム別ベストセラー Part１]'+'\n'+(comm)+'\n'+(msg5)+'\n'+'\n'+(msg7)+'\n'+'\n'+(msg8)+'\n'+'\n'+(msg9))#+'\n'+'\n'+(msg5)+'\n'+'\n'+(msg7)+'\n'+'\n'+(msg8)+'\n'+'\n'+(msg9)+'\n'+'\n'+(msg7)+'\n'+'\n'+(msg7))
#+'\n'+(msg1))+'\n'

#payload = {'message': message}

#requests.post(api_url, headers=headers, params=payload)   


#ーーーーーー 第1メッセージ ーーーーーーー
#ボトムスカテゴリに変更【PT、SK】
message_1 = ('\n'+'[アイテム別ベストセラー トップス部門]'+'\n'+(comm)+'\n'+(msg4)+'\n'+'\n'+(msg5)+'\n'+'\n'+(msg7)+'\n')#+'\n'+(msg12))#+'\n'+'\n'+(msg9))
payload = {'message': message_1}

requests.post(api_url, headers=headers, params=payload)  

#ーーーーーー 第2メッセージ ーーーーーーー
#ボトムスカテゴリに変更【PT、SK】
message_2 = ('\n'+'[アイテム別ベストセラー ボトムス/OP,TR部門]'+'\n'+(comm)+'\n'+(msg8)+'\n'+'\n'+(msg9)+'\n'+'\n'+(msg1)+'\n'+'\n'+(msg10))#+'\n'+'\n'+(msg9))
payload = {'message': message_2}


requests.post(api_url, headers=headers, params=payload)  
#ーーーーーー 第3メッセージ ーーーーーーー
#羽織カテゴリに変更【CD、JK、CT】

message_3 = ('\n'+'[アイテム別ベストセラー 羽織部門]'+'\n'+(comm)+'\n'+(msg2)+'\n'+'\n'+(msg3)+'\n'+'\n'+(msg6)+'\n')#+'\n'+(msg12))#+'\n'+'\n'+(msg9))



payload = {'message': message_3}

requests.post(api_url, headers=headers, params=payload)  
print("SUCCESSFULL!!")
  

#while True:
  #schedule.run_pending()
  #time.sleep(1)  
  