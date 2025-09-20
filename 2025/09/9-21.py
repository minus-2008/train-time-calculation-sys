#駅間時間計算

# 認証のためのコード
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 認証情報の設定
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\wbc54\\Documents\\python\\wbcAPIkey.json", scope)

# クライアントの作成
client = gspread.authorize(creds)

# スプレッドシートの取得
ss = client.open('駅間所要時分計算ツール用試験台')
#以上まではローカル環境用認証コード

target = ss.worksheet("シート1")
cars1 = ss.worksheet("車両加速力スプレットシート")
cars2 = ss.worksheet("車両減速力スプレットシート")

import time#API管制用(臨時措置)
from matplotlib import pyplot as pyp#グラフ描画
pyp.ioff()
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# プログラムの注意：四捨五入はpython標準のraund関数を使用しており、仕様上、.5の場合、大きい方ではなく偶数の方が出力されます
#                   例：1.5→2.0, 2.5→2.0

def cell_exchange(x,y): #セルの変換
  line='' #変換先を作成
  alphabet_list=['Z','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y']#変換リスト
  while True: #列を指定
    line_value=x%26 #余りを求める
    x=(x-1)//26 #商
    line_add=alphabet_list[line_value] #読み込み
    line=line_add+line #書き込み
    if x==0:
      break #条件で終了
  answer=line+str(y)
  return answer

def read_limit(start,end):  #速度制限読み込み
  answer=target.get(start+':'+end)
  for i in range(0,len(answer)):
    for j in range(0,len(answer[i])):
      answer[i][j]=int(answer[i][j])
  return answer

def limit_exchange(a): #制限リスト変換
  for i in range(len(a)):
    temp1=a[i]#操作対象

def calculate_air_resistance(C_d, C_f, A, S, rho, V):#空気抵抗計算
    D_p=0.5*C_d*rho*(V**2)*A    # 圧力抵抗
    D_f=0.5*C_f*rho*(V**2)*S    # 摩擦抵抗の計算
    D_total = D_p + D_f# 合計の空気抵抗
    return D_total

def collor():#グラフ色
  collorlist=[(1,0,0),(1,0.5,0),(1,1,0),(0.5,1,0),(0,1,0),(0,1,0.5),(0,1,1),(0,0.5,1),(0,0,1),(0.5,0,1),(1,0,1),(1,0,0.5)]
  if not hasattr(collor,'Number'):
    collor.Number=0
  try:
    collorcode=collorlist[collor.Number]
  except IndexError:
    collorcode=collorlist[0]
    collor.Number=-1
  collor.Number=collor.Number+1
  return collorcode

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
print('〜〜〜〜〜〜〜〜〜〜〜')
print('start to read speed limit')
lr_range=target.range('A1:B2')
limitrange_start=int(lr_range[0].value)  #計算、制限読み込み範囲の指定
limitrange_end=int(lr_range[1].value)
limitrange_number=int((limitrange_end-limitrange_start)/2+1)
runningrange_start=int(lr_range[2].value)
runningrange_end=int(lr_range[3].value)
runningrange_number=int((runningrange_end-runningrange_start+1)/2)   #計算、制限読み込み範囲の指定終了
running_column=runningrange_start+1

print('〜〜〜〜〜〜〜〜〜〜〜')
print('start to read car-spec')
car_typelist=[]#車両性能取り込み car_typelist▷車両種類リスト
i=1#i▷繰り返し変数
car_dv=float(cars1.cell(2,1).value)#car_dv▷加減速記述間隔
car_fainalspeed=int(cars1.cell(2,2).value)#car_fainalspeed▷演算最高速度
while True:
  i=i+1
  temp=cars1.cell(1,i).value
  if temp=='E':#Eの場合、終了指示
    print(car_typelist)
    break
  else:
    car_typelist.append(temp)#ここまで、リストに車両種別追加
#車両加減速設定
car_weight=cars1.get('B3:'+str(cell_exchange(2+len(car_typelist),3)))
car_length=cars1.get('B4:'+str(cell_exchange(2+len(car_typelist),4)))
print(str(cell_exchange(len(car_typelist)+1,3+int(car_fainalspeed/car_dv))))#str以降読み込み終了地点
car_acclist=cars1.get('B5:'+str(cell_exchange(len(car_typelist)+1,3+int(car_fainalspeed/car_dv))))#car_acclist▷車両加速リスト
car_declist=cars2.get('B5:'+str(cell_exchange(len(car_typelist)+1,3+int(car_fainalspeed/car_dv))))#car_declist▷車両減速度リスト
print('weight',car_weight)#編成重量
print('length',car_length)#編成長
#print('acc',car_acclist)#加速
#print('dec',car_declist)#減速

print('〜〜〜〜〜〜〜〜〜〜〜')
print('the main roop starting')
for column in range(runningrange_number):#メインループ
  line=4
  pattern=[]
  car_type=car_typelist.index(str(target.cell(2,running_column).value))#car_type車両形式
  print(car_type)
  weight=int(car_weight[0][car_type])*1000#weight▷編成重量kg
  length=int(car_length[0][car_type])#length▷編成長m
  sum_time=0
  while True: #停車通過パターン読み込み
    pattern_data=target.cell(line,running_column-1) #D出発,A到着,P通過,S停車,F次列 pattern_data読み込み値
    pattern.append(str(pattern_data.value)) #pattern停車情報リスト
    if pattern[-1]=='F':#Fで改行
      print('F')
      break
    if pattern[-1]=='S' or pattern[-1]=='A': #A,Sのときすなわち停車時に計算かつクリア
      print(pattern)
      limit_range=read_limit(cell_exchange(limitrange_start,line-len(pattern)+1),cell_exchange(limitrange_end,line))#limit_range▷リスト
      print(limit_range)

      exchange=[]#制限リストを1列に置き換える
      for i in range(len(limit_range)-1):
        limit_range[i+1].pop(0)
      for i in range(len(limit_range)):
        for j in range(len(limit_range[i])):
          exchange.append(limit_range[i][j])

      exchange.pop(0)#直前制限削除
      speedlimit=[[],[]]#speedlimit▷[場所,速度]

      for i in range(int(len(exchange)/2)):#制限リストを場所、速度に分ける
        adition=0
        if speedlimit[1]!=[]:
          if exchange[1]>speedlimit[1][-1] and speedlimit[1][-1]!=0:
            adition=length
        temp=adition+exchange.pop(0)#場所書き込み
        speedlimit[0].append(temp)
        temp=exchange.pop(0)#速度書き込み
        speedlimit[1].append(temp)
      print(speedlimit)

      temp=0
      for i in range(len(speedlimit[1])):#不要部を削除
        if speedlimit[1][temp]==0:
          speedlimit[1].pop(temp)
          speedlimit[0].pop(temp)
        else:
          temp=temp+1

      temp=0
      temp1=1
      for i in range(len(speedlimit[0])-1):#通過駅を含むときの距離程加算
        if speedlimit[0][temp1]==0:
          speedlimit[0][temp1]=temp
          speedlimit[0].pop(temp1)
          speedlimit[1].pop(temp1)
        else:
          speedlimit[0][temp1]=speedlimit[0][temp1]+temp
          temp=speedlimit[0][temp1]
          temp1=temp1+1
      speedlimit[1][-1]=0#停車操作
      speedlimit[1].insert(0,0)
      print(speedlimit)
      #print(speedlimit[1])
      #print(speedlimit[0])

      #速度計算ループ
      for i in range(len(speedlimit[0])-1):
        #フラグリセット
        flag_non_acc=False#flag_non_acc▷加速なし
        flag_non_break=False#flag_non_break▷減速なし
        flag_non_lenge=False#flag_non_lenge▷残距離なし

        #制限入力(仮)
        if 'acc_curve' not in globals():#acc_curveが存在しない際に仮に無限と置く
          acc_curve=[[float('inf')]]
        start_speed=min(speedlimit[1][i],acc_curve[0][-1])#start_speed▷開始速度
        start_point=speedlimit[0][i]#start_point▷開始地点
        range_speed=speedlimit[1][i+1]#range_speed▷区間速度
        end_speed=speedlimit[1][i+2]#end_speed▷終了速度
        end_point=speedlimit[0][i+1]#end_point▷終了地点
        cal_range=end_point-start_point#cal_range▷計算区間

        #区間制限に向けた開始、終了制限修正
        if start_speed>=range_speed:#加速なし設定
          start_speed=range_speed#開始速度上書き
          flag_non_acc=True#加速なしフラグ成立
        if end_speed>=range_speed:#減速なし設定
          end_speed=range_speed#終了速度上書き
          flag_non_break=True#減速なしフラグ成立
        print('st',start_speed)
        print('rg',range_speed)
        print('ed',end_speed)
        print('range',cal_range)

        #減速計算⇨加速計算⇨惰行計算
        print('〜〜〜〜〜〜〜〜〜〜〜')
        print('time calculation start')
        dt=0.01 #dt▷演算間隔(s) 速度演算時隔推奨0.1~0.01

        #減速計算ループ初期設定
        speed=end_speed #speed▷現在速度　ここでは初速度を代入(終末速度)
        location=end_point #location▷現在位置　ここでは終了地点を代入
        break_curve=[[speed],[location]] #減速曲線描画リスト[[速度],[地点]]
        runtime=0 #runtime▷減速時間　初期値0s

        #減速計算ループ
        while flag_non_break==False:
          roundspeed=round(speed/car_dv)*car_dv#roundspeed▷丸め後の速度
          dec=float(car_declist[int(roundspeed/car_dv)][car_type])#dec▷加速度
          speed=speed+(dec*dt)#速度加算
          location=location-speed*dt/3.6#現在位置減算(/3.6はm/sからkm/hへの対応)
          #print(speed)
          #print(location)
          runtime=runtime+dt#減速時間加算
          break_curve[0].append(round(speed,1))#ブレーキ曲線速度記録 照査処理のために丸める
          break_curve[1].append(location)#ブレーキ曲線位置記録
          if speed>=range_speed:#現在速度が区間速度を超えた場合
            sum_time=sum_time+runtime#sum_time▷合計時間　加算
            break#減速処理終了

        #加速計算初期設定
        speed=start_speed#初速度(仮)
        location=start_point#location▷現在位置
        acc_curve=[[speed],[location]]#acc_curve▷加速曲線[[速度][位置]]
        runtime=0#runtime▷加速時間　初期値0s

        #加速計算ループ
        while flag_non_acc==False: #加速計算
          roundspeed=round(speed/car_dv)*car_dv#roundspeed▷丸め後の速度
          force=float(car_acclist[int(roundspeed/car_dv)][car_type])#force▷加速力　参照
          acc=force/weight#acc▷加速度km/h/s
          speed=speed+(acc*dt)#速度加算
          location=location+speed*dt/3.6#位置加算
          runtime=runtime+dt#走行時間加算
          #print(speed)
          #print(location)
          acc_curve[0].append(round(speed,1))#加速曲線記録
          acc_curve[1].append(location)

          #減速パターン超過照査
          check=round(speed,1)#超過照査に向けた速度丸め定義
          if check>=min(break_curve[0]):#パターン照査
            while True:
              try:
                restlenge=break_curve[1][break_curve[0].index(check)]-location#速度に該当する位置を確認
                #⇑残距離=ブレーキ曲線[位置][ブレーキ曲線[速度]にcheckの値の位置]-現在距離
                if restlenge/(speed/3.6)<=10:#もし残距離が10秒で走り切る距離以下なら
                  flag_non_lenge=True#加速途中終了フラグ
                break
              except ValueError:#ない場合そのまま-0.1再度確認
                check=round(check-0.1,1)
            if flag_non_lenge==True:#加速終了処理(加速途中終了フラグ成立)
              flag_non_lenge=False#フラグリセット
              print('残距離不足',restlenge)
              break
            if speed>=range_speed:#最高速到達停止
              print('最高速到達')
          if location>=end_point:#距離超過停止
            print('超過')
            break
        sum_time=sum_time+runtime#合計時間加算
        print('加速距離',location)
        print('time calculation finish')
        print('〜〜〜〜〜〜〜〜〜〜〜')
        pyp.plot(break_curve[1],break_curve[0],'#00FF00')#減速曲線描画
        pyp.plot(acc_curve[1],acc_curve[0],'#ff0000')#加速曲線描画
      target.update_cell(line,running_column,sum_time)#書き込み
      print('描画しました')
      pyp.show()

      pattern=[]#停車パターンリセット
    line=line+1#改行
  running_column=running_column+2#改列 データ列を飛ばすため2列飛ばし
#