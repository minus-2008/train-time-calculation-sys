#駅間時間計算

# 認証のためのコード
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.auth.exceptions import TransportError#ネットワークエラー
from google.auth.exceptions import RefreshError#認証エラー
import csv#csv読み取り用
import time#API管制用
import bisect#リスト管理など
from matplotlib import pyplot as pyp#グラフ描画

pyp.ioff()
# 指定対象の読み込み
#情報のcsvを読み込み
try:
    defn = open('definition.csv','r',encoding='UTF-8')#csv取り込み
    defn = list(csv.reader(defn))#リストに書き換え
    json_Key = defn[0][0]#jsonキー
    seat = defn[0][1]#シート名
except (FileNotFoundError,IndexError):#csvエラー
    print('definition.csvが存在しない、または正しく記入されていません。')
    exit()

# 認証情報の設定
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
    ]
try:
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_Key, scope)
except Exception as err:
    print(
      '認証関連にエラーが発生していると考えられます、JsonKeyを確認してください。' \
      'それでもこのエラーが発生する場合は、開発者にこのログと共にお問い合わせください',err
      )
    exit()
print('認証情報が定義されました')
# クライアントの作成
client = gspread.authorize(creds)

# スプレッドシートの取得
try:
    ss = client.open(seat)
    print('スプレッドシートに接続しました') 
except TransportError:#ネットワークエラー
    print(
        'ネットワークエラーが発生しました、接続を確認後、再度実行してください'
        )
    exit()
except RefreshError:#認証エラー(おそらく)
    print(
        '認証関連でエラーが発生しています、JsonKeyを確認してください。'
        )
    exit()

target = ss.worksheet("新書き込み先")
cars_meta = ss.worksheet("メタデータ")
cars1 = ss.worksheet("車両加速力スプレットシート")
cars2 = ss.worksheet("車両減速力スプレットシート")
line_seat=ss.worksheet("路線情報")
sta_seat=ss.worksheet("駅情報")

print('シートが定義されました')
import time#API管制用
import bisect#リスト管理など
from matplotlib import pyplot as pyp#グラフ描画
pyp.ioff()
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# プログラムの注意：四捨五入はpython標準のround関数を使用しており、仕様上、.5の場合、大きい方ではなく偶数の方が出力されます
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

def read_API_seat(seat):
  try:
    answer=seat.get_all_values()
  except gspread.exceptions.APIError:
    print('API速度上限に到達しました 数秒間動作を停止しています。。。')
    time.sleep(10)
    answer=read_API_seat(seat)#待機後再起呼び出しで再実行
  return answer

def read_API_cells(seat,range):#API規制用読み込み複数セル
  try:
    answer=seat.get(range)
  except gspread.exceptions.APIError:
    print('API速度上限に到達しました 数秒間動作を停止しています。。。')
    time.sleep(10)
    answer=read_API_cells(seat,range)#待機後再起呼び出しで再実行
  return answer

def read_API_acell(seat,x,y):#API規制用読み込み単セル
  try:
    answer=seat.cell(x,y).value
  except gspread.exceptions.APIError:
    print('API速度上限に到達しました 数秒間動作を停止しています。。。')
    time.sleep(10)
    answer=read_API_acell(seat,x,y)
  return answer

def write_API_acell(seat,x,y,write):
  try:
    seat.update_cell(x,y,write)
  except gspread.exceptions.APIError:
    print('API速度上限に到達しました 数秒間動作を停止しています。。。')
    time.sleep(10)
    read_API_acell(seat,x,y,write)

def read_limit(start,end):  #速度制限読み込み
  print('制限を読み込みました')
  answer=read_API_cells(target,start+':'+end)
  for i in range(0,len(answer)):
    for j in range(0,len(answer[i])):
      answer[i][j]=int(answer[i][j])
  return answer

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

print('def関数が定義されました')

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

sta_datasorce=read_API_seat(sta_seat)
sta_datasorce.pop(0)
event_data=read_API_seat(line_seat)
event_data.pop(0)

sta_list=[[],[]]#駅番、場所
gra_list=[[],[]]#勾配、場所
lim_list=[[],[]]#速度、場所

event_dec={'STA':sta_list,'GRA':gra_list,'LIM':lim_list}#イベントリストを駅、勾配、制限リストに書き換え
for loc,typ,data in event_data:#event_all_listの1つ目loc,2つ目typ,3つ目dataが格納される
  if typ in event_dec:#typがevent_decの'aa'の中に当てはまる場合
    _list=event_dec[typ]#書き込みリストを該当リストと定義する
    _list[1].append(int(loc))#場所を記録
    _list[0].append(data)#情報を記録
    if typ=='GRA':
      _list[0][-1]=float(_list[0][-1])#イベントが勾配の場合、情報をstrからfloatに書き換える
    elif typ=='LIM':
      _list[0][-1]=int(_list[0][-1])#イベントが制限の場合、情報をstrからintに書き換える


sta_data=[]
for index,stanum in enumerate(sta_list[0]):#駅情報をリスト、それを引き出す辞書型で構成する
  sta_data.append([stanum,int(sta_list[1][index]),{}])


stanum_dec=dict(zip(sta_list[0],sta_data))#stanum{駅名:当該のリスト}
for i,track in enumerate(sta_datasorce):#読み込む行
  _list=stanum_dec[track[0]]#駅名からリストを特定
  temp=[]
  temp.append([track[1]])#番線名を格納
  temp.append([int(x) for x in sta_datasorce[i][2:6]])#単純加算情報を格納
  temp.append([[],[]])#速度制限用の空配列を格納
  for j,(limspead,limplace) in enumerate(zip(sta_datasorce[i][6::2],sta_datasorce[i][7::2])):#2つ抜き出し
    if limspead=='':#情報が記載されていないなら
      pass#パス
    else:
      temp[-1][0].append(int(limplace))#制限地点を格納
      if limspead=='ED':
        temp[-1][1].append(float('inf'))#制限速度を格納(速度解除標識を無限と書き換え)
      else:
        temp[-1][1].append(int(limspead))#制限速度を格納
  _list[2][track[1]]=temp
#print(stanum_dec)
#↓は駅情報の格納形式(一例)
#{'MT09': ['MT09', 0, {'1': [['1'], [20, 0, 20, 0], [[-400, -300, 50], [50, 45, inf]]],
#                      '2': [['2'], [0, 0, 0, 0], [[], []]],
#                      '3': [['3'], [20, 0, 20, 0], [[-400, 50], [50, inf]]]}],
#　　   　　              　番線↗ 単純加算情報↗ 追加速度制限(場所)↖　↗追加速度制限(速度) ※EDは解除、本線の制限を適応を意味する
#
#　　↙駅番号 ↙距離程↙番線リスト(空)↖
# 'MT08': ['MT08', 1750, {}],
# 'MT07': ['MT07', 2900, {}], 
# 'MT06': ['MT06', 3520, {}], 
# 'MT05': ['MT05', 4650, {'1': [['1'], [20, 0, 20, 0], [[-400, 50], [50, inf]]], 
#                         '2': [['2'], [0, 0, 0, 0], [[], []]], 
#                         '3': [['3'], [20, 0, 20, 0], [[-400, 50], [50, inf]]]}], 
# 'MT04': ['MT04', 5330, {}], 
# 'MT03': ['MT03', 6430, {}], 
# 'MT02': ['MT02', 7120, {}], 
# 'MT01': ['MT01', 8420, {'1': [['1'], [20, 0, 20, 0], [[-400, 50], [50, inf]]], 
#                         '2': [['2'], [0, 0, 0, 0], [[], []]], 
#                         '3': [['3'], [20, 0, 20, 0], [[-400, 50], [50, inf]]]}]}

car_typelist=[]#車両性能取り込み car_typelist▷車両種類リスト
i=1#i▷繰り返し変数
car_dv_acc=float(read_API_acell(cars1,3,1))-float(read_API_acell(cars1,2,1))#car_dv_acc▷加速記述間隔 初項と次項の差から計算
car_dv_dec=float(read_API_acell(cars2,3,1))-float(read_API_acell(cars2,2,1))#car_dv_dec▷減速記述間隔 初項と次項の差から計算
car_fainalspeed=int(read_API_acell(cars_meta,6,2))#car_fainalspeed▷演算最高速度
dt= float(read_API_acell(cars_meta,7,2)) #dt▷演算間隔(s)
min_coast_time=int(read_API_acell(cars_meta,8,2)) #min_coast_time▷最短惰行時間(s)、惰行予想時間がこれ以下になった場合、自動的に惰行に移行します

while True:
  i=i+1
  temp=read_API_acell(cars_meta,1,i)
  if temp=='E':#Eの場合、終了指示
    print('車種：',car_typelist)
    break
  else:
    car_typelist.append(temp)#ここまで、リストに車両種別追加
#車両加減速設定
car_weight=read_API_cells(cars_meta,'B2:'+str(cell_exchange(2+len(car_typelist),3)))#車両重量リスト
car_length=read_API_cells(cars_meta,'B3:'+str(cell_exchange(2+len(car_typelist),4)))#車両長リスト
car_acclist=read_API_cells(cars1,'B5:'+str(cell_exchange(len(car_typelist)+1,2+int(car_fainalspeed/car_dv_acc))))#car_acclist▷車両加速リスト
car_declist=read_API_cells(cars2,'B5:'+str(cell_exchange(len(car_typelist)+1,2+int(car_fainalspeed/car_dv_dec))))#car_declist▷車両減速度リスト
print('メタデータが読み込まれました')

print('演算処理を開始します')
running_column=6
while True:#メインループ
  line=4
  pattern=[]
  car_type_name=str(read_API_acell(target,2,running_column-2))#API節約のために当該列の車種をいったん書き起こす
  if car_type_name=='':#終了検知
    break
  if car_type_name not in car_typelist:#車種がリストに存在しない場合
    print('該当車種が存在しません 訂正し、再度実行してください 当該列 '+str(running_column-2))
    exit()   
  car_type=car_typelist.index(car_type_name)#車種を列の番号に置き換える
  weight=int(car_weight[0][car_type])*1000#weight▷編成重量kg
  length=int(car_length[0][car_type])#length▷編成長m
  print('演算列の車両情報が定義されました')

  while True:#開始地点を捜索
    a=read_API_acell(target,line,running_column-2)
    if a=='ST':
      next_start=read_API_acell(target,line,2)
      track_num=read_API_acell(target,line,running_column-1)
      line=line+1
      sta_passing=[[next_start,track_num]]#通過する(発着含む)駅リスト[[駅名,番線]]
      break
    else:
      line=line+1

  while True: #停車通過パターン読み込み
    sta_start=next_start#開始駅を設定
    pattern_data=read_API_acell(target,line,running_column-2)#D出発,A到着,P通過,S停車,F次列 pattern_data読み込み値
    if pattern_data=='ED':#Fで改行
      break
    write_API_acell(target,line,running_column,'↓')#読んだ行に書き込み
    track_num=read_API_acell(target,line,running_column-1)
    if track_num!=None:
      sta_passing.append([read_API_acell(target,line,2),track_num])#通過駅一覧に加算

    if pattern_data=='S': #Sのときすなわち停車時に計算かつクリア
      sta_end=next_start=read_API_acell(target,line,2)#終了駅、次開始駅を設定
      running_range=(stanum_dec[sta_start][1],stanum_dec[sta_end][1])#走行区間(開始、終了)
      limit_start=bisect.bisect_right(lim_list[1],running_range[0])-1#制限開始地点の速度(インデックス)
      limit_end=bisect.bisect_left(lim_list[1],running_range[1])#制限終了地点の速度(インデックス)
      limit_range=[[],[]]#区間速度に書き換え
      print(sta_passing)

      for index_,limspead in enumerate(lim_list[0][limit_start:limit_end+1]):
        limit_range[1].append(int(limspead))
        limit_range[0].append(max(lim_list[1][index_+limit_start]-running_range[0],0))
      print('演算区間の制限速度が読み込まれました')

      limit_range[0][-1]=running_range[1]#停車処理、最終項を停止位置、速度0に上書き
      limit_range[1][-1]=running_range[0]
      limit_range[1].insert(0,0)#発車処理、初項に0を挿入
      print(limit_range)


      for sta in reversed(sta_passing):#駅構内制限処理
        _sta=stanum_dec[sta[0]][2]
        sta_loc=stanum_dec[sta[0]][1]-running_range[0]
        track_data=_sta[sta[1]]
        print(track_data)
        for check_loc in reversed(limit_range[0]):#距離程加算
          if check_loc<=sta_loc:
            break
          limit_range[0][limit_range[0].index(check_loc)]+=track_data[1][2]#出発距離程加算
        
        lim_adi=list(zip(track_data[2][0],track_data[2][1]))
        print(lim_adi)
        flag_backrange=True
        index_start=bisect.bisect_right(limit_range[0],sta_loc+lim_adi[0][0])#制限追加区間を確認し、indexで指定
        index_end=bisect.bisect_right(limit_range[0],sta_loc+lim_adi[-1][0])
        for i in range(index_start,index_end):#すでに変更区間内にある制限をstrに書き換え(のちに削除を行う分)
          limit_range[1][i+1]=str(limit_range[1][i+1])
        for point_talp in reversed(lim_adi): #番線別制限速度追加(後ろから)
          point_loc=sta_loc+point_talp[0]#位置
          point_lim=point_talp[1]#速度
          if point_loc>=limit_range[0][-1]:#制限地点が計算範囲外(後ろ)の場合
            pass
          elif point_loc<limit_range[0][0]:#制限地点が計算範囲外(前)の場合
            if flag_backrange:#初回の場合は、発車制限速度に上書き
              limit_range[1][1]=point_lim
              flag_backrange=False
              print('発車制限が上書きされました')
            else:#二回目以降はなし
              pass
          else:#制限追加処理
            
            for check_loc in reversed(limit_range[0]):#設定箇所にすでに制限速度ある場合→上書き
              if point_loc in limit_range[0]:
                print('aa',point_loc,point_lim,index_)
                index_=limit_range[0].index(point_loc)
                limit_range[1][index_]=point_lim
                break
              else:
                index_=bisect.bisect_left(limit_range[0],point_loc)#挿入する場所を取得
                print('bb',point_loc,point_lim,index_)
                print(limit_range)
                limit_range[0].insert(index_,point_loc)#位置を挿入
                limit_range[1].insert(index_+1,min(point_lim,int(limit_range[1][index_])))#速度を挿入(今までの制限速度から小さいほう)
                print('挿入しました',limit_range)
                break
        for i,content in enumerate(limit_range[1]):#すでにあり、上書きされていない制限速度を削除
          if type(content)==str:
            print('popします',limit_range)
            limit_range[1].pop(i)
            limit_range[0].pop(i)
            print('pop',limit_range)
          
        for check_loc in reversed(limit_range[0]):
          if check_loc<sta_loc:
            break
          limit_range[0][limit_range[0].index(check_loc)]+=track_data[1][0]#到着距離程加算
          #出発駅の到着も加算されているが、差を取るので問題はない
        print(limit_range)

      for i in range(1,len(limit_range[0])):#編成長加算
        if limit_range[1][i]<limit_range[1][i+1]:
          limit_range[0][i]+=length


      print(limit_range)
      section_list=[]#計算区間に分けた形式
      for i in range(len(limit_range[0])-1):
        start_point=limit_range[0][i]
        end_point=limit_range[0][i+1]
        distance=end_point-start_point
        max_speed=limit_range[1][i+1]
        start_speed=limit_range[1][i]#フラグ系統の処理があるため、最高速を上回っても上書きをここでは行わない
        end_speed=limit_range[1][i+2]
        section_list.append([start_point,end_point,distance,start_speed,max_speed,end_speed])
      all_distance=limit_range[0][-1]-limit_range[0][0]#総距離
      print(section_list)
      print('書き換えの処理が完了しました')

      sum_time=0#合計経過時間をリセット

      #速度計算ループ

      cum_acc_distance=0#累積加速距離(惰行加速を除く)
      acc_curve=[[],[],[]]#acc_curve▷加速曲線[[速度][位置][加速時間]]　加速時間はその加速段階における累積時間である
      for section in section_list:
        #print(section)
        Flag_acceleration=True#加速フラグ
        fwd_acc_speed=float('inf')#ひとつ前の加速の終端速度(加速不足の場合)
        
        start_point=section[0]
        end_point=section[1]
        distance=section[2]#distance▷計算距離
        sat_speed=min(section[3],fwd_acc_speed)#start_speed▷開始速度
        max_speed=section[4]#max_speed▷区間速度
        end_speed=section[5]#end_speed▷終了速度

        if sat_speed>=max_speed:#加速無し処理
          sat_speed=max_speed
          Flag_acceleration=False
          #print('加速無し')

        speed=sat_speed#初速度(仮)
        location=start_point#location▷現在位置
        runtime=0
        while Flag_acceleration:#加速本処理
          roundspeed=round(speed/car_dv_acc)*car_dv_acc#roundspeed▷丸め後の速度
          force=float(car_acclist[int(roundspeed/car_dv_acc)][car_type])#force▷加速力N　参照
          acc=(force/weight)*3.6#acc▷加速度km/h/s
          speed=speed+(acc*dt)#速度加算
          location=location+speed*dt/3.6#位置加算
          runtime=runtime+dt#走行時間加算
          acc_curve[0].append(round(speed,1))#加速曲線記録
          acc_curve[1].append(location)
          acc_curve[2].append(round(runtime,1))
          if speed>=max_speed:
            Flag_acceleration=False
          #  print('加速終了')
          elif location>=end_point:
            Flag_acceleration=False
            fwd_acc_speed=speed
          #  print('距離不足')

      pyp.plot(acc_curve[1],acc_curve[0],'#ff0000')#加速曲線描画
      

      #減速計算ループ
      
      cum_dec_distance=0#累積減速距離(惰行減速を除く)
      dec_curve=[[],[],[]]
      for section in reversed(section_list):
        print(section)
        Flag_deceleration=True
        fwd_dec_speed=float('inf')#次の減速の開始速度(減速不足の場合)

        start_point=section[0]
        end_point=section[1]
        distance=section[2]#distance▷計算距離
        sat_speed=section[3]#start_speed▷開始速度
        max_speed=section[4]#max_speed▷区間速度
        end_speed=min(section[5],fwd_dec_speed)#end_speed▷終了速度

        if end_speed>=max_speed:#加速無し処理
          end_speed=max_speed
          Flag_deceleration=False
        #  print('減速無し')
        
        speed=end_speed#初速度(仮)
        location=end_point#location▷現在位置
        runtime=0
        while Flag_deceleration:
          roundspeed=round(speed/car_dv_dec)*car_dv_dec#roundspeed▷丸め後の速度
          force=float(car_declist[int(roundspeed/car_dv_dec)][car_type])#force▷加速力N　参照
          dec=(force/weight)*3.6#acc▷加速度km/h/s
          speed=speed+(dec*dt)#速度加算
          location=location-speed*dt/3.6#位置加算
          runtime=runtime+dt#走行時間加算
          dec_curve[0].append(round(speed,1))#加速曲線記録
          dec_curve[1].append(location)
          dec_curve[2].append(round(runtime,1))
          if speed>=max_speed:
            Flag_deceleration=False
          #  print('減速終了')
          elif location<=start_point:
            Flag_deceleration=False
            fwd_dec_speed=speed
          #  print('距離不足')
      pyp.plot(dec_curve[1],dec_curve[0],'#00FF00')
      pyp.show()


      exit()
      for section in section_list:
        #フラグリセット
        flag_non_acc=False#flag_non_acc▷加速なし
        flag_non_break=False#flag_non_break▷減速なし
        flag_non_lenge=False#flag_non_lenge▷残距離なし

        #制限入力(仮)
        if 'acc_curve' not in globals():#acc_curveが存在しない際に仮に無限と置く
          acc_curve=[[float('inf')]]
        start_speed=min(section[1],acc_curve[0][-1])#start_speed▷開始速度
        range_speed=section[2]#max_speed▷区間速度
        end_speed=section[3]#end_speed▷終了速度
        distance=section[0]#distance▷計算距離

        #区間制限に向けた開始、終了制限修正
        if start_speed>=range_speed:#加速なし設定
          start_speed=range_speed#開始速度上書き
          flag_non_acc=True#加速なしフラグ成立
        if end_speed>=range_speed:#減速なし設定
          end_speed=range_speed#終了速度上書き
          flag_non_break=True#減速なしフラグ成立

        #減速計算⇨加速計算⇨惰行計算

        #減速計算ループ初期設定
        speed=end_speed #speed▷現在速度　ここでは初速度を代入(終末速度)
        location=end_point #location▷現在位置　ここでは終了地点を代入
        break_curve=[[speed],[location]] #減速曲線描画リスト[[速度],[地点]]
        runtime=0 #runtime▷減速時間　初期値0s

        #減速計算ループ
        while flag_non_break==False:
          roundspeed=round(speed/car_dv_dec)*car_dv_dec#roundspeed▷丸め後の速度
          force=float(car_declist[int(roundspeed/car_dv_dec)][car_type])#force▷減速力N　参照
          dec=(force/weight)*3.6 #dec▷加速度
          speed=speed+(dec*dt)#速度加算
          location=location-speed*dt/3.6#現在位置減算(/3.6はm/sからkm/hへの対応)
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
          roundspeed=round(speed/car_dv_acc)*car_dv_acc#roundspeed▷丸め後の速度
          force=float(car_acclist[int(roundspeed/car_dv_acc)][car_type])#force▷加速力N　参照
          acc=(force/weight)*3.6#acc▷加速度km/h/s
          speed=speed+(acc*dt)#速度加算
          location=location+speed*dt/3.6#位置加算
          runtime=runtime+dt#走行時間加算
          acc_curve[0].append(round(speed,1))#加速曲線記録
          acc_curve[1].append(location)

          #減速パターン超過照査
          check=round(speed,1)#超過照査に向けた速度丸め定義
          if check>=min(break_curve[0]):#パターン照査
            while True:
              if check not in break_curve[0]:
                check=round(check-0.1,1)
              restlenge=break_curve[1][break_curve[0].index(check)]-location
                #⇑残距離=ブレーキ曲線[位置][ブレーキ曲線[速度]にcheckの値の位置]-現在距離
              if restlenge/(speed/3.6)<=min_coast_time:#もし残距離が10秒で走り切る距離以下なら
                flag_non_lenge=True#加速途中終了フラグ
              break
#              try:
#                restlenge=break_curve[1][break_curve[0].index(check)]-location#速度に該当する位置を確認
#                #⇑残距離=ブレーキ曲線[位置][ブレーキ曲線[速度]にcheckの値の位置]-現在距離
#                if restlenge/(speed/3.6)<=10:#もし残距離が10秒で走り切る距離以下なら
#                  flag_non_lenge=True#加速途中終了フラグ
#                break
#              except ValueError:#ない場合そのまま-0.1再度確認
#                check=round(check-0.1,1)
            if flag_non_lenge==True:#加速終了処理(加速途中終了フラグ成立)
              flag_non_lenge=False#フラグリセット
              break
            if speed>=range_speed:#最高速到達停止
              break
          if location>=end_point:#距離超過停止
            restlenge=end_point-location
            break

        coast_time=restlenge/(speed/3.6)#惰行処理(仮)
        sum_time=sum_time+runtime+coast_time#合計時間加算
        #pyp.plot(break_curve[1],break_curve[0],'#00FF00')#減速曲線描画
        #pyp.plot(acc_curve[1],acc_curve[0],'#ff0000')#加速曲線描画
        start_point+=distance#開始距離加算
      sum_time=sum_time//1+1#所要時間切り上げ計算(切り捨て除算ののち1足して秒数倍)
      print('時間計算が終了しました')
      write_API_acell(target,line,running_column,sum_time)#書き込み
      #print('描画しました')
      #pyp.show()
      pattern=[]#停車パターンリセット
    line=line+1#改行
  running_column=running_column+3#改列 データ列を飛ばすため3列飛ばし
print('終了しました')