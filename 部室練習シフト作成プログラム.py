import streamlit as st
import datetime
from datetime import date
import openpyxl
from openpyxl import load_workbook
from io import BytesIO
import jpholiday
from openpyxl.styles import Border, Side, Font
from openpyxl.styles.alignment import Alignment
from mip import Model, xsum, minimize, BINARY, OptimizationStatus

import json

border_topthick = Border(top=Side(style='thick', color='000000'),
                left=Side(style='thick', color='000000'),
                right=Side(style='thick', color='000000')
)

border_bottomthick = Border(bottom=Side(style='thick', color='000000'),
                left=Side(style='thick', color='000000'),
                right=Side(style='thick', color='000000')
)

border_sidethick = Border(left=Side(style='thick', color='000000'),
                right=Side(style='thick', color='000000')
)

border_topleft = Border(top=Side(style='thick', color='000000'),
                left=Side(style='thick', color='000000'),
                right=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
)

border_topcenter = Border(top=Side(style='thick', color='000000'),
                right=Side(style='thin', color='000000')
)

border_topright = Border(top=Side(style='thick', color='000000'),
                right=Side(style='thick', color='000000'),
                bottom=Side(style='thin', color='000000')
)

border_bottomleft = Border(bottom=Side(style='thick', color='000000'),
                left=Side(style='thick', color='000000'),
                right=Side(style='thin', color='000000')
)

border_bottomright = Border(bottom=Side(style='thick', color='000000'),
                right=Side(style='thick', color='000000')
)

border_bottomcenter = Border(bottom=Side(style='thick', color='000000'),
                right=Side(style='thin', color='000000')
)

border_left = Border(top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000'),
                left=Side(style='thick', color='000000'),
                right=Side(style='thin', color='000000')
)

border_right = Border(top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000'),
                left=Side(style='thin', color='000000'),
                right=Side(style='thick', color='000000')
)

border_allthin = Border(top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000'),
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000')
)


band_list = {}
week = {}


st.session_state["page_control"] = 0
st.session_state["kinshi"] = {}

def change_page():
  # ページ切り替えボタンコールバック
  st.session_state["page_control"] += 1

def band_list_making():
  i = 1
  while st.session_state["sheet"].cell(row=5 + i, column=2).value is not None:
      band_list[i] = st.session_state["sheet"].cell(row=5 + i, column=2).value
      i += 1
  band_sum = len(band_list)
  return band_sum

def option_select():
    max_practice = st.selectbox(
        '最大練習回数',
        [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        index=0,
        placeholder="練習回数を選択してください"
    )
    return max_practice

def kinshi_select():
  # 日付を保存するためのセッション状態を初期化
  if 'dates_list' not in st.session_state:
    st.session_state['dates_list'] = []

  # 日付入力
  selected_date = st.date_input('日付を選択してください:', date.today())


  # 日付を追加するボタン
  if st.button('日付を追加'):
    if selected_date < st.session_state["start_day"] or selected_date > st.session_state["end_day"]:
      st.error('シフト期間内の日付を入力してください。')
      st.stop()
    st.session_state['dates_list'].append(selected_date)
    st.success(f'{selected_date} が追加されました。')

  if st.button('リセット'):
    st.session_state['dates_list'] = []
    st.success('リセットされました。')

  # 追加された日付と今日の日付の差分を辞書に登録（キーを1, 2, 3, ...とする）
  st.session_state["kinshi"] = {int(i + 1): (d - st.session_state["start_day"] + datetime.timedelta(days=1)).days for i, d in enumerate(st.session_state['dates_list'])}

  # 現在の追加済みの日付を表示
  st.write('追加された日付一覧:')
  i = 0
  while i < len(st.session_state['dates_list']):
    st.write(st.session_state['dates_list'][i])
    i += 1


  # 日付の差分を含む辞書を表示
  st.write('日付と今日との差分（日数）の辞書:', st.session_state["kinshi"])

def week_judge(start_day, vacation):
  # 平日と土日祝の判別
  current_day = start_day
  for i in range(1, day_sum + 1):
    if not vacation:
      if jpholiday.is_holiday(current_day) or current_day.weekday() in [5, 6]:
        week[i] = 1  # 祝日または土日
      else:
        week[i] = 0  # 平日
    else:
      week[i] = 1  # 長期休暇中
    current_day += datetime.timedelta(days=1)



def input_date():
  book1 = openpyxl.Workbook()
  for i in range(1, band_sum + 1):
    book1.create_sheet(index=0, title=band_list[i])
    sheet = book1[band_list[i]]
    for t in range(1, 8):
        sheet.cell(row=2 + t, column=2).value = str(t) + "限"
    calc_day = st.session_state["start_day"]
    j = 1
    while calc_day <= st.session_state["end_day"]:
        sheet.cell(row=2, column=2 + j).value = str(calc_day.month) + "/" + str(calc_day.day)
        j += 1
        calc_day += datetime.timedelta(days=1)

    #枠線作成
    #B列目
    sheet.cell(row=2, column=2).border = border_topleft
    for t in range(1,8):
      sheet.cell(row=2+t, column=2).border = border_left
    sheet.cell(row=9, column=2).border = border_bottomleft
    #2行目
    for j in range(1, day_sum):
      sheet.cell(row=2, column=2+j).border = border_topcenter
    #右端
    sheet.cell(row=2, column=day_sum+2).border = border_topright
    for t in range(1,8):
      sheet.cell(row=2+t, column=day_sum+2).border = border_right
    sheet.cell(row=9, column=day_sum+2).border = border_bottomright
    #中
    for i in range(1, day_sum):
      for t in range(1,7):
        sheet.cell(row=2+t, column=2+i).border = border_allthin
    #下
    for j in range(1, day_sum):
      sheet.cell(row=9, column=2+j).border = border_bottomcenter

    #書式設定
    font = Font(name="游ゴシック",size=11,bold=True)
    for t in range(1, 9):
      for j in range(1, day_sum+2):
        sheet.cell(row=1+t, column=1+j).font = font
        sheet.cell(row=1+t, column=1+j).alignment = Alignment(horizontal = 'left', vertical = 'center')


    # デフォルトで作成されるシートを削除
  if 'Sheet' in book1.sheetnames:
      book1.remove(book1['Sheet'])

    # バイトストリームにExcelファイルを保存
  buffer = BytesIO()
  book1.save(buffer)
  buffer.seek(0)

  # StreamlitのダウンロードボタンでExcelファイルをダウンロード
  st.download_button(
      label="ダウンロード",
      data=buffer,
      file_name='シフト希望記入表.xlsx',
      mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  )




def saitekika():
  # 最適化モデルの作成
  model = Model('PracticeShiftTime')
  # 決定変数の作成
  y = {(i, d, t): model.add_var(var_type='B') for i in st.session_state["I"] for d in st.session_state["D"] for t in st.session_state["T"]}
  st.session_state["y2"] = {(i, d, t): model.add_var(var_type='B') for i in st.session_state["I"] for d in st.session_state["D"] for t in st.session_state["T"]}
  s = {i: model.add_var(var_type='B') for i in st.session_state["I"]}

  # ハード制約の追加
  # 各バンドの制約
  for i in st.session_state["I"]:
      model += xsum(y[i, d, t] for d in st.session_state["D"] for t in st.session_state["T"]) >= 1  # 最低1回は練習
      model += xsum(y[i, d, t] for d in st.session_state["D"] for t in st.session_state["T"]) <= max_practice  # 練習回数は最大 max_practice
      for d in st.session_state["D"]:
          model += xsum(y[i, d, t] for t in st.session_state["T"]) <= 1  # 1日に1回のみ練習

  # 希望しない時間に練習を入れない
  for i in st.session_state["I"]:
      for d in st.session_state["D"]:
          for t in st.session_state["T"]:
              if st.session_state["kibou_time"][f"{i}_{d}_{t}"] == 0:
                  model += y[i, d, t] == 0

  
  # 部室利用禁止日に練習を割り当てない
  if st.session_state["kinshi"] is not None:
    for d in st.session_state["D"]:
        for k in st.session_state["kinshi"]:
            if d == st.session_state["kinshi"][k]:
              for i in st.session_state["I"]:
                for t in T:
                  model += y[i, d, t] == 0


  # 最終週に希望がある場合、必ず練習を入れる
  for i in st.session_state["I"]:
      if st.session_state["last_week"][i] == 1:
          model += xsum(y[i, d, t] for d in range(day_sum - 6, day_sum + 1) for t in st.session_state["T"]) >= 1 - s[i]

  # 同じ時間に練習するバンドは1つまで
  for d in st.session_state["D"]:
      for t in st.session_state["T"]:
          model += xsum(y[i, d, t] for i in st.session_state["I"]) <= 1

  # 連続して練習しない（1日以上あける）
  for i in st.session_state["I"]:
      for d in range(1, day_sum):
          model += xsum(y[i, d, t] for t in st.session_state["T"]) + xsum(y[i, d + 1, t] for t in st.session_state["T"]) <= 1

  # 目的関数の設定
  model.objective = minimize(-xsum(y[i, d, t] for i in st.session_state["I"] for d in st.session_state["D"] for t in st.session_state["T"]) + 10 * s[i])

  # 最適化の実行
  status = model.optimize()
  
  if status == OptimizationStatus.OPTIMAL:
    st.write('最適値 =', model.objective_value)
    st.session_state["y2"] = y
    for i in st.session_state["I"]:
      for d in st.session_state["D"]:
        for t in st.session_state["T"]:
          st.session_state["y2"][f"{i}_{d}_{t}"] = y[i, d, t].x

    result()


def result():
  print(st.session_state["y2"])
  book2 = openpyxl.Workbook()
  book2.create_sheet(index=0, title="結果出力")
  sheet = book2["結果出力"]
  for i in band_list:
    for d in range(1, day_sum + 1):
        for t in range(1, 8):
            if st.session_state["y2"][f"{i}_{d}_{t}"] > 0.01:
                sheet.cell(row=2 + t, column=2 + d).value = band_list[i]
  buffer2 = BytesIO()
  book2.save(buffer2)
  buffer2.seek(0)
  st.download_button(
    label="結果をダウンロード",
    data=buffer2,
    file_name="最適化結果.xlsx",
  mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')




# Webアプリのタイトル
st.title('シフトスケジュール最適化')


#ページ１：参加バンド登録
uploaded_file_path = 'バンドリスト_テンプレート.xlsx'
# ファイルをバイトとして読み込む
with open(uploaded_file_path, 'rb') as file:
    band_listfile = file.read()

if st.session_state["page_control"] == 0:
  st.header('１．参加バンドの登録')
  st.caption('ダウンロードボタンからテンプレートをダウンロードして、出演バンドを記入してください。')
  st.caption('記入を終えたファイルをアップロードしてください。')

  st.download_button(
      label="テンプレートをダウンロード",
      data=band_listfile,
      file_name='参加バンド登録＿テンプレート.xlsx',
      mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  )

  st.session_state["uploaded_file1"] = st.file_uploader("バンド名簿をアップロード", type=["xlsx"],key = "バンド名簿")
  if st.session_state["uploaded_file1"] is not None:
    change_page()



#ページ２：ライブ情報の入力
if "page_control" in st.session_state and st.session_state["page_control"] == 1:
    # st.session_state['uploaded'] = True
    st.header('２．ライブ情報の入力')
    st.session_state["book"] = load_workbook(st.session_state["uploaded_file1"])
    st.session_state["sheet"] = st.session_state["book"]["概要"]
    band_sum = band_list_making()

    st.session_state["start_day"] = st.date_input('シフト開始日を入力してください。', datetime.date(2024, 10, 10))
    st.session_state["end_day"] = st.date_input('シフト終了日を入力してください。', datetime.date(2024, 10, 31))
    if st.session_state["start_day"] > st.session_state["end_day"]:
        st.error('開始日は終了日より後の日付を入力してください。')
        st.stop()

    day_sum = (st.session_state["end_day"] - st.session_state["start_day"] + datetime.timedelta(days=1)).days
    max_practice = option_select()
    vacation = st.toggle("長期休暇期間")
    d = st.toggle("部室利用禁止日あり")
    if d == True:
      kinshi_select()
    week_judge(st.session_state["start_day"], vacation)

    if st.button("入力完了"):
      st.session_state["3pagenext"] = True
    if st.session_state["3pagenext"]:
      change_page()

#ページ３：希望日入力
if "page_control" in st.session_state and st.session_state["page_control"] == 2:
  st.header('３．練習希望日時の入力')
  input_date()

  #定数データ作成
  st.session_state["I"] = [i for i in range(1, band_sum + 1)]
  st.session_state["D"] = [i for i in range(1, day_sum + 1)]
  st.session_state["T"] = [i for i in range(1, 8)]



  st.write("記入を終えたファイルをアップロードしてください。")

  st.session_state["kibou_file"] = st.file_uploader(label="シフト希望表をアップロード", type=["xlsx"])
  
  if st.session_state["kibou_file"] is not None:
    change_page()

if "page_control" in st.session_state and st.session_state["page_control"] == 3:
  st.header('４．最適化の実行')
  book1 = load_workbook(st.session_state["kibou_file"])
  st.session_state["kibou_time"] = {}
  st.session_state["last_week"] = {}

  #希望ファイルの読み込み
  for i in st.session_state["I"]:
      sheet_band = book1[band_list[i]]  # シートを取得
      for d in st.session_state["D"]:
          for t in st.session_state["T"]:
              value = sheet_band.cell(row=2 + t, column=2 + d).value
              if value is not 1:
                  value = 0      
              # キーを文字列に変換して保存
              key_str = f"{i}_{d}_{t}"
              st.session_state["kibou_time"][key_str] = value

  for i in st.session_state["I"]:
    st.session_state["last_week"][i] = int(any(st.session_state["kibou_time"][f"{i}_{d}_{t}"] for d in range(day_sum - 6, day_sum + 1) for t in st.session_state["T"]))


  if st.session_state["kibou_time"] is not None:
    st.write("希望の読み込みに成功しました。")
    st.write("実行ボタンを押してシフトを作成します。")
    if st.button("実行ボタン"):
      saitekika()
      
  else:
    st.write("希望の読み込みに失敗しました。もう一度ファイルを読み込ませてください。")
