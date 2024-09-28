import os
import shutil
import sys
import glob

import datetime
import re

import DB_Process
import Tool
import numpy as np

# 日付変数
currentDateTime = datetime.datetime.now()
# date化
date = currentDateTime.date()
# 年変数
year = date.strftime("%Y")
# 年変数
month = date.strftime("%m")
# 日付変数
day = date.strftime("%d")
# 給与支給日
kyuyo_day = 25

dt = datetime.datetime.today()

# OCR結果フォルダパス
Shogo_Folda = 'OCRの対象管理フォルダまでのパス'
# # 現場リストパス
#  人事労務の顧問先管理に変更
komon_List = "顧問管理.xlsm"

# 所属リストの名前管理を定義する
shozoku = "所属リスト"

# === Excel位置関係変数
# サムネイルの画像位置
Img_pos = "O7"

# == セルの結合範囲
# 主現場セルの結合範囲,副現場セルの結合範囲,処理月の結合
# Merge_ary = ["A2:B2","C2:D2","A1:A2"]

# OCRの処理を行う
class OCR_process():
    def __init__(self,Shogo_Path,shzoku_name):
      
      #  ツールオブジェクトのインスタンス化
      self.tool = Tool.Tool_object()
      #  ﾀｲﾑｶｰﾄﾞ集計templateパス
      TimeCard_totalling_tmp_Path = r"C:\XXXX共有\最新_タイムカード集計_" + shzoku_name + ".xlsx"
      #  ﾀｲﾑｶｰﾄﾞのファイルパス
      self.TimeCard_totalling_Path ="./asset/etc_files/Excel/" + os.path.basename(TimeCard_totalling_tmp_Path).replace(shzoku_name,shzoku_name + "_" + self.tool.year_converter_to_wareki(year) + "." + str(month))
      
      if os.path.isfile(TimeCard_totalling_tmp_Path) == True:
         if os.path.isfile(self.TimeCard_totalling_Path) == False:
            shutil.copy(TimeCard_totalling_tmp_Path,self.TimeCard_totalling_Path)
          # 所属名
         self.shzoku_name = shzoku_name
         # OCR結果作成フォルダパス
         # self.Ocr_path = self.tool.Folda_Create(Shogo_Folda + "/" + "タイムカード/" + year + "_" + month + "/" + day + "/" + shzoku_name)
         # OCR対象パス
         self.thumnail_Folda_ary = Shogo_Path
         
         # Excel操作オブジェクトのインスタンス化   
         self.excel_Obj =Tool.Excel_Operations(self.TimeCard_totalling_Path)
         #   シートの作成
         self.sheet = self.excel_Obj.Sheet_Create()
         # 社員リストの作成（商号名が返り値になる）
         self.Name_Only_List,self.Bumon_and_name_List,self.yyyy_mm_dd = self.excel_Obj.Write_employee_list()
         # 日付リスト
         self.DayList = self.excel_Obj.Days_List_Create(self.yyyy_mm_dd)
         # サムネイルフォルダと智目カードフォルダのソート
         self.Time_Foda,self.thum_File = self.tool.Process_Folda_Create(Shogo_Path)
         
         # タイムカード種類(Mx,a,b,c,TA)
         self.Time_Card_type = self.tool.Txt_Read(r"C:\XXXX\Time_Card_Status.txt",0)

      else:
         print("---ﾀｲﾑｶｰﾄﾞ集計テンプレートが見つかりません。")
         sys.exit(1)


    # OCR処理関数
    def Main(self):
         # excelに画像添付させる  
        self.excel_Obj.Excel_Img_Past(self.sheet,self.thum_File,Img_pos) 
        # 社員現場詳細リストの作成
        self.excel_Obj.Shain_Shosai_Create(
            self.sheet.title, self.shzoku_name,self.Name_Only_List,self.Bumon_and_name_List)
      
        # OCR結果取得関数（OCR結果を二次元配列で受け取る）
        Ocr_Result_ary = self.tool.Ocr_Process(self.Time_Foda,self.Time_Card_type)
        # 削除インデックス配列の作成   
        del_ary = self.tool.Ary_Concat(Ocr_Result_ary,self.DayList)
        # 結果配列に格納  
        Ocr_Result_ary = self.tool.ary_del_loop(Ocr_Result_ary,del_ary)
        
        # ここからシートレイアウトの編集   
        # シートのレイアウトの作成
        self.excel_Obj.Sheet_Ocr_framework_create(self.sheet,month,Ocr_Result_ary,self.yyyy_mm_dd)
        # 作成したファイルの保存
        self.excel_Obj.Excel_Save()

# テストコード  
if __name__ == "__main__":
    Ocr_process = OCR_process(sys.argv[1],sys.argv[2])
    Ocr_process.Main()