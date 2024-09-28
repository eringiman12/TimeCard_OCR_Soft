from multiprocessing.dummy import Process
import shutil
import tkinter as tk
import os
import datetime
import time
from natsort import natsorted
import glob
from tkinter import messagebox

import Template_Matching
import Cut_into_pieces_TimeCard
import Tool
import OCR_Process

import datetime

# 処理フォルダ
Process_Folda = os.getcwd() + "\\asset\\img\\tmp\\scandata"

# 開始ボタン押下
def push_button(str_r):
   print("=処理開始")
   if len(glob.glob("./asset/img/tmp/scandata/*")) != 0:
      # 時間計測開始
      time_sta = time.perf_counter()
      # ツールクラス読み込み   
      Tools = Tool.Tool_object()
      # 日付変数
      currentDateTime = datetime.datetime.now()
      # date化
      date = currentDateTime.date()
      # 年変数
      year = date.strftime("%Y")
      # 年変数
      month = date.strftime("%m")
      
      shorigo_psth = os.getcwd() + "\\asset\\img\\tmp\\処理後データ\\" + year+"年"+month+"月"
      # 処理完了フォルダ
      Tools.Folda_Create(shorigo_psth)
      
      # 個人フォルダに仕分ける。（ファイル数が奇数の場合は、エラー）
      bo,er_path = Tools.TimeCard_Folda_Create(Process_Folda)
      
      if bo:
         for Shogo in natsorted(glob.glob(Process_Folda + "/*")):
            Shogo_Name = os.path.basename(Shogo)
            
            print("=== 商号名：[" + Shogo_Name +"]の処理を開始します。")
            
            try:
               for Kojin in natsorted(glob.glob(Shogo + "/*")):
                  Nomber_Folda = os.path.basename(Kojin)
                  if "Thumbs" not in Nomber_Folda:
                     print("---- 人物カウントフォルダ：" + Nomber_Folda +"人目の処理を開始します。")             
                     # 処理する商号パス変数
                     Shori_Shogo_Path = ""
                     cnt = 1
                     
                     tmp_path = True
                     
                     for omote_ura in natsorted(glob.glob(Kojin + "/*")):
                        # 商号 / 個人ディレクトリ
                        Concat_Folda = Shogo_Name + "/" + Nomber_Folda

                        print("----- マッチング処理：" + Nomber_Folda +"フォルダ内画像"+str(cnt)+"回目の表裏マッチング処理を開始します。")   
                        
                        # パターンマッチングインスタンス化
                        ma = Template_Matching.Matching(omote_ura,Concat_Folda)
                        concat_path,thum_path,judge = ma.Main()
                        
                        if judge:
                              
                           print("----- マッチング処理：" + Nomber_Folda +"フォルダ内画像"+str(cnt)+"回目の表裏マッチング処理を終了しました。") 
                           
                           if concat_path != "" and thum_path != "":
                              print("----- 結合処理：" + Nomber_Folda +"フォルダ内の画像ファイルの結合処理を開始します。") 
                              cu = Cut_into_pieces_TimeCard.Cut_Pieces(concat_path,Shogo_Name,thum_path)
                              Shori_Shogo_Path = cu.Main()
                              print("----- 結合処理：" + Nomber_Folda +"フォルダ内の画像ファイルの結合処理を終了しました。") 
                           
                           #　表裏カウント 
                           cnt += 1  
                        
                        else:
                           # マッチングに失敗したとき
                           tmp_path = False
                           break
                     print(tmp_path)      
                     #　マッチング成功時のみ処理 
                     if tmp_path:
                        #　商号パスが空でなければ 
                        if Shori_Shogo_Path != "":
                           print("---- OCR処理：" + Nomber_Folda +"フォルダ内の画像ファイルのOCR処理を開始します。") 
                           ocr = OCR_Process.OCR_process(Shori_Shogo_Path,Shogo_Name)
                           ocr.Main()
                           print("---- OCR処理：" + Nomber_Folda +"フォルダ内の画像ファイルのOCR処理を終了しました。") 
                        
                        print("---- フォルダ移動：" + Nomber_Folda +"フォルダを処理後フォルダに移動します。") 
                        
                        if os.path.exists(shorigo_psth + "\\" + Shogo_Name + "\\" + Nomber_Folda) == False:
                              # 商号内の個人フォルダの異動
                           shutil.move(Kojin,shorigo_psth + "\\" + Shogo_Name + "\\" + Nomber_Folda)
                        else:
                           shutil.rmtree(Kojin)
                     
                     else:
                         Move_Path =os.getcwd() + "\\asset\\img\\tmp\\エラー画像\\" + Shogo_Name + "\\" + Nomber_Folda
                         print("---- OCR処理：" + Nomber_Folda +"フォルダ内の画像ファイルのマッチングに失敗しました。")
                         if os.path.exists(os.getcwd() + "\\asset\\img\\tmp\\エラー画像\\" + Shogo_Name + "\\" + Nomber_Folda) == True:
                           t_delta = datetime.timedelta(hours=9)
                           JST = datetime.timezone(t_delta, 'JST')
                           now = datetime.datetime.now(JST)
                           d = '{:%Y%m%d%H%M%S}'.format(now) 
                           Move_Path = os.getcwd() + "\\asset\\img\\tmp\\エラー画像\\" + Shogo_Name + "\\" + Nomber_Folda + "_" + d
                           
                         os.makedirs(Move_Path)
                         for Move_Files in glob.glob(Kojin + "/*"):
                           shutil.move(Move_Files,Move_Path)
                         
                         shutil.rmtree(Kojin)   
                        
                     tmp_path = True
  
               
               print("=== 商号名：[" + Shogo_Name +"]の処理を終了します。")
               
            finally:
               
               # shutil.rmtree(Shogo)  
               # 結合ファイルの後処理
               Shori_Subject_Path = [
                  os.getcwd() + "\\asset\\img\\tmp\\Concat_Imge",
                  os.getcwd() + "\\asset\\img\\tmp\\Tmp_thumnail",
                  os.getcwd() + "\\asset\\img\\tmp\\Genba_Img",
                  os.getcwd() + "\\asset\\img\\tmp\\template_matching_tmp",
                  os.getcwd() + "\\asset\\img\\tmp\\tmp_img",
                  os.getcwd() + "\\asset\\img\\tmp\\Tmp_thumnail",
                  os.getcwd() + "\\asset\\img\\tmp\\OCR_img",
                  # os.getcwd() + "\\asset\\img\\tmp\\img_procesing"
               ]
               
               for S_path in Shori_Subject_Path:
                  File_After_Process(S_path)
                  
               if len(glob.glob(os.getcwd() + "\\asset\\img\\tmp\\scandata\\" + Shogo_Name + "\\*")) == 0:
                  #  File_After_Process(os.getcwd() + "\\asset\\img\\tmp\\scandata\\"+ Shogo_Name)
                   shutil.rmtree(os.getcwd() + "\\asset\\img\\tmp\\scandata\\"+ Shogo_Name)
                  #  os.removedirs(os.getcwd() + "\\asset\\img\\tmp\\scandata\\"+ Shogo_Name)
                   if os.path.exists(os.getcwd() + "\\asset\\img\\tmp\\scandata") == False: 
                      os.mkdir(os.getcwd() + "\\asset\\img\\tmp\\scandata")
               
               if os.path.exists(os.getcwd() + "\\asset\\img\\tmp\\scandata") == False:
                  os.mkdir(os.getcwd() + "\\asset\\img\\tmp\\scandata")
               
               # 時間計測終了
               time_end = time.perf_counter()
               # 経過時間（秒）
               tim = time_end- time_sta
               print("=処理終了" + str(tim) + "秒で処理終了")

      else:
         # メッセージボックス（エラー） 
         messagebox.showerror("エラー", er_path + "の画像数が奇数です。")
      
      # 時間計測終了
      time_end = time.perf_counter()
      # 経過時間（秒）
      tim = time_end- time_sta
      # print("=処理終了" + str(tim) + "秒で処理終了")
   else:
       print("--scandataフォルダ内に画像が存在しません。")
       print("=処理終了")

def File_After_Process(path):
   shutil.rmtree(path)
   # if os.path.exists(os.getcwd() + "\\asset\\img\\tmp\\scandata") == False: 
   os.mkdir(path)
   

if __name__ == "__main__": 
   main_win = tk.Tk()
   main_win.title("実行ウィンドウ")
   main_win.geometry("300x100")
   button = tk.Button(text="開始", command=lambda:push_button("test"))
   button.pack()

   main_win.mainloop()