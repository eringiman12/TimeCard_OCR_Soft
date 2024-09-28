import shutil
import  sys
import datetime
import os
from PIL import Image
import Tool
import DB_Process


class Cut_Pieces():
    def __init__(self, Shori_Img_Path,Shogo_Name,thumnail_path):
        # DBクラス呼び出し  
        self.Db_client = DB_Process.DB_Operation()
        # ツールクラス読み込み   
        self.Tools = Tool.Tool_object()
       
        # 現在時刻   
        dt_now = datetime.datetime.now()
        # 処理対象ファイルパス  
        self.Shori_Img_Path = Shori_Img_Path
        # ﾀｲﾑｶｰﾄﾞの種類と行列の入った配列
        TimeCard_rc_syurui= open("./asset/etc/bin/Time_Card_Status.txt", 'r')
        # ﾀｲﾑｶｰﾄﾞ情報の読み取り   
        self.TimeCard_rc_syurui_ary = TimeCard_rc_syurui.read().split('\n')
        # ﾀｲﾑｶｰﾄﾞテキスト閉じる   
        TimeCard_rc_syurui.close()
        # 切り取り画像の保存先フォルダ
        self.Concat_shogo_Folda = self.Tools.Folda_Create("./asset/img/tmp/OCR_img/" + Shogo_Name + "/" + dt_now.strftime('%Y年%m月%d日_%H時%M分%S秒'))
        # 処理対象画像を開く
        self.im = Image.open(Shori_Img_Path)
        self.thum_path = thumnail_path
        
    def Cropp_Process(self,TimeCard_Barshi_Fol,row,col,strow,x,Width,y,Height,sabun_x,sabun_y):

       for rows in range(strow,row):
          # 行　フォルダ指定  
          row_Fol = TimeCard_Barshi_Fol + "/" + str(int(rows) + 1)
          os.makedirs(row_Fol)
          
         #  x位置に代入
          x_1 = x 
         #  横幅に代入
          width_1 = Width
          
           # 列ループ  
          for cols in range(0,col):               
            im_crop = self.im.crop((x_1,y, width_1, Height))
            im_crop.save(row_Fol + "/" + str(cols + 1) + ".jpg", quality=100)
            
            x_1 = width_1
            width_1 += sabun_x
          
          y = Height
          Height += sabun_y
        
    def Main(self): 
       if len(self.TimeCard_rc_syurui_ary) > 1:
          #  ﾀｲﾑｶｰﾄﾞの切り取った画像の保管フォルダ作成
         TimeCard_Barshi_Fol = self.Tools.Folda_Create(self.Concat_shogo_Folda + "/" + "TimeCard")
         # サムネイルフォルダ  
         TimeCard_Tum_Fol = self.Tools.Folda_Create(self.Concat_shogo_Folda + "/" + "Thumnail")
         
         #  DB上の数式情報の一覧を取得
         Db_get_val = self.Db_client.Select_value(
                     "select * from hdr_cut_option where Time_Card_Syurui=\"" + self.TimeCard_rc_syurui_ary[0] + "\"")
         
         row = int(self.TimeCard_rc_syurui_ary[2])
         row2 = int(self.TimeCard_rc_syurui_ary[4])
         col = int(self.TimeCard_rc_syurui_ary[1])
         
         time_cardsyurui = self.TimeCard_rc_syurui_ary[0]
         
         Date_x = 0
         # x座標  
         if time_cardsyurui == "Mx":
             Date_x = int(Db_get_val[0][2]) + 17
             Date_x2 = int(Db_get_val[0][7]) + 17
         else:
             Date_x = int(Db_get_val[0][2])
             Date_x2 = int(Db_get_val[0][7])
             
         x = Date_x
         # 横幅  
         Width = Db_get_val[0][4]
         # y座標
         y = Db_get_val[0][3] 
         #  高さ
         Height = Db_get_val[0][5]
         # 差分変化量（y） 
         sabun_y = Db_get_val[0][6]
         # テーブル２のx1座標
         x2 = Date_x2
         # テーブル２のwidth座標
         Width2 =   Db_get_val[0][8]
         # 差分変化量（x1） 
         sabun_x = Db_get_val[0][9]
         
         # ﾀｲﾑｶｰﾄﾞ表クロップ
         self.Cropp_Process(TimeCard_Barshi_Fol,row,col,0,x,Width,y,Height,sabun_x,sabun_y)
         # ﾀｲﾑｶｰﾄﾞ裏クロップ
         self.Cropp_Process(TimeCard_Barshi_Fol,row2 + row,col,row,x2,Width2,y,Height,sabun_x,sabun_y)
         
         # # サムネイルファイルを各商号個人クロップフォルダ Thumnailフォルダに移動
         shutil.move(self.thum_path,TimeCard_Tum_Fol)
         os.remove(self.Shori_Img_Path)
         # ナンバーフォルダ削除
         os.rmdir(os.path.dirname(self.Shori_Img_Path))
         
         return self.Concat_shogo_Folda