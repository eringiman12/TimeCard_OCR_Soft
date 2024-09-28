from pickle import TRUE
import sys 
import cv2
import shutil
import os
import glob
import numpy as np
from PIL import Image
import Tool

Process_Folda = os.getcwd() + "\\asset\\img\\tmp\\img_procesing\\tmp.png"
Moto_img = os.getcwd() + "\\asset\\img\\tmp\\img_procesing\\color.png"
Tmp_Thumnail = os.getcwd() + "\\asset\\img\\tmp\\Tmp_thumnail\\"

class Matching():
    def __init__(self, img_path,shogo_kojin_num_folda):    
      # 処理対象画像
      self.img_path = img_path
      # 商号名+個人フォルダ
      self.shogo_kojin_num_folda = shogo_kojin_num_folda
      # 結合フォルダ
      self.concat_img_folda = "./asset/img/tmp/Concat_Imge/"
      # サムネイルフォルダ
      self.Tmp_Thumnail = Tmp_Thumnail + self.shogo_kojin_num_folda
      # 一時フォルダパス
      self.tmp_img_path = "./asset/img/tmp/tmp_img/"
      # 結合ファイル格納先フォルダパス
      self.Concat_img_1_2_Fols = self.concat_img_folda +self.shogo_kojin_num_folda
      
      # マッチング用画像ファイル配列
      self.template_img_path = glob.glob(
             os.getcwd() + "\\asset\\img\\tmp\\Matching_File\\*")
      # ステータスファイルのファイルパス
      self.Process_File = "./asset/etc/bin/Time_Card_Status.txt"
      # ツールオブジェクトのインスタンス化
      self.to = Tool.Tool_object()
   
      # プロセスファイル内のクリア
      with open(self.Process_File, 'r+') as f:
         f.truncate(0)
      
    # ファイルの準備  
    def File_Move(self):
      # 一時フォルダ内ゴミ掃除
      self.File_del(os.getcwd() + "\\asset\\img\\tmp\\img_procesing")
      # 一時フォルダ内ゴミ掃除
      self.File_del(os.getcwd() + "\\asset\\img\\tmp\\tmp_img")
      # 処理対象画像を一時フォルダにコピー
      shutil.copyfile(self.img_path,Process_Folda)
      # 変更予定画像
      shutil.copyfile(self.img_path,Moto_img)
      # 一時フォルダに移動した処理対象ファイルをグレースケール化
      self.Img_Color_ch()
   
   # 画像のグレースケール化
    def Img_Color_ch(self):
        # 画像指定
        img = cv2.imread(Process_Folda)
        im_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        cv2.imwrite(Process_Folda, im_gray)
    
    # 引数にフォルダパスを指定すると直下のファイルはすべて削除
    def File_del(self,f_path):
       filelist = glob.glob(f_path +"/*")
       for f in filelist:
         os.remove(f)
    
    def Matching_Process(self,Match_File,Shoritaisho_img):
      Matching_img = cv2.imread(Match_File)
      # 処理対象画像に対して、テンプレート画像との類似度を算出する
      res = cv2.matchTemplate(Shoritaisho_img, Matching_img, cv2.TM_CCOEFF_NORMED)
      
      # 類似度の高い部分を検出する
      threshold = 0.8
      loc = np.where(res >= threshold)

      # テンプレートマッチング画像の高さ、幅を取得する
      h, w,s = Matching_img.shape
      
      # マッチング成功有無判定変数
      Matching_Judge = False
      
      # 検出した部分に赤枠をつける
      for pt in zip(*loc[::-1]):
         if pt != "":
            cv2.rectangle(Matching_img, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)
            Matching_Judge = True
            break
      
      return Matching_Judge
    
   #  ファイルとフォルダの作成
    def Fol_File_Create(self,Matching_Img,Rotete_img):

      # 同じファイルが存在しているか
      Same_FileName_Judge = True
      # 画像パス
      Imge_Path = ""
   
      # 表裏判定 
      if Matching_Img != "":
         # 画像結合フォルダパス
         concat_folda_path =self.concat_img_folda
      
         # tmpフォルダパス
         tmp_folda_psth = self.tmp_img_path
         
         if Rotete_img != 0:
            Color_Img = cv2.imread(Moto_img)
            
            # 傾き分岐
            if Rotete_img == 90:
               Rotete_img = cv2.rotate(Color_Img, cv2.ROTATE_90_CLOCKWISE)
            elif Rotete_img == 180:
               Rotete_img = cv2.rotate(Color_Img, cv2.ROTATE_180)
            elif Rotete_img == 270:
               Rotete_img = cv2.rotate(Color_Img, cv2.ROTATE_90_COUNTERCLOCKWISE)
            
            # 一時フォルダ書き込み
            cv2.imwrite(tmp_folda_psth + "1.jpg", Rotete_img, (cv2.IMWRITE_JPEG_QUALITY, 100))
         
         # ファイル作成 
         if os.path.exists(tmp_folda_psth + "1.jpg") == True:
            # 変更予定画像
            shutil.copyfile(tmp_folda_psth + "1.jpg",concat_folda_path + "tmp.jpg")
         else:
            shutil.copyfile(self.img_path,concat_folda_path + "tmp.jpg")
         
         # 結合フォルダ作成
         self.to.Folda_Create(self.Concat_img_1_2_Fols)
         self.to.Folda_Create(self.Tmp_Thumnail)
         


         
         #　ﾀｲﾑｶｰﾄﾞの表裏確認
         if "Omote" in Matching_Img:
            Imge_Path = self.Concat_img_1_2_Fols + "/1.jpg"
            os.rename(concat_folda_path + "tmp.jpg",Imge_Path) 
         elif "Ura" in Matching_Img:
            Imge_Path = self.Concat_img_1_2_Fols + "/2.jpg"
            os.rename(concat_folda_path + "tmp.jpg",Imge_Path) 
      else:
         Same_FileName_Judge = False 
      
      
      return Imge_Path,Same_FileName_Judge
      
    
    # 画像の結合 
    def Concat_Imges(self,Path1,Path2,save_path):
       img_1 = Image.open(Path1)
       img_2 = Image.open(Path2)
       concat_path = save_path + "/concat.jpg"
       
       dst = Image.new('RGB', (img_1.width +
                  img_2.width, img_1.height))
       dst.paste(img_1, (0, 0))
       dst.paste(img_2, (img_1.width, 0))

       dst.save(concat_path)
       return concat_path
    
    def Folda_senbetu(self,path):
       Concat_img_fol = glob.glob(path + "/*")
       concat_path = ""
       # 結合ファイルが複数の場合
       if len(Concat_img_fol) > 1:
         # 画像の結合を行う
         concat_path = self.Concat_Imges(Concat_img_fol[0],Concat_img_fol[1],path)
         
         # 結合する前の画像は削除
         for Files in Concat_img_fol:
            os.remove(Files)
       return concat_path
               
    # メイン処理関数  
    def Main(self):
      self.File_Move()
      Matching_Img = ""
      Rotete_img = 0
      
      # 一時フォルダパス
      tmp_img_fol = os.getcwd() + "\\asset\\img\\tmp\\template_matching_tmp\\"
      
      # テンプレートマッチング画像フォルダ文ループ
      for Match_File in self.template_img_path:
         # 傾き率
         rotate = 0
         # 処理用パス
         shori_path = ""
         # プロセスカウント
         process_cnt = 1
         # マッチした画像変数
         Matching_Img = ""
         
         While_Judge = False
         
         
         
         # 360度回転終わるまで
         while rotate < 360:
            
            if shori_path == "":
               shori_path =  Process_Folda

            # 処理対象画像の読み込み
            Shoritaisho_img = cv2.imread(shori_path)
            
            # テンプレートファイルにマッチしたかどうか
            judge = self.Matching_Process(Match_File,Shoritaisho_img)
    
            if judge:
                # マッチした画像のファイル名
                Matching_Img = os.path.basename(Match_File)
                
                # 回転数を格納  
                Rotete_img = rotate
                # whilejudge  
                While_Judge = True
                break
            else:
               # 処理テンプレートパス
               shori_path = tmp_img_fol + "tmp_" + str(process_cnt)+".png"
               
               img_rotate_90_clockwise = cv2.rotate(Shoritaisho_img, cv2.ROTATE_90_CLOCKWISE)
               cv2.imwrite(shori_path, img_rotate_90_clockwise, (cv2.IMWRITE_JPEG_QUALITY, 100))

            process_cnt += 1
            # 回転を９０足す
            rotate += 90
         
         # 一時フォルダ内掃除
         self.File_del(r"C:\XXXX\htdocs\System\Ocr_Web_Sys\asset\img\tmp\template_matching_tmp")   
         
         if While_Judge:
            break
      
      # プロセスファイルを開く
      f = open(self.Process_File, 'w')
      
      # ファイルとフォルダの作成
      Macthing_Af_Img_PATH,Same_FileName_Judge = self.Fol_File_Create(Matching_Img,Rotete_img)
      
      thum_path = ""
      if Same_FileName_Judge:

            
         judge = False
         
         # ﾀｲﾑｶｰﾄﾞの種類を記載する
         if "Mx" in Matching_Img:
            f.write('Mx\n')
            f.write('4\n')
            f.write('16\n')
            f.write('4\n')
            f.write('16\n')
            
            x1 = 0
            y1 = 395
            x2 = 395
            y2 = 965
            
            judge = True
         
         elif "TA" in Matching_Img:
            f.write('TA\n')
            f.write('4\n')
            f.write('18\n')
            f.write('4\n')
            f.write('13\n')
            
            if "Omote" in Matching_Img:
               x1 = 110
               y1 = 370
               x2 = 150
               y2 = 1045
            elif "Ura" in Matching_Img:
               x1 = 105
               y1 = 360
               x2 = 150
               y2 = 795
            
            judge = True

         # テンプレートマッチング済みの画像を二値化
         if Macthing_Af_Img_PATH != "":    
            if judge:
               # ツールオブジェクトクラスの二値化関数に渡す
               self.to.Time_Parts_2_Ch_Img(Macthing_Af_Img_PATH,x1,y1,x2,y2)
               
         f.close()
         
         # 結合フォルダフォルダ内の処理
         concat_path = self.Folda_senbetu(self.Concat_img_1_2_Fols)
         tmp_img_Fl_cn = glob.glob(os.getcwd() + "\\asset\\img\\tmp\\tmp_img\\*")

         if len(tmp_img_Fl_cn) == 0:
            # 元の画像の移動
            shutil.move(Moto_img,self.Tmp_Thumnail + "\\" + os.path.basename(Macthing_Af_Img_PATH))
         else:
            shutil.move(os.getcwd() + "\\asset\\img\\tmp\\tmp_img\\1.jpg",self.Tmp_Thumnail + "\\" + os.path.basename(Macthing_Af_Img_PATH))  
      
         
         # サムネイルフォルダ内の処理
         thum_path = self.Folda_senbetu(self.Tmp_Thumnail)
            
         # 一時フォルダ内掃除
         # self.File_del(os.getcwd() + "\\asset\\img\\tmp\\img_procesing")

         return concat_path,thum_path,Same_FileName_Judge
      
      else:
            
         concat_path = ""
         return concat_path,thum_path,Same_FileName_Judge
            
            

# テストコード（テスト時にコメントアウト外す）
# ma = Matching(sys.argv[1],sys.argv[2])
# ma.Main()
