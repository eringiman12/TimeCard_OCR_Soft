

from google.cloud import vision
import io
import os
import glob
import re

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "C:\\XXXX共有\\GCP_key.json"
client = vision.ImageAnnotatorClient()


test_path = glob.glob("C:\\XXXX\\htdocs\\System\\Ocr_Web_Sys\\asset\\img\\tmp\\OCR_img\\会社名\\20xx年xx月xx日_xx時xx分xx秒 copy\\TimeCard\\2\\*")

def detect_text(path):
   # 返り値変数の取得  
   rtn_text = ""
   
   # 画像の読み込み
   with io.open(path, 'rb') as image_file:
      content = image_file.read()
   # vision APIの呼び出し
   image = vision.Image(content=content)

   # レスポンス帰り井
   response = client.text_detection(image=image)
   # テキスト配列を返す
   texts = response.text_annotations
   
   for item in texts:
      rtn_text = item.description

   return rtn_text

if __name__ == "__main__":
   new_list = sorted(test_path)
   for b in new_list:
      
      test = detect_text(b)
      print(test)
          
       