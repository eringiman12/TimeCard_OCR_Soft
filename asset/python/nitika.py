import os
import cv2
import numpy as np
from PIL import Image

############################################################################ 
# 画像の指定領域を二値化した画像を生成する
# 引数
# imgpath:画像ファイルのパス
# expdirpath:出力先フォルダのパス 
# x,y:二値化の開始点のx,y座標
# x2,y2:二値化の終了点のx,y座標
############################################################################
def Mix_ThresholdImg(imgpath,expdirpath,x,y,x2,y2):
    #print('今回の処理対象ファイルのパス =',imgpath)
    # Pillowで画像ファイルを開く
    pil_img = Image.open(imgpath)
    # PillowからNumPyへ変換
    img = np.array(pil_img)
    # RGBからBGRへ変換する
    if img.ndim == 3:
        im = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    else:
        im = img
    # print('im =',im)
    #print('im.shape =',im.shape)
    
    # 出力先フォルダのパス
    exppath = expdirpath

    filenm = os.path.splitext(os.path.basename(imgpath))[0]
    
    # グレースケール化
    im_gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
    
    # 二値化(閾値を超えた画素を255にする。)
    threshold = 225 # 閾値の設定
    ret, img_thresh = cv2.threshold(im_gray, threshold, 255, cv2.THRESH_BINARY)

    pt1 = (x, y)   # 打刻データ部の始点(左上)x,y
    pt2 = (x2, y2) # 打刻データ部の終点(右下)x,y

    cropped = img_thresh[pt2[0]:pt2[1],pt1[0]:pt1[1]]
    fname = exppath + '\\croppedimg.jpg'
    #print('fname =',fname)
    # 二値化画像の保存
    pil_image = Image.fromarray(cropped)
    
    # Pillowで画像ファイルへ保存
    pil_image.save(fname)
    
    pil_cropped = Image.open(fname)
    cropimg = np.array(pil_cropped)
    crpim = cv2.cvtColor(cropimg, cv2.COLOR_RGB2BGR)
    print('cropimg.shape =',cropimg.shape)

    # 画像の合成
    fore_img = crpim
    back_img = im
    h, w = back_img.shape[:2]

    #合成用座標
    dx = pt1[0]      
    dy = pt2[0] 
    
    M = np.array([[1, 0, dx], [0, 1, dy]], dtype=float)
    img_warped = cv2.warpAffine(fore_img, M, (w, h), back_img, borderMode=cv2.BORDER_TRANSPARENT)
    # expfilenm = expdirpath + '\\' + filenm + '_mix.jpg' # 同一フォルダ内に作成する場合
    expfilenm = expdirpath + '\\' + filenm + '.jpg' # 別フォルダに作成する場合
    #print('出力先フォルダのパス =',expfilenm)
    
    #cv2.imwrite(expfilenm,img_warped)
    warped = cv2.cvtColor(img_warped, cv2.COLOR_BGR2RGB)
    pil_image = Image.fromarray(warped)
    
    # Pillowで画像ファイルへ保存
    pil_image.save(expfilenm)

    # 一時データ削除
    os.remove(fname)

# フォルダ内一括処理
basepath = r'C:\xxxx\会社名\1'
s = 'Thumbs.db'
for curDir, dirs, files in os.walk(basepath):
   #  print('===================')
   #  print('現在のディレクトリ:'  , curDir)
   #  print('内包するディレクトリ:' , dirs)
   #  print('内包するファイル:' , files)
    for file in files:
        newDir = curDir.replace(os.path.basename(basepath), os.path.basename(basepath)+'_mixedimgs')
      #   if os.path.exists(newDir):
      #       print('指定フォルダは存在します。')
      #   else:
      #       print('指定パスが存在しませんでした。フォルダを作成します。')
      #       os.makedirs(newDir)

        fileflp = os.path.join(curDir, file)
      #   print('元画像パス =',fileflp)
      #   print('出力先パス =',newDir)
        if fileflp.find(s) != -1:
            break
        # 引数（処理フォルダ、出力フォルダ、left、right,top,bottom）   
        Mix_ThresholdImg(fileflp,newDir ,110,370,150,1045)