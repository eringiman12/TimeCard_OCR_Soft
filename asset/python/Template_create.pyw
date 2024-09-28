from tkinter import messagebox
import sys
import cv2
import numpy as np
from PIL import Image

class create_template():
    def __init__(self):
        im = Image.open(sys.argv[1])
        im_crop = im.crop((40, 990, 57, 1040))
        im_crop.save(
            'C:/XXXX/htdocs/System/Ocr_Web_Sys/asset/img/tmp/Matching_File/Mx_Omote_05.png', quality=95)

if __name__ == '__main__':
    pm = create_template()
