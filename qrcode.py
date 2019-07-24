#!/usr/bin/python
# -*- coding: UTF-8 -*-
import pyzbar.pyzbar as pyzbar
import cv2
import zxing
import os
from openpyxl import Workbook


if __name__ == "__main__":
    xlsfile = 'output.xlsx'
    imgs=[]
    for root, dirs, files in os.walk(os.path.abspath('.')):
        if len(files)>0:
            imgs=files
            break
    if len(imgs)== 0:
        print("目录下没有文件")
    else:
        wb = Workbook()    #创建文件对象
        ws = wb.active
        ws.append(["图片名", "序号（左上角）", "成绩（右上角）"])
        dd=root.replace("\\", "/")
        dd="file:/"+dd+"/tmp/"
        if not os.path.exists(root+"\\tmp\\"):
            os.makedirs(root+"\\tmp\\")
        for img in imgs:
            if img.endswith('.jpg') or img.endswith('.png'):
                 print("开始解析文件："+img+"\n")
                 frame= cv2.imread(img)
                 cropped = frame[0:290, 100:512]  # 裁剪坐标为[y0:y1, x0:x1]
                 cv2.imwrite(root+"\\tmp\\"+img, cropped)

                 #x, y = frame.shape[0:2]
                 #frame = cv2.resize(frame, (int(y * 2), int(x * 2)))
                 reader1 = zxing.BarCodeReader()
                 barcode1 = reader1.decode(dd+img)
                 print(barcode1)
                 print(img+"解析出二维码1数值："+barcode1.parsed)
                 # 转为灰度图像
                 gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                 barcodes = pyzbar.decode(gray)
                 barcodeData=''
                 if len(barcodes)>0:
                     barcode=barcodes[0]
                     barcodeData = barcode.data.decode("utf-8")
                     print(img+"解析出二维码2数值："+barcodeData)
                 data=[img,barcode1.parsed,barcodeData]
                 ws.append(data)
        wb.save(xlsfile)





