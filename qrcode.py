#!/usr/bin/python
# -*- coding: UTF-8 -*-
import pyzbar.pyzbar as pyzbar
import cv2
import zxing
import os
from openpyxl import Workbook


if __name__ == "__main__":
    s=os.getcwd() #获取当前文件的位置如 'C:\\Users\\Administrator\\Desktop\\TEXT'
    # s=os.path.split(s)[0] #去掉末尾\\后面的内容，变成 'C:\\Users\\Administrator\\Desktop'
    s=s+'\\jdk' #变成 'C:\\Users\\Administrator\\Desktop\\Mozilla Firefox'
    w=s+"\\bin;"+s+"\\jre\\bin;"
    os.environ['JAVA_HOME']=s
    os.environ['Path']=os.environ['Path']+";"+w #在环境变量中Path后添加字符串s

    xlsfile = 'output.xlsx'
    imgs=[]
    for root, dirs, files in os.walk(os.path.abspath('./work/')):
        print(root)
        if len(files) > 0:
            imgs = files
            break
    if len(imgs) == 0:
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
            if img.endswith('.jpg') or img.endswith('.png') or img.endswith('.tiff'):
                 print("开始解析文件："+root+"\\"+img+"\n")
                 frame= cv2.imread(root+"\\"+img)
                 #x, y = frame.shape[0:2]
                 #frame = cv2.resize(frame, (int(y * 2), int(x * 2)))
                 cropped = frame[0:270, 0:512]  # 裁剪坐标为[y0:y1, x0:x1]
                 # x, y = cropped.shape[0:2]
                 # cropped = cv2.resize(cropped, (int(y * 2), int(x * 2)))
                 nameArr=img.split(".")
                 cv2.imwrite(root+"\\tmp\\"+nameArr[0]+".jpg", cropped)


                 reader1 = zxing.BarCodeReader()
                 barcode1 = reader1.decode(dd+nameArr[0]+".jpg")
                 print(img+"解析出条形码1数值："+barcode1.parsed)
                 # 转为灰度图像
                 gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                 barcodes = pyzbar.decode(gray)
                 barcodeData=''
                 if len(barcodes)>0:
                     barcode=barcodes[0]
                     barcodeData = barcode.data.decode("utf-8")
                     print(img+"解析出条形码2数值："+barcodeData)
                 data=[img,barcode1.parsed,barcodeData]
                 ws.append(data)
        wb.save(xlsfile)





