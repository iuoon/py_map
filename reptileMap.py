# -*- coding:UTF-8 -*-
# author iuoon

import requests
import threading
import time
import xlwt
import os
import sys
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger


class LocaDiv(object):
    def __init__(self, loc_all):
        self.loc_all = loc_all

    def lat_all(self):
        lat_sw = float(self.loc_all.split(',')[0])
        lat_ne = float(self.loc_all.split(',')[2])
        lat_list = []
        for i in range(0, int((lat_ne - lat_sw + 0.0001) / 0.05)):  # 0.1为网格大小，可更改
            lat_list.append(round(lat_sw + round(0.05 * i, 2), 6))  # 0.05
        lat_list.append(lat_ne)
        return lat_list

    def lng_all(self):
        lng_sw = float(self.loc_all.split(',')[1])
        lng_ne = float(self.loc_all.split(',')[3])
        lng_list = []
        for i in range(0, int((lng_ne - lng_sw + 0.0001) / 0.04)):  # 0.1为网格大小，可更改
            lng_list.append(round(lng_sw + round(0.04 * i, 2), 6))  # 0.1为网格大小，可更改
        lng_list.append(lng_ne)
        return lng_list

    def ls_com(self):
        l1 = self.lat_all()
        l2 = self.lng_all()
        ab_list = []
        for i in range(0, len(l1)):
            a = str(l1[i])
            for i2 in range(0, len(l2)):
                b = str(l2[i2])
                ab = a + ',' + b
                print(ab)
                ab_list.append(ab)
        return ab_list

    def ls_row(self):
        l1 = self.lat_all()
        l2 = self.lng_all()
        ls_com_v = self.ls_com()
        ls = []
        for n in range(0, len(l1) - 1):
            for i in range(0 + len(l1) * n, len(l2) + (len(l2)) * n - 1):
                a = ls_com_v[i]
                b = ls_com_v[i + len(l2) + 1]
                ab = a + ';' + b
                ls.append(ab)
        return ls


def LocaDiv2(ploy):
    list = []
    p0 = float(ploy.split(',')[0])
    p1 = float(ploy.split(',')[1])
    p2 = float(ploy.split(',')[2])
    p3 = float(ploy.split(',')[3])
    len1 = int((p2 - p0 + 0.0001) / 0.05)
    len2 = int((p3 - p1 + 0.0001) / 0.04)
    for i in range(0, len1):
        for j in range(0, len2):
            a = round(p0 + round(0.05 * i, 2), 6)
            b = round(p1 + round(0.04 * j, 2), 6)
            c = round(a+round(0.05 * 1, 2), 6)
            d = round(b+round(0.04 * 1, 2), 6)
            pos = str(a)+','+str(b)+';'+str(c)+','+str(d)
            list.append(pos)
    return list

isreptiling = False
def reptileMap(key):
    print('key='+key)
    print('[info]开始爬取数据...')
    startTime = time.time()
    locs = LocaDiv2('116.208904,39.747315,116.550123,40.028783')
    #locs = loc0.ls_row()
    date = time.strftime("%Y%m%d-%H")

    dirs = os.path.abspath('.')+'\\'+time.strftime("%Y%m%d")
    # 创建文件夹
    if not os.path.exists(dirs):
        os.makedirs(dirs)
    # 删除旧文件
    file1 = dirs+'\\'+ date+'.xls'
    if os.path.exists(file1):
       os.remove(file1)

    dttime = time.strftime("%Y-%m-%d %H:%M:%S")
    count = 0
    workbook = xlwt.Workbook()
    sheet1 = workbook.add_sheet('beijing')
    keys1 = ['angle', 'direction', 'lcodes', 'name', 'polyline', 'speed', 'status', 'datetime']
    for i in range(0, len(keys1)):
        sheet1.write(0, i, keys1[i])       # 写入表头
    global isreptiling
    isreptiling = True

    for loc in locs:
        pa = {
            'key': str(key),
            # 'level': 6,                   # 道路等级为6，即返回的道路路况等级最小到无名道路这一级别
            'rectangle': str(loc),          # 矩形区域
            'extensions': 'all'
            # 'output': 'JSON'
        }
        print('[info]探测区块：'+loc)
        obj = '{}'
        while True:
            try:
                obj = requests.get('http://restapi.amap.com/v3/traffic/status/rectangle?', params=pa, timeout=30)
                break
            except requests.exceptions.ConnectionError:
                print('[ERROR]ConnectionError -- will retry connect')
                time.sleep(3)

        data = obj.json()
        if data['status'] == '0':
            print('[info]'+str(data))
            print('[warn]请求参数错误')
            continue

        for road in data['trafficinfo']['roads']:
            count = count+1

            rangle = road['angle'] if 'angle' in road else ''
            rdirection = road['direction'] if 'direction' in road else ''
            rlcodes = road['lcodes'] if 'lcodes' in road else ''
            rname = road['name'] if 'name' in road else ''
            rpolyline = road['polyline'] if 'polyline' in road else ''
            rspeed = road['speed'] if 'speed' in road else '0'
            rstatus = road['status'] if 'status' in road else ''

            sheet1.write(count, 0, rangle)
            sheet1.write(count, 1, rdirection)
            sheet1.write(count, 2, rlcodes)
            sheet1.write(count, 3, rname)
            sheet1.write(count, 4, rpolyline)
            sheet1.write(count, 5, rspeed)
            sheet1.write(count, 6, rstatus)
            sheet1.write(count, 7, dttime)

            # 判断有无路线，有些没有，直接取值会报错
            # if 'polyline' in road:
            #   rps = road['polyline'].split(";")
            #   for j in range(0, len(rps)):

        time.sleep(1)    # 间隔1s执行一次分块请求，避免并发度高被限制
    workbook.save(file1)
    endTime = time.time()
    print('[info]数据爬取完毕，用时%.2f秒' % (endTime-startTime))
    print('[info]数据存储路径：'+file1)
    isreptiling = False

def test(key):
    global isreptiling
    isreptiling = True
    print(key)

scheduler = BlockingScheduler()
def startWork(key, loop):
    test(key)
    if loop:
        # 周一到周日,每小时执行一次   每5秒second='*/5' hour='0-23'
        trigger = CronTrigger(day_of_week='0-6', second='*/5')
        scheduler.add_job(test, trigger, args=(key,))
        # 周六到周日,24
        # scheduler.add_job(test, 'cron', day_of_week='5-6', hour='0-23')
        scheduler.start()

def onEnter():
    # 监听输入,没有输入时该线程会阻塞住
    while True:
      line = input()
      if str(line).lower() == 'exit':
          if isreptiling:
              print('[warn]历史数据已保存，当前正在爬取中,是否丢弃本次爬取  -- 输入 Y or N')
          else:
              scheduler.shutdown(wait=False)
              exit('已退出')
              break
      if str(line).lower() == 'y' and isreptiling:
          scheduler.shutdown(wait=False)
          exit('已退出')
          break



    return True


if __name__ == '__main__':
    print('**************************EXE执行方式**************************************')
    print('[info]执行命令案例：reptileMap.exe -key 0b1804994cd63974f873a29a269d65e7')
    print('**************************脚本执行方式**************************************')
    print('[info]请输入高德地图key,缺省时将使用开发者预留key（保存5天，5天后预留key删除）')
    print('[info]执行命令(带key)案例：python reptileMap.py -key 0b1804994cd63974f873a29a269d65e7')
    print('[info]每小时爬取一次,文件存放在当前目录')
    print('[info]当前窗口输入exit按回车键，即可退出程序,文件自动保存')
    param = sys.argv
    key = '0b1804994cd63974f873a29a269d65e7'
    loop = True  # 改为默认一直执行
    for i in range(0, len(param)):
        if param[i] == '-key':
            key = param[i+1]
    #   if param[i] == '-loop':
    #       loop = True
    t1=threading.Thread(target=onEnter)
    t1.start()
    startWork(key, loop)


