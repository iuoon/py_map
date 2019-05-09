#!/usr/bin/python
# -*- coding: UTF-8 -*-
import requests
import time

def rep(bounds):
    print(bounds)
    obj = '{}'
    while True:
        try:
            obj = requests.get('http://api.map.baidu.com/traffic/v1/bound?ak=zynT7H4G44DIeLB443zDXFsTwfsDjN8q&bounds='+bounds, timeout=30)
            break
        except requests.exceptions.ConnectionError:
            print('[ERROR]ConnectionError -- will retry connect')
            time.sleep(3)
    print(obj)
    data = obj.json()
    print(data)


def LocaDiv(ploy):
    list = []
    p0 = float(ploy.split(',')[0])
    p1 = float(ploy.split(',')[1])
    p2 = float(ploy.split(',')[2])
    p3 = float(ploy.split(',')[3])
    len1 = int((p2 - p0 + 0.0001) / 0.012)
    len2 = int((p3 - p1 + 0.0001) / 0.010)
    for i in range(0, len1):
        for j in range(0, len2):
            a = round(p0 + round(0.012 * i, 2), 6)
            b = round(p1 + round(0.010 * j, 2), 6)
            c = round(a+round(0.012 * 1, 2), 6)
            d = round(b+round(0.010 * 1, 2), 6)
            pos = str(b)+','+str(a)+';'+str(d)+','+str(c)
            list.append(pos)
    return list

if __name__ == "__main__":
  locs=LocaDiv('108.273095,22.78576,108.350709,22.859311')
  print(len(locs))
  for loc in locs:
      rep(loc)