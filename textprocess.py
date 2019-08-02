#!/usr/bin/python
# -*- coding: UTF-8 -*-

def is_contain_chinese(check_str):
    """
    判断字符串中是否包含中文
    :param check_str: {str} 需要检测的字符串
    :return: {bool} 包含返回True， 不包含返回False
    """
    for ch in check_str:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False


if __name__ == "__main__":
    with open("新建文本文档.txt","r+",encoding='UTF-8') as fr,open("test_2.txt","w+", encoding='UTF-8') as fw:                   #以读的方式打开
      list = fr.readlines()
      for str in list:
          flag=True
          r1=str.find(".com")
          r2=str.find(".cn")
          if r1 !=-1 or r2 !=-1:
              flag=False
              continue
          if len(str)>30:
              flag=False
              continue
          if is_contain_chinese(str):
              flag=False
              continue
          arr=str.split("----")
          if len(arr)>1:
             if len(arr[1].strip())<5:
                 flag=False
                 continue
          if flag:
              fw.write(str)                                        #修改的内容重写进文件
              print(str)