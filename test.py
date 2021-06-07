import pandas as pd
from bs4 import BeautifulSoup
import re

if __name__ == "__main__":
    with open("C:\\Users\\iuoon\\Desktop\\test2.html", "r", encoding="UTF-8") as f:  # 打开文件
        data = f.read()  # 读取文件
        soup = BeautifulSoup(data, 'html.parser')
        a_ctx = soup.findAll("a", {'class': 'btn-base btn-noborder icon-download'})  # 抓取a标签
        for ax in a_ctx:
            data_herf = ax.get('href')
            data_id = re.findall("\d+",data_herf)[0]
            print('获取到数据id:', data_id)
