import pandas as pd
from bs4 import BeautifulSoup
import re

if __name__ == "__main__":
    with open("C:\\Users\\iuoon\\Desktop\\test2.html", "r", encoding="UTF-8") as f:  # 打开文件
        data = f.read()  # 读取文件
        # df = pd.read_html(data)[0]
        # print(df)
        soup = BeautifulSoup(data, 'html.parser')

        a_ctx = soup.findAll("a", {'class': 'btn-base btn-noborder icon-download'})  # 抓取a标签
        for ax in a_ctx:
            # 获取元素的父级元素
            parent_text = ax.parent.parent.text
            if parent_text.find("年报") == -1:
               continue
            data_herf = ax.get('href')
            data_id = re.findall("\d+",data_herf)[0]
            print('获取到数据id:', data_id)
