# coding=utf-8
import wx
import requests
import threading
import os
from urllib.request import urlretrieve
import urllib
from urllib import parse
from openpyxl import load_workbook
import time
# import pandas as pd
from bs4 import BeautifulSoup
import re

APP_TITLE = u'下载文件工具'
APP_ICON = 'res/python.ico'


class mainFrame(wx.Frame):


    def __init__(self, parent):
        wx.Frame.__init__(self, parent, -1, APP_TITLE)
        self.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.SetSize((600, 480))
        self.Center()

        self.selectExcelBtn = wx.Button(self, -1, u'选择企业Excel', pos=(10, 20), size=(100, -1), style=wx.ALIGN_LEFT)

        # wx.StaticText(self, -1, u'请求地址：', pos=(10, 20), size=(60, -1), style=wx.ALIGN_LEFT)
        self.excelFile = wx.TextCtrl(self, -1, '', pos=(130, 20), size=(260, -1), name='excelFile', style=wx.TE_LEFT)

        self.selectOutPathBtn = wx.Button(self, -1, u'下载文件路径：', pos=(10, 50), size=(100, -1), style=wx.ALIGN_LEFT)
        self.outPath = wx.TextCtrl(self, -1, '', pos=(130, 50), size=(260, -1), name='outPath', style=wx.TE_LEFT)

        wx.StaticText(self, -1, u'Cookie：', pos=(10, 82), size=(60, -1), style=wx.ALIGN_LEFT)
        self.cookie = wx.TextCtrl(self, -1, '', pos=(70, 80), size=(260, -1), name='Cookie', style=wx.TE_LEFT)

        self.area = wx.TextCtrl(self, -1, '', pos=(10, 190), size=(320, 200), name='area',
                                style=wx.TE_LEFT | wx.TE_MULTILINE)

        self.btn_start = wx.Button(self, -1, u'开始下载', pos=(10, 160), size=(80, -1))

        self.Bind(wx.EVT_BUTTON, self.OnSelectExcel, self.selectExcelBtn)
        self.Bind(wx.EVT_BUTTON, self.OnSelectOutPath, self.selectOutPathBtn)
        self.btn_start.Bind(wx.EVT_LEFT_DOWN, self.startWork)

    def OnSelectExcel(self, event):
        dlg = wx.FileDialog(self, message=u"选择文件",
                            defaultDir=os.getcwd(),
                            defaultFile="")
        if dlg.ShowModal() == wx.ID_OK:
            print(dlg.GetPath())  # 文件夹路径
            self.excelFile.SetLabelText(dlg.GetPath())
        dlg.Destroy()

    def OnSelectOutPath(self, event):
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            print(dlg.GetPath())  # 文件夹路径
            self.outPath.SetLabelText(dlg.GetPath())
        dlg.Destroy()

    def startWork(self, event):
        if self.excelFile.GetValue() == '':
            self.area.AppendText("请选择Excel文件\n")
            return
        if self.outPath.GetValue() == '':
            self.area.AppendText("请选择输出文件夹\n")
            return

        t1 = threading.Thread(target=self.pre_work)
        t1.start()

    def pre_work(self):

        self.excelFile.Disable()
        self.outPath.Disable()
        self.cookie.Disable()
        self.btn_start.Disable()
        self.area.AppendText("开始加载企业列表\n")
        ent_list = self.read_excel()
        size = len(ent_list)
        if size <= 0:
            self.area.AppendText("加载企业失败,结束下载\n")
            return

        self.area.AppendText("加载完毕，开始下载文件\n")
        for r in range(0, size):
            ent_name = ent_list.pop(r)
            self.download(ent_name)
            time.sleep(0.5)

        self.area.AppendText("认证结束\n")
        self.excelFile.Enable()
        self.outPath.Enable()
        self.cookie.Enable()
        self.btn_start.Enable()

    def read_excel(self):
        ent_list = []
        print('excel：', self.excelFile.GetValue())
        wb = load_workbook(self.excelFile.GetValue())
        sheet = wb.active
        rnum = sheet.max_row + 1
        cnum = sheet.max_column
        for r in range(3, rnum):
            ent_name = sheet.cell(row=r, column=2).value
            print(ent_name)
            if ent_name is None:
                continue
            ent_list.append(ent_name)
        return ent_list

    def download(self, ent_name):
        down_url = "http://www.baidu.com/" + parse.quote(ent_name) + ".pdf"
        self.download_file2(down_url, self.outPath.GetLabel())
        return True

    def download_file1(self, url, store_path):
        filename = url.split("/")[-1]
        filepath = os.path.join(store_path, filename)

        file_data = requests.get(url, allow_redirects=True).content
        with open(filepath, 'wb') as handler:
            handler.write(file_data)

    def download_file2(self, url, store_path):
        while True:
            try:
                # store_path = store_path.encode(encoding='UTF-8',errors='strict')
                filename = url.split("/")[-1]
                filepath = os.path.join(store_path, filename)
                urlretrieve(url, filepath)
                break
            except urllib.error.HTTPError as e:
                self.area.AppendText(filename + '下载异常：' + e.reason + '\n')
                print(e.reason)
                print(e.code)
                time.sleep(3)

    def searchEnt(self, ent_name):
        header = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Referer': 'http://10.100.248.214/manager/reportInfo/list?hasReport=false',
            'x-requested-with': 'XMLHttpRequest',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Accept-Encoding': 'gzip, deflate',
            'Host': '10.100.248.214',
            'Origin': 'http://10.100.248.214',
            'Pragma': 'no-cache',
            'Content-Type': 'application/x-www-form-urlencoded',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.77 Safari/537.36',
            'Cookie': self.cookie.GetLabel()
        }
        paramData = {
            'industryName': '',
            'industryCode': '',
            'provinceCode': '',
            'cityCode': '',
            'countyCode': '',
            'permitCode': '',
            'companyName': ent_name,
            'startFztime': '',
            'endFztime': '',
            'fzjg': False,
            'first': '',
            'isUpdate': '',
            'page': 1
        }
        try:
            ret = requests.post('http://10.100.248.214/manager/reportInfo/list?hasReport=true', data=paramData,
                                timeout=30, headers=self.header)
            # df = pd.read_html(ret.text)[0]   # 该页面只返回一个表格，所以取第0个表格
            # print(df)
            soup = BeautifulSoup(ret.text, 'html.parser')
            a_ctx = soup.findAll("a", {'target': 'rightFrame'})  # 抓取a标签
            req_url2 = ''
            for ax in a_ctx:
                req_url2 = ax.get('href')
                print('获取到企业详情url', req_url2)

            if len(req_url2) <= 0:
                print('没有查询到企业')
                return
            req_url2 = 'http://10.100.248.214' + req_url2

        except requests.exceptions.ConnectionError:
            print('[ERROR]ConnectionError -- will retry connect')

    def searchPdf(self, req_url2):
        header = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Referer': 'http://10.100.248.214/manager/reportInfo/list?hasReport=false',
            'x-requested-with': 'XMLHttpRequest',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Accept-Encoding': 'gzip, deflate',
            'Host': '10.100.248.214',
            'Origin': 'http://10.100.248.214',
            'Pragma': 'no-cache',
            'Content-Type': 'application/x-www-form-urlencoded',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.77 Safari/537.36',
            'Cookie': self.cookie.GetLabel()
        }

        token = "a247c4ab8ab04ff1b470c9ff3bdd95fb"
        req_url3 = 'http://10.100.248.214/report//pdf/month?reportId=dataid&token=a247c4ab8ab04ff1b470c9ff3bdd95fb'
        try:
            # 获取2018年的数据
            ret = requests.get(req_url2 + '&year=2018',
                               timeout=30, headers=self.header)
            soup = BeautifulSoup(ret.text, 'html.parser')
            a_ctx = soup.findAll("a", {'class': 'btn-base btn-noborder icon-download'})  # 抓取a标签
            for ax in a_ctx:
                data_herf = ax.get('href')
                data_id = re.findall("\d+", data_herf)[0]
                print('获取到数据id:', data_id)

                # 在循环内一个个开始下载文件
                req_url4 = req_url3.replace("dataid", data_id)


        except requests.exceptions.ConnectionError:
            print('[ERROR]ConnectionError -- will retry connect')


class mainApp(wx.App):
    def OnInit(self):
        self.SetAppName(APP_TITLE)
        self.Frame = mainFrame(None)
        self.Frame.Show()
        return True


if __name__ == "__main__":
    app = mainApp()
    app.MainLoop()
