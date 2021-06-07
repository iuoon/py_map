import wx
import requests
import time
import threading

APP_TITLE = u'身份认证'
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

        self.btn_start.Bind(wx.EVT_LEFT_DOWN, self.startWork)

        self.currentStep = ''
        self.entityName = ''

    def startWork(self, event):
        if self.host.GetValue() == '':
            self.area.AppendText("请输入host\n")
            return
        if self.host.GetValue() == '':
            self.area.AppendText("请输入cookie\n")
            return
        if self.currentStep == '':
            self.area.AppendText("请选择类型\n")
            return

        t1 = threading.Thread(target=self.preReadCard, )
        t1.start()

    def preReadCard(self):
        self.area.AppendText("开始认证\n")
        self.host.Disable()
        self.cookie.Disable()
        self.btn_start.Disable()
        self.ch1.Disable()
        for line in open("card.txt", "r", encoding="utf-8"):  # 设置文件对象并读取每一行文件
            arr = line.split("#")
            print(arr[1])
            print(arr[2])
            name = arr[1].replace("\n", '')
            cardNo = arr[2].replace("\n", '')
            if self.readCard(name, cardNo) == False:
                print("认证方式一失败，开始尝试第二种方式")
                self.area.AppendText(name + cardNo + "认证方式一失败，开始尝试第二种方式\n")
                if self.readCard2(name, cardNo) == False:
                    self.area.AppendText(name + cardNo + "认证失败,请更换cookie重试\n")
                else:
                    self.area.AppendText(name + cardNo + "认证成功\n")
            else:
                self.area.AppendText(name + cardNo + "认证成功\n")
            time.sleep(10)
        self.area.AppendText("认证结束\n")
        self.host.Enable()
        self.cookie.Enable()
        self.btn_start.Enable()
        self.ch1.Enable()

    def readCard(self, name, cardNo):

        header = {'Accept': 'application/json, text/javascript, */*; q=0.01',
                  'Referer': self.host.GetValue() + '/E-office/sit/employeeReadCard',
                  'x-requested-with': 'XMLHttpRequest',
                  'Accept-Encoding': 'gzip, deflate',
                  'Accept-Language': 'zh-CN',
                  'Cache-Control': 'max-age=0',
                  'Connection': 'keep-alive',
                  'Accept-Encoding': 'gzip, deflate',
                  # 'Host': '222.82.237.75:8070',
                  'Pragma': 'no-cache',
                  'Content-Type': 'application/x-www-form-urlencoded',
                  'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; InfoPath.3; .NET4.0C)',
                  'Cookie': self.cookie.GetValue()
                  }
        paramData = {
            "pkId": self.pkId.GetValue(),  # 27694
            "cardNo": cardNo,
            "entityName": self.entityName
        }
        ret = '{"status":False}'
        data = '{}'
        while True:
            try:
                ret = requests.post(self.host.GetValue() + '/E-office/sit/getEmployeeByIdCard', data=paramData,
                                    timeout=30, headers=header)
                print(ret.text)
                break
            except requests.exceptions.ConnectionError:
                print('[ERROR]ConnectionError -- will retry connect')
                time.sleep(3)
        try:
            data = ret.json()
        except:
            return False
        try:
            if data['status']:
                employeeData = data['data'][0]
                idCardPhoto = ''
                employeeId = '',
                for de in employeeData:
                    if de['field'] == 'employeeId':
                        employeeId = de['value']
                        idCardPhoto = de['idCardPhoto']
                        break
                print(idCardPhoto)
                print(str(employeeId))
                pa = {
                    "pkId": self.pkId.GetValue(),
                    "idCard": cardNo,
                    "idCardPhoto": str(idCardPhoto),
                    "idCardName": name,
                    "employeeId": employeeId,
                    "currentStep": self.currentStep,
                    "confirmStatus": "Y",
                    "sync": "false"
                }
                while True:
                    try:
                        obj = requests.post(self.host.GetValue() + '/E-office/sit/employeeIdConfirm', data=pa,
                                            timeout=30, headers=header)
                        try:
                            dt1 = obj.json()
                            if dt1['status'] == False:
                                return False
                        except:
                            return False
                        break
                    except requests.exceptions.ConnectionError:
                        print('[ERROR]ConnectionError -- will retry connect')
                        time.sleep(3)
            else:
                return False
        except:
            return False
        return True

    def readCard2(self, name, card):
        header = {'Accept': 'application/json, text/javascript, */*; q=0.01',
                  'Referer': self.host.GetValue() + '/E-office/sit/employeeReadCard',
                  'x-requested-with': 'XMLHttpRequest',
                  'Accept-Encoding': 'gzip, deflate',
                  'Accept-Language': 'zh-CN',
                  'Cache-Control': 'max-age=0',
                  'Connection': 'keep-alive',
                  'Accept-Encoding': 'gzip, deflate',
                  'Host': '222.82.237.75:8070',
                  'Pragma': 'no-cache',
                  'Content-Type': 'application/x-www-form-urlencoded',
                  'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; InfoPath.3; .NET4.0C)',
                  'Cookie': self.cookie.GetValue()
                  }
        paramData = {
            "licenseNum": self.licenseNum.GetValue(),
            "cardNo": str(card),
            "employeeName": str(name),
            "entityName": self.entityName
        }
        ret = '{"status":False}'
        data = '{}'
        while True:
            try:
                ret = requests.post(self.host.GetValue() + '/E-office/sit/getEmployeeByIdCard', data=paramData,
                                    timeout=30, headers=header)
                print(ret.text)
                break
            except requests.exceptions.ConnectionError:
                print('[ERROR]ConnectionError -- will retry connect')
                time.sleep(3)
        try:
            data = ret.json()
        except:
            return False
        if data['status']:
            employeeData = data['data'][0]
            idCardPhoto = ''
            employeeId = '',
            for de in employeeData:
                if de['field'] == 'employeeId':
                    employeeId = de['value']
                    idCardPhoto = de['idCardPhoto']
                    break
            pa = {
                "licenseNum": self.licenseNum.GetValue(),
                "idCard": card,
                "idCardPhoto": str(idCardPhoto),
                "idCardName": name,
                "employeeId": str(employeeId),
                "currentStep": self.currentStep,
                "confirmStatus": "Y",
                "sync": "false"
            }
            print(pa)
            while True:
                try:
                    obj = requests.post(self.host.GetValue() + '/E-office/sit/employeeIdConfirm', data=pa, timeout=30,
                                        headers=header)
                    try:
                        dt1 = obj.json()
                        print(obj.text)
                        if dt1['status'] == False:
                            return False
                    except:
                        return False
                    break
                except requests.exceptions.ConnectionError:
                    print('[ERROR]ConnectionError -- will retry connect')
                    time.sleep(3)
        else:
            return False
        return True


class mainApp(wx.App):
    def OnInit(self):
        self.SetAppName(APP_TITLE)
        self.Frame = mainFrame(None)
        self.Frame.Show()
        return True


if __name__ == "__main__":
    app = mainApp()
    app.MainLoop()
