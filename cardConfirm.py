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

        wx.StaticText(self, -1, u'请求地址：', pos=(10, 20), size=(60, -1), style=wx.ALIGN_LEFT)
        self.host = wx.TextCtrl(self, -1, '', pos=(70, 20), size=(260, -1), name='host', style=wx.TE_LEFT)
        wx.StaticText(self, -1, u'Cookie：', pos=(10, 50), size=(60, -1), style=wx.ALIGN_LEFT)
        self.cookie = wx.TextCtrl(self, -1, '', pos=(70, 50), size=(260, -1), name='host', style=wx.TE_LEFT)
        self.area = wx.TextCtrl(self, -1, '', pos=(10, 130), size=(320, 200), name='area', style=wx.TE_LEFT | wx.TE_MULTILINE)

        self.btn_start = wx.Button(self, -1, u'开始批量认证', pos=(350, 20), size=(100, 25))

        wx.StaticText(self, -1, u'请选择认证类型：', pos=(10, 90), size=(100, -1), style=wx.ALIGN_LEFT)
        pros=['现场管理人员', '注册建造师', '技术工人']
        self.ch1 = wx.ComboBox(self,-1,value='请选择',choices=pros,pos=(110, 90))

        self.btn_start.Bind(wx.EVT_LEFT_DOWN, self.startWork)
        self.Bind(wx.EVT_COMBOBOX, self.OnTypeChoice, self.ch1)

        self.currentStep=''
        self.entityName=''

    def OnTypeChoice(self,evt):
        type= evt.GetString()
        if type == '现场管理人员':
            self.currentStep='xcglryInfo'
            self.entityName='Administrator'
        if type == '注册建造师':
            self.currentStep='zyryInfo'
            self.entityName='Practitioners'
        if type == '技术工人':
            self.currentStep='jsgrInfo'
            self.entityName='Workers'


    def startWork(self,event):
        if self.host.GetValue()=='':
            self.area.AppendText("请输入host\n")
            return
        if self.host.GetValue()=='':
            self.area.AppendText("请输入cookie\n")
            return
        if self.currentStep=='':
            self.area.AppendText("请选择类型\n")
            return

        t1 = threading.Thread(target=self.preReadCard,)
        t1.start()

    def preReadCard(self):
        self.area.AppendText("开始认证\n")
        self.host.Disable()
        self.cookie.Disable()
        self.btn_start.Disable()
        self.ch1.Disable()
        for line in open("card.txt","r",encoding="utf-8"): #设置文件对象并读取每一行文件
            arr=line.split("#")
            print(arr[0])
            print(arr[1])
            name=arr[0].replace("\n",'')
            cardNo=arr[1].replace("\n",'')
            if self.readCard(name,cardNo) == False:
                print("认证方式一失败，开始尝试第二种方式")
                self.area.AppendText(name+cardNo+"认证方式一失败，开始尝试第二种方式\n")
                if self.readCard2(name,cardNo) ==False:
                    self.area.AppendText(name+cardNo+"认证失败,请更换cookie重试\n")
                else:
                    self.area.AppendText(name+cardNo+"认证成功\n")
            time.sleep(10)
        self.area.AppendText("认证结束\n")
        self.host.Enable()
        self.cookie.Enable()
        self.btn_start.Enable()
        self.ch1.Enable()

    def readCard(self,name,cardNo):

        header = {'Accept': 'application/json, text/javascript, */*; q=0.01',
                  'Referer': self.host.GetValue()+'/E-office/sit/employeeReadCard',
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
        paramData={
            "pkId":"27694",
            "cardNo":cardNo,
            "entityName":self.entityName
        }
        ret='{"status":False}'
        data='{}'
        while True:
            try:
                ret = requests.post(self.host.GetValue()+'/E-office/sit/getEmployeeByIdCard', data = paramData, timeout=30, headers=header)
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
                print(data['data'][0][5])
                employee=data['data'][0][5]
                print(employee['idCardPhoto'])
                print(str(employee['value']))
                pa = {
                    "pkId":"27694",
                    "idCard":cardNo,
                    "idCardPhoto":str(employee['idCardPhoto']),
                    "idCardName":name,
                    "employeeId":str(employee['value']),
                    "currentStep":self.currentStep,
                    "confirmStatus":"Y",
                    "sync":"false"
                }
                while True:
                    try:
                        obj = requests.post(self.host.GetValue()+'/E-office/sit/employeeIdConfirm', data = pa, timeout=30, headers=header)
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


    def readCard2(self,name,card):
        header = { 'Accept': 'application/json, text/javascript, */*; q=0.01',
                   'Referer': self.host.GetValue()+'/E-office/sit/employeeReadCard',
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
        paramData={
            "licenseNum":'91650100397711723R',
            "cardNo":str(card),
            "employeeName":str(name),
            "entityName":self.entityName
        }
        ret='{"status":False}'
        data='{}'
        while True:
            try:
                ret = requests.post(self.host.GetValue()+'/E-office/sit/getEmployeeByIdCard', data = paramData, timeout=30, headers=header)
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
            print(data['data'][0][4])
            employee=data['data'][0][4]
            print(employee['idCardPhoto'])
            print(str(employee['value']))
            pa = {
                "licenseNum":"91650100397711723R",
                "idCard":card,
                "idCardPhoto":str(employee['idCardPhoto']),
                "idCardName":name,
                "employeeId":str(employee['value']),
                "currentStep":self.currentStep,
                "confirmStatus":"Y",
                "sync":"false"
            }
            while True:
                try:
                    obj = requests.post(self.host.GetValue()+'/E-office/sit/employeeIdConfirm', data = pa, timeout=30, headers=header)
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
