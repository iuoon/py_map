import wx
import wx.lib.filebrowsebutton
from openpyxl import load_workbook

APP_TITLE = u'表格处理'
APP_ICON = 'res/python.ico'

class mainFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, -1, APP_TITLE)
        self.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.SetSize((520, 300))
        self.Center()

        self.btn_selectFile = wx.lib.filebrowsebutton.FileBrowseButton(self,pos=(20, 40),size=(400,-1),style=wx.ALIGN_RIGHT,labelText='', fileMask='*.xlsx',buttonText='打开',changeCallback=self.OnRead)

        self.tip = wx.StaticText(self, -1, u'', pos=(20, 100), size=(400, -1), style=wx.ST_NO_AUTORESIZE)

        self.obj1={'X':[],'YY':[],'OOO':[],'UUUU':[]}
        self.obj2={'X':[],'YY':[],'OOO':[],'UUUU':[]}


    def On_size(self, evt):
        # 改变窗口大小事件函数
        self.Refresh()
        evt.Skip()

    def OnRead(self, event):
        """"""
        self.btn_selectFile.Disable()
        self.tip.SetLabel('数据处理中，请不要关闭')
        print(self.btn_selectFile.GetValue())
        wb = load_workbook(self.btn_selectFile.GetValue())
        sheet = wb.active
        rnum=sheet.max_row
        cnum=sheet.max_column
        for r in range(3,rnum):
            value1=sheet.cell(row=r, column=3).value
            value2=sheet.cell(row=r, column=5).value
            value3=sheet.cell(row=r, column=6).value
            # print(value1,value2,value3)
            if value1 is None:
                continue
            a1=value1.split("=")
            atype=a1[0].split(" ")[1]
            #print(atype,',',a1[1],',',value2,',',value3)
            if atype == 'A':
                atype='X'
            if atype == 'BB':
                atype='YY'
            if atype == 'CC':
                atype='OOO'
            if atype == 'DX':
                atype='UUUU'
            obj1={'mc':0,'type':atype, 'value':a1[1], 'row':r}  # mc--匹配度  type--类型  value--字符串 row--所在字符串的行数
            obj2={'type':value3, 'value':value2, 'row':r}
            self.obj1[atype].append(obj1)
            self.obj2[value3].append(obj2)
        print('---------------------------------------------')
        for key in self.obj1:
            arr1=self.obj1[key]
            arr2=self.obj2[key]
            for obj1 in arr1:
                value=obj1['value']
                nlen=len(value)
                tarr=[]
                tarr.append(value)
                for i in range (0,nlen):
                    tarr.append(value[i:i+1])
                    if i+2 < nlen:
                        tarr.append(value[i:i+2])
                    if i+3 < nlen:
                        tarr.append(value[i:i+3])
                    if i+4 < nlen:
                        tarr.append(value[i:i+4])
                    if i+5 < nlen:
                        tarr.append(value[i:i+5])
                    if i+6 < nlen:
                        tarr.append(value[i:i+6])
                    if i+7 < nlen:
                        tarr.append(value[i:i+7])
                    if i+8 < nlen:
                        tarr.append(value[i:i+8])
                for i in range (1,nlen):
                    tarr.append(value[0:i])

                for t in tarr:
                    for obj2 in arr2:
                        if obj2['value'].find(t) == 1:
                            obj1['mc']+=1
            # 查找完毕，开始排序,默认最大放在最前面
            arr1.sort(key=lambda obj: obj['mc'], reverse=True)
            print(arr1[0])
            if arr1[0]['mc'] > 0:
                sheet.cell(row=arr1[0]['row'], column=4).value =arr1[0]['value']

        wb.save(self.btn_selectFile.GetValue().split(".xls")[0]+'_new.xlsx')
        self.btn_selectFile.Enable()
        self.tip.SetLabel('数据处理完成')

    # 数据处理 核心块
    def process(self):
        print(1)
        for key in self.obj1:
            arr1=self.obj1[key]
            arr2=self.obj2[key]
            for obj1 in arr1:
                value=obj1['value']
                nlen=len(value)
                tarr=[]
                for i in range (1,nlen):
                    tarr.append(value[0:i])
                for t in tarr:
                    for obj2 in arr2:
                        if obj2['value'].find(t) == 1:
                            obj1['mc']+=1
            # 查找完毕，开始排序,默认最大放在最前面
            arr1.sort(key=lambda obj: obj['mc'], reverse=True)
            print(arr1[0])
        self.btn_selectFile.Enable()
        self.tip.SetLabel('数据处理完成')




class mainApp(wx.App):
    def OnInit(self):
        self.SetAppName(APP_TITLE)
        self.Frame = mainFrame(None)
        self.Frame.Show()
        return True

if __name__ == "__main__":
    app = mainApp()
    app.MainLoop()
