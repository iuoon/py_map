import wx
import pytesseract
from PIL import Image

APP_TITLE = u'文字识别'
APP_ICON = 'res/python.ico'


class MainFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, -1, APP_TITLE)
        self.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.SetSize((400, 160))
        self.Center()

        self.pathLabel = wx.StaticText(self, -1, u'文件路径：', pos=(10, 20), size=(60, -1), name='pathLabel', style=wx.TE_LEFT)
        self.path = wx.StaticText(self, -1, u'未选择', pos=(80, 20), size=(200, -1), name='path', style=wx.TE_LEFT)

        self.selectBtn = wx.Button(self, -1, u'选择', pos=(10, 50), size=(60, -1), style=wx.ALIGN_LEFT)
        self.startBtn = wx.Button(self, -1, u'开始', pos=(100, 50), size=(60, -1), style=wx.ALIGN_LEFT)

        self.Bind(wx.EVT_BUTTON, self.OnSelect, self.selectBtn)
        self.Bind(wx.EVT_BUTTON, self.OnStart, self.startBtn)

    def OnSelect(self, event):
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            print(dlg.GetPath())  # 文件夹路径
            self.path.SetLabelText(dlg.GetPath())
        dlg.Destroy()

    def OnStart(self, event):
        print(1)
        self.selectBtn.Disable()
        self.startBtn.Disable()
        print(self.path.GetLabelText())
        image = Image.open("C:\\Users\\iuoon\\Desktop\\10个版样\\20180609075358350.jpg")  # 打开图片
        image.load()  # 加载一下图片，防止报错，此处可省略
        # image.show()  # 调用show来展示图片，调试用，可省略
        pytesseract.pytesseract.tesseract_cmd = 'G:/tesseract-4.0.0/tesseract-4.0.0-alpha/tesseract.exe'
        tessdata_dir_config = '--tessdata-dir "G:/tesseract-4.0.0/tesseract-4.0.0-alpha/tessdata"'
        text = pytesseract.image_to_string(image,lang='chi_sim', config=tessdata_dir_config)
        print(text)
        self.selectBtn.Enable()
        self.startBtn.Enable()


class MainApp(wx.App):
    def OnInit(self):
        self.SetAppName(APP_TITLE)
        self.Frame = MainFrame(None)
        self.Frame.Show()
        return True


if __name__ == "__main__":
    app = MainApp()
    app.MainLoop()
