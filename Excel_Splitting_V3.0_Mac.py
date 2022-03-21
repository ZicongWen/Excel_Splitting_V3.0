import wx
import os
import shutil
import sys
import pandas as pd
#import win32api


APP_TITLE = u'Excel_Splitting_V3.0'
APP_ICON = 'imag/Excel_Processing.ico'

INPUT_FILE_PATH = ""
OUTPUT_FILE_PATH = ""
KEYWORD = ""
KEYNUM = -1

def excel(input_file_path, keyword, output_file_path):
    if keyword.isdigit():
        data = pd.read_excel(input_file_path)
        key_col = data.columns[int(keyword)]
        area_list = list(set(data[key_col]))
        # data.to_excel(writer,sheet_name="总表",index=False)
        for j in area_list:
            df = data[data[key_col] == j]
            writer = pd.ExcelWriter(str(output_file_path) + str(j) + ".xlsx")
            df.to_excel(writer, index=False)
            writer.save()
        return True
    else:
        data = pd.read_excel(input_file_path)

        area_list = list(set(data[keyword]))

        # data.to_excel(writer,sheet_name="总表",index=False)

        for j in area_list:
            df = data[data[keyword] == j]
            writer = pd.ExcelWriter(str(output_file_path) + str(j) + ".xlsx")
            df.to_excel(writer, index=False)
            writer.save()
        return True


class mainFrame(wx.Frame):
    def __init__(self, parent):
            """构造函数"""

            wx.Frame.__init__(self, parent, -1, APP_TITLE)
            self.SetBackgroundColour(wx.Colour(224, 224, 224))
            self.SetSize((520, 220))
            self.Center()

            icon = wx.Icon(APP_ICON, wx.BITMAP_TYPE_ICO)
            self.SetIcon(icon)

            # 以下可以添加各类控件

            # 添加文本框
            wx.StaticText(self, -1, u'请选择文件：', pos=(40, 50), size=(200, -1), style=wx.ALIGN_LEFT)
            wx.StaticText(self, -1, u'请输入关键字/列：', pos=(40, 80), size=(200, -1), style=wx.ALIGN_LEFT)

            # 添加输入
            self.tc1 = wx.TextCtrl(self, -1, '', pos=(145, 50), size=(150, -1), name='TC01',
                                   style=wx.TE_CENTER | wx.TE_READONLY)
            self.tc2 = wx.TextCtrl(self, -1, '', pos=(145, 80), size=(150, -1), name='TC02', style=wx.TE_CENTER)
            self.tc3 = wx.TextCtrl(self, -1, '未开始', pos=(145, 110), size=(150, -1), name='TC03',
                                   style=wx.TE_CENTER | wx.TE_READONLY)

            # 添加按钮
            btn_mea = wx.Button(self, -1, u'选择文件位置..', pos=(350, 50), size=(100, 25))
            btn_meb = wx.Button(self, -1, u'开始', pos=(350, 80), size=(100, 25))
            btn_close = wx.Button(self, -1, u'关闭窗口', pos=(350, 110), size=(100, 25))

            # 控件事件 (Bind 方法)
            self.tc2.Bind(wx.EVT_TEXT, self.EvtText)
            self.Bind(wx.EVT_BUTTON, self.OnClose, btn_close)

            # 鼠标事件
            btn_mea.Bind(wx.EVT_LEFT_DOWN, self.OnLeftDown1)  # 左键点击按钮a
            btn_meb.Bind(wx.EVT_LEFT_DOWN, self.OnLeftDown2)

            # 系统事件
            self.Bind(wx.EVT_CLOSE, self.OnClose)
            self.Bind(wx.EVT_SIZE, self.On_size)
            # self.Bind(wx.EVT_PAINT, self.On_paint)
            # self.Bind(wx.EVT_ERASE_BACKGROUND, lambda event: None)

    def EvtText(self, evt):
        """输入框事件函数"""
        global KEYWORD
        KEYWORD = evt.GetString()

    def On_size(self, evt):
        """改变窗口大小事件函数"""
        self.Refresh()
        evt.Skip()  # 体会作用

    def OnClose(self, evt):
        """关闭窗口事件函数"""

        dlg = wx.MessageDialog(None, u'确定要关闭本窗口？', u'操作提示', wx.YES_NO | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_YES:
            self.Destroy()

    def OnLeftDown1(self, evt):
        """btn_A 左键按下事件函数"""
        dialog = wx.FileDialog(None, "Choose a directory:", style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        if dialog.ShowModal() == wx.ID_OK:
            # print(dialog.GetPath())
            global INPUT_FILE_PATH
            INPUT_FILE_PATH = str(dialog.GetPath())
            self.tc1.SetValue(str(dialog.GetPath()))
        dialog.Destroy()

    def OnLeftDown2(self, evt):
        """btn_B 左键按下事件"""
        self.tc3.SetValue("运行中")
        global OUTPUT_FILE_PATH, KEYWORD, KEYNUM
        indicator, OUTPUT_FILE_PATH = self.create_folder(INPUT_FILE_PATH)
        if indicator:
            for key in KEYWORD:
                if key.isdigit():
                    KEYNUM = key


        if KEYNUM != -1:
            flag = excel(INPUT_FILE_PATH, str(KEYNUM), OUTPUT_FILE_PATH)
            if flag:
                self.tc3.SetValue("完成！")
            else:
                self.tc3.SetValue("错误！请重试。")
        elif KEYNUM == -1:
            flag = excel(INPUT_FILE_PATH, KEYWORD, OUTPUT_FILE_PATH)

            if flag:
                self.tc3.SetValue("完成！")
            else:
                self.tc3.SetValue("错误！请重试。")
        else:
            self.tc3.SetValue("错误！请重试。")

    def create_folder(self, input_path):
        foldername = str(input_path[:-5]) + "拆分" + "/"
        isCreated = os.path.exists(foldername)
        if not isCreated:
            os.makedirs(foldername)
            return True, foldername
        else:
            YN2 = wx.MessageDialog(
                self,
                "文件夹已存在，是否重新生成？", "确认",
                wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION).ShowModal()

            if YN2 == wx.ID_YES:
                shutil.rmtree(foldername)
                os.makedirs(foldername)
                return True, foldername
            else:
                return False, "ERROR"
class mainApp(wx.App):
    def OnInit(self):
        self.SetAppName(APP_TITLE)
        self.Frame = mainFrame(None)
        self.Frame.Show()
        return True


if __name__ == "__main__":
    app = mainApp()
    app.MainLoop()