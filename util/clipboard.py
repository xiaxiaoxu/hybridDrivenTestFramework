#encoding=utf-8
import win32clipboard as w
import win32con
import time
class Clipboard(object):
    #模拟Windows设置剪贴板

    #读取剪贴板
    @staticmethod
    def getText():
        #打开剪贴板
        w.OpenClipboard()
        #获取剪贴板中的数据
        content=w.GetClipboardData(win32con.CF_TEXT)
        #关闭剪贴板
        w.CloseClipboard()
        #返回剪贴板数据
        return content


    #设置剪贴板内容
    @staticmethod
    def setText(aString):
        #打开剪贴板
        w.OpenClipboard()
        #清空剪贴板
        w.EmptyClipboard()
        #将数据aString写入剪贴板
        w.SetClipboardData(win32con.CF_UNICODETEXT,aString)
        #关闭剪贴板
        w.CloseClipboard()

if __name__=='__main__':

    Clipboard.setText(u'hey buddy!')
    time.sleep(3)
    print Clipboard.getText()

