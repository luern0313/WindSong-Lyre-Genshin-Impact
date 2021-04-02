import sys, os
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
from system_hotkey import SystemHotkey
from PyQt5.QtCore import QSize,QPoint,Qt,QRect,pyqtSignal
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QKeySequence,QIcon
from PyQt5.QtWidgets import QWidget, QVBoxLayout,QLabel,QListWidget,QApplication
import ctypes
from 疯物之诗琴 import PlayThread,is_admin
import time

class playWindow(QWidget):
    sig_keyhot = pyqtSignal(str)
    def __init__(self,parent = None):
        super(playWindow,self).__init__(parent)
        #1. 设置演奏线程
        self.playThread = PlayThread()
        #2. 设置我们的自定义热键响应函数
        self.sig_keyhot.connect(self.MKey_pressEvent)
        #3. 初始化两个热键
        self.hk_stop = SystemHotkey()
        #4. 绑定快捷键和对应的信号发送函数
        try:
            self.hk_stop.register(('control', 'shift', 'g'), callback=lambda x: self.send_key_event("stop"))
        except InvalidKeyError as e:
            QMessageBox(QMessageBox.Warning,'警告','热键设置失败').exec_()
            print(e)
        except SystemRegisterError as e:
            QMessageBox(QMessageBox.Warning,'警告','热键设置冲突').exec_()
            print(e)
        #5.设置pyqt5的快捷键
        QShortcut(QKeySequence("Escape"), self, self.stopTool)
        #6.设置图形界面
        self.setupUi()

    def setupUi(self):
        self.setWindowTitle("风神之诗琴")
        self.setWindowIcon(QIcon('icon.ico'))
        self.setFixedSize(QSize(360,200))
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool) 
        #self.setAttribute(Qt.WA_TranslucentBackground)
        

        self.widgetLayout = QVBoxLayout()#创建垂直布局
        self.widgetLayout.setObjectName("widgetLayout")
        self.setLayout(self.widgetLayout)
        self.playList = QListWidget()
        self.playList.setGeometry(QRect(0, 40, 340, 60))
        self.playList.setMinimumSize(QSize(340, 60))
        self.playList.setBaseSize(QSize(340, 60))
        try:
            self.fileList = os.listdir("midi/")
            self.playList.addItems(self.fileList)
        except FileNotFoundError as e:
            QMessageBox(QMessageBox.Warning,'警告','没有找到midi文件').exec_()
            print(e)
        
        self.msgLabel = QLabel('双击列表选项开始或停止演奏\nEsc退出程序，Ctrl+Shift+G停止演奏')
        self.msgLabel.setGeometry(QRect(0, 0, 360, 40))
        self.msgLabel.setMinimumSize(QSize(360, 40))
        self.msgLabel.setBaseSize(QSize(360, 40))
        self.msgLabel.setAlignment(Qt.AlignLeft)
        self.msgLabel.setObjectName("msgLabel")

        self.playStatus = QLabel('请选择一首音乐开始演奏')
        self.playStatus.setGeometry(QRect(0, 120, 360, 20))
        self.playStatus.setMinimumSize(QSize(360, 20))
        self.playStatus.setBaseSize(QSize(360, 20))
        #添加控件到布局中
        self.widgetLayout.addWidget(self.msgLabel)
        self.widgetLayout.addWidget(self.playList)
        self.widgetLayout.addWidget(self.playStatus)
        #绑定操作
        self.playList.itemClicked.connect(self.playItemClicked)
        self.playList.doubleClicked.connect(self.playMidi)
        self.playThread.playSignal.connect(self.showStopPlay)

    def playItemClicked(self,item):
        print('你选择了：' + item.text())
        self.playStatus.setText('你选择了：' + item.text())

    
    #热键处理函数
    def MKey_pressEvent(self,i_str):
        print("按下的按键是%s" % (i_str,))
        self.stopPlayThread()
        
    #热键信号发送函数(将外部信号，转化成qt信号)
    def send_key_event(self,i_str):
        self.sig_keyhot.emit(i_str)
    
    def playMidi(self,index):
        self.stopPlayThread()
        print('开始演奏：'+self.fileList[index.row()])
        self.playStatus.setText('开始演奏：'+self.fileList[index.row()])
        self.playThread.setFilePath("midi/"+self.fileList[index.row()])
        self.playThread.start()
        pass

    def showStopPlay(self,msg):
        self.playStatus.setText(msg)

    def stopPlayThread(self):
        self.playStatus.setText('停止演奏')
        self.playThread.stopPlay()
        time.sleep(0.1)
        if not self.playThread.isFinished():
                self.playThread.terminate()
                self.playThread.wait()
        return

    def stopTool(self):
        self.stopPlayThread()
        time.sleep(0.1)
        try:
            self.hk_stop.unregister(('control', 'shift', 'g'))
        except UnregisterError as e:
            QMessageBox(QMessageBox.Warning,'警告','热键注销失败').exec_()
            print(e)
        self.close()
        print('退出了应用！！！')

def main():
    app = QApplication(sys.argv)
    mainWindow = playWindow()
    mainWindow.show()
    sys.exit(app.exec_())
    pass

if __name__ == '__main__':
    if is_admin():
        main()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)