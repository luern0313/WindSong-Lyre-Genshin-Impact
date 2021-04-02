import ctypes
import os
import sys
import time

from PyQt5.QtCore import QSize, Qt, QRect, pyqtSignal, QCoreApplication
from PyQt5.QtGui import QKeySequence, QIcon
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QListWidget, QApplication, QShortcut, QMessageBox
from system_hotkey import SystemHotkey, SystemRegisterError, InvalidKeyError, UnregisterError

from 疯物之诗琴 import PlayThread, is_admin

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']


class playWindow(QWidget):
    sig_hot_key = pyqtSignal(str)

    def __init__(self, parent=None):
        super(playWindow, self).__init__(parent)
        # 创建自动演奏线程
        self.playThread = PlayThread()
        # ---------设置全局快捷键----------
        # 设置我们的自定义热键响应函数
        self.sig_hot_key.connect(self.mkey_press_event)
        # 初始化热键
        self.hk_stop = SystemHotkey()
        # 绑定快捷键和对应的信号发送函数
        try:
            self.hk_stop.register(('control', 'shift', 'g'), callback=lambda x: self.send_key_event("stop"))
        except InvalidKeyError as e:
            QMessageBox(QMessageBox.Warning, '警告', '热键设置失败').exec_()
            print(e)
        except SystemRegisterError as e:
            QMessageBox(QMessageBox.Warning, '警告', '热键设置冲突').exec_()
            print(e)

        # 5.设置pyqt5的快捷键，ESC退出工具
        QShortcut(QKeySequence("Escape"), self, self.stop_tool)
        # 6.设置图形界面
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("疯物之诗琴")
        self.setWindowIcon(QIcon('icon.ico'))
        self.setFixedSize(QSize(360, 400))
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        # self.setAttribute(Qt.WA_TranslucentBackground)

        self.widgetLayout = QVBoxLayout()  # 创建垂直布局
        self.widgetLayout.setObjectName("widgetLayout")
        self.setLayout(self.widgetLayout)
        self.playList = QListWidget()
        self.playList.setGeometry(QRect(0, 50, 340, 60))
        self.playList.setMinimumSize(QSize(340, 60))
        self.playList.setBaseSize(QSize(340, 60))
        try:
            self.fileList = os.listdir("midi/")
            self.playList.addItems(self.fileList)
        except FileNotFoundError as e:
            QMessageBox(QMessageBox.Warning, '警告', '没有找到midi文件').exec_()
            print(e)

        self.msgLabel = QLabel('双击列表选项开始或停止演奏\nEsc退出程序，Ctrl+Shift+G停止演奏\n目前一共有%d条曲目' % (len(self.playList)))
        self.msgLabel.setGeometry(QRect(0, 0, 360, 50))
        self.msgLabel.setMinimumSize(QSize(360, 50))
        self.msgLabel.setBaseSize(QSize(360, 50))
        self.msgLabel.setAlignment(Qt.AlignLeft)
        self.msgLabel.setObjectName("msgLabel")

        self.playStatus = QLabel('请选择一首音乐开始演奏')
        self.playStatus.setGeometry(QRect(0, 130, 360, 20))
        self.playStatus.setMinimumSize(QSize(360, 20))
        self.playStatus.setBaseSize(QSize(360, 20))
        # 添加控件到布局中
        self.widgetLayout.addWidget(self.msgLabel)
        self.widgetLayout.addWidget(self.playList)
        self.widgetLayout.addWidget(self.playStatus)
        # 绑定操作函数
        self.playList.itemClicked.connect(self.play_item_clicked)
        self.playList.doubleClicked.connect(self.play_midi)
        self.playThread.playSignal.connect(self.show_stop_play)

    # 在界面显示选择的状态
    def play_item_clicked(self, item):
        print('你选择了：' + item.text())
        self.playStatus.setText('你选择了：' + item.text())

    # 热键处理函数
    def mkey_press_event(self, i_str):
        print("按下的按键是%s" % (i_str,))
        self.stop_play_thread()  # 按下全局快捷键终止演奏线程

    # 热键信号发送函数(将外部信号，转化成qt信号)
    def send_key_event(self, i_str):
        self.sig_hot_key.emit(i_str)

    # 启动playThread进行自动演奏
    def play_midi(self, index):
        self.stop_play_thread()
        print('开始演奏：' + self.fileList[index.row()])
        # 显示演奏的状态
        self.playStatus.setText('开始演奏：' + self.fileList[index.row()])
        self.playThread.set_file_path("midi/" + self.fileList[index.row()])
        self.playThread.start()
        pass

    def show_stop_play(self, msg):
        self.playStatus.setText(msg)

    # 终止演奏线程，停止自动演奏
    def stop_play_thread(self):
        self.playStatus.setText('停止演奏')  # 在工具界面显示状态
        self.playThread.stop_play()
        time.sleep(0.1)
        if not self.playThread.isFinished():
            self.playThread.terminate()
            self.playThread.wait()
        return

    # 工具退出函数，主要用来停止演奏线程和退出注销热键
    def stop_tool(self):
        self.stop_play_thread()
        time.sleep(0.1)
        try:
            self.hk_stop.unregister(('control', 'shift', 'g'))
        except UnregisterError as e:
            QMessageBox(QMessageBox.Warning, '警告', '热键注销失败').exec_()
            print(e)
        QCoreApplication.instance().quit()
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
