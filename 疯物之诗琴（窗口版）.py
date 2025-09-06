import ctypes
import os
import sys
import time

from PyQt5.QtCore import QSize, Qt, QRect, pyqtSignal, QCoreApplication, QFileSystemWatcher
from PyQt5.QtGui import QKeySequence, QIcon
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QListWidget, QApplication, QShortcut, QMessageBox, QLineEdit, QHBoxLayout
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
        # 存储原始文件列表
        self.allFileList = []
        self.fileList = []
        # 创建文件系统监控器
        self.file_watcher = QFileSystemWatcher()
        self.midi_path = "midi/"
        if os.path.exists(self.midi_path):
            self.file_watcher.addPath(self.midi_path)
            self.file_watcher.directoryChanged.connect(self.on_directory_changed)
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
        self.setFixedSize(QSize(360, 440))
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        # self.setAttribute(Qt.WA_TranslucentBackground)

        self.widgetLayout = QVBoxLayout()  # 创建垂直布局
        self.widgetLayout.setObjectName("widgetLayout")
        self.setLayout(self.widgetLayout)
        
        # 添加搜索框
        self.searchLayout = QHBoxLayout()
        self.searchLabel = QLabel('搜索：')
        self.searchInput = QLineEdit()
        self.searchInput.setPlaceholderText('输入歌曲名称搜索...')
        self.searchInput.textChanged.connect(self.on_search_text_changed)
        self.searchLayout.addWidget(self.searchLabel)
        self.searchLayout.addWidget(self.searchInput)
        self.playList = QListWidget()
        self.playList.setGeometry(QRect(0, 50, 340, 60))
        self.playList.setMinimumSize(QSize(340, 60))
        self.playList.setBaseSize(QSize(340, 60))
        # 初始加载文件列表
        self.reload_file_list()

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
        self.widgetLayout.addLayout(self.searchLayout)
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

    # 重新加载文件列表
    def reload_file_list(self):
        try:
            # 获取midi文件夹中的所有文件
            all_files = os.listdir(self.midi_path)
            # 只保留midi和mid文件
            self.allFileList = [f for f in all_files if f.lower().endswith(('.mid', '.midi'))]
            # 应用当前的搜索过滤
            self.apply_search_filter()
        except FileNotFoundError as e:
            QMessageBox(QMessageBox.Warning, '警告', '没有找到midi文件夹').exec_()
            print(e)
            self.allFileList = []
            self.fileList = []
    
    # 文件夹变化时的处理函数
    def on_directory_changed(self, path):
        print(f'检测到文件夹变化: {path}')
        # 保存当前正在播放的文件（如果有）
        current_playing = None
        if self.playThread.isRunning():
            current_row = self.playList.currentRow()
            if current_row >= 0 and current_row < len(self.fileList):
                current_playing = self.fileList[current_row]
        
        # 重新加载文件列表
        self.reload_file_list()
        
        # 如果之前有正在播放的文件，尝试重新选中它
        if current_playing and current_playing in self.fileList:
            index = self.fileList.index(current_playing)
            self.playList.setCurrentRow(index)
    
    # 应用搜索过滤
    def apply_search_filter(self):
        search_text = self.searchInput.text() if hasattr(self, 'searchInput') else ''
        
        if search_text:
            # 过滤文件列表
            self.fileList = [f for f in self.allFileList if search_text.lower() in f.lower()]
        else:
            # 如果搜索框为空，显示所有文件
            self.fileList = self.allFileList.copy()
        
        # 更新列表显示
        self.playList.clear()
        self.playList.addItems(self.fileList)
        
        # 更新消息标签
        if hasattr(self, 'msgLabel'):
            if search_text:
                self.msgLabel.setText('双击列表选项开始或停止演奏\nEsc退出程序，Ctrl+Shift+G停止演奏\n搜索到%d条曲目（共%d条）' % (len(self.fileList), len(self.allFileList)))
            else:
                self.msgLabel.setText('双击列表选项开始或停止演奏\nEsc退出程序，Ctrl+Shift+G停止演奏\n目前一共有%d条曲目' % len(self.fileList))
    
    # 搜索过滤功能
    def on_search_text_changed(self, text):
        self.apply_search_filter()
    
    # 工具退出函数，主要用来停止演奏线程和退出注销热键
    def stop_tool(self):
        self.stop_play_thread()
        time.sleep(0.1)
        # 移除文件系统监控
        if self.file_watcher and self.midi_path:
            self.file_watcher.removePath(self.midi_path)
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
