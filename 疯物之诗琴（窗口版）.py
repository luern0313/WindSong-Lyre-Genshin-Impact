import ctypes
import os
import sys
import time
import mido
import qtawesome as qta

# 修复 system_hotkey 与 pywin32 的兼容性问题
import win32con
if not hasattr(win32con, 'VK_MEDIA_STOP'):
    win32con.VK_MEDIA_STOP = 0xB2
if not hasattr(win32con, 'VK_MEDIA_PLAY_PAUSE'):
    win32con.VK_MEDIA_PLAY_PAUSE = 0xB3
if not hasattr(win32con, 'VK_MEDIA_PREV_TRACK'):
    win32con.VK_MEDIA_PREV_TRACK = 0xB1
if not hasattr(win32con, 'VK_MEDIA_NEXT_TRACK'):
    win32con.VK_MEDIA_NEXT_TRACK = 0xB0

from PyQt5.QtCore import QSize, Qt, QRect, pyqtSignal, QCoreApplication, QFileSystemWatcher, QTimer
from PyQt5.QtGui import QKeySequence, QIcon, QFont, QFontDatabase
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QLabel, QListWidget, QApplication, 
                             QShortcut, QMessageBox, QLineEdit, QHBoxLayout, QSlider, 
                             QPushButton, QFrame, QGraphicsDropShadowEffect)
from system_hotkey import SystemHotkey, SystemRegisterError, InvalidKeyError, UnregisterError

from 疯物之诗琴 import PlayThread, is_admin

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']


def load_stylesheet():
    """加载QSS样式表"""
    style_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'styles', 'theme.qss')
    if os.path.exists(style_path):
        with open(style_path, 'r', encoding='utf-8') as f:
            return f.read()
    return ""


class playWindow(QWidget):
    sig_hot_key = pyqtSignal(str)

    def __init__(self, parent=None):
        super(playWindow, self).__init__(parent)
        # 创建自动演奏线程
        self.playThread = PlayThread()
        # 存储原始文件列表
        self.allFileList = []
        self.fileList = []
        # 当前播放的文件和总时长
        self.current_midi_file = None
        self.total_duration = 0
        self.current_time = 0
        self.is_dragging = False
        self.is_paused = False  # 添加暂停状态
        self.pause_time = 0  # 记录暂停时的时间
        # 创建文件系统监控器
        self.file_watcher = QFileSystemWatcher()
        self.midi_path = "midi/"
        if os.path.exists(self.midi_path):
            self.file_watcher.addPath(self.midi_path)
            self.file_watcher.directoryChanged.connect(self.on_directory_changed)
        # 创建定时器更新进度
        self.progress_timer = QTimer()
        self.progress_timer.timeout.connect(self.update_progress)
        self.progress_timer.setInterval(100)  # 每100ms更新一次
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

    def setup_custom_title_bar(self):
        self.titleBar = QWidget()
        self.titleBar.setObjectName("titleBar")
        self.titleBar.setFixedHeight(40)
        
        layout = QHBoxLayout(self.titleBar)
        layout.setContentsMargins(15, 0, 10, 0)
        layout.setSpacing(10)
        
        # 图标
        iconLabel = QLabel()
        iconLabel.setPixmap(QIcon('icon.ico').pixmap(20, 20))
        layout.addWidget(iconLabel)
        
        # 标题
        titleLabel = QLabel("疯物之诗琴")
        titleLabel.setObjectName("windowTitle")
        layout.addWidget(titleLabel)
        
        layout.addStretch()
        
        # 最小化按钮
        self.btnMin = QPushButton()
        self.btnMin.setObjectName("btnMin")
        self.btnMin.setIcon(qta.icon('fa5s.minus', color='#5c5c5c'))
        self.btnMin.setFixedSize(30, 30)
        self.btnMin.clicked.connect(self.showMinimized)
        layout.addWidget(self.btnMin)
        
        # 关闭按钮
        self.btnClose = QPushButton()
        self.btnClose.setObjectName("btnClose")
        self.btnClose.setIcon(qta.icon('fa5s.times', color='#5c5c5c'))
        self.btnClose.setFixedSize(30, 30)
        self.btnClose.clicked.connect(self.close)
        layout.addWidget(self.btnClose)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if self.titleBar.geometry().contains(event.pos()):
                self.is_dragging = True
                self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
                event.accept()

    def mouseMoveEvent(self, event):
        if self.is_dragging and event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.is_dragging = False

    def setup_ui(self):
        self.setWindowTitle("疯物之诗琴")
        self.setWindowIcon(QIcon('icon.ico'))
        self.setFixedSize(QSize(960, 540))  # 16:9 宽屏
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setObjectName("mainWindow")
        
        # 根布局 - 垂直布局 (标题栏 + 内容区)
        self.rootLayout = QVBoxLayout()
        self.rootLayout.setContentsMargins(0, 0, 0, 0)
        self.rootLayout.setSpacing(0)
        self.setLayout(self.rootLayout)
        
        # 自定义标题栏
        self.setup_custom_title_bar()
        self.rootLayout.addWidget(self.titleBar)
        
        # 内容区域容器
        self.contentWidget = QWidget()
        self.contentWidget.setObjectName("contentWidget")
        self.rootLayout.addWidget(self.contentWidget)
        
        # 主布局 - 水平布局 (在内容区域内)
        self.mainLayout = QHBoxLayout(self.contentWidget)
        self.mainLayout.setContentsMargins(20, 20, 20, 20)
        self.mainLayout.setSpacing(20)
        
        # ========== 左侧区域 (列表 + 搜索) ==========
        self.leftWidget = QWidget()
        self.leftWidget.setObjectName("leftWidget")
        self.leftLayout = QVBoxLayout(self.leftWidget)
        self.leftLayout.setContentsMargins(0, 0, 0, 0)
        self.leftLayout.setSpacing(10)
        
        # 搜索框
        self.searchLayout = QHBoxLayout()
        self.searchLabel = QLabel()
        self.searchLabel.setPixmap(qta.icon('fa5s.search', color='#4A90D9').pixmap(16, 16))
        self.searchInput = QLineEdit()
        self.searchInput.setPlaceholderText('搜索曲目...')
        self.searchInput.textChanged.connect(self.on_search_text_changed)
        self.searchInput.setMinimumHeight(40)
        self.searchLayout.addWidget(self.searchLabel)
        self.searchLayout.addWidget(self.searchInput)
        
        # 播放列表
        self.playList = QListWidget()
        self.reload_file_list()
        
        self.leftLayout.addLayout(self.searchLayout)
        self.leftLayout.addWidget(self.playList)
        
        # ========== 右侧区域 (控制 + 信息) ==========
        self.rightWidget = QWidget()
        self.rightWidget.setObjectName("rightWidget")
        self.rightLayout = QVBoxLayout(self.rightWidget)
        self.rightLayout.setContentsMargins(0, 0, 0, 0)
        self.rightLayout.setSpacing(20)
        
        # 标题
        self.titleLabel = QLabel('🎵 疯物之诗琴')
        self.titleLabel.setObjectName("titleLabel")
        self.titleLabel.setAlignment(Qt.AlignCenter)
        
        # 提示信息
        self.msgLabel = QLabel('🎹 双击列表选项开始演奏\nEsc 退出程序 | Ctrl+Shift+G 停止演奏')
        self.msgLabel.setObjectName("msgLabel")
        self.msgLabel.setAlignment(Qt.AlignCenter)
        self.msgLabel.setWordWrap(True)
        
        # 控制面板容器
        self.controlFrame = QFrame()
        self.controlFrame.setObjectName("controlFrame")
        self.controlFrameLayout = QVBoxLayout(self.controlFrame)
        self.controlFrameLayout.setContentsMargins(20, 20, 20, 20)
        self.controlFrameLayout.setSpacing(15)
        
        # 按钮行
        self.controlLayout = QHBoxLayout()
        self.controlLayout.setSpacing(20)
        
        self.playPauseButton = QPushButton(' 播放')
        self.playPauseButton.setObjectName("playPauseButton")
        self.playPauseButton.setIcon(qta.icon('fa5s.play', color='white'))
        self.playPauseButton.setIconSize(QSize(16, 16))
        self.playPauseButton.setMinimumHeight(45)
        self.playPauseButton.setCursor(Qt.PointingHandCursor)
        
        self.stopButton = QPushButton(' 停止')
        self.stopButton.setObjectName("stopButton")
        self.stopButton.setIcon(qta.icon('fa5s.stop', color='white'))
        self.stopButton.setIconSize(QSize(16, 16))
        self.stopButton.setMinimumHeight(45)
        self.stopButton.setCursor(Qt.PointingHandCursor)
        
        self.playPauseButton.clicked.connect(self.on_play_pause_button_clicked)
        self.stopButton.clicked.connect(self.on_stop_button_clicked)
        
        self.controlLayout.addWidget(self.playPauseButton)
        self.controlLayout.addWidget(self.stopButton)
        
        # 进度条区域
        self.progressLayout = QVBoxLayout()
        self.progressLayout.setSpacing(8)
        
        self.timeLayout = QHBoxLayout()
        self.currentTimeLabel = QLabel('00:00')
        self.currentTimeLabel.setObjectName("timeLabel")
        self.totalTimeLabel = QLabel('00:00')
        self.totalTimeLabel.setObjectName("timeLabel")
        self.timeLayout.addWidget(self.currentTimeLabel)
        self.timeLayout.addStretch()
        self.timeLayout.addWidget(self.totalTimeLabel)
        
        self.progressSlider = QSlider(Qt.Horizontal)
        self.progressSlider.setMinimum(0)
        self.progressSlider.setMaximum(1000)
        self.progressSlider.setValue(0)
        self.progressSlider.setMinimumHeight(25)
        self.progressSlider.setCursor(Qt.PointingHandCursor)
        self.progressSlider.sliderPressed.connect(self.on_slider_pressed)
        self.progressSlider.sliderReleased.connect(self.on_slider_released)
        self.progressSlider.sliderMoved.connect(self.on_slider_moved)
        
        self.progressLayout.addLayout(self.timeLayout)
        self.progressLayout.addWidget(self.progressSlider)
        
        self.controlFrameLayout.addLayout(self.controlLayout)
        self.controlFrameLayout.addLayout(self.progressLayout)
        
        # 状态栏
        self.playStatus = QLabel('✨ 请选择一首音乐开始演奏')
        self.playStatus.setObjectName("playStatus")
        self.playStatus.setAlignment(Qt.AlignCenter)
        self.playStatus.setMinimumHeight(40)
        self.playStatus.setWordWrap(True)
        
        # 添加到右侧布局
        self.rightLayout.addStretch()
        self.rightLayout.addWidget(self.titleLabel)
        self.rightLayout.addWidget(self.msgLabel)
        self.rightLayout.addStretch()
        self.rightLayout.addWidget(self.controlFrame)
        self.rightLayout.addStretch()
        self.rightLayout.addWidget(self.playStatus)
        self.rightLayout.addStretch()
        
        # 添加到主布局
        self.mainLayout.addWidget(self.leftWidget, 4) # 左侧占 40%
        self.mainLayout.addWidget(self.rightWidget, 6) # 右侧占 60%
        
        # 绑定操作函数
        self.playList.itemClicked.connect(self.play_item_clicked)
        self.playList.doubleClicked.connect(self.on_list_double_clicked)
        self.playThread.playSignal.connect(self.show_stop_play)
        self.playThread.progressSignal.connect(self.on_play_progress)

    # 在界面显示选择的状态
    def play_item_clicked(self, item):
        print('你选择了：' + item.text())
        self.playStatus.setText('✨ 已选择：' + item.text())

    # 热键处理函数
    def mkey_press_event(self, i_str):
        print("按下的按键是%s" % (i_str,))
        self.stop_play_thread()  # 按下全局快捷键终止演奏线程

    # 热键信号发送函数(将外部信号，转化成qt信号)
    def send_key_event(self, i_str):
        self.sig_hot_key.emit(i_str)

    # 双击列表项
    def on_list_double_clicked(self, index):
        selected_file = self.fileList[index.row()]
        self.current_midi_file = "midi/" + selected_file
        self.is_paused = False
        self.pause_time = 0
        self.play_midi_from_position(0)
    
    # 播放/暂停按钮点击
    def on_play_pause_button_clicked(self):
        if self.playThread.isRunning():
            # 当前正在播放，执行暂停
            self.pause_play()
        else:
            # 当前未播放，开始播放
            if self.is_paused and self.current_midi_file:
                # 从暂停位置继续播放
                self.resume_play()
            elif self.playList.currentRow() >= 0:
                # 新开始播放
                selected_file = self.fileList[self.playList.currentRow()]
                self.current_midi_file = "midi/" + selected_file
                self.is_paused = False
                self.play_midi_from_position(0)
            else:
                QMessageBox(QMessageBox.Warning, '提示', '请先选择一首歌曲').exec_()
    
    # 停止按钮点击
    def on_stop_button_clicked(self):
        self.is_paused = False
        self.pause_time = 0
        self.current_time = 0
        self.stop_play_thread()
        self.progressSlider.setValue(0)
        self.currentTimeLabel.setText('00:00')
    
    # 暂停播放
    def pause_play(self):
        if self.playThread.isRunning():
            self.is_paused = True
            self.pause_time = self.current_time
            self.stop_play_thread()
            self.playPauseButton.setText('▶ 继续')
            self.playStatus.setText('⏸️ 已暂停')
            # 停止进度更新
            self.progress_timer.stop()
    
    # 继续播放
    def resume_play(self):
        if self.is_paused and self.current_midi_file:
            self.is_paused = False
            self.play_midi_from_position(self.pause_time)
    
    # 从指定位置播放
    def play_midi_from_position(self, start_time):
        # 如果正在播放，先停止
        if self.playThread.isRunning():
            self.playThread.stop_play()
            self.progress_timer.stop()
            time.sleep(0.2)  # 等待线程完全停止
            if not self.playThread.isFinished():
                self.playThread.terminate()
                self.playThread.wait()
        
        if not self.current_midi_file:
            return
            
        # 获取MIDI文件总时长
        try:
            midi = mido.MidiFile(self.current_midi_file)
            self.total_duration = midi.length
            self.totalTimeLabel.setText(self.format_time(self.total_duration))
            self.progressSlider.setMaximum(int(self.total_duration * 10))  # 0.1秒精度
        except:
            self.total_duration = 300  # 默认5分钟
            
        file_name = os.path.basename(self.current_midi_file)
        print(f'开始演奏：{file_name}，从第{start_time:.1f}秒开始')
        
        # 显示演奏的状态
        if start_time > 0:
            self.playStatus.setText(f'🎵 演奏中：{file_name} (从{start_time:.1f}秒)')
        else:
            self.playStatus.setText('🎵 演奏中：' + file_name)
            
        self.playThread.set_file_path(self.current_midi_file)
        self.playThread.set_start_time(start_time)
        self.current_time = start_time
        self.is_paused = False
        self.playThread.start()
        self.progress_timer.start()
        self.playPauseButton.setText('⏸ 暂停')

    def show_stop_play(self, msg):
        self.playStatus.setText('✅ ' + msg)

    # 终止演奏线程，停止自动演奏
    def stop_play_thread(self):
        if not self.is_paused:  # 只有非暂停状态才显示停止
            self.playStatus.setText('⏹️ 已停止演奏')
        self.playThread.stop_play()
        self.progress_timer.stop()  # 停止进度更新
        self.playPauseButton.setText('▶ 播放')
        time.sleep(0.1)
        if not self.playThread.isFinished():
            self.playThread.terminate()
            self.playThread.wait()
        return
    
    # 格式化时间显示
    def format_time(self, seconds):
        minutes = int(seconds // 60)
        secs = int(seconds % 60)
        return f'{minutes:02d}:{secs:02d}'
    
    # 更新播放进度
    def update_progress(self):
        if not self.is_dragging and self.playThread.isRunning() and not self.is_paused:
            self.current_time += 0.1
            if self.current_time > self.total_duration:
                self.current_time = self.total_duration
                self.on_stop_button_clicked()  # 播放完成，停止
            self.currentTimeLabel.setText(self.format_time(self.current_time))
            self.progressSlider.setValue(int(self.current_time * 10))
    
    # 播放进度信号处理
    def on_play_progress(self, current_time):
        self.current_time = current_time
        if not self.is_dragging:
            self.currentTimeLabel.setText(self.format_time(current_time))
            self.progressSlider.setValue(int(current_time * 10))
    
    # 进度条按下
    def on_slider_pressed(self):
        self.is_dragging = True
    
    # 进度条释放
    def on_slider_released(self):
        self.is_dragging = False
        # 跳转到新位置播放
        new_time = self.progressSlider.value() / 10.0
        self.current_time = new_time
        if self.playThread.isRunning():
            self.play_midi_from_position(new_time)
    
    # 进度条移动
    def on_slider_moved(self, value):
        if self.is_dragging:
            time_pos = value / 10.0
            self.currentTimeLabel.setText(self.format_time(time_pos))

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
                self.msgLabel.setText('🎹 双击列表选项开始演奏\nEsc 退出程序 | Ctrl+Shift+G 停止演奏\n🔍 搜索到 %d 条曲目（共 %d 条）' % (len(self.fileList), len(self.allFileList)))
            else:
                self.msgLabel.setText('🎹 双击列表选项开始演奏\nEsc 退出程序 | Ctrl+Shift+G 停止演奏\n📂 共 %d 条曲目' % len(self.fileList))
    
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
    
    # 加载样式表
    stylesheet = load_stylesheet()
    if stylesheet:
        app.setStyleSheet(stylesheet)
    
    mainWindow = playWindow()
    mainWindow.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    if is_admin():
        main()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
