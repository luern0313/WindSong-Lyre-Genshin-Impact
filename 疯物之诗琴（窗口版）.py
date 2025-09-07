import ctypes
import os
import sys
import time
import mido

from PyQt5.QtCore import QSize, Qt, QRect, pyqtSignal, QCoreApplication, QFileSystemWatcher, QTimer
from PyQt5.QtGui import QKeySequence, QIcon
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QListWidget, QApplication, QShortcut, QMessageBox, QLineEdit, QHBoxLayout, QSlider, QPushButton
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

    def setup_ui(self):
        self.setWindowTitle("疯物之诗琴")
        self.setWindowIcon(QIcon('icon.ico'))
        self.setFixedSize(QSize(400, 500))
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

        # 播放控制按钮
        self.controlLayout = QHBoxLayout()
        self.playPauseButton = QPushButton('▶ 播放')
        self.stopButton = QPushButton('■ 停止')
        self.playPauseButton.clicked.connect(self.on_play_pause_button_clicked)
        self.stopButton.clicked.connect(self.on_stop_button_clicked)
        self.controlLayout.addWidget(self.playPauseButton)
        self.controlLayout.addWidget(self.stopButton)
        
        # 进度条和时间显示
        self.progressLayout = QVBoxLayout()
        
        # 时间标签
        self.timeLayout = QHBoxLayout()
        self.currentTimeLabel = QLabel('00:00')
        self.totalTimeLabel = QLabel('00:00')
        self.timeLayout.addWidget(self.currentTimeLabel)
        self.timeLayout.addStretch()
        self.timeLayout.addWidget(self.totalTimeLabel)
        
        # 进度条
        self.progressSlider = QSlider(Qt.Horizontal)
        self.progressSlider.setMinimum(0)
        self.progressSlider.setMaximum(1000)  # 使用1000作为精度
        self.progressSlider.setValue(0)
        self.progressSlider.sliderPressed.connect(self.on_slider_pressed)
        self.progressSlider.sliderReleased.connect(self.on_slider_released)
        self.progressSlider.sliderMoved.connect(self.on_slider_moved)
        
        self.progressLayout.addLayout(self.timeLayout)
        self.progressLayout.addWidget(self.progressSlider)
        
        self.playStatus = QLabel('请选择一首音乐开始演奏')
        self.playStatus.setGeometry(QRect(0, 130, 400, 20))
        self.playStatus.setMinimumSize(QSize(400, 20))
        self.playStatus.setBaseSize(QSize(400, 20))
        
        # 添加控件到布局中
        self.widgetLayout.addWidget(self.msgLabel)
        self.widgetLayout.addLayout(self.searchLayout)
        self.widgetLayout.addWidget(self.playList)
        self.widgetLayout.addLayout(self.controlLayout)
        self.widgetLayout.addLayout(self.progressLayout)
        self.widgetLayout.addWidget(self.playStatus)
        # 绑定操作函数
        self.playList.itemClicked.connect(self.play_item_clicked)
        self.playList.doubleClicked.connect(self.on_list_double_clicked)
        self.playThread.playSignal.connect(self.show_stop_play)
        self.playThread.progressSignal.connect(self.on_play_progress)

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
            self.playStatus.setText('暂停中')
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
            self.playStatus.setText(f'演奏中：{file_name} (从{start_time:.1f}秒开始)')
        else:
            self.playStatus.setText('演奏中：' + file_name)
            
        self.playThread.set_file_path(self.current_midi_file)
        self.playThread.set_start_time(start_time)
        self.current_time = start_time
        self.is_paused = False
        self.playThread.start()
        self.progress_timer.start()
        self.playPauseButton.setText('⏸ 暂停')

    def show_stop_play(self, msg):
        self.playStatus.setText(msg)

    # 终止演奏线程，停止自动演奏
    def stop_play_thread(self):
        if not self.is_paused:  # 只有非暂停状态才显示停止
            self.playStatus.setText('停止演奏')
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
