#!/usr/bin/env python3
# coding=utf-8
import mido
import ctypes
import win32api
import time
import json
import os
import sys
import threading
from PyQt5.QtCore import QThread, pyqtSignal

# ==================== 诗琴模式 (21键，无黑键) ====================
# 键位映射：低音(Z-M) 中音(A-J) 高音(Q-U)
key_lyre = ["z", "x", "c", "v", "b", "n", "m",
            "a", "s", "d", "f", "g", "h", "j",
            "q", "w", "e", "r", "t", "y", "u"]

vk_lyre = {"z": 0x2c, "x": 0x2d, "c": 0x2e, "v": 0x2f, "b": 0x30, "n": 0x31, "m": 0x32,
           "a": 0x1e, "s": 0x1f, "d": 0x20, "f": 0x21, "g": 0x22, "h": 0x23, "j": 0x24,
           "q": 0x10, "w": 0x11, "e": 0x12, "r": 0x13, "t": 0x14, "y": 0x15, "u": 0x16}

# 诗琴模式音符 (只有白键，C大调音阶)
note_lyre = [12, 14, 16, 17, 19, 21, 23,   # 低音 C D E F G A B
             24, 26, 28, 29, 31, 33, 35,   # 中音 C D E F G A B
             36, 38, 40, 41, 43, 45, 47]   # 高音 C D E F G A B

# ==================== 钢琴模式 (36键，有黑键) ====================
# 白键：低音(.,/IOP[) 中音(ZXCVBNM) 高音(QWERTYU)
# 黑键：低音(L;90-) 中音(SDGHJ) 高音(23567)
key_piano = [
    # 低音区 (C3-B3) - 白键 + 黑键
    ",", "l", ".", ";", "/", "i", "9", "o", "0", "p", "-", "[",
    # 中音区 (C4-B4) - 白键 + 黑键  
    "z", "s", "x", "d", "c", "v", "g", "b", "h", "n", "j", "m",
    # 高音区 (C5-B5) - 白键 + 黑键
    "q", "2", "w", "3", "e", "r", "5", "t", "6", "y", "7", "u"
]

vk_piano = {
    # 低音区白键
    ",": 0x33, ".": 0x34, "/": 0x35, "i": 0x17, "o": 0x18, "p": 0x19, "[": 0x1a,
    # 低音区黑键
    "l": 0x26, ";": 0x27, "9": 0x0a, "0": 0x0b, "-": 0x0c,
    # 中音区白键
    "z": 0x2c, "x": 0x2d, "c": 0x2e, "v": 0x2f, "b": 0x30, "n": 0x31, "m": 0x32,
    # 中音区黑键
    "s": 0x1f, "d": 0x20, "g": 0x22, "h": 0x23, "j": 0x24,
    # 高音区白键
    "q": 0x10, "w": 0x11, "e": 0x12, "r": 0x13, "t": 0x14, "y": 0x15, "u": 0x16,
    # 高音区黑键
    "2": 0x03, "3": 0x04, "5": 0x06, "6": 0x07, "7": 0x08
}

# 钢琴模式音符 (包含所有半音，完整的12音阶)
note_piano = [
    # 低音区 C3-B3 (所有12个半音)
    12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23,
    # 中音区 C4-B4 (所有12个半音)
    24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35,
    # 高音区 C5-B5 (所有12个半音)
    36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47
]

# ==================== 当前使用的映射 (默认诗琴模式) ====================
key = key_lyre.copy()
vk = vk_lyre.copy()
note = note_lyre.copy()

pressed_key = set()
note_map, configure = None, {}

# 调式识别相关
KEY_SIGNATURES = {
    # 升号调
    "C": 0,   # 无升降号
    "G": 1,   # 1个升号 F#
    "D": 2,   # 2个升号 F#, C#
    "A": 3,   # 3个升号 F#, C#, G#
    "E": 4,   # 4个升号 F#, C#, G#, D#
    "B": 5,   # 5个升号 F#, C#, G#, D#, A#
    "F#": 6,  # 6个升号
    # 降号调
    "F": -1,  # 1个降号 Bb
    "Bb": -2, # 2个降号 Bb, Eb
    "Eb": -3, # 3个降号 Bb, Eb, Ab
    "Ab": -4, # 4个降号 Bb, Eb, Ab, Db
    "Db": -5, # 5个降号
    "Gb": -6, # 6个降号
}

# 各调主音相对于C的半音数
KEY_ROOT_OFFSET = {
    "C": 0, "G": 7, "D": 2, "A": 9, "E": 4, "B": 11, "F#": 6,
    "F": 5, "Bb": 10, "Eb": 3, "Ab": 8, "Db": 1, "Gb": 6
}

configure_attr = {
    "instrument_mode": {
        "set_tip": "选择乐器模式",
        "get_tip": "乐器模式",
        "default": 0,
        "mode": "option",
        "option": [
            "诗琴模式 (21键，无黑键)",
            "钢琴模式 (36键，有黑键)"
        ]
    },
    "lowest_pitch_name": {
        "set_tip": "游戏仅支持演奏三个八度，请设定最低八度的音名（输入C后的数字，如输入4则演奏C4-B6，直接回车则程序按照不同乐谱自动判断）",
        "get_tip": "最低八度的音名",
        "default": -1,
        "mode": "int"
    },
    "auto_transpose": {
        "set_tip": "是否自动将其他调转换为C调（钢琴模式下建议关闭）",
        "get_tip": "自动转调到C调",
        "default": 1,
        "mode": "option",
        "option": [
            "不转换（只适合C调曲子）",
            "自动转换到C调"
        ]
    },
    "below_limit": {
        "set_tip": "当乐谱出现低于最低八度的音符时",
        "get_tip": "出现低于最低八度的音符",
        "default": 2,
        "mode": "option",
        "option": [
            "不演奏",
            "上升一个八度演奏",
            "升高八度直到可以演奏"
        ]
    },
    "above_limit": {
        "set_tip": "当乐谱出现高于最高八度的音符时",
        "get_tip": "出现高于最高八度的音符",
        "default": 2,
        "mode": "option",
        "option": [
            "不演奏",
            "降低一个八度演奏",
            "降低八度直到可以演奏"
        ]
    },
    "black_key_1": {
        "set_tip": "游戏不支持演奏黑键，当乐谱第一个八度（多为和弦部分）出现钢琴黑键时",
        "get_tip": "乐谱第一个八度出现钢琴黑键",
        "default": 0,
        "mode": "option",
        "option": [
            "不演奏",
            "降低一个半音演奏",
            "升高一个半音演奏",
            "同时演奏上一个半音和下一个半音"
        ]
    },
    "black_key_2": {
        "set_tip": "游戏不支持演奏黑键，当乐谱第二个八度（和弦或主旋律部分）出现钢琴黑键时",
        "get_tip": "乐谱第二个八度出现钢琴黑键",
        "default": 3,
        "mode": "option",
        "option": [
            "不演奏",
            "降低一个半音演奏",
            "升高一个半音演奏",
            "同时演奏上一个半音和下一个半音"
        ]
    },
    "black_key_3": {
        "set_tip": "游戏不支持演奏黑键，当乐谱第三个八度（多为主旋律部分）出现钢琴黑键时",
        "get_tip": "乐谱第三个八度出现钢琴黑键",
        "default": 3,
        "mode": "option",
        "option": [
            "不演奏",
            "降低一个半音演奏",
            "升高一个半音演奏",
            "同时演奏上一个半音和下一个半音"
        ]
    }
}


def switch_instrument_mode(mode):
    """切换乐器模式
    mode: 0 = 诗琴模式 (21键，无黑键)
          1 = 钢琴模式 (36键，有黑键)
    """
    global key, vk, note
    
    if mode == 0:
        key = key_lyre.copy()
        vk = vk_lyre.copy()
        note = note_lyre.copy()
        print("已切换到诗琴模式 (21键，无黑键)")
    elif mode == 1:
        key = key_piano.copy()
        vk = vk_piano.copy()
        note = note_piano.copy()
        print("已切换到钢琴模式 (36键，有黑键)")
    
    return mode


def is_piano_mode():
    """检查当前是否为钢琴模式"""
    return configure.get("instrument_mode", 0) == 1


def detect_key_signature(tracks):
    """自动检测MIDI文件的调式"""
    # 统计每个音符类别的出现次数
    note_count = {i: 0 for i in range(12)}  # 12个半音
    
    for track in tracks:
        for msg in track:
            if msg.type == 'note_on' and msg.velocity > 0:
                note_class = msg.note % 12  # 获取音符类别(0-11)
                note_count[note_class] += 1
    
    # 检查每个可能的调式
    best_key = "C"
    best_score = 0
    
    # C大调的自然音阶（相对于C的半音数）
    c_major_scale = [0, 2, 4, 5, 7, 9, 11]  # C D E F G A B
    
    for key_name, root_offset in KEY_ROOT_OFFSET.items():
        score = 0
        # 计算该调式的自然音阶
        key_scale = [(note + root_offset) % 12 for note in c_major_scale]
        
        # 计算匹配度
        for note_class in key_scale:
            score += note_count[note_class]
        
        # 特别加权主音和属音（第1和第5音）
        tonic = root_offset % 12
        dominant = (root_offset + 7) % 12
        score += note_count[tonic] * 2  # 主音权重更高
        score += note_count[dominant] * 1.5  # 属音次之
        
        if score > best_score:
            best_score = score
            best_key = key_name
    
    print(f"检测到的调式: {best_key}")
    return best_key


def transpose_to_c(midi_note, from_key):
    """将其他调的音符转换到C调"""
    if from_key == "C":
        return midi_note
    
    # 获取调式的偏移量
    offset = KEY_ROOT_OFFSET[from_key]
    
    # 转调：减去偏移量就能转到C调
    transposed = midi_note - offset
    
    return transposed


def read_configure():
    global configure
    if os.path.exists("configure.json"):
        with open("configure.json", encoding="utf-8") as file:
            configure = json.loads(file.read())
    else:
        print("配置文件不存在")
        set_configure()
        save_configure()

    # 根据配置切换乐器模式
    instrument_mode = configure.get("instrument_mode", 0)
    switch_instrument_mode(instrument_mode)

    print_split_line()
    print("当前配置：")
    for conf_key in configure.keys():
        if conf_key in configure_attr:
            conf = configure_attr[conf_key]
            if conf["mode"] == "int":
                print(conf["get_tip"] + "：" + str(configure[conf_key]))
            elif conf["mode"] == "option":
                print(conf["get_tip"] + "：" + conf["option"][configure[conf_key]])
    print_split_line()


def save_configure():
    with open("configure.json", "w", encoding="utf-8") as file:
        file.write(json.dumps(configure))
    print("配置文件已保存")


def set_configure():
    print_split_line()
    print("创建新配置：")
    for conf_key in configure_attr.keys():
        try:
            while True:
                print()
                conf = configure_attr[conf_key]
                if conf["mode"] == "int":
                    print(conf["set_tip"])
                    value = input("请输入（整数或置空）：")
                    if value.isdigit() or (value.startswith('-') and value[1:].isdigit()):
                        configure[conf_key] = int(value)
                        break
                    elif value == "":
                        configure[conf_key] = conf["default"]
                        break
                    else:
                        print("格式错误，请重新输入")
                elif conf["mode"] == "option":
                    print(conf["set_tip"] + "（直接回车则为选项" + str(conf["default"]) + "）")
                    print("\n".join([str(index) + "：" + conf["option"][index] for index in range(len(conf["option"]))]))
                    value = input("请输入（选项的序号或置空）：")
                    if value.isdigit() and int(value) < len(conf["option"]):
                        configure[conf_key] = int(value)
                        break
                    elif value == "":
                        configure[conf_key] = conf["default"]
                        break
                    else:
                        print("格式错误，请重新输入")

        except RuntimeError:
            print("ERR")


def get_base_note(bn_tracks):
    note_count = {i: 0 for i in range(9)}
    for bn_track in bn_tracks:
        for bn_msg in bn_track:
            if bn_msg.type == "note_on":
                note_count[(bn_msg.note - 24) // 12 + 1] += 1

    c = [note_count[i] + note_count[i + 1] + note_count[i + 2] for i in range(len(note_count) - 2)]
    return c.index(max(c))


def get_note(n):
    """处理音符，包括超范围和黑键处理"""
    n_list = []
    note_map_keys = list(note_map.keys())
    
    # 处理超出范围的音符
    while n < note_map_keys[0] and configure["below_limit"] > 0:
        n += 12
        if configure["below_limit"] == 1:
            break

    while n > note_map_keys[-1] and configure["above_limit"] > 0:
        n -= 12
        if configure["above_limit"] == 1:
            break

    # 钢琴模式：直接支持黑键，无需转换
    if is_piano_mode():
        n_list.append(n)
        return n_list
    
    # 诗琴模式：处理黑键（用邻近白键代替）
    if note_map_keys[0] <= n <= note_map_keys[6] and n not in note_map_keys:
        if configure["black_key_1"] == 1:
            n -= 1
        elif configure["black_key_1"] > 1:
            n += 1
            if configure["black_key_1"] == 3:
                n_list.append(n - 2)
    elif note_map_keys[7] <= n <= note_map_keys[13] and n not in note_map_keys:
        if configure["black_key_2"] == 1:
            n -= 1
        elif configure["black_key_2"] > 1:
            n += 1
            if configure["black_key_2"] == 3:
                n_list.append(n - 2)
    elif note_map_keys[14] <= n <= note_map_keys[20] and n not in note_map_keys:
        if configure["black_key_3"] == 1:
            n -= 1
        elif configure["black_key_3"] > 1:
            n += 1
            if configure["black_key_3"] == 3:
                n_list.append(n - 2)

    n_list.append(n)
    return n_list


def print_split_line():
    print("_" * 50)


# Windows键盘控制相关
PUL = ctypes.POINTER(ctypes.c_ulong)
SendInput = ctypes.windll.user32.SendInput


class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]


class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                ("mi", MouseInput),
                ("hi", HardwareInput)]


class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]


class PlayThread(QThread):
    """演奏线程"""
    playSignal = pyqtSignal(str)
    progressSignal = pyqtSignal(float)  # 添加进度信号
    file_path = None
    start_time = 0  # 添加起始时间属性

    def __init__(self, parent=None):
        super(PlayThread, self).__init__(parent)
        self.playFlag = False
        self._stop_event = threading.Event()  # 使用Event实现可中断等待
        read_configure()

    def stop_play(self):
        self.playFlag = False
        self._stop_event.set()  # 触发事件，中断等待

    def set_file_path(self, file_path):
        self.file_path = file_path
    
    def set_start_time(self, start_time):
        self.start_time = start_time

    def _interruptible_sleep(self, seconds):
        """可中断的睡眠，返回True表示正常完成，False表示被中断"""
        if seconds <= 0:
            return self.playFlag
        # 使用Event.wait实现可中断等待
        # wait返回True表示事件被set（即被中断），False表示超时
        interrupted = self._stop_event.wait(timeout=seconds)
        return not interrupted and self.playFlag

    def run(self):
        self.playFlag = True
        self._stop_event.clear()  # 重置停止事件
        global note_map
        
        midi = mido.MidiFile(self.file_path)
        print_split_line()
        tracks = midi.tracks
        
        # 检测调式
        detected_key = "C"
        if configure.get("auto_transpose", 1) == 1:
            detected_key = detect_key_signature(tracks)
            if detected_key != "C":
                print(f"将从{detected_key}调自动转换到C调")
        
        # 获取基础音高
        base_note = get_base_note(tracks) if configure["lowest_pitch_name"] == -1 else configure["lowest_pitch_name"]
        
        # 创建C调的音符映射
        note_map = {note[i] + base_note * 12: key[i] for i in range(len(note))}
        
        # 使用可中断的等待
        if not self._interruptible_sleep(1):
            self.playSignal.emit('停止演奏！')
            return
        
        # 如果设置了起始时间，跳过前面的消息
        elapsed_time = 0
        messages_to_play = []
        skip_time = self.start_time
        
        for msg in midi:
            # 累加时间
            if elapsed_time < skip_time:
                elapsed_time += msg.time
                if elapsed_time >= skip_time:
                    # 调整第一个消息的时间
                    adjusted_msg = msg.copy()
                    adjusted_msg.time = elapsed_time - skip_time
                    messages_to_play.append(adjusted_msg)
                # 跳过这个消息
            else:
                messages_to_play.append(msg)
        
        # 如果没有设置起始时间，使用所有消息
        if self.start_time == 0:
            messages_to_play = list(midi)
        
        # 记录开始播放的时间
        play_start_time = time.time() - self.start_time
        
        # 播放消息
        for msg in messages_to_play:
            if not self.playFlag:
                self.playSignal.emit('停止演奏！')
                print('停止演奏！')
                break
            
            # 使用可中断的等待替代 time.sleep
            if msg.time > 0:
                if not self._interruptible_sleep(msg.time):
                    self.playSignal.emit('停止演奏！')
                    print('停止演奏！')
                    break
            
            # 发送当前播放进度
            current_play_time = time.time() - play_start_time
            self.progressSignal.emit(current_play_time)
                
            if msg.type == "note_on" or msg.type == "note_off":
                # 如果需要转调，先转换音符
                original_note = msg.note
                if configure.get("auto_transpose", 1) == 1 and detected_key != "C":
                    original_note = transpose_to_c(msg.note, detected_key)
                
                # 获取要按的键
                note_list = get_note(original_note)
                
                for n in note_list:
                    if not self.playFlag:
                        self.playSignal.emit('停止演奏！！')
                        print('停止演奏！！')
                        break
                    if n in note_map:
                        if msg.type == "note_on":
                            if vk[note_map[n]] in pressed_key:
                                release_key(vk[note_map[n]])
                            press_key(vk[note_map[n]])
                        elif msg.type == "note_off":
                            release_key(vk[note_map[n]])


def press_key(hex_key_code):
    global pressed_key
    pressed_key.add(hex_key_code)
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hex_key_code, 0x0008, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def release_key(hex_key_code):
    global pressed_key
    pressed_key.discard(hex_key_code)
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hex_key_code, 0x0008 | 0x0002, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except RuntimeError:
        return False


def main():
    global note_map
    print("疯物之诗琴 - 自动转调版")
    print("世界线变动率：1.2.0.61745723")
    print("新功能：自动将其他调转换为C调")
    read_configure()
    
    while True:
        try:
            file_list = os.listdir("midi/")
            print("\n选择要打开的文件：")
            print("\n".join([str(i) + "、" + file_list[i] for i in range(len(file_list))]))

            midi_file = mido.MidiFile("midi/" + file_list[int(input("请输入文件前数字序号："))])
            print_split_line()
            tracks = midi_file.tracks
            
            # 检测并显示调式
            detected_key = "C"
            if configure.get("auto_transpose", 1) == 1:
                detected_key = detect_key_signature(tracks)
                if detected_key != "C":
                    print(f"检测到{detected_key}调，将自动转换到C调演奏")
                else:
                    print("检测到C调，无需转换")
            
            base_note = get_base_note(tracks) if configure["lowest_pitch_name"] == -1 else configure["lowest_pitch_name"]
            note_map = {note[i] + base_note * 12: key[i] for i in range(len(note))}
            
            time.sleep(1)
            
            for msg in midi_file.play():
                if msg.type == "note_on" or msg.type == "note_off":
                    # 转调处理
                    original_note = msg.note
                    if configure.get("auto_transpose", 1) == 1 and detected_key != "C":
                        original_note = transpose_to_c(msg.note, detected_key)
                    
                    note_list = get_note(original_note)
                    for n in note_list:
                        if n in note_map:
                            if msg.type == "note_on":
                                if vk[note_map[n]] in pressed_key:
                                    release_key(vk[note_map[n]])
                                press_key(vk[note_map[n]])
                            elif msg.type == "note_off":
                                release_key(vk[note_map[n]])
                                
        except Exception as e:
            print("错误:" + str(e))


if __name__ == "__main__":
    if is_admin():
        main()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)