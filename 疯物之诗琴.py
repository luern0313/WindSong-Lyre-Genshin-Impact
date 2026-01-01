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
# 使用相对值，表示3个八度内的白键位置
# 实际 MIDI 音符 = note_lyre[i] + base_note * 12
note_lyre = [0, 2, 4, 5, 7, 9, 11,     # 低音八度 C D E F G A B (相对值)
             12, 14, 16, 17, 19, 21, 23,   # 中音八度 C D E F G A B
             24, 26, 28, 29, 31, 33, 35]   # 高音八度 C D E F G A B

# ==================== 钢琴模式 (36键，有黑键) ====================
# 实际游戏只有36键，但我们创建虚拟映射来支持更宽的音域
# 超出36键范围的音符会自动折叠到可演奏范围
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
# 使用相对值 0-35，代表 36 个半音
# 实际 MIDI 音符 = note_piano[i] + base_note * 12
note_piano = [
    # 低音区 (相对音符 0-11，12个半音)
    0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11,
    # 中音区 (相对音符 12-23，12个半音)
    12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23,
    # 高音区 (相对音符 24-35，12个半音)
    24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35
]

# 钢琴模式的实际可演奏范围 (36键)
PIANO_RANGE = 36     # 36键

# ==================== 当前使用的映射 (默认诗琴模式) ====================
key = key_lyre.copy()
vk = vk_lyre.copy()
note = note_lyre.copy()

pressed_key = set()
configure = {}

# 初始化默认 note_map (诗琴模式，base_note=3，即 C4-B6)
# 这样在 PlayThread.run() 之前也能正常使用 get_note() 函数
note_map = {note_lyre[i] + 3 * 12: key_lyre[i] for i in range(len(note_lyre))}

# 线程锁，保护共享资源的并发访问
_pressed_key_lock = threading.Lock()
_note_map_lock = threading.Lock()
_configure_lock = threading.Lock()

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
    },
    "midi_directory": {
        "set_tip": "设置MIDI文件所在目录（相对路径或绝对路径）",
        "get_tip": "MIDI文件目录",
        "default": "midi",
        "mode": "string"
    }
}

# 默认 MIDI 目录
DEFAULT_MIDI_DIR = "midi"


def get_midi_directory():
    """获取 MIDI 目录路径（线程安全）"""
    with _configure_lock:
        midi_dir = configure.get("midi_directory", DEFAULT_MIDI_DIR)
    # 确保路径以分隔符结尾
    if not midi_dir.endswith(os.sep) and not midi_dir.endswith("/"):
        midi_dir = midi_dir + os.sep
    return midi_dir


def switch_instrument_mode(mode):
    """切换乐器模式
    mode: 0 = 诗琴模式 (21键，无黑键)
          1 = 钢琴模式 (36键，有黑键)
    
    注意：使用原地修改 (clear + extend/update) 而不是重新赋值，
    这样通过 'from ... import' 导入的变量也能看到更新。
    同时更新 note_map 的默认值。
    """
    global key, vk, note, note_map
    
    if mode == 0:
        # 原地修改列表
        key.clear()
        key.extend(key_lyre)
        # 原地修改字典
        vk.clear()
        vk.update(vk_lyre)
        # 原地修改列表
        note.clear()
        note.extend(note_lyre)
        # 更新默认 note_map (base_note=3, 即 C4-B6)
        with _note_map_lock:
            note_map = {note_lyre[i] + 3 * 12: key_lyre[i] for i in range(len(note_lyre))}
        print("已切换到诗琴模式 (21键，无黑键)")
    elif mode == 1:
        # 原地修改列表
        key.clear()
        key.extend(key_piano)
        # 原地修改字典
        vk.clear()
        vk.update(vk_piano)
        # 原地修改列表
        note.clear()
        note.extend(note_piano)
        # 更新默认 note_map (base_note=3, 即 C3-B5, MIDI 36-71)
        with _note_map_lock:
            note_map = {note_piano[i] + 3 * 12: key_piano[i] for i in range(len(note_piano))}
        print("已切换到钢琴模式 (36键，有黑键)")
    
    return mode


def is_piano_mode():
    """检查当前是否为钢琴模式（线程安全）"""
    with _configure_lock:
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
    """读取配置文件（线程安全）"""
    global configure
    with _configure_lock:
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
                elif conf["mode"] == "string":
                    print(conf["get_tip"] + "：" + str(configure[conf_key]))
        print_split_line()


def save_configure():
    """保存配置文件（线程安全）"""
    with _configure_lock:
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
                elif conf["mode"] == "string":
                    print(conf["set_tip"] + f"（直接回车使用默认值：{conf['default']}）")
                    value = input("请输入：")
                    if value == "":
                        configure[conf_key] = conf["default"]
                    else:
                        configure[conf_key] = value
                    break

        except RuntimeError:
            print("ERR")


def get_base_note(bn_tracks):
    """
    智能计算最佳基础音高（base_note）
    
    对于36键钢琴模式，需要将MIDI音符映射到3个八度范围内。
    这个函数会分析MIDI文件中所有音符的分布，找到能覆盖最多音符的3个八度窗口。
    
    返回值: base_note (1-6)
        - base_note=1: 映射范围 C1-B3 (MIDI 12-47)
        - base_note=2: 映射范围 C2-B4 (MIDI 24-59)
        - base_note=3: 映射范围 C3-B5 (MIDI 36-71)  <- 标准钢琴中音区
        - base_note=4: 映射范围 C4-B6 (MIDI 48-83)
        - base_note=5: 映射范围 C5-B7 (MIDI 60-95)
        - base_note=6: 映射范围 C6-B8 (MIDI 72-107)
    """
    # 统计每个八度的音符数量 (覆盖 C0-B8，共9个八度)
    note_count = {i: 0 for i in range(9)}
    min_note = 127
    max_note = 0
    
    for bn_track in bn_tracks:
        for bn_msg in bn_track:
            if bn_msg.type == "note_on" and bn_msg.velocity > 0:
                octave = (bn_msg.note - 12) // 12  # C1=12 对应 octave=0
                if 0 <= octave < 9:
                    note_count[octave] += 1
                min_note = min(min_note, bn_msg.note)
                max_note = max(max_note, bn_msg.note)
    
    # 打印音域信息
    if min_note <= max_note:
        note_names = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
        min_name = note_names[min_note % 12] + str(min_note // 12 - 1)
        max_name = note_names[max_note % 12] + str(max_note // 12 - 1)
        print(f"MIDI文件音域: {min_name} (MIDI {min_note}) - {max_name} (MIDI {max_note})")
        print(f"跨越约 {(max_note - min_note) // 12 + 1} 个八度")
    
    # 计算每个可能的3八度窗口能覆盖多少音符
    # 窗口起始八度: 0=C1, 1=C2, 2=C3, ...
    window_scores = []
    for start_octave in range(7):  # 最多到C7开始的窗口
        score = sum(note_count[start_octave + i] for i in range(3) if start_octave + i < 9)
        window_scores.append(score)
    
    # 找到最佳窗口
    best_window = window_scores.index(max(window_scores))
    base_note = best_window + 1  # 转换为 base_note (1-7)
    
    # 打印映射信息
    start_midi = base_note * 12
    end_midi = start_midi + 35
    print(f"选择 base_note={base_note}，映射范围: MIDI {start_midi}-{end_midi}")
    
    return base_note


def get_note(n):
    """
    处理音符，包括超范围和黑键处理（线程安全）
    
    对于36键钢琴模式：
    - 超出范围的音符会被折叠到可演奏范围内（升/降八度）
    - 支持完整的12音阶，无需黑键转换
    
    对于21键诗琴模式：
    - 超出范围的音符会被折叠到可演奏范围内
    - 黑键会用邻近的白键代替
    """
    n_list = []
    
    # 获取 note_map 的快照（线程安全）
    with _note_map_lock:
        if note_map is None:
            return n_list
        note_map_keys = list(note_map.keys())
    
    if not note_map_keys:
        return n_list
    
    # 获取配置的快照（线程安全）
    with _configure_lock:
        below_limit = configure.get("below_limit", 2)
        above_limit = configure.get("above_limit", 2)
        black_key_1 = configure.get("black_key_1", 0)
        black_key_2 = configure.get("black_key_2", 3)
        black_key_3 = configure.get("black_key_3", 3)
    
    min_key = note_map_keys[0]
    max_key = note_map_keys[-1]
    
    # 处理超出范围的音符 - 折叠到可演奏范围
    while n < min_key and below_limit > 0:
        n += 12
        if below_limit == 1:
            break

    while n > max_key and above_limit > 0:
        n -= 12
        if above_limit == 1:
            break
    
    # 再次检查是否在范围内（below_limit=0或above_limit=0时可能仍超出范围）
    if n < min_key or n > max_key:
        # 音符超出范围且配置为不演奏
        return n_list

    # 钢琴模式：直接支持黑键，无需转换
    if is_piano_mode():
        n_list.append(n)
        return n_list
    
    # 诗琴模式：处理黑键（用邻近白键代替）
    # 检查音符是否需要处理（不在note_map中表示是黑键）
    if n in note_map_keys:
        n_list.append(n)
        return n_list
    
    # 确定音符所在的八度区域
    octave_size = len(note_map_keys) // 3  # 每个八度的键数 (21键模式下为7)
    if octave_size <= 0:
        octave_size = 7
    
    if note_map_keys[0] <= n <= note_map_keys[min(6, len(note_map_keys)-1)]:
        # 第一个八度
        if black_key_1 == 1:
            n -= 1
        elif black_key_1 > 1:
            n += 1
            if black_key_1 == 3:
                n_list.append(n - 2)
    elif len(note_map_keys) > 7 and note_map_keys[7] <= n <= note_map_keys[min(13, len(note_map_keys)-1)]:
        # 第二个八度
        if black_key_2 == 1:
            n -= 1
        elif black_key_2 > 1:
            n += 1
            if black_key_2 == 3:
                n_list.append(n - 2)
    elif len(note_map_keys) > 14 and note_map_keys[14] <= n <= note_map_keys[min(20, len(note_map_keys)-1)]:
        # 第三个八度
        if black_key_3 == 1:
            n -= 1
        elif black_key_3 > 1:
            n += 1
            if black_key_3 == 3:
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
        """协作式停止播放，确保资源正确释放"""
        self.playFlag = False
        self._stop_event.set()  # 触发事件，中断等待
        # 释放所有按下的键，防止按键卡住
        release_all_keys()

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
        
        # 检测调式（线程安全读取配置）
        detected_key = "C"
        with _configure_lock:
            auto_transpose = configure.get("auto_transpose", 1)
            lowest_pitch_name = configure["lowest_pitch_name"]
        
        if auto_transpose == 1:
            detected_key = detect_key_signature(tracks)
            if detected_key != "C":
                print(f"将从{detected_key}调自动转换到C调")
        
        # 获取基础音高
        base_note = get_base_note(tracks) if lowest_pitch_name == -1 else lowest_pitch_name
        
        # 创建本次播放的音符映射（局部变量，整首曲子保持不变）
        local_note_map = {note[i] + base_note * 12: key[i] for i in range(len(note))}
        
        # 同时更新全局 note_map（供其他地方参考）
        with _note_map_lock:
            note_map = local_note_map.copy()
        
        # 获取配置快照（整首曲子使用相同配置，不受中途配置变更影响）
        with _configure_lock:
            local_below_limit = configure.get("below_limit", 2)
            local_above_limit = configure.get("above_limit", 2)
        
        # 预计算映射范围（整首曲子保持不变）
        local_note_map_keys = sorted(local_note_map.keys())
        local_min_key = local_note_map_keys[0]
        local_max_key = local_note_map_keys[-1]
        
        print(f"本次演奏音符映射范围: MIDI {local_min_key} - {local_max_key}")
        
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
        
        # 定义本地音符处理函数（使用本次播放的固定映射）
        def process_note_local(n):
            """处理音符，使用本次播放的固定映射"""
            # 处理超出范围的音符 - 折叠到可演奏范围
            while n < local_min_key and local_below_limit > 0:
                n += 12
                if local_below_limit == 1:
                    break

            while n > local_max_key and local_above_limit > 0:
                n -= 12
                if local_above_limit == 1:
                    break
            
            return n
        
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
                if auto_transpose == 1 and detected_key != "C":
                    original_note = transpose_to_c(msg.note, detected_key)
                
                # 使用本地处理函数处理音符（保持整首曲子映射一致）
                processed_note = process_note_local(original_note)
                
                # 检查音符是否在映射范围内
                if processed_note not in local_note_map:
                    continue
                
                target_key = local_note_map[processed_note]
                
                if not self.playFlag:
                    self.playSignal.emit('停止演奏！！')
                    print('停止演奏！！')
                    break
                    
                if msg.type == "note_on":
                    with _pressed_key_lock:
                        key_pressed = vk[target_key] in pressed_key
                    if key_pressed:
                        release_key(vk[target_key])
                    press_key(vk[target_key])
                elif msg.type == "note_off":
                    release_key(vk[target_key])
        
        # 播放结束后，确保释放所有按键
        release_all_keys()


def press_key(hex_key_code):
    """按下指定键（线程安全）"""
    global pressed_key
    with _pressed_key_lock:
        pressed_key.add(hex_key_code)
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hex_key_code, 0x0008, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def release_key(hex_key_code):
    """释放指定键（线程安全）"""
    global pressed_key
    with _pressed_key_lock:
        pressed_key.discard(hex_key_code)
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hex_key_code, 0x0008 | 0x0002, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def release_all_keys():
    """释放所有当前按下的键，防止资源泄漏（线程安全）"""
    global pressed_key
    with _pressed_key_lock:
        # 复制一份，避免在迭代时修改集合
        keys_to_release = pressed_key.copy()
    for hex_key_code in keys_to_release:
        release_key(hex_key_code)
    with _pressed_key_lock:
        pressed_key.clear()


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
            # 使用配置的 MIDI 目录
            midi_dir = get_midi_directory()
            if not os.path.exists(midi_dir):
                print(f"MIDI目录不存在: {midi_dir}")
                print("请在配置中设置正确的MIDI目录，或创建该目录")
                break
            
            file_list = os.listdir(midi_dir)
            if not file_list:
                print(f"MIDI目录为空: {midi_dir}")
                break
            
            print("\n选择要打开的文件：")
            print("\n".join([str(i) + "、" + file_list[i] for i in range(len(file_list))]))

            midi_file = mido.MidiFile(midi_dir + file_list[int(input("请输入文件前数字序号："))])
            print_split_line()
            tracks = midi_file.tracks
            
            # 检测并显示调式（线程安全读取配置）
            detected_key = "C"
            with _configure_lock:
                auto_transpose = configure.get("auto_transpose", 1)
                lowest_pitch_name = configure["lowest_pitch_name"]
            
            if auto_transpose == 1:
                detected_key = detect_key_signature(tracks)
                if detected_key != "C":
                    print(f"检测到{detected_key}调，将自动转换到C调演奏")
                else:
                    print("检测到C调，无需转换")
            
            base_note = get_base_note(tracks) if lowest_pitch_name == -1 else lowest_pitch_name
            with _note_map_lock:
                note_map = {note[i] + base_note * 12: key[i] for i in range(len(note))}
            
            time.sleep(1)
            
            for msg in midi_file.play():
                if msg.type == "note_on" or msg.type == "note_off":
                    # 转调处理
                    original_note = msg.note
                    if auto_transpose == 1 and detected_key != "C":
                        original_note = transpose_to_c(msg.note, detected_key)
                    
                    note_list = get_note(original_note)
                    with _note_map_lock:
                        current_note_map = note_map.copy()
                    for n in note_list:
                        if n in current_note_map:
                            if msg.type == "note_on":
                                with _pressed_key_lock:
                                    key_pressed = vk[current_note_map[n]] in pressed_key
                                if key_pressed:
                                    release_key(vk[current_note_map[n]])
                                press_key(vk[current_note_map[n]])
                            elif msg.type == "note_off":
                                release_key(vk[current_note_map[n]])
                                
        except Exception as e:
            print("错误:" + str(e))


if __name__ == "__main__":
    if is_admin():
        main()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)