#!/usr/bin/env python3
# coding=utf-8
import mido
import ctypes
import win32api
import time
import json
import os
import sys
from PyQt5.QtCore import QThread, pyqtSignal

# 游戏键盘映射
key = ["z", "x", "c", "v", "b", "n", "m",
       "a", "s", "d", "f", "g", "h", "j",
       "q", "w", "e", "r", "t", "y", "u"]

vk = {"z": 0x2c, "x": 0x2d, "c": 0x2e, "v": 0x2f, "b": 0x30, "n": 0x31, "m": 0x32,
      "a": 0x1e, "s": 0x1f, "d": 0x20, "f": 0x21, "g": 0x22, "h": 0x23, "j": 0x24,
      "q": 0x10, "w": 0x11, "e": 0x12, "r": 0x13, "t": 0x14, "y": 0x15, "u": 0x16}

# C大调音阶的MIDI值（相对值）
note = [12, 14, 16, 17, 19, 21, 23,
        24, 26, 28, 29, 31, 33, 35,
        36, 38, 40, 41, 43, 45, 47]

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
    "lowest_pitch_name": {
        "set_tip": "游戏仅支持演奏三个八度，请设定最低八度的音名（输入C后的数字，如输入4则演奏C4-B6，直接回车则程序按照不同乐谱自动判断）",
        "get_tip": "最低八度的音名",
        "default": -1,
        "mode": "int"
    },
    "auto_transpose": {
        "set_tip": "是否自动将其他调转换为C调",
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

    # 处理黑键
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
    file_path = None

    def __init__(self, parent=None):
        super(PlayThread, self).__init__(parent)
        self.playFlag = False
        read_configure()

    def stop_play(self):
        self.playFlag = False

    def set_file_path(self, file_path):
        self.file_path = file_path

    def run(self):
        self.playFlag = True
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
        
        time.sleep(1)
        
        for msg in midi.play():
            if not self.playFlag:
                self.playSignal.emit('停止演奏！')
                print('停止演奏！')
                break
                
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