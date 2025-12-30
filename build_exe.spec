# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['疯物之诗琴（窗口版）.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('styles', 'styles'),          # 包含样式文件夹
        ('icon.ico', '.'),             # 包含图标
        # ('midi', 'midi'),            # 如果想内置midi文件，取消此行注释
    ],
    hiddenimports=[
        'mido',
        'mido.midifiles',
        'mido.messages',
        'system_hotkey',
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'PyQt5.sip',
        'qtawesome',
        'win32con',
        'win32api',
        'win32gui',
    ],
    excludes=['rtmidi', 'mido.backends.rtmidi'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='疯物之诗琴',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                          # 使用UPX压缩（需要安装UPX）
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                     # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',                   # 应用图标
    uac_admin=True,                    # 请求管理员权限
)
