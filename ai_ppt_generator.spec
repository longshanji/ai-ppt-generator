# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['ai_ppt_generator.py'],
    pathex=[],
    binaries=[],
    datas=[('config.ini', '.')],  # 添加配置文件
    hiddenimports=['pptx', 'pptx.dml.color', 'pptx.opc.constants', 'pptx.opc.package'],  # 添加 python-pptx 相关依赖
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AI PPT生成器',  # 修改可执行文件名称
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 设置为False以隐藏控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # 如果有图标文件，请提供正确的路径
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ai_ppt_generator',
)
