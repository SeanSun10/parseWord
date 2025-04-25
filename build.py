# -*- coding: utf-8 -*-
import PyInstaller.__main__

PyInstaller.__main__.run([
    'main.py',
    '--name=WordParser',
    '--onefile',
    '--windowed',
    '--add-data=word_parser.py;.',
    '--icon=app.ico',  # 如果有图标文件的话
    '--clean',
    '--noconfirm'
]) 