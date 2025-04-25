# Word题库解析工具

这是一个用于解析Word文档中题目的工具，可以将题目、选项、答案和解析等内容提取并导出到Excel文件中。

## 功能特点

- 支持Word文档(.docx)的解析
- 可视化操作界面
- 导出结果到Excel文件
- 支持自定义导出格式

## 安装要求

- Python 3.8+
- 依赖包：见requirements.txt

## 使用方法

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 运行程序：
```bash
python main.py
```

## 打包方法

使用PyInstaller打包成可执行文件：
```bash
pyinstaller --onefile --windowed main.py
``` 