# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from word_parser import WordParser
import os

class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word题库解析工具")
        self.root.geometry("600x400")
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加说明标签
        info_label = ttk.Label(main_frame, text="请选择Word文档进行解析，支持.docx格式")
        info_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # 创建按钮
        self.select_file_btn = ttk.Button(main_frame, text="选择Word文件", command=self.select_word_file)
        self.select_file_btn.grid(row=1, column=0, padx=5, pady=5)
        
        self.parse_btn = ttk.Button(main_frame, text="开始解析", command=self.parse_document, state="disabled")
        self.parse_btn.grid(row=1, column=1, padx=5, pady=5)
        
        # 添加进度条
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # 状态标签
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.grid(row=3, column=0, columnspan=2, pady=10)
        
        self.word_file = None
        self.parser = WordParser()
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

    def select_word_file(self):
        file_name = filedialog.askopenfilename(
            title="选择Word文档",
            filetypes=[("Word Documents", "*.docx")]
        )
        if file_name:
            self.word_file = file_name
            self.parse_btn.config(state="normal")
            self.status_label.config(text="已选择文件: {}".format(os.path.basename(file_name)))

    def parse_document(self):
        if not self.word_file:
            return
            
        save_path = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if save_path:
            self.progress_bar.start()
            self.parse_btn.config(state="disabled")
            self.select_file_btn.config(state="disabled")
            
            try:
                self.parser.parse_document(self.word_file, save_path)
                self.status_label.config(text="解析完成！文件已保存。")
                messagebox.showinfo("成功", "解析完成！文件已保存。")
            except Exception as e:
                self.status_label.config(text="解析失败：{}".format(str(e)))
                messagebox.showerror("错误", "解析失败：{}".format(str(e)))
            finally:
                self.progress_bar.stop()
                self.parse_btn.config(state="normal")
                self.select_file_btn.config(state="normal")

    def run(self):
        self.root.mainloop()

def main():
    app = MainWindow()
    app.run()

if __name__ == "__main__":
    main()