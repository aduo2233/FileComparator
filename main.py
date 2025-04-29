import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from difflib import SequenceMatcher
import time
import os
from pdfminer.high_level import extract_text
from docx import Document

class FileComparator(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("文档查重工具 v2.0")
        self.geometry("600x400")
        self.configure(bg='#F5F5F5')
        
        # 程序图标（需准备icon.ico）
        if os.path.exists("icon.ico"):
            self.iconbitmap("icon.ico")

        # 拖放区域
        self.create_drop_zone(1, "拖放第一个文件", 50)
        self.create_drop_zone(2, "拖放第二个文件", 150)
        
        # 控制面板
        control_frame = tk.Frame(self, bg='#F5F5F5')
        control_frame.pack(pady=20)
        
        self.btn_compare = tk.Button(control_frame, text="开始比对", 
                                   command=self.start_comparison,
                                   width=15, height=2,
                                   bg='#2196F3', fg='white',
                                   font=('微软雅黑', 12))
        self.btn_compare.pack(side=tk.LEFT, padx=10)

        self.btn_clear = tk.Button(control_frame, text="清空记录", 
                                 command=self.clear_input,
                                 width=15, height=2,
                                 bg='#607D8B', fg='white',
                                 font=('微软雅黑', 12))
        self.btn_clear.pack(side=tk.LEFT)

        # 结果展示
        self.result_var = tk.StringVar()
        result_label = tk.Label(self, textvariable=self.result_var,
                               font=('微软雅黑', 14, 'bold'),
                               bg='#F5F5F5', fg='#009688')
        result_label.pack(pady=20)

        # 状态栏
        self.status_var = tk.StringVar()
        status_bar = tk.Label(self, textvariable=self.status_var,
                            bd=1, relief=tk.SUNKEN, anchor=tk.W,
                            bg='#E0E0E0', fg='#616161')
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 文件路径存储
        self.file_paths = {1: "", 2: ""}

    def create_drop_zone(self, num, text, y_pos):
        frame = tk.LabelFrame(self, text=text, 
                            bg='white', fg='#757575',
                            font=('微软雅黑', 10),
                            width=500, height=80)
        frame.place(x=50, y=y_pos)
        
        label = tk.Label(frame, text="", bg='white',
                       font=('微软雅黑', 9), wraplength=400)
        label.pack(pady=10)
        
        frame.drop_target_register(DND_FILES)
        frame.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, num))
        setattr(self, f"label{num}", label)

    def on_drop(self, event, num):
        path = event.data.strip('{}')  # Windows路径处理
        if os.path.isfile(path):
            self.file_paths[num] = path
            getattr(self, f"label{num}").config(text=path)
            self.status_var.set(f"已加载文件{num}: {os.path.basename(path)}")

    def read_file(self, path):
        """智能读取不同格式文件"""
        ext = os.path.splitext(path)[1].lower()
        
        try:
            if ext == '.pdf':
                return extract_text(path)
            elif ext == '.docx':
                doc = Document(path)
                return '\n'.join([p.text for p in doc.paragraphs])
            else:  # 默认按文本文件读取
                with open(path, 'r', encoding='utf-8') as f:
                    return f.read()
        except Exception as e:
            raise RuntimeError(f"读取文件失败: {str(e)}")

    def start_comparison(self):
        if not all(self.file_paths.values()):
            self.show_error("请先拖放两个文件")
            return

        start_time = time.time()
        
        try:
            text1 = self.read_file(self.file_paths[1]).lower()
            text2 = self.read_file(self.file_paths[2]).lower()
            
            similarity = SequenceMatcher(None, text1, text2).ratio()
            elapsed = time.time() - start_time
            
            self.result_var.set(f"文档重复率: {similarity*100:.2f}%\n耗时: {elapsed:.2f}秒")
            self.status_var.set("比对完成")
        except Exception as e:
            self.show_error(str(e))

    def show_error(self, msg):
        self.result_var.set("")
        self.status_var.set(f"错误: {msg}")
        self.after(3000, lambda: self.status_var.set(""))

    def clear_input(self):
        for num in [1, 2]:
            self.file_paths[num] = ""
            getattr(self, f"label{num}").config(text="")
        self.result_var.set("")
        self.status_var.set("已清空所有记录")

if __name__ == "__main__":
    app = FileComparator()
    app.mainloop()