import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from difflib import SequenceMatcher
import time
import os
from pdfminer.high_level import extract_text
from docx import Document
from datetime import datetime

'''auther:杜明静；版本：v2.1；日期：20250430'''

class HTMLFileComparator(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("数科文档查重工具 v2.1")
        self.geometry("680x480")
        self.configure(bg='#F5F5F5')
        
        # 初始化组件
        self.create_widgets()
        self.file_paths = {1: "", 2: ""}

    def create_widgets(self):
        """创建界面组件"""
        # 程序图标
        if os.path.exists("icon.ico"):
            self.iconbitmap("icon.ico")

        # 拖放区域
        self.create_drop_zone(1, "拖放第一个文件", 50)
        self.create_drop_zone(2, "拖放第二个文件", 150)
        
        # 控制面板
        control_frame = tk.Frame(self, bg='#F5F5F5')
        control_frame.pack(pady=15)
        
        # 操作按钮
        self.btn_compare = tk.Button(control_frame, text="开始比对", 
                                   command=self.start_comparison,
                                   width=15, height=1,
                                   bg='#2196F3', fg='white',
                                   font=('微软雅黑', 11))
        self.btn_compare.pack(side=tk.LEFT, padx=10)

        self.btn_clear = tk.Button(control_frame, text="清空记录", 
                                 command=self.clear_input,
                                 width=15, height=1,
                                 bg='#607D8B', fg='white',
                                 font=('微软雅黑', 11))
        self.btn_clear.pack(side=tk.LEFT)

        # 结果展示
        self.result_var = tk.StringVar()
        result_label = tk.Label(self, textvariable=self.result_var,
                               font=('微软雅黑', 12),
                               bg='#F5F5F5', fg='#009688')
        result_label.pack(pady=15)

        # 状态栏
        self.status_var = tk.StringVar()
        status_bar = tk.Label(self, textvariable=self.status_var,
                            bd=1, relief=tk.SUNKEN, anchor=tk.W,
                            bg='#E0E0E0', fg='#616161',
                            font=('微软雅黑', 9))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def create_drop_zone(self, num, text, y_pos):
        """创建文件拖放区域"""
        frame = tk.LabelFrame(self, text=text, 
                            bg='white', fg='#757575',
                            font=('微软雅黑', 10),
                            width=580, height=80)
        frame.place(x=50, y=y_pos)
        
        label = tk.Label(frame, text="", bg='white',
                       font=('微软雅黑', 9), wraplength=500)
        label.pack(pady=10)
        
        frame.drop_target_register(DND_FILES)
        frame.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, num))
        setattr(self, f"label{num}", label)

    def on_drop(self, event, num):
        """处理文件拖放事件"""
        path = event.data.strip('{}')
        if os.path.isfile(path):
            self.file_paths[num] = path
            getattr(self, f"label{num}").config(text=path)
            self.status_var.set(f"已加载文件{num}: {os.path.basename(path)}")
            self.result_var.set("")

    def read_file(self, path):
        """读取不同格式的文件内容"""
        ext = os.path.splitext(path)[1].lower()
        
        try:
            if ext == '.pdf':
                return extract_text(path)
            elif ext == '.docx':
                doc = Document(path)
                return '\n'.join([p.text for p in doc.paragraphs])
            else:
                with open(path, 'r', encoding='utf-8') as f:
                    return f.read()
        except Exception as e:
            raise RuntimeError(f"文件读取失败: {str(e)}")

    def start_comparison(self):
        """启动比对流程"""
        if not all(self.file_paths.values()):
            self.show_error("请先拖放两个文件")
            return

        try:
            start_time = time.time()
            self.status_var.set("正在分析文件...")
            self.update()
            
            # 读取文件内容
            text1 = self.read_file(self.file_paths[1]).lower()
            text2 = self.read_file(self.file_paths[2]).lower()
            
            # 计算相似度
            similarity = SequenceMatcher(None, text1, text2).ratio()
            
            # 生成标注文件
            self.generate_highlighted_files(text1, text2)
            
            elapsed = time.time() - start_time
            self.result_var.set(f"文档重复率: {similarity*100:.2f}%\n耗时: {elapsed:.2f}秒")
            self.status_var.set("比对完成 - HTML文件已生成在源文件目录")
        except Exception as e:
            self.show_error(str(e))

    def generate_highlighted_files(self, text1, text2):
        """生成HTML标注文件"""
        matcher = SequenceMatcher(None, text1, text2)
        blocks = matcher.get_matching_blocks()

        # 为两个文件生成标注
        self.create_html_file(self.file_paths[1], text1, 
                            [(b.a, b.a + b.size) for b in blocks if b.size > 0])
        self.create_html_file(self.file_paths[2], text2,
                            [(b.b, b.b + b.size) for b in blocks if b.size > 0])

    def create_html_file(self, orig_path, text, matches):
        """创建HTML输出文件"""
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        base_name = os.path.splitext(os.path.basename(orig_path))[0]
        output_dir = os.path.join(os.path.dirname(orig_path), "对比结果")
        os.makedirs(output_dir, exist_ok=True)
        
        output_path = os.path.join(output_dir, f"{base_name}_对比_{timestamp}.html")
        
        # 生成高亮HTML内容
        highlighted = []
        last_pos = 0
        for start, end in sorted(matches, key=lambda x: x[0]):
            highlighted.append(text[last_pos:start])
            highlighted.append(f'<mark style="background:yellow">{text[start:end]}</mark>')
            last_pos = end
        highlighted.append(text[last_pos:])
        
        # 构建完整HTML文档
        html_template = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{base_name} 对比结果</title>
    <style>
        body {{ 
            font-family: 'Microsoft YaHei', sans-serif;
            line-height: 1.6;
            margin: 2rem;
        }}
        mark {{
            background-color: #FFFF00;
            padding: 2px 4px;
            border-radius: 3px;
        }}
        .container {{
            max-width: 800px;
            margin: 0 auto;
        }}
        h1 {{ color: #333; }}
        pre {{
            white-space: pre-wrap;
            word-wrap: break-word;
            background: #f8f9fa;
            padding: 1rem;
            border-radius: 5px;
            border: 1px solid #ddd;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>文档对比结果</h1>
        <pre>{''.join(highlighted)}</pre>
    </div>
</body>
</html>"""
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_template)
        
        return output_path

    def show_error(self, msg):
        """显示错误信息"""
        self.result_var.set("")
        self.status_var.set(f"错误: {msg}")
        self.after(5000, lambda: self.status_var.set(""))

    def clear_input(self):
        """清空所有输入"""
        for num in [1, 2]:
            self.file_paths[num] = ""
            getattr(self, f"label{num}").config(text="")
        self.result_var.set("")
        self.status_var.set("状态已重置")

if __name__ == "__main__":
    app = HTMLFileComparator()
    app.mainloop()