"""
WPS文件转换工具 - GUI版本
支持WPS&DOC转DOCX和WPS转Markdown
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
import sys
import os

# 添加当前目录到路径，确保能导入wps_parse模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from wps_parse.wps_to_docx import wps_to_docx
from wps_parse.wps_to_markdown import wps_to_md


class WPSConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("WPS文件转换工具 v1.0")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 设置窗口图标
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        self.setup_ui()
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="WPS文件转换工具", 
                               font=("微软雅黑", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 输入文件选择
        ttk.Label(main_frame, text="选择WPS文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.input_var = tk.StringVar()
        input_entry = ttk.Entry(main_frame, textvariable=self.input_var, width=50)
        input_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_input_file).grid(row=1, column=2, pady=5)
        
        # 输出目录选择
        ttk.Label(main_frame, text="输出目录:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar()
        output_entry = ttk.Entry(main_frame, textvariable=self.output_var, width=50)
        output_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_output_dir).grid(row=2, column=2, pady=5)
        
        # 转换格式选择
        format_frame = ttk.LabelFrame(main_frame, text="转换格式", padding="10")
        format_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)
        
        self.format_var = tk.StringVar(value="docx")
        ttk.Radiobutton(format_frame, text="转换为DOCX格式", variable=self.format_var, 
                       value="docx").grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Radiobutton(format_frame, text="转换为Markdown格式", variable=self.format_var, 
                       value="markdown").grid(row=0, column=1, sticky=tk.W)
        
        # 高级选项
        advanced_frame = ttk.LabelFrame(main_frame, text="高级选项", padding="10")
        advanced_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(advanced_frame, text="可读性阈值:").grid(row=0, column=0, sticky=tk.W)
        self.threshold_var = tk.DoubleVar(value=1.0)
        threshold_scale = ttk.Scale(advanced_frame, from_=0.5, to=1.0, 
                                   variable=self.threshold_var, orient=tk.HORIZONTAL)
        threshold_scale.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        self.threshold_label = ttk.Label(advanced_frame, text="1.0")
        self.threshold_label.grid(row=0, column=2)
        
        # 绑定阈值变化事件
        threshold_scale.configure(command=self.update_threshold_label)
        
        # 转换按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        self.convert_button = ttk.Button(button_frame, text="开始转换", 
                                        command=self.start_conversion, 
                                        style="Accent.TButton")
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="清空", command=self.clear_fields).pack(side=tk.LEFT)
        
        # 进度条
        self.progress_var = tk.StringVar(value="准备就绪")
        ttk.Label(main_frame, textvariable=self.progress_var).grid(row=6, column=0, columnspan=3, pady=5)
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 日志输出区域
        log_frame = ttk.LabelFrame(main_frame, text="转换日志", padding="5")
        log_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(8, weight=1)
        
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
    def update_threshold_label(self, value):
        """更新阈值标签显示"""
        self.threshold_label.config(text=f"{float(value):.1f}")
        
    def browse_input_file(self):
        """浏览选择输入文件"""
        filename = filedialog.askopenfilename(
            title="选择WPS文件",
            filetypes=[("WPS文件", "*.wps"), ("所有文件", "*.*")]
        )
        if filename:
            self.input_var.set(filename)
            # 自动设置输出目录为输入文件所在目录
            if not self.output_var.get():
                self.output_var.set(str(Path(filename).parent))
                
    def browse_output_dir(self):
        """浏览选择输出目录"""
        dirname = filedialog.askdirectory(title="选择输出目录")
        if dirname:
            self.output_var.set(dirname)
            
    def clear_fields(self):
        """清空所有字段"""
        self.input_var.set("")
        self.output_var.set("")
        self.threshold_var.set(1.0)
        self.log_text.delete(1.0, tk.END)
        self.progress_var.set("准备就绪")
        
    def log_message(self, message):
        """添加日志消息"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def start_conversion(self):
        """开始转换过程"""
        # 验证输入
        if not self.input_var.get():
            messagebox.showerror("错误", "请选择要转换的WPS文件")
            return
            
        if not self.output_var.get():
            messagebox.showerror("错误", "请选择输出目录")
            return
            
        input_path = Path(self.input_var.get())
        if not input_path.exists():
            messagebox.showerror("错误", "输入文件不存在")
            return
            
        output_dir = Path(self.output_var.get())
        if not output_dir.exists():
            messagebox.showerror("错误", "输出目录不存在")
            return
            
        # 在新线程中执行转换
        self.convert_button.config(state="disabled")
        self.progress_bar.start()
        
        thread = threading.Thread(target=self.perform_conversion)
        thread.daemon = True
        thread.start()
        
    def perform_conversion(self):
        """执行实际的转换操作"""
        try:
            input_path = Path(self.input_var.get())
            output_dir = Path(self.output_var.get())
            format_type = self.format_var.get()
            threshold = self.threshold_var.get()
            
            self.progress_var.set("正在转换...")
            self.log_message(f"开始转换: {input_path.name}")
            self.log_message(f"输出目录: {output_dir}")
            self.log_message(f"转换格式: {format_type.upper()}")
            self.log_message(f"可读性阈值: {threshold}")
            
            if format_type == "docx":
                output_path = output_dir / f"{input_path.stem}_converted.docx"
                wps_to_docx(input_path, output_path, threshold)
                self.log_message(f"DOCX文件已生成: {output_path}")
            else:
                output_path = output_dir / f"{input_path.stem}_converted.md"
                wps_to_md(input_path, output_path, threshold)
                self.log_message(f"Markdown文件已生成: {output_path}")
                
            self.progress_var.set("转换完成！")
            self.log_message("转换成功完成！")
            
            # 询问是否打开输出目录
            self.root.after(0, lambda: self.ask_open_output(output_dir))
            
        except Exception as e:
            error_msg = f"转换失败: {str(e)}"
            self.progress_var.set("转换失败")
            self.log_message(error_msg)
            self.root.after(0, lambda: messagebox.showerror("转换失败", error_msg))
            
        finally:
            self.root.after(0, self.conversion_finished)
            
    def ask_open_output(self, output_dir):
        """询问是否打开输出目录"""
        if messagebox.askyesno("转换完成", "转换完成！是否打开输出目录？"):
            try:
                os.startfile(str(output_dir))
            except:
                pass
                
    def conversion_finished(self):
        """转换完成后的清理工作"""
        self.progress_bar.stop()
        self.convert_button.config(state="normal")


def main():
    root = tk.Tk()
    app = WPSConverterGUI(root)
    
    # 设置窗口居中
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    root.mainloop()


if __name__ == "__main__":
    main()