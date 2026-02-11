# _*_coding : UTF-8 _*_
# @Time : 2026/2/6 22:00
# @Author : Murchey
# @File : gui
# @Project : python
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import getData as gd
import compare as cp
import multipleFiles as mf
import getStandardData as gsd
from pathlib import Path

class HandleTheBillsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("账单处理系统")
        
        # 获取屏幕分辨率
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        
        # 设置窗口大小为屏幕的70%，确保在不同分辨率下都有合适的大小
        window_width = int(screen_width * 0.7)
        window_height = int(screen_height * 0.73)
        
        # 设置窗口位置在屏幕中央
        x = int((screen_width - window_width) / 2)
        y = int((screen_height - window_height) / 2)
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(True, True)
        
        # 设置主题
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 初始化暗黑模式状态
        self.dark_mode = False
        
        # 初始化行列参数变量，这样在切换页面时不会丢失用户输入
        self.name_col_var = tk.StringVar(value="C")
        self.money_col_var = tk.StringVar(value="K")
        self.begin_row_var = tk.StringVar(value="3")
        self.standard_name_col_var = tk.StringVar(value="B")
        self.standard_money_col_var = tk.StringVar(value="F")
        
        # 创建必要的文件夹
        try:
            mf.newFolders()
        except Exception as e:
            print(f"创建文件夹时出错: {e}")
            # 即使创建文件夹失败，也继续运行程序
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建顶部框架，包含暗黑模式切换按钮和标题
        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.pack(fill=tk.X, pady=10)
        
        # 创建暗黑模式切换按钮
        self.dark_mode_var = tk.BooleanVar(value=self.dark_mode)
        self.dark_mode_button = ttk.Checkbutton(
            self.top_frame, 
            text="显示模式切换",
            variable=self.dark_mode_var, 
            command=self.toggle_dark_mode, 
            style='ToggleButton.TCheckbutton'
        )
        self.dark_mode_button.pack(side=tk.LEFT, padx=10)
        
        # 创建标题
        self.title_label = ttk.Label(self.top_frame, text="账单处理系统", font=("SimHei", 28, "bold"))
        self.title_label.pack(side=tk.LEFT, padx=20)
        
        # 创建功能按钮框架
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(pady=20)
        
        # 横向排列功能按钮
        # 数据提取按钮
        self.extract_button = ttk.Button(self.button_frame, text="1.需核对表格处理", command=self.show_extract_frame, width=18)
        self.extract_button.pack(side=tk.LEFT, padx=10, pady=10)
        
        # 标准表格处理按钮
        self.standard_button = ttk.Button(self.button_frame, text="2.标准表格处理", command=self.show_standard_frame, width=15)
        self.standard_button.pack(side=tk.LEFT, padx=10, pady=10)
        
        # 数据比对按钮
        self.compare_button = ttk.Button(self.button_frame, text="3.数据比对", command=self.show_compare_frame, width=15)
        self.compare_button.pack(side=tk.LEFT, padx=10, pady=10)
        
        # 退出按钮
        self.exit_button = ttk.Button(self.button_frame, text="退出", command=root.quit, width=15)
        self.exit_button.pack(side=tk.LEFT, padx=10, pady=10)
        
        # 创建内容框架
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=20)
        
        # 初始化功能框架
        self.extract_frame = None
        self.compare_frame = None
        
        # 显示欢迎信息
        self.show_welcome_frame()
        
        # 初始化样式
        self.init_styles()
    
    def init_styles(self):
        """初始化样式"""
        # 扁平化按钮样式
        self.style.configure('TButton', 
            relief='flat',
            padding=8,
            font=('SimHei', 12)
        )
        
        # 切换按钮样式
        self.style.configure('ToggleButton.TCheckbutton',
            relief='flat',
            font=('SimHei', 14)
        )
        
        # 标签框架样式
        self.style.configure('TLabelframe',
            relief='flat',
            borderwidth=1
        )
        
        # 应用初始样式
        if self.dark_mode:
            # 暗黑模式
            bg_color = '#2d2d2d'
            fg_color = '#e0e0e0'
            frame_bg = '#3d3d3d'
            button_bg = '#4d4d4d'
            button_fg = '#e0e0e0'
            
            # 设置根窗口背景
            self.root.configure(bg=bg_color)
            
            # 设置样式
            self.style.configure('.', 
                background=bg_color,
                foreground=fg_color
            )
            
            self.style.configure('TFrame', background=bg_color)
            self.style.configure('TLabel', background=bg_color, foreground=fg_color)
            self.style.configure('TButton', 
                background=button_bg,
                foreground=button_fg,
                relief='flat'
            )
            self.style.configure('TLabelframe', 
                background=bg_color,
                foreground=fg_color,
                borderwidth=1,
                relief='flat'
            )
            self.style.configure('TLabelframe.Label', 
                background=bg_color,
                foreground=fg_color
            )
            self.style.configure('TListbox', 
                background=frame_bg,
                foreground=fg_color,
                relief='flat'
            )
            self.style.configure('Text', 
                background=frame_bg,
                foreground=fg_color
            )
        else:
            # 亮色模式
            bg_color = '#ffffff'
            fg_color = '#000000'
            frame_bg = '#f0f0f0'
            button_bg = '#e0e0e0'
            button_fg = '#000000'
            
            # 设置根窗口背景
            self.root.configure(bg=bg_color)
            
            # 设置样式
            self.style.configure('.', 
                background=bg_color,
                foreground=fg_color
            )
            
            self.style.configure('TFrame', background=bg_color)
            self.style.configure('TLabel', background=bg_color, foreground=fg_color)
            self.style.configure('TButton', 
                background=button_bg,
                foreground=button_fg,
                relief='flat'
            )
            self.style.configure('TLabelframe', 
                background=bg_color,
                foreground=fg_color,
                borderwidth=1,
                relief='flat'
            )
            self.style.configure('TLabelframe.Label', 
                background=bg_color,
                foreground=fg_color
            )
            self.style.configure('TListbox', 
                background=frame_bg,
                foreground=fg_color,
                relief='flat'
            )
            self.style.configure('Text', 
                background=frame_bg,
                foreground=fg_color
            )
    
    def toggle_dark_mode(self):
        """切换暗黑模式"""
        self.dark_mode = self.dark_mode_var.get()
        
        if self.dark_mode:
            # 暗黑模式
            bg_color = '#2d2d2d'
            fg_color = '#e0e0e0'
            frame_bg = '#3d3d3d'
            button_bg = '#4d4d4d'
            button_fg = '#e0e0e0'
            
            # 设置根窗口背景
            self.root.configure(bg=bg_color)
            
            # 设置样式
            self.style.configure('.', 
                background=bg_color,
                foreground=fg_color
            )
            
            self.style.configure('TFrame', background=bg_color)
            self.style.configure('TLabel', background=bg_color, foreground=fg_color)
            self.style.configure('TButton', 
                background=button_bg,
                foreground=button_fg,
                relief='flat'
            )
            self.style.configure('TLabelframe', 
                background=bg_color,
                foreground=fg_color,
                borderwidth=1,
                relief='flat'
            )
            self.style.configure('TLabelframe.Label', 
                background=bg_color,
                foreground=fg_color
            )
            self.style.configure('TListbox', 
                background=frame_bg,
                foreground=fg_color,
                relief='flat'
            )
            self.style.configure('Text', 
                background=frame_bg,
                foreground=fg_color
            )
            
            # 更新文本框背景
            if hasattr(self, 'result_text') and self.result_text.winfo_exists():
                self.result_text.configure(bg=frame_bg, fg='#000000')  # 纯黑色
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.configure(bg=frame_bg, fg='#000000')  # 纯黑色
            
            # 更新暗黑模式按钮文本
            self.dark_mode_button.configure(text="☀️")
        else:
            # 亮色模式
            bg_color = '#ffffff'
            fg_color = '#000000'
            frame_bg = '#f0f0f0'
            button_bg = '#e0e0e0'
            button_fg = '#000000'
            
            # 设置根窗口背景
            self.root.configure(bg=bg_color)
            
            # 设置样式
            self.style.configure('.', 
                background=bg_color,
                foreground=fg_color
            )
            
            self.style.configure('TFrame', background=bg_color)
            self.style.configure('TLabel', background=bg_color, foreground=fg_color)
            self.style.configure('TButton', 
                background=button_bg,
                foreground=button_fg,
                relief='flat'
            )
            self.style.configure('TLabelframe', 
                background=bg_color,
                foreground=fg_color,
                borderwidth=1,
                relief='flat'
            )
            self.style.configure('TLabelframe.Label', 
                background=bg_color,
                foreground=fg_color
            )
            self.style.configure('TListbox', 
                background=frame_bg,
                foreground=fg_color,
                relief='flat'
            )
            self.style.configure('Text', 
                background=frame_bg,
                foreground=fg_color
            )
            
            # 更新文本框背景
            if hasattr(self, 'result_text') and self.result_text.winfo_exists():
                self.result_text.configure(bg=frame_bg, fg=fg_color)
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.configure(bg=frame_bg, fg=fg_color)
            
            # 更新暗黑模式按钮文本
            self.dark_mode_button.configure(text="🌙")
    
    def show_welcome_frame(self):
        """显示欢迎信息"""
        # 清空内容框架
        self.clear_content_frame()
        
        welcome_frame = ttk.Frame(self.content_frame)
        welcome_frame.pack(fill=tk.BOTH, expand=True)
        
        welcome_label = ttk.Label(welcome_frame, text="欢迎使用账单处理系统", font=("SimHei", 20))
        welcome_label.pack(pady=20)

        info_text = "\n使用说明：\n"
        info_text += "1. 点击“需核对表格处理”按钮，放入需核对表格，设置参数\n"
        info_text += "2. 点击“标准表格提取”按钮，放入标准表格，设置参数\n"
        info_text += "3. 点击“数据比对”按钮，进行数据比对\n"
        info_text += "4. 处理结果将作为表格文件保存在相应的文件夹中\n"
        info_text += "\n详细设置见 使用说明书.docx\n"
        info_label = ttk.Label(welcome_frame, text=info_text, font=("SimHei", 14), justify=tk.LEFT)
        info_label.pack(padx=20, pady=10)
        
        # 显示目录结构
        dir_info = "\n当前目录结构：\n"
        dirs = ["需核对表格", "标准表格", "已提取数据", "比对结果"]
        # 获取应用程序的根目录
        import sys
        if getattr(sys, 'frozen', False):
            # 打包后的exe模式
            app_dir = Path(sys.executable).resolve().parent
        else:
            # 脚本模式
            app_dir = Path(__file__).resolve().parent
        for dir_name in dirs:
            dir_path = app_dir / dir_name
            if dir_path.exists():
                files = list(dir_path.glob('*.xlsx'))
                dir_info += f"{dir_name}: {len(files)} 个文件\n"
                for file in files[:3]:  # 只显示前3个文件
                    dir_info += f"  - {file.name}\n"
                if len(files) > 3:
                    dir_info += f"  ... 等{len(files)}个文件\n"
            else:
                dir_info += f"{dir_name}: 目录不存在\n"
        
        dir_label = ttk.Label(welcome_frame, text=dir_info, font=("SimHei", 12), justify=tk.LEFT)
        dir_label.pack(padx=20, pady=10)
    
    def show_extract_frame(self):
        """显示数据提取界面"""
        # 清空内容框架
        self.clear_content_frame()
        
        # 创建数据提取框架
        self.extract_frame = ttk.Frame(self.content_frame)
        self.extract_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        extract_title = ttk.Label(self.extract_frame, text="需核对表格设置", font=("SimHei", 16, "bold"))
        extract_title.pack(pady=10)
        
        # 列设置框架
        col_frame = ttk.LabelFrame(self.extract_frame, text="列设置", padding="10")
        col_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 姓名列设置
        ttk.Label(col_frame, text="姓名列（如A、B、C）:", font=("SimHei", 12)).pack(side=tk.LEFT, padx=5)
        ttk.Entry(col_frame, textvariable=self.name_col_var, width=5, font=("SimHei", 12)).pack(side=tk.LEFT, padx=5)
        
        # 金额列设置
        ttk.Label(col_frame, text="金额列（如A、B、C）:", font=("SimHei", 12)).pack(side=tk.LEFT, padx=5)
        ttk.Entry(col_frame, textvariable=self.money_col_var, width=5, font=("SimHei", 12)).pack(side=tk.LEFT, padx=5)
        
        # 文件列表框架
        file_frame = ttk.LabelFrame(self.extract_frame, text="待处理文件", padding="10")
        file_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建滚动条
        scrollbar = ttk.Scrollbar(file_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建文件列表
        self.file_listbox = tk.Listbox(file_frame, yscrollcommand=scrollbar.set, font=("SimHei", 12))
        self.file_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # 结果文本框
        result_frame = ttk.LabelFrame(self.extract_frame, text="提取结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 设置文本框背景颜色
        text_bg = '#3d3d3d' if self.dark_mode else '#f0f0f0'
        text_fg = '#000000'  # 始终使用纯黑色
        
        self.result_text = tk.Text(result_frame, height=10, font=("SimHei", 12), bg=text_bg, fg=text_fg)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        result_scrollbar = ttk.Scrollbar(self.result_text)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=result_scrollbar.set)
        result_scrollbar.config(command=self.result_text.yview)
        
        # 按钮框架
        button_frame = ttk.Frame(self.extract_frame)
        button_frame.pack(pady=10)
        

        
        # 刷新按钮
        refresh_button = ttk.Button(button_frame, text="刷新列表", command=self.load_extract_files, width=15)
        refresh_button.pack(side=tk.LEFT, padx=10)
        
        # 清除已提取数据按钮
        clear_button = ttk.Button(button_frame, text="清除已提取数据", command=self.clear_extracted_data, width=15)
        clear_button.pack(side=tk.LEFT, padx=10)
        
        # 打开需核对表格文件夹按钮
        open_check_folder_button = ttk.Button(button_frame, text="打开需核对表格文件夹", command=self.open_check_folder, width=20)
        open_check_folder_button.pack(side=tk.LEFT, padx=10)
        
        # 加载文件列表（在所有 UI 元素创建后）
        self.load_extract_files()
    
    def load_extract_files(self):
        """加载待提取文件列表"""
        # 清空列表
        self.file_listbox.delete(0, tk.END)
        
        # 获取文件列表
        try:
            files = mf.getFilesNames("需核对表格")
            if files:
                for file in files:
                    self.file_listbox.insert(tk.END, file.name)
                self.result_text.insert(tk.END, f"已找到 {len(files)} 个待处理文件\n")
            else:
                self.result_text.insert(tk.END, "未找到待处理文件，请将文件放入 '需核对表格' 文件夹\n")
        except Exception as e:
            messagebox.showerror("错误", f"加载文件列表失败: {e}")
    
    def start_extract(self):
        """开始数据提取"""
        try:
            # 获取文件列表
            files = mf.getFilesNames("需核对表格")
            if not files:
                messagebox.showwarning("警告", "未找到待处理文件，请将文件放入 '需核对表格' 文件夹")
                return
            
            # 清空结果文本
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "开始提取数据...\n")
            self.result_text.update()
            
            # 获取用户设置的列参数
            name_col = self.name_col_var.get().strip().upper()
            money_col = self.money_col_var.get().strip().upper()
            
            # 验证输入
            if not name_col:
                name_col = "C"
                self.name_col_var.set("C")
            if not money_col:
                money_col = "J"
                self.money_col_var.set("J")
            
            # 处理每个文件
            for file in files:
                self.result_text.insert(tk.END, f"处理文件: {file.name}\n")
                self.result_text.insert(tk.END, f"使用列: 姓名={name_col}, 金额={money_col}\n")
                self.result_text.update()
                
                try:
                    # 调用数据提取函数
                    print(f"调用gd.mainFunc: {file}, 已提取数据, {name_col}, {money_col}")
                    gd.mainFunc(file, "已提取数据", name_col, money_col)
                    print(f"gd.mainFunc调用完成: {file.name}")
                except Exception as e:
                    print(f"处理文件 {file.name} 时出错: {e}")
                    self.result_text.insert(tk.END, f"处理文件 {file.name} 时出错: {e}\n")
                
                self.result_text.insert(tk.END, f"文件 {file.name} 提取完成\n")
                self.result_text.update()
            
            self.result_text.insert(tk.END, "\n所有文件提取完成！\n")
            self.result_text.insert(tk.END, "提取结果已保存到 '已提取数据' 文件夹\n")
            messagebox.showinfo("成功", "数据提取完成！")
        except Exception as e:
            print(f"提取数据失败: {e}")
            messagebox.showerror("错误", f"提取数据失败: {e}")
            self.result_text.insert(tk.END, f"提取过程中出错: {e}\n")
    
    def clear_extracted_data(self):
        """清除已提取数据文件夹下的所有文件以及子文件"""
        try:
            # 确认用户操作
            confirm = messagebox.askyesno("确认", "确定要清除已提取数据文件夹下的所有文件以及子文件吗？")
            if not confirm:
                return
            
            # 获取应用程序的根目录，无论是脚本还是打包后的exe
            import multipleFiles as mf
            app_dir = mf.get_application_dir()
            
            # 获取已提取数据文件夹路径
            extracted_data_path = app_dir / "已提取数据"
            
            # 检查文件夹是否存在
            if not extracted_data_path.exists():
                messagebox.showinfo("信息", "已提取数据文件夹不存在")
                return
            
            # 获取文件夹中的所有文件（包括子文件夹中的文件）
            all_files = list(extracted_data_path.rglob('*'))
            if not all_files:
                messagebox.showinfo("信息", "已提取数据文件夹为空")
                return
            
            # 删除所有文件
            deleted_count = 0
            for file in all_files:
                if file.is_file():
                    file.unlink()
                    deleted_count += 1
            
            # 删除所有空文件夹
            for file in all_files:
                if file.is_dir() and not list(file.glob('*')):
                    file.rmdir()
            
            # 显示结果
            messagebox.showinfo("成功", f"已成功清除 {deleted_count} 个文件")
            
            # 刷新结果文本
            if hasattr(self, 'result_text') and self.result_text.winfo_exists():
                self.result_text.insert(tk.END, f"\n已成功清除 {deleted_count} 个已提取数据文件\n")
        except Exception as e:
            messagebox.showerror("错误", f"清除数据失败: {e}")
    
    def open_check_folder(self):
        """打开需核对表格文件夹"""
        try:
            # 获取应用程序的根目录，无论是脚本还是打包后的exe
            import multipleFiles as mf
            app_dir = mf.get_application_dir()
            
            # 获取需核对表格文件夹路径
            check_folder_path = app_dir / "需核对表格"
            
            # 检查文件夹是否存在
            if not check_folder_path.exists():
                # 如果文件夹不存在，创建它
                check_folder_path.mkdir(parents=True, exist_ok=True)
                messagebox.showinfo("信息", "需核对表格文件夹不存在，已自动创建")
            
            # 打开文件夹
            import os
            if os.name == 'nt':  # Windows
                os.startfile(check_folder_path)
            else:  # macOS or Linux
                import subprocess
                subprocess.run(['open', check_folder_path] if os.name == 'posix' else ['xdg-open', check_folder_path])
        except Exception as e:
            messagebox.showerror("错误", f"打开文件夹失败: {e}")
    
    def open_export_folder(self):
        """打开比对结果文件夹"""
        try:
            # 获取应用程序的根目录，无论是脚本还是打包后的exe
            import multipleFiles as mf
            app_dir = mf.get_application_dir()
            
            # 获取比对结果文件夹路径
            export_folder_path = app_dir / "比对结果"
            
            # 检查文件夹是否存在
            if not export_folder_path.exists():
                # 如果文件夹不存在，创建它
                export_folder_path.mkdir(parents=True, exist_ok=True)
                messagebox.showinfo("信息", "比对结果文件夹不存在，已自动创建")
            
            # 打开文件夹
            import os
            if os.name == 'nt':  # Windows
                os.startfile(export_folder_path)
            else:  # macOS or Linux
                import subprocess
                subprocess.run(['open', export_folder_path] if os.name == 'posix' else ['xdg-open', export_folder_path])
        except Exception as e:
            messagebox.showerror("错误", f"打开文件夹失败: {e}")
    
    def show_compare_frame(self):
        """显示数据比对界面"""
        # 清空内容框架
        self.clear_content_frame()
        
        # 创建数据比对框架
        self.compare_frame = ttk.Frame(self.content_frame)
        self.compare_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        compare_title = ttk.Label(self.compare_frame, text="数据比对", font=("SimHei", 16, "bold"))
        compare_title.pack(pady=10)
        
        # 说明文本和按钮框架
        top_frame = ttk.Frame(self.compare_frame)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 说明文本
        info_frame = ttk.LabelFrame(top_frame, text="比对说明", padding="10")
        info_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        info_text = "比对规则：\n"
        info_text += "1. 程序会自动按照同名文件进行比对\n"
        info_text += "2. 标准表格文件名：例如 ""ClassOne.xlsx""\n"
        info_text += "3. 已提取数据文件名：例如 ""ClassOne_new.xlsx""\n"
        info_text += "4. 比对结果会保存为 ""原文件名_比对结果.txt""\n"
        
        info_label = ttk.Label(info_frame, text=info_text, font=("SimHei", 12), justify=tk.LEFT)
        info_label.pack(fill=tk.X, padx=5)
        
        # 按钮框架
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(side=tk.RIGHT, padx=(10, 0))
        
        # 开始比对按钮
        start_button = ttk.Button(button_frame, text="开始比对", command=self.start_compare, width=15)
        start_button.pack(pady=5)
        
        # 刷新按钮
        refresh_button = ttk.Button(button_frame, text="刷新列表", command=self.load_compare_files, width=15)
        refresh_button.pack(pady=5)
        
        # 打开文件夹按钮
        open_folder_button = ttk.Button(button_frame, text="打开结果文件夹", command=self.open_export_folder, width=15)
        open_folder_button.pack(pady=5)
        
        # 标准表格列表
        std_frame = ttk.LabelFrame(self.compare_frame, text="标准表格", padding="10")
        std_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 创建滚动条
        std_scrollbar = ttk.Scrollbar(std_frame)
        std_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建文件列表
        self.std_listbox = tk.Listbox(std_frame, yscrollcommand=std_scrollbar.set, font=("SimHei", 12), height=5)
        self.std_listbox.pack(fill=tk.X, padx=5)
        std_scrollbar.config(command=self.std_listbox.yview)
        
        # 待比对文件列表
        compare_files_frame = ttk.LabelFrame(self.compare_frame, text="待比对文件", padding="10")
        compare_files_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建滚动条
        scrollbar = ttk.Scrollbar(compare_files_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建文件列表
        self.compare_listbox = tk.Listbox(compare_files_frame, yscrollcommand=scrollbar.set, font=("SimHei", 12))
        self.compare_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.compare_listbox.yview)
        
        # 结果文本框
        result_frame = ttk.LabelFrame(self.compare_frame, text="比对结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 设置文本框背景颜色
        text_bg = '#3d3d3d' if self.dark_mode else '#f0f0f0'
        text_fg = '#000000'  # 始终使用纯黑色
        
        self.compare_result_text = tk.Text(result_frame, height=10, font=("SimHei", 12), bg=text_bg, fg=text_fg)
        self.compare_result_text.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        result_scrollbar = ttk.Scrollbar(self.compare_result_text)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.compare_result_text.config(yscrollcommand=result_scrollbar.set)
        result_scrollbar.config(command=self.compare_result_text.yview)
        
        # 加载文件列表（在所有 UI 元素创建后）
        self.load_compare_files()
    
    def browse_std_file(self):
        """浏览选择标准表格文件"""
        file_path = filedialog.askopenfilename(
            title="选择标准表格文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.std_file_var.set(file_path)
    
    def load_compare_files(self):
        """加载待比对文件列表"""
        # 清空列表
        if hasattr(self, 'std_listbox'):
            self.std_listbox.delete(0, tk.END)
        if hasattr(self, 'compare_listbox'):
            self.compare_listbox.delete(0, tk.END)
        
        # 获取文件列表
        try:
            # 获取应用程序的根目录，无论是脚本还是打包后的exe
            import multipleFiles as mf
            app_dir = mf.get_application_dir()
            
            # 打印调试信息
            print(f"应用程序目录: {app_dir}")
            
            # 加载标准表格数据（从已提取数据/标准表格数据文件夹）
            std_data_dir = app_dir / "已提取数据" / "标准表格数据"
            std_data_dir.mkdir(parents=True, exist_ok=True)  # 确保文件夹存在
            std_files = list(p for p in std_data_dir.glob('*.xlsx'))
            print(f"标准表格数据文件: {[f.name for f in std_files]}")
            
            if hasattr(self, 'std_listbox'):
                if std_files:
                    for file in std_files:
                        self.std_listbox.insert(tk.END, file.name)
                else:
                    self.std_listbox.insert(tk.END, "未找到标准表格数据文件，请先运行标准表格处理")
            
            # 加载待比对文件（从已提取数据文件夹）
            print("开始获取已提取数据文件夹中的文件...")
            files = mf.getFilesNames("已提取数据")
            print(f"获取到的待比对文件: {[f.name for f in files]}")
            
            # 清空结果文本
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.delete(1.0, tk.END)
                if files:
                    self.compare_result_text.insert(tk.END, f"已找到 {len(std_files)} 个标准表格数据文件\n")
                    self.compare_result_text.insert(tk.END, f"已找到 {len(files)} 个待比对文件\n")
                else:
                    self.compare_result_text.insert(tk.END, "未找到待比对文件，请先进行数据提取\n")
            
            if hasattr(self, 'compare_listbox'):
                if files:
                    for file in files:
                        self.compare_listbox.insert(tk.END, file.name)
                else:
                    self.compare_listbox.insert(tk.END, "未找到待比对文件，请先进行数据提取")
        except Exception as e:
            messagebox.showerror("错误", f"加载文件列表失败: {e}")
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, f"加载文件列表失败: {e}\n")
            print(f"加载文件列表失败: {e}")
    
    def start_compare(self):
        """开始数据比对"""
        try:
            # 获取应用程序的根目录，无论是脚本还是打包后的exe
            import multipleFiles as mf
            app_dir = mf.get_application_dir()
            
            # 清空结果文本
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.delete(1.0, tk.END)
                self.compare_result_text.insert(tk.END, "开始数据比对流程...\n")
                self.compare_result_text.update()
            
            # 1. 删除已提取数据文件夹下的所有文件
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, "步骤1: 清理已提取数据文件夹...\n")
                self.compare_result_text.update()
            
            extracted_data_path = app_dir / "已提取数据"
            if extracted_data_path.exists():
                # 获取文件夹中的所有文件（包括子文件夹中的文件）
                all_files = list(extracted_data_path.rglob('*'))
                for file in all_files:
                    if file.is_file():
                        file.unlink()
                # 删除所有空文件夹
                for file in all_files:
                    if file.is_dir() and not list(file.glob('*')):
                        file.rmdir()
            
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, "已清理已提取数据文件夹\n\n")
                self.compare_result_text.update()
            
            # 2. 处理“需核对表格”文件夹中的文件
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, "步骤2: 处理需核对表格...\n")
                self.compare_result_text.update()
            
            # 获取需核对表格文件列表
            check_files = mf.getFilesNames("需核对表格")
            if not check_files:
                messagebox.showwarning("警告", "未找到需核对表格文件，请将文件放入 '需核对表格' 文件夹")
                return
            
            # 获取用户设置的列参数
            name_col = "C"  # 默认值
            money_col = "J"  # 默认值
            
            if hasattr(self, 'name_col_var'):
                name_col = self.name_col_var.get().strip().upper()
                if not name_col:
                    name_col = "C"
            
            if hasattr(self, 'money_col_var'):
                money_col = self.money_col_var.get().strip().upper()
                if not money_col:
                    money_col = "J"
            
            # 处理每个需核对表格文件
            for file in check_files:
                if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                    self.compare_result_text.insert(tk.END, f"处理文件: {file.name}\n")
                    self.compare_result_text.insert(tk.END, f"使用列: 姓名={name_col}, 金额={money_col}\n")
                    self.compare_result_text.update()
                
                # 调用数据提取函数
                try:
                    import getData as gd
                    gd.mainFunc(file, "已提取数据", name_col, money_col)
                except Exception as e:
                    if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                        self.compare_result_text.insert(tk.END, f"处理文件 {file.name} 时出错: {e}\n")
            
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, "需核对表格处理完成\n\n")
                self.compare_result_text.update()
            
            # 3. 处理“标准表格”文件夹中的文件
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, "步骤3: 处理标准表格...\n")
                self.compare_result_text.update()
            
            # 获取标准表格文件列表
            std_files = mf.getFilesNames("标准表格")
            if not std_files:
                messagebox.showwarning("警告", "未找到标准表格文件，请将文件放入 '标准表格' 文件夹")
                return
            
            # 获取标准表格处理参数
            begin_row = 3  # 默认值
            std_name_col = "B"  # 默认值
            std_money_col = "F"  # 默认值
            
            if hasattr(self, 'begin_row_var'):
                begin_row_str = self.begin_row_var.get().strip()
                if begin_row_str:
                    try:
                        begin_row = int(begin_row_str)
                    except ValueError:
                        begin_row = 3
            
            if hasattr(self, 'standard_name_col_var'):
                std_name_col = self.standard_name_col_var.get().strip().upper()
                if not std_name_col:
                    std_name_col = "B"
            
            if hasattr(self, 'standard_money_col_var'):
                std_money_col = self.standard_money_col_var.get().strip().upper()
                if not std_money_col:
                    std_money_col = "F"
            
            # 处理每个标准表格文件
            import getStandardData as gsd
            for file in std_files:
                if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                    self.compare_result_text.insert(tk.END, f"处理文件: {file.name}\n")
                    self.compare_result_text.update()
                
                # 调用标准表格处理函数
                try:
                    data = gsd.getValuableData(file, begin_row, std_name_col, std_money_col)
                    # 保存处理结果
                    output_file_name = f"{file.stem}_处理结果.xlsx"
                    output_path = extracted_data_path / "标准表格数据" / output_file_name
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    gsd.saveFile(data, output_path)
                except Exception as e:
                    if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                        self.compare_result_text.insert(tk.END, f"处理文件 {file.name} 时出错: {e}\n")
            
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, "标准表格处理完成\n\n")
                self.compare_result_text.update()
            
            # 4. 进行数据比对
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, "步骤4: 开始数据比对...\n")
                self.compare_result_text.update()
            
            # 获取标准表格数据（从已提取数据/标准表格数据文件夹）
            std_data_dir = extracted_data_path / "标准表格数据"
            std_data_files = list(p for p in std_data_dir.glob('*.xlsx'))
            
            if not std_data_files:
                messagebox.showwarning("警告", "未找到标准表格数据文件，请检查标准表格处理是否成功")
                return
            
            # 获取待比对文件（从已提取数据文件夹）
            compare_files = mf.getFilesNames("已提取数据")
            if not compare_files:
                messagebox.showwarning("警告", "未找到待比对文件，请检查需核对表格处理是否成功")
                return
            
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, f"找到 {len(std_data_files)} 个标准表格数据文件\n")
                self.compare_result_text.insert(tk.END, f"找到 {len(compare_files)} 个待比对文件\n\n")
                self.compare_result_text.update()
            
            # 创建标准表格字典，以文件名（不含扩展名）为键
            std_file_dict = {}
            for file in std_data_files:
                # 获取文件名（不含扩展名），去掉末尾的 "_处理结果"
                file_name = file.stem
                if file_name.endswith("_处理结果"):
                    file_name = file_name[:-5]  # 去掉 "_处理结果"（5个字符）
                std_file_dict[file_name] = file
            
            # 处理每个待比对文件
            matched_count = 0
            unmatched_count = 0
            
            import compare as cp
            for file in compare_files:
                # 获取文件名（不含扩展名），去掉末尾的 "_new"
                file_name = file.stem
                if file_name.endswith("_new"):
                    file_name = file_name[:-4]  # 去掉 "_new"
                
                if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                    self.compare_result_text.insert(tk.END, f"处理文件: {file.name}\n")
                
                # 查找对应的标准表格
                if file_name in std_file_dict:
                    std_file = std_file_dict[file_name]
                    if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                        self.compare_result_text.insert(tk.END, f"匹配标准表格: {std_file.name}\n")
                        self.compare_result_text.update()
                    
                    # 调用比对函数
                    cp.compare_and_save(std_file, file)
                    
                    if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                        self.compare_result_text.insert(tk.END, f"文件 {file.name} 比对完成\n\n")
                    matched_count += 1
                else:
                    if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                        self.compare_result_text.insert(tk.END, f"未找到对应标准表格，跳过比对\n\n")
                    unmatched_count += 1
                
                if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                    self.compare_result_text.update()
            
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, f"比对完成！\n")
                self.compare_result_text.insert(tk.END, f"成功比对: {matched_count} 个文件\n")
                self.compare_result_text.insert(tk.END, f"未匹配: {unmatched_count} 个文件\n")
                self.compare_result_text.insert(tk.END, "比对结果已保存到 '比对结果' 文件夹\n")
            messagebox.showinfo("成功", f"数据比对流程完成！成功比对 {matched_count} 个文件，未匹配 {unmatched_count} 个文件")
        except Exception as e:
            messagebox.showerror("错误", f"比对数据失败: {e}")
            if hasattr(self, 'compare_result_text') and self.compare_result_text.winfo_exists():
                self.compare_result_text.insert(tk.END, f"比对过程中出错: {e}\n")
    
    def show_standard_frame(self):
        """显示标准表格处理界面"""
        # 清空内容框架
        self.clear_content_frame()
        
        # 创建标准表格处理框架
        self.standard_frame = ttk.Frame(self.content_frame)
        self.standard_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        standard_title = ttk.Label(self.standard_frame, text="标准表格处理", font=("SimHei", 16, "bold"))
        standard_title.pack(pady=10)
        
        # 参数设置框架
        param_frame = ttk.LabelFrame(self.standard_frame, text="参数设置", padding="10")
        param_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 表头行设置
        ttk.Label(param_frame, text="表头所在行（数字）:", font=("SimHei", 12)).grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Entry(param_frame, textvariable=self.begin_row_var, width=10, font=("SimHei", 12)).grid(row=0, column=1, padx=10, pady=5)
        
        # 姓名列设置
        ttk.Label(param_frame, text="姓名列（如A、B、C）:", font=("SimHei", 12)).grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Entry(param_frame, textvariable=self.standard_name_col_var, width=10, font=("SimHei", 12)).grid(row=1, column=1, padx=10, pady=5)
        
        # 金额列设置
        ttk.Label(param_frame, text="金额列（如A、B、C）:", font=("SimHei", 12)).grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Entry(param_frame, textvariable=self.standard_money_col_var, width=10, font=("SimHei", 12)).grid(row=2, column=1, padx=10, pady=5)
        
        # 文件选择框架
        file_frame = ttk.LabelFrame(self.standard_frame, text="批量处理", padding="10")
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 批量处理说明
        ttk.Label(file_frame, text="将处理 '标准表格' 文件夹下的所有文件", font=("SimHei", 10)).pack(padx=10, pady=5)
        
        # 显示标准表格文件夹中的文件
        import multipleFiles as mf
        std_files = mf.getFilesNames("标准表格")
        if std_files:
            file_list_text = "标准表格文件夹中的文件：\n"
            for file in std_files:
                file_list_text += f"- {file.name}\n"
        else:
            file_list_text = "标准表格文件夹为空，请先放入标准表格文件"
        
        ttk.Label(file_frame, text=file_list_text, font=("SimHei", 9), justify=tk.LEFT).pack(padx=10, pady=5)
        
        # 按钮框架
        button_frame = ttk.Frame(self.standard_frame)
        button_frame.pack(pady=10)
        

        
        # 打开标准表格文件夹按钮
        open_folder_button = ttk.Button(button_frame, text="打开标准表格文件夹", command=self.open_standard_folder, width=20)
        open_folder_button.pack(side=tk.LEFT, padx=10)
        
        # 结果文本框
        result_frame = ttk.LabelFrame(self.standard_frame, text="处理结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 设置文本框背景颜色
        text_bg = '#3d3d3d' if self.dark_mode else '#f0f0f0'
        text_fg = '#000000'  # 始终使用纯黑色
        
        self.standard_result_text = tk.Text(result_frame, height=10, font=("SimHei", 12), bg=text_bg, fg=text_fg)
        self.standard_result_text.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        result_scrollbar = ttk.Scrollbar(self.standard_result_text)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.standard_result_text.config(yscrollcommand=result_scrollbar.set)
        result_scrollbar.config(command=self.standard_result_text.yview)
    
    def browse_standard_file(self):
        """浏览选择标准表格文件"""
        file_path = filedialog.askopenfilename(
            title="选择标准表格文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.standard_file_var.set(file_path)
    
    def start_standard_process(self):
        """开始标准表格处理（批量）"""
        try:
            # 获取参数
            begin_row_str = self.begin_row_var.get().strip()
            name_col = self.standard_name_col_var.get().strip().upper()
            money_col = self.standard_money_col_var.get().strip().upper()
            
            # 验证参数
            if not begin_row_str or not name_col or not money_col:
                messagebox.showwarning("警告", "请填写所有参数")
                return
            
            try:
                begin_row = int(begin_row_str)
                if begin_row < 1:
                    messagebox.showwarning("警告", "表头所在行必须大于0")
                    return
            except ValueError:
                messagebox.showwarning("警告", "表头所在行必须是数字")
                return
            
            # 清空结果文本
            self.standard_result_text.delete(1.0, tk.END)
            self.standard_result_text.insert(tk.END, "开始批量处理标准表格...\n")
            self.standard_result_text.insert(tk.END, f"表头所在行: {begin_row}\n")
            self.standard_result_text.insert(tk.END, f"姓名列: {name_col}\n")
            self.standard_result_text.insert(tk.END, f"金额列: {money_col}\n\n")
            self.standard_result_text.update()
            
            # 获取应用程序的根目录，无论是脚本还是打包后的exe
            import sys
            if getattr(sys, 'frozen', False):
                # 打包后的exe模式
                app_dir = Path(sys.executable).resolve().parent
            else:
                # 脚本模式
                app_dir = Path(__file__).resolve().parent
            
            # 获取标准表格文件夹中的所有文件
            std_dir = app_dir / "标准表格"
            std_files = [p for p in std_dir.glob('*.xlsx')]
            
            if not std_files:
                messagebox.showwarning("警告", "标准表格文件夹为空，请先放入标准表格文件")
                return
            
            # 构建保存路径
            save_dir = app_dir / "已提取数据" / "标准表格数据"
            save_dir.mkdir(parents=True, exist_ok=True)
            
            # 处理每个文件
            processed_count = 0
            failed_count = 0
            
            for file_path in std_files:
                try:
                    self.standard_result_text.insert(tk.END, f"处理文件: {file_path.name}\n")
                    self.standard_result_text.update()
                    
                    # 提取数据
                    data = gsd.getValuableData(file_path, begin_row, name_col, money_col)
                    
                    self.standard_result_text.insert(tk.END, f"成功提取 {len(data)} 条数据\n")
                    
                    # 获取文件名
                    file_name = file_path.stem
                    save_path = save_dir / f"{file_name}_处理结果.xlsx"
                    
                    # 保存数据
                    success = gsd.saveFile(data, save_path)
                    
                    if success:
                        self.standard_result_text.insert(tk.END, f"数据保存成功！\n")
                        self.standard_result_text.insert(tk.END, f"保存路径: {save_path}\n\n")
                        processed_count += 1
                    else:
                        self.standard_result_text.insert(tk.END, "数据保存失败！\n\n")
                        failed_count += 1
                except Exception as e:
                    self.standard_result_text.insert(tk.END, f"处理过程中出错: {e}\n\n")
                    failed_count += 1
                
                self.standard_result_text.update()
            
            # 显示处理结果
            if processed_count > 0:
                self.standard_result_text.insert(tk.END, f"批量处理完成！\n")
                self.standard_result_text.insert(tk.END, f"成功处理: {processed_count} 个文件\n")
                self.standard_result_text.insert(tk.END, f"失败: {failed_count} 个文件\n")
                messagebox.showinfo("成功", f"批量处理完成！成功处理 {processed_count} 个文件，失败 {failed_count} 个文件")
            else:
                messagebox.showerror("错误", "所有文件处理失败，请检查参数和文件格式")
        except Exception as e:
            messagebox.showerror("错误", f"处理标准表格失败: {e}")
            if hasattr(self, 'standard_result_text') and self.standard_result_text.winfo_exists():
                self.standard_result_text.insert(tk.END, f"处理过程中出错: {e}\n")
    
    def open_standard_folder(self):
        """打开标准表格文件夹"""
        try:
            # 获取应用程序的根目录，无论是脚本还是打包后的exe
            import sys
            if getattr(sys, 'frozen', False):
                # 打包后的exe模式
                app_dir = Path(sys.executable).resolve().parent
            else:
                # 脚本模式
                app_dir = Path(__file__).resolve().parent
            
            # 获取标准表格文件夹路径
            standard_folder_path = app_dir / "标准表格"
            
            # 检查文件夹是否存在
            if not standard_folder_path.exists():
                # 如果文件夹不存在，创建它
                standard_folder_path.mkdir(parents=True, exist_ok=True)
                messagebox.showinfo("信息", "标准表格文件夹不存在，已自动创建")
            
            # 打开文件夹
            import os
            if os.name == 'nt':  # Windows
                os.startfile(standard_folder_path)
            else:  # macOS or Linux
                import subprocess
                subprocess.run(['open', standard_folder_path] if os.name == 'posix' else ['xdg-open', standard_folder_path])
        except Exception as e:
            messagebox.showerror("错误", f"打开文件夹失败: {e}")
    
    def clear_content_frame(self):
        """清空内容框架"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    try:
        # 创建主窗口
        root = tk.Tk()
        
        # 创建应用实例
        app = HandleTheBillsGUI(root)
        
        # 运行主循环
        root.mainloop()
    except Exception as e:
        print(f"应用程序启动失败: {e}")
        # 尝试显示错误信息
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()  # 隐藏主窗口
            messagebox.showerror("错误", f"应用程序启动失败: {e}")
            root.destroy()
        except:
            pass
