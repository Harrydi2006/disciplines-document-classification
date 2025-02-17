import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import configparser
from pathlib import Path
import threading
from typing import Dict, Optional
import json
import queue
import logging
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

# 设置日志记录器
logger = logging.getLogger(__name__)

class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))

class MainWindow:
    def __init__(self, config: configparser.ConfigParser):
        self.root = tk.Tk()
        self.root.title("文件分类助手")
        self.root.geometry("1200x900")
        
        self.config = config
        self.processing = False
        self.files_status: Dict[str, str] = {}
        self.files_results: Dict[str, str] = {}
        self.log_queue = queue.Queue()
        
        self._setup_logger()
        self._init_ui()
        self._center_window()
        self._start_log_monitor()
        
    def _setup_logger(self):
        """设置日志处理器"""
        # 创建队列处理器
        self.queue_handler = QueueHandler(self.log_queue)
        self.queue_handler.setLevel(logging.INFO)  # 设置处理器级别
        self.queue_handler.setFormatter(
            logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', 
                            datefmt='%Y-%m-%d %H:%M:%S')
        )
        
        # 获取根日志记录器
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        
        # 移除所有现有的处理器
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
        
        # 添加队列处理器
        root_logger.addHandler(self.queue_handler)
        
        # 确保其他模块的日志也能显示
        logging.getLogger('file_classifier').propagate = True
        logging.getLogger('audio_processor').propagate = True
        
    def _init_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 添加关于按钮到主窗口右上角
        about_button = ttk.Button(
            self.root,  # 注意这里改为 self.root
            text="关于",
            command=self._show_about,
            width=8
        )
        about_button.pack(anchor=tk.NE, padx=10, pady=5)
        
        # 左侧：文件列表和控制按钮
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 路径设置
        path_frame = ttk.LabelFrame(left_frame, text="路径设置", padding="5")
        path_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 源文件夹
        source_frame = ttk.Frame(path_frame)
        source_frame.pack(fill=tk.X, pady=2)
        ttk.Label(source_frame, text="源文件夹:").pack(side=tk.LEFT)
        self.source_var = tk.StringVar(value=self.config['Paths']['source_folder'])
        source_entry = ttk.Entry(source_frame, textvariable=self.source_var)
        source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(
            source_frame,
            text="浏览",
            command=lambda: self._select_folder(self.source_var)
        ).pack(side=tk.LEFT)
        
        # 目标文件夹
        target_frame = ttk.Frame(path_frame)
        target_frame.pack(fill=tk.X, pady=2)
        ttk.Label(target_frame, text="目标文件夹:").pack(side=tk.LEFT)
        self.target_var = tk.StringVar(value=self.config['Paths']['target_folder'])
        target_entry = ttk.Entry(target_frame, textvariable=self.target_var)
        target_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(
            target_frame,
            text="浏览",
            command=lambda: self._select_folder(self.target_var)
        ).pack(side=tk.LEFT)
        
        # 文件列表
        list_frame = ttk.LabelFrame(left_frame, text="文件列表", padding="5")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建树形视图
        columns = ("文件名", "状态", "分类结果")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
        self.tree.column("文件名", width=300)
        self.tree.column("状态", width=100)
        self.tree.column("分类结果", width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 控制按钮
        control_frame = ttk.Frame(left_frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        self.start_button = ttk.Button(
            control_frame,
            text="开始分类",
            command=self._start_classification,
            state='disabled'  # 初始状态为禁用
        )
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.refresh_button = ttk.Button(
            control_frame,
            text="刷新列表",
            command=self._refresh_file_list,
            state='disabled'  # 初始状态为禁用
        )
        self.refresh_button.pack(side=tk.LEFT, padx=5)
        
        # 右侧：设置面板和日志
        right_frame = ttk.Frame(main_frame, width=400)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10)
        right_frame.pack_propagate(False)
        
        # 设置面板
        settings_frame = ttk.LabelFrame(right_frame, text="设置", padding="10")
        settings_frame.pack(fill=tk.X, expand=False)
        
        # 模型设置
        model_frame = ttk.LabelFrame(settings_frame, text="模型设置", padding="5")
        model_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(model_frame, text="模型名称:").pack(anchor=tk.W)
        self.model_name_var = tk.StringVar(value=self.config['Model']['model_name'])
        model_name_entry = ttk.Entry(model_frame, textvariable=self.model_name_var)
        model_name_entry.pack(fill=tk.X, pady=2)
        
        # API 设置
        api_frame = ttk.LabelFrame(settings_frame, text="API设置", padding="5")
        api_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(api_frame, text="API Key:").pack(anchor=tk.W)
        self.api_key_var = tk.StringVar(value=self.config['API']['api_key'])
        api_key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="*")
        api_key_entry.pack(fill=tk.X, pady=2)
        
        ttk.Label(api_frame, text="API Host:").pack(anchor=tk.W)
        self.api_host_var = tk.StringVar(value=self.config['API']['host'])
        api_host_entry = ttk.Entry(api_frame, textvariable=self.api_host_var)
        api_host_entry.pack(fill=tk.X, pady=2)
        
        # 添加测试按钮
        ttk.Button(
            api_frame,
            text="测试 API 连接",
            command=self._test_api
        ).pack(pady=5)
        
        # 提示词设置
        prompt_frame = ttk.LabelFrame(settings_frame, text="提示词设置", padding="5")
        prompt_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(prompt_frame, text="分类提示词:").pack(anchor=tk.W)
        self.prompt_var = tk.StringVar(value=self.config['Prompt']['classification_prompt'])
        prompt_entry = ttk.Entry(prompt_frame, textvariable=self.prompt_var)
        prompt_entry.pack(fill=tk.X, pady=2)
        
        # 功能设置
        features_frame = ttk.LabelFrame(settings_frame, text="功能设置", padding="5")
        features_frame.pack(fill=tk.X, pady=5)
        
        self.ocr_var = tk.BooleanVar(value=self.config.getboolean('Features', 'enable_ocr'))
        ttk.Checkbutton(features_frame, text="启用OCR", variable=self.ocr_var).pack(anchor=tk.W)
        
        self.audio_var = tk.BooleanVar(value=self.config.getboolean('Features', 'enable_audio'))
        ttk.Checkbutton(features_frame, text="启用音频识别", variable=self.audio_var).pack(anchor=tk.W)
        
        self.archive_var = tk.BooleanVar(value=self.config.getboolean('Features', 'enable_archive'))
        ttk.Checkbutton(features_frame, text="启用压缩包处理", variable=self.archive_var).pack(anchor=tk.W)
        
        # 线程设置
        thread_frame = ttk.LabelFrame(settings_frame, text="线程设置", padding="5")
        thread_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(thread_frame, text="最大线程数:").pack(side=tk.LEFT)
        self.max_workers_var = tk.StringVar(value=self.config['Threading']['max_workers'])
        max_workers_entry = ttk.Entry(thread_frame, textvariable=self.max_workers_var, width=5)
        max_workers_entry.pack(side=tk.LEFT, padx=5)
        
        # 外部程序设置
        external_frame = ttk.LabelFrame(settings_frame, text="外部程序设置", padding="5")
        external_frame.pack(fill=tk.X, pady=5)
        
        # OCR 路径
        ocr_path_frame = ttk.Frame(external_frame)
        ocr_path_frame.pack(fill=tk.X, pady=2)
        ttk.Label(ocr_path_frame, text="Tesseract 路径:").pack(side=tk.LEFT)
        self.tesseract_path_var = tk.StringVar(value=self.config['OCR']['tesseract_path'])
        tesseract_entry = ttk.Entry(ocr_path_frame, textvariable=self.tesseract_path_var)
        tesseract_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(
            ocr_path_frame,
            text="浏览",
            command=lambda: self._select_file(self.tesseract_path_var, "选择 Tesseract 可执行文件", "exe")
        ).pack(side=tk.LEFT)
        
        # FFmpeg 路径
        ffmpeg_path_frame = ttk.Frame(external_frame)
        ffmpeg_path_frame.pack(fill=tk.X, pady=2)
        ttk.Label(ffmpeg_path_frame, text="FFmpeg 路径:").pack(side=tk.LEFT)
        self.ffmpeg_path_var = tk.StringVar(value=self.config['Audio']['ffmpeg_path'])
        ffmpeg_entry = ttk.Entry(ffmpeg_path_frame, textvariable=self.ffmpeg_path_var)
        ffmpeg_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(
            ffmpeg_path_frame,
            text="浏览",
            command=lambda: self._select_file(self.ffmpeg_path_var, "选择 FFmpeg 可执行文件", "exe")
        ).pack(side=tk.LEFT)
        
        # 保存按钮
        ttk.Button(
            settings_frame,
            text="保存设置",
            command=self._save_settings
        ).pack(pady=10)
        
        # 日志显示
        log_frame = ttk.LabelFrame(right_frame, text="日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(
            log_frame, 
            wrap=tk.WORD, 
            height=20  # 增加日志文本框高度
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 检查并更新按钮状态
        self._update_buttons_state()
        
    def _select_folder(self, var: tk.StringVar):
        """选择文件夹"""
        folder = filedialog.askdirectory(
            initialdir=var.get() if var.get() else os.getcwd()
        )
        if folder:
            var.set(folder)
            # 保存设置
            self._save_settings()
            # 更新按钮状态
            self._update_buttons_state()
            
    def _select_file(self, var: tk.StringVar, title: str, file_type: str):
        filetypes = [("可执行文件", f"*.{file_type}")]
        filename = filedialog.askopenfilename(
            initialdir=str(Path(var.get()).parent),
            title=title,
            filetypes=filetypes
        )
        if filename:
            var.set(filename)
        
    def _start_log_monitor(self):
        """监控日志队列并更新显示"""
        def check_queue():
            while True:
                try:
                    record = self.log_queue.get_nowait()
                    self.log_text.insert(tk.END, record + '\n')
                    self.log_text.see(tk.END)
                    self.log_text.update_idletasks()  # 强制更新显示
                except queue.Empty:
                    break
            self.root.after(100, check_queue)
        
        check_queue()
        
    def _save_settings(self):
        """保存设置"""
        try:
            # 更新配置
            self.config['API']['api_key'] = self.api_key_var.get()
            self.config['API']['host'] = self.api_host_var.get()
            
            self.config['Model']['model_name'] = self.model_name_var.get()
            
            self.config['Paths']['source_folder'] = self.source_var.get()
            self.config['Paths']['target_folder'] = self.target_var.get()
            
            self.config['Prompt']['classification_prompt'] = self.prompt_var.get()
            
            self.config['Features']['enable_ocr'] = str(self.ocr_var.get())
            self.config['Features']['enable_audio'] = str(self.audio_var.get())
            self.config['Features']['enable_archive'] = str(self.archive_var.get())
            
            self.config['Threading']['max_workers'] = self.max_workers_var.get()
            
            self.config['OCR']['tesseract_path'] = self.tesseract_path_var.get()
            self.config['Audio']['ffmpeg_path'] = self.ffmpeg_path_var.get()
            
            # 保存到文件
            with open('config.conf', 'w', encoding='utf-8') as f:
                self.config.write(f)
                
            # 更新按钮状态
            self._update_buttons_state()
            
            messagebox.showinfo("成功", "设置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {str(e)}")
            
    def _refresh_file_list(self):
        """刷新文件列表"""
        # 清空现有列表
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 清空状态缓存
        self.files_status.clear()
        self.files_results.clear()
        
        try:
            # 重新加载配置
            self.config.read('config.conf', encoding='utf-8')
            
            # 获取源文件夹中的文件
            source_folder = Path(self.config['Paths']['source_folder'])
            if not source_folder.exists():
                logger.warning(f"源文件夹不存在: {source_folder}")
                return
            
            # 只获取当前目录下的文件（不遍历子目录）
            files = [f for f in source_folder.iterdir() if f.is_file()]
            
            # 添加到树形视图
            for file in files:
                file_str = str(file)
                self.tree.insert('', tk.END, values=(file.name, "等待处理", "未分类"))
                self.files_status[file_str] = "等待处理"
                self.files_results[file_str] = "未分类"
            
            logger.info(f"刷新文件列表完成，找到 {len(files)} 个文件")
            
        except Exception as e:
            logger.error(f"刷新文件列表失败: {str(e)}")
            messagebox.showerror("错误", f"刷新文件列表失败: {str(e)}")
        
    def _test_api(self):
        """测试 API 连接"""
        try:
            # 保存当前设置
            self._save_settings()
            
            # 创建临时分类器进行测试
            from file_classifier import FileClassifier
            classifier = FileClassifier()
            
            # 测试 API 调用
            test_content = "这是一个测试文本"
            result = classifier._call_api(test_content)
            
            messagebox.showinfo("测试成功", f"API 连接测试成功！\n返回结果: {result}")
        except Exception as e:
            messagebox.showerror("测试失败", f"API 连接测试失败：\n{str(e)}")

    def _check_config(self) -> bool:
        """检查配置是否完整"""
        # 检查 API 设置
        if not self.config['API']['api_key'] or not self.config['API']['host']:
            messagebox.showwarning("配置不完整", "请先配置 API 信息（主机地址和密钥）")
            return False
        
        if not self.config['Model']['model_name']:
            messagebox.showwarning("配置不完整", "请先配置模型名称")
            return False
        
        # 检查文件夹
        source_dir = Path(self.config['Paths']['source_folder'])
        target_dir = Path(self.config['Paths']['target_folder'])
        if not source_dir.exists() or not target_dir.exists():
            messagebox.showwarning("配置不完整", "请先设置并创建源文件夹和目标文件夹")
            return False
        
        return True

    def _start_classification(self):
        """开始分类处理"""
        if self.processing:
            return
        
        # 检查配置是否完整
        if not self._check_config():
            return
        
        self.processing = True
        self.start_button.config(state='disabled')
        
        # 先刷新文件列表
        self._refresh_file_list()
        
        def process_thread():
            logger = logging.getLogger('file_classifier')
            
            try:
                # 重新加载配置
                self.config.read('config.conf', encoding='utf-8')
                
                # 检查目录
                source_dir = Path(self.config['Paths']['source_folder'])
                target_dir = Path(self.config['Paths']['target_folder'])
                
                if not source_dir.exists():
                    logger.error(f"源文件夹不存在: {source_dir}")
                    messagebox.showerror("错误", f"源文件夹不存在: {source_dir}")
                    return
                
                if not target_dir.exists():
                    logger.info(f"创建目标文件夹: {target_dir}")
                    target_dir.mkdir(parents=True, exist_ok=True)
                
                from file_classifier import FileClassifier
                classifier = FileClassifier()
                
                # 只获取当前目录下的文件（不遍历子目录）
                files = [f for f in source_dir.iterdir() if f.is_file()]
                
                if not files:
                    logger.info("没有找到需要处理的文件")
                    messagebox.showinfo("提示", "没有找到需要处理的文件")
                    return
                
                logger.info(f"开始处理 {len(files)} 个文件")
                
                # 创建线程池
                max_workers = int(self.config['Threading']['max_workers'])
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    # 创建任务列表
                    future_to_file = {
                        executor.submit(classifier.process_single_file, file): file 
                        for file in files
                    }
                    
                    # 处理完成的任务
                    for future in as_completed(future_to_file):
                        file = future_to_file[future]
                        file_str = str(file)
                        try:
                            # 更新状态为"处理中"
                            self.files_status[file_str] = "处理中"
                            self.root.after(0, self._update_file_status, file.name, "处理中", "")
                            
                            # 获取处理结果
                            subject = future.result()
                            
                            # 更新状态和结果
                            self.files_status[file_str] = "已完成"
                            self.files_results[file_str] = subject
                            self.root.after(0, self._update_file_status, file.name, "已完成", subject)
                            
                        except Exception as e:
                            logger.error(f"处理文件 {file.name} 失败: {str(e)}")
                            self.files_status[file_str] = "失败"
                            self.root.after(0, self._update_file_status, file.name, "失败", "错误")
                
                logger.info("文件处理完成")
                # 最后再刷新一次文件列表
                self.root.after(0, self._refresh_file_list)
                messagebox.showinfo("完成", "文件处理完成")
                
            except Exception as e:
                error_msg = f"处理文件时发生错误: {str(e)}"
                logger.error(error_msg)
                messagebox.showerror("错误", error_msg)
            finally:
                self.processing = False
                self.root.after(0, lambda: self.start_button.config(state='normal'))
        
        threading.Thread(target=process_thread, daemon=True).start()
        
    def _update_file_status(self, filename: str, status: str, result: str):
        """更新文件状态"""
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][0] == filename:
                self.tree.item(item, values=(filename, status, result))
                break
        
        # 更新缓存
        for file_str in list(self.files_status.keys()):  # 使用 list 避免在迭代时修改
            if Path(file_str).name == filename:
                self.files_status[file_str] = status
                if result:  # 只在有结果时更新
                    self.files_results[file_str] = result
                
    def _center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def _update_buttons_state(self):
        """更新按钮状态"""
        source_path = self.source_var.get()
        target_path = self.target_var.get()
        
        if source_path and target_path and Path(source_path).exists() and Path(target_path).exists():
            self.start_button.config(state='normal')
            self.refresh_button.config(state='normal')
            self._refresh_file_list()  # 只在路径有效时刷新列表
        else:
            self.start_button.config(state='disabled')
            self.refresh_button.config(state='disabled')
            # 清空文件列表
            for item in self.tree.get_children():
                self.tree.delete(item)

    def _show_about(self):
        """显示关于对话框"""
        about_window = tk.Toplevel(self.root)
        about_window.title("关于文件分类助手")
        about_window.geometry("400x350")  # 增加关于窗口高度
        
        # 设置模态
        about_window.transient(self.root)
        about_window.grab_set()
        
        # 创建框架
        frame = ttk.Frame(about_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 软件标题
        title_label = ttk.Label(
            frame,
            text="文件分类助手",
            font=("微软雅黑", 16, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # 版本信息
        version_label = ttk.Label(
            frame,
            text="版本 1.0.0",
            font=("微软雅黑", 10)
        )
        version_label.pack()
        
        # 软件描述
        description = (
            "文件分类助手是一个基于人工智能的文档分类工具，\n"
            "能够自动识别并分类不同学科的文档。\n\n"
            "支持的功能：\n"
            "• 文本文件分类\n"
            "• OCR 图片识别\n"
            "• 音频文件转写\n"
            "• 压缩包处理"
        )
        desc_label = ttk.Label(
            frame,
            text=description,
            justify=tk.CENTER,
            wraplength=350
        )
        desc_label.pack(pady=20)
        
        # 作者信息和 GitHub 链接
        author_frame = ttk.Frame(frame)
        author_frame.pack(pady=(10, 20))  # 增加上下边距
        
        # 加载 GitHub 图标
        try:
            # 创建并保存 GitHub 图标的 base64 字符串
            github_icon = """
            iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAAdgAAAHYBTnsmCAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAADSSURBVDiNxZIxbsJAEEX/NyshCqcgBS2iQEEJBXfIBXKAXCGHyBFyAe5AQUGBUkSJkE2BIhcrcbFjr7WrxJGmGM3M/zPz96+Nqv4CeAZwB2DQtYEEPgG8A/gwAF4ADHE7pQC2AKbGOYcVgB2+x78H8Gic8/B7gdtpBbDQQ+P0zLG6vgWQmUMi4o1z+5/+1MnPyqW1XgGYA8gy9Oc3gIlzzn+rUNU+gFcAj1gvFxuPEuK99dMmJtLc2xBvrbUPYBeHlwBe1PJ0PLnFBs89gB9z7QlxfhBEIQAAAABJRU5ErkJggg==
            """
            github_image = tk.PhotoImage(data=github_icon)
            
            author_label = ttk.Label(
                author_frame,
                text="作者: "
            )
            author_label.pack(side=tk.LEFT)
            
            github_link = ttk.Label(
                author_frame,
                text="Harrydi2006",
                cursor="hand2",
                image=github_image,
                compound=tk.LEFT
            )
            github_link.image = github_image  # 保持引用
            github_link.pack(side=tk.LEFT)
            
            # 绑定点击事件
            def open_github(event):
                import webbrowser
                webbrowser.open("https://github.com/Harrydi2006")
            
            github_link.bind("<Button-1>", open_github)
            
        except Exception as e:
            # 如果图标加载失败，使用纯文本
            author_label = ttk.Label(
                author_frame,
                text="作者: Harrydi2006 (GitHub)",
                cursor="hand2"
            )
            author_label.pack()
            author_label.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/Harrydi2006"))
        
        # 版权信息
        copyright_label = ttk.Label(
            frame,
            text="© 2024 All Rights Reserved",
            font=("微软雅黑", 8)
        )
        copyright_label.pack(side=tk.BOTTOM, pady=20)  # 增加底部边距
        
        # 居中窗口
        about_window.update_idletasks()
        width = about_window.winfo_width()
        height = about_window.winfo_height()
        x = (about_window.winfo_screenwidth() // 2) - (width // 2)
        y = (about_window.winfo_screenheight() // 2) - (height // 2)
        about_window.geometry(f'{width}x{height}+{x}+{y}')

    def run(self):
        self.root.mainloop() 