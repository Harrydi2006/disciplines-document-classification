import tkinter as tk
from tkinter import ttk
import requests
from pathlib import Path
import zipfile
import os
import threading
from tqdm import tqdm
import shutil
import sys
import configparser
import time
import tkinter.messagebox as messagebox
import logging
from tkinter import filedialog

logger = logging.getLogger(__name__)

class SetupWindow:
    # 主要下载地址
    FFMPEG_URL = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
    TESSERACT_URL = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe"
    
    # 镜像地址
    FFMPEG_MIRRORS = [
        "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip",
        "https://huggingface.co/datasets/your-mirror/ffmpeg/resolve/main/ffmpeg-master-latest-win64-gpl.zip",
        "https://mirror.ghproxy.com/https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
    ]
    
    TESSERACT_MIRRORS = [
        "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe",
        "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.1.20230401/tesseract-ocr-w64-setup-5.3.1.20230401.exe",
        "https://mirror.ghproxy.com/https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.1.20230401/tesseract-ocr-w64-setup-5.3.1.20230401.exe"
    ]
    
    # 更新中文语言包下载地址
    TESSERACT_CHI_SIM_MIRRORS = [
        "https://github.com/tesseract-ocr/tessdata/raw/4.1.0/chi_sim.traineddata",  # 使用稳定版本
        "https://raw.githubusercontent.com/tesseract-ocr/tessdata/4.1.0/chi_sim.traineddata",
        "https://ghproxy.com/https://raw.githubusercontent.com/tesseract-ocr/tessdata/4.1.0/chi_sim.traineddata"
    ]
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("文件分类助手 - 首次运行设置")
        self.root.geometry("800x600")  # 调整初始窗口大小
        
        # 设置窗口最小尺寸
        self.root.minsize(600, 500)
        
        # 添加窗口关闭处理
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # 设置窗口样式
        self.style = ttk.Style()
        self.style.configure('Horizontal.TProgressbar', thickness=20)
        
        # 检查是否是完整版
        self.is_full_version = self._check_full_version()
        
        self._init_ui()
        self._center_window()
        
    def _check_full_version(self) -> bool:
        """检查是否是完整版安装包"""
        deps_dir = Path('dependencies')
        if not deps_dir.exists():
            if getattr(sys, '_MEIPASS', None):
                deps_dir = Path(sys._MEIPASS) / 'dependencies'
            
        if deps_dir.exists():
            tesseract_installer = deps_dir / 'tesseract-installer.exe'
            ffmpeg_zip = deps_dir / 'ffmpeg.zip'
            return tesseract_installer.exists() and ffmpeg_zip.exists()
        return False

    def _init_ui(self):
        # 创建主框架，添加内边距
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text="欢迎使用文件分类助手",
            font=('Microsoft YaHei UI', 16, 'bold')
        )
        title_label.pack(pady=(0, 20))
        
        # 说明文本
        if self.is_full_version:
            desc_text = "首次运行需要安装必要的组件，所有依赖已包含在安装包中"
        else:
            desc_text = "首次运行需要下载并安装必要的组件，请保持网络连接"
            
        desc_label = ttk.Label(
            main_frame,
            text=desc_text,
            font=('Microsoft YaHei UI', 10)
        )
        desc_label.pack(pady=(0, 20))
        
        # 进度框架
        self.progress_frame = ttk.LabelFrame(main_frame, text="安装进度", padding=10)
        self.progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        # FFmpeg 进度
        self.ffmpeg_label = ttk.Label(self.progress_frame, text="FFmpeg: 等待安装")
        self.ffmpeg_label.pack(fill=tk.X, pady=(0, 5))
        self.ffmpeg_progress = ttk.Progressbar(
            self.progress_frame, 
            style='Horizontal.TProgressbar',
            mode='determinate'
        )
        self.ffmpeg_progress.pack(fill=tk.X, pady=(0, 10))
        
        # Tesseract 进度
        self.tesseract_label = ttk.Label(self.progress_frame, text="Tesseract: 等待安装")
        self.tesseract_label.pack(fill=tk.X, pady=(0, 5))
        self.tesseract_progress = ttk.Progressbar(
            self.progress_frame, 
            style='Horizontal.TProgressbar',
            mode='determinate'
        )
        self.tesseract_progress.pack(fill=tk.X)
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="安装日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建日志文本框和滚动条的容器
        log_container = ttk.Frame(log_frame)
        log_container.pack(fill=tk.BOTH, expand=True)
        
        # 日志文本框
        self.log_text = tk.Text(
            log_container, 
            wrap=tk.WORD, 
            height=12,  # 增加默认高度
            font=('Consolas', 9)  # 使用等宽字体
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        log_scrollbar = ttk.Scrollbar(log_container, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 底部控制区域
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 状态信息
        self.status_label = ttk.Label(
            control_frame, 
            text="准备开始安装...",
            font=('Microsoft YaHei UI', 9)
        )
        self.status_label.pack(side=tk.LEFT, pady=5)
        
        # 开始按钮
        self.start_button = ttk.Button(
            control_frame, 
            text="开始安装",
            command=self._start_installation,
            width=15  # 设置按钮宽度
        )
        self.start_button.pack(side=tk.RIGHT, pady=5)
        
        # 文件列表
        list_frame = ttk.LabelFrame(main_frame, text="文件列表", padding="5")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 添加按钮框架
        buttons_frame = ttk.Frame(list_frame)
        buttons_frame.pack(fill=tk.X, pady=(0, 5))
        
        # 添加文件和文件夹按钮
        ttk.Button(
            buttons_frame,
            text="添加文件",
            command=self._add_files
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            buttons_frame,
            text="添加文件夹",
            command=self._add_folder
        ).pack(side=tk.LEFT, padx=2)
        
        # 选择按钮
        ttk.Button(
            buttons_frame,
            text="全选",
            command=self._select_all
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            buttons_frame,
            text="反选",
            command=self._invert_selection
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            buttons_frame,
            text="取消全选",
            command=self._deselect_all
        ).pack(side=tk.LEFT, padx=2)
        
        # 文件类型选择按钮
        ttk.Button(
            buttons_frame,
            text="选择文档",
            command=lambda: self._select_by_type('doc')
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            buttons_frame,
            text="选择压缩包",
            command=lambda: self._select_by_type('archive')
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            buttons_frame,
            text="选择音频",
            command=lambda: self._select_by_type('audio')
        ).pack(side=tk.LEFT, padx=2)
        
        # 创建树形视图
        columns = ("选择", "文件名", "状态", "分类结果")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        # 设置列
        self.tree.heading("选择", text="选择")
        self.tree.heading("文件名", text="文件名")
        self.tree.heading("状态", text="状态")
        self.tree.heading("分类结果", text="分类结果")
        
        self.tree.column("选择", width=50)
        self.tree.column("文件名", width=300)
        self.tree.column("状态", width=100)
        self.tree.column("分类结果", width=100)
        
        # 添加复选框状态字典
        self.checkboxes = {}
        
        # 绑定点击事件
        self.tree.bind('<Button-1>', self._on_click)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
    def _center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def _download_file(self, urls, progress_bar, status_label, name):
        """从多个镜像下载文件，带重试功能"""
        if isinstance(urls, str):
            urls = [urls]
            
        downloads_dir = Path('downloads')
        downloads_dir.mkdir(exist_ok=True)
        
        for url in urls:
            try:
                response = requests.get(url, stream=True)
                response.raise_for_status()
                
                # 获取文件大小
                total_size = int(response.headers.get('content-length', 0))
                
                # 设置进度条最大值
                progress_bar['maximum'] = total_size
                
                # 更新状态标签
                status_label.config(text=f"正在下载 {name}...")
                
                # 准备文件路径
                file_path = downloads_dir / Path(url).name
                
                # 下载文件
                with open(file_path, 'wb') as f, tqdm(
                    desc=name,
                    total=total_size,
                    unit='iB',
                    unit_scale=True,
                    unit_divisor=1024,
                ) as pbar:
                    downloaded = 0
                    for data in response.iter_content(chunk_size=1024):
                        size = f.write(data)
                        downloaded += size
                        pbar.update(size)
                        progress_bar['value'] = downloaded
                        self.root.update()
                        
                return file_path
                
            except Exception as e:
                logger.error(f"从 {url} 下载 {name} 失败: {str(e)}")
                self._add_log(f"从 {url} 下载失败，尝试其他镜像...")
                continue
                
        raise Exception(f"所有镜像下载 {name} 均失败")

    def _check_ffmpeg(self) -> bool:
        """检查 FFmpeg 是否已安装"""
        try:
            ffmpeg_path = Path('C:/ffmpeg/bin/ffmpeg.exe')
            if ffmpeg_path.exists():
                # 测试是否可用
                result = os.system(f'"{ffmpeg_path}" -version')
                if result == 0:
                    self.ffmpeg_label.config(text="FFmpeg: 已安装")
                    return True
            return False
        except Exception:
            return False

    def _check_tesseract(self) -> bool:
        """检查 Tesseract 是否已安装且包含中文支持"""
        try:
            tesseract_path = Path(r'C:\Program Files\Tesseract-OCR\tesseract.exe')
            if not tesseract_path.exists():
                return False
                
            # 测试 Tesseract 是否可用
            result = os.system(f'"{tesseract_path}" --version')
            if result != 0:
                return False
                
            # 检查中文语言包
            self._add_log("检查中文语言包...")
            lang_file = Path(r'C:\Program Files\Tesseract-OCR\tessdata\chi_sim.traineddata')
            if not lang_file.exists():
                self._add_log("未找到中文语言包")
                return False
                
            # 测试中文支持
            test_cmd = f'"{tesseract_path}" --list-langs'
            test_result = os.popen(test_cmd).read()
            if 'chi_sim' not in test_result:
                self._add_log("中文语言包无效")
                return False
                
            self.tesseract_label.config(text="Tesseract: 已安装（含中文支持）")
            return True
            
        except Exception as e:
            self._add_log(f"检查 Tesseract 失败: {str(e)}")
            return False

    def _add_log(self, message: str):
        """添加日志到显示区域"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.update_idletasks()

    def _install_ffmpeg(self):
        """安装 FFmpeg"""
        try:
            self.ffmpeg_label.config(text="FFmpeg: 正在安装...")
            self.ffmpeg_progress['value'] = 0
            self.root.update()
            
            # 检查是否是完整版
            if self.is_full_version:
                # 从本地安装
                deps_dir = Path('dependencies')
                if not deps_dir.exists() and getattr(sys, '_MEIPASS', None):
                    deps_dir = Path(sys._MEIPASS) / 'dependencies'
                    
                ffmpeg_zip = deps_dir / 'ffmpeg.zip'
                if not ffmpeg_zip.exists():
                    raise FileNotFoundError("找不到 FFmpeg 安装文件")
                    
                self._add_log("正在解压 FFmpeg...")
                
                # 解压 FFmpeg
                with zipfile.ZipFile(ffmpeg_zip) as zf:
                    # 获取压缩包中的所有文件
                    total_files = len(zf.filelist)
                    self.ffmpeg_progress['maximum'] = total_files
                    
                    # 创建临时目录
                    temp_dir = Path('temp_ffmpeg')
                    temp_dir.mkdir(exist_ok=True)
                    
                    # 解压文件
                    for i, file in enumerate(zf.filelist):
                        zf.extract(file, temp_dir)
                        self.ffmpeg_progress['value'] = i + 1
                        self.root.update()
                        
                    # 移动 FFmpeg 目录
                    ffmpeg_dir = Path('C:/ffmpeg')
                    if ffmpeg_dir.exists():
                        shutil.rmtree(ffmpeg_dir)
                        
                    ffmpeg_extracted = next(temp_dir.glob('ffmpeg-*'))
                    shutil.move(str(ffmpeg_extracted), str(ffmpeg_dir))
                    
                    # 清理临时目录
                    shutil.rmtree(temp_dir)
                    
            else:
                # 从网络下载安装
                self._add_log("正在下载 FFmpeg...")
                ffmpeg_file = self._download_file(
                    self.FFMPEG_MIRRORS,
                    self.ffmpeg_progress,
                    self.ffmpeg_label,
                    "FFmpeg"
                )
                
                self._add_log("正在解压 FFmpeg...")
                
                # 解压 FFmpeg
                with zipfile.ZipFile(ffmpeg_file) as zf:
                    # 获取压缩包中的所有文件
                    total_files = len(zf.filelist)
                    self.ffmpeg_progress['maximum'] = total_files
                    
                    # 创建临时目录
                    temp_dir = Path('temp_ffmpeg')
                    temp_dir.mkdir(exist_ok=True)
                    
                    # 解压文件
                    for i, file in enumerate(zf.filelist):
                        zf.extract(file, temp_dir)
                        self.ffmpeg_progress['value'] = i + 1
                        self.root.update()
                        
                    # 移动 FFmpeg 目录
                    ffmpeg_dir = Path('C:/ffmpeg')
                    if ffmpeg_dir.exists():
                        shutil.rmtree(ffmpeg_dir)
                        
                    ffmpeg_extracted = next(temp_dir.glob('ffmpeg-*'))
                    shutil.move(str(ffmpeg_extracted), str(ffmpeg_dir))
                    
                    # 清理临时目录
                    shutil.rmtree(temp_dir)
                    
                # 清理下载文件
                ffmpeg_file.unlink()
                
            self._add_log("FFmpeg 安装完成")
            self.ffmpeg_label.config(text="FFmpeg: 安装完成")
            self.ffmpeg_progress['value'] = self.ffmpeg_progress['maximum']
            
        except Exception as e:
            self._add_log(f"FFmpeg 安装失败: {str(e)}")
            self.ffmpeg_label.config(text="FFmpeg: 安装失败")
            raise

    def _install_chinese_language_pack(self):
        """安装中文语言包"""
        try:
            self._add_log("开始安装中文语言包...")
            tessdata_dir = Path(r'C:\Program Files\Tesseract-OCR\tessdata')
            lang_file = tessdata_dir / 'chi_sim.traineddata'
            
            if lang_file.exists():
                self._add_log("中文语言包已存在")
                return True
            
            # 确保tessdata目录存在
            tessdata_dir.mkdir(parents=True, exist_ok=True)
            
            # 下载语言包
            self._add_log("下载中文语言包...")
            for url in self.TESSERACT_CHI_SIM_MIRRORS:
                try:
                    self._add_log(f"尝试从 {url.split('/')[2]} 下载...")
                    response = requests.get(url, stream=True)
                    if response.status_code != 200:
                        continue
                    
                    total_size = int(response.headers.get('content-length', 0))
                    self.tesseract_progress['maximum'] = total_size
                    self.tesseract_progress['value'] = 0
                    
                    with open(lang_file, 'wb') as f:
                        downloaded = 0
                        for chunk in response.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                                downloaded += len(chunk)
                                self.tesseract_progress['value'] = downloaded
                                self.root.update()
                    
                    self._add_log("中文语言包下载完成")
                    return True
                    
                except Exception as e:
                    self._add_log(f"从 {url} 下载失败: {str(e)}")
                    continue
            
            raise Exception("所有镜像下载失败")
            
        except Exception as e:
            self._add_log(f"安装中文语言包失败: {str(e)}")
            return False

    def _install_tesseract(self):
        """安装 Tesseract"""
        try:
            self.tesseract_label.config(text="Tesseract: 正在安装...")
            self.tesseract_progress['value'] = 0
            self.root.update()
            
            # 检查是否是完整版
            if self.is_full_version:
                # 从本地安装
                deps_dir = Path('dependencies')
                if not deps_dir.exists() and getattr(sys, '_MEIPASS', None):
                    deps_dir = Path(sys._MEIPASS) / 'dependencies'
                    
                tesseract_installer = deps_dir / 'tesseract-installer.exe'
                if not tesseract_installer.exists():
                    raise FileNotFoundError("找不到 Tesseract 安装文件")
                    
                self._add_log("正在安装 Tesseract...")
                
                # 运行安装程序
                import subprocess
                process = subprocess.Popen(
                    [str(tesseract_installer), '/S'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                
                # 等待安装完成
                while process.poll() is None:
                    time.sleep(0.1)
                    self.root.update()
                    
                if process.returncode != 0:
                    raise Exception(f"Tesseract 安装失败，返回代码: {process.returncode}")
                    
            else:
                # 从网络下载安装
                self._add_log("正在下载 Tesseract...")
                tesseract_file = self._download_file(
                    self.TESSERACT_MIRRORS,
                    self.tesseract_progress,
                    self.tesseract_label,
                    "Tesseract"
                )
                
                self._add_log("正在安装 Tesseract...")
                
                # 运行安装程序
                import subprocess
                process = subprocess.Popen(
                    [str(tesseract_file), '/S'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                
                # 等待安装完成
                while process.poll() is None:
                    time.sleep(0.1)
                    self.root.update()
                    
                if process.returncode != 0:
                    raise Exception(f"Tesseract 安装失败，返回代码: {process.returncode}")
                    
                # 清理下载文件
                tesseract_file.unlink()
                
            # 安装中文语言包
            self._install_chinese_language_pack()
            
            self._add_log("Tesseract 安装完成")
            self.tesseract_label.config(text="Tesseract: 安装完成")
            self.tesseract_progress['value'] = self.tesseract_progress['maximum']
            
        except Exception as e:
            self._add_log(f"Tesseract 安装失败: {str(e)}")
            self.tesseract_label.config(text="Tesseract: 安装失败")
            raise

    def _start_installation(self):
        self.start_button.config(state='disabled')
        self.status_label.config(text="检查已安装组件...")
        
        def install_thread():
            self._add_log("\n=== 开始检查系统环境 ===")
            success = True
            need_install = False
            tesseract_needs_chinese = False
            
            # 检查 FFmpeg
            self._add_log("\n正在检查 FFmpeg...")
            ffmpeg_path = Path('C:/ffmpeg/bin/ffmpeg.exe')
            if ffmpeg_path.exists():
                try:
                    result = os.system(f'"{ffmpeg_path}" -version')
                    if result == 0:
                        self._add_log("✓ FFmpeg 已安装且可用")
                    else:
                        self._add_log("✗ FFmpeg 安装可能损坏，需要重新安装")
                        need_install = True
                except Exception as e:
                    self._add_log(f"✗ FFmpeg 检测出错: {str(e)}")
                    need_install = True
            else:
                self._add_log("✗ 未找到 FFmpeg，需要安装")
                need_install = True
            
            # 检查 Tesseract
            self._add_log("\n正在检查 Tesseract OCR...")
            tesseract_path = Path(r'C:\Program Files\Tesseract-OCR\tesseract.exe')
            if tesseract_path.exists():
                try:
                    # 检查版本
                    version_result = os.popen(f'"{tesseract_path}" --version').read()
                    self._add_log(f"发现 Tesseract: {version_result.split()[1] if version_result else '未知版本'}")
                    
                    # 检查中文支持
                    self._add_log("检查中文语言包...")
                    test_cmd = f'"{tesseract_path}" --list-langs'
                    test_result = os.popen(test_cmd).read()
                    
                    if 'chi_sim' in test_result:
                        self._add_log("✓ 中文语言包已安装")
                    else:
                        self._add_log("✗ 未找到中文语言包，需要重新安装")
                        tesseract_needs_chinese = True
                        need_install = True
                        
                    # 检查环境变量
                    tesseract_dir = r'C:\Program Files\Tesseract-OCR'
                    if tesseract_dir in os.environ['PATH']:
                        self._add_log("✓ Tesseract 已添加到环境变量")
                    else:
                        self._add_log("! Tesseract 未添加到环境变量，可能影响使用")
                        
                except Exception as e:
                    self._add_log(f"✗ Tesseract 检测出错: {str(e)}")
                    need_install = True
            else:
                self._add_log("✗ 未找到 Tesseract，需要安装")
                need_install = True
            
            self._add_log("\n=== 环境检查完成 ===\n")
            
            # 根据检查结果进行安装
            if need_install:
                self._add_log("开始安装缺失组件...")
                
                # 安装 FFmpeg
                if not self._check_ffmpeg():
                    if not self._install_ffmpeg():
                        success = False
                
                # 处理 Tesseract 安装
                if tesseract_needs_chinese:
                    self._add_log("\n需要重新安装 Tesseract 以添加中文支持")
                    if messagebox.askyesno(
                        "需要中文支持",
                        "检测到 Tesseract 缺少中文支持，需要重新安装以添加中文语言包。\n是否现在安装？"
                    ):
                        if not self._install_tesseract():
                            success = False
                    else:
                        success = False
                        self._add_log("用户取消安装中文支持")
                elif not self._check_tesseract():
                    if not self._install_tesseract():
                        success = False
                
                # 清理下载文件
                shutil.rmtree('downloads', ignore_errors=True)
                
                if success:
                    self._add_log("\n=== 所有组件安装完成 ===")
                    self.status_label.config(text="安装完成！程序将在3秒后启动...")
                    self.root.after(3000, self._start_main_program)
                else:
                    error_msg = "安装失败或被取消。\n\n注意：缺少中文支持可能会影响中文文档的识别。"
                    messagebox.showerror("安装未完成", error_msg)
                    self.status_label.config(text="安装失败，请检查错误信息并重试")
                    self.start_button.config(state='normal')
                    return
            else:
                self._add_log("\n=== 所有组件已正确安装 ===")
                self.status_label.config(text="所有组件已安装，程序将在3秒后启动...")
                self.root.after(3000, self._start_main_program)
        
        threading.Thread(target=install_thread, daemon=True).start()
        
    def _start_main_program(self):
        self.root.destroy()
        # 这里启动主程序
        
    def _on_closing(self):
        """处理窗口关闭事件"""
        if messagebox.askokcancel("退出", "确定要退出安装程序吗？\n退出后程序将无法正常运行。"):
            self.root.quit()
            sys.exit(0)  # 直接退出程序

    def run(self):
        try:
            self.root.mainloop()
        except Exception as e:
            logger.error(f"安装程序运行失败: {str(e)}")
            sys.exit(1)

    def _add_files(self):
        """添加文件到列表"""
        files = filedialog.askopenfilenames(
            title="选择文件",
            filetypes=[
                ("所有支持的文件", "*.txt *.doc *.docx *.pdf *.ppt *.pptx *.jpg *.png *.jpeg *.zip *.rar *.7z *.mp3 *.wav *.m4a"),
                ("文档", "*.txt *.doc *.docx *.pdf *.ppt *.pptx"),  # 添加 ppt 和 pptx
                ("图片", "*.jpg *.png *.jpeg"),
                ("压缩包", "*.zip *.rar *.7z"),
                ("音频", "*.mp3 *.wav *.m4a"),
                ("所有文件", "*.*")
            ]
        )
        
        for file in files:
            file_path = Path(file)
            if not self._is_file_in_tree(file_path):
                item = self.tree.insert('', tk.END, values=('☐', file_path.name, "等待处理", "未分类"))
                self.checkboxes[item] = False
                self.files_status[str(file_path)] = "等待处理"
                self.files_results[str(file_path)] = "未分类"

    def _add_folder(self):
        """添加文件夹中的文件到列表"""
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            folder_path = Path(folder)
            for file_path in folder_path.rglob('*'):
                if file_path.is_file() and not self._is_file_in_tree(file_path):
                    item = self.tree.insert('', tk.END, values=('☐', file_path.name, "等待处理", "未分类"))
                    self.checkboxes[item] = False
                    self.files_status[str(file_path)] = "等待处理"
                    self.files_results[str(file_path)] = "未分类"

    def _is_file_in_tree(self, file_path: Path) -> bool:
        """检查文件是否已在列表中"""
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][1] == file_path.name:
                return True
        return False

    def _on_click(self, event):
        """处理点击事件"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#1":  # 选择列
                item = self.tree.identify_row(event.y)
                if item:
                    status = self.tree.item(item)['values'][2]
                    if status != "已完成":  # 只允许选择未完成的文件
                        self.checkboxes[item] = not self.checkboxes[item]
                        self.tree.set(item, "选择", '☑' if self.checkboxes[item] else '☐')

    def _select_all(self):
        """全选"""
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][2] != "已完成":
                self.checkboxes[item] = True
                self.tree.set(item, "选择", '☑')

    def _deselect_all(self):
        """取消全选"""
        for item in self.tree.get_children():
            self.checkboxes[item] = False
            self.tree.set(item, "选择", '☐')

    def _invert_selection(self):
        """反选"""
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][2] != "已完成":
                self.checkboxes[item] = not self.checkboxes[item]
                self.tree.set(item, "选择", '☑' if self.checkboxes[item] else '☐')

    def _select_by_type(self, file_type: str):
        """根据文件类型选择"""
        extensions = {
            'doc': ('.txt', '.doc', '.docx', '.pdf', '.ppt', '.pptx'),  # 添加 ppt 和 pptx
            'archive': ('.zip', '.rar', '.7z'),
            'audio': ('.mp3', '.wav', '.m4a')
        }
        
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            if values[2] != "已完成":  # 只选择未完成的文件
                file_name = values[1]
                ext = Path(file_name).suffix.lower()
                if ext in extensions[file_type]:
                    self.checkboxes[item] = True
                    self.tree.set(item, "选择", '☑')

    def _update_file_status(self, filename: str, status: str, result: str):
        """更新文件状态"""
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][1] == filename:
                self.tree.item(item, values=('☐', filename, status, result))
                if status == "已完成":
                    # 使用灰色标记已完成的文件
                    self.tree.tag_configure('completed', foreground='gray')
                    self.tree.item(item, tags=('completed',))
                    # 取消选择并禁用复选框
                    self.checkboxes[item] = False
                break

def check_first_run() -> bool:
    """检查是否首次运行"""
    # 检查安装标记文件
    if Path('.installed').exists():
        return False
        
    # 检查配置文件
    if not Path('config.conf').exists():
        return True
        
    # 检查必要的目录
    config = configparser.ConfigParser()
    config.read('config.conf', encoding='utf-8')
    
    source_dir = Path(config['Paths']['source_folder'])
    target_dir = Path(config['Paths']['target_folder'])
    
    if not source_dir.exists() or not target_dir.exists():
        return True
    
    return False

def main():
    try:
        # 配置日志
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        logger.info("启动安装程序...")
        
        if check_first_run():
            logger.info("需要进行初始化安装")
            setup = SetupWindow()
            setup.run()
        else:
            logger.info("无需安装，直接启动主程序")
            
    except Exception as e:
        logger.error(f"程序执行失败: {str(e)}", exc_info=True)
        messagebox.showerror("错误", f"程序执行失败: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 