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
from concurrent.futures import ThreadPoolExecutor
import json
import py7zr
import hashlib

logger = logging.getLogger(__name__)

class SetupWindow:
    # 主要下载地址
    FFMPEG_URL = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
    TESSERACT_URL = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe"
    
    # GitHub 镜像加速服务
    GITHUB_MIRRORS = [
        'https://bgithub.xyz',
        'https://kkgithub.com',
        'https://gitclone.com',
        'https://github.ur1.fun',
        'https://moeyy.cn/gh-proxy/',
        'https://ghp.ci/',
        'https://gh-proxy.com/',
        'https://ghproxy.net/',
        'https://ghproxy.homeboyc.cn/',
        'http://toolwa.com/github/'
    ]
    
    # 更新中文语言包下载地址
    TESSERACT_CHI_SIM_MIRRORS = [
        'https://raw.fastgit.org/tesseract-ocr/tessdata/4.1.0/chi_sim.traineddata',
        'https://cdn.jsdelivr.net/gh/tesseract-ocr/tessdata@4.1.0/chi_sim.traineddata',
        'https://ghproxy.net/https://raw.githubusercontent.com/tesseract-ocr/tessdata/4.1.0/chi_sim.traineddata',
        'https://gitee.com/mirrors/tesseract/raw/master/tessdata/chi_sim.traineddata',
    ]
    
    # 更新镜像地址
    mirrors = {
        'default': {
            'tesseract': [
                'https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe',
                'https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.1.20230401/tesseract-ocr-w64-setup-5.3.1.20230401.exe'
            ],
            'ffmpeg': 'https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip',
        },
        'china': {
            'tesseract': [
                # 国内镜像
                'https://download.fastgit.org/UB-Mannheim/tesseract/releases/download/v5.3.1.20230401/tesseract-ocr-w64-setup-5.3.1.20230401.exe',
                'https://hub.fastgit.xyz/UB-Mannheim/tesseract/releases/download/v5.3.1.20230401/tesseract-ocr-w64-setup-5.3.1.20230401.exe',
                'https://ghproxy.net/https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.1.20230401/tesseract-ocr-w64-setup-5.3.1.20230401.exe',
                # 备用下载地址
                'https://raw.fastgit.org/UB-Mannheim/tesseract/master/installer/tesseract-ocr-w64-setup-5.3.1.20230401.exe',
                'https://cdn.jsdelivr.net/gh/UB-Mannheim/tesseract@master/installer/tesseract-ocr-w64-setup-5.3.1.20230401.exe',
                # 百度网盘镜像（如果有的话）
                'https://pan.baidu.com/s/1xxx',  # 需要替换为实际的分享链接
                # 阿里云 OSS 镜像（如果有的话）
                'https://your-bucket.oss-cn-hangzhou.aliyuncs.com/tesseract-ocr-w64-setup-5.3.1.20230401.exe',
            ],
            'ffmpeg': [
                'https://gitee.com/mirrors/ffmpeg/raw/master/ffmpeg-master-latest-win64-gpl.zip',
                'https://mirrors.cloud.tencent.com/ffmpeg/ffmpeg-master-latest-win64-gpl.zip',
                'https://mirrors.aliyun.com/ffmpeg/ffmpeg-master-latest-win64-gpl.zip',
            ],
            'pip': 'https://pypi.tuna.tsinghua.edu.cn/simple'
        }
    }
    
    # 添加文件校验和
    FILE_CHECKSUMS = {
        'tesseract': {
            'tesseract-ocr-w64-setup-5.3.1.20230401.exe': 'a7d4c69e5b336c4f19eb3cca6c4a2d558fed685b9d679c60393886b4b8641481',
            'chi_sim.traineddata': '06eac7a56c20f1f66889d65a3d7c2e9871f0d3fea0b683475a4c8c21c35ca646',
        },
        'ffmpeg': {
            'ffmpeg-master-latest-win64-gpl.zip': None  # 动态文件，每次构建的哈希值不同
        }
    }
    
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
        
        self._init_ui()
        self._center_window()
        
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
        desc_label = ttk.Label(
            main_frame,
            text="首次运行需要下载并安装必要的组件，请保持网络连接",
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
        
    def _center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def _verify_file(self, file_path: Path, expected_hash: str = None, file_type: str = None) -> bool:
        """验证文件完整性"""
        if not file_path.exists():
            return False
            
        if not expected_hash and file_type == 'ffmpeg':
            # FFmpeg 是动态构建的，只检查文件大小
            min_size = 100 * 1024 * 1024  # 最小 100MB
            return file_path.stat().st_size >= min_size
            
        try:
            sha256_hash = hashlib.sha256()
            with open(file_path, "rb") as f:
                for byte_block in iter(lambda: f.read(4096), b""):
                    sha256_hash.update(byte_block)
            actual_hash = sha256_hash.hexdigest()
            
            if actual_hash != expected_hash:
                self._add_log(f"文件校验失败: {file_path.name}")
                self._add_log(f"预期哈希值: {expected_hash}")
                self._add_log(f"实际哈希值: {actual_hash}")
                return False
                
            self._add_log(f"文件校验成功: {file_path.name}")
            return True
            
        except Exception as e:
            self._add_log(f"文件校验出错: {str(e)}")
            return False

    def _download_file(self, urls, progress_bar, status_label, name):
        """从最快的镜像下载文件，并验证完整性"""
        try:
            url = self.select_fastest_mirror(urls, name)
            file_path = self._download_from_url(url, progress_bar, status_label, name)
            
            # 获取文件类型和预期哈希值
            file_type = name.lower()
            file_name = file_path.name
            expected_hash = self.FILE_CHECKSUMS.get(file_type, {}).get(file_name)
            
            # 验证文件完整性
            self._add_log(f"正在验证 {name} 文件完整性...")
            if not self._verify_file(file_path, expected_hash, file_type):
                # 如果验证失败，删除文件并重试其他镜像
                file_path.unlink()
                raise Exception(f"{name} 文件校验失败")
                
            return file_path
            
        except Exception as e:
            self._add_log(f"下载 {name} 失败: {str(e)}")
            raise

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
        try:
            if self._check_ffmpeg():
                self._add_log("检测到 FFmpeg 已安装")
                self.ffmpeg_progress['value'] = 100
                return True

            self._add_log("开始安装 FFmpeg...")
            self.ffmpeg_label.config(text="FFmpeg: 准备下载...")
            
            # 选择最佳镜像源
            mirror = self.select_best_mirror()
            url = self.mirrors[mirror]['ffmpeg']
            
            # 下载文件
            zip_path = self._download_file(
                url,
                self.ffmpeg_progress,
                self.ffmpeg_label,
                "FFmpeg"
            )
            self._add_log("FFmpeg 下载完成，开始解压安装...")
            
            # 安装过程
            self.ffmpeg_label.config(text="FFmpeg: 正在安装...")
            ffmpeg_dir = Path('C:/ffmpeg')
            if ffmpeg_dir.exists():
                self._add_log("删除旧版本 FFmpeg...")
                shutil.rmtree(ffmpeg_dir)
            
            # 使用 zipfile 解压 zip 文件
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                root_dir = zip_ref.namelist()[0].split('/')[0]
                self._add_log("解压 FFmpeg 文件...")
                zip_ref.extractall('C:/')
                
                extracted_path = Path(f'C:/{root_dir}')
                if extracted_path != ffmpeg_dir:
                    self._add_log("重命名 FFmpeg 目录...")
                    extracted_path.rename(ffmpeg_dir)
            
            # 环境变量设置
            ffmpeg_bin = str(ffmpeg_dir / 'bin')
            if ffmpeg_bin not in os.environ['PATH']:
                self._add_log("添加 FFmpeg 到环境变量...")
                os.environ['PATH'] += os.pathsep + ffmpeg_bin
            
            # 更新配置
            self._add_log("更新配置文件...")
            config = configparser.ConfigParser()
            config.read('config.conf', encoding='utf-8')
            config['Audio']['ffmpeg_path'] = str(Path('C:/ffmpeg/bin/ffmpeg.exe'))
            with open('config.conf', 'w', encoding='utf-8') as f:
                config.write(f)
            
            self._add_log("FFmpeg 安装完成！")
            self.ffmpeg_label.config(text="FFmpeg: 安装完成")
            return True
            
        except Exception as e:
            error_msg = f"FFmpeg 安装失败: {str(e)}"
            self._add_log(error_msg)
            self.ffmpeg_label.config(text=f"FFmpeg: 安装失败 - {str(e)}")
            return False
            
    def _install_chinese_language_pack(self):
        """安装中文语言包（添加校验）"""
        try:
            self._add_log("开始安装中文语言包...")
            tessdata_dir = Path(r'C:\Program Files\Tesseract-OCR\tessdata')
            lang_file = tessdata_dir / 'chi_sim.traineddata'
            
            if lang_file.exists():
                # 验证已存在的文件
                if self._verify_file(
                    lang_file,
                    self.FILE_CHECKSUMS['tesseract']['chi_sim.traineddata']
                ):
                    self._add_log("中文语言包已存在且验证通过")
                    return True
                else:
                    self._add_log("现有中文语言包验证失败，将重新下载")
                    lang_file.unlink()
            
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
        try:
            if self._check_tesseract():
                self._add_log("检测到 Tesseract 已安装")
                self.tesseract_progress['value'] = 100
                return True

            self._add_log("开始安装 Tesseract...")
            self.tesseract_label.config(text="Tesseract: 准备下载...")
            
            # 下载安装程序
            exe_path = self._download_file(
                self.mirrors['china']['tesseract'],
                self.tesseract_progress,
                self.tesseract_label,
                "Tesseract"
            )
            
            # 直接使用手动安装
            self._add_log("准备安装 Tesseract 主程序...")
            self._add_log("1. 安装位置保持默认：C:\\Program Files\\Tesseract-OCR")
            self._add_log("2. 确保勾选'Add Tesseract to PATH'选项")
            
            # 显示安装提示对话框
            messagebox.showinfo(
                "安装说明", 
                "即将打开 Tesseract 安装程序，请：\n\n"
                "1. 使用默认安装位置\n"
                "2. 勾选'Add Tesseract to PATH'\n\n"
                "完成上述步骤后点击安装。"
            )
            
            # 运行安装程序
            self._add_log("启动安装程序...")
            result = os.system(f'"{exe_path}"')
            
            if result != 0:
                raise Exception("安装程序异常退出")
            
            # 等待安装完成
            self._add_log("等待安装完成...")
            tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            
            for i in range(60):
                if Path(tesseract_path).exists():
                    self._add_log("检测到 Tesseract 已安装")
                    break
                self._add_log(f"等待安装完成... {i+1}/60")
                time.sleep(1)
                self.root.update()
            
            if not Path(tesseract_path).exists():
                raise Exception("未检测到 Tesseract 安装，请确认是否安装成功")
            
            # 测试安装
            self._add_log("测试 Tesseract 安装...")
            test_result = os.system(f'"{tesseract_path}" --version')
            if test_result != 0:
                raise Exception("Tesseract 安装测试失败")
            
            # 安装中文语言包
            self._add_log("\n开始安装中文语言包...")
            if not self._install_chinese_language_pack():
                if not messagebox.askyesno(
                    "继续安装", 
                    "中文语言包安装失败，这可能会影响中文文档的识别。\n是否继续？"
                ):
                    raise Exception("用户取消安装")
                self._add_log("用户选择继续，但中文支持可能不可用")
            else:
                self._add_log("中文语言包安装成功")
            
            # 更新配置
            self._add_log("更新配置文件...")
            config = configparser.ConfigParser()
            config.read('config.conf', encoding='utf-8')
            config['OCR']['tesseract_path'] = tesseract_path
            with open('config.conf', 'w', encoding='utf-8') as f:
                config.write(f)
            
            self._add_log("Tesseract 安装完成！")
            self.tesseract_label.config(text="Tesseract: 安装完成")
            return True
            
        except Exception as e:
            error_msg = f"Tesseract 安装失败: {str(e)}"
            self._add_log(error_msg)
            self.tesseract_label.config(text=f"Tesseract: 安装失败 - {str(e)}")
            return False
            
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

    def check_connection(self, url, timeout=5):
        """检查连接速度"""
        try:
            start_time = time.time()
            response = requests.get(url, timeout=timeout, stream=True)
            if response.status_code == 200:
                # 只下载前 8KB 来测试速度
                for _ in response.iter_content(chunk_size=8192):
                    break
                elapsed = time.time() - start_time
                return elapsed
        except Exception as e:
            self._add_log(f"连接测试失败: {url}, 错误: {str(e)}")
            return float('inf')
        return float('inf')

    def check_is_china_ip(self):
        """检查是否为中国 IP"""
        try:
            # 使用 ipapi.co 检查 IP 归属地
            response = requests.get('https://ipapi.co/json/', timeout=5)
            if response.status_code == 200:
                data = response.json()
                country_code = data.get('country_code')
                self._add_log(f"检测到 IP 归属地: {data.get('country_name', '未知')}")
                return country_code == 'CN'
        except Exception as e:
            self._add_log(f"IP 归属地检测失败: {str(e)}")
            
            # 备用方案：使用 cip.cc
            try:
                response = requests.get('http://cip.cc', timeout=5)
                if response.status_code == 200 and '中国' in response.text:
                    self._add_log("备用检测确认为中国 IP")
                    return True
            except Exception as e2:
                self._add_log(f"备用 IP 检测也失败: {str(e2)}")
        
        return False

    def select_best_mirror(self):
        """选择最佳镜像源"""
        self._add_log("正在检测网络环境...")
        
        # 首先检查是否为中国 IP
        if self.check_is_china_ip():
            self._add_log("检测到中国 IP，将优先使用国内镜像")
            return 'china'
        
        # 如果不是中国 IP，测试镜像速度
        self._add_log("正在测试镜像源连接速度...")
        test_urls = {
            'default': 'https://github.com',
            'china': 'https://gitee.com'
        }
        
        results = {}
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {executor.submit(self.check_connection, url): name 
                      for name, url in test_urls.items()}
            
            for future in futures:
                name = futures[future]
                try:
                    speed = future.result()
                    results[name] = speed
                    self._add_log(f"镜像源 {name} 响应时间: {speed:.2f}秒")
                except Exception as e:
                    self._add_log(f"测试镜像源 {name} 失败: {str(e)}")
                    results[name] = float('inf')
        
        # 选择响应最快的镜像源
        best_mirror = min(results.items(), key=lambda x: x[1])[0]
        self._add_log(f"选择镜像源: {best_mirror}")
        return best_mirror

    def download_dependencies(self):
        """下载依赖"""
        # 选择最佳镜像源
        mirror = self.select_best_mirror()
        urls = self.mirrors[mirror]
        
        # 如果使用中国镜像，设置 pip 镜像
        if mirror == 'china':
            self._add_log("使用中国镜像源下载依赖")
            import subprocess
            try:
                subprocess.run([
                    'pip', 'config', 'set', 'global.index-url',
                    'https://pypi.tuna.tsinghua.edu.cn/simple'
                ], check=True)
                self._add_log("已设置 pip 镜像源")
            except Exception as e:
                self._add_log(f"设置 pip 镜像源失败: {str(e)}")
        
        # 下载 Tesseract
        self._download_file(urls['tesseract'], self.tesseract_progress, self.tesseract_label, "Tesseract")
        
        # 下载 FFmpeg
        self._download_file(urls['ffmpeg'], self.ffmpeg_progress, self.ffmpeg_label, "FFmpeg")
        
        # ... 其他下载和安装步骤 ...

    def download_file(self, url, filename):
        """带进度条的文件下载"""
        try:
            response = requests.get(url, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            
            with open(filename, 'wb') as f, tqdm(
                desc=f"下载 {filename}",
                total=total_size,
                unit='iB',
                unit_scale=True
            ) as pbar:
                for data in response.iter_content(chunk_size=8192):
                    size = f.write(data)
                    pbar.update(size)
                    
        except Exception as e:
            self._add_log(f"下载 {filename} 失败: {str(e)}")
            raise

    def select_fastest_mirror(self, urls, name="文件"):
        """选择最快的镜像"""
        if isinstance(urls, str):
            urls = [urls]
            
        self._add_log(f"正在测试 {name} 的下载速度...")
        results = []
        
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {executor.submit(self.check_connection, url): url for url in urls}
            
            for future in futures:
                url = futures[future]
                try:
                    speed = future.result()
                    if speed != float('inf'):
                        results.append((url, speed))
                        self._add_log(f"镜像 {url.split('/')[2]} 响应时间: {speed:.2f}秒")
                except Exception as e:
                    self._add_log(f"测试镜像 {url.split('/')[2]} 失败: {str(e)}")
        
        if not results:
            raise Exception(f"所有 {name} 镜像都无法访问")
            
        # 选择最快的镜像
        fastest_url = min(results, key=lambda x: x[1])[0]
        self._add_log(f"选择最快的镜像: {fastest_url.split('/')[2]}")
        return fastest_url

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