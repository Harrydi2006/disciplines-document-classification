import configparser
import os
import logging
import json
import threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import tqdm
import requests
from typing import Optional
import sys
from utils.logger import setup_logger
import time
import shutil
from tkinter import messagebox
import subprocess

# 添加必要的路径
if getattr(sys, 'frozen', False):
    # 运行于 PyInstaller 打包后的环境
    base_path = sys._MEIPASS
    if base_path not in sys.path:
        sys.path.insert(0, base_path)
else:
    # 运行于开发环境
    base_path = os.path.dirname(os.path.abspath(__file__))
    if base_path not in sys.path:
        sys.path.insert(0, base_path)

# 设置日志记录器
logger = setup_logger('file_classifier', 'file_classifier.log')

class APIError(Exception):
    """API调用错误的自定义异常"""
    pass

class FileClassifier:
    def __init__(self):
        logger.info("初始化文件分类器")
        
        # 先创建默认配置（如果需要）
        self._create_default_config()
        
        try:
            # 加载配置
            self.config = self._load_config()
            self.subjects = ['语文', '数学', '英语', '物理', '化学', '生物', '未知']
            
            # 在初始化时检查环境
            if not self._check_environment():
                raise EnvironmentError("环境检查失败，请重新运行安装程序")
            
            self._setup_folders()
            
        except Exception as e:
            logger.error(f"初始化失败: {str(e)}")
            # 显示错误对话框
            messagebox.showerror("初始化失败", 
                "程序初始化失败，可能是因为：\n"
                "1. 配置文件不存在或损坏\n"
                "2. 必要的文件夹无法创建\n"
                "3. 外部程序未正确安装\n\n"
                "程序将重新进行首次运行设置。"
            )
            
            # 删除可能损坏的配置文件
            if Path('config.conf').exists():
                Path('config.conf').unlink()
            
            # 重新创建配置
            self._create_default_config()
            
            # 重新加载配置
            self.config = self._load_config()
            self.subjects = ['语文', '数学', '英语', '物理', '化学', '生物', '未知']
            self._setup_folders()

    def _create_default_config(self):
        """如果配置文件不存在，创建默认配置文件"""
        if not Path('config.conf').exists():
            config = configparser.ConfigParser()
            
            # 设置默认配置
            config['API'] = {
                'host': '',  # 留空，等待用户填写
                'api_key': ''  # 留空，等待用户填写
            }
            
            config['Model'] = {
                'model_name': ''  # 留空，等待用户填写
            }
            
            config['Paths'] = {
                'source_folder': '',  # 留空，等待用户选择
                'target_folder': ''   # 留空，等待用户选择
            }
            
            config['Prompt'] = {
                'classification_prompt': '请判断以下内容属于哪个学科（语文、数学、英语、物理、化学、生物）？如果无法判断，请回答"未知"。内容：'
            }
            
            config['Features'] = {
                'enable_ocr': 'true',
                'enable_audio': 'true',
                'enable_archive': 'true'
            }
            
            config['Threading'] = {
                'max_workers': '4'
            }
            
            config['OCR'] = {
                'tesseract_path': r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            }
            
            config['Audio'] = {
                'ffmpeg_path': r'C:\ffmpeg\bin\ffmpeg.exe'
            }
            
            # 创建配置文件
            with open('config.conf', 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            
            # 显示首次运行提示
            messagebox.showinfo("首次运行提示", 
                "欢迎使用文件分类助手！\n\n"
                "首次运行需要完成以下设置：\n"
                "1. 配置 API 信息（主机地址、密钥和模型）\n"
                "2. 设置源文件夹和目标文件夹\n"
                "3. 安装必要的外部程序\n\n"
                "点击确定开始配置。"
            )
            
            # 启动安装程序
            from gui.setup_window import SetupWindow
            setup = SetupWindow()
            setup.run()

    def _load_config(self) -> configparser.ConfigParser:
        config = configparser.ConfigParser()
        try:
            if not config.read('config.conf', encoding='utf-8'):
                raise FileNotFoundError("配置文件不存在")
            return config
        except Exception as e:
            logger.error(f"配置文件读取失败: {str(e)}", exc_info=True)
            raise

    def _setup_folders(self):
        """设置必要的文件夹结构"""
        try:
            # 检查路径设置
            source_path = self.config['Paths']['source_folder']
            target_path = self.config['Paths']['target_folder']
            
            if not source_path or not target_path:
                logger.warning("源文件夹或目标文件夹未设置")
                messagebox.showwarning(
                    "路径未设置",
                    "请先设置源文件夹和目标文件夹路径。\n"
                    "可以在主界面的路径设置中选择合适的目录。"
                )
                return False
            
            # 创建基本目录
            base_dir = Path(target_path)
            source_dir = Path(source_path)
            
            # 创建源文件夹
            source_dir.mkdir(parents=True, exist_ok=True)
            
            # 创建目标文件夹及其子目录
            for subject in self.subjects:
                subject_dir = base_dir / subject
                subject_dir.mkdir(parents=True, exist_ok=True)
            
            logger.info("文件夹结构已设置完成")
            return True
            
        except Exception as e:
            logger.error(f"设置文件夹结构失败: {str(e)}")
            messagebox.showerror(
                "错误",
                f"创建文件夹结构失败：\n{str(e)}\n\n"
                "请确保有足够的权限创建文件夹。"
            )
            return False

    def classify_file(self, file_path: Path) -> str:
        try:
            # 首先通过文件名判断
            subject = self._classify_by_filename(file_path)
            if subject != '未知':
                return subject

            # 根据文件类型选择处理方法
            if file_path.suffix.lower() in ['.txt', '.doc', '.docx', '.pdf']:
                subject = self._classify_by_content(file_path)
            elif file_path.suffix.lower() in ['.jpg', '.png', '.jpeg'] and \
                 self.config.getboolean('Features', 'enable_ocr'):
                subject = self._classify_by_ocr(file_path)
            elif file_path.suffix.lower() in ['.zip', '.rar', '.7z'] and \
                 self.config.getboolean('Features', 'enable_archive'):
                subject = self._classify_archive(file_path)
            elif file_path.suffix.lower() in ['.mp3', '.wav', '.m4a'] and \
                 self.config.getboolean('Features', 'enable_audio'):
                subject = self._classify_audio(file_path)
            
            return subject or '未知'
        except Exception as e:
            logger.error(f"文件分类失败 {file_path}: {str(e)}", exc_info=True)
            return '未知'

    def _call_api(self, content: str) -> str:
        try:
            headers = {
                'Authorization': f"Bearer {self.config['API']['api_key']}",
                'Content-Type': 'application/json'
            }
            
            data = {
                "model": self.config['Model']['model_name'],
                "messages": [
                    {"role": "system", "content": "你是一个文件分类助手。请从以下选项中选择一个：语文、数学、英语、物理、化学、生物、未知"},
                    {"role": "user", "content": f"{self.config['Prompt']['classification_prompt']}{content}"}
                ]
            }

            response = requests.post(
                f"{self.config['API']['host']}/chat/completions",
                headers=headers,
                json=data,
                timeout=30
            )
            
            if response.status_code != 200:
                error_msg = f"API调用失败: HTTP {response.status_code} - {response.text}"
                logger.error(error_msg)
                raise APIError(error_msg)

            try:
                response_data = response.json()
                logger.debug(f"API响应: {response_data}")
                
                if 'choices' not in response_data or not response_data['choices']:
                    raise APIError("API响应中没有choices字段")
                
                message_content = response_data['choices'][0]['message'].get('content', '')
                if not message_content:
                    raise APIError("API响应中没有content字段")
                
                # 从响应内容中提取学科
                for subject in self.subjects:
                    if subject in message_content:
                        return subject
                
                # 如果没有找到匹配的学科，返回未知
                return '未知'
                
            except json.JSONDecodeError as e:
                error_msg = f"API响应解析失败: {str(e)} - {response.text}"
                logger.error(error_msg)
                raise APIError(error_msg)
            except KeyError as e:
                error_msg = f"API响应格式错误: {str(e)} - {response.text}"
                logger.error(error_msg)
                raise APIError(error_msg)

        except requests.RequestException as e:
            error_msg = f"API请求异常: {str(e)}"
            logger.error(error_msg)
            raise APIError(error_msg)
        except Exception as e:
            error_msg = f"API调用过程中发生未知错误: {str(e)}"
            logger.error(error_msg)
            raise APIError(error_msg)

    def process_files(self):
        try:
            logger.info("开始处理文件")
            source_folder = Path(self.config['Paths']['source_folder'])
            if not source_folder.exists():
                logger.error(f"源文件夹不存在: {source_folder}")
                return

            files = list(source_folder.rglob('*'))
            files = [f for f in files if f.is_file()]
            
            if not files:
                logger.info("没有找到需要处理的文件")
                return

            # 创建分类缓存
            self._classification_cache = {}

            logger.info(f"开始处理 {len(files)} 个文件")
            with ThreadPoolExecutor(max_workers=int(self.config['Threading']['max_workers'])) as executor:
                with tqdm.tqdm(total=len(files), desc="处理进度") as pbar:
                    # 先进行分类
                    for file_path in files:
                        try:
                            subject = self.classify_file(file_path)
                            self._classification_cache[str(file_path)] = subject
                            logger.info(f"文件 {file_path.name} 分类为: {subject}")
                            
                            # 直接移动文件，不再调用 process_single_file
                            target_dir = Path(self.config['Paths']['target_folder']) / subject
                            target_path = target_dir / file_path.name
                            
                            # 确保目标目录存在
                            target_dir.mkdir(parents=True, exist_ok=True)
                            
                            # 如果目标文件已存在，添加序号
                            counter = 1
                            while target_path.exists():
                                new_name = f"{file_path.stem}_{counter}{file_path.suffix}"
                                target_path = target_dir / new_name
                                counter += 1
                            
                            # 移动文件
                            shutil.move(str(file_path), str(target_path))
                            logger.info(f"文件 {file_path.name} 已成功移动到 {subject} 目录")
                            
                        except Exception as e:
                            logger.error(f"处理失败 {file_path}: {str(e)}", exc_info=True)
                            self._classification_cache[str(file_path)] = '未知'
                        pbar.update(1)

            logger.info("文件处理完成")

        except Exception as e:
            logger.error(f"处理文件过程中发生错误: {str(e)}", exc_info=True)
            raise

    def process_single_file(self, file_path: Path) -> str:
        """处理单个文件（为了保持兼容性）"""
        try:
            logger.info(f"开始处理文件: {file_path.name}")
            
            # 从缓存获取分类结果
            if not hasattr(self, '_classification_cache'):
                self._classification_cache = {}
            
            # 检查缓存
            if str(file_path) not in self._classification_cache:
                # 先进行分类
                subject = self.classify_file(file_path)
                self._classification_cache[str(file_path)] = subject
                logger.info(f"文件 {file_path.name} 分类为: {subject}")
            else:
                # 使用缓存的分类结果
                subject = self._classification_cache[str(file_path)]
                logger.info(f"使用缓存的分类结果: {subject}")
            
            # 移动文件
            target_dir = Path(self.config['Paths']['target_folder']) / subject
            target_path = target_dir / file_path.name
            
            logger.info(f"准备移动文件 {file_path.name} 到 {target_dir}")
            
            # 确保目标目录存在
            target_dir.mkdir(parents=True, exist_ok=True)
            
            # 如果目标文件已存在，添加序号
            counter = 1
            while target_path.exists():
                new_name = f"{file_path.stem}_{counter}{file_path.suffix}"
                target_path = target_dir / new_name
                counter += 1
            
            # 使用 shutil.move 而不是 rename，以支持跨设备移动
            shutil.move(str(file_path), str(target_path))
            logger.info(f"文件 {file_path.name} 已成功移动到 {subject} 目录")
            return subject
            
        except Exception as e:
            logger.error(f"处理文件失败 {file_path}: {str(e)}", exc_info=True)
            raise

    def _move_file(self, file_path: Path):
        """移动文件到目标目录"""
        try:
            # 从缓存获取分类结果
            subject = self._classification_cache.get(str(file_path), '未知')
            
            target_dir = Path(self.config['Paths']['target_folder']) / subject
            target_path = target_dir / file_path.name
            
            logger.info(f"准备移动文件 {file_path.name} 到 {target_dir}")
            
            # 确保目标目录存在
            target_dir.mkdir(parents=True, exist_ok=True)
            
            # 如果目标文件已存在，添加序号
            counter = 1
            while target_path.exists():
                new_name = f"{file_path.stem}_{counter}{file_path.suffix}"
                target_path = target_dir / new_name
                counter += 1
            
            # 使用 shutil.move 而不是 rename，以支持跨设备移动
            shutil.move(str(file_path), str(target_path))
            logger.info(f"文件 {file_path.name} 已成功移动到 {subject} 目录")
            return subject
            
        except Exception as e:
            logger.error(f"移动文件失败 {file_path}: {str(e)}", exc_info=True)
            raise

    def _classify_by_filename(self, file_path: Path) -> str:
        try:
            filename = file_path.stem.lower()
            content = f"文件名：{filename}"
            return self._call_api(content)
        except Exception as e:
            logger.error(f"文件名分类失败 {file_path}: {str(e)}", exc_info=True)
            return '未知'

    def _classify_by_content(self, file_path: Path) -> str:
        try:
            content = ""
            if file_path.suffix.lower() == '.txt':
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read(2000)  # 读取前2000个字符
            elif file_path.suffix.lower() in ['.doc', '.docx']:
                from docx import Document
                doc = Document(file_path)
                content = '\n'.join([p.text for p in doc.paragraphs][:10])  # 读取前10段
            elif file_path.suffix.lower() == '.pdf':
                from PyPDF2 import PdfReader
                reader = PdfReader(file_path)
                content = reader.pages[0].extract_text()[:2000]  # 读取第一页前2000字符

            # 清理文本
            import re
            content = re.sub(r'\s+', ' ', content)  # 替换多个空白字符为单个空格
            content = re.sub(r'[^\w\s\u4e00-\u9fff]', '', content)  # 只保留中文、英文、数字和空格
            
            if content:
                return self._call_api(content[:500])  # 只使用前500字符
            return '未知'
        except Exception as e:
            logger.error(f"内容分类失败 {file_path}: {str(e)}", exc_info=True)
            return '未知'

    def _classify_by_ocr(self, file_path: Path) -> str:
        try:
            import pytesseract
            from PIL import Image
            import subprocess
            
            logger.info(f"开始OCR识别: {file_path.name}")
            
            # 设置 Tesseract 路径
            tesseract_path = self.config['OCR']['tesseract_path']
            logger.info(f"使用 Tesseract 路径: {tesseract_path}")
            
            # 打开并预处理图像
            image = Image.open(file_path)
            
            # 使用 subprocess 直接调用 Tesseract
            try:
                result = subprocess.run(
                    [
                        str(tesseract_path),
                        'stdin',
                        'stdout',
                        '-l', 'chi_sim'
                    ],
                    input=image.tobytes(),
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW  # 隐藏窗口
                )
                
                if result.returncode == 0 and result.stdout:
                    text = result.stdout
                    # 清理OCR文本
                    text = ' '.join(text.split())
                    logger.info(f"OCR识别结果: {text[:200]}...")  # 只显示前200个字符
                    return self._call_api(text[:500])
                else:
                    logger.warning(f"OCR识别结果为空: {file_path.name}")
                    
            except Exception as e:
                logger.error(f"OCR识别失败: {str(e)}")
                raise
            
            return '未知'
            
        except Exception as e:
            logger.error(f"OCR分类失败 {file_path}: {str(e)}", exc_info=True)
            return '未知'

    def _classify_archive(self, file_path: Path) -> str:
        try:
            import py7zr
            import rarfile
            import zipfile
            import tempfile
            
            # 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                
                # 根据文件类型解压
                if file_path.suffix.lower() == '.zip':
                    with zipfile.ZipFile(file_path) as zip_ref:
                        file_list = zip_ref.namelist()
                        # 先尝试用文件名列表判断
                        content = "压缩包内文件：" + ", ".join(file_list)
                        result = self._call_api(content)
                        if result != '未知':
                            return result
                            
                        # 如果无法判断，解压第一个文本文件
                        for name in file_list:
                            if name.lower().endswith(('.txt', '.doc', '.docx', '.pdf')):
                                zip_ref.extract(name, temp_path)
                                return self._classify_by_content(temp_path / name)
                
                elif file_path.suffix.lower() == '.7z':
                    with py7zr.SevenZipFile(file_path, mode='r') as z:
                        file_list = z.getnames()
                        content = "压缩包内文件：" + ", ".join(file_list)
                        result = self._call_api(content)
                        if result != '未知':
                            return result
                            
                        for name in file_list:
                            if name.lower().endswith(('.txt', '.doc', '.docx', '.pdf')):
                                z.extract(temp_path, [name])
                                return self._classify_by_content(temp_path / name)
                
                elif file_path.suffix.lower() == '.rar':
                    with rarfile.RarFile(file_path) as rf:
                        file_list = rf.namelist()
                        content = "压缩包内文件：" + ", ".join(file_list)
                        result = self._call_api(content)
                        if result != '未知':
                            return result
                            
                        for name in file_list:
                            if name.lower().endswith(('.txt', '.doc', '.docx', '.pdf')):
                                rf.extract(name, temp_path)
                                return self._classify_by_content(temp_path / name)
            
            return '未知'
        except Exception as e:
            logger.error(f"压缩包分类失败 {file_path}: {str(e)}", exc_info=True)
            return '未知'

    def _classify_audio(self, file_path: Path) -> str:
        """分类音频文件"""
        try:
            logger.info(f"开始处理音频文件: {file_path.name}")
            # 首先尝试使用 vosk 进行离线识别
            try:
                from audio_processor import AudioProcessor
                processor = AudioProcessor()
                text = processor.transcribe_audio(file_path)
                if text:
                    logger.info(f"音频识别结果: {text[:200]}...")  # 只显示前200个字符
                    return self._call_api(text[:500])
            except ImportError:
                logger.warning("Vosk 模块未安装，将使用在线语音识别")
                
                # 使用 speech_recognition 作为备选
                import speech_recognition as sr
                from pydub import AudioSegment
                
                # 转换音频为 WAV 格式
                audio = AudioSegment.from_file(str(file_path))
                wav_path = file_path.with_suffix('.wav')
                audio.export(str(wav_path), format="wav")
                
                try:
                    # 使用 Google Speech Recognition
                    recognizer = sr.Recognizer()
                    with sr.AudioFile(str(wav_path)) as source:
                        audio = recognizer.record(source, duration=30)  # 只处理前30秒
                        text = recognizer.recognize_google(audio, language='zh-CN')
                        if text:
                            logger.info(f"在线语音识别结果: {text[:200]}...")  # 只显示前200个字符
                            return self._call_api(text[:500])
                finally:
                    # 清理临时文件
                    if wav_path.exists():
                        wav_path.unlink()
            
            logger.warning(f"音频识别结果为空: {file_path.name}")
            return '未知'
            
        except Exception as e:
            logger.error(f"音频分类失败 {file_path}: {str(e)}", exc_info=True)
            return '未知'

    def _check_environment(self) -> bool:
        """检查运行环境是否完整"""
        try:
            logger.info("开始检查运行环境...")
            
            # 检查 FFmpeg
            logger.info("检查 FFmpeg...")
            ffmpeg_path = Path(self.config['Audio']['ffmpeg_path'])
            if not ffmpeg_path.exists():
                logger.error("未找到 FFmpeg")
                return False
            
            # 测试 FFmpeg
            import subprocess
            try:
                result = subprocess.run(
                    [str(ffmpeg_path), '-version'],
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW  # 隐藏窗口
                )
                if result.returncode != 0:
                    logger.error("FFmpeg 测试失败")
                    return False
            except Exception as e:
                logger.error(f"FFmpeg 测试失败: {str(e)}")
                return False
            
            logger.info("FFmpeg 检查通过")
            
            # 检查 Tesseract
            logger.info("检查 Tesseract OCR...")
            tesseract_path = Path(self.config['OCR']['tesseract_path'])
            if not tesseract_path.exists():
                logger.error("未找到 Tesseract")
                return False
            
            # 测试 Tesseract
            try:
                result = subprocess.run(
                    [str(tesseract_path), '--version'],
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW  # 隐藏窗口
                )
                if result.returncode != 0:
                    logger.error("Tesseract 测试失败")
                    return False
                
                # 检查中文支持
                logger.info("检查 Tesseract 中文支持...")
                result = subprocess.run(
                    [str(tesseract_path), '--list-langs'],
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW  # 隐藏窗口
                )
                test_result = result.stdout
                
                needs_chinese = False
                if 'chi_sim' not in test_result:
                    logger.warning("未检测到中文支持，尝试安装...")
                    needs_chinese = True
                
            except Exception as e:
                logger.error(f"Tesseract 测试失败: {str(e)}")
                return False
            
            # 检查语言包文件
            lang_file = Path(tesseract_path).parent / 'tessdata' / 'chi_sim.traineddata'
            if not lang_file.exists():
                logger.warning("未找到中文语言包文件，尝试安装...")
                needs_chinese = True
            
            # 如果需要安装中文支持
            if needs_chinese:
                try:
                    # 下载中文语言包
                    logger.info("开始下载中文语言包...")
                    tessdata_dir = Path(tesseract_path).parent / 'tessdata'
                    tessdata_dir.mkdir(exist_ok=True)
                    
                    # 中文语言包下载地址
                    urls = [
                        "https://github.com/tesseract-ocr/tessdata/raw/4.1.0/chi_sim.traineddata",
                        "https://raw.githubusercontent.com/tesseract-ocr/tessdata/4.1.0/chi_sim.traineddata",
                        "https://ghproxy.com/https://raw.githubusercontent.com/tesseract-ocr/tessdata/4.1.0/chi_sim.traineddata"
                    ]
                    
                    success = False
                    for url in urls:
                        try:
                            logger.info(f"尝试从 {url.split('/')[2]} 下载...")
                            response = requests.get(url, stream=True)
                            if response.status_code == 200:
                                with open(lang_file, 'wb') as f:
                                    for chunk in response.iter_content(chunk_size=8192):
                                        if chunk:
                                            f.write(chunk)
                                success = True
                                logger.info("中文语言包下载完成")
                                break
                        except Exception as e:
                            logger.warning(f"从 {url} 下载失败: {str(e)}")
                            continue
                    
                    if not success:
                        logger.error("所有中文语言包下载地址均失败")
                        return False
                    
                    # 验证安装
                    test_result = os.popen(f'"{tesseract_path}" --list-langs').read()
                    if 'chi_sim' not in test_result:
                        logger.error("中文语言包安装验证失败")
                        return False
                        
                except Exception as e:
                    logger.error(f"安装中文语言包失败: {str(e)}")
                    return False
            
            logger.info("Tesseract 中文支持检查通过")
            
            # 检查配置文件
            logger.info("检查配置文件...")
            required_sections = ['API', 'Model', 'Paths', 'Threading', 'OCR', 'Audio']
            for section in required_sections:
                if section not in self.config:
                    logger.error(f"配置文件缺少 {section} 部分")
                    return False
            logger.info("配置文件检查通过")
            
            # 检查文件夹结构
            logger.info("检查文件夹结构...")
            source_dir = Path(self.config['Paths']['source_folder'])
            target_dir = Path(self.config['Paths']['target_folder'])
            if not source_dir.exists() or not target_dir.exists():
                logger.error("基本文件夹结构不完整")
                return False
            logger.info("文件夹结构检查通过")
            
            logger.info("环境检查完成：所有组件正常")
            return True
            
        except Exception as e:
            logger.error(f"环境检查失败: {str(e)}", exc_info=True)
            return False

    def _check_config(self) -> bool:
        """检查配置是否完整"""
        try:
            # 检查 API 设置
            if not self.config['API']['api_key']:
                return False
            
            # 检查文件夹
            source_path = self.config['Paths']['source_folder']
            target_path = self.config['Paths']['target_folder']
            
            if not source_path or not target_path:
                messagebox.showwarning(
                    "路径未设置",
                    "请先设置源文件夹和目标文件夹路径。\n"
                    "可以在主界面的路径设置中选择合适的目录。"
                )
                return False
            
            source_dir = Path(source_path)
            target_dir = Path(target_path)
            
            if not source_dir.exists() or not target_dir.exists():
                if not self._setup_folders():
                    return False
            
            # 检查外部程序
            if self.config.getboolean('Features', 'enable_ocr'):
                tesseract_path = Path(self.config['OCR']['tesseract_path'])
                if not tesseract_path.exists():
                    return False
            
            if self.config.getboolean('Features', 'enable_audio'):
                ffmpeg_path = Path(self.config['Audio']['ffmpeg_path'])
                if not ffmpeg_path.exists():
                    return False
            
            return True
            
        except Exception:
            return False

def main():
    try:
        # 设置日志
        log_dir = os.environ.get('LOG_DIR', 'logs')
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, 'file_classifier.log')
        
        # 配置日志记录
        logging.basicConfig(
            filename=log_file,
            level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        logger = logging.getLogger('file_classifier')
        logger.info("程序启动")
        
        # 创建主窗口
        from gui.main_window import MainWindow
        window = MainWindow()
        
        # 运行主循环
        window.run()
        
    except Exception as e:
        print(f"Error: {str(e)}")
        if logger:
            logger.error(f"程序异常退出: {str(e)}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main() 