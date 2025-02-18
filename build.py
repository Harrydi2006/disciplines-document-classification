import PyInstaller.__main__
import os
import shutil
import sys
import argparse
from pathlib import Path

# 设置控制台输出编码
if sys.platform.startswith('win'):
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleCP(65001)
        kernel32.SetConsoleOutputCP(65001)
    except:
        pass

# 确保输出使用 UTF-8 编码
sys.stdout.reconfigure(encoding='utf-8') if hasattr(sys.stdout, 'reconfigure') else None

def clean_build():
    """清理旧的构建文件"""
    print("清理旧的构建文件...")
    dirs_to_clean = ['build', 'dist']
    files_to_clean = ['*.spec']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            
    for pattern in files_to_clean:
        for file in Path('.').glob(pattern):
            file.unlink()

def copy_resources():
    """复制必要的资源文件到 dist 目录"""
    print("Copying resource files...")  # 使用英文输出
    dist_dir = Path('dist')
    
    # 创建基本目录结构
    base_dir = dist_dir / 'FileClassifier'  # 使用英文名称
    base_dir.mkdir(exist_ok=True)
    
    # 创建必要的目录
    (base_dir / 'source').mkdir(exist_ok=True)
    (base_dir / 'target').mkdir(exist_ok=True)
    (base_dir / 'logs').mkdir(exist_ok=True)
    
    # 复制配置文件
    if Path('config.conf').exists():
        shutil.copy2('config.conf', base_dir)
    else:
        print("Warning: Configuration file not found")  # 使用英文输出
    
    # 复制可执行文件
    if (dist_dir / 'FileClassifier.exe').exists():
        shutil.move(dist_dir / 'FileClassifier.exe', base_dir / 'FileClassifier.exe')
    
    print("Resource files copied")  # 使用英文输出
    return base_dir

def find_python_dll():
    """查找 Python DLL 文件"""
    # 获取 Python 版本信息
    version = sys.version_info
    dll_name = f'python{version.major}{version.minor}.dll'
    
    # 可能的路径列表
    possible_paths = [
        sys.prefix,  # Python 安装根目录
        os.path.dirname(sys.executable),  # Python 可执行文件目录
        os.path.join(sys.prefix, 'DLLs'),  # DLLs 目录
        os.path.join(sys.prefix, 'Library', 'bin'),  # Library/bin 目录
        os.environ.get('WINDIR', ''),  # Windows 目录
        os.path.join(os.environ.get('WINDIR', ''), 'System32'),  # System32
        os.path.join(os.environ.get('WINDIR', ''), 'SysWOW64'),  # SysWOW64
    ]
    
    print(f"正在查找 {dll_name}...")
    print("搜索路径:")
    
    # 搜索所有可能的路径
    for path in possible_paths:
        if not path:
            continue
            
        dll_path = os.path.join(path, dll_name)
        print(f"检查: {dll_path}")
        
        if os.path.exists(dll_path):
            print(f"找到 DLL: {dll_path}")
            return dll_path
    
    # 如果没找到，尝试使用 ctypes 查找
    try:
        import ctypes
        from ctypes.util import find_library
        
        lib_path = find_library(f'python{version.major}{version.minor}')
        if lib_path:
            print(f"通过 ctypes 找到 DLL: {lib_path}")
            return lib_path
    except Exception as e:
        print(f"ctypes 查找失败: {str(e)}")
    
    # 如果还是没找到，尝试从注册表查找
    try:
        import winreg
        key_path = f"SOFTWARE\\Python\\PythonCore\\{version.major}.{version.minor}\\InstallPath"
        
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            install_path = winreg.QueryValue(key, None)
            dll_path = os.path.join(install_path, dll_name)
            
            if os.path.exists(dll_path):
                print(f"通过注册表找到 DLL: {dll_path}")
                return dll_path
    except Exception as e:
        print(f"注册表查找失败: {str(e)}")
    
    print("未找到 Python DLL")
    return None

def build_exe(include_dependencies=False):
    """构建可执行文件"""
    print("开始构建可执行文件...")
    
    # 清理旧的构建文件
    clean_build()
    
    # 基本的 PyInstaller 参数
    args = [
        'file_classifier.py',  # 主程序文件
        '--name=FileClassifier',  # 输出文件名
        '--onefile',  # 打包成单个文件
        '--noconsole',  # 不显示控制台
        '--icon=resources/icon.ico',  # 程序图标
        '--hidden-import=PIL._tkinter_finder',  # 添加隐藏导入
        '--hidden-import=win32gui',
        '--hidden-import=win32con',
        '--hidden-import=pythoncom',
        '--hidden-import=pywintypes',
        '--hidden-import=win32api',
        '--hidden-import=tkinter',
        '--hidden-import=tkinter.ttk',
        '--hidden-import=tkinter.messagebox',
        '--hidden-import=tkinter.filedialog',
        '--clean',  # 清理临时文件
        '--uac-admin',  # 请求管理员权限
    ]
    
    if include_dependencies:
        # 添加依赖文件
        deps_dir = Path('resources/dependencies')
        if not deps_dir.exists():
            print("错误：依赖目录不存在")
            sys.exit(1)
            
        # 检查依赖文件是否存在
        tesseract_installer = deps_dir / 'tesseract-installer.exe'
        ffmpeg_zip = deps_dir / 'ffmpeg.zip'
        
        if not tesseract_installer.exists() or not ffmpeg_zip.exists():
            print("错误：依赖文件不完整")
            sys.exit(1)
            
        # 添加依赖文件到构建
        args.extend([
            '--add-data=resources/dependencies;dependencies'
        ])
        
        # 创建运行时钩子
        with open('runtime_hook.py', 'w', encoding='utf-8') as f:
            f.write("""
import os
import sys
import shutil
import subprocess
import zipfile
from pathlib import Path

def setup_dependencies():
    try:
        if getattr(sys, '_MEIPASS', None):
            base_path = Path(sys._MEIPASS)
            deps_path = base_path / 'dependencies'
            
            # 设置 FFmpeg
            ffmpeg_zip = deps_path / 'ffmpeg.zip'
            if ffmpeg_zip.exists():
                ffmpeg_dir = Path('C:/ffmpeg')
                if not ffmpeg_dir.exists():
                    print("正在安装 FFmpeg...")
                    with zipfile.ZipFile(ffmpeg_zip) as zf:
                        # 解压到临时目录
                        temp_dir = Path('temp_ffmpeg')
                        zf.extractall(temp_dir)
                        # 移动 FFmpeg 目录
                        ffmpeg_extracted = next(temp_dir.glob('ffmpeg-*'))
                        shutil.move(str(ffmpeg_extracted), str(ffmpeg_dir))
                        # 清理临时目录
                        shutil.rmtree(temp_dir)
                    print("FFmpeg 安装完成")
            
            # 设置 Tesseract
            tesseract_installer = deps_path / 'tesseract-installer.exe'
            if tesseract_installer.exists():
                tesseract_path = Path('C:/Program Files/Tesseract-OCR/tesseract.exe')
                if not tesseract_path.exists():
                    print("正在安装 Tesseract OCR...")
                    subprocess.run([str(tesseract_installer), '/S'], 
                                creationflags=subprocess.CREATE_NO_WINDOW)
                    print("Tesseract OCR 安装完成")
                    
    except Exception as e:
        print(f"设置依赖时出错: {str(e)}")
        import traceback
        traceback.print_exc()

# 设置依赖
setup_dependencies()
""")
            
        # 添加运行时钩子
        args.extend(['--runtime-hook=runtime_hook.py'])
    
    # 运行 PyInstaller
    PyInstaller.__main__.run(args)
    
    # 清理运行时钩子文件
    if include_dependencies and os.path.exists('runtime_hook.py'):
        os.remove('runtime_hook.py')
    
    print("构建完成！")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='构建文件分类助手')
    parser.add_argument('--include-dependencies', action='store_true',
                      help='包含 Tesseract 和 FFmpeg 依赖')
    
    args = parser.parse_args()
    
    try:
        build_exe(include_dependencies=args.include_dependencies)
    except Exception as e:
        print(f"构建失败: {str(e)}")
        sys.exit(1) 