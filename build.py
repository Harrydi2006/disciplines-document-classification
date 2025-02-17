import PyInstaller.__main__
import os
import shutil
import sys
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
    print("Cleaning old build files...")
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

def build():
    """构建可执行文件"""
    print("Starting build process...")
    
    # 清理旧的构建文件
    clean_build()
    
    # 查找 Tcl/Tk 库路径
    import tkinter
    import _tkinter
    tcl_tk_paths = []
    
    # 尝试多个可能的路径
    possible_paths = [
        os.path.dirname(tkinter.__file__),
        os.path.dirname(_tkinter.__file__),
        os.path.join(sys.prefix, 'tcl'),
        os.path.join(sys.prefix, 'lib', 'tcl'),
        os.path.join(sys.prefix, 'lib', 'tk'),
    ]
    
    for base_path in possible_paths:
        if os.path.exists(base_path):
            # 查找 tcl*.dll 和 tk*.dll
            for root, dirs, files in os.walk(base_path):
                for file in files:
                    if file.startswith(('tcl', 'tk')) and file.endswith('.dll'):
                        dll_path = os.path.join(root, file)
                        tcl_tk_paths.append((dll_path, '.'))
    
    # 创建运行时钩子文件
    with open('runtime_hook.py', 'w', encoding='utf-8') as f:
        f.write("""
import os
import sys
import tkinter as tk

def setup_environment():
    try:
        print("Setting up environment...")
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
            print(f"MEIPASS path: {base_path}")
            
            # 设置资源目录
            os.environ['RESOURCE_DIR'] = base_path
            print(f"Resource dir set to: {base_path}")
            
            # 只创建日志目录
            log_dir = os.path.join(os.getcwd(), 'logs')
            os.makedirs(log_dir, exist_ok=True)
            print(f"Created logs directory: {log_dir}")
            
            # 处理配置文件
            config_template = os.path.join(base_path, 'config.conf.template')
            config_file = 'config.conf'
            if not os.path.exists(config_file) and os.path.exists(config_template):
                import shutil
                shutil.copy2(config_template, config_file)
                print(f"Copied config template to: {config_file}")
            
            # 添加模块搜索路径
            for path in [base_path, os.path.join(base_path, 'gui'), os.path.join(base_path, 'utils')]:
                if path not in sys.path:
                    sys.path.insert(0, path)
                    print(f"Added to sys.path: {path}")
                    
            # 设置日志文件路径
            os.environ['LOG_DIR'] = log_dir
            
            # 初始化并隐藏 tkinter 根窗口
            try:
                # 在 Windows 上使用 win32gui 隐藏窗口
                if sys.platform == 'win32':
                    try:
                        import win32gui
                        import win32con
                        
                        def hide_tk_window(hwnd, extra):
                            classname = win32gui.GetClassName(hwnd)
                            title = win32gui.GetWindowText(hwnd)
                            if 'tk' in classname.lower():
                                win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
                        
                        win32gui.EnumWindows(hide_tk_window, None)
                    except ImportError:
                        pass
                
                # 创建并配置根窗口
                root = tk.Tk()
                root.withdraw()
                
                # 设置窗口属性
                root.title('文件分类助手')
                root.attributes('-alpha', 0)
                root.attributes('-topmost', True)
                root.overrideredirect(True)
                root.geometry('0x0+0+0')
                
                # 在 Windows 上额外设置
                if sys.platform == 'win32':
                    root.wm_attributes('-toolwindow', True)
                    
                # 保存根窗口引用
                global _root
                _root = root
                
            except Exception as e:
                print(f"Tkinter initialization error: {str(e)}")
                import traceback
                traceback.print_exc()
            
    except Exception as e:
        print(f"Error in setup_environment: {str(e)}")
        import traceback
        traceback.print_exc()

# 设置环境
setup_environment()
""")

    # 更新 PyInstaller 参数
    params = [
        'file_classifier.py',
        '--name=FileClassifier',
        '--noconsole',
        '--windowed',
        '--noconfirm',
        '--clean',
        '--add-data=gui;gui',
        '--add-data=utils;utils',
        '--add-data=config.conf.template;.',
        '--add-data=resources;resources',
        '--icon=resources/icon.ico',
        '--onefile',
        '--debug=all',
        '--hidden-import=tkinter',
        '--hidden-import=tkinter.ttk',
        '--hidden-import=PIL',
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=pytesseract',
        '--hidden-import=docx',
        '--hidden-import=PyPDF2',
        '--hidden-import=speech_recognition',
        '--hidden-import=pydub',
        '--hidden-import=py7zr',
        '--hidden-import=rarfile',
        '--hidden-import=zipfile',
        '--hidden-import=utils.logger',
        '--hidden-import=gui.main_window',
        '--hidden-import=gui.setup_window',
        '--hidden-import=win32gui',
        '--collect-all=gui',
        '--runtime-hook=runtime_hook.py'
    ]
    
    # 添加找到的 Tcl/Tk DLL
    for src, dst in tcl_tk_paths:
        params.extend(['--add-binary', f'{src};{dst}'])
    
    try:
        # 运行 PyInstaller
        PyInstaller.__main__.run(params)
        print("Build completed successfully!")
        
    except Exception as e:
        print(f"Build failed: {str(e)}")
        raise
    finally:
        # 清理临时文件
        if os.path.exists('runtime_hook.py'):
            os.remove('runtime_hook.py')

if __name__ == "__main__":
    build() 