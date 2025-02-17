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
    print("Cleaning old build files...")  # 使用英文输出
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
    tcl_path = os.path.dirname(tkinter.__file__)
    tk_path = os.path.join(tcl_path, 'tk')
    tcl_lib = os.path.join(tcl_path, 'tcl')
    
    # PyInstaller 参数
    params = [
        'file_classifier.py',  # 主程序文件
        '--name=FileClassifier',  # 使用英文名称
        '--noconsole',          # 不显示控制台
        '--windowed',           # 使用 GUI 模式
        '--noconfirm',
        '--clean',
        '--add-data=gui;gui',   # 添加 GUI 模块
        '--add-data=config.conf.template;.',  # 添加配置模板
        f'--add-data={tcl_lib};tcl',  # 添加 Tcl 库
        f'--add-data={tk_path};tk',   # 添加 Tk 库
        '--icon=resources/icon.ico',
        '--onefile',            # 打包成单个文件
        # 添加所有必要的导入
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
        # 添加运行时钩子
        '--runtime-hook=runtime_hook.py'
    ]
    
    # 创建运行时钩子文件
    with open('runtime_hook.py', 'w', encoding='utf-8') as f:
        f.write("""
import os
import sys
import tkinter as tk

# 设置正确的 DLL 搜索路径
if hasattr(sys, '_MEIPASS'):
    os.environ['TCL_LIBRARY'] = os.path.join(sys._MEIPASS, 'tcl')
    os.environ['TK_LIBRARY'] = os.path.join(sys._MEIPASS, 'tk')
    os.environ['TCL_LIBRARY'] = os.path.join(sys._MEIPASS, 'tcl', 'tcl8.6')
    os.environ['TK_LIBRARY'] = os.path.join(sys._MEIPASS, 'tk', 'tk8.6')

# 修复 tkinter 标题
try:
    root = tk.Tk()
    root.withdraw()
    root.title('文件分类助手')
    root.destroy()
except Exception:
    pass
""")
    
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