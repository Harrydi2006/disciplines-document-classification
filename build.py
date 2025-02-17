import PyInstaller.__main__
import os
import shutil
import sys
from pathlib import Path

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
    print("复制资源文件...")
    dist_dir = Path('dist')
    
    # 创建基本目录结构
    base_dir = dist_dir / '文件分类助手'
    base_dir.mkdir(exist_ok=True)
    
    # 创建必要的目录
    (base_dir / 'source').mkdir(exist_ok=True)
    (base_dir / 'target').mkdir(exist_ok=True)
    (base_dir / 'logs').mkdir(exist_ok=True)
    
    # 复制配置文件
    if Path('config.conf').exists():
        shutil.copy2('config.conf', base_dir)
    else:
        print("警告: 未找到配置文件")
    
    # 复制可执行文件
    if (dist_dir / '文件分类助手.exe').exists():
        shutil.move(dist_dir / '文件分类助手.exe', base_dir / '文件分类助手.exe')
    
    print("资源文件复制完成")
    return base_dir  # 返回基本目录路径

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
    print("开始构建...")
    
    # 清理旧的构建文件
    clean_build()
    
    # 查找 Python DLL
    python_dll = find_python_dll()
    if not python_dll:
        print("警告: 未找到 Python DLL 文件")
        if not input("是否继续构建？(y/n): ").lower().startswith('y'):
            print("构建已取消")
            return
    else:
        print(f"找到 Python DLL: {python_dll}")
    
    # PyInstaller 参数
    params = [
        'file_classifier.py',  # 主程序文件
        '--name=文件分类助手',  # 程序名称
        '--noconsole',  # 不显示控制台窗口
        '--windowed',  # 使用 GUI 模式
        '--noconfirm',  # 覆盖现有文件
        '--clean',  # 清理临时文件
        '--add-data=gui;gui',  # 添加 GUI 模块
        '--icon=resources/icon.ico',  # 程序图标（如果有）
        '--onefile',  # 打包成单个文件
        # 添加隐式导入
        '--hidden-import=PIL',
        '--hidden-import=pytesseract',
        '--hidden-import=docx',
        '--hidden-import=PyPDF2',
        '--hidden-import=speech_recognition',
        '--hidden-import=pydub',
        '--hidden-import=py7zr',
        '--hidden-import=rarfile',
        '--hidden-import=zipfile',
    ]
    
    # 添加 Python DLL
    if python_dll:
        params.extend(['--add-binary', f'{python_dll};.'])
    
    try:
        # 运行 PyInstaller
        PyInstaller.__main__.run(params)
        
        # 复制资源文件并获取目标目录
        base_dir = copy_resources()
        
        # 创建 README.txt
        readme_content = """文件分类助手

使用说明：
1. 运行 文件分类助手.exe
2. 将需要分类的文件放入 source 文件夹
3. 分类后的文件会自动移动到 target 文件夹的对应子目录中
4. 日志文件保存在 logs 文件夹中

注意事项：
- 首次运行时请按提示配置相关设置
- 确保已安装必要的外部程序（Tesseract、FFmpeg等）
"""
        
        with open(base_dir / 'README.txt', 'w', encoding='utf-8') as f:
            f.write(readme_content)
        
        print("构建完成！")
        print(f"\n程序包已生成在: {base_dir}")
        print("请将整个文件夹复制到目标计算机使用")
        
    except Exception as e:
        print(f"构建失败: {str(e)}")
        raise

if __name__ == "__main__":
    build() 