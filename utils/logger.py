import logging
import os
from pathlib import Path
from datetime import datetime
import shutil

def setup_logger(name: str, log_file: str) -> logging.Logger:
    """设置日志记录器"""
    # 创建 logs 目录
    log_dir = Path('logs')
    log_dir.mkdir(exist_ok=True)
    
    # 获取当前时间作为文件名
    current_time = datetime.now()
    current_timestamp = current_time.strftime('%Y-%m-%d_%H-%M-%S')
    
    # 如果 latest.log 存在，将其重命名为时间戳格式
    latest_log = log_dir / 'latest.log'
    if latest_log.exists():
        # 获取文件的修改时间作为旧日志的时间戳
        mtime = datetime.fromtimestamp(latest_log.stat().st_mtime)
        old_timestamp = mtime.strftime('%Y-%m-%d_%H-%M-%S')
        old_log = log_dir / f'{old_timestamp}.log'
        
        # 如果文件已存在，添加序号
        counter = 1
        while old_log.exists():
            old_log = log_dir / f'{old_timestamp}_{counter}.log'
            counter += 1
            
        # 重命名旧的日志文件
        shutil.move(str(latest_log), str(old_log))
    
    # 清理旧日志文件，只保留最新的10个
    log_files = sorted(
        [f for f in log_dir.glob('*.log') if f.name != 'latest.log'],
        key=lambda x: x.stat().st_mtime,
        reverse=True
    )
    
    # 删除多余的日志文件
    for old_file in log_files[9:]:  # 保留最新的9个 + 当前的latest.log = 10个
        try:
            old_file.unlink()
        except Exception as e:
            print(f"删除旧日志文件失败: {old_file}, 错误: {e}")
    
    # 创建新的日志记录器
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    
    # 创建文件处理器
    file_handler = logging.FileHandler(latest_log, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    
    # 创建控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # 设置日志格式
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # 清除现有的处理器
    logger.handlers.clear()
    
    # 添加处理器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger 