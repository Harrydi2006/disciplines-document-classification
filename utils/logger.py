import logging
import os
from pathlib import Path
from datetime import datetime
import shutil

def setup_logger(name: str, log_file: str) -> logging.Logger:
    """设置日志记录器"""
    try:
        # 创建 logs 目录
        log_dir = Path('logs')
        log_dir.mkdir(exist_ok=True)
        
        # 获取当前时间作为文件名
        current_time = datetime.now()
        current_timestamp = current_time.strftime('%Y-%m-%d_%H-%M-%S')
        
        # 使用时间戳作为日志文件名，避免使用latest.log
        log_file = log_dir / f'{name}_{current_timestamp}.log'
        
        # 如果文件已存在，添加序号
        counter = 1
        while log_file.exists():
            log_file = log_dir / f'{name}_{current_timestamp}_{counter}.log'
            counter += 1
        
        # 清理旧日志文件，只保留最新的10个
        log_files = sorted(
            [f for f in log_dir.glob(f'{name}_*.log')],
            key=lambda x: x.stat().st_mtime,
            reverse=True
        )
        
        # 删除多余的日志文件
        for old_file in log_files[9:]:  # 保留最新的10个
            try:
                old_file.unlink()
            except Exception as e:
                print(f"删除旧日志文件失败: {old_file}, 错误: {e}")
        
        # 创建新的日志记录器
        logger = logging.getLogger(name)
        logger.setLevel(logging.INFO)
        
        # 创建文件处理器，使用 delay=True 延迟创建文件
        max_retries = 3
        retry_count = 0
        file_handler = None
        
        while retry_count < max_retries:
            try:
                file_handler = logging.FileHandler(log_file, encoding='utf-8', mode='w', delay=True)
                break
            except (PermissionError, OSError) as e:
                retry_count += 1
                if retry_count == max_retries:
                    # 如果所有重试都失败，使用标准错误输出
                    import sys
                    file_handler = logging.StreamHandler(sys.stderr)
                    print(f"无法创建日志文件，将使用标准错误输出: {str(e)}")
                else:
                    import time
                    time.sleep(0.1)  # 短暂延迟后重试
        
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
        
    except Exception as e:
        # 如果设置日志失败，至少创建一个基本的控制台日志记录器
        print(f"设置日志记录器时出错: {str(e)}")
        fallback_logger = logging.getLogger(name)
        fallback_logger.setLevel(logging.INFO)
        
        if not fallback_logger.handlers:
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            fallback_logger.addHandler(console_handler)
        
        return fallback_logger 