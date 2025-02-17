import speech_recognition as sr
from pathlib import Path
import configparser
from pydub import AudioSegment
import os
from pydub.utils import which
from utils.logger import setup_logger
import time
from vosk import Model, KaldiRecognizer, SetLogLevel
import wave
import json
import requests
from tqdm import tqdm
import zipfile
import shutil

# 设置日志记录器
logger = setup_logger('audio_processor', 'audio_processor.log')

# 设置 Vosk 日志级别
SetLogLevel(-1)

class AudioProcessor:
    VOSK_MODEL_URL = "https://alphacephei.com/vosk/models/vosk-model-small-cn-0.22.zip"
    
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self._setup_ffmpeg()
        self._setup_vosk()
        
    def _download_model(self, model_path: Path):
        """下载 Vosk 模型"""
        try:
            # 创建模型目录
            model_path.parent.mkdir(parents=True, exist_ok=True)
            zip_path = model_path.parent / "model.zip"
            
            logger.info("开始下载 Vosk 中文语音模型...")
            response = requests.get(self.VOSK_MODEL_URL, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            
            with open(zip_path, 'wb') as f, tqdm(
                desc="下载模型",
                total=total_size,
                unit='iB',
                unit_scale=True,
                unit_divisor=1024,
            ) as pbar:
                for data in response.iter_content(chunk_size=1024):
                    size = f.write(data)
                    pbar.update(size)
            
            logger.info("解压模型文件...")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                # 获取压缩包内的根目录名
                root_dir = zip_ref.namelist()[0].split('/')[0]
                zip_ref.extractall(model_path.parent)
                
                # 如果解压后的目录名不是我们想要的，重命名它
                extracted_path = model_path.parent / root_dir
                if extracted_path != model_path:
                    if model_path.exists():
                        shutil.rmtree(model_path)
                    extracted_path.rename(model_path)
            
            # 清理下载的zip文件
            zip_path.unlink()
            logger.info("模型安装完成")
            return True
            
        except Exception as e:
            logger.error(f"下载模型失败: {str(e)}", exc_info=True)
            # 清理可能的部分下载文件
            if zip_path.exists():
                zip_path.unlink()
            if model_path.exists():
                shutil.rmtree(model_path)
            return False

    def _setup_ffmpeg(self):
        """设置 ffmpeg 路径"""
        try:
            config = configparser.ConfigParser()
            config.read('config.conf', encoding='utf-8')
            ffmpeg_dir = str(Path(config['Audio']['ffmpeg_path']).parent)
            
            # 设置 ffmpeg 相关路径
            ffmpeg_path = str(Path(ffmpeg_dir) / 'ffmpeg.exe')
            ffprobe_path = str(Path(ffmpeg_dir) / 'ffprobe.exe')
            
            if os.path.exists(ffmpeg_path) and os.path.exists(ffprobe_path):
                # 设置环境变量
                os.environ["PATH"] += os.pathsep + ffmpeg_dir
                
                # 设置 pydub 的路径
                AudioSegment.converter = ffmpeg_path
                AudioSegment.ffmpeg = ffmpeg_path
                AudioSegment.ffprobe = ffprobe_path
                
                # 验证设置是否成功
                if which("ffmpeg") and which("ffprobe"):
                    logger.info(f"已成功设置 ffmpeg 环境: {ffmpeg_dir}")
                else:
                    logger.warning("ffmpeg 环境设置可能不完整")
            else:
                logger.warning(f"ffmpeg 文件不完整: {ffmpeg_dir}")
        except Exception as e:
            logger.error(f"设置 ffmpeg 路径失败: {str(e)}", exc_info=True)
        
    def _setup_vosk(self):
        """设置 Vosk 模型"""
        try:
            model_path = Path('models/vosk-model-small-cn')
            if not model_path.exists():
                logger.info(f"Vosk 模型不存在，开始下载...")
                if not self._download_model(model_path):
                    logger.error("模型下载失败，语音识别功能将不可用")
                    self.model = None
                    return
            
            self.model = Model(str(model_path))
            logger.info("已加载 Vosk 模型")
        except Exception as e:
            logger.error(f"加载 Vosk 模型失败: {str(e)}", exc_info=True)
            self.model = None

    def convert_to_wav(self, audio_path: Path) -> Path:
        try:
            # 验证 ffmpeg 是否可用
            if not which("ffmpeg"):
                logger.error("ffmpeg 不可用，无法转换音频")
                return None
                
            audio = AudioSegment.from_file(str(audio_path))
            # 转换为16kHz采样率的单声道WAV文件（Vosk要求）
            audio = audio.set_frame_rate(16000).set_channels(1)
            wav_path = audio_path.with_suffix('.wav')
            audio.export(str(wav_path), format="wav")
            return wav_path
        except Exception as e:
            logger.error(f"音频转换失败 {audio_path}: {str(e)}", exc_info=True)
            return None

    def transcribe_audio(self, audio_path: Path) -> str:
        wav_path = None
        wf = None
        try:
            wav_path = self.convert_to_wav(audio_path) if audio_path.suffix.lower() != '.wav' else audio_path
            
            if not wav_path or not self.model:
                return ""

            logger.info(f"开始离线语音识别: {audio_path.name}")
            # 使用 Vosk 进行离线识别
            wf = wave.open(str(wav_path), "rb")
            rec = KaldiRecognizer(self.model, wf.getframerate())
            rec.SetWords(True)

            # 只读取前30秒的音频
            max_frames = int(30 * wf.getframerate())
            frames_read = 0
            results = []

            while True:
                data = wf.readframes(4000)
                if len(data) == 0 or frames_read >= max_frames:
                    break
                frames_read += 4000
                if rec.AcceptWaveform(data):
                    result = json.loads(rec.Result())
                    if 'text' in result and result['text']:
                        results.append(result['text'])
                        logger.debug(f"识别片段: {result['text']}")

            # 获取最后的识别结果
            final_result = json.loads(rec.FinalResult())
            if 'text' in final_result and final_result['text']:
                results.append(final_result['text'])
                logger.debug(f"最终片段: {final_result['text']}")

            text = " ".join(results)
            if text:
                logger.info(f"离线语音识别完成: {audio_path.name}")
                logger.info(f"完整识别结果: {text}")
                return text
            else:
                logger.warning(f"音频转写结果为空: {audio_path.name}")
                return ""

        except Exception as e:
            logger.error(f"音频转写失败 {audio_path}: {str(e)}", exc_info=True)
            return ""
        finally:
            # 确保文件句柄被关闭
            if wf:
                try:
                    wf.close()
                except Exception as e:
                    logger.error(f"关闭音频文件失败: {str(e)}")
            
            # 清理临时文件
            if wav_path and wav_path != audio_path:
                try:
                    # 等待一小段时间确保文件不再被使用
                    time.sleep(0.1)
                    if os.path.exists(wav_path):
                        os.remove(wav_path)
                except Exception as e:
                    logger.error(f"临时文件清理失败 {wav_path}: {str(e)}")

def main():
    try:
        processor = AudioProcessor()
        # 这里可以添加测试代码
        test_file = Path("test.mp3")
        if test_file.exists():
            result = processor.transcribe_audio(test_file)
            logger.info(f"转写结果: {result}")
    except Exception as e:
        logger.error(f"音频处理程序执行失败: {str(e)}", exc_info=True)

if __name__ == "__main__":
    main() 