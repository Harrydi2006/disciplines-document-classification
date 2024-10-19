import configparser
import json
import os
import shutil
import traceback
import re
from docx import Document
from PyPDF2 import PdfReader
from openai import OpenAI  

# 配置文件路径
config_file = 'config.conf'

# 检查 .conf 文件是否存在
if not os.path.exists(config_file):
    config = configparser.ConfigParser()
    config['settings'] = {
        'api_key': '',
        'path': '',
        'subjects': '语文,英语,数学,物理,化学,生物,未知',
        'allowed_extensions': '',
        'output_directory': ''
    }
    with open(config_file, 'w', encoding='utf-8') as configfile:
        config.write(configfile)
    print(f"{config_file} 文件已创建，请补全 api_key 和 path 信息后再运行此程序。")
    exit()

config = configparser.ConfigParser()
config.read(config_file, encoding='utf-8')

api_key = config.get('settings', 'api_key')
path = config.get('settings', 'path')
output_directory = config.get('settings', 'output_directory')
subjects = config.get('settings', 'subjects').split(',')
allowed_extensions = config.get('settings', 'allowed_extensions')

if allowed_extensions:
    allowed_extensions = [ext.strip().lower() for ext in allowed_extensions.split(',')]
else:
    allowed_extensions = None

if not api_key or not path or not output_directory:
    print(f"{config_file} 文件中的 api_key、path 或 output_directory 信息为空，请补全配置后再运行此程序。")
    exit()

# 初始化OpenAI客户端
client = OpenAI(
    api_key=api_key,
    base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
)

# 创建目标分类文件夹
for subject in subjects:
    folder_path = os.path.join(output_directory, subject.strip())
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

# 获取指定路径下的文件列表
files = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]

# 如果 allowed_extensions 不为空，则过滤符合扩展名的文件
if allowed_extensions:
    files = [f for f in files if os.path.splitext(f)[1][1:].lower() in allowed_extensions]

# 如果没有符合条件的文件，给出提示
if not files:
    print(f"没有符合扩展名 {allowed_extensions} 的文件。")
    exit()


def get_unique_filename(directory, file_name):
    base_name, extension = os.path.splitext(file_name)
    counter = 1
    new_file_name = file_name

    # 检查目标文件夹是否已经存在该文件
    while os.path.exists(os.path.join(directory, new_file_name)):
        new_file_name = f"{base_name}_{counter}{extension}"
        counter += 1

    return new_file_name


# 清理文本：去除特殊字符并处理空格
def clean_text(text):
    text = re.sub(r'[^\w\s]', '', text)  # 去掉所有非字母、数字、空格的字符
    text = re.sub(r'\s+', ' ', text)  # 将多个连续的空格替换为一个空格
    return text.strip()


def extract_text_from_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.docx':
            return extract_text_from_docx(file_path)
        elif ext == '.pdf':
            return extract_text_from_pdf(file_path)
        elif ext == '.txt':
            return extract_text_from_txt(file_path)
        else:
            print(f"不支持的文件类型: {ext}")
            return ''
    except Exception as e:
        print(f"读取文件内容失败: {e}")
        return ''


def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return clean_text('\n'.join(full_text))
    except Exception as e:
        print(f"读取 .docx 文件失败: {e}")
        return ''


def extract_text_from_pdf(file_path):
    try:
        with open(file_path, 'rb') as pdf_file:
            reader = PdfReader(pdf_file)
            text = ''
            for page in reader.pages:
                text += page.extract_text() or ''
            return clean_text(text)
    except Exception as e:
        print(f"读取 .pdf 文件失败: {e}")
        return ''


def extract_text_from_txt(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return clean_text(f.read())
    except UnicodeDecodeError:
        try:
            with open(file_path, 'r', encoding='gbk') as f:
                return clean_text(f.read())
        except Exception as e:
            print(f"读取 .txt 文件失败: {e}")
            return ''


def classify_file(file_name, file_path):
    # 从配置文件中读取subjects并替换消息中的相关内容
    subjects_from_conf = config.get('settings', 'subjects')

    system_message = f'You are a subject file classification tool. Please respond in JSON format with a key called "subject". The value should be one of the following: {subjects_from_conf}.'
    user_message = f'文件名：{file_name}, 判断其所属学科(回复仅限:{subjects_from_conf})不确定返回未知'

    messages = [
        {'role': 'system', 'content': system_message},
        {'role': 'user', 'content': user_message}
    ]

    subject = "未知"

    try:
        # 调用OpenAI API进行文件名分类
        completion = client.chat.completions.create(
            model="qwen-plus",
            messages=messages,
            response_format={"type": "json_object"}
        )

        # 解析返回结果
        if completion and completion.choices:
            # 获取消息内容
            message_content = completion.choices[0].message.content
            result_json = json.loads(message_content)
            subject = result_json.get('subject', '未知')

        # 打印调试信息
        print(f"API 返回的学科分类: {subject}")

    except Exception as e:
        print(f"API 调用失败: {e}")
        traceback.print_exc()

    # 如果文件名分类返回 "未知"，尝试读取文件内容进行分类
    if subject == "未知":
        file_text = extract_text_from_file(file_path)
        if file_text:
            # 打印调试信息，查看截取到的文件内容
            print(f"读取到的文件内容（前500字符）：\n{file_text[:500]}")

            # 同样地，在使用文件内容进行分类时也应保持一致的消息格式
            messages = [
                {'role': 'system', 'content': system_message},
                {'role': 'user', 'content': file_text[:500]}  # 传递前500个字符
            ]
            try:
                # 调用API进行文件内容分类
                completion = client.chat.completions.create(
                    model="qwen-plus",
                    messages=messages,
                    response_format={"type": "json_object"}
                )

                if completion and completion.choices:
                    # 获取消息内容
                    message_content = completion.choices[0].message.content
                    result_json = json.loads(message_content)
                    subject = result_json.get('subject', '未知')

                # 打印调试信息
                print(f"API 返回的学科分类（文件内容）：{subject}")

            except Exception as e:
                print(f"API 调用失败: {e}")
                traceback.print_exc()

    return subject


# 遍历指定路径下的文件，进行分类并移动
for file_name in files:
    source_path = os.path.join(path, file_name)
    subject = classify_file(file_name, source_path)
    destination_directory = os.path.join(output_directory, subject)

    # 检查文件名是否重复，获取唯一的文件名
    unique_file_name = get_unique_filename(destination_directory, file_name)
    destination_path = os.path.join(destination_directory, unique_file_name)

    # 移动文件到对应学科文件夹
    shutil.move(source_path, destination_path)
    print(
        f"文件 '{file_name}' 已分类到 {subject} 文件夹，重命名为 '{unique_file_name}'。" if unique_file_name != file_name else f"文件 '{file_name}' 已分类到 {subject} 文件夹。")
