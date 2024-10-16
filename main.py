import configparser
import json
import os
import shutil
import traceback
import textract
import re  # 新增导入正则模块
from PyPDF2 import PdfReader
from zhipuai import ZhipuAI

# 配置文件路径
config_file = 'config.conf'

# 检查 .conf 文件是否存在
if not os.path.exists(config_file):
    config = configparser.ConfigParser()
    config['settings'] = {
        'api_key': '',
        'path': '',
        'description': '根据文件名称推测其属于的学科(如语文、英语、数学、物理、化学、生物），不确定返回未知',
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
description = config.get('settings', 'description')
subjects = config.get('settings', 'subjects').split(',')
allowed_extensions = config.get('settings', 'allowed_extensions')

if allowed_extensions:
    allowed_extensions = [ext.strip().lower() for ext in allowed_extensions.split(',')]
else:
    allowed_extensions = None

if not api_key or not path or not output_directory:
    print(f"{config_file} 文件中的 api_key、path 或 output_directory 信息为空，请补全配置后再运行此程序。")
    exit()

client = ZhipuAI(api_key=api_key)

tools = [
    {
        "type": "function",
        "function": {
            "name": "分类文件",
            "description": description,
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {
                        "type": "string",
                        "description": "文件名称",
                    },
                    "subject": {
                        "type": "string",
                        "description": "学科",
                    },
                },
                "required": ["filename", "subject"],
            },
        }
    }
]

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
    # 去掉所有非字母、数字、空格的字符
    text = re.sub(r'[^\w\s]', '', text)
    # 将多个连续的空格替换为一个空格
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_text_from_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext in ['.doc', '.docx', '.ppt', '.pptx', '.txt']:
            # 尝试读取文件内容
            raw_text = textract.process(file_path).decode('utf-8', errors='ignore')
            return clean_text(raw_text)
        elif ext == '.pdf':
            with open(file_path, 'rb') as pdf_file:
                reader = PdfReader(pdf_file)
                text = ''
                for page in reader.pages:
                    text += page.extract_text() or ''  # 确保不返回 None
                return clean_text(text)
    except Exception as e:
        print(f"读取文件内容失败: {e}")
        return ''

def classify_file(file_name, file_path):
    messages = [{"role": "user", "content": file_name}]
    subject = "未知"

    try:
        # 调用API进行文件名分类
        response = client.chat.completions.create(
            model="glm-4",
            messages=messages,
            tools=tools,
            tool_choice="auto"
        )

        if response.choices:
            tool_calls = response.choices[0].message.tool_calls
            if tool_calls:
                result = json.loads(tool_calls[0].function.arguments)
                subject = result.get("subject", "未知")

        # 打印调试信息，查看分类返回的内容
        print(f"API 返回的学科分类: {subject}")

    except Exception as e:
        print(f"API 调用失败: {e}")
        traceback.print_exc()

    # 如果文件名分类返回 "未知"，尝试读取文件内容进行分类
    if subject == "未知":
        file_text = extract_text_from_file(file_path)
        if file_text:
            # 打印调试信息，查看截取到的文件内容
            print(f"读取到的文件内容（前500字符）：\n{file_text[:50]}")

            messages = [{"role": "user", "content": file_text[:50]}]  # 传递前500个字符
            try:
                # 调用API进行文件内容分类
                response = client.chat.completions.create(
                    model="glm-4-plus",
                    messages=messages,
                    tools=tools,
                    tool_choice="auto"
                )

                if response.choices:
                    tool_calls = response.choices[0].message.tool_calls
                    if tool_calls:
                        result = json.loads(tool_calls[0].function.arguments)
                        subject = result.get("subject", "未知")

                # 打印调试信息，查看分类返回的内容
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
    print(f"文件 '{file_name}' 已分类到 {subject} 文件夹，重命名为 '{unique_file_name}'。" if unique_file_name != file_name else f"文件 '{file_name}' 已分类到 {subject} 文件夹。")
