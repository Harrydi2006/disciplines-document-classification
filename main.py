import configparser
import json
import os
import shutil
import traceback
from zhipuai import ZhipuAI

# 配置文件路径
config_file = 'config.conf'

# 检查 .conf 文件是否存在
if not os.path.exists(config_file):
    # 如果文件不存在，则创建并填充模板
    config = configparser.ConfigParser()

    config['settings'] = {
        'api_key': '',  # 留空，用户手动填写
        'path': '',  # 留空，用户手动填写
        'output_directory': '',  # 新增分类输出目录字段，用户手动填写
        'description': '根据文件名称推测其属于的学科(如语文、英语、数学、物理、化学、生物），不确定返回未知',
        'subjects': '语文,英语,数学,物理,化学,生物,未知',  # 默认值，用户可根据需要修改
        'allowed_extensions': ''  # 新增扩展名白名单，多个扩展名用逗号分隔，如 "pdf,docx,txt"
    }

    # 写入到文件中，确保以 UTF-8 编码保存
    with open(config_file, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

    # 提示用户补全配置文件
    print(f"{config_file} 文件已创建，请补全 api_key、path 和 output_directory 信息后再运行此程序。")
    exit()  # 退出程序，等待用户填充配置文件

# 读取 .conf 文件
config = configparser.ConfigParser()
config.read(config_file, encoding='utf-8')

# 从 .conf 文件中获取 api_key 和路径
api_key = config.get('settings', 'api_key')
path = config.get('settings', 'path')
output_directory = config.get('settings', 'output_directory')  # 新增输出目录
description = config.get('settings', 'description')
subjects = config.get('settings', 'subjects').split(',')
allowed_extensions = config.get('settings', 'allowed_extensions')

# 如果 allowed_extensions 不为空，则转换为列表
if allowed_extensions:
    allowed_extensions = [ext.strip().lower() for ext in allowed_extensions.split(',')]
else:
    allowed_extensions = None  # 如果为空，则允许所有文件

# 检查配置文件内容是否为空
if not api_key or not path or not output_directory:
    print(f"{config_file} 文件中的 api_key、path 或 output_directory 信息为空，请补全配置后再运行此程序。")
    exit()

# 设置API Key
client = ZhipuAI(api_key=api_key)

# 定义工具函数，用于调用API
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

# 创建目标分类文件夹（基于 output_directory，而不是 path）
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

# 定义函数用于检测文件名是否重复，并添加数字后缀
def get_unique_filename(directory, file_name):
    base_name, extension = os.path.splitext(file_name)
    counter = 1
    new_file_name = file_name

    # 检查目标文件夹是否已经存在该文件
    while os.path.exists(os.path.join(directory, new_file_name)):
        new_file_name = f"{base_name}_{counter}{extension}"
        counter += 1

    return new_file_name


# 定义调用智谱AI的函数
def classify_file(file_name):
    # 准备要发送的消息
    messages = [
        {
            "role": "user",
            "content": file_name
        }
    ]

    try:
        # 使用 chat.completions.create() 调用API
        response = client.chat.completions.create(
            model="glm-4",  # 模型名称
            messages=messages,
            tools=tools,
            tool_choice="auto"
        )

        # 获取返回的工具调用内容
        if response.choices:
            tool_calls = response.choices[0].message.tool_calls  # 直接访问属性
            if tool_calls:
                # 从返回的JSON中提取subject
                result = json.loads(tool_calls[0].function.arguments)
                subject = result.get("subject", "未知")
                return subject
    except Exception as e:
        # 打印完整的异常信息
        print(f"API 调用失败: {e}")
        traceback.print_exc()  # 打印异常堆栈信息
        return "未知"


# 遍历指定路径下的文件，进行分类并移动
for file_name in files:
    subject = classify_file(file_name)  # 调用API获取学科分类
    source_path = os.path.join(path, file_name)
    destination_directory = os.path.join(output_directory, subject)

    # 检查文件名是否重复，获取唯一的文件名
    unique_file_name = get_unique_filename(destination_directory, file_name)
    destination_path = os.path.join(destination_directory, unique_file_name)

    # 移动文件到对应学科文件夹
    shutil.move(source_path, destination_path)
    print(
        f"文件 '{file_name}' 已分类到 {subject} 文件夹，重命名为 '{unique_file_name}'." if unique_file_name != file_name else f"文件 '{file_name}' 已分类到 {subject} 文件夹。")
