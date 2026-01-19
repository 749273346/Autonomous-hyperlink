import os
import random
import time

TEST_DIR = r"e:\QC-攻关小组\正在进行项目\自主超链接\Autonomous-hyperlink\测试文件"

# 支持的文件扩展名，包括 Office 和 WPS
EXTENSIONS = [
    '.doc', '.docx', '.wps',
    '.xls', '.xlsx', '.et',
    '.ppt', '.pptx', '.dps',
    '.pdf', '.txt'
]

# 随机文件名主题
TOPICS = [
    "关于做好2025年防洪工作的通知",
    "2025年安全生产月活动方案",
    "关于加强网络安全管理的通知",
    "季度工作总结与计划",
    "职工技能培训考核表",
    "关于进一步规范办公用品管理的通知",
    "党支部会议纪要",
    "突发事件应急预案",
    "关于开展卫生大检查的通知",
    "设备维护保养记录",
    "关于落实全员安全生产责任制的意见",
    "关于调整作息时间的通知",
    "财务报销管理办法解读"
]

# 部门/来源
DEPARTMENTS = ["供电", "工务", "电务", "车务", "机务", "客运", "人事", "财务"]

def generate_random_filename(folder_name):
    # 从文件夹名提取前缀，例如 "1-上级文" -> "上级文"
    prefix = folder_name.split('-')[-1] if '-' in folder_name else folder_name
    
    topic = random.choice(TOPICS)
    dept = random.choice(DEPARTMENTS)
    year = "2025"
    
    # 随机组合文件名格式
    name_formats = [
        f"{prefix}_{year}_{topic}",
        f"（{dept}函〔{year}〕{random.randint(1, 100)}号）{topic}",
        f"{topic}",
        f"{year}年{dept}段{prefix}材料",
        f"{dept}段关于{prefix}工作的汇报"
    ]
    
    base_name = random.choice(name_formats)
    ext = random.choice(EXTENSIONS)
    return f"{base_name}{ext}"

def create_dummy_file(file_path):
    try:
        # 写入一些简单的文本内容，避免文件为空（虽然空文件也没事，但有内容更逼真）
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(f"This is a dummy file generated for testing.\n")
            f.write(f"File Name: {os.path.basename(file_path)}\n")
            f.write(f"Created Time: {time.ctime()}\n")
            f.write(f"Type: {os.path.splitext(file_path)[1]}\n")
        return True
    except Exception as e:
        print(f"Error creating {file_path}: {e}")
        return False

def populate_folders():
    print(f"开始填充测试文件到: {TEST_DIR}")
    if not os.path.exists(TEST_DIR):
        print(f"目录不存在: {TEST_DIR}")
        return

    total_created = 0
    # 遍历一级目录
    for item in os.listdir(TEST_DIR):
        item_path = os.path.join(TEST_DIR, item)
        
        # 只处理目录，且看起来像分类目录（数字开头，如 "1-", "01-" 等）
        # 或者包含中文的目录也可以考虑，但主要是针对分类文件夹
        if os.path.isdir(item_path):
            # 简单的过滤：如果是 __pycache__ 或 .trae 等则跳过
            if item.startswith('.') or item == "__pycache__":
                continue
                
            print(f"正在处理文件夹: {item}")
            
            # 目标是确保 25 文件夹里有文件
            # 如果没有 25 文件夹，则创建它
            dir_25 = os.path.join(item_path, "25")
            
            if not os.path.exists(dir_25):
                try:
                    os.makedirs(dir_25)
                    print(f"  已创建子目录: 25")
                except Exception as e:
                    print(f"  创建子目录失败: {e}")
                    # 如果无法创建 25，就退而求其次用当前目录
                    dir_25 = item_path
            
            # 开始在 dir_25 中填充文件
            target_dir = dir_25
            
            # 1. 随机生成 5-8 个文件
            num_random = random.randint(5, 8)
            for _ in range(num_random):
                filename = generate_random_filename(item)
                file_path = os.path.join(target_dir, filename)
                # 避免覆盖已存在的文件
                if not os.path.exists(file_path):
                    if create_dummy_file(file_path):
                        total_created += 1
            
            # 2. 确保每种扩展名至少存在一个
            for ext in EXTENSIONS:
                # 检查当前目录下是否已有该后缀的文件
                has_ext = any(f.endswith(ext) for f in os.listdir(target_dir))
                if not has_ext:
                    # 如果没有，强制创建一个
                    filename = f"测试样本_{item}_{random.randint(100,999)}{ext}"
                    file_path = os.path.join(target_dir, filename)
                    if create_dummy_file(file_path):
                        total_created += 1
                        print(f"    补充缺失类型 {ext}: {filename}")
            
            print(f"  - 完成处理 {item}")

    print(f"\n任务完成！总共新创建了 {total_created} 个文件。")

if __name__ == "__main__":
    populate_folders()
