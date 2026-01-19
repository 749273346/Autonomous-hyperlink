import os
import random
import time

TEST_DIR = r"e:\QC-攻关小组\正在进行项目\自主超链接\Autonomous-hyperlink\测试文件"

# 定义文件类型组，确保每个组都至少有一个文件
TYPE_GROUPS = {
    "Word": ['.doc', '.docx', '.wps'],
    "Excel": ['.xls', '.xlsx', '.et'],
    "PPT": ['.ppt', '.pptx', '.dps'],
    "Other": ['.pdf', '.txt']
}

# 所有扩展名扁平列表
ALL_EXTENSIONS = [ext for group in TYPE_GROUPS.values() for ext in group]

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
    "财务报销管理办法解读",
    "专项整治行动方案",
    "关于节假日值班安排的通知"
]

# 部门/来源
DEPARTMENTS = ["供电", "工务", "电务", "车务", "机务", "客运", "人事", "财务", "安监", "科信"]

def generate_random_filename(folder_name, ext):
    # 尝试从文件夹路径中提取有意义的名称作为前缀
    # 例如 ...\1-上级文\25 -> "上级文"
    parts = folder_name.split(os.sep)
    prefix = parts[-1]
    # 如果是纯数字（如 25, 26），尝试取上一级目录名
    if prefix.isdigit() and len(parts) > 1:
        prefix = parts[-2]
    
    # 清理前缀中的数字编号，如 "1-上级文" -> "上级文"
    if '-' in prefix:
        prefix = prefix.split('-')[-1]
        
    topic = random.choice(TOPICS)
    dept = random.choice(DEPARTMENTS)
    year = "2025"
    
    # 随机组合文件名格式
    name_formats = [
        f"{prefix}_{year}_{topic}",
        f"（{dept}函〔{year}〕{random.randint(1, 100)}号）{topic}",
        f"{dept}段{topic}",
        f"{year}年{dept}段{prefix}材料汇总",
        f"附件：{topic}说明",
        f"{prefix}工作汇报_{random.randint(1000,9999)}"
    ]
    
    base_name = random.choice(name_formats)
    return f"{base_name}{ext}"

def create_dummy_file(file_path):
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(f"This is a dummy file generated for recursive testing.\n")
            f.write(f"File Path: {file_path}\n")
            f.write(f"Created Time: {time.ctime()}\n")
            f.write(f"Random Content: {'*'*random.randint(10, 50)}\n")
        return True
    except Exception as e:
        print(f"Error creating {file_path}: {e}")
        return False

def populate_recursive():
    print(f"开始递归填充测试文件到: {TEST_DIR}")
    if not os.path.exists(TEST_DIR):
        print(f"目录不存在: {TEST_DIR}")
        return

    total_created = 0
    
    # os.walk 递归遍历所有子目录
    for root, dirs, files in os.walk(TEST_DIR):
        # 排除根目录本身（根目录通常只放脚本和Excel，不放乱七八糟的模拟文件）
        if root == TEST_DIR:
            continue
            
        # 排除隐藏目录和系统目录
        if any(part.startswith('.') or part == "__pycache__" for part in root.split(os.sep)):
            continue
            
        print(f"正在检查目录: {root}")
        
        # 统计当前目录下已有的扩展名
        existing_exts = {os.path.splitext(f)[1] for f in files}
        
        # 确保每个大类（Word, Excel, PPT, Other）至少有一个文件
        for group_name, exts in TYPE_GROUPS.items():
            # 检查该组是否已有文件存在
            group_has_file = any(ext in existing_exts for ext in exts)
            
            if not group_has_file:
                # 如果该组没有文件，随机选一个扩展名生成一个文件
                target_ext = random.choice(exts)
                filename = generate_random_filename(root, target_ext)
                file_path = os.path.join(root, filename)
                
                if create_dummy_file(file_path):
                    print(f"  + 补充 {group_name} 类文件: {filename}")
                    total_created += 1
            else:
                # 如果该组已有文件，有一定概率（30%）再增加一个不同后缀的，增加丰富度
                if random.random() < 0.3:
                     # 找出该组中还不存在的后缀
                     missing_exts = [e for e in exts if e not in existing_exts]
                     if missing_exts:
                         target_ext = random.choice(missing_exts)
                         filename = generate_random_filename(root, target_ext)
                         file_path = os.path.join(root, filename)
                         if create_dummy_file(file_path):
                            print(f"  + (丰富度) 补充 {target_ext} 文件: {filename}")
                            total_created += 1

    print(f"\n任务完成！总共递归新创建了 {total_created} 个文件。")

if __name__ == "__main__":
    populate_recursive()
