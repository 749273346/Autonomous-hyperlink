import os
import shutil
import random
import time

TEST_DIR = r"e:\QC-攻关小组\正在进行项目\自主超链接\Autonomous-hyperlink\测试文件"

# 定义文件类型组
TYPE_GROUPS = {
    "Word": ['.doc', '.docx', '.wps'],
    "Excel": ['.xls', '.xlsx', '.et'],
    "PPT": ['.ppt', '.pptx', '.dps'],
    "Other": ['.pdf', '.txt']
}

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

# 部门
DEPARTMENTS = ["供电", "工务", "电务", "车务", "机务", "客运", "人事", "财务"]

def generate_random_filename(folder_name, ext, year_folder):
    parts = folder_name.split(os.sep)
    prefix = parts[-1]
    if prefix in ["25", "26"]:
        prefix = parts[-2]
    
    if '-' in prefix:
        prefix = prefix.split('-')[-1]
        
    topic = random.choice(TOPICS)
    dept = random.choice(DEPARTMENTS)
    year = "2025" if year_folder == "25" else "2026"
    
    name_formats = [
        f"{prefix}_{year}_{topic}",
        f"（{dept}函〔{year}〕{random.randint(1, 100)}号）{topic}",
        f"{dept}段{topic}",
        f"{year}年{dept}段{prefix}材料汇总"
    ]
    
    base_name = random.choice(name_formats)
    return f"{base_name}{ext}"

def create_dummy_file(file_path):
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(f"This is a dummy file.\nFile: {file_path}\nTime: {time.ctime()}\n")
        return True
    except Exception as e:
        print(f"Error creating {file_path}: {e}")
        return False

def cleanup_and_populate():
    print(f"开始清理并填充: {TEST_DIR}")
    if not os.path.exists(TEST_DIR):
        print("目录不存在")
        return

    # 遍历一级目录
    for item in os.listdir(TEST_DIR):
        item_path = os.path.join(TEST_DIR, item)
        
        # 只处理目录，跳过文件（如 Excel 表格本身）
        if os.path.isdir(item_path) and not item.startswith('.'):
            # 排除 __pycache__
            if item == "__pycache__":
                continue
                
            print(f"处理文件夹: {item}")
            
            # 1. 清理一级目录下的所有文件（只保留文件夹）
            for sub_item in os.listdir(item_path):
                sub_item_path = os.path.join(item_path, sub_item)
                if os.path.isfile(sub_item_path):
                    try:
                        os.remove(sub_item_path)
                        print(f"  - 已删除散乱文件: {sub_item}")
                    except Exception as e:
                        print(f"  ! 删除失败: {sub_item} - {e}")
            
            # 2. 确保 25 和 26 文件夹存在并有文件
            for year_folder in ["25", "26"]:
                target_dir = os.path.join(item_path, year_folder)
                
                if not os.path.exists(target_dir):
                    os.makedirs(target_dir)
                    print(f"  + 创建目录: {year_folder}")
                
                # 检查文件数量和类型
                existing_files = [f for f in os.listdir(target_dir) if os.path.isfile(os.path.join(target_dir, f))]
                existing_exts = {os.path.splitext(f)[1] for f in existing_files}
                
                print(f"    检查 {year_folder} (现有 {len(existing_files)} 个文件)...")
                
                # 确保每种类型都有
                for group_name, exts in TYPE_GROUPS.items():
                    if not any(ext in existing_exts for ext in exts):
                        # 缺失该类型，补充一个
                        ext = random.choice(exts)
                        filename = generate_random_filename(item, ext, year_folder)
                        if create_dummy_file(os.path.join(target_dir, filename)):
                            print(f"      + 补充 {group_name}: {filename}")
                
                # 确保总数至少有 5 个
                current_count = len(os.listdir(target_dir))
                while current_count < 5:
                    ext = random.choice([e for g in TYPE_GROUPS.values() for e in g])
                    filename = generate_random_filename(item, ext, year_folder)
                    if create_dummy_file(os.path.join(target_dir, filename)):
                        print(f"      + 补充数量: {filename}")
                        current_count += 1

    print("\n所有任务完成！已清理一级目录散乱文件，并保证 25/26 目录内容丰富。")

if __name__ == "__main__":
    cleanup_and_populate()
