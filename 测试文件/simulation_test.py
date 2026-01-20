import os
import time
import xlrd
import subprocess
import shutil
import threading
import sys

# 配置路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_FILE_SUBDIR = r"10-事项通知\25"
TEST_FILENAME = f"自动化测试通知_（测函〔2025〕{int(time.time())}号）.doc"
TEST_FILE_PATH = os.path.join(BASE_DIR, TEST_FILE_SUBDIR, TEST_FILENAME)
EXCEL_NAME = "2025工区收文目录.xls"
EXCEL_PATH = os.path.join(BASE_DIR, EXCEL_NAME)
MONITOR_SCRIPT = "auto_hyperlink.py"

def _extract_doc_no_simple(filename):
    import re
    m = re.search(r"[（(]([^）)]+号)[)）]", filename)
    if m:
        return m.group(1).strip()
    return ""

def print_status(msg, status="INFO"):
    print(f"[{status}] {msg}")

def ensure_monitor_running():
    print_status(f"正在启动监控脚本 {MONITOR_SCRIPT} ...")
    # 捕获输出以便调试
    process = subprocess.Popen(
        [sys.executable, "-u", MONITOR_SCRIPT], 
        cwd=BASE_DIR, 
        stdout=subprocess.PIPE, 
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1
    )
    print_status("监控脚本已启动 (PID: {})".format(process.pid), "SUCCESS")
    
    # 启动线程读取输出
    def read_output(pipe, prefix):
        for line in iter(pipe.readline, ''):
            print(f"[MONITOR-{prefix}] {line.strip()}")
    
    t_out = threading.Thread(target=read_output, args=(process.stdout, "OUT"))
    t_out.daemon = True
    t_out.start()
    
    t_err = threading.Thread(target=read_output, args=(process.stderr, "ERR"))
    t_err.daemon = True
    t_err.start()
    
    return process

def create_test_file():
    target_dir = os.path.dirname(TEST_FILE_PATH)
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
    
    # 确保文件是新的
    if os.path.exists(TEST_FILE_PATH):
        print_status(f"清理旧的测试文件: {TEST_FILE_PATH}")
        os.remove(TEST_FILE_PATH)
        time.sleep(2) # 等待 watchdog 处理删除事件

    print_status(f"创建测试文件: {TEST_FILE_PATH}")
    with open(TEST_FILE_PATH, "w") as f:
        f.write("This is a test file.")
    print_status("测试文件创建完成", "SUCCESS")

def verify_excel():
    print_status(f"正在检查 Excel 文件: {EXCEL_PATH}")
    if not os.path.exists(EXCEL_PATH):
        print_status("找不到 Excel 文件！", "FAIL")
        return False

    try:
        # 给一点时间让 Excel 写入完成
        time.sleep(2) 
        rb = xlrd.open_workbook(EXCEL_PATH, formatting_info=True)
        
        print_status(f"Excel 包含的工作表: {rb.sheet_names()}", "DEBUG")
        
        target_sheet_name = "事项通知" 
        sheet = None
        for s in rb.sheets():
            if target_sheet_name in s.name:
                sheet = s
                print_status(f"找到目标工作表: {s.name}", "DEBUG")
                break
        
        if not sheet:
            print_status(f"未找到包含'{target_sheet_name}'的工作表，将检查所有工作表...", "WARN")
            sheets_to_check = rb.sheets()
        else:
            sheets_to_check = [sheet]

        found = False
        for s in sheets_to_check:
            # 全表检查（跳过表头）
            print_status(f"检查工作表 '{s.name}' (共 {s.nrows} 行)...", "DEBUG")
            
            start_row = 1 # 跳过 Row 0 (可能是标题) 和 Row 1 (表头)
            if s.nrows > 1:
                # 打印表头
                try:
                    header = [str(s.cell_value(1, c)) for c in range(s.ncols)]
                    print_status(f"表头 (Row 2): {header}", "DEBUG")
                except:
                    pass

            for r in range(start_row, s.nrows):
                row_values = [str(s.cell_value(r, c)).strip() for c in range(s.ncols)]
                # 跳过空行
                if not any(row_values):
                    continue
                
                row_str = " ".join(row_values)
                # 打印所有非空行
                # print_status(f"Row {r+1}: {row_str}", "DEBUG")
                
                # 匹配逻辑修改：优先匹配文号，因为 xlrd 可能读不到公式计算后的文件名
                expected_doc_no = _extract_doc_no_simple(TEST_FILENAME)
                
                # 检查文号是否在行数据中
                if expected_doc_no and expected_doc_no in row_str:
                     print_status(f"Row {r+1}: {row_values}", "DEBUG") 
                     print_status(f"在工作表 '{s.name}' 第 {r+1} 行找到测试记录（通过文号匹配）！", "SUCCESS")
                     print_status(f"行数据: {row_values}", "INFO")
                     
                     today_dot = time.strftime("%Y.%m.%d")
                     if today_dot in row_str:
                         print_status("日期验证通过", "SUCCESS")
                     else:
                         print_status(f"日期验证警告: 未找到 {today_dot}", "WARN")

                     found = True
                     break
                
                # 保留原来的文件名匹配作为备选
                if TEST_FILENAME in row_str:
                     print_status(f"Row {r+1}: {row_values}", "DEBUG") 
                     print_status(f"在工作表 '{s.name}' 第 {r+1} 行找到测试记录！", "SUCCESS")
                     found = True
                     break
            
            if found: break
            
        if not found:
            print_status("在 Excel 中未找到测试文件的记录！", "FAIL")
            return False
            
        return True

    except Exception as e:
        print_status(f"读取 Excel 失败: {e}", "FAIL")
        return False

def main():
    print("========================================")
    print("      真实环境模拟测试脚本启动 (调试版)   ")
    print("========================================")
    
    monitor_process = ensure_monitor_running()
    
    try:
        # 等待监控完全启动
        time.sleep(5)
        
        create_test_file()
        
        print_status("等待监控脚本处理 (10秒)...")
        time.sleep(10)
        
        if verify_excel():
            print("\n========================================")
            print("             测试通过！                 ")
            print("========================================")
        else:
            print("\n========================================")
            print("             测试失败！                 ")
            print("========================================")
            
    finally:
        print("\n[CLEANUP] 正在关闭监控进程...")
        monitor_process.terminate()
        print(f"[CLEANUP] 测试文件保留在: {TEST_FILE_PATH}")

if __name__ == "__main__":
    main()
