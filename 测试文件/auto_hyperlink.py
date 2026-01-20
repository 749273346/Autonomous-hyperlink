import os
import re
import time
import sys
from datetime import datetime
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
import pythoncom
import win32com.client

# Determine the directory to watch:
if getattr(sys, 'frozen', False):
    WATCH_DIR = os.path.dirname(sys.executable)
else:
    WATCH_DIR = os.path.dirname(os.path.abspath(__file__))

RETRY_DELAY = 1
RETRIES = 8

CATEGORY_PREFIX_MAP = {
    "上级文": "SJW",
    "其他": "QT",
    "事项通知": "SXTZ",
}

def _find_year_two_digits(path):
    parts = os.path.normpath(path).split(os.sep)
    for part in parts:
        if part in ("25", "26"):
            return part
    return None

def _category_label_from_path(path):
    rel = os.path.relpath(path, WATCH_DIR)
    parts = rel.split(os.sep)
    if len(parts) < 2:
        return None
    folder = parts[0]
    label = re.sub(r"^\s*\d+\s*[-－]\s*", "", folder).strip()
    return label or folder

def _extract_doc_no(text):
    stem = os.path.splitext(text)[0]
    patterns = [
        r"([^\s（）()]*?[〔【]\s*20\d{2}\s*[〕】]\s*\d+\s*号)",
        r"([^\s（）()]*?\[\s*20\d{2}\s*\]\s*\d+\s*号)",
    ]
    for pat in patterns:
        m = re.search(pat, stem)
        if m:
            return re.sub(r"\s+", "", m.group(1)).strip()
    m = re.search(r"[（(]([^）)]+号)[)）]", stem)
    if m:
        inner = m.group(1).strip()
        if re.search(r"20\d{2}", inner):
            return re.sub(r"\s+", "", inner).strip()
    return ""

def _normalize_doc_no(v):
    return re.sub(r"\s+", "", str(v or "")).strip()

def _find_sheet_index_com(wb, category_label):
    if not category_label:
        return None
    candidates = []
    # wb.Sheets is 1-based collection, but we can iterate
    for i in range(1, wb.Sheets.Count + 1):
        name = wb.Sheets(i).Name
        if name == "Sheet1":
            continue
        if name == category_label:
            return i
        if category_label in name:
            candidates.append((len(name), i))
    if candidates:
        candidates.sort()
        return candidates[0][1]
    return None

def _find_header_map_com(ws):
    # Scan first 50 rows for header
    max_rows = 50
    used_rows = ws.UsedRange.Rows.Count
    limit = min(max_rows, used_rows)
    
    header_row = -1
    hm = {}
    
    # Iterate rows (1-based)
    for r in range(1, limit + 1):
        # Read entire row efficiently? 
        # For simplicity, read specific cells or UsedRange intersection
        # Let's just read the first 20 columns to find "序号" and "文件名"
        row_vals = []
        for c in range(1, 20):
            val = str(ws.Cells(r, c).Value or "").strip()
            row_vals.append(val)
        
        if "序号" in row_vals and "文件名" in row_vals:
            header_row = r
            for idx, val in enumerate(row_vals):
                if val:
                    hm[val] = idx + 1 # Store 1-based column index
            if "存放位置" in hm and "存盒位置" not in hm:
                hm["存盒位置"] = hm["存放位置"]
            break
            
    return header_row, hm

def _find_existing_row_com(ws, header_row, hm, doc_no, rel_path, filename):
    # Scan from header_row + 1 to end
    last_row = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    # If UsedRange is unreliable, we might need to be careful, but usually it's fine for reading.
    
    target_doc_no = _normalize_doc_no(doc_no)
    
    doc_col = hm.get("文号")
    file_col = hm.get("文件名")
    
    if not file_col: return None
    
    # To optimize, we could read the whole range into a list of lists, but for now loop is safer for COM logic simplicity
    # Reading huge sheets cell-by-cell is slow via COM. 
    # Optimization: Read the column into a variant array.
    
    # However, for simplicity and robustness in this script, simple iteration is okay if rows < 1000.
    # Let's try to be slightly efficient.
    
    for r in range(header_row + 1, last_row + 2): # Go a bit beyond
        # Check Doc No
        if doc_col and target_doc_no:
            val = str(ws.Cells(r, doc_col).Value or "").strip()
            if _normalize_doc_no(val) == target_doc_no:
                return r
        
        # Check Filename/Path
        val_file = str(ws.Cells(r, file_col).Value or "").strip()
        if not val_file:
            continue # Skip empty file cells? Or is it end of data?
            # Don't stop, there might be gaps.
            
        if rel_path and rel_path in val_file:
            return r
        if filename and filename in val_file:
            return r
            
    return None

def _is_row_empty_com(ws, r, content_cols):
    for c in content_cols:
        val = str(ws.Cells(r, c).Value or "").strip()
        if val:
            return False
    return True

def _find_first_empty_row_com(ws, header_row, hm):
    content_cols = [hm[k] for k in ("收文日期", "文号", "文件名", "自编号") if k in hm]
    if not content_cols:
        return header_row + 1
        
    last_row = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    
    for r in range(header_row + 1, last_row + 2):
        if _is_row_empty_com(ws, r, content_cols):
            return r
    return last_row + 1

def _next_seq_com(ws, header_row, hm, target_row):
    # 如果是紧接着表头的第一行，强制序号为1
    if target_row == header_row + 1:
        return 1

    seq_col = hm.get("序号")
    if not seq_col: return 1
    
    # Check if target_row already has a sequence (if we are overwriting a blank row inside)
    # But usually we want max seq + 1
    
    max_seq = 0
    last_row = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    
    for r in range(header_row + 1, last_row + 1):
        if r == target_row: continue
        val = str(ws.Cells(r, seq_col).Value or "").strip()
        if val:
            try:
                v = int(float(val))
                max_seq = max(max_seq, v)
            except:
                pass
    return max_seq + 1

def _wait_for_file_unlock(filepath, timeout=3):
    """尝试等待文件解锁"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # 尝试以追加模式打开文件，如果成功说明未被锁定
            with open(filepath, 'a+'):
                pass
            return True
        except PermissionError:
            time.sleep(0.5)
        except Exception:
            # 其他错误（如文件不存在）暂不处理，交给后续逻辑
            return True
    return False

def _infer_date_format_com(ws, header_row, date_col):
    if not date_col: return "dot"
    last_row = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    for r in range(last_row, header_row, -1):
        val = str(ws.Cells(r, date_col).Value or "").strip()
        if val:
            if "/" in val: return "slash"
            if "." in val: return "dot"
    return "dot"

def _generate_self_id_com(ws, header_row, self_col, year_full, category_label):
    prefix = CATEGORY_PREFIX_MAP.get(category_label, "QT")
    if not self_col: return f"{prefix}-{year_full}-1"
    
    max_num = 0
    pattern = re.compile(rf"^{re.escape(prefix)}-{year_full}-(\d+)$")
    
    last_row = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    for r in range(header_row + 1, last_row + 1):
        val = str(ws.Cells(r, self_col).Value or "").strip()
        m = pattern.match(val)
        if m:
            try:
                num = int(m.group(1))
                max_num = max(max_num, num)
            except:
                pass
    return f"{prefix}-{year_full}-{max_num + 1}"

def _infer_last_nonempty_com(ws, header_row, col):
    last_row = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    for r in range(last_row, header_row, -1):
        val = str(ws.Cells(r, col).Value or "").strip()
        if val:
            return val
    return ""

def _update_workbook(excel_path, file_path):
    # COM Implementation
    
    # 1. 检查文件锁定状态，给予一定的缓冲时间让之前的进程释放
    if not _wait_for_file_unlock(excel_path, timeout=5):
        print(f"文件仍被锁定，跳过本次更新: {excel_path}")
        return

    try:
        pythoncom.CoInitialize()
    except:
        pass
        
    app = None
    wb = None
    try:
        try:
            app = win32com.client.DispatchEx("Excel.Application")
        except Exception as e:
            # Fallback to standard Dispatch if Ex fails
            app = win32com.client.Dispatch("Excel.Application")
            
        app.Visible = False
        app.DisplayAlerts = False
        
        # Check if workbook is already open?
        # DispatchEx creates a new instance usually.
        
        try:
            wb = app.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=False)
        except Exception as e:
            print(f"Failed to open workbook: {e}")
            return

        category_label = _category_label_from_path(file_path)
        sheet_index = _find_sheet_index_com(wb, category_label)
        
        if sheet_index is None:
            print(f"找不到对应工作表: {category_label}")
            return
            
        ws = wb.Worksheets(sheet_index)
        header_row, hm = _find_header_map_com(ws)
        
        if not hm:
            print(f"工作表缺少表头: {ws.Name}")
            return
            
        filename = os.path.basename(file_path)
        rel_path = os.path.relpath(file_path, WATCH_DIR)
        doc_no = _extract_doc_no(filename)
        
        existing_row = _find_existing_row_com(ws, header_row, hm, doc_no, rel_path, filename)
        
        if existing_row:
            target_row = existing_row
            # Update Filename
            cell = ws.Cells(target_row, hm["文件名"])
            cell.Value = filename
            
            # Re-add hyperlink (safe to delete old one first if needed, but Add usually overwrites or adds)
            # To be clean, delete existing hyperlinks on that cell
            try:
                cell.Hyperlinks.Delete()
            except:
                pass
                
            ws.Hyperlinks.Add(Anchor=cell, Address=rel_path, TextToDisplay=filename)
            
            if "备注" in hm:
                ws.Cells(target_row, hm["备注"]).Value = ""
                
        else:
            target_row = _find_first_empty_row_com(ws, header_row, hm)
            
            # Prepare data
            year_two = _find_year_two_digits(file_path) or "25"
            year_full = f"20{year_two}"
            
            # DATE FIX: Use current time instead of file mtime
            date_fmt = _infer_date_format_com(ws, header_row, hm.get("收文日期"))
            if date_fmt == "slash":
                received_date = datetime.now().strftime("%Y/%m/%d")
            else:
                received_date = datetime.now().strftime("%Y.%m.%d")
                
            self_id = _generate_self_id_com(ws, header_row, hm.get("自编号"), year_full, category_label)
            transmit = ""
            if "传阅方式" in hm:
                transmit = _infer_last_nonempty_com(ws, header_row, hm["传阅方式"])
            
            seq = _next_seq_com(ws, header_row, hm, target_row)
            
            # Write Data
            if "序号" in hm: ws.Cells(target_row, hm["序号"]).Value = seq
            if "收文日期" in hm: ws.Cells(target_row, hm["收文日期"]).Value = received_date
            if "文号" in hm: ws.Cells(target_row, hm["文号"]).Value = doc_no
            if "自编号" in hm: ws.Cells(target_row, hm["自编号"]).Value = self_id
            if "传阅方式" in hm: ws.Cells(target_row, hm["传阅方式"]).Value = transmit
            if "存盒位置" in hm: ws.Cells(target_row, hm["存盒位置"]).Value = ""
            if "备注" in hm: ws.Cells(target_row, hm["备注"]).Value = ""
            
            # Filename & Link
            cell = ws.Cells(target_row, hm["文件名"])
            cell.Value = filename
            ws.Hyperlinks.Add(Anchor=cell, Address=rel_path, TextToDisplay=filename)

        wb.Save()
        print(f"已更新: {excel_path} (Row {target_row})")
        
    except Exception as e:
        print(f"更新失败: {e}")
    finally:
        # 强制释放资源
        if wb:
            try:
                wb.Close(SaveChanges=True)
            except: pass
            del wb
        if app:
            try:
                app.Quit()
            except: pass
            del app
        try:
            pythoncom.CoUninitialize()
        except: pass

def _excel_path_for_year(year_two_digits):
    excel_name = f"20{year_two_digits}工区收文目录.xls"
    p1 = os.path.join(WATCH_DIR, excel_name)
    if os.path.exists(p1):
        return p1
    return None

class AutoHyperlinkHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        self._handle(event.src_path, "created")

    def on_moved(self, event):
        if event.is_directory:
            return
        self._handle(event.dest_path, "moved")
        
    def _handle(self, file_path, kind):
        filename = os.path.basename(file_path)
        if filename.startswith("~$") or filename.lower().endswith(".tmp") or \
           filename.lower() in ["autohyperlink.exe", "auto_hyperlink.py", "auto_hyperlink.spec"]:
            return
        if re.match(r"^20\d{2}工区收文目录\.xls$", filename):
            return
            
        year = _find_year_two_digits(file_path)
        if not year:
            print(f"跳过（无法识别年份目录 25/26）: {file_path}")
            return
            
        excel_path = _excel_path_for_year(year)
        if not excel_path:
            print(f"跳过（找不到收文目录表）: 20{year}工区收文目录.xls")
            return
            
        print(f"File {kind}: {file_path}")
        
        # Retry loop
        for attempt in range(RETRIES):
            try:
                _update_workbook(excel_path, file_path)
                return
            except Exception as e:
                print(f"Attempt {attempt+1} failed: {e}")
                time.sleep(RETRY_DELAY)
        print("多次重试仍失败。")

def main():
    if not os.path.exists(WATCH_DIR):
        print(f"目录不存在: {WATCH_DIR}")
        return
        
    handler = AutoHyperlinkHandler()
    observer = Observer()
    observer.schedule(handler, WATCH_DIR, recursive=True)
    observer.start()
    print(f"Monitoring {WATCH_DIR} for changes (COM Mode)...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    main()
