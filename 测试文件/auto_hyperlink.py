import os
import re
import shutil
import time
from datetime import datetime

import xlrd
import xlwt
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
import sys
from xlutils.copy import copy

# Determine the directory to watch:
# If frozen (exe), use the executable's directory.
# If script, use the script's directory.
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


def _find_sheet_index(wb, category_label):
    if not category_label:
        return None
    candidates = []
    for i, name in enumerate(wb.sheet_names()):
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


def _find_header_row(sheet, max_rows=50):
    for r in range(min(max_rows, sheet.nrows)):
        row = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
        if "序号" in row and "文件名" in row:
            return r
    return None


def _header_map(sheet, header_row):
    header = [str(sheet.cell_value(header_row, c)).strip() for c in range(sheet.ncols)]
    mapping = {v: i for i, v in enumerate(header) if v}
    if "存放位置" in mapping and "存盒位置" not in mapping:
        mapping["存盒位置"] = mapping["存放位置"]
    return mapping


def _last_data_row(sheet, header_row, seq_col):
    for r in range(sheet.nrows - 1, header_row, -1):
        v = sheet.cell_value(r, seq_col)
        if str(v).strip():
            try:
                return r, int(float(v))
            except Exception:
                return r, None
    return header_row, 0


def _next_seq(sheet, header_row, seq_col):
    _, last_seq = _last_data_row(sheet, header_row, seq_col)
    return (last_seq or 0) + 1


def _is_blank_cell(v):
    return not str(v).strip()


def _find_first_empty_content_row(sheet, header_row, hm):
    content_cols = [hm[k] for k in ("收文日期", "文号", "文件名", "自编号") if k in hm]
    for r in range(header_row + 1, sheet.nrows):
        if all(_is_blank_cell(sheet.cell_value(r, c)) for c in content_cols):
            return r
    return sheet.nrows


def _next_seq_by_content(sheet, header_row, hm):
    seq_col = hm["序号"]
    content_cols = [hm[k] for k in ("收文日期", "文号", "文件名", "自编号") if k in hm]
    last_seq = 0
    for r in range(header_row + 1, sheet.nrows):
        if all(_is_blank_cell(sheet.cell_value(r, c)) for c in content_cols):
            continue
        v = str(sheet.cell_value(r, seq_col)).strip()
        if not v:
            continue
        try:
            last_seq = max(last_seq, int(float(v)))
        except Exception:
            continue
    return last_seq + 1


def _infer_date_format(sheet, header_row, date_col):
    for r in range(sheet.nrows - 1, header_row, -1):
        v = str(sheet.cell_value(r, date_col)).strip()
        if v:
            if "/" in v:
                return "slash"
            if "." in v:
                return "dot"
            return "dot"
    return "dot"


def _format_received_date(file_path, fmt, year_full):
    dt = datetime.fromtimestamp(os.path.getmtime(file_path))
    return f"{dt.year}.{dt.month:02d}.{dt.day:02d}"


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


def _generate_self_id(sheet, header_row, self_col, year_full, category_label):
    prefix = CATEGORY_PREFIX_MAP.get(category_label, "QT")
    max_num = 0
    pattern = re.compile(rf"^{re.escape(prefix)}-{year_full}-(\d+)$")

    for r in range(header_row + 1, sheet.nrows):
        v = str(sheet.cell_value(r, self_col)).strip()
        m = pattern.match(v)
        if m:
            try:
                num = int(m.group(1))
                max_num = max(max_num, num)
            except ValueError:
                pass
    return f"{prefix}-{year_full}-{max_num + 1}"


def _infer_last_nonempty(sheet, header_row, col):
    for r in range(sheet.nrows - 1, header_row, -1):
        v = str(sheet.cell_value(r, col)).strip()
        if v:
            return v
    return ""


def _escape_excel_formula_str(s):
    return str(s).replace('"', '""')


def _hyperlink_formula(target, display):
    t = _escape_excel_formula_str(target)
    d = _escape_excel_formula_str(display)
    return xlwt.Formula(f'HYPERLINK("{t}","{d}")')


def _file_uri_from_path(p):
    ap = os.path.abspath(p)
    uri = ap.replace("\\", "/")
    if re.match(r"^[A-Za-z]:/", uri):
        return "file:///" + uri
    return "file://" + uri


def _try_add_hyperlink_com(excel_path, sheet_name, row0, col0, address, text):
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return False

    def _do(prog_id):
        pythoncom.CoInitialize()
        app = None
        wb = None
        try:
            app = win32com.client.DispatchEx(prog_id)
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=False)
            ws = wb.Worksheets(sheet_name)
            cell = ws.Cells(row0 + 1, col0 + 1)
            try:
                cell.Hyperlinks.Delete()
            except Exception:
                pass
            ws.Hyperlinks.Add(Anchor=cell, Address=address, TextToDisplay=text)
            wb.Save()
            wb.Close(SaveChanges=True)
            wb = None
            app.Quit()
            app = None
            return True
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=True)
            except Exception:
                pass
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    for prog_id in ("Excel.Application", "ket.Application"):
        try:
            if _do(prog_id):
                return True
        except Exception:
            continue
    return False


def _normalize_doc_no(v):
    return re.sub(r"\s+", "", str(v or "")).strip()


def _find_existing_row(sheet, header_row, doc_col, file_col, doc_no, rel_path, filename):
    target_doc_no = _normalize_doc_no(doc_no)
    for r in range(header_row + 1, sheet.nrows):
        if target_doc_no:
            existing_doc_no = _normalize_doc_no(sheet.cell_value(r, doc_col))
            if existing_doc_no and existing_doc_no == target_doc_no:
                return r
        v = str(sheet.cell_value(r, file_col)).strip()
        if not v:
            continue
        if rel_path and rel_path in v:
            return r
        if filename and filename in v:
            return r
    return None


def _update_workbook(excel_path, file_path):
    temp_read_path = excel_path + ".read.tmp"
    temp_write_path = excel_path + ".write.tmp"
    shutil.copy2(excel_path, temp_read_path)
    try:
        rb = xlrd.open_workbook(temp_read_path, formatting_info=False)
        category_label = _category_label_from_path(file_path)
        sheet_index = _find_sheet_index(rb, category_label)
        if sheet_index is None:
            raise ValueError(f"找不到对应工作表: {category_label}")
        rs = rb.sheet_by_index(sheet_index)
        header_row = _find_header_row(rs)
        if header_row is None:
            raise ValueError(f"工作表缺少表头: {rb.sheet_names()[sheet_index]}")
        hm = _header_map(rs, header_row)
        required = ["序号", "收文日期", "文号", "文件名", "自编号", "传阅方式", "存盒位置", "备注"]
        missing = [k for k in required if k not in hm]
        if missing:
            raise ValueError(f"工作表列缺失: {rb.sheet_names()[sheet_index]} {missing}")

        rel_path = os.path.relpath(file_path, WATCH_DIR)
        filename = os.path.basename(file_path)
        ws = copy(rb)
        sheet = ws.get_sheet(sheet_index)

        doc_no = _extract_doc_no(filename)
        existing_row = _find_existing_row(
            rs,
            header_row,
            hm["文号"],
            hm["文件名"],
            doc_no,
            rel_path,
            filename,
        )
        sheet_name = rb.sheet_names()[sheet_index]
        if existing_row is not None:
            sheet.write(existing_row, hm["文件名"], filename)
            sheet.write(existing_row, hm["备注"], "")
            ws.save(temp_write_path)
            os.replace(temp_write_path, excel_path)
            if not _try_add_hyperlink_com(excel_path, sheet_name, existing_row, hm["文件名"], rel_path, filename):
                rb2 = xlrd.open_workbook(excel_path, formatting_info=False)
                ws2 = copy(rb2)
                sheet2 = ws2.get_sheet(sheet_index)
                sheet2.write(existing_row, hm["文件名"], _hyperlink_formula(rel_path, filename))
                ws2.save(temp_write_path)
                os.replace(temp_write_path, excel_path)
            return

        target_row = _find_first_empty_content_row(rs, header_row, hm)
        year_two = _find_year_two_digits(file_path) or ""
        year_full = f"20{year_two}" if year_two else ""
        date_fmt = _infer_date_format(rs, header_row, hm["收文日期"])
        received_date = _format_received_date(file_path, date_fmt, year_full)
        self_id = _generate_self_id(rs, header_row, hm["自编号"], year_full, category_label)
        transmit = _infer_last_nonempty(rs, header_row, hm["传阅方式"])

        existing_seq = str(rs.cell_value(target_row, hm["序号"])).strip() if target_row < rs.nrows else ""
        if existing_seq:
            try:
                seq_value = int(float(existing_seq))
            except Exception:
                seq_value = _next_seq_by_content(rs, header_row, hm)
        else:
            seq_value = _next_seq_by_content(rs, header_row, hm)

        sheet.write(target_row, hm["序号"], seq_value)
        sheet.write(target_row, hm["收文日期"], received_date)
        sheet.write(target_row, hm["文号"], doc_no)
        sheet.write(target_row, hm["文件名"], filename)
        sheet.write(target_row, hm["自编号"], self_id)
        sheet.write(target_row, hm["传阅方式"], transmit)
        sheet.write(target_row, hm["存盒位置"], "")
        sheet.write(target_row, hm["备注"], "")

        ws.save(temp_write_path)
        os.replace(temp_write_path, excel_path)
        if not _try_add_hyperlink_com(excel_path, sheet_name, target_row, hm["文件名"], rel_path, filename):
            rb2 = xlrd.open_workbook(excel_path, formatting_info=False)
            ws2 = copy(rb2)
            sheet2 = ws2.get_sheet(sheet_index)
            sheet2.write(target_row, hm["文件名"], _hyperlink_formula(rel_path, filename))
            ws2.save(temp_write_path)
            os.replace(temp_write_path, excel_path)
    finally:
        if os.path.exists(temp_read_path):
            try:
                os.remove(temp_read_path)
            except Exception:
                pass
        if os.path.exists(temp_write_path):
            try:
                os.remove(temp_write_path)
            except Exception:
                pass


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

    def on_deleted(self, event):
        if event.is_directory:
            return
        print(f"File deleted: {event.src_path}")

    def _handle(self, file_path, kind):
        filename = os.path.basename(file_path)
        # Filter temp files and the executable/script itself
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
        for attempt in range(RETRIES):
            try:
                _update_workbook(excel_path, file_path)
                print(f"已更新: {excel_path}")
                return
            except PermissionError:
                time.sleep(RETRY_DELAY)
            except OSError:
                time.sleep(RETRY_DELAY)
            except Exception as e:
                print(f"更新失败: {e}")
                return
        print(f"多次重试仍失败（文件可能正在被占用）: {excel_path}")


def main():
    if not os.path.exists(WATCH_DIR):
        raise SystemExit(f"目录不存在: {WATCH_DIR}")
    handler = AutoHyperlinkHandler()
    observer = Observer()
    observer.schedule(handler, WATCH_DIR, recursive=True)
    observer.start()
    print(f"Monitoring {WATCH_DIR} for changes...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()
