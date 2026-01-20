import os
import re
import shutil
import time
from datetime import datetime

import xlrd
import xlwt
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
from xlutils.copy import copy

WATCH_DIR = r"e:\QC-攻关小组\正在进行项目\自主超链接\Autonomous-hyperlink\测试文件"
RETRY_DELAY = 1
RETRIES = 8


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
    if fmt == "slash":
        year = year_full or str(dt.year)
        return f"{year}/{dt.month:02d}/{dt.day:02d}"
    return f"{dt.month}.{dt.day}"


def _extract_doc_no(text):
    stem = os.path.splitext(text)[0]
    patterns = [
        r"([^\s（）()]*?[〔【]\s*20\d{2}\s*[〕】]\s*\d+\s*号)",
        r"([^\s（）()]*?\[\s*20\d{2}\s*\]\s*\d+\s*号)",
    ]
    for pat in patterns:
        m = re.search(pat, stem)
        if m:
            return m.group(1).strip()
    m = re.search(r"[（(]([^）)]+号)[)）]", stem)
    if m:
        inner = m.group(1).strip()
        if re.search(r"20\d{2}", inner):
            return inner
    return ""


def _infer_self_id_pattern(sheet, header_row, self_col, year_full):
    prefix = ""
    width = 3
    for r in range(header_row + 1, sheet.nrows):
        v = str(sheet.cell_value(r, self_col)).strip()
        if not v:
            continue
        m = re.match(r"^([A-Za-z]+)-(\d{4})-(\d+)$", v)
        if m:
            prefix = m.group(1)
            width = len(m.group(3))
            break
    if not prefix:
        prefix = "AUTO"
    max_num = 0
    for r in range(header_row + 1, sheet.nrows):
        v = str(sheet.cell_value(r, self_col)).strip()
        m = re.match(rf"^{re.escape(prefix)}-(\d{{4}})-(\d+)$", v)
        if not m:
            continue
        if m.group(1) != year_full:
            continue
        try:
            num = int(m.group(2))
        except Exception:
            continue
        max_num = max(max_num, num)
    return prefix, width, max_num + 1


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


def _find_existing_row(sheet, header_row, file_col, rel_path, filename):
    for r in range(header_row + 1, sheet.nrows):
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

        existing_row = _find_existing_row(rs, header_row, hm["文件名"], rel_path, filename)
        if existing_row is not None:
            existing_display = str(rs.cell_value(existing_row, hm["文件名"])).strip() or rel_path
            sheet.write(existing_row, hm["文件名"], _hyperlink_formula(rel_path, existing_display))
            ws.save(temp_write_path)
            os.replace(temp_write_path, excel_path)
            return

        target_row, _ = _last_data_row(rs, header_row, hm["序号"])
        target_row += 1
        year_two = _find_year_two_digits(file_path) or ""
        year_full = f"20{year_two}" if year_two else ""
        date_fmt = _infer_date_format(rs, header_row, hm["收文日期"])
        received_date = _format_received_date(file_path, date_fmt, year_full)
        doc_no = _extract_doc_no(filename)
        prefix, width, next_num = _infer_self_id_pattern(rs, header_row, hm["自编号"], year_full)
        self_id = f"{prefix}-{year_full}-{str(next_num).zfill(width)}" if year_full else f"{prefix}-{str(next_num).zfill(width)}"
        transmit = _infer_last_nonempty(rs, header_row, hm["传阅方式"])

        sheet.write(target_row, hm["序号"], _next_seq(rs, header_row, hm["序号"]))
        sheet.write(target_row, hm["收文日期"], received_date)
        sheet.write(target_row, hm["文号"], doc_no)
        sheet.write(target_row, hm["文件名"], _hyperlink_formula(rel_path, rel_path))
        sheet.write(target_row, hm["自编号"], self_id)
        sheet.write(target_row, hm["传阅方式"], transmit)
        sheet.write(target_row, hm["存盒位置"], "")
        sheet.write(target_row, hm["备注"], "")

        ws.save(temp_write_path)
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
        if filename.startswith("~$") or filename.lower().endswith(".tmp"):
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
