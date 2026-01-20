import os
import shutil
import time
from datetime import datetime

import pythoncom
import win32com.client

import auto_hyperlink as ah


def _open_excel():
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    return app


def _sheet_and_headers(wb, category_label):
    sheet_index = ah._find_sheet_index_com(wb, category_label)
    if sheet_index is None:
        raise RuntimeError(f"找不到工作表: {category_label}")
    ws = wb.Worksheets(sheet_index)
    header_row, hm = ah._find_header_map_com(ws)
    if header_row <= 0 or not hm:
        raise RuntimeError(f"找不到表头: {ws.Name}")
    return ws, header_row, hm


def _first_row_with_hyperlink(ws, header_row, file_col):
    last_row = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    for r in range(header_row + 1, last_row + 1):
        cell = ws.Cells(r, file_col)
        try:
            if cell.Hyperlinks.Count:
                return r
        except Exception:
            pass
        try:
            formula = str(cell.Formula or "")
            if "HYPERLINK(" in formula.upper():
                return r
        except Exception:
            pass
    return None


def main():
    pythoncom.CoInitialize()
    base_dir = os.path.dirname(os.path.abspath(__file__))
    src_xls = os.path.join(base_dir, "2026工区收文目录.xls")
    tmp_xls = os.path.join(base_dir, f"2026工区收文目录.verify.{int(time.time())}.tmp.xls")
    test_dir_rel = os.path.join("1-上级文", "26")
    test_dir_abs = os.path.join(base_dir, test_dir_rel)
    os.makedirs(test_dir_abs, exist_ok=True)

    test_files = [
        os.path.join(test_dir_abs, "（上级文〔2026〕998号）自动化测试文件A.doc"),
        os.path.join(test_dir_abs, "（上级文〔2026〕999号）自动化测试文件B.doc"),
    ]

    shutil.copy2(src_xls, tmp_xls)

    app = None
    wb = None
    app2 = None
    wb2 = None
    try:
        app = _open_excel()
        wb = app.Workbooks.Open(tmp_xls, UpdateLinks=0, ReadOnly=False)
        ws, header_row, hm = _sheet_and_headers(wb, "上级文")

        file_col = hm["文件名"]
        date_col = hm.get("收文日期")
        widths_before = [ws.Columns(c).ColumnWidth for c in range(1, 9)]

        wb.Close(SaveChanges=True)
        wb = None
        app.Quit()
        app = None

        for p in test_files:
            with open(p, "wb") as f:
                f.write(b"test")
            ah._update_workbook(tmp_xls, p)

        app2 = _open_excel()
        wb2 = app2.Workbooks.Open(tmp_xls, UpdateLinks=0, ReadOnly=False)
        ws2, header_row2, hm2 = _sheet_and_headers(wb2, "上级文")

        today_dot = datetime.now().strftime("%Y.%m.%d")
        today_slash = datetime.now().strftime("%Y/%m/%d")

        for p in test_files:
            target_row = ah._find_existing_row_com(
                ws2,
                header_row2,
                hm2,
                ah._extract_doc_no(os.path.basename(p)),
                os.path.relpath(p, ah.WATCH_DIR),
                os.path.basename(p),
            )
            if target_row is None:
                raise RuntimeError(f"未找到新写入的记录行: {os.path.basename(p)}")

            cell2 = ws2.Cells(target_row, hm2["文件名"])
            value2 = str(cell2.Value or "").strip()
            if value2 != os.path.basename(p):
                raise RuntimeError(f"文件名写入异常: {value2}")

            try:
                link_count = int(cell2.Hyperlinks.Count)
            except Exception:
                link_count = 0
            if link_count < 1:
                raise RuntimeError(f"未生成超链接: {os.path.basename(p)}")

            link_addr = str(cell2.Hyperlinks(1).Address or "")
            if not link_addr:
                raise RuntimeError(f"超链接地址为空: {os.path.basename(p)}")

            if date_col:
                date_value = str(ws2.Cells(target_row, date_col).Value or "").strip()
                if date_value not in (today_dot, today_slash):
                    raise RuntimeError(f"日期写入异常: {date_value} (期望 {today_dot} 或 {today_slash})")

        widths_after = [ws2.Columns(c).ColumnWidth for c in range(1, 9)]
        if widths_before != widths_after:
            raise RuntimeError("列宽发生变化（疑似触发表格蜷缩）")

        wb2.Close(SaveChanges=True)
        wb2 = None
        app2.Quit()
        app2 = None

        print("检查通过：超链接未覆盖、格式未变、文件名正确、日期正确。")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            if wb2 is not None:
                wb2.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if app2 is not None:
                app2.Quit()
        except Exception:
            pass
        try:
            for p in test_files:
                if os.path.exists(p):
                    os.remove(p)
        except Exception:
            pass
        try:
            if os.path.exists(tmp_xls):
                os.remove(tmp_xls)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    main()
