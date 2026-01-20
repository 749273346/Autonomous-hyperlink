import os
import time
import threading
import win32com.client
import pythoncom
from auto_hyperlink import _update_workbook, _wait_for_file_unlock

# Mock environment
TEST_XLS = "2026工区收文目录.xls"
TEST_FILE = "上级文/2026/测试通知.doc"

def create_dummy_xls():
    pythoncom.CoInitialize()
    if os.path.exists(TEST_XLS):
        try: os.remove(TEST_XLS)
        except: pass
    
    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    wb = app.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "上级文"
    
    # Headers
    headers = ["序号", "收文日期", "文号", "文件名", "自编号", "传阅方式", "存盒位置", "备注"]
    for i, h in enumerate(headers):
        ws.Cells(1, i+1).Value = h # Header at row 1
        
    wb.SaveAs(os.path.abspath(TEST_XLS), FileFormat=56) # xls
    wb.Close()
    app.Quit()
    
    # Verify creation
    app = win32com.client.Dispatch("Excel.Application")
    wb = app.Workbooks.Open(os.path.abspath(TEST_XLS))
    ws = wb.Worksheets(1)
    val = ws.Cells(1, 1).Value
    print(f"Verified header at A1: {val}")
    wb.Close(False)
    app.Quit()

def check_seq_is_1():
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    wb = app.Workbooks.Open(os.path.abspath(TEST_XLS))
    ws = wb.Worksheets("上级文")
    
    # Data should be at row 2
    seq = ws.Cells(2, 1).Value
    print(f"Sequence at row 2 is: {seq}")
    
    wb.Close(False)
    app.Quit()
    
    if int(seq) == 1:
        print("PASS: Sequence is 1")
    else:
        print("FAIL: Sequence is not 1")

def test_lock_mechanism():
    print("\nTesting lock mechanism (COM lock)...")
    # Lock the file using Excel
    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    wb = app.Workbooks.Open(os.path.abspath(TEST_XLS))
    print("File locked by Excel (Main thread).")
    
    def run_update():
        print("Thread: Attempting update...")
        # Create a dummy file to trigger update logic
        if not os.path.exists(os.path.dirname(TEST_FILE)):
            os.makedirs(os.path.dirname(TEST_FILE))
        with open(TEST_FILE, "w") as f: f.write("test")
            
        _update_workbook(os.path.abspath(TEST_XLS), os.path.abspath(TEST_FILE))
        print("Thread: Update finished.")

    t = threading.Thread(target=run_update)
    t.start()
    
    time.sleep(4) # Wait longer than timeout (5s) to see if it gives up, or wait shorter to see if it retries?
    # Wait 2 seconds, then close. _update_workbook timeout is 5s.
    # It should retry for 2 seconds then succeed.
    
    print("Main: Closing Excel...")
    wb.Close(SaveChanges=True)
    app.Quit()
    del wb
    del app
    print("Main: Excel closed.")
    
    t.join()
    print("Lock test finished.")

if __name__ == "__main__":
    try:
        print("Creating dummy xls...")
        create_dummy_xls()
        
        print("Updating workbook (should start at seq 1)...")
        if not os.path.exists(os.path.dirname(TEST_FILE)):
            os.makedirs(os.path.dirname(TEST_FILE))
        with open(TEST_FILE, "w") as f: f.write("test")
            
        _update_workbook(os.path.abspath(TEST_XLS), os.path.abspath(TEST_FILE))
        
        check_seq_is_1()
        
        test_lock_mechanism()
        
    except Exception as e:
        print(f"Test failed: {e}")
    finally:
        # Cleanup
        if os.path.exists(TEST_XLS):
            try: os.remove(TEST_XLS)
            except: pass
