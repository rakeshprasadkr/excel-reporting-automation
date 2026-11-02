import os
import time
import pythoncom
from datetime import datetime
import win32com.client

folder_path =(
    r"C:\Users\Rakesh K R\I Fund Cities"
    r"\Originations - Daily_Sales_Reporting"
    r"\Borrower_Contacts_Management_Category"
    r"\Daily Reports Backup"
)
file_prefix = "Management_report_Borrower_Contacts_Details_V1.04_"
file_extension = ".xlsx"

def log(message):
    print(message, flush=True)

today = datetime.now().strftime("%Y%m%d")
file_name: str = (file_prefix + today + file_extension)
file_path = os.path.join(folder_path,file_name)

log("=" * 60)
log("Borrower_Contacts_file REFRESH STARTED")
log(f"File: {file_name}")

if not os.path.isfile(file_path):
    raise FileNotFoundError(f"Workbook does not exist")

log("Workbook found")

excel = None
workbook = None
com_initialized = False

try:
    pythoncom.CoInitialize()
    com_initialized = True

    excel = win32com.client.DispatchEx("Excel.Application")

    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    excel.AskToUpdateLinks = False

    log("Excel Initialized")
    
    open_start = time.time()
    log("Opening Workbook")
    
    workbook = excel.Workbooks.Open(
        file_path,
        UpdateLinks = 0,
        ReadOnly = False,
        AddToMru = False
    )
    
    if workbook.ReadOnly:
        raise RuntimeError("It is Read-Only, Check if it is locked")
    
    log("\nConnections to refresh:")
    
    for i in range(1,workbook.Connections.Count +1):
        connection = workbook.Connections.Item(i)
        log(f"{i}.{connection.Name}")
    log(f"Total Connections: {workbook.Connections.Count}")    
    
    workbook.RefreshAll()
    
    log(f"Refreshing All Connections")
    excel.CalculateUntilAsyncQueriesDone()
    excel.Calculate()
    log("\nRecalculating workbook")
    
    while excel.CalculationState != 0:
        pythoncom.PumpWaitingMessages()
        time.sleep(2)
    
    workbook.Save()
    log("\nSaving Refreshed Workbook")
    log(f"Total execution time: "
            f"{time.time() - open_start:.2f} seconds.")
    
except Exception as e:
    log(f"Refresh Error: {repr(e)}")
    raise

finally:
    if workbook is not None:
        try:
            workbook.Close(SaveChanges=False)
            log("Workbook Saved and Closed")
        except Exception as e:
            log(f"Workbook close error: {repr(e)}")

    if excel is not None:
        try:
            excel.Quit()
            log("Excel application closed")
        except Exception as e:
            log(f"Excel close error: {repr(e)}")

    if com_initialized:
        pythoncom.CoUninitialize()

    log("Borrower_Contacts_file REFRESH FINISHED.")
    log("=" * 60)