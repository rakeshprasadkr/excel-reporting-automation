import time
import os
import datetime
import win32com.client
import pywintypes
import pythoncom


def wait_until_excel_ready(xlapp, timeout=120):

    start = time.time()

    while True:

        try:
            if xlapp.Ready:
                time.sleep(0.3)
                return

        except Exception:
            pass

        if time.time() - start > timeout:
            raise TimeoutError("Excel never became ready")

        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass

        time.sleep(0.5)

def refresh_and_save(
    xlapp,
    wb_path,
    max_save_retries=8,
    retry_delay=0.5
):

    print("\nOpening:", wb_path)

    wb = xlapp.Workbooks.Open(wb_path)

    wait_until_excel_ready(xlapp)

    try:

        try:

            for ws in wb.Worksheets:

                for qt in ws.QueryTables:

                    try:
                        qt.BackgroundQuery = False
                    except Exception:
                        pass
                for lo in ws.ListObjects:

                    try:
                        if hasattr(lo, "QueryTable"):
                            lo.QueryTable.BackgroundQuery = False
                    except Exception:
                        pass

        except Exception:
            pass

        print("Refreshing:", os.path.basename(wb_path))

        wb.RefreshAll()
        try:
            xlapp.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        wait_seconds = 0
        max_wait = 60

        while wait_seconds < max_wait:

            app_ready = True

            try:
                app_ready = bool(xlapp.Ready)
            except Exception:
                app_ready = True

            if app_ready:

                time.sleep(0.3)
                break

            time.sleep(0.5)
            wait_seconds += 0.5

        attempt = 0
        current_retry_delay = retry_delay

        while True:

            try:

                wb.Save()

                print("Saved:", wb_path)

                break

            except pywintypes.com_error as e:

                attempt += 1

                if attempt > max_save_retries:

                    print(
                        f"Save failed after {attempt} attempts."
                    )

                    raise

                print(
                    f"Save busy, retrying "
                    f"{attempt}/{max_save_retries} "
                    f"after {current_retry_delay} seconds... "
                    f"({e})"
                )

                try:
                    pythoncom.PumpWaitingMessages()
                except Exception:
                    pass

                time.sleep(current_retry_delay)

                current_retry_delay = min(
                    current_retry_delay * 1.5,
                    5.0
                )

    finally:
        try:
            wb.Close(SaveChanges=False)

        except Exception:
            try:
                wb.Close()
            except Exception:
                pass

if __name__ == "__main__":

    xlapp = win32com.client.DispatchEx(
        "Excel.Application"
    )

    xlapp.Visible = False
    xlapp.DisplayAlerts = False
    xlapp.AutomationSecurity = 3

    today = datetime.datetime.today().strftime("%Y%m%d")

    land_base = (
        r"C:\Users\Rakesh K R\I Fund Cities"
        r"\Land Banking - Documents\Reporting"
    )

    land_files = [

        os.path.join(
            land_base,
            "Pipeline Report",
            f"LMX_Land_Bank_Pipeline_{today}.xlsx"
        )

    ]

    sales_base = (
        r"C:\Users\Rakesh K R\I Fund Cities"
        r"\Originations - Documents\Daily_Sales_Reporting"
    )

    sales_files = [
        os.path.join(
            sales_base,
            "Daily_PnL_Pipeline",
            "Daily Reports Backup",
            f"New_Daily_PnL_Pipeline_SQL_Refresh_V1.04_{today}.xlsx"
        ),

        os.path.join(
            sales_base,
            "New_Sales_Discovery_V2",
            "Daily Reports Backup",
            f"New_Sales_Discovery_SQL_Refresh_V2.07_{today}.xlsx"
        ),

        os.path.join(
            sales_base,
            "Borrower_Contacts_Management_Category",
            "Daily Reports Backup",
            f"Management_report_Borrower_Contacts_Details_V1.04_{today}.xlsx"
        )

    ]

    marketing_base = (
        r"C:\Users\Rakesh K R\I Fund Cities"
        r"\Servicing - Documents\Dwaraka"
        r"\Reports Rebuilding\Priority\Priority-01"
        r"\Adhoc requests\Tina\Marketing Reports"
        r"\Daily_Reporting"
    )

    marketing_files = [

        os.path.join(
            marketing_base,
            "Daily_Clicks",
            "Current_Day_Reports",
            f"Daily_Clicks_V1.01_{today}.xlsx"
        ),

        os.path.join(
            marketing_base,
            "Email_Counts_1",
            "Current_Day_Reports",
            f"Email_counts_V1.01_{today}.xlsx"
        ),

        os.path.join(
            marketing_base,
            "Marketing_Leads",
            "Current_Day_Reports",
            f"Marketing_Leads_V1.01_{today}.xlsx"
        )

    ]

    report_groups = {

        "LAND BANKING REPORTS": land_files,

        "ORIGINATIONS / SALES REPORTS": sales_files,

        "MARKETING REPORTS": marketing_files,

    }

    all_files = []

    for group_name, group_files in report_groups.items():

        print("\n")
        print("=" * 65)
        print(group_name)
        print("=" * 65)

        for fpath in group_files:

            if not os.path.exists(fpath):

                print("File not found, skipping:", fpath)

                continue

            all_files.append(fpath)


    total_files = len(all_files)
    successful = 0
    failed = []

    print("\n")
    print("=" * 65)
    print("MASTER REPORT REFRESH STARTED")
    print("Total workbooks found:", total_files)
    print("=" * 65)

    try:

        for index, fpath in enumerate(all_files, start=1):

            print("\n")
            print("-" * 65)

            print(
                f"Processing workbook {index}/{total_files}"
            )

            print("-" * 65)

            try:

                refresh_and_save(xlapp, fpath)

                successful += 1

            except Exception as e:

                print("ERROR processing:", fpath)
                print("Error:", e)

                failed.append({
                    "file": fpath,
                    "error": str(e)
                })

                # Continue with next workbook

                continue

    finally:
        try:
            xlapp.Quit()
        except Exception:
            pass

        del xlapp

    print("\n")
    print("=" * 65)
    print("MASTER REPORT REFRESH COMPLETED")
    print("=" * 65)

    print("Total workbooks found :", total_files)
    print("Successfully refreshed:", successful)
    print("Failed workbooks      :", len(failed))

    if failed:

        print("\nFAILED REPORTS:")

        for item in failed:

            print("\nFile:", item["file"])
            print("Error:", item["error"])

    print("\nLand Banking + Sales + Marketing refresh completed.")
