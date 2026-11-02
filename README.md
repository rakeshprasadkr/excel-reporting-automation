# Excel Reporting Automation Using Python

     Project Overview

A Python-based automation solution developed to streamline recurring Excel reporting workflows by automatically refreshing workbook data connections, executing queries, recalculating reports, and saving updated workbooks.

The solution uses Python and Windows COM automation to interact with Microsoft Excel Desktop, reducing repetitive manual reporting activities.

     Business Problem

Daily reporting processes often require users to:
- Open multiple Excel workbooks manually.
- Refresh SQL and other data connections.
- Wait for queries to complete.
- Recalculate workbook formulas.
- Save updated reports.

These repetitive tasks consume time and introduce opportunities for manual errors.

This project automates the workbook refresh and save process.

     Key Features

-     Automated File Identification:     Identifies daily reporting workbooks using date-based filenames.
-     File Validation:     Checks whether the expected workbook exists before execution.
-     Excel Automation:     Opens and controls Microsoft Excel using Python COM automation.
-     Data Connection Refresh:     Refreshes all configured workbook connections.
-     Asynchronous Query Handling:     Waits for background queries to finish before proceeding.
-     Workbook Recalculation:     Recalculates formulas after data refresh.
-     Automated Saving:     Saves the refreshed workbook.
-     Error Handling:     Captures exceptions and reports execution failures.
-     Resource Management:     Closes workbook and Excel application and releases COM resources.

     Technology Stack

| Technology | Purpose |

| Python | Automation logic |
| pywin32 | Windows COM integration |
| Microsoft Excel | Reporting and data connections |
| SQL Server-connected workbooks | Data source integration |
| Windows | Execution environment |

     Workflow

1. Generate the expected daily filename.
2. Locate and validate the Excel workbook.
3. Initialize the Excel application through COM.
4. Open the workbook in editable mode.
5. Identify configured workbook connections.
6. Trigger a refresh of all connections.
7. Wait for asynchronous queries and calculations.
8. Save the refreshed workbook.
9. Close Excel and release resources.

     Project Structure

excel-reporting-automation/
│
├── README.md
├── requirements.txt
├── .gitignore
│
└── scripts/
    └── daily_excel_refresh.py

     Prerequisites

- Windows operating system.
- Microsoft Excel Desktop installed.
- Python 3.x.
- Access to the relevant workbook and its configured data sources.

     Installation

Clone the repository:


Navigate to the project directory:

Install dependencies:


     Execution

Run the automation script:

     Business Impact

The automation helps:
- Reduce repetitive manual reporting activities.
- Improve consistency in daily workbook refreshes.
- Minimize manual intervention in recurring reporting workflows.
- Streamline existing Excel-based reporting processes.

     Limitations

- Requires Windows and Microsoft Excel Desktop.
- Uses COM automation through `pywin32`.
- Requires appropriate access to workbook data sources.
- Not designed to run directly in standard Linux environments or Google Colab.

     Future Enhancements

- Configurable workbook paths and filenames.
- Detailed execution logs with timestamps.
- Automated failure notifications.
- Scheduled execution using Windows Task Scheduler.
- Workbook validation after refresh.
- Support for multiple reporting workbooks.

     Disclaimer

This repository contains demonstration code for portfolio purposes. No proprietary company workbooks, confidential business data, credentials, or client-specific information are included.
