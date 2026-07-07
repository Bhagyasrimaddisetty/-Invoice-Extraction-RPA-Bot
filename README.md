# Invoice Processing Automation Bot Using UiPath

## Project Overview

This project automates invoice processing using UiPath. The bot reads invoices from a folder, extracts invoice details, validates mandatory fields, stores structured data in Excel, and generates an automation summary. It simulates a real accounts payable automation process commonly used in enterprises.

## Features

- **Folder Automation** — automatically scans `sample_invoices/` for all PDF invoices
- **PDF Reading** — extracts full text from each invoice using PDF activities
- **Data Extraction** — pulls Invoice Number, Invoice Date, Vendor, and Amount via regex
- **Validation** — flags invoices with missing mandatory fields instead of failing silently
- **Excel Reporting** — writes a structured, color-coded report to `output/Invoice_Report.xlsx`
- **Exception Handling** — Try/Catch around each invoice so one bad file doesn't stop the run
- **Logging** — a log message after every major step for traceability

## Tools

- UiPath Studio (Community Edition)
- Microsoft Excel
- PDF Activities
- Windows

## 🎥 Project Demo

▶️ **Watch the automation in action:**

[Invoice Processing Automation Bot Demo](https://drive.google.com/file/d/1wpUZFII6JAg-BklBPdYQJdp9kdPQxaLE/view?usp=sharing)


## Workflow Architecture

```
Start
  |
Read Invoice Folder
  |
For Each Invoice
  |
Extract Invoice Data
  |
Validate Fields
  |
Valid? --- No --> Store Error
  |
 Yes
  |
Store Data
  |
Continue Loop
  |
Write Excel Report
  |
Generate Summary
  |
End
```

## Folder Structure

```
Invoice-Extraction-RPA-Bot
│
├── sample_invoices/                    10 sample invoice PDFs (8 valid, 2 with intentional errors)
├── output/
│      Invoice_Report.xlsx              Generated Excel report
├── screenshots/                        Add UiPath workflow + Excel output screenshots here
├── Main.xaml                           UiPath workflow (open in UiPath Studio)
├── project.json                        UiPath project descriptor
├── process_invoices_reference.py       Python reference implementation (see note below)
└── README.md
```

## How to Run in UiPath Studio

1. Install **UiPath Studio Community Edition** (Windows).
2. Open `project.json` from UiPath Studio ("Open" → select this folder).
3. In the **Manage Packages** panel, restore the dependencies listed in `project.json`:
   `UiPath.Excel.Activities`, `UiPath.PDF.Activities`, `UiPath.System.Activities`, and optionally `UiPath.Mail.Activities`.
4. Open `Main.xaml`. It contains the full workflow described above, built from standard activities:
   `Assign`, `For Each`, `Read PDF Text`, `Add Data Row`, `If`, `Write Range (Workbook)`, `Log Message`, and `Try Catch`.
5. Press **Run** (F5). The bot will:
   - Read every PDF in `sample_invoices/`
   - Extract fields and validate them
   - Write `output/Invoice_Report.xlsx`
   - Print an automation summary to the Output panel

> **Note on `Main.xaml`:** this file was hand-authored as valid UiPath XAML matching the workflow architecture above. Since UiPath Studio itself only runs on Windows, this file has **not** been executed inside UiPath Studio in this environment — open it in Studio and click through each activity once to confirm the packages/arguments resolve on your machine before recording the demo. If any activity shows a red exclamation mark, it almost always means a missing package from step 3.

## Reference Implementation (`process_invoices_reference.py`)

Because UiPath Studio doesn't run in this environment, `process_invoices_reference.py` is a Python script that mirrors the exact same steps (read PDF → regex extraction → validation → Excel write → summary) so the logic could be built and verified end-to-end. It was run against the 10 sample invoices and produced:

```
Automation Summary
-------------------
Total Invoices : 10
Processed      : 10
Valid          : 8
Errors         : 2
```

This matches `output/Invoice_Report.xlsx`, included in this project, so you can see exactly what the UiPath bot should produce once run in Studio. Use this script as a sanity check if you change the regex patterns or add new sample invoices — you don't need UiPath installed to test extraction logic changes.

## Sample Invoices

10 invoices are included in `sample_invoices/`, generated to match the spec:

| File | Invoice Number | Status |
|---|---|---|
| INV1001.pdf | INV-1001 | Valid |
| INV1002.pdf | INV-1002 | Valid |
| INV1003.pdf | INV-1003 | Valid |
| INV1004.pdf | INV-1004 | Valid |
| INV1005.pdf | INV-1005 | Valid |
| INV1006.pdf | INV-1006 | Valid |
| INV1007.pdf | INV1007  | **Missing Date** |
| INV1008.pdf | INV-1008 | Valid |
| INV1009.pdf | INV1009  | **Missing Amount** |
| INV1010.pdf | INV-1010 | Valid |

## Workflow Screenshot

![UiPath Workflow](screenshots/Screenshot%202026-07-07%20180911.png)

## Output Screenshot

![Invoice Report](screenshots/Screenshot%202026-07-07%20182056.png)

## Optional: Email Notification

`Main.xaml` includes a commented-out **Send Outlook Mail Message** activity. Uncomment it and set your Outlook account, recipient, subject, and body to email the report automatically after each run.

## Exception Handling

Each invoice is processed inside a `Try Catch`. If reading or extracting a single invoice fails, the bot logs the error and continues to the next invoice rather than stopping the entire run.

## GitHub Repository Structure

```
Invoice-Extraction-RPA-Bot
├── Main.xaml
├── project.json
├── sample_invoices/
├── output/
├── screenshots/
├── process_invoices_reference.py
└── README.md
```
