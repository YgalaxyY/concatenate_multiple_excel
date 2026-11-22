# Excel Merge Core

A Python desktop prototype that consolidates rows from multiple XLSX files into one workbook. It scans the current folder, skips the output workbook and merges the inputs with pandas.

[Source code](COOPERATE.py) · [Русский](README.ru.md)

## Example

Two headerless input sheets contain the rows `A, 10`, `B, 20` and `C, 30`. The output sheet contains all three rows. Headers are treated as ordinary rows. The prototype uses pandas, openpyxl, PySimpleGUI and xlwings.

## Try it with sample data

1. Create a separate test folder with two sample XLSX files and a blank `report.xlsx` containing a sheet named `Sheet1`. Copy [COOPERATE.py](COOPERATE.py) there.
2. Activate a virtual environment and install [requirements.txt](requirements.txt). Package versions are unpinned and setup may depend on the PySimpleGUI release.
3. From the test folder, run `python COOPERATE.py`. In the window, choose **Create a new report about clients**. After the GUI closes, the script attempts to open the report with xlwings, which needs Microsoft Excel in this workflow.

## Current limits

- The current code expects a pre-existing `report.xlsx`. Use a blank output and input copies each time. Without the prepared output, the script can point at an input workbook.
- The prototype merges row values. It does not align mismatched schemas, remove duplicates or preserve workbook styling.
- It looks for inputs in the current working directory and does not clear older rows left in the output worksheet.
- Compatibility with newly installed dependencies and large workbooks has not been measured.
