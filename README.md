# EDGAR DCF Excel Integration (`edgar-dcf-excel`)

Automated 5-year historical financial statement extraction from **SEC EDGAR (10-K / 10-Q XBRL)** directly into a structured **Excel DCF Valuation Model** (`DCF_Model_Template.xlsm`).

---

## Features

- **Live Excel Ticker Trigger (`C7`)**: Type any public US stock ticker (e.g., `GOOGL`, `UBER`, `AVGO`, `NVDA`) into cell `C7` of the `Statements` tab and press **Enter** to automatically fetch and populate 5 years of historical financials.
- **Exact Balance Sheet Reconciliation (`$0.00` Discrepancy)**:
  - Parses raw XBRL 10-K Balance Sheets, filters out dimensional/breakdown duplicates, and reconciles every section against reported GAAP totals (`AssetsCurrent`, `Assets`, `LiabilitiesCurrent`, `Liabilities`, and `LiabilitiesAndStockholdersEquity`) so `Total assets` (`Row 78`) always equals `Total liabilities & equity` (`Row 100`).
- **Complete Income Statement & EBITDA Reconciliation**:
  - Extracts 5-year historical `Revenue`, `COGS`, `Operating & Support`, `S&M`, `R&D`, `G&A`, `D&A` (`Row 23`), `Operating Income` (`Row 25`), `EBITDA` (`Row 37`), `Stock-Based Compensation` (`Row 38`), and multi-class `Diluted Shares` (`Row 41`).
  - Automatically falls back to Cash Flow Statement `D&A` when not reported as a separate line on the GAAP Income Statement while preserving exact GAAP `Operating Income` and `EBITDA`.
- **PP&E & D&A Schedule Population (`Rows 121–133`)**:
  - Populates historical `Beginning of period` (`Row 123`), `Plus: CapEX` (`Row 124`), `Less: Depreciation` (`Row 125`), `End of Period` (`Row 126` linked to Balance Sheet `PP&E` in `Row 73`), and `Total D&A` (`Row 133` linked to Income Statement `D&A` in `Row 23`) across `2022A–2025A` (`Cols F–I`) without overwriting projected formulas (`Cols J–N`).
- **Latest Quarterly Overview Metrics**:
  - Automatically pulls latest 10-Q / 10-K `Cash & Short-Term Investments` (`I6`), `Total Debt` (`I7`), and `Diluted Shares` (`F7`), while preserving live Excel 365 `STOCKHISTORY` share price formulas (`F6`) and two-variable sensitivity Data Tables on the `Valuation` sheet.

---

## Repository Structure

| File | Description |
| :--- | :--- |
| `DCF_Model_Template.xlsm` | Main macro-enabled Excel DCF model with embedded `EDGAR_Fetcher` VBA module and `Statements` worksheet change trigger. |
| `DCF_Model_Template.xlsx` | Base non-macro Excel template. |
| `fetch_edgar.py` | Python SEC EDGAR XBRL fetcher (`edgartools`) that outputs standardized & reconciled 5-year historical data to `edgar_data.csv`. |
| `EDGAR_Fetcher.bas` | VBA module (`FetchEdgarData`, `LoadEdgarCSVIntoSheet`, `AutoHideEmptyRows`, `GetRowForKey`). |
| `Statements_Worksheet_Code.txt` | VBA `Worksheet_Change` handler for the `Statements` tab (`C7` ticker cell). |
| `populate_dcf.py` | Standalone Python utility to populate `DCF_Model_Template.xlsm` from `edgar_data.csv` via `openpyxl`. |
| `create_xlsm_with_macros.py` | Helper script to inject `EDGAR_Fetcher.bas` and `Statements_Worksheet_Code.txt` into the Excel template via COM (`pywin32`). |
| `pyproject.toml` / `uv.lock` | Python project dependencies managed with `uv` (`edgartools`, `openpyxl`, `pandas`, `pywin32`). |

---

## Setup & Usage

### 1. Install Dependencies
```powershell
uv sync
```

### 2. Use Directly in Excel
1. Open `DCF_Model_Template.xlsm` in Microsoft Excel and enable macros.
2. Enter your SEC EDGAR contact email in cell `C8` (e.g., `name@example.com`).
3. Enter any ticker symbol in cell `C7` (e.g., `GOOGL`, `UBER`, `AVGO`, `NVDA`) and press **Enter**.

### 3. Run via Command Line
```powershell
# Fetch 5 years of SEC EDGAR data into edgar_data.csv
.\.venv\Scripts\python.exe fetch_edgar.py --ticker GOOGL --email your_email@example.com

# Populate a standalone workbook from edgar_data.csv
.\.venv\Scripts\python.exe populate_dcf.py GOOGL
```
