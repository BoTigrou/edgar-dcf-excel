import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_template_v2():
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    # ----------------------------------------------------
    # Styles and Palettes (Rich & Sleek Aesthetics)
    # ----------------------------------------------------
    font_family = "Segoe UI"
    
    # Fonts
    font_title = Font(name=font_family, size=16, bold=True, color="000000")
    font_subtitle = Font(name=font_family, size=9, italic=True, color="595959")
    font_sec_header = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    font_sub_header = Font(name=font_family, size=10, bold=True, color="000000")
    
    # Color-coded Fonts (AS PER USER REQUEST)
    font_regular = Font(name=font_family, size=10, color="000000") # Black for labels/local calculations
    font_bold = Font(name=font_family, size=10, bold=True, color="000000") # Black bold
    font_italic = Font(name=font_family, size=10, italic=True, color="000000") # Black italic
    font_gray = Font(name=font_family, size=9, color="595959") # Gray for assumptions
    font_gray_italic = Font(name=font_family, size=9, italic=True, color="595959") # Gray italic
    
    # RGB Blue (0 - 0 - 255) for hardcoded inputs
    font_blue_input = Font(name=font_family, size=10, color="0000FF") 
    font_blue_bold = Font(name=font_family, size=10, bold=True, color="0000FF")
    
    # RGB Green (0 - 128 - 0) for formulas linking another sheet
    font_green_link = Font(name=font_family, size=10, color="008000") 
    font_green_bold = Font(name=font_family, size=10, bold=True, color="008000")
    font_green_target_price = Font(name=font_family, size=11, bold=True, color="008000")
    
    # Fills
    fill_sec_header = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")
    fill_sub_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_navy_band = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
    
    # Alignments
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    
    # Borders
    thin_line = Side(border_style="thin", color="D9D9D9")
    thick_line = Side(border_style="medium", color="1A1A1A")
    double_line = Side(border_style="double", color="1A1A1A")
    
    border_all_thin = Border(left=thin_line, right=thin_line, top=thin_line, bottom=thin_line)
    border_top_thin = Border(top=thin_line)
    border_bottom_double = Border(bottom=double_line)
    border_total = Border(top=thin_line, bottom=double_line)
    
    cols = ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
    years = [2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]
    
    # Helper to style section headers
    def format_sec_row(ws_obj, row_idx):
        ws_obj.row_dimensions[row_idx].height = 24
        for c in range(2, 15):
            col_letter = get_column_letter(c)
            cell = ws_obj[f"{col_letter}{row_idx}"]
            cell.fill = fill_sec_header
            cell.font = font_sec_header

    # =========================================================================
    # TAB 1: Statements
    # =========================================================================
    ws_stmt = wb.create_sheet(title="Statements")
    ws_stmt.views.sheetView[0].showGridLines = True
    
    ws_stmt.row_dimensions[2].height = 25
    ws_stmt.row_dimensions[3].height = 18
    
    ws_stmt["B2"] = "Statements"
    ws_stmt["B2"].font = font_title
    ws_stmt["B3"] = "$ and shares in million, except per share data"
    ws_stmt["B3"].font = font_subtitle
    
    # Company Overview Block
    format_sec_row(ws_stmt, 5)
    ws_stmt.merge_cells("B5:D5")
    ws_stmt["B5"] = "Company Overview"
    ws_stmt["B5"].alignment = align_left
    
    ws_stmt["B7"] = "Company name"
    ws_stmt["B7"].font = font_regular
    ws_stmt["C7"] = "Uber Technologies, Inc."
    ws_stmt["C7"].font = font_bold
    ws_stmt["C7"].border = border_all_thin
    
    ws_stmt["E7"] = "Share price"
    ws_stmt["E7"].font = font_regular
    ws_stmt["F7"] = 72.66
    ws_stmt["F7"].font = font_blue_input # Blue (Hardcoded input)
    ws_stmt["F7"].number_format = "$#,##0.00"
    ws_stmt["F7"].border = border_all_thin
    
    ws_stmt["H7"] = "Cash"
    ws_stmt["H7"].font = font_regular
    ws_stmt["I7"] = "=I61+I62" # Black (Local formula link to cash + short-term investments)
    ws_stmt["I7"].font = font_bold
    ws_stmt["I7"].number_format = "$#,##0"
    ws_stmt["I7"].border = border_all_thin
    
    ws_stmt["K7"] = "WACC"
    ws_stmt["K7"].font = font_regular
    ws_stmt["L7"] = "=WACC!E36" # Green (Formula linking from WACC sheet!)
    ws_stmt["L7"].font = font_green_bold
    ws_stmt["L7"].number_format = "0.0%"
    ws_stmt["L7"].border = border_all_thin
    
    ws_stmt["B8"] = "Ticker"
    ws_stmt["B8"].font = font_regular
    ws_stmt["C8"] = "UBER"
    ws_stmt["C8"].font = font_bold
    ws_stmt["C8"].border = border_all_thin
    
    ws_stmt["E8"] = "Shares outstanding"
    ws_stmt["E8"].font = font_regular
    ws_stmt["F8"] = "=I38" # Black (Local formula link to latest year shares outstanding in Row 38)
    ws_stmt["F8"].font = font_bold
    ws_stmt["F8"].number_format = "#,##0"
    ws_stmt["F8"].border = border_all_thin
    
    ws_stmt["H8"] = "Debt"
    ws_stmt["H8"].font = font_regular
    ws_stmt["I8"] = "=I80+I85" # Black (Local formula link to short-term + long-term debt)
    ws_stmt["I8"].font = font_bold
    ws_stmt["I8"].number_format = "$#,##0"
    ws_stmt["I8"].border = border_all_thin
    
    ws_stmt["K8"] = "Terminal Growth"
    ws_stmt["K8"].font = font_regular
    ws_stmt["L8"] = 0.02
    ws_stmt["L8"].font = font_blue_input # Blue (Hardcoded input)
    ws_stmt["L8"].number_format = "0.0%"
    ws_stmt["L8"].border = border_all_thin
    
    ws_stmt["E9"] = "Market Cap."
    ws_stmt["E9"].font = font_regular
    ws_stmt["F9"] = "=F8*F7" # Black (Local formula)
    ws_stmt["F9"].font = font_bold
    ws_stmt["F9"].number_format = "$#,##0"
    ws_stmt["F9"].border = border_all_thin
    
    ws_stmt["H9"] = "EV"
    ws_stmt["H9"].font = font_regular
    ws_stmt["I9"] = "=F9-I7+I8" # Black (Local formula)
    ws_stmt["I9"].font = font_bold
    ws_stmt["I9"].number_format = "$#,##0"
    ws_stmt["I9"].border = border_all_thin
    
    ws_stmt["B4"] = "SEC Email"
    ws_stmt["B4"].font = font_regular
    ws_stmt["C4"] = "julie@example.com"
    ws_stmt["C4"].font = font_blue_input
    ws_stmt["C4"].border = border_all_thin
    
    # ----------------------------------------------------
    # Income Statement Section (Row 11)
    # ----------------------------------------------------
    format_sec_row(ws_stmt, 11)
    ws_stmt.merge_cells("B11:D11")
    ws_stmt["B11"] = "Income Statement"
    ws_stmt["B11"].alignment = align_left
    
    for col, yr in zip(cols, years):
        cell_ref = f"{col}11"
        suffix = "A" if yr <= 2025 else "E"
        ws_stmt[cell_ref] = f"{yr}{suffix}"
        ws_stmt[cell_ref].alignment = align_right
        
    # Revenue (Row 13)
    ws_stmt["B13"] = "Revenue"
    ws_stmt["B13"].font = font_regular
    revs = [17455, 31877, 37281, 43793, 52017]
    for col, val in zip(cols[:5], revs):
        ws_stmt[f"{col}13"] = val
        ws_stmt[f"{col}13"].font = font_blue_input
        ws_stmt[f"{col}13"].number_format = "#,##0"
        ws_stmt[f"{col}13"].alignment = align_right
    # Projected Revenue
    proj_rev_growth = [0.15, 0.12, 0.10, 0.08, 0.06]
    for i in range(5, 10):
        col = cols[i]
        prev_col = cols[i-1]
        ws_stmt[f"{col}13"] = f"={prev_col}13*(1+{col}47)" # Black (Local formula)
        ws_stmt[f"{col}13"].font = font_regular
        ws_stmt[f"{col}13"].number_format = "#,##0"
        ws_stmt[f"{col}13"].alignment = align_right
        
    # COGS (Row 16)
    ws_stmt["B16"] = "Cost of revenue"
    ws_stmt["B16"].font = font_regular
    cogs = [-9351, -15655, -16900, -21000, -31338]
    for col, val in zip(cols[:5], cogs):
        ws_stmt[f"{col}16"] = val
        ws_stmt[f"{col}16"].font = font_blue_input
        ws_stmt[f"{col}16"].number_format = "#,##0"
        ws_stmt[f"{col}16"].alignment = align_right
    # Projected COGS
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}16"] = f"={col}13*{col}48" # Black (Local formula)
        ws_stmt[f"{col}16"].font = font_regular
        ws_stmt[f"{col}16"].number_format = "#,##0"
        ws_stmt[f"{col}16"].alignment = align_right
        
    # Gross Profit (Row 17)
    ws_stmt["B17"] = "Gross profit"
    ws_stmt["B17"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}17"] = f"={col}13+{col}16" # Black (Local formula)
        ws_stmt[f"{col}17"].font = font_bold
        ws_stmt[f"{col}17"].number_format = "#,##0"
        ws_stmt[f"{col}17"].alignment = align_right
        ws_stmt[f"{col}17"].border = border_top_thin
        
    # Operating Expenses (Row 19-24)
    ws_stmt["B19"] = "Operating expenses"
    ws_stmt["B19"].font = Font(name=font_family, size=10, italic=True, bold=True)
    
    op_exp_items = [
        ("  Operation & support", 20, [1877, 2413, 2689, 2732, 2854], 49),
        ("  Sales & marketing", 21, [4789, 4756, 4356, 4337, 4898], 50),
        ("  Research & development", 22, [2054, 2798, 3164, 3109, 3402], 51),
        ("  General & administrative", 23, [2316, 3136, 2682, 3639, 3241], 52),
        ("  Depreciation & amortization", 24, [902, 947, 823, 711, 719], 53)
    ]
    
    for label, r_idx, hist_vals, margin_row in op_exp_items:
        ws_stmt[f"B{r_idx}"] = label
        ws_stmt[f"B{r_idx}"].font = font_regular
        
        for col, val in zip(cols[:5], hist_vals):
            ws_stmt[f"{col}{r_idx}"] = val
            ws_stmt[f"{col}{r_idx}"].font = font_blue_input
            ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
            ws_stmt[f"{col}{r_idx}"].alignment = align_right
            
        for i in range(5, 10):
            col = cols[i]
            ws_stmt[f"{col}{r_idx}"] = f"={col}13*{col}{margin_row}" # Black (Local formula)
            ws_stmt[f"{col}{r_idx}"].font = font_regular
            ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
            ws_stmt[f"{col}{r_idx}"].alignment = align_right
            
    # Total Operating Expenses (Row 25)
    ws_stmt["B25"] = "Total operating expenses"
    ws_stmt["B25"].font = font_regular
    for col in cols:
        ws_stmt[f"{col}25"] = f"=SUM({col}20:{col}24)" # Black (Local formula)
        ws_stmt[f"{col}25"].font = font_regular
        ws_stmt[f"{col}25"].number_format = "#,##0"
        ws_stmt[f"{col}25"].alignment = align_right
        ws_stmt[f"{col}25"].border = border_top_thin
        
    # Operating Income (loss) (Row 26)
    ws_stmt["B26"] = "Operating Income (loss)"
    ws_stmt["B26"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}26"] = f"={col}17-{col}25" # Black (Local formula)
        ws_stmt[f"{col}26"].font = font_bold
        ws_stmt[f"{col}26"].number_format = "#,##0"
        ws_stmt[f"{col}26"].alignment = align_right
        ws_stmt[f"{col}26"].border = border_total
        
    # Interest expense (Row 28)
    ws_stmt["B28"] = "Interest income (expense)"
    ws_stmt["B28"].font = font_regular
    interest = [-483, -565, -633, -523, 303]
    for col, val in zip(cols[:5], interest):
        ws_stmt[f"{col}28"] = val
        ws_stmt[f"{col}28"].font = font_blue_input
        ws_stmt[f"{col}28"].number_format = "#,##0"
        ws_stmt[f"{col}28"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}28"] = 303 # Assumed constant
        ws_stmt[f"{col}28"].font = font_regular
        ws_stmt[f"{col}28"].number_format = "#,##0"
        ws_stmt[f"{col}28"].alignment = align_right
        
    # Other Income (loss) (Row 29)
    ws_stmt["B29"] = "Other Income (loss)"
    ws_stmt["B29"].font = font_regular
    other_inc = [3292, -7029, 1844, 1849, -68]
    for col, val in zip(cols[:5], other_inc):
        ws_stmt[f"{col}29"] = val
        ws_stmt[f"{col}29"].font = font_blue_input
        ws_stmt[f"{col}29"].number_format = "#,##0"
        ws_stmt[f"{col}29"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}29"] = 0 # Assumed constant
        ws_stmt[f"{col}29"].font = font_regular
        ws_stmt[f"{col}29"].number_format = "#,##0"
        ws_stmt[f"{col}29"].alignment = align_right
        
    # Income (loss) before taxes & equity (Row 31)
    ws_stmt["B31"] = "Income (loss) before taxes & equity"
    ws_stmt["B31"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}31"] = f"={col}26+{col}28+{col}29" # Black (Local formula)
        ws_stmt[f"{col}31"].font = font_bold
        ws_stmt[f"{col}31"].number_format = "#,##0"
        ws_stmt[f"{col}31"].alignment = align_right
        ws_stmt[f"{col}31"].border = border_top_thin
        
    # tax provision (Row 32)
    ws_stmt["B32"] = "provision for (benefit from) income tax"
    ws_stmt["B32"].font = font_regular
    tax = [-492, -181, 213, -5758, -4346]
    for col, val in zip(cols[:5], tax):
        ws_stmt[f"{col}32"] = val
        ws_stmt[f"{col}32"].font = font_blue_input
        ws_stmt[f"{col}32"].number_format = "#,##0"
        ws_stmt[f"{col}32"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}32"] = f"={col}31*{col}55" # Black (Local formula)
        ws_stmt[f"{col}32"].font = font_regular
        ws_stmt[f"{col}32"].number_format = "#,##0"
        ws_stmt[f"{col}32"].alignment = align_right
        
    # equity inv (Row 33)
    ws_stmt["B33"] = "Income (loss) from equity inv."
    ws_stmt["B33"].font = font_regular
    equity = [-37, 107, 48, -38, -53]
    for col, val in zip(cols[:5], equity):
        ws_stmt[f"{col}33"] = val
        ws_stmt[f"{col}33"].font = font_blue_input
        ws_stmt[f"{col}33"].number_format = "#,##0"
        ws_stmt[f"{col}33"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}33"] = 0
        ws_stmt[f"{col}33"].font = font_regular
        ws_stmt[f"{col}33"].number_format = "#,##0"
        ws_stmt[f"{col}33"].alignment = align_right
        
    # Net income (loss) incl. NCI (Row 34)
    ws_stmt["B34"] = "Net income (loss) incl. NCI"
    ws_stmt["B34"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}34"] = f"={col}31-{col}32+{col}33" # Black (Local formula)
        ws_stmt[f"{col}34"].font = font_bold
        ws_stmt[f"{col}34"].number_format = "#,##0"
        ws_stmt[f"{col}34"].alignment = align_right
        ws_stmt[f"{col}34"].border = border_top_thin
        
    # Less NCI (Row 35)
    ws_stmt["B35"] = "Less: net income (loss) attributable to NCI"
    ws_stmt["B35"].font = font_regular
    nci = [-74, 3, 269, -11, 40]
    for col, val in zip(cols[:5], nci):
        ws_stmt[f"{col}35"] = val
        ws_stmt[f"{col}35"].font = font_blue_input
        ws_stmt[f"{col}35"].number_format = "#,##0"
        ws_stmt[f"{col}35"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}35"] = 0
        ws_stmt[f"{col}35"].font = font_regular
        ws_stmt[f"{col}35"].number_format = "#,##0"
        ws_stmt[f"{col}35"].alignment = align_right
        
    # Net income (loss) (Row 36)
    ws_stmt["B36"] = "Net income (loss)"
    ws_stmt["B36"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}36"] = f"={col}34-{col}35" # Black (Local formula)
        ws_stmt[f"{col}36"].font = font_bold
        ws_stmt[f"{col}36"].number_format = "#,##0"
        ws_stmt[f"{col}36"].alignment = align_right
        ws_stmt[f"{col}36"].border = border_total
        
    # Shares (Row 38)
    ws_stmt["B38"] = "Diluted shares outstanding"
    ws_stmt["B38"].font = font_regular
    shares = [1896, 1975, 2092, 2151, 2120]
    for col, val in zip(cols[:5], shares):
        ws_stmt[f"{col}38"] = val
        ws_stmt[f"{col}38"].font = font_blue_input
        ws_stmt[f"{col}38"].number_format = "#,##0"
        ws_stmt[f"{col}38"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}38"] = f"={cols[4]}38" # Keep constant at latest
        ws_stmt[f"{col}38"].font = font_bold
        ws_stmt[f"{col}38"].number_format = "#,##0"
        ws_stmt[f"{col}38"].alignment = align_right
        
    # Diluted EPS (Row 39)
    ws_stmt["B39"] = "Diluted earnings per share"
    ws_stmt["B39"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}39"] = f"=IF({col}38<=0, \"NA\", {col}36/{col}38)" # Black (Local formula)
        ws_stmt[f"{col}39"].font = font_bold
        ws_stmt[f"{col}39"].number_format = "$#,##0.00"
        ws_stmt[f"{col}39"].alignment = align_right
        ws_stmt[f"{col}39"].border = border_total

    # Model Assumptions Header (Row 45)
    format_sec_row(ws_stmt, 45)
    ws_stmt.merge_cells("B45:D45")
    ws_stmt["B45"] = "Historical & Projected Model Assumptions"
    ws_stmt["B45"].alignment = align_left
    
    # Assumption Rows (Rows 47-56)
    ass_items = [
        ("Revenue growth rate", 47, "0.0%", "=IF({col}13_prev=0, 0, ({col}13-{col}13_prev)/ABS({col}13_prev))"),
        ("COGS % of sales", 48, "0.0%", "=-{col}16/{col}13"),
        ("Operation & support % of sales", 49, "0.0%", "={col}20/{col}13"),
        ("Sales & marketing % of sales", 50, "0.0%", "={col}21/{col}13"),
        ("R&D % of sales", 51, "0.0%", "={col}22/{col}13"),
        ("G&A % of sales", 52, "0.0%", "={col}23/{col}13"),
        ("D&A % of sales", 53, "0.0%", "={col}24/{col}13"),
        ("Operating margin", 54, "0.0%", "={col}26/{col}13"),
        ("Tax rate", 55, "0.0%", "=IF({col}31<=0, \"NA\", {col}32/{col}31)"),
        ("SBC % of sales", 56, "0.0%", "=StockBasedComp_PlaceHolder")
    ]
    
    for label, r_idx, num_fmt, formula_tpl in ass_items:
        ws_stmt[f"B{r_idx}"] = label
        ws_stmt[f"B{r_idx}"].font = font_bold if "margin" in label else font_regular
        
        # Historical Assumptions (calculated by formulas)
        for col_idx in range(5, 10): # E, F, G, H, I
            col = cols[col_idx-5]
            if label == "Revenue growth rate" and col == "E":
                ws_stmt[f"E{r_idx}"] = 0.15 # hardcoded for first historical year
                ws_stmt[f"E{r_idx}"].font = font_blue_input
            else:
                if label == "Revenue growth rate":
                    prev_col = cols[col_idx-6]
                    ws_stmt[f"{col}{r_idx}"] = f"=IF({prev_col}13=0, 0, ({col}13-{prev_col}13)/ABS({prev_col}13))"
                elif label == "SBC % of sales":
                    ws_stmt[f"{col}{r_idx}"] = f"={col}104/{col}13"
                else:
                    ws_stmt[f"{col}{r_idx}"] = formula_tpl.replace("{col}", col)
                ws_stmt[f"{col}{r_idx}"].font = font_bold if "margin" in label else font_regular
                ws_stmt[f"{col}{r_idx}"].number_format = num_fmt
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
                
        # Projected Assumptions (Blue hardcoded inputs!)
        proj_vals = {
            47: [0.15, 0.12, 0.10, 0.08, 0.06],  # Revenue growth
            48: [0.60, 0.60, 0.60, 0.60, 0.60],  # COGS % of sales
            49: [0.05, 0.05, 0.05, 0.05, 0.05],  # Op support
            50: [0.09, 0.09, 0.09, 0.08, 0.08],  # Sales & mkt
            51: [0.06, 0.06, 0.06, 0.06, 0.05],  # R&D
            52: [0.06, 0.05, 0.05, 0.05, 0.04],  # G&A
            53: [0.01, 0.01, 0.01, 0.01, 0.01],  # D&A
            55: [0.21, 0.21, 0.21, 0.21, 0.21],  # Tax rate
            56: [0.03, 0.03, 0.03, 0.02, 0.02]   # SBC
        }
        
        for col_idx in range(10, 15): # J, K, L, M, N
            col = cols[col_idx-5]
            if r_idx == 54: # Operating margin formula
                ws_stmt[f"{col}54"] = f"={col}26/{col}13"
                ws_stmt[f"{col}54"].font = font_bold
                ws_stmt[f"{col}54"].number_format = "0.0%"
                ws_stmt[f"{col}54"].alignment = align_right
            else:
                ws_stmt[f"{col}{r_idx}"] = proj_vals[r_idx][col_idx-10]
                ws_stmt[f"{col}{r_idx}"].font = font_blue_input # Blue (Hardcoded input)
                ws_stmt[f"{col}{r_idx}"].number_format = num_fmt
                ws_stmt[f"{col}{r_idx}"].alignment = align_right

    # ----------------------------------------------------
    # Balance Sheet Section (Row 58)
    # ----------------------------------------------------
    format_sec_row(ws_stmt, 58)
    ws_stmt.merge_cells("B58:D58")
    ws_stmt["B58"] = "Balance Sheet"
    ws_stmt["B58"].alignment = align_left
    
    for col, yr in zip(cols, years):
        cell_ref = f"{col}58"
        suffix = "A" if yr <= 2025 else "E"
        ws_stmt[cell_ref] = f"{yr}{suffix}"
        ws_stmt[cell_ref].alignment = align_right
        
    # Assets & Liabilities
    ws_stmt["B60"] = "Assets"
    ws_stmt["B60"].font = Font(name=font_family, size=10, bold=True, italic=True)
    
    bs_assets = [
        ("  Cash & cash equivalent", 61, [4295, 4208, 4680, 5893, 7105]),
        ("  Short-term investments", 62, [0, 103, 727, 1084, 528]),
        ("  Restricted cash & cash equivalent", 63, [631, 680, 805, 545, 631]),
        ("  Accounts receivable", 64, [2439, 2779, 3404, 3333, 3827]),
        ("  Prepaid expenses & other current assets", 65, [1454, 1479, 1681, 1390, 1902]),
        ("Total current assets", 66, None),
        ("  Restricted cash & cash equivalent", 67, [2879, 1789, 1519, 2172, 1911]),
        ("  Restricted investments", 68, [0, 1614, 4779, 7019, 8874]),
        ("  Investments", 69, [11806, 4401, 6101, 8460, 9178]),
        ("  Equity method investments", 70, [800, 870, 353, 302, 287]),
        ("  PP&E", 71, [1853, 2082, 2073, 1952, 1897]),
        ("  Operating lease right-of-use assets", 72, [1388, 1449, 1241, 1158, 1114]),
        ("  Intangible assets", 73, [2412, 1874, 1425, 1125, 1048]),
        ("  Goodwill", 74, [8420, 8263, 8151, 8066, 8931]),
        ("  Other assets", 75, [397, 518, 1760, 8745, 14569]),
        ("Total assets", 76, None)
    ]
    
    for label, r_idx, hist_vals in bs_assets:
        ws_stmt[f"B{r_idx}"] = label
        ws_stmt[f"B{r_idx}"].font = font_bold if "Total" in label else font_regular
        
        if hist_vals is not None:
            for col, val in zip(cols[:5], hist_vals):
                ws_stmt[f"{col}{r_idx}"] = val
                ws_stmt[f"{col}{r_idx}"].font = font_blue_input
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
            for i in range(5, 10):
                col = cols[i]
                ws_stmt[f"{col}{r_idx}"] = f"={cols[4]}{r_idx}" # Keep constant at latest
                ws_stmt[f"{col}{r_idx}"].font = font_regular
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
        else:
            if label == "Total current assets":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}61:{col}65)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Total assets":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}66+SUM({col}67:{col}75)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total

    # Liabilities & Equity (starts at Row 78)
    ws_stmt["B78"] = "Liabilities"
    ws_stmt["B78"].font = Font(name=font_family, size=10, bold=True, italic=True)
    
    bs_liabs = [
        ("  Accounts payable", 79, [860, 728, 790, 858, 1013]),
        ("  Short-term debt & current portion of LT debt", 80, [0, 0, 0, 0, 0]),
        ("  Operating lease liabilities, current", 81, [185, 201, 190, 175, 169]),
        ("  Accrued and other current liabilities", 82, [6537, 6232, 6397, 7689, 7751]),
        ("Total current liabilities", 83, None),
        ("  Long-term insurance reserves", 84, [2546, 3028, 4909, 7042, 9076]),
        ("  Long-term debt, net of current portion", 85, [9276, 9265, 9459, 8347, 10521]),
        ("  Operating lease liabilities, non-current", 86, [1644, 1673, 1550, 1454, 1390]),
        ("  Other long-term liabilities", 87, [935, 786, 645, 449, 412]),
        ("Total liabilities", 88, None),
        ("  Redeemable non-controlling interests", 89, [204, 430, 654, 93, 165]),
        ("Equity", 91, None),
        ("  Common stock", 92, [0, 0, 0, 0, 0]), # APIC holds equity value in templates
        ("  Accumulated other comprehensive income (loss)", 93, [12, -43, 56, -11, 40]),
        ("  Accumulated deficit", 94, [-1423, -2001, -3450, -5678, -9876]),
        ("  Non-controlling interests", 96, [122, 133, 145, 154, 158]),
        ("Total equity", 97, None),
        ("Total liabilities & equity", 98, None)
    ]
    
    for label, r_idx, hist_vals in bs_liabs:
        if label == "Equity":
            ws_stmt[f"B{r_idx}"] = label
            ws_stmt[f"B{r_idx}"].font = Font(name=font_family, size=10, bold=True, italic=True)
            continue
            
        ws_stmt[f"B{r_idx}"] = label
        ws_stmt[f"B{r_idx}"].font = font_bold if "Total" in label else font_regular
        
        if hist_vals is not None:
            for col, val in zip(cols[:5], hist_vals):
                ws_stmt[f"{col}{r_idx}"] = val
                ws_stmt[f"{col}{r_idx}"].font = font_blue_input
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
            for i in range(5, 10):
                col = cols[i]
                ws_stmt[f"{col}{r_idx}"] = f"={cols[4]}{r_idx}"
                ws_stmt[f"{col}{r_idx}"].font = font_regular
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
        else:
            if label == "Total current liabilities":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}79:{col}82)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Total liabilities":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}83+SUM({col}84:{col}87)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Total equity":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}92:{col}94)+{col}96"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Total liabilities & equity":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}88+{col}89+{col}97"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total

    # ----------------------------------------------------
    # Cash Flow Statement (Row 100)
    # ----------------------------------------------------
    format_sec_row(ws_stmt, 100)
    ws_stmt.merge_cells("B100:D100")
    ws_stmt["B100"] = "Statement of Cash Flows"
    ws_stmt["B100"].alignment = align_left
    
    for col, yr in zip(cols, years):
        cell_ref = f"{col}100"
        suffix = "A" if yr <= 2025 else "E"
        ws_stmt[cell_ref] = f"{yr}{suffix}"
        ws_stmt[cell_ref].alignment = align_right
        
    cf_items = [
        ("Net income", 102, "=E36", "Operating"),
        ("Depreciation & Amortization (D&A)", 103, "=E24", "Operating"),
        ("Stock-based compensation", 104, [213, 254, 301, 345, 412], "Operating"),
        ("Change in net working capital assets", 105, [-110, -234, -412, -311, -298], "Operating"),
        ("Change in net working capital liabilities", 106, [145, 342, 287, 432, 290], "Operating"),
        ("Other non-cash operating assets", 107, [0, 0, 0, 0, 0], "Operating"),
        ("Other non-cash operating liabilities", 108, [0, 0, 0, 0, 0], "Operating"),
        ("Cash from operating activities", 109, None, "OperatingTotal"),
        ("Capital expenditures", 110, [-234, -412, -567, -612, -528], "Investing"),
        ("Cash from investing activities", 111, None, "InvestingTotal"),
        ("Net change in long-term debt", 112, [1452, 2341, -124, -432, 856], "Financing"),
        ("Commercial paper / Revolver change", 113, [0, 0, 0, 0, 0], "Financing"),
        ("Share repurchases", 114, [0, 0, -43, -112, -234], "Financing"),
        ("Common dividends", 115, [0, 0, 0, 0, 0], "Financing"),
        ("Cash from financing activities", 116, None, "FinancingTotal"),
        ("Net change in cash during period", 117, None, "NetChangeTotal")
    ]
    
    for label, r_idx, val_data, group in cf_items:
        ws_stmt[f"B{r_idx}"] = "  " + label if "Total" not in group else label
        ws_stmt[f"B{r_idx}"].font = font_bold if "Total" in group else font_regular
        
        if val_data is not None and not isinstance(val_data, str):
            # Write historical values
            for col, val in zip(cols[:5], val_data):
                ws_stmt[f"{col}{r_idx}"] = val
                ws_stmt[f"{col}{r_idx}"].font = font_blue_input
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
            # Write projected values (formulas)
            for i in range(5, 10):
                col = cols[i]
                ws_stmt[f"{col}{r_idx}"] = f"={cols[4]}{r_idx}"
                ws_stmt[f"{col}{r_idx}"].font = font_regular
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
        elif isinstance(val_data, str):
            # Formula linking other rows/cells locally
            for col in cols:
                cell_formula = val_data.replace("E", col)
                ws_stmt[f"{col}{r_idx}"] = cell_formula
                ws_stmt[f"{col}{r_idx}"].font = font_regular
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
        else:
            # Totals
            if group == "OperatingTotal":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}102:{col}108)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif group == "InvestingTotal":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}110"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif group == "FinancingTotal":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}112:{col}115)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif group == "NetChangeTotal":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}109+{col}111+{col}116"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total

    # Set column widths
    ws_stmt.column_dimensions["A"].width = 3
    ws_stmt.column_dimensions["B"].width = 45
    ws_stmt.column_dimensions["C"].width = 25
    ws_stmt.column_dimensions["D"].width = 3
    for col_let in cols:
        ws_stmt.column_dimensions[col_let].width = 15

    # =========================================================================
    # TAB 2: WACC (USING TEMPLATE IMAGE LAYOUT)
    # =========================================================================
    ws_wacc = wb.create_sheet(title="WACC")
    ws_wacc.views.sheetView[0].showGridLines = True
    
    ws_wacc.row_dimensions[2].height = 25
    ws_wacc.row_dimensions[4].height = 24
    
    # Title
    ws_wacc["B2"] = "Weighted Average Cost of Capital"
    ws_wacc["B2"].font = font_title
    
    # Navy Band
    ws_wacc.merge_cells("B4:G4")
    ws_wacc["B4"] = "Weighted Average Cost of Capital"
    ws_wacc["B4"].font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws_wacc["B4"].fill = fill_navy_band
    ws_wacc["B4"].alignment = align_left
    
    # Cost of Equity section
    ws_wacc["B6"] = "Cost of Equity:"
    ws_wacc["B6"].font = font_bold
    
    ws_wacc["B7"] = "  Risk-Free Rate"
    ws_wacc["E7"] = 0.0421
    ws_wacc["E7"].font = font_blue_input # Blue (Hardcoded input)
    ws_wacc["E7"].number_format = "0.00%"
    ws_wacc["E7"].alignment = align_right
    ws_wacc["G7"] = "10-year Treasury bond"
    ws_wacc["G7"].font = font_gray_italic
    
    ws_wacc["B8"] = "  Beta"
    ws_wacc["E8"] = 1.12
    ws_wacc["E8"].font = font_blue_input # Blue (Hardcoded input)
    ws_wacc["E8"].number_format = "0.00"
    ws_wacc["E8"].alignment = align_right
    ws_wacc["G8"] = "Yahoo 5 year beta"
    ws_wacc["G8"].font = font_gray_italic
    
    ws_wacc["B9"] = "  Equity Risk Premium"
    ws_wacc["E9"] = 0.0438
    ws_wacc["E9"].font = font_blue_input # Blue (Hardcoded input)
    ws_wacc["E9"].number_format = "0.00%"
    ws_wacc["E9"].alignment = align_right
    ws_wacc["G9"] = "Damodaran"
    ws_wacc["G9"].font = font_gray_italic
    
    ws_wacc["B10"] = "  Cost of Equity"
    ws_wacc["B10"].font = font_bold
    ws_wacc["E10"] = "=E7+E8*E9" # Black (Local formula)
    ws_wacc["E10"].font = font_bold
    ws_wacc["E10"].number_format = "0.00%"
    ws_wacc["E10"].alignment = align_right
    ws_wacc["E10"].border = border_top_thin
    
    # Cost of Debt section
    ws_wacc["B12"] = "Cost of Debt:"
    ws_wacc["B12"].font = font_bold
    
    ws_wacc["B13"] = "  Risk-Free Rate"
    ws_wacc["E13"] = "=E7" # Black (Local formula)
    ws_wacc["E13"].font = font_regular
    ws_wacc["E13"].number_format = "0.00%"
    ws_wacc["E13"].alignment = align_right
    ws_wacc["G13"] = "10-year Treasury bond"
    ws_wacc["G13"].font = font_gray_italic
    
    ws_wacc["B14"] = "  Implied Default Premium"
    ws_wacc["E14"] = 0.0065
    ws_wacc["E14"].font = font_blue_input # Blue (Hardcoded input)
    ws_wacc["E14"].number_format = "0.00%"
    ws_wacc["E14"].alignment = align_right
    ws_wacc["G14"] = "Bond yield premium"
    ws_wacc["G14"].font = font_gray_italic
    
    ws_wacc["B15"] = "  Pre-Tax Cost of Debt"
    ws_wacc["B15"].font = font_regular
    ws_wacc["E15"] = "=E13+E14" # Black (Local formula)
    ws_wacc["E15"].font = font_regular
    ws_wacc["E15"].number_format = "0.00%"
    ws_wacc["E15"].alignment = align_right
    
    ws_wacc["B16"] = "  Marginal Tax Rate"
    ws_wacc["E16"] = "=Statements!I55" # Green (Formula referencing Statements sheet!)
    ws_wacc["E16"].font = font_green_link
    ws_wacc["E16"].number_format = "0.00%"
    ws_wacc["E16"].alignment = align_right
    
    ws_wacc["B17"] = "  After-Tax Cost of Debt"
    ws_wacc["B17"].font = font_bold
    ws_wacc["E17"] = "=E15*(1-E16)" # Black (Local formula)
    ws_wacc["E17"].font = font_bold
    ws_wacc["E17"].number_format = "0.00%"
    ws_wacc["E17"].alignment = align_right
    ws_wacc["E17"].border = border_top_thin
    
    # Separator Line
    for c in range(2, 8):
        ws_wacc.cell(row=19, column=c).border = Border(bottom=thin_line)
        
    # Market Value of Common Equity section
    ws_wacc["B20"] = "Market Value of Common Equity:"
    ws_wacc["B20"].font = font_bold
    
    ws_wacc["B21"] = "  Total Shares Outstanding"
    ws_wacc["E21"] = "=Statements!F8" # Green (Formula referencing Statements sheet!)
    ws_wacc["E21"].font = font_green_link
    ws_wacc["E21"].number_format = "#,##0"
    ws_wacc["E21"].alignment = align_right
    
    ws_wacc["B22"] = "  Current Stock Price"
    ws_wacc["E22"] = "=Statements!F7" # Green (Formula referencing Statements sheet!)
    ws_wacc["E22"].font = font_green_link
    ws_wacc["E22"].number_format = "0.00"
    ws_wacc["E22"].alignment = align_right
    
    ws_wacc["B23"] = "  MV of Equity"
    ws_wacc["B23"].font = font_bold
    ws_wacc["E23"] = "=E21*E22" # Black (Local formula)
    ws_wacc["E23"].font = font_bold
    ws_wacc["E23"].number_format = "#,##0"
    ws_wacc["E23"].alignment = align_right
    ws_wacc["E23"].border = border_top_thin
    
    ws_wacc["B24"] = "    Weight %"
    ws_wacc["B24"].font = font_gray_italic
    ws_wacc["E24"] = "=E23/E33" # Black (Local formula)
    ws_wacc["E24"].font = font_gray_italic
    ws_wacc["E24"].number_format = "0.00%"
    ws_wacc["E24"].alignment = align_right
    
    # Market Value of Debt section
    ws_wacc["B26"] = "Market Value of Debt:"
    ws_wacc["B26"].font = font_bold
    
    ws_wacc["B27"] = "  Short-Term Debt"
    ws_wacc["E27"] = "=Statements!I80" # Green (Formula referencing Statements sheet!)
    ws_wacc["E27"].font = font_green_link
    ws_wacc["E27"].number_format = "#,##0"
    ws_wacc["E27"].alignment = align_right
    
    ws_wacc["B28"] = "  Long-Term Debt"
    ws_wacc["E28"] = "=Statements!I85" # Green (Formula referencing Statements sheet!)
    ws_wacc["E28"].font = font_green_link
    ws_wacc["E28"].number_format = "#,##0"
    ws_wacc["E28"].alignment = align_right
    
    ws_wacc["B29"] = "  PV of Operating Leases"
    ws_wacc["E29"] = "=Statements!I81+Statements!I86" # Green (Formula referencing Statements sheet!)
    ws_wacc["E29"].font = font_green_link
    ws_wacc["E29"].number_format = "#,##0"
    ws_wacc["E29"].alignment = align_right
    
    ws_wacc["B30"] = "  MV of Total Debt"
    ws_wacc["B30"].font = font_bold
    ws_wacc["E30"] = "=E27+E28+E29" # Black (Local formula)
    ws_wacc["E30"].font = font_bold
    ws_wacc["E30"].number_format = "#,##0"
    ws_wacc["E30"].alignment = align_right
    ws_wacc["E30"].border = border_top_thin
    
    ws_wacc["B31"] = "    Weight %"
    ws_wacc["B31"].font = font_gray_italic
    ws_wacc["E31"] = "=E30/E33" # Black (Local formula)
    ws_wacc["E31"].font = font_gray_italic
    ws_wacc["E31"].number_format = "0.00%"
    ws_wacc["E31"].alignment = align_right
    
    # Market Value of Company
    ws_wacc["B33"] = "Market Value of Company"
    ws_wacc["B33"].font = font_bold
    ws_wacc["E33"] = "=E23+E30" # Black (Local formula)
    ws_wacc["E33"].font = font_bold
    ws_wacc["E33"].number_format = "#,##0"
    ws_wacc["E33"].alignment = align_right
    ws_wacc["E33"].border = border_total
    
    # Final WACC Output
    for c in range(2, 8):
        ws_wacc.cell(row=35, column=c).border = Border(bottom=thin_line)
        
    ws_wacc["B36"] = "Weighted Average Cost of Capital"
    ws_wacc["B36"].font = font_bold
    ws_wacc["E36"] = "=E10*E24+E17*E31" # Black (Local formula)
    ws_wacc["E36"].font = font_bold
    ws_wacc["E36"].number_format = "0.00%"
    ws_wacc["E36"].alignment = align_right
    ws_wacc["E36"].border = border_all_thin # Boxed WACC cell!
    
    ws_wacc.column_dimensions["A"].width = 3
    ws_wacc.column_dimensions["B"].width = 45
    ws_wacc.column_dimensions["C"].width = 3
    ws_wacc.column_dimensions["D"].width = 3
    ws_wacc.column_dimensions["E"].width = 15
    ws_wacc.column_dimensions["F"].width = 3
    ws_wacc.column_dimensions["G"].width = 35

    # =========================================================================
    # TAB 3: Valuation (Discounted Cash Flow Model)
    # =========================================================================
    ws_val = wb.create_sheet(title="Valuation")
    ws_val.views.sheetView[0].showGridLines = True
    
    ws_val.row_dimensions[2].height = 25
    ws_val.row_dimensions[4].height = 24
    
    ws_val["B2"] = "Discounted Cash Flow (DCF) Valuation Model"
    ws_val["B2"].font = font_title
    
    # Section Header band
    format_sec_row(ws_val, 4)
    ws_val.merge_cells("B4:D4")
    ws_val["B4"] = "Discounted Cash Flow Projection"
    ws_val["B4"].alignment = align_left
    
    # Projected and Historical column years (linked from Statements sheet!)
    for col_idx, yr in enumerate(years):
        col_let = cols[col_idx]
        stmt_col = cols[col_idx]
        ws_val[f"{col_let}4"] = f"=Statements!{stmt_col}11" # Green (Formula referencing another sheet!)
        ws_val[f"{col_let}4"].font = font_green_bold
        ws_val[f"{col_let}4"].alignment = align_right
        
    # EBIT (Row 6)
    ws_val["B6"] = "Operating Income (EBIT)"
    ws_val["B6"].font = font_regular
    for col in cols:
        ws_val[f"{col}6"] = f"=Statements!{col}26" # Green (Formula referencing another sheet!)
        ws_val[f"{col}6"].font = font_green_link
        ws_val[f"{col}6"].number_format = "#,##0"
        ws_val[f"{col}6"].alignment = align_right
        
    # Tax Rate (Row 7)
    ws_val["B7"] = "Tax Rate"
    ws_val["B7"].font = font_regular
    for col in cols:
        ws_val[f"{col}7"] = f"=Statements!{col}55" # Green (Formula referencing another sheet!)
        ws_val[f"{col}7"].font = font_green_link
        ws_val[f"{col}7"].number_format = "0.0%"
        ws_val[f"{col}7"].alignment = align_right
        
    # Taxes on EBIT (Row 8)
    ws_val["B8"] = "Taxes on EBIT"
    ws_val["B8"].font = font_regular
    for col in cols:
        ws_val[f"{col}8"] = f"={col}6*{col}7" # Black (Local formula)
        ws_val[f"{col}8"].font = font_regular
        ws_val[f"{col}8"].number_format = "#,##0"
        ws_val[f"{col}8"].alignment = align_right
        
    # NOPAT (Row 9)
    ws_val["B9"] = "NOPAT (Net Operating Profit After Tax)"
    ws_val["B9"].font = font_bold
    for col in cols:
        ws_val[f"{col}9"] = f"={col}6-{col}8" # Black (Local formula)
        ws_val[f"{col}9"].font = font_bold
        ws_val[f"{col}9"].number_format = "#,##0"
        ws_val[f"{col}9"].alignment = align_right
        ws_val[f"{col}9"].border = border_top_thin
        
    # Add: D&A (Row 11)
    ws_val["B11"] = "Add: Depreciation & Amortization (D&A)"
    ws_val["B11"].font = font_regular
    for col in cols:
        ws_val[f"{col}11"] = f"=Statements!{col}24" # Green (Formula referencing another sheet!)
        ws_val[f"{col}11"].font = font_green_link
        ws_val[f"{col}11"].number_format = "#,##0"
        ws_val[f"{col}11"].alignment = align_right
        
    # Less: Capex (Row 12)
    ws_val["B12"] = "Less: Capital Expenditures (Capex)"
    ws_val["B12"].font = font_regular
    for col in cols:
        ws_val[f"{col}12"] = f"=Statements!{col}110" # Green (Formula referencing another sheet!)
        ws_val[f"{col}12"].font = font_green_link
        ws_val[f"{col}12"].number_format = "#,##0"
        ws_val[f"{col}12"].alignment = align_right
        
    # Less: NWC Change (Row 13)
    ws_val["B13"] = "Less: Change in Net Working Capital"
    ws_val["B13"].font = font_regular
    for col in cols:
        ws_val[f"{col}13"] = f"=Statements!{col}105+Statements!{col}106" # Green (Formula referencing another sheet!)
        ws_val[f"{col}13"].font = font_green_link
        ws_val[f"{col}13"].number_format = "#,##0"
        ws_val[f"{col}13"].alignment = align_right
        
    # FCFF (Row 14)
    ws_val["B14"] = "Free Cash Flow to Firm (FCFF)"
    ws_val["B14"].font = font_bold
    for col in cols:
        ws_val[f"{col}14"] = f"={col}9+{col}11+{col}12+{col}13" # Black (Local formula)
        ws_val[f"{col}14"].font = font_bold
        ws_val[f"{col}14"].number_format = "#,##0"
        ws_val[f"{col}14"].alignment = align_right
        ws_val[f"{col}14"].border = border_total
        
    # Discount Period (Row 16)
    ws_val["B16"] = "Discount Period"
    ws_val["B16"].font = font_gray
    for idx, col in enumerate(cols):
        ws_val[f"{col}16"] = "" if idx < 5 else idx - 4 # only project years have periods
        ws_val[f"{col}16"].font = font_gray
        ws_val[f"{col}16"].alignment = align_right
        
    # Discount Factor (Row 17)
    ws_val["B17"] = "Discount Factor"
    ws_val["B17"].font = font_gray
    for idx, col in enumerate(cols):
        if idx < 5:
            ws_val[f"{col}17"] = ""
        else:
            ws_val[f"{col}17"] = f"=1/((1+WACC!$E$36)^{col}16)" # Green (Formula referencing WACC sheet!)
            ws_val[f"{col}17"].font = font_green_link
            ws_val[f"{col}17"].number_format = "0.0000"
            ws_val[f"{col}17"].alignment = align_right
            
    # Present Value of FCF (Row 18)
    ws_val["B18"] = "Present Value of FCF"
    ws_val["B18"].font = font_bold
    for idx, col in enumerate(cols):
        if idx < 5:
            ws_val[f"{col}18"] = ""
        else:
            ws_val[f"{col}18"] = f"={col}14*{col}17" # Black (Local formula)
            ws_val[f"{col}18"].font = font_bold
            ws_val[f"{col}18"].number_format = "#,##0"
            ws_val[f"{col}18"].alignment = align_right
            ws_val[f"{col}18"].border = border_total

    # ----------------------------------------------------
    # Valuation Output Block (Column B and C, Rows 21-31)
    # ----------------------------------------------------
    ws_val["B21"] = "Valuation Summary"
    ws_val["B21"].font = font_sub_header
    ws_val["B21"].fill = fill_sub_header
    ws_val["B21"].alignment = align_left
    
    ws_val["B22"] = "Cumulative PV of FCFs"
    ws_val["B22"].font = font_regular
    ws_val["C22"] = "=SUM(J18:N18)" # Black (Local formula)
    ws_val["C22"].font = font_bold
    ws_val["C22"].number_format = "$#,##0"
    ws_val["C22"].alignment = align_right
    ws_val["C22"].border = border_all_thin
    
    ws_val["B23"] = "Terminal Value"
    ws_val["B23"].font = font_regular
    ws_val["C23"] = "=N14*(1+Statements!$L$8)/(WACC!$E$36-Statements!$L$8)" # Green (Formula referencing WACC and Statements!)
    ws_val["C23"].font = font_green_link
    ws_val["C23"].number_format = "$#,##0"
    ws_val["C23"].alignment = align_right
    ws_val["C23"].border = border_all_thin
    
    ws_val["B24"] = "PV of Terminal Value"
    ws_val["B24"].font = font_regular
    ws_val["C24"] = "=C23*N17" # Black (Local formula)
    ws_val["C24"].font = font_bold
    ws_val["C24"].number_format = "$#,##0"
    ws_val["C24"].alignment = align_right
    ws_val["C24"].border = border_all_thin
    
    ws_val["B25"] = "Implied Enterprise Value"
    ws_val["B25"].font = font_bold
    ws_val["C25"] = "=C22+C24" # Black (Local formula)
    ws_val["C25"].font = font_bold
    ws_val["C25"].number_format = "$#,##0"
    ws_val["C25"].alignment = align_right
    ws_val["C25"].border = border_all_thin
    
    ws_val["B26"] = "Less: Net Debt"
    ws_val["B26"].font = font_regular
    ws_val["C26"] = "=Statements!I8-Statements!I7" # Green (Formula referencing Statements sheet!)
    ws_val["C26"].font = font_green_link
    ws_val["C26"].number_format = "$#,##0"
    ws_val["C26"].alignment = align_right
    ws_val["C26"].border = border_all_thin
    
    ws_val["B27"] = "Implied Equity Value"
    ws_val["B27"].font = font_bold
    ws_val["C27"] = "=C25-C26" # Black (Local formula)
    ws_val["C27"].font = font_bold
    ws_val["C27"].number_format = "$#,##0"
    ws_val["C27"].alignment = align_right
    ws_val["C27"].border = border_all_thin
    
    ws_val["B28"] = "Diluted Shares Outstanding"
    ws_val["B28"].font = font_regular
    ws_val["C28"] = "=Statements!F8" # Green (Formula referencing Statements sheet!)
    ws_val["C28"].font = font_green_link
    ws_val["C28"].number_format = "#,##0"
    ws_val["C28"].alignment = align_right
    ws_val["C28"].border = border_all_thin
    
    ws_val["B29"] = "Implied Share Price"
    ws_val["B29"].font = font_green_target_price
    ws_val["C29"] = "=C27/C28" # Black (Local formula)
    ws_val["C29"].font = font_green_target_price
    ws_val["C29"].number_format = "$#,##0.00"
    ws_val["C29"].alignment = align_right
    ws_val["C29"].border = Border(left=thin_line, right=thin_line, top=thin_line, bottom=double_line)
    
    ws_val["B30"] = "Current Share Price"
    ws_val["B30"].font = font_regular
    ws_val["C30"] = "=Statements!F7" # Green (Formula referencing Statements sheet!)
    ws_val["C30"].font = font_bold
    ws_val["C30"].number_format = "$#,##0.00"
    ws_val["C30"].alignment = align_right
    ws_val["C30"].border = border_all_thin
    
    ws_val["B31"] = "Implied Upside / (Downside)"
    ws_val["B31"].font = font_bold
    ws_val["C31"] = "=(C29-C30)/C30" # Black (Local formula)
    ws_val["C31"].font = font_bold
    ws_val["C31"].number_format = "0.0%"
    ws_val["C31"].alignment = align_right
    ws_val["C31"].border = border_all_thin
    
    # Column widths
    ws_val.column_dimensions["A"].width = 3
    ws_val.column_dimensions["B"].width = 45
    ws_val.column_dimensions["C"].width = 25
    ws_val.column_dimensions["D"].width = 3
    for col_let in cols:
        ws_val.column_dimensions[col_let].width = 15
        
    wb.save("DCF_Model_Template.xlsx")
    print("Multi-tab template V2 generated successfully.")

if __name__ == "__main__":
    create_template_v2()
