import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_template_v3():
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
    
    # Color-coded Fonts
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
    
    # Company Overview Block (Row 4-8)
    format_sec_row(ws_stmt, 4)
    ws_stmt.merge_cells("B4:D4")
    ws_stmt["B4"] = "Company Overview"
    ws_stmt["B4"].alignment = align_left
    
    # Row 6
    ws_stmt["B6"] = "Company name"
    ws_stmt["B6"].font = font_regular
    ws_stmt["C6"] = "Uber Technologies, Inc."
    ws_stmt["C6"].font = font_bold
    ws_stmt["C6"].border = border_all_thin
    
    ws_stmt["E6"] = "Share price"
    ws_stmt["E6"].font = font_regular
    ws_stmt["F6"] = 72.66
    ws_stmt["F6"].font = font_blue_input # Blue (Hardcoded input)
    ws_stmt["F6"].number_format = "$#,##0.00"
    ws_stmt["F6"].border = border_all_thin
    
    ws_stmt["H6"] = "Cash"
    ws_stmt["H6"].font = font_regular
    ws_stmt["I6"] = 0.0 # Will be populated statically by macro/populator from latest quarter
    ws_stmt["I6"].font = font_bold
    ws_stmt["I6"].number_format = "$#,##0"
    ws_stmt["I6"].border = border_all_thin
    
    ws_stmt["K6"] = "WACC"
    ws_stmt["K6"].font = font_regular
    ws_stmt["L6"] = "=WACC!E36" # Green (Formula linking from WACC sheet!)
    ws_stmt["L6"].font = font_green_bold
    ws_stmt["L6"].number_format = "0.0%"
    ws_stmt["L6"].border = border_all_thin
    
    # Row 7
    ws_stmt["B7"] = "Ticker"
    ws_stmt["B7"].font = font_regular
    ws_stmt["C7"] = "UBER"
    ws_stmt["C7"].font = font_bold
    ws_stmt["C7"].border = border_all_thin
    
    ws_stmt["E7"] = "Shares outstanding"
    ws_stmt["E7"].font = font_regular
    ws_stmt["F7"] = 0.0 # Will be populated statically by macro/populator from latest quarter
    ws_stmt["F7"].font = font_bold
    ws_stmt["F7"].number_format = "#,##0"
    ws_stmt["F7"].border = border_all_thin
    
    ws_stmt["H7"] = "Debt"
    ws_stmt["H7"].font = font_regular
    ws_stmt["I7"] = 0.0 # Will be populated statically by macro/populator from latest quarter
    ws_stmt["I7"].font = font_bold
    ws_stmt["I7"].number_format = "$#,##0"
    ws_stmt["I7"].border = border_all_thin
    
    ws_stmt["K7"] = "Terminal Growth"
    ws_stmt["K7"].font = font_regular
    ws_stmt["L7"] = 0.02
    ws_stmt["L7"].font = font_blue_input # Blue (Hardcoded input)
    ws_stmt["L7"].number_format = "0.0%"
    ws_stmt["L7"].border = border_all_thin
    
    # Row 8
    ws_stmt["B8"] = "SEC Email"
    ws_stmt["B8"].font = font_regular
    ws_stmt["C8"] = "julie@example.com"
    ws_stmt["C8"].font = font_blue_input # Blue (Hardcoded input)
    ws_stmt["C8"].border = border_all_thin
    
    ws_stmt["E8"] = "Market Cap."
    ws_stmt["E8"].font = font_regular
    ws_stmt["F8"] = "=F7*F6" # Black (Local formula)
    ws_stmt["F8"].font = font_bold
    ws_stmt["F8"].number_format = "$#,##0"
    ws_stmt["F8"].border = border_all_thin
    
    ws_stmt["H8"] = "EV"
    ws_stmt["H8"].font = font_regular
    ws_stmt["I8"] = "=F8-I6+I7" # Black (Local formula)
    ws_stmt["I8"].font = font_bold
    ws_stmt["I8"].number_format = "$#,##0"
    ws_stmt["I8"].border = border_all_thin
    
    # ----------------------------------------------------
    # Income Statement Section (Row 10)
    # ----------------------------------------------------
    format_sec_row(ws_stmt, 10)
    ws_stmt.merge_cells("B10:D10")
    ws_stmt["B10"] = "Income Statement"
    ws_stmt["B10"].alignment = align_left
    
    for col, yr in zip(cols, years):
        cell_ref = f"{col}10"
        suffix = "A" if yr <= 2025 else "E"
        ws_stmt[cell_ref] = f"{yr}{suffix}"
        ws_stmt[cell_ref].alignment = align_right
        
    # Revenue (Row 12)
    ws_stmt["B12"] = "Revenue"
    ws_stmt["B12"].font = font_regular
    revs = [17455, 31877, 37281, 43793, 52017]
    for col, val in zip(cols[:5], revs):
        ws_stmt[f"{col}12"] = val
        ws_stmt[f"{col}12"].font = font_blue_input
        ws_stmt[f"{col}12"].number_format = "#,##0"
        ws_stmt[f"{col}12"].alignment = align_right
    # Projected Revenue
    for i in range(5, 10):
        col = cols[i]
        prev_col = cols[i-1]
        ws_stmt[f"{col}12"] = f"={prev_col}12*(1+{col}48)" # Black (Local formula)
        ws_stmt[f"{col}12"].font = font_regular
        ws_stmt[f"{col}12"].number_format = "#,##0"
        ws_stmt[f"{col}12"].alignment = align_right
        
    # COGS (Row 15)
    ws_stmt["B15"] = "Cost of revenue"
    ws_stmt["B15"].font = font_regular
    cogs = [-9351, -15655, -16900, -21000, -31338]
    for col, val in zip(cols[:5], cogs):
        ws_stmt[f"{col}15"] = val
        ws_stmt[f"{col}15"].font = font_blue_input
        ws_stmt[f"{col}15"].number_format = "#,##0"
        ws_stmt[f"{col}15"].alignment = align_right
    # Projected COGS
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}15"] = f"={col}12*{col}49" # Black (Local formula)
        ws_stmt[f"{col}15"].font = font_regular
        ws_stmt[f"{col}15"].number_format = "#,##0"
        ws_stmt[f"{col}15"].alignment = align_right
        
    # Gross Profit (Row 16)
    ws_stmt["B16"] = "Gross profit"
    ws_stmt["B16"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}16"] = f"={col}12+{col}15" # Black (Local formula)
        ws_stmt[f"{col}16"].font = font_bold
        ws_stmt[f"{col}16"].number_format = "#,##0"
        ws_stmt[f"{col}16"].alignment = align_right
        ws_stmt[f"{col}16"].border = border_top_thin
        
    # Operating Expenses (Row 18-23)
    ws_stmt["B18"] = "Operating expenses"
    ws_stmt["B18"].font = Font(name=font_family, size=10, italic=True, bold=True)
    
    op_exp_items = [
        ("  Operation & support", 19, [1877, 2413, 2689, 2732, 2854], 50),
        ("  Sales & marketing", 20, [4789, 4756, 4356, 4337, 4898], 51),
        ("  Research & development", 21, [2054, 2798, 3164, 3109, 3402], 52),
        ("  General & administrative", 22, [2316, 3136, 2682, 3639, 3241], 53),
        ("  Depreciation & amortization", 23, [902, 947, 823, 711, 719], 54)
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
            ws_stmt[f"{col}{r_idx}"] = f"={col}12*{col}{margin_row}" # Black (Local formula)
            ws_stmt[f"{col}{r_idx}"].font = font_regular
            ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
            ws_stmt[f"{col}{r_idx}"].alignment = align_right
            
    # Total Operating Expenses (Row 24)
    ws_stmt["B24"] = "Total operating expenses"
    ws_stmt["B24"].font = font_regular
    for col in cols:
        ws_stmt[f"{col}24"] = f"=SUM({col}19:{col}23)" # Black (Local formula)
        ws_stmt[f"{col}24"].font = font_regular
        ws_stmt[f"{col}24"].number_format = "#,##0"
        ws_stmt[f"{col}24"].alignment = align_right
        ws_stmt[f"{col}24"].border = border_top_thin
        
    # Operating Income (loss) (Row 25)
    ws_stmt["B25"] = "Operating Income (loss)"
    ws_stmt["B25"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}25"] = f"={col}16-{col}24" # Black (Local formula)
        ws_stmt[f"{col}25"].font = font_bold
        ws_stmt[f"{col}25"].number_format = "#,##0"
        ws_stmt[f"{col}25"].alignment = align_right
        ws_stmt[f"{col}25"].border = border_total
        
    # Interest expense (Row 27)
    ws_stmt["B27"] = "Interest income (expense)"
    ws_stmt["B27"].font = font_regular
    interest = [-483, -565, -633, -523, 303]
    for col, val in zip(cols[:5], interest):
        ws_stmt[f"{col}27"] = val
        ws_stmt[f"{col}27"].font = font_blue_input
        ws_stmt[f"{col}27"].number_format = "#,##0"
        ws_stmt[f"{col}27"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}27"] = 303 # Assumed constant
        ws_stmt[f"{col}27"].font = font_regular
        ws_stmt[f"{col}27"].number_format = "#,##0"
        ws_stmt[f"{col}27"].alignment = align_right
        
    # Other Income (loss) (Row 28)
    ws_stmt["B28"] = "Other Income (loss)"
    ws_stmt["B28"].font = font_regular
    other_inc = [3292, -7029, 1844, 1849, -68]
    for col, val in zip(cols[:5], other_inc):
        ws_stmt[f"{col}28"] = val
        ws_stmt[f"{col}28"].font = font_blue_input
        ws_stmt[f"{col}28"].number_format = "#,##0"
        ws_stmt[f"{col}28"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}28"] = 0 # Assumed constant
        ws_stmt[f"{col}28"].font = font_regular
        ws_stmt[f"{col}28"].number_format = "#,##0"
        ws_stmt[f"{col}28"].alignment = align_right
        
    # Income (loss) before taxes & equity (Row 29)
    ws_stmt["B29"] = "Income (loss) before taxes & equity"
    ws_stmt["B29"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}29"] = f"={col}25+{col}27+{col}28" # Black (Local formula)
        ws_stmt[f"{col}29"].font = font_bold
        ws_stmt[f"{col}29"].number_format = "#,##0"
        ws_stmt[f"{col}29"].alignment = align_right
        ws_stmt[f"{col}29"].border = border_top_thin
        
    # tax provision (Row 31)
    ws_stmt["B31"] = "provision for (benefit from) income tax"
    ws_stmt["B31"].font = font_regular
    tax = [-492, -181, 213, -5758, -4346]
    for col, val in zip(cols[:5], tax):
        ws_stmt[f"{col}31"] = val
        ws_stmt[f"{col}31"].font = font_blue_input
        ws_stmt[f"{col}31"].number_format = "#,##0"
        ws_stmt[f"{col}31"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}31"] = f"={col}29*{col}56" # Black (Local formula)
        ws_stmt[f"{col}31"].font = font_regular
        ws_stmt[f"{col}31"].number_format = "#,##0"
        ws_stmt[f"{col}31"].alignment = align_right
        
    # equity inv (Row 32)
    ws_stmt["B32"] = "Income (loss) from equity inv."
    ws_stmt["B32"].font = font_regular
    equity = [-37, 107, 48, -38, -53]
    for col, val in zip(cols[:5], equity):
        ws_stmt[f"{col}32"] = val
        ws_stmt[f"{col}32"].font = font_blue_input
        ws_stmt[f"{col}32"].number_format = "#,##0"
        ws_stmt[f"{col}32"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}32"] = 0
        ws_stmt[f"{col}32"].font = font_regular
        ws_stmt[f"{col}32"].number_format = "#,##0"
        ws_stmt[f"{col}32"].alignment = align_right
        
    # Net income (loss) incl. NCI (Row 33)
    ws_stmt["B33"] = "Net income (loss) incl. NCI"
    ws_stmt["B33"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}33"] = f"={col}29-{col}31+{col}32" # Black (Local formula)
        ws_stmt[f"{col}33"].font = font_bold
        ws_stmt[f"{col}33"].number_format = "#,##0"
        ws_stmt[f"{col}33"].alignment = align_right
        ws_stmt[f"{col}33"].border = border_top_thin
        
    # Less NCI (Row 34)
    ws_stmt["B34"] = "Less: net income (loss) attributable to NCI"
    ws_stmt["B34"].font = font_regular
    nci = [-74, 3, 269, -11, 40]
    for col, val in zip(cols[:5], nci):
        ws_stmt[f"{col}34"] = val
        ws_stmt[f"{col}34"].font = font_blue_input
        ws_stmt[f"{col}34"].number_format = "#,##0"
        ws_stmt[f"{col}34"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}34"] = 0
        ws_stmt[f"{col}34"].font = font_regular
        ws_stmt[f"{col}34"].number_format = "#,##0"
        ws_stmt[f"{col}34"].alignment = align_right
        
    # Net income (loss) (Row 35)
    ws_stmt["B35"] = "Net income (loss)"
    ws_stmt["B35"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}35"] = f"={col}33-{col}34" # Black (Local formula)
        ws_stmt[f"{col}35"].font = font_bold
        ws_stmt[f"{col}35"].number_format = "#,##0"
        ws_stmt[f"{col}35"].alignment = align_right
        ws_stmt[f"{col}35"].border = border_total
        
    # EBITDA / ADJ EBITDA SECTION (Rows 37-39)
    ws_stmt["B37"] = "EBITDA"
    ws_stmt["B37"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}37"] = f"={col}25+{col}23" # Black (Operating Income Row 25 + D&A Row 23)
        ws_stmt[f"{col}37"].font = font_bold
        ws_stmt[f"{col}37"].number_format = "#,##0"
        ws_stmt[f"{col}37"].alignment = align_right
        ws_stmt[f"{col}37"].border = border_top_thin
        
    ws_stmt["B38"] = "Stock based compensation"
    ws_stmt["B38"].font = font_regular
    for col in cols:
        # Links to SBC in Cash Flow (which shifts to Row 106!)
        ws_stmt[f"{col}38"] = f"={col}106"
        ws_stmt[f"{col}38"].font = font_regular
        ws_stmt[f"{col}38"].number_format = "#,##0"
        ws_stmt[f"{col}38"].alignment = align_right
        
    ws_stmt["B39"] = "Adjusted EBITDA"
    ws_stmt["B39"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}39"] = f"={col}37+{col}38" # EBITDA + SBC
        ws_stmt[f"{col}39"].font = font_bold
        ws_stmt[f"{col}39"].number_format = "#,##0"
        ws_stmt[f"{col}39"].alignment = align_right
        ws_stmt[f"{col}39"].border = border_total
        
    # Diluted Shares outstanding (Row 41)
    ws_stmt["B41"] = "Diluted shares outstanding"
    ws_stmt["B41"].font = font_regular
    shares = [1896, 1975, 2092, 2151, 2120]
    for col, val in zip(cols[:5], shares):
        ws_stmt[f"{col}41"] = val
        ws_stmt[f"{col}41"].font = font_blue_input
        ws_stmt[f"{col}41"].number_format = "#,##0"
        ws_stmt[f"{col}41"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws_stmt[f"{col}41"] = f"={cols[4]}41" # Keep constant at latest
        ws_stmt[f"{col}41"].font = font_bold
        ws_stmt[f"{col}41"].number_format = "#,##0"
        ws_stmt[f"{col}41"].alignment = align_right
        
    # Diluted EPS (Row 42)
    ws_stmt["B42"] = "Diluted earnings per share"
    ws_stmt["B42"].font = font_bold
    for col in cols:
        ws_stmt[f"{col}42"] = f"=IF({col}41<=0, \"NA\", {col}35/{col}41)" # Black (Local formula)
        ws_stmt[f"{col}42"].font = font_bold
        ws_stmt[f"{col}42"].number_format = "$#,##0.00"
        ws_stmt[f"{col}42"].alignment = align_right
        ws_stmt[f"{col}42"].border = border_total
        
    # Model Assumptions Header (Row 46)
    format_sec_row(ws_stmt, 46)
    ws_stmt.merge_cells("B46:D46")
    ws_stmt["B46"] = "Historical & Projected Model Assumptions"
    ws_stmt["B46"].alignment = align_left
    
    # Assumption Rows (Rows 48-57)
    ass_items = [
        ("Revenue growth rate", 48, "0.0%", "=IF({col}12_prev=0, 0, ({col}12-{col}12_prev)/ABS({col}12_prev))"),
        ("COGS % of sales", 49, "0.0%", "=-{col}15/{col}12"),
        ("Operation & support % of sales", 50, "0.0%", "={col}19/{col}12"),
        ("Sales & marketing % of sales", 51, "0.0%", "={col}20/{col}12"),
        ("R&D % of sales", 52, "0.0%", "={col}21/{col}12"),
        ("G&A % of sales", 53, "0.0%", "={col}22/{col}12"),
        ("D&A % of sales", 54, "0.0%", "={col}23/{col}12"),
        ("Operating margin", 55, "0.0%", "={col}25/{col}12"),
        ("Tax rate", 56, "0.0%", "=IF({col}29<=0, \"NA\", {col}31/{col}29)"),
        ("SBC % of sales", 57, "0.0%", "={col}38/{col}12")
    ]
    
    for label, r_idx, num_fmt, formula_tpl in ass_items:
        ws_stmt[f"B{r_idx}"] = label
        ws_stmt[f"B{r_idx}"].font = font_gray
        
        # Hist values (formulas)
        for col_idx, col in enumerate(cols[:5]):
            if "growth" in label.lower() and col_idx == 0:
                ws_stmt[f"{col}{r_idx}"] = 0.0
            else:
                prev_col = cols[col_idx-1] if col_idx > 0 else "E"
                f_str = formula_tpl.replace("{col}", col).replace("{col}_prev", prev_col)
                ws_stmt[f"{col}{r_idx}"] = f_str
            ws_stmt[f"{col}{r_idx}"].font = font_gray
            ws_stmt[f"{col}{r_idx}"].number_format = num_fmt
            ws_stmt[f"{col}{r_idx}"].alignment = align_right
            
        # Projected values (hardcoded inputs or formulas)
        for col_idx in range(5, 10):
            col = cols[col_idx]
            if "growth" in label.lower():
                # Growth inputs are blue hardcoded
                ws_stmt[f"{col}{r_idx}"] = 0.10 if r_idx == 48 else 0.05
                ws_stmt[f"{col}{r_idx}"].font = font_blue_input
            elif r_idx in [49, 50, 51, 52, 53, 54, 57]:
                # Operating cost percentages are blue hardcoded
                prev_col = cols[col_idx-1]
                ws_stmt[f"{col}{r_idx}"] = f"={prev_col}{r_idx}" # links to previous year as default
                ws_stmt[f"{col}{r_idx}"].font = font_blue_input
            elif r_idx == 56:
                # Tax rate is blue hardcoded
                ws_stmt[f"{col}{r_idx}"] = 0.15
                ws_stmt[f"{col}{r_idx}"].font = font_blue_input
            else:
                # Margins / formulas are black
                prev_col = cols[col_idx-1]
                f_str = formula_tpl.replace("{col}", col).replace("{col}_prev", prev_col)
                ws_stmt[f"{col}{r_idx}"] = f_str
                ws_stmt[f"{col}{r_idx}"].font = font_regular
                
            ws_stmt[f"{col}{r_idx}"].number_format = num_fmt
            ws_stmt[f"{col}{r_idx}"].alignment = align_right
            ws_stmt[f"{col}{r_idx}"].border = border_all_thin
            
    # ----------------------------------------------------
    # Balance Sheet Section (Row 59)
    # ----------------------------------------------------
    format_sec_row(ws_stmt, 59)
    ws_stmt.merge_cells("B59:D59")
    ws_stmt["B59"] = "Balance Sheet"
    ws_stmt["B59"].alignment = align_left
    
    for col, yr in zip(cols, years):
        cell_ref = f"{col}59"
        suffix = "A" if yr <= 2025 else "E"
        ws_stmt[cell_ref] = f"{yr}{suffix}"
        ws_stmt[cell_ref].alignment = align_right
        
    ws_stmt["B61"] = "Assets"
    ws_stmt["B61"].font = Font(name=font_family, size=10, italic=True, bold=True)
    
    bs_items = [
        # Assets
        ("  Cash equivalents", 63, [3850, 4830, 4680, 5680, 6120], "Asset"),
        ("  Short-term investments", 64, [0, 0, 0, 0, 0], "Asset"),
        ("  Restricted cash & cash equivalents (current)", 65, [0, 0, 0, 0, 0], "Asset"),
        ("  Accounts receivable", 66, [1920, 2410, 2580, 3110, 3350], "Asset"),
        ("  Prepaid expenses & other current assets", 67, [1210, 1340, 1420, 1560, 1680], "Asset"),
        ("Total current assets", 68, None, "AssetTotalCurrent"),
        ("  Restricted cash & cash equivalents (non-current)", 69, [0, 0, 0, 0, 0], "Asset"),
        ("  Restricted investments (non-current)", 70, [0, 0, 0, 0, 0], "Asset"),
        ("  Investments (non-current)", 71, [1560, 2110, 2340, 2450, 2560], "Asset"),
        ("  Equity method investments", 72, [0, 0, 0, 0, 0], "Asset"),
        ("  Property, plant & equipment (PP&E)", 73, [1890, 2140, 2310, 2560, 2680], "Asset"),
        ("  Operating lease assets", 74, [1120, 1340, 1420, 1560, 1690], "Asset"),
        ("  Intangible assets", 75, [2450, 3110, 2980, 3450, 3120], "Asset"),
        ("  Goodwill", 76, [18900, 21500, 22100, 23400, 24500], "Asset"),
        ("  Other non-current assets", 77, [1230, 1450, 1560, 1780, 1890], "Asset"),
        ("Total assets", 78, None, "AssetTotal"),
        
        # Liabilities
        ("Liabilities & Shareholders' Equity", 80, None, "LabelOnly"),
        ("  Accounts payable", 81, [1450, 1890, 1920, 2110, 2240], "Liab"),
        ("  Short-term debt & current portion of LT debt", 82, [0, 0, 0, 0, 0], "Liab"),
        ("  Operating lease liabilities (current)", 83, [240, 310, 340, 390, 412], "Liab"),
        ("  Accrued liabilities & other current liabilities", 84, [4320, 5110, 5680, 6120, 6450], "Liab"),
        ("Total current liabilities", 85, None, "LiabTotalCurrent"),
        ("  Long-term insurance reserves", 86, [0, 0, 0, 0, 0], "Liab"),
        ("  Long-term debt", 87, [7450, 9120, 9450, 11200, 12100], "Liab"),
        ("  Operating lease liabilities (non-current)", 88, [920, 1050, 1140, 1230, 1312], "Liab"),
        ("  Other non-current liabilities", 89, [2340, 2890, 3110, 3450, 3680], "Liab"),
        ("Total liabilities", 90, None, "LiabTotal"),
        ("  Redeemable non-controlling interests (NCI)", 91, [0, 0, 0, 0, 0], "Liab"),
        
        # Equity
        ("  Common stock & APIC", 94, [14500, 15600, 16100, 17800, 18900], "Equity"),
        ("  Accumulated other comprehensive income (OCI)", 95, [-110, -230, -310, -290, -320], "Equity"),
        ("  Accumulated deficit / Retained earnings", 96, [-4500, -6100, -7400, -9100, -10093], "Equity"),
        ("  Non-controlling interests", 98, [0, 0, 0, 0, 0], "Equity"),
        ("Total equity", 99, None, "EquityTotal"),
        ("Total liabilities & equity", 100, None, "LiabEquityTotal")
    ]
    
    for label, r_idx, val_data, group in bs_items:
        if group == "LabelOnly":
            ws_stmt[f"B{r_idx}"] = label
            ws_stmt[f"B{r_idx}"].font = Font(name=font_family, size=10, italic=True, bold=True)
            continue
            
        ws_stmt[f"B{r_idx}"] = label
        ws_stmt[f"B{r_idx}"].font = font_bold if "Total" in group else font_regular
        
        if val_data is not None:
            # Write historical values
            for col, val in zip(cols[:5], val_data):
                ws_stmt[f"{col}{r_idx}"] = val
                ws_stmt[f"{col}{r_idx}"].font = font_blue_input
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
                
            # Write projected values (keep constant at latest historical value)
            for i in range(5, 10):
                col = cols[i]
                ws_stmt[f"{col}{r_idx}"] = f"={cols[4]}{r_idx}"
                ws_stmt[f"{col}{r_idx}"].font = font_regular
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
                ws_stmt[f"{col}{r_idx}"].border = border_all_thin
        else:
            # Formula / Totals
            if label == "Total current assets":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}63:{col}67)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Total assets":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}68+{col}69+{col}70+{col}71+{col}72+{col}73+{col}74+{col}75+{col}76+{col}77"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Total current liabilities":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}81:{col}84)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Total liabilities":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}85+SUM({col}86:{col}89)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Total equity":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}94:{col}96)+{col}98"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Total liabilities & equity":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}90+{col}91+{col}99"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
                    
    # ----------------------------------------------------
    # Cash Flow Statement (Row 102)
    # ----------------------------------------------------
    format_sec_row(ws_stmt, 102)
    ws_stmt.merge_cells("B102:D102")
    ws_stmt["B102"] = "Statement of Cash Flows"
    ws_stmt["B102"].alignment = align_left
    
    for col, yr in zip(cols, years):
        cell_ref = f"{col}102"
        suffix = "A" if yr <= 2025 else "E"
        ws_stmt[cell_ref] = f"{yr}{suffix}"
        ws_stmt[cell_ref].alignment = align_right
        
    cf_items = [
        ("Net income", 104, "=E35", "Operating"),
        ("Depreciation & Amortization (D&A)", 105, "=E23", "Operating"),
        ("Stock-based compensation", 106, [213, 254, 301, 345, 412], "Operating"),
        ("Change in net working capital assets", 107, [-110, -234, -412, -311, -298], "Operating"),
        ("Change in net working capital liabilities", 108, [145, 342, 287, 432, 290], "Operating"),
        ("Other non-cash operating assets", 109, [0, 0, 0, 0, 0], "Operating"),
        ("Other non-cash operating liabilities", 110, [0, 0, 0, 0, 0], "Operating"),
        ("Cash from operating activities", 111, None, "OperatingTotal"),
        ("Capital expenditures", 112, [-234, -412, -567, -612, -528], "Investing"),
        ("Cash from investing activities", 113, None, "InvestingTotal"),
        ("Net change in long-term debt", 114, [1452, 2341, -124, -432, 856], "Financing"),
        ("Commercial paper / Revolver change", 115, [0, 0, 0, 0, 0], "Financing"),
        ("Share repurchases", 116, [0, 0, -43, -112, -234], "Financing"),
        ("Common dividends", 117, [0, 0, 0, 0, 0], "Financing"),
        ("Cash from financing activities", 118, None, "FinancingTotal"),
        ("Net change in cash during period", 119, None, "NetChangeTotal")
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
                ws_stmt[f"{col}{r_idx}"].border = border_all_thin
        elif isinstance(val_data, str):
            # Formula linking other rows/cells locally
            for col in cols:
                cell_formula = val_data.replace("E", col)
                ws_stmt[f"{col}{r_idx}"] = cell_formula
                ws_stmt[f"{col}{r_idx}"].font = font_regular
                ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                ws_stmt[f"{col}{r_idx}"].alignment = align_right
                ws_stmt[f"{col}{r_idx}"].border = border_all_thin
        else:
            # Totals
            if label == "Cash from operating activities":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}104:{col}110)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Cash from investing activities":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}112"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Cash from financing activities":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"=SUM({col}114:{col}117)"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
            elif label == "Net change in cash during period":
                for col in cols:
                    ws_stmt[f"{col}{r_idx}"] = f"={col}111+{col}113+{col}118"
                    ws_stmt[f"{col}{r_idx}"].font = font_bold
                    ws_stmt[f"{col}{r_idx}"].number_format = "#,##0"
                    ws_stmt[f"{col}{r_idx}"].alignment = align_right
                    ws_stmt[f"{col}{r_idx}"].border = border_total
                    
    # Set Statements column widths
    ws_stmt.column_dimensions["A"].width = 3
    ws_stmt.column_dimensions["B"].width = 45
    ws_stmt.column_dimensions["C"].width = 25
    ws_stmt.column_dimensions["D"].width = 3
    for col_let in cols:
        ws_stmt.column_dimensions[col_let].width = 15
        
    # =========================================================================
    # TAB 2: WACC (Weighted Average Cost of Capital)
    # =========================================================================
    ws_wacc = wb.create_sheet(title="WACC")
    ws_wacc.views.sheetView[0].showGridLines = True
    
    ws_wacc.row_dimensions[2].height = 25
    ws_wacc.row_dimensions[4].height = 24
    
    ws_wacc["B2"] = "Weighted Average Cost of Capital (WACC) Calculation"
    ws_wacc["B2"].font = font_title
    
    format_sec_row(ws_wacc, 4)
    ws_wacc.merge_cells("B4:D4")
    ws_wacc["B4"] = "Cost of Capital Components"
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
    ws_wacc["E16"] = "=Statements!I56" # Green (Formula referencing Statements sheet Row 56!)
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
    ws_wacc["E21"] = "=Statements!F7" # Green (Formula referencing Statements sheet cell F7!)
    ws_wacc["E21"].font = font_green_link
    ws_wacc["E21"].number_format = "#,##0.00"
    ws_wacc["E21"].alignment = align_right
    
    ws_wacc["B22"] = "  Current Stock Price"
    ws_wacc["E22"] = "=Statements!F6" # Green (Formula referencing Statements sheet cell F6!)
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
    ws_wacc["E27"] = "=Statements!I82" # Green (Formula referencing Statements sheet Row 82!)
    ws_wacc["E27"].font = font_green_link
    ws_wacc["E27"].number_format = "#,##0"
    ws_wacc["E27"].alignment = align_right
    
    ws_wacc["B28"] = "  Long-Term Debt"
    ws_wacc["E28"] = "=Statements!I87" # Green (Formula referencing Statements sheet Row 87!)
    ws_wacc["E28"].font = font_green_link
    ws_wacc["E28"].number_format = "#,##0"
    ws_wacc["E28"].alignment = align_right
    
    ws_wacc["B29"] = "  PV of Operating Leases"
    ws_wacc["E29"] = "=Statements!I83+Statements!I88" # Green (Formula referencing Statements sheet Row 83+88!)
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
    
    # Projected and Historical column years (linked from Statements sheet Row 10!)
    for col_idx, yr in enumerate(years):
        col_let = cols[col_idx]
        stmt_col = cols[col_idx]
        ws_val[f"{col_let}4"] = f"=Statements!{stmt_col}10" # Green (Formula referencing another sheet!)
        ws_val[f"{col_let}4"].font = font_green_bold
        ws_val[f"{col_let}4"].alignment = align_right
        
    # EBIT (Row 6)
    ws_val["B6"] = "Operating Income (EBIT)"
    ws_val["B6"].font = font_regular
    for col in cols:
        ws_val[f"{col}6"] = f"=Statements!{col}25" # Green (Formula referencing another sheet Row 25!)
        ws_val[f"{col}6"].font = font_green_link
        ws_val[f"{col}6"].number_format = "#,##0"
        ws_val[f"{col}6"].alignment = align_right
        
    # Tax Rate (Row 7)
    ws_val["B7"] = "Tax Rate"
    ws_val["B7"].font = font_regular
    for col in cols:
        ws_val[f"{col}7"] = f"=Statements!{col}56" # Green (Formula referencing another sheet Row 56!)
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
        ws_val[f"{col}11"] = f"=Statements!{col}23" # Green (Formula referencing another sheet Row 23!)
        ws_val[f"{col}11"].font = font_green_link
        ws_val[f"{col}11"].number_format = "#,##0"
        ws_val[f"{col}11"].alignment = align_right
        
    # Less: Capex (Row 12)
    ws_val["B12"] = "Less: Capital Expenditures (Capex)"
    ws_val["B12"].font = font_regular
    for col in cols:
        ws_val[f"{col}12"] = f"=Statements!{col}112" # Green (Formula referencing another sheet Row 112!)
        ws_val[f"{col}12"].font = font_green_link
        ws_val[f"{col}12"].number_format = "#,##0"
        ws_val[f"{col}12"].alignment = align_right
        
    # Less: NWC Change (Row 13)
    ws_val["B13"] = "Less: Change in Net Working Capital"
    ws_val["B13"].font = font_regular
    for col in cols:
        ws_val[f"{col}13"] = f"=Statements!{col}107+Statements!{col}108" # Green (Formula referencing another sheet Row 107+108!)
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
        ws_val[f"{col}16"] = "" if idx < 5 else idx - 4
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
    ws_val["C23"] = "=N14*(1+Statements!$L$7)/(WACC!$E$36-Statements!$L$7)" # Green (Formula referencing WACC and Statements!)
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
    ws_val["C26"] = "=Statements!I7-Statements!I6" # Green (Formula referencing Statements sheet Debt in I7 and Cash in I6!)
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
    ws_val["C28"] = "=Statements!F7" # Green (Formula referencing Statements sheet cell F7!)
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
    ws_val["C30"] = "=Statements!F6" # Green (Formula referencing Statements sheet cell F6!)
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
    print("Multi-tab template V3 generated successfully.")

if __name__ == "__main__":
    create_template_v3()
