import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_template():
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    # Create Statements sheet
    ws = wb.create_sheet(title="Statements")
    
    # Enable grid lines
    ws.views.sheetView[0].showGridLines = True
    
    # Define styles
    font_family = "Segoe UI"
    
    font_title = Font(name=font_family, size=16, bold=True, color="000000")
    font_subtitle = Font(name=font_family, size=9, italic=True, color="595959")
    
    font_sec_header = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    fill_sec_header = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")
    
    font_sub_header = Font(name=font_family, size=10, bold=True, color="000000")
    fill_sub_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    font_bold = Font(name=font_family, size=10, bold=True, color="000000")
    font_regular = Font(name=font_family, size=10, color="000000")
    font_gray = Font(name=font_family, size=9, color="595959")
    font_blue_input = Font(name=font_family, size=10, color="003399") # Blue for hardcoded inputs
    
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
    
    # Helper function to fill a range of cells with the section header style
    def format_sec_row(row_idx):
        ws.row_dimensions[row_idx].height = 24
        # Format B to N
        for c in range(2, 15):
            col_letter = get_column_letter(c)
            cell = ws[f"{col_letter}{row_idx}"]
            cell.fill = fill_sec_header
            cell.font = font_sec_header
            
    # Row Heights
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 18
    
    # Write Title
    ws["B2"] = "Statements"
    ws["B2"].font = font_title
    ws["B3"] = "$ and shares in million, except per share data; share price as of July 1st, 4 p.m."
    ws["B3"].font = font_subtitle
    
    # ----------------------------------------------------
    # Company Overview
    # ----------------------------------------------------
    format_sec_row(5)
    ws.merge_cells("B5:D5")
    ws["B5"] = "Company Overview"
    ws["B5"].alignment = align_left
    
    ws["B7"] = "Company name"
    ws["B7"].font = font_regular
    ws["C7"] = "Uber Technologies, Inc."
    ws["C7"].font = font_bold
    ws["C7"].border = border_all_thin
    
    ws["E7"] = "Share price"
    ws["E7"].font = font_regular
    ws["F7"] = 72.66
    ws["F7"].font = font_blue_input
    ws["F7"].number_format = "$#,##0.00"
    ws["F7"].border = border_all_thin
    
    ws["H7"] = "Cash"
    ws["H7"].font = font_regular
    ws["I7"] = 7633
    ws["I7"].font = font_blue_input
    ws["I7"].number_format = "$#,##0"
    ws["I7"].border = border_all_thin
    
    ws["K7"] = "WACC"
    ws["K7"].font = font_regular
    ws["L7"] = 0.09
    ws["L7"].font = font_blue_input
    ws["L7"].number_format = "0.0%"
    ws["L7"].border = border_all_thin
    
    ws["B8"] = "Ticker"
    ws["B8"].font = font_regular
    ws["C8"] = "UBER"
    ws["C8"].font = font_bold
    ws["C8"].border = border_all_thin
    
    ws["E8"] = "Shares outstanding"
    ws["E8"].font = font_regular
    ws["F8"] = 2036
    ws["F8"].font = font_blue_input
    ws["F8"].number_format = "#,##0"
    ws["F8"].border = border_all_thin
    
    ws["H8"] = "Debt"
    ws["H8"].font = font_regular
    ws["I8"] = 10590
    ws["I8"].font = font_blue_input
    ws["I8"].number_format = "$#,##0"
    ws["I8"].border = border_all_thin
    
    ws["K8"] = "Terminal Growth"
    ws["K8"].font = font_regular
    ws["L8"] = 0.02
    ws["L8"].font = font_blue_input
    ws["L8"].number_format = "0.0%"
    ws["L8"].border = border_all_thin
    
    ws["E9"] = "Market Cap."
    ws["E9"].font = font_regular
    ws["F9"] = "=F8*F7"
    ws["F9"].font = font_bold
    ws["F9"].number_format = "$#,##0"
    ws["F9"].border = border_all_thin
    
    ws["H9"] = "EV"
    ws["H9"].font = font_regular
    ws["I9"] = "=F9-I7+I8"
    ws["I9"].font = font_bold
    ws["I9"].number_format = "$#,##0"
    ws["I9"].border = border_all_thin
    
    # ----------------------------------------------------
    # Income Statement Headers
    # ----------------------------------------------------
    format_sec_row(11)
    ws.merge_cells("B11:D11")
    ws["B11"] = "Income Statement"
    ws["B11"].alignment = align_left
    
    cols = ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
    years = [2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]
    
    for col, yr in zip(cols, years):
        cell_ref = f"{col}11"
        suffix = "A" if yr <= 2025 else "E"
        ws[cell_ref] = f"{yr}{suffix}"
        ws[cell_ref].alignment = align_right
    
    # ----------------------------------------------------
    # Income Statement Rows
    # ----------------------------------------------------
    # Revenue (Row 13)
    ws["B13"] = "Revenue"
    ws["B13"].font = font_bold
    # Historical Revenue
    rev_vals = [17455, 31877, 37281, 43978, 52017]
    for col, val in zip(cols[:5], rev_vals):
        ws[f"{col}13"] = val
        ws[f"{col}13"].font = font_blue_input
        ws[f"{col}13"].number_format = "#,##0"
        ws[f"{col}13"].alignment = align_right
    # Projected Revenue
    for i in range(5, 10):
        col = cols[i]
        prev_col = cols[i-1]
        ws[f"{col}13"] = f"={prev_col}13*(1+{col}47)"
        ws[f"{col}13"].font = font_bold
        ws[f"{col}13"].number_format = "#,##0"
        ws[f"{col}13"].alignment = align_right
        
    # Growth % (Row 14)
    ws["B14"] = "Growth %"
    ws["B14"].font = font_gray
    for i in range(1, 10):
        col = cols[i]
        prev_col = cols[i-1]
        ws[f"{col}14"] = f"=({col}13-{prev_col}13)/{prev_col}13"
        ws[f"{col}14"].font = font_gray
        ws[f"{col}14"].number_format = "0.0%"
        ws[f"{col}14"].alignment = align_right
        
    # COGS (Row 16)
    ws["B16"] = "COSG"
    ws["B16"].font = font_regular
    cogs_vals = [-9351, -19659, -22457, -26651, -31338]
    for col, val in zip(cols[:5], cogs_vals):
        ws[f"{col}16"] = val
        ws[f"{col}16"].font = font_blue_input
        ws[f"{col}16"].number_format = "#,##0"
        ws[f"{col}16"].alignment = align_right
    # Projected COGS
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}16"] = f"={col}13*{col}48"
        ws[f"{col}16"].font = font_regular
        ws[f"{col}16"].number_format = "#,##0"
        ws[f"{col}16"].alignment = align_right
        
    # Gross Profit (Row 17)
    ws["B17"] = "Gross profit"
    ws["B17"].font = font_bold
    for col in cols:
        ws[f"{col}17"] = f"={col}13+{col}16"
        ws[f"{col}17"].font = font_bold
        ws[f"{col}17"].number_format = "#,##0"
        ws[f"{col}17"].alignment = align_right
        ws[f"{col}17"].border = border_top_thin
        
    # Operating Expenses Header (Row 19)
    ws["B19"] = "Operating expenses"
    ws["B19"].font = Font(name=font_family, size=10, italic=True, bold=True)
    
    # Operation & support (Row 20)
    ws["B20"] = "  Operation & support"
    ws["B20"].font = font_regular
    op_support = [1877, 2413, 2689, 2732, 2854]
    for col, val in zip(cols[:5], op_support):
        ws[f"{col}20"] = val
        ws[f"{col}20"].font = font_blue_input
        ws[f"{col}20"].number_format = "#,##0"
        ws[f"{col}20"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}20"] = f"={col}13*{col}49"
        ws[f"{col}20"].font = font_regular
        ws[f"{col}20"].number_format = "#,##0"
        ws[f"{col}20"].alignment = align_right
        
    # Sales & marketing (Row 21)
    ws["B21"] = "  Sales & marketing"
    ws["B21"].font = font_regular
    sales_mkt = [4789, 4756, 4356, 4337, 4898]
    for col, val in zip(cols[:5], sales_mkt):
        ws[f"{col}21"] = val
        ws[f"{col}21"].font = font_blue_input
        ws[f"{col}21"].number_format = "#,##0"
        ws[f"{col}21"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}21"] = f"={col}13*{col}50"
        ws[f"{col}21"].font = font_regular
        ws[f"{col}21"].number_format = "#,##0"
        ws[f"{col}21"].alignment = align_right
        
    # R&D (Row 22)
    ws["B22"] = "  R&D"
    ws["B22"].font = font_regular
    rd = [2054, 2798, 3164, 3109, 3402]
    for col, val in zip(cols[:5], rd):
        ws[f"{col}22"] = val
        ws[f"{col}22"].font = font_blue_input
        ws[f"{col}22"].number_format = "#,##0"
        ws[f"{col}22"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}22"] = f"={col}13*{col}51"
        ws[f"{col}22"].font = font_regular
        ws[f"{col}22"].number_format = "#,##0"
        ws[f"{col}22"].alignment = align_right
        
    # G&A (Row 23)
    ws["B23"] = "  G&A"
    ws["B23"].font = font_regular
    ga = [2316, 3136, 2682, 3639, 3241]
    for col, val in zip(cols[:5], ga):
        ws[f"{col}23"] = val
        ws[f"{col}23"].font = font_blue_input
        ws[f"{col}23"].number_format = "#,##0"
        ws[f"{col}23"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}23"] = f"={col}13*{col}52"
        ws[f"{col}23"].font = font_regular
        ws[f"{col}23"].number_format = "#,##0"
        ws[f"{col}23"].alignment = align_right
        
    # D&A (Row 24)
    ws["B24"] = "  D&A"
    ws["B24"].font = font_regular
    da = [902, 947, 823, 711, 719]
    for col, val in zip(cols[:5], da):
        ws[f"{col}24"] = val
        ws[f"{col}24"].font = font_blue_input
        ws[f"{col}24"].number_format = "#,##0"
        ws[f"{col}24"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}24"] = f"={col}13*{col}53"
        ws[f"{col}24"].font = font_regular
        ws[f"{col}24"].number_format = "#,##0"
        ws[f"{col}24"].alignment = align_right
        
    # Total operating expenses (Row 25)
    ws["B25"] = "Total operating expenses"
    ws["B25"].font = font_regular
    for col in cols:
        ws[f"{col}25"] = f"=SUM({col}20:{col}24)"
        ws[f"{col}25"].font = font_regular
        ws[f"{col}25"].number_format = "#,##0"
        ws[f"{col}25"].alignment = align_right
        ws[f"{col}25"].border = border_top_thin
        
    # Operating Income (loss) (Row 26)
    ws["B26"] = "Operating Income (loss)"
    ws["B26"].font = font_bold
    for col in cols:
        ws[f"{col}26"] = f"={col}17-{col}25"
        ws[f"{col}26"].font = font_bold
        ws[f"{col}26"].number_format = "#,##0"
        ws[f"{col}26"].alignment = align_right
        ws[f"{col}26"].border = border_total
        
    # Interest income (expense) (Row 28)
    ws["B28"] = "Interest income (expense)"
    ws["B28"].font = font_regular
    interest = [-483, -565, -633, -523, 303]
    for col, val in zip(cols[:5], interest):
        ws[f"{col}28"] = val
        ws[f"{col}28"].font = font_blue_input
        ws[f"{col}28"].number_format = "#,##0"
        ws[f"{col}28"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}28"] = 303 # Assumed constant
        ws[f"{col}28"].font = font_regular
        ws[f"{col}28"].number_format = "#,##0"
        ws[f"{col}28"].alignment = align_right
        
    # Other Income (loss) (Row 29)
    ws["B29"] = "Other Income (loss)"
    ws["B29"].font = font_regular
    other_inc = [3292, -7029, 1844, 1849, -68]
    for col, val in zip(cols[:5], other_inc):
        ws[f"{col}29"] = val
        ws[f"{col}29"].font = font_blue_input
        ws[f"{col}29"].number_format = "#,##0"
        ws[f"{col}29"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}29"] = 0 # Assumed constant
        ws[f"{col}29"].font = font_regular
        ws[f"{col}29"].number_format = "#,##0"
        ws[f"{col}29"].alignment = align_right
        
    # Income (loss) before taxes & equity (Row 31)
    ws["B31"] = "Income (loss) before taxes & equity"
    ws["B31"].font = font_bold
    for col in cols:
        ws[f"{col}31"] = f"={col}26+{col}28+{col}29"
        ws[f"{col}31"].font = font_bold
        ws[f"{col}31"].number_format = "#,##0"
        ws[f"{col}31"].alignment = align_right
        ws[f"{col}31"].border = border_top_thin
        
    # tax provision (Row 32)
    ws["B32"] = "provision for (benefit from) income tax"
    ws["B32"].font = font_regular
    tax = [-492, -181, 213, -5758, -4346]
    for col, val in zip(cols[:5], tax):
        ws[f"{col}32"] = val
        ws[f"{col}32"].font = font_blue_input
        ws[f"{col}32"].number_format = "#,##0"
        ws[f"{col}32"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}32"] = f"={col}31*{col}55"
        ws[f"{col}32"].font = font_regular
        ws[f"{col}32"].number_format = "#,##0"
        ws[f"{col}32"].alignment = align_right
        
    # equity inv (Row 33)
    ws["B33"] = "Income (loss) from equity inv."
    ws["B33"].font = font_regular
    equity = [-37, 107, 48, -38, -53]
    for col, val in zip(cols[:5], equity):
        ws[f"{col}33"] = val
        ws[f"{col}33"].font = font_blue_input
        ws[f"{col}33"].number_format = "#,##0"
        ws[f"{col}33"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}33"] = 0
        ws[f"{col}33"].font = font_regular
        ws[f"{col}33"].number_format = "#,##0"
        ws[f"{col}33"].alignment = align_right
        
    # Net income (loss) incl. NCI (Row 34)
    ws["B34"] = "Net income (loss) incl. NCI"
    ws["B34"].font = font_bold
    for col in cols:
        ws[f"{col}34"] = f"={col}31-{col}32+{col}33"
        ws[f"{col}34"].font = font_bold
        ws[f"{col}34"].number_format = "#,##0"
        ws[f"{col}34"].alignment = align_right
        ws[f"{col}34"].border = border_top_thin
        
    # Less NCI (Row 35)
    ws["B35"] = "Less: net income (loss) attributable to NCI"
    ws["B35"].font = font_regular
    nci = [-74, 3, 269, -11, 40]
    for col, val in zip(cols[:5], nci):
        ws[f"{col}35"] = val
        ws[f"{col}35"].font = font_blue_input
        ws[f"{col}35"].number_format = "#,##0"
        ws[f"{col}35"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}35"] = 0
        ws[f"{col}35"].font = font_regular
        ws[f"{col}35"].number_format = "#,##0"
        ws[f"{col}35"].alignment = align_right
        
    # Net income (loss) (Row 36)
    ws["B36"] = "Net income (loss)"
    ws["B36"].font = font_bold
    for col in cols:
        ws[f"{col}36"] = f"={col}34-{col}35"
        ws[f"{col}36"].font = font_bold
        ws[f"{col}36"].number_format = "#,##0"
        ws[f"{col}36"].alignment = align_right
        ws[f"{col}36"].border = border_total
        
    # Shares (Row 38)
    ws["B38"] = "Diluted shares outstanding"
    ws["B38"].font = font_regular
    shares = [1896, 1975, 2092, 2151, 2120]
    for col, val in zip(cols[:5], shares):
        ws[f"{col}38"] = val
        ws[f"{col}38"].font = font_blue_input
        ws[f"{col}38"].number_format = "#,##0"
        ws[f"{col}38"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}38"] = f"={cols[4]}38" # Keep constant at latest
        ws[f"{col}38"].font = font_bold
        ws[f"{col}38"].number_format = "#,##0"
        ws[f"{col}38"].alignment = align_right
        
    # Share growth (Row 39)
    ws["B39"] = "% growth"
    ws["B39"].font = font_gray
    for i in range(1, 10):
        col = cols[i]
        prev_col = cols[i-1]
        ws[f"{col}39"] = f"=({col}38-{prev_col}38)/{prev_col}38"
        ws[f"{col}39"].font = font_gray
        ws[f"{col}39"].number_format = "0.0%"
        ws[f"{col}39"].alignment = align_right
        
    # Diluted EPS (Row 40)
    ws["B40"] = "Diluted EPS"
    ws["B40"].font = font_bold
    for col in cols:
        ws[f"{col}40"] = f"={col}36/{col}38"
        ws[f"{col}40"].font = font_bold
        ws[f"{col}40"].number_format = "$#,##0.00"
        ws[f"{col}40"].alignment = align_right
        ws[f"{col}40"].border = border_total
        
    # D&A (Row 42)
    ws["B42"] = "Depreciation & amortization"
    ws["B42"].font = font_regular
    for col in cols:
        ws[f"{col}42"] = f"={col}24"
        ws[f"{col}42"].font = font_regular
        ws[f"{col}42"].number_format = "#,##0"
        ws[f"{col}42"].alignment = align_right
        
    # EBITDA (Row 43)
    ws["B43"] = "EBITDA"
    ws["B43"].font = font_bold
    for col in cols:
        ws[f"{col}43"] = f"={col}26+{col}42"
        ws[f"{col}43"].font = font_bold
        ws[f"{col}43"].number_format = "#,##0"
        ws[f"{col}43"].alignment = align_right
        ws[f"{col}43"].border = border_top_thin
        
    # SBC (Row 44)
    ws["B44"] = "Stock-based compensation"
    ws["B44"].font = font_regular
    sbc_vals = [1168, 1793, 1935, 1796, 1826]
    for col, val in zip(cols[:5], sbc_vals):
        ws[f"{col}44"] = val
        ws[f"{col}44"].font = font_blue_input
        ws[f"{col}44"].number_format = "#,##0"
        ws[f"{col}44"].alignment = align_right
    for i in range(5, 10):
        col = cols[i]
        ws[f"{col}44"] = f"={col}13*{col}56"
        ws[f"{col}44"].font = font_regular
        ws[f"{col}44"].number_format = "#,##0"
        ws[f"{col}44"].alignment = align_right
        
    # Adjusted EBITDA (Row 45)
    ws["B45"] = "Adjusted EBITDA"
    ws["B45"].font = font_bold
    for col in cols:
        ws[f"{col}45"] = f"={col}43+{col}44"
        ws[f"{col}45"].font = font_bold
        ws[f"{col}45"].number_format = "#,##0"
        ws[f"{col}45"].alignment = align_right
        ws[f"{col}45"].border = border_total
        
    # ----------------------------------------------------
    # Assumptions & Margins (Rows 47-56)
    # ----------------------------------------------------
    # Rev growth (Row 47)
    ws["B47"] = "Revenue growth"
    ws["B47"].font = font_bold
    for i in range(1, 5):
        col = cols[i]
        ws[f"{col}47"] = f"={col}14"
        ws[f"{col}47"].font = font_gray
        ws[f"{col}47"].number_format = "0.0%"
        ws[f"{col}47"].alignment = align_right
    # Projected growth assumptions (Inputs)
    proj_growths = [0.18, 0.17, 0.17, 0.16, 0.15]
    for col, val in zip(cols[5:], proj_growths):
        ws[f"{col}47"] = val
        ws[f"{col}47"].font = font_blue_input
        ws[f"{col}47"].number_format = "0.0%"
        ws[f"{col}47"].alignment = align_right
        
    # Gross margin (Row 48)
    ws["B48"] = "Gross margin"
    ws["B48"].font = font_regular
    for col in cols[:5]:
        ws[f"{col}48"] = f"={col}17/{col}13"
        ws[f"{col}48"].font = font_gray
        ws[f"{col}48"].number_format = "0.0%"
        ws[f"{col}48"].alignment = align_right
    # Projected COGS % of sales (Input)
    proj_cogs_pcts = [-0.60, -0.60, -0.60, -0.60, -0.60]
    for col, val in zip(cols[5:], proj_cogs_pcts):
        ws[f"{col}48"] = val
        ws[f"{col}48"].font = font_blue_input
        ws[f"{col}48"].number_format = "0.0%"
        ws[f"{col}48"].alignment = align_right
        
    # Operation & support % of sales (Row 49)
    ws["B49"] = "Operation & support % of sales"
    ws["B49"].font = font_regular
    for col in cols[:5]:
        ws[f"{col}49"] = f"={col}20/{col}13"
        ws[f"{col}49"].font = font_gray
        ws[f"{col}49"].number_format = "0.0%"
        ws[f"{col}49"].alignment = align_right
    proj_op_supp = [0.05, 0.05, 0.05, 0.04, 0.04]
    for col, val in zip(cols[5:], proj_op_supp):
        ws[f"{col}49"] = val
        ws[f"{col}49"].font = font_blue_input
        ws[f"{col}49"].number_format = "0.0%"
        ws[f"{col}49"].alignment = align_right
        
    # Sales & marketing % of sales (Row 50)
    ws["B50"] = "Sales & marketing % of sales"
    ws["B50"].font = font_regular
    for col in cols[:5]:
        ws[f"{col}50"] = f"={col}21/{col}13"
        ws[f"{col}50"].font = font_gray
        ws[f"{col}50"].number_format = "0.0%"
        ws[f"{col}50"].alignment = align_right
    proj_sm = [0.09, 0.09, 0.08, 0.07, 0.07]
    for col, val in zip(cols[5:], proj_sm):
        ws[f"{col}50"] = val
        ws[f"{col}50"].font = font_blue_input
        ws[f"{col}50"].number_format = "0.0%"
        ws[f"{col}50"].alignment = align_right
        
    # R&D % of sales (Row 51)
    ws["B51"] = "R&D % of sales"
    ws["B51"].font = font_regular
    for col in cols[:5]:
        ws[f"{col}51"] = f"={col}22/{col}13"
        ws[f"{col}51"].font = font_gray
        ws[f"{col}51"].number_format = "0.0%"
        ws[f"{col}51"].alignment = align_right
    proj_rd = [0.07, 0.07, 0.06, 0.06, 0.05]
    for col, val in zip(cols[5:], proj_rd):
        ws[f"{col}51"] = val
        ws[f"{col}51"].font = font_blue_input
        ws[f"{col}51"].number_format = "0.0%"
        ws[f"{col}51"].alignment = align_right
        
    # G&A % of sales (Row 52)
    ws["B52"] = "G&A % of sales"
    ws["B52"].font = font_regular
    for col in cols[:5]:
        ws[f"{col}52"] = f"={col}23/{col}13"
        ws[f"{col}52"].font = font_gray
        ws[f"{col}52"].number_format = "0.0%"
        ws[f"{col}52"].alignment = align_right
    proj_ga = [0.06, 0.05, 0.05, 0.05, 0.04]
    for col, val in zip(cols[5:], proj_ga):
        ws[f"{col}52"] = val
        ws[f"{col}52"].font = font_blue_input
        ws[f"{col}52"].number_format = "0.0%"
        ws[f"{col}52"].alignment = align_right
        
    # D&A % of sales (Row 53)
    ws["B53"] = "D&A % of sales"
    ws["B53"].font = font_regular
    for col in cols[:5]:
        ws[f"{col}53"] = f"={col}24/{col}13"
        ws[f"{col}53"].font = font_gray
        ws[f"{col}53"].number_format = "0.0%"
        ws[f"{col}53"].alignment = align_right
    proj_da = [0.01, 0.01, 0.01, 0.01, 0.01]
    for col, val in zip(cols[5:], proj_da):
        ws[f"{col}53"] = val
        ws[f"{col}53"].font = font_blue_input
        ws[f"{col}53"].number_format = "0.0%"
        ws[f"{col}53"].alignment = align_right
        
    # Operating margin (Row 54)
    ws["B54"] = "Operating margin"
    ws["B54"].font = font_bold
    for col in cols:
        ws[f"{col}54"] = f"={col}26/{col}13"
        ws[f"{col}54"].font = font_bold
        ws[f"{col}54"].number_format = "0.0%"
        ws[f"{col}54"].alignment = align_right
        
    # Tax rate (Row 55)
    ws["B55"] = "Tax rate"
    ws["B55"].font = font_regular
    for col in cols[:5]:
        ws[f"{col}55"] = f"=IF({col}31<=0, \"NA\", {col}32/{col}31)"
        ws[f"{col}55"].font = font_gray
        ws[f"{col}55"].number_format = "0.0%"
        ws[f"{col}55"].alignment = align_right
    # Projected tax rate (Input)
    for col in cols[5:]:
        ws[f"{col}55"] = 0.21
        ws[f"{col}55"].font = font_blue_input
        ws[f"{col}55"].number_format = "0.0%"
        ws[f"{col}55"].alignment = align_right
        
    # SBC % of sales (Row 56)
    ws["B56"] = "SBC % of sales"
    ws["B56"].font = font_regular
    for col in cols[:5]:
        ws[f"{col}56"] = f"={col}44/{col}13"
        ws[f"{col}56"].font = font_gray
        ws[f"{col}56"].number_format = "0.0%"
        ws[f"{col}56"].alignment = align_right
    proj_sbc = [0.03, 0.03, 0.03, 0.02, 0.02]
    for col, val in zip(cols[5:], proj_sbc):
        ws[f"{col}56"] = val
        ws[f"{col}56"].font = font_blue_input
        ws[f"{col}56"].number_format = "0.0%"
        ws[f"{col}56"].alignment = align_right
        
    # ----------------------------------------------------
    # Balance Sheet Headers (Row 58)
    # ----------------------------------------------------
    format_sec_row(58)
    ws.merge_cells("B58:D58")
    ws["B58"] = "Balance Sheet"
    ws["B58"].alignment = align_left
    
    for col, yr in zip(cols, years):
        cell_ref = f"{col}58"
        suffix = "A" if yr <= 2025 else "E"
        ws[cell_ref] = f"{yr}{suffix}"
        ws[cell_ref].alignment = align_right
        
    # Balance Sheet Items
    ws["B60"] = "Assets"
    ws["B60"].font = Font(name=font_family, size=10, bold=True, italic=True)
    
    # Let's populate the row labels and default historical values for balance sheet
    bs_items = [
        ("Cash & cash equivalent", "CashEquivalents", [4295, 4208, 4680, 5893, 7105]),
        ("Short-term investments", "ShortTermInvestments", [0, 103, 727, 1084, 528]),
        ("Restricted cash & cash equivalent", "RestrictedCashCurrent", [631, 680, 805, 545, 631]),
        ("Accounts receivable", "AccountsReceivable", [2439, 2779, 3404, 3333, 3827]),
        ("Prepaid expenses & other current assets", "PrepaidExpenses", [1454, 1479, 1681, 1390, 1902]),
        ("Total current assets", "TotalCurrentAssets", None), # Formula: sum of above
        ("Restricted cash & cash equivalent", "RestrictedCashNoncurrent", [2879, 1789, 1519, 2172, 1911]),
        ("Restricted investments", "RestrictedInvestments", [0, 1614, 4779, 7019, 8874]),
        ("Investments", "Investments", [11806, 4401, 6101, 8460, 9178]),
        ("Equity method investments", "EquityMethodInvestments", [800, 870, 353, 302, 287]),
        ("PP&E", "PP_E", [1853, 2082, 2073, 1952, 1897]),
        ("Operating lease right-of-use assets", "LeaseAssets", [1388, 1449, 1241, 1158, 1114]),
        ("Intangible assets", "Intangibles", [2412, 1874, 1425, 1125, 1048]),
        ("Goodwill", "Goodwill", [8420, 8263, 8151, 8066, 8931]),
        ("Other assets", "OtherAssets", [397, 518, 1760, 8745, 14569]),
        ("Total assets", "TotalAssets", None) # Formula: Total current + rest
    ]
    
    current_row = 61
    for label, key, vals in bs_items:
        ws.cell(row=current_row, column=2, value="  " + label if vals is not None or "Total" not in label else label)
        ws.cell(row=current_row, column=2).font = font_bold if "Total" in label else font_regular
        
        if vals is not None:
            # write historical
            for col, val in zip(cols[:5], vals):
                ws[f"{col}{current_row}"] = val
                ws[f"{col}{current_row}"].font = font_blue_input
                ws[f"{col}{current_row}"].number_format = "#,##0"
                ws[f"{col}{current_row}"].alignment = align_right
            # write projection
            for col in cols[5:]:
                ws[f"{col}{current_row}"] = f"={cols[4]}{current_row}"
                ws[f"{col}{current_row}"].font = font_regular
                ws[f"{col}{current_row}"].number_format = "#,##0"
                ws[f"{col}{current_row}"].alignment = align_right
        else:
            # Formula rows
            if label == "Total current assets":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"=SUM({col}61:{col}65)"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
            elif label == "Total assets":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"={col}66+SUM({col}67:{col}75)"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
                    
        current_row += 1
        
    # Liabilities & Equity (starts at Row 78)
    ws["B78"] = "Liabilities"
    ws["B78"].font = Font(name=font_family, size=10, bold=True, italic=True)
    
    liab_items = [
        ("Accounts payable", "AccountsPayable", [860, 728, 790, 858, 1013]),
        ("Short-term debt & current portion of LT debt", "ShortTermDebt", [0, 0, 0, 0, 0]),
        ("Operating lease liabilities, current", "LeaseLiabilitiesCurrent", [185, 201, 190, 175, 169]),
        ("Accrued and other current liabilities", "AccruedLiabilitiesCurrent", [6537, 6232, 6397, 7689, 7751]),
        ("Total current liabilities", "TotalCurrentLiabilities", None),
        ("Long-term insurance reserves", "LongTermInsuranceReserves", [2546, 3028, 4909, 7042, 9076]),
        ("Long-term debt, net of current portion", "LongTermDebt", [9276, 9265, 9459, 8347, 10521]),
        ("Operating lease liabilities, non-current", "LeaseLiabilitiesNoncurrent", [1644, 1673, 1550, 1454, 1390]),
        ("Other long-term liabilities", "OtherLiabilitiesNoncurrent", [935, 786, 645, 449, 412]),
        ("Total liabilities", "TotalLiabilities", None),
        ("Redeemable non-controlling interests", "RedeemableNCI", [204, 430, 654, 93, 165])
    ]
    
    current_row = 79
    for label, key, vals in liab_items:
        ws.cell(row=current_row, column=2, value="  " + label if vals is not None or "Total" not in label else label)
        ws.cell(row=current_row, column=2).font = font_bold if "Total" in label else font_regular
        
        if vals is not None:
            # write historical
            for col, val in zip(cols[:5], vals):
                ws[f"{col}{current_row}"] = val
                ws[f"{col}{current_row}"].font = font_blue_input
                ws[f"{col}{current_row}"].number_format = "#,##0"
                ws[f"{col}{current_row}"].alignment = align_right
            # write projection
            for col in cols[5:]:
                ws[f"{col}{current_row}"] = f"={cols[4]}{current_row}"
                ws[f"{col}{current_row}"].font = font_regular
                ws[f"{col}{current_row}"].number_format = "#,##0"
                ws[f"{col}{current_row}"].alignment = align_right
        else:
            # Formula rows
            if label == "Total current liabilities":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"=SUM({col}79:{col}82)"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
            elif label == "Total liabilities":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"={col}83+SUM({col}84:{col}87)"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
                    
        current_row += 1
        
    # Equity Section (starts at Row 91)
    ws["B91"] = "Equity"
    ws["B91"].font = Font(name=font_family, size=10, bold=True, italic=True)
    
    eq_items = [
        ("Additional paid-in capital", "APIC", [38608, 40550, 42264, 42801, 38101]),
        ("Accumulated other comprehensive loss", "AccumOCI", [-524, -443, -421, -517, -432]),
        ("Accumulated deficit", "AccumDeficit", [-23626, -32767, -30594, -20726, -10628]),
        ("Total stockholders' equity", "ParentStockholdersEquity", None),
        ("Non-redeemable non-controlling interests", "NoncontrollingInterest", [687, 734, 779, 825, 877]),
        ("Total equity", "TotalEquity", None),
        ("Total liabilities, redeemable NCI & equity", "TotalLiabEquity", None)
    ]
    
    current_row = 92
    for label, key, vals in eq_items:
        ws.cell(row=current_row, column=2, value="  " + label if vals is not None or "Total" not in label else label)
        ws.cell(row=current_row, column=2).font = font_bold if "Total" in label else font_regular
        
        if vals is not None:
            # write historical
            for col, val in zip(cols[:5], vals):
                ws[f"{col}{current_row}"] = val
                ws[f"{col}{current_row}"].font = font_blue_input
                ws[f"{col}{current_row}"].number_format = "#,##0"
                ws[f"{col}{current_row}"].alignment = align_right
            # write projection
            for col in cols[5:]:
                ws[f"{col}{current_row}"] = f"={cols[4]}{current_row}"
                ws[f"{col}{current_row}"].font = font_regular
                ws[f"{col}{current_row}"].number_format = "#,##0"
                ws[f"{col}{current_row}"].alignment = align_right
        else:
            # Formula rows
            if label == "Total stockholders' equity":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"=SUM({col}92:{col}94)"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
            elif label == "Total equity":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"={col}95+{col}96"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
            elif label == "Total liabilities, redeemable NCI & equity":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"={col}88+{col}89+{col}97"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
                    
        current_row += 1
        
    # ----------------------------------------------------
    # Cash Flow Statement Headers (Row 100)
    # ----------------------------------------------------
    format_sec_row(100)
    ws.merge_cells("B100:D100")
    ws["B100"] = "Cash Flow Statements"
    ws["B100"].alignment = align_left
    
    for col, yr in zip(cols, years):
        cell_ref = f"{col}100"
        suffix = "A" if yr <= 2025 else "E"
        ws[cell_ref] = f"{yr}{suffix}"
        ws[cell_ref].alignment = align_right
        
    # Cash Flow Items
    cf_items = [
        ("Net income", "CF_NetIncome", "=E36"), # Link to Income Statement
        ("Depreciation and amortization", "CF_Depreciation", "=E24"), # Link to Income Statement
        ("Stock based compensation", "CF_SBC", "=E44"), # Link to Income Statement
        ("Decrease / (increase) in working capital assets", "CF_WorkingCapitalAssets", [-343, -196, -214, -305, -450]), # Sample historical
        ("Decrease / (increase) in working capital liabilities", "CF_WorkingCapitalLiabilities", [520, 240, 185, 290, 480]), # Sample historical
        ("Other non current assets", "CF_OtherNonCurrentAssets", [100, 50, -30, -80, -95]),
        ("Other non current liabilities", "CF_OtherNonCurrentLiabilities", [150, -80, 120, -110, 45]),
        ("Cash from operating activities", "CF_TotalOperating", None), # Formula
        ("Capital expenditures", "CF_Capex", [-252, -223, -242, -336, -350]), # Sample historical
        ("Cash from investing activities", "CF_TotalInvesting", None), # Formula
        ("Long term debt", "CF_LongTermDebt", [500, -200, 450, -800, 1200]),
        ("Commercial paper / Revolver", "CF_CommercialPaper", [0, 0, 0, 0, 0]),
        ("Shares repurchases", "CF_ShareRepurchases", [0, 0, 0, -1252, -6523]),
        ("Common dividends", "CF_CommonDividends", [0, 0, 0, 0, 0]),
        ("Cash from financing activities", "CF_TotalFinancing", None), # Formula
        ("Net change in cash during period", "CF_NetChange", None) # Formula
    ]
    
    current_row = 102
    for label, key, val_ref in cf_items:
        ws.cell(row=current_row, column=2, value="  " + label if "Total" not in label and "Cash from" not in label and "Net change" not in label else label)
        ws.cell(row=current_row, column=2).font = font_bold if "Cash from" in label or "Net change" in label else font_regular
        
        if val_ref is not None and not isinstance(val_ref, list) and val_ref.startswith("="):
            # Formula/Link for all columns
            for col in cols:
                # Adjust cell column letter in the formula string
                formula_adjusted = val_ref.replace("E", col)
                ws[f"{col}{current_row}"] = formula_adjusted
                ws[f"{col}{current_row}"].font = font_bold if "Total" in label else font_regular
                ws[f"{col}{current_row}"].number_format = "#,##0"
                ws[f"{col}{current_row}"].alignment = align_right
        elif isinstance(val_ref, list):
            # Historical list
            for col, val in zip(cols[:5], val_ref):
                ws[f"{col}{current_row}"] = val
                ws[f"{col}{current_row}"].font = font_blue_input
                ws[f"{col}{current_row}"].number_format = "#,##0"
                ws[f"{col}{current_row}"].alignment = align_right
            # Projection
            for col in cols[5:]:
                ws[f"{col}{current_row}"] = f"={cols[4]}{current_row}" # Keep constant
                ws[f"{col}{current_row}"].font = font_regular
                ws[f"{col}{current_row}"].number_format = "#,##0"
                ws[f"{col}{current_row}"].alignment = align_right
        else:
            # Totals
            if label == "Cash from operating activities":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"=SUM({col}102:{col}108)"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
            elif label == "Cash from investing activities":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"={col}110" # Only Capex for simplicity
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
            elif label == "Cash from financing activities":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"=SUM({col}112:{col}115)"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
            elif label == "Net change in cash during period":
                for col in cols:
                    ws[f"{col}{current_row}"] = f"={col}109+{col}111+{col}116"
                    ws[f"{col}{current_row}"].font = font_bold
                    ws[f"{col}{current_row}"].number_format = "#,##0"
                    ws[f"{col}{current_row}"].alignment = align_right
                    ws[f"{col}{current_row}"].border = border_total
                    
        current_row += 1
        
    # ----------------------------------------------------
    # DCF Valuation Section (starts at Row 122)
    # ----------------------------------------------------
    format_sec_row(122)
    ws.merge_cells("B122:D122")
    ws["B122"] = "Discounted Cash Flow (DCF) Valuation Model"
    ws["B122"].alignment = align_left
    
    # DCF Headers
    for col, yr in zip(cols[5:], years[5:]):
        cell_ref = f"{col}123"
        ws[cell_ref] = f"{yr}E"
        ws[cell_ref].font = font_sub_header
        ws[cell_ref].alignment = align_right
        
    # DCF Rows
    ws["B125"] = "Operating Income (EBIT)"
    ws["B125"].font = font_regular
    for col in cols[5:]:
        ws[f"{col}125"] = f"={col}26" # link to Operating Income
        ws[f"{col}125"].font = font_regular
        ws[f"{col}125"].number_format = "#,##0"
        ws[f"{col}125"].alignment = align_right
        
    ws["B126"] = "Tax Rate"
    ws["B126"].font = font_regular
    for col in cols[5:]:
        ws[f"{col}126"] = f"={col}55" # link to projected tax rate
        ws[f"{col}126"].font = font_regular
        ws[f"{col}126"].number_format = "0.0%"
        ws[f"{col}126"].alignment = align_right
        
    ws["B127"] = "Taxes on EBIT"
    ws["B127"].font = font_regular
    for col in cols[5:]:
        ws[f"{col}127"] = f"={col}125*{col}126"
        ws[f"{col}127"].font = font_regular
        ws[f"{col}127"].number_format = "#,##0"
        ws[f"{col}127"].alignment = align_right
        
    ws["B128"] = "NOPAT (Net Operating Profit After Tax)"
    ws["B128"].font = font_bold
    for col in cols[5:]:
        ws[f"{col}128"] = f"={col}125-{col}127"
        ws[f"{col}128"].font = font_bold
        ws[f"{col}128"].number_format = "#,##0"
        ws[f"{col}128"].alignment = align_right
        ws[f"{col}128"].border = border_top_thin
        
    ws["B129"] = "Add: Depreciation & Amortization (D&A)"
    ws["B129"].font = font_regular
    for col in cols[5:]:
        ws[f"{col}129"] = f"={col}24" # link to D&A (Row 24)
        ws[f"{col}129"].font = font_regular
        ws[f"{col}129"].number_format = "#,##0"
        ws[f"{col}129"].alignment = align_right
        
    ws["B130"] = "Less: Capital Expenditures (Capex)"
    ws["B130"].font = font_regular
    for col in cols[5:]:
        ws[f"{col}130"] = f"={col}110" # link to Capex (stored as negative cash outflow in Row 110)
        ws[f"{col}130"].font = font_regular
        ws[f"{col}130"].number_format = "#,##0"
        ws[f"{col}130"].alignment = align_right
        
    ws["B131"] = "Less: Change in Net Working Capital"
    ws["B131"].font = font_regular
    for col in cols[5:]:
        ws[f"{col}131"] = f"={col}105+{col}106" # sum of working capital asset & liab changes
        ws[f"{col}131"].font = font_regular
        ws[f"{col}131"].number_format = "#,##0"
        ws[f"{col}131"].alignment = align_right
        
    ws["B132"] = "Free Cash Flow to Firm (FCFF)"
    ws["B132"].font = font_bold
    for col in cols[5:]:
        ws[f"{col}132"] = f"={col}128+{col}129+{col}130+{col}131"
        ws[f"{col}132"].font = font_bold
        ws[f"{col}132"].number_format = "#,##0"
        ws[f"{col}132"].alignment = align_right
        ws[f"{col}132"].border = border_total
        
    ws["B134"] = "Discount Period"
    ws["B134"].font = font_gray
    for idx, col in enumerate(cols[5:]):
        ws[f"{col}134"] = idx + 1
        ws[f"{col}134"].font = font_gray
        ws[f"{col}134"].alignment = align_right
        
    ws["B135"] = "Discount Factor"
    ws["B135"].font = font_gray
    for col in cols[5:]:
        ws[f"{col}135"] = f"=1/((1+$L$7)^{col}134)" # $L$7 is WACC
        ws[f"{col}135"].font = font_gray
        ws[f"{col}135"].number_format = "0.0000"
        ws[f"{col}135"].alignment = align_right
        
    ws["B136"] = "Present Value of FCF"
    ws["B136"].font = font_bold
    for col in cols[5:]:
        ws[f"{col}136"] = f"={col}132*{col}135"
        ws[f"{col}136"].font = font_bold
        ws[f"{col}136"].number_format = "#,##0"
        ws[f"{col}136"].alignment = align_right
        ws[f"{col}136"].border = border_total
        
    # ----------------------------------------------------
    # Valuation Output Block (placed in Column B and C, Rows 139-148)
    # ----------------------------------------------------
    ws["B139"] = "Valuation Summary"
    ws["B139"].font = font_sub_header
    ws["B139"].fill = fill_sub_header
    ws["B139"].alignment = align_left
    
    ws["B140"] = "Cumulative PV of FCFs"
    ws["B140"].font = font_regular
    ws["C140"] = "=SUM(J136:N136)"
    ws["C140"].font = font_bold
    ws["C140"].number_format = "$#,##0"
    ws["C140"].alignment = align_right
    ws["C140"].border = border_all_thin
    
    ws["B141"] = "Terminal Value"
    ws["B141"].font = font_regular
    ws["C141"] = "=N132*(1+$L$8)/($L$7-$L$8)" # $L$8 is terminal growth rate, $L$7 is WACC
    ws["C141"].font = font_regular
    ws["C141"].number_format = "$#,##0"
    ws["C141"].alignment = align_right
    ws["C141"].border = border_all_thin
    
    ws["B142"] = "PV of Terminal Value"
    ws["B142"].font = font_regular
    ws["C142"] = "=C141*N135" # N135 is discount factor for year 5
    ws["C142"].font = font_bold
    ws["C142"].number_format = "$#,##0"
    ws["C142"].alignment = align_right
    ws["C142"].border = border_all_thin
    
    ws["B143"] = "Implied Enterprise Value"
    ws["B143"].font = font_bold
    ws["C143"] = "=C140+C142"
    ws["C143"].font = font_bold
    ws["C143"].number_format = "$#,##0"
    ws["C143"].alignment = align_right
    ws["C143"].border = border_all_thin
    
    ws["B144"] = "Less: Net Debt"
    ws["B144"].font = font_regular
    ws["C144"] = "=I8-I7" # Debt (I8) - Cash (I7)
    ws["C144"].font = font_regular
    ws["C144"].number_format = "$#,##0"
    ws["C144"].alignment = align_right
    ws["C144"].border = border_all_thin
    
    ws["B145"] = "Implied Equity Value"
    ws["B145"].font = font_bold
    ws["C145"] = "=C143-C144"
    ws["C145"].font = font_bold
    ws["C145"].number_format = "$#,##0"
    ws["C145"].alignment = align_right
    ws["C145"].border = border_all_thin
    
    ws["B146"] = "Diluted Shares Outstanding"
    ws["B146"].font = font_regular
    ws["C146"] = "=F8" # Links to Shares Outstanding in Overview
    ws["C146"].font = font_regular
    ws["C146"].number_format = "#,##0"
    ws["C146"].alignment = align_right
    ws["C146"].border = border_all_thin
    
    ws["B147"] = "Implied Share Price"
    ws["B147"].font = Font(name=font_family, size=11, bold=True, color="008000") # Green for target price!
    ws["C147"] = "=C145/C146"
    ws["C147"].font = Font(name=font_family, size=11, bold=True, color="008000")
    ws["C147"].number_format = "$#,##0.00"
    ws["C147"].alignment = align_right
    ws["C147"].border = Border(left=thin_line, right=thin_line, top=thin_line, bottom=double_line)
    
    ws["B148"] = "Current Share Price"
    ws["B148"].font = font_regular
    ws["C148"] = "=F7"
    ws["C148"].font = font_bold
    ws["C148"].number_format = "$#,##0.00"
    ws["C148"].alignment = align_right
    ws["C148"].border = border_all_thin
    
    ws["B149"] = "Implied Upside / (Downside)"
    ws["B149"].font = font_bold
    ws["C149"] = "=(C147-C148)/C148"
    ws["C149"].font = font_bold
    ws["C149"].number_format = "0.0%"
    ws["C149"].alignment = align_right
    ws["C149"].border = border_all_thin
    
    # ----------------------------------------------------
    # Adjust Column Widths
    # ----------------------------------------------------
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 45 # Label column needs to be wide
    ws.column_dimensions["C"].width = 25 # Company name / value column
    ws.column_dimensions["D"].width = 3
    
    for col_letter in cols:
        ws.column_dimensions[col_letter].width = 15
        
    wb.save("DCF_Model_Template.xlsx")
    print("DCF_Model_Template.xlsx generated successfully.")

if __name__ == "__main__":
    create_template()
