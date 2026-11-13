import sys
import os
import csv
import openpyxl
from openpyxl.utils import get_column_letter

def main():
    ticker = "AVGO"
    if len(sys.argv) > 1:
        ticker = sys.argv[1].upper()
        
    template_path = "DCF_Model_Template.xlsm"
    csv_path = "edgar_data.csv"
    output_path = f"DCF_Model_{ticker}.xlsx"
    
    if not os.path.exists(template_path):
        print(f"Error: Template {template_path} not found.")
        sys.exit(1)
        
    if not os.path.exists(csv_path):
        print(f"Error: Edgar data {csv_path} not found.")
        sys.exit(1)
        
    print(f"Loading template {template_path}...")
    wb = openpyxl.load_workbook(template_path, keep_vba=True)
    ws = wb["Statements"]
    
    # Read the CSV data
    print(f"Reading data from {csv_path}...")
    data = {}
    years = []
    company_name = "Broadcom Inc."
    
    q_cash = 0.0
    q_debt = 0.0
    q_shares = 1.0
    
    with open(csv_path, mode='r') as f:
        reader = csv.reader(f)
        header = next(reader)
        years = [int(y) for y in header[1:]]
        for row in reader:
            if not row:
                continue
            key = row[0]
            if key == 'CompanyName':
                company_name = row[-1]
            elif key == 'LatestQuarterlyCash':
                q_cash = float(row[-1])
            elif key == 'LatestQuarterlyDebt':
                q_debt = float(row[-1])
            elif key == 'LatestQuarterlyShares':
                q_shares = float(row[-1])
            else:
                vals = [float(v) for v in row[1:]]
                data[key] = vals
                
    print(f"Company: {company_name}")
    print(f"Years found: {years}")
    
    # Reset all hidden rows to visible (up to row 135 to cover the PP&E schedule!)
    for r in range(1, 135):
        ws.row_dimensions[r].hidden = False
        
    # Update Company name and ticker in C6 and C7
    ws["C6"] = company_name
    ws["C7"] = ticker
    
    # Update Overview Cash, Debt, and Shares (statically from latest quarter!)
    ws["I6"] = q_cash
    ws["I7"] = q_debt
    ws["F7"] = q_shares
    
    # Column mapping: historical columns in sheet are E, F, G, H, I (columns 5, 6, 7, 8, 9)
    col_indices = [5, 6, 7, 8, 9]
    
    # Write years to Row 10, Row 59, Row 102, and Row 121
    for idx, yr in enumerate(years):
        col_let = get_column_letter(col_indices[idx])
        ws[f"{col_let}10"] = f"{yr}A"
        ws[f"{col_let}59"] = f"{yr}A"
        ws[f"{col_let}102"] = f"{yr}A"
        ws[f"{col_let}121"] = f"{yr}A"
        
    # Sheet row mappings (Excluding Cash Flow statements as per user request)
    row_mappings = {
        # Income Statement
        'Revenue': 12,
        'COGS': 15,
        'OperationSupport': 19,
        'SalesMarketing': 20,
        'ResearchDevelopment': 21,
        'GeneralAdministrative': 22,
        'DepreciationAmortizationIS': 23,
        'InterestExpense': 27,
        'OtherIncome': 28,
        'TaxProvision': 31,
        'EquityInvestmentIncome': 32,
        'NCI_Income': 34,
        'StockBasedCompensation': 38,
        'DilutedShares': 41,
        
        # Balance Sheet
        'CashEquivalents': 63,
        'ShortTermInvestments': 64,
        'RestrictedCashCurrent': 65,
        'AccountsReceivable': 66,
        'PrepaidExpenses': 67,
        'RestrictedCashNoncurrent': 69,
        'RestrictedInvestmentsNoncurrent': 70,
        'InvestmentsNoncurrent': 71,
        'EquityMethodInvestments': 72,
        'PP_E': 73,
        'LeaseAssets': 74,
        'Intangibles': 75,
        'Goodwill': 76,
        'OtherAssets': 77,
        'AccountsPayable': 81,
        'ShortTermDebt': 82,
        'LeaseLiabilitiesCurrent': 83,
        'AccruedLiabilitiesCurrent': 84,
        'LongTermInsuranceReserves': 86,
        'LongTermDebt': 87,
        'LeaseLiabilitiesNoncurrent': 88,
        'OtherLiabilitiesNoncurrent': 89,
        'RedeemableNCI': 91,
        'APIC': 94,
        'AccumOCI': 95,
        'AccumDeficit': 96,
        'NoncontrollingInterest': 98
    }
    
    # Write the values row-by-row
    for key, row_num in row_mappings.items():
        if key in data:
            vals = data[key]
            for idx, val in enumerate(vals):
                col_let = get_column_letter(col_indices[idx])
                ws[f"{col_let}{row_num}"] = val
                
    # Clear historical Cash Flow Statement input values (columns E to I) to keep them empty
    # We do NOT clear formula rows (111, 113, 118) to preserve the Excel subtotal formulas!
    cf_inputs = [106, 107, 108, 109, 110, 112, 114, 115, 116, 117]
    for r in cf_inputs:
        for col_idx in col_indices:
            ws.cell(row=r, column=col_idx).value = None
            
    # Populate PP&E Schedule (Rows 123 to 126, 133) for historical years 2022A-2025A (Columns F to I)
    # Do NOT overwrite projected columns J to N so template projection formulas remain untouched!
    ws["E123"] = None
    ws["E124"] = None
    ws["E125"] = None
    ws["E126"] = None
    ws["E133"] = None
    
    # 2022A to 2025A (Columns F to I, indices 1 to 4)
    for idx in range(1, 5):
        col_let = get_column_letter(col_indices[idx])
        prev_col_let = get_column_letter(col_indices[idx - 1])
        
        # Row 123: Beginning of period = Prior year's EOP PP&E
        if idx == 1:
            ws[f"{col_let}123"] = "=E73"
        else:
            ws[f"{col_let}123"] = f"={prev_col_let}126"
        
        # Row 124: Plus: CapEX = positive historical CapEX from EDGAR data
        capex_val = 0.0
        if 'CF_Capex' in data and idx < len(data['CF_Capex']):
            capex_val = abs(data['CF_Capex'][idx])
        ws[f"{col_let}124"] = capex_val
        
        # Row 125: Less: Depreciation = negative historical D&A from EDGAR data
        da_val = 0.0
        if 'DepreciationAmortizationIS' in data and idx < len(data['DepreciationAmortizationIS']):
            da_val = abs(data['DepreciationAmortizationIS'][idx])
        ws[f"{col_let}125"] = -da_val
        
        # Row 126: End of Period = Balance Sheet PP&E (Row 73)
        ws[f"{col_let}126"] = f"={col_let}73"
        
        # Row 133: Total D&A = Income Statement D&A (Row 23)
        ws[f"{col_let}133"] = f"={col_let}23"

    # Ensure Row 132 (% of revenue) in projected columns J to N links to prior column (=I132..=M132)
    # to prevent circular reference with Row 131 (=J132*J12)
    ws["J132"] = "=I132"
    ws["K132"] = "=J132"
    ws["L132"] = "=K132"
    ws["M132"] = "=L132"
    ws["N132"] = "=M132"
                 
    # Hide empty rows (mirroring VBA AutoHideEmptyRows)
    # Row 23 (D&A) is excluded from rows_to_check so it is NEVER hidden!
    rows_to_check = [19, 20, 21, 32, 34, 65, 69, 70, 72, 82, 86, 91, 98, 115, 117]
    for r in rows_to_check:
        is_empty = True
        for col_idx in col_indices:
            cell_val = ws.cell(row=r, column=col_idx).value
            if cell_val is not None:
                try:
                    val_float = float(cell_val)
                    if abs(val_float) > 0.0001:
                        is_empty = False
                        break
                except ValueError:
                    # If it's a formula or string, don't treat it as zero
                    is_empty = False
                    break
                
        if r == 34:
            # If NCI is zero, hide both 33 and 34
            ws.row_dimensions[33].hidden = is_empty
            ws.row_dimensions[34].hidden = is_empty
        else:
            ws.row_dimensions[r].hidden = is_empty
            
    # Save the populated model (both as XLSX and XLSM to ensure user receives the format they expect)
    print(f"Saving populated model to {output_path}...")
    wb.save(output_path)
    output_path_xlsm = output_path.replace(".xlsx", ".xlsm")
    print(f"Saving macro-enabled copy to {output_path_xlsm}...")
    wb.save(output_path_xlsm)
    print("Done!")

if __name__ == "__main__":
    main()
