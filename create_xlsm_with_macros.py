import os
import openpyxl
import win32com.client as win32

def clear_excel_template(file_path):
    print("Clearing historical data from template to make it clean...")
    wb = openpyxl.load_workbook(file_path)
    ws = wb["Statements"]
    
    # 1. Clear Company Name and Ticker
    ws["C6"] = ""
    ws["C7"] = ""
    
    # 2. Clear Overview Cash, Debt, and Shares
    ws["I6"] = ""
    ws["I7"] = ""
    ws["F7"] = ""
    
    # 3. Mapped rows to clear across columns E to I (columns 5 to 9)
    rows_to_clear = [
        # Income Statement
        12, 15, 19, 20, 21, 22, 23, 27, 28, 31, 32, 34, 41,
        # Balance Sheet
        63, 64, 65, 66, 67, 69, 70, 71, 72, 73, 74, 75, 76, 77, 81, 82, 83, 84, 86, 87, 88, 89, 91, 94, 95, 96, 98,
        # Cash Flow
        106, 107, 108, 109, 110, 112, 114, 115, 116, 117
    ]
    
    # Year headers
    year_rows = [10, 59, 102]
    
    for col_idx in range(5, 10): # Columns E to I
        # Clear year headers
        for yr_row in year_rows:
            ws.cell(row=yr_row, column=col_idx, value="")
            
        # Clear data values
        for r in rows_to_clear:
            ws.cell(row=r, column=col_idx, value="")
            
    # Reset row dimensions (make all visible again)
    for r in range(1, 160):
        ws.row_dimensions[r].hidden = False
        
    temp_clean_path = "DCF_Model_Template_Clean.xlsx"
    wb.save(temp_clean_path)
    print(f"Cleaned template saved to {temp_clean_path}")
    return temp_clean_path

def inject_macros(clean_xlsx_path):
    user_profile = os.environ["USERPROFILE"]
    project_dir = os.path.join(user_profile, ".gemini", "antigravity", "scratch", "edgar_dcf_excel")
    
    xlsx_abs_path = os.path.abspath(clean_xlsx_path)
    xlsm_abs_path = os.path.join(project_dir, "DCF_Model_Template.xlsm")
    
    # Load VBA scripts
    vba_module_path = os.path.join(project_dir, "EDGAR_Fetcher.bas")
    vba_sheet_path = os.path.join(project_dir, "Statements_Worksheet_Code.txt")
    
    if not os.path.exists(vba_module_path) or not os.path.exists(vba_sheet_path):
        print("Error: VBA source files not found.")
        return False
        
    with open(vba_module_path, "r") as f:
        module_code = f.read()
        
    with open(vba_sheet_path, "r") as f:
        sheet_code = f.read()
        
    print("Launching Excel to inject macros...")
    excel = None
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open workbook
        wb = excel.Workbooks.Open(xlsx_abs_path)
        
        # Access VBA Project
        try:
            vba_project = wb.VBProject
        except Exception as e:
            print("\n" + "="*60)
            print("SECURITY BLOCK: Excel is blocking programmatic access to VBA Project.")
            print("To fix this, please enable trust access in Excel:")
            print(" 1. Open Excel.")
            print(" 2. Go to File -> Options -> Trust Center -> Trust Center Settings.")
            print(" 3. Go to Macro Settings.")
            print(" 4. Check the box 'Trust access to the VBA project object model'.")
            print(" 5. Click OK, close Excel, and let me know so I can try again!")
            print("="*60 + "\n")
            wb.Close(False)
            excel.Quit()
            return False
            
        # Add Module EDGAR_Fetcher
        print("Adding VBA Module 'EDGAR_Fetcher'...")
        vba_module = vba_project.VBComponents.Add(1) # 1 = vbext_ct_StdModule
        vba_module.Name = "EDGAR_Fetcher"
        vba_module.CodeModule.AddFromString(module_code)
        
        # Find Statements sheet VBComponent and add Worksheet_Change event code
        print("Adding Worksheet event handler to 'Statements' sheet...")
        for comp in vba_project.VBComponents:
            if comp.Name == "Sheet1": # Sheet1 represents the Statements sheet codename
                comp.CodeModule.AddFromString(sheet_code)
                break
                
        # Save as Macro-Enabled Workbook (.xlsm)
        # FileFormat 52 = xlOpenXMLWorkbookMacroEnabled
        print(f"Saving as {xlsm_abs_path}...")
        wb.SaveAs(xlsm_abs_path, FileFormat=52)
        wb.Close(True)
        excel.Quit()
        print("VBA Macros successfully injected!")
        return True
        
    except Exception as e:
        print(f"Error during macro injection: {str(e)}")
        if excel:
            try:
                excel.Quit()
            except:
                pass
        return False

def main():
    clean_xlsx = clear_excel_template("DCF_Model_Template.xlsx")
    success = inject_macros(clean_xlsx)
    if success:
        # Clean up temporary clean xlsx file
        if os.path.exists(clean_xlsx):
            os.remove(clean_xlsx)
            
if __name__ == "__main__":
    main()
