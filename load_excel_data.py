import clr
from System.Runtime.InteropServices import Marshal

def load_excel_data(file_path, log_file):
    """
    Read load cases from an Excel file and return list of (Fx, Fy, Fz, Mx, My).
    """
    data = []
    try:
        clr.AddReference('Microsoft.Office.Interop.Excel')
        from Microsoft.Office.Interop import Excel
        excel_app = Excel.ApplicationClass()
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Open(file_path)
        sheet = workbook.Worksheets[1]
        last_row = sheet.Cells(sheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        for i in range(1, last_row + 1):
            fx = sheet.Cells(i, 1).Value2
            fy = sheet.Cells(i, 2).Value2
            fz = sheet.Cells(i, 3).Value2
            mx = sheet.Cells(i, 4).Value2
            my = sheet.Cells(i, 5).Value2
            if None in (fx, fy, fz, mx, my):
                continue
            data.append((float(fx), float(fy), float(fz), float(mx), float(my)))
        workbook.Close(SaveChanges=False)
        excel_app.Quit()
        Marshal.ReleaseComObject(sheet)
        Marshal.ReleaseComObject(workbook)
        Marshal.ReleaseComObject(excel_app)
    except Exception as err:
        log_file.write(f"ERROR: Failed to read Excel file '{file_path}'. Exception: {err}\n")
        raise
    log_file.write(f"Loaded {len(data)} load cases from '{file_path}'.\n")
    return data