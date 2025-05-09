import clr
from System.Runtime.InteropServices import Marshal

def write_excel_results(load_cases, max_stress_values, output_excel_path, log_file):
    """
    Write load cases and results into a new Excel file.
    """
    try:
        clr.AddReference('Microsoft.Office.Interop.Excel')
        from Microsoft.Office.Interop import Excel

        excel_app = Excel.ApplicationClass()
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Add()
        sheet = workbook.Worksheets[1]

        headers = ["Case", "Fx (N)", "Fy (N)", "Fz (N)", "Mx (N*m)", "My (N*m)", "Max Von Mises (Pa)"]
        for col, header in enumerate(headers, start=1):
            sheet.Cells(1, col).Value2 = header

        for idx, (loads, stress) in enumerate(zip(load_cases, max_stress_values), start=2):
            fx, fy, fz, mx, my = loads
            sheet.Cells(idx, 1).Value2 = idx - 1
            sheet.Cells(idx, 2).Value2 = fx
            sheet.Cells(idx, 3).Value2 = fy
            sheet.Cells(idx, 4).Value2 = fz
            sheet.Cells(idx, 5).Value2 = mx
            sheet.Cells(idx, 6).Value2 = my
            sheet.Cells(idx, 7).Value2 = stress if stress is not None else "SolveFailed"

        workbook.SaveAs(output_excel_path)
        workbook.Close(SaveChanges=True)
        excel_app.Quit()
        Marshal.ReleaseComObject(sheet)
        Marshal.ReleaseComObject(workbook)
        Marshal.ReleaseComObject(excel_app)
        log_file.write(f"Results written to Excel file: {output_excel_path}\n")
    except Exception as err:
        log_file.write(f"ERROR: Could not write results. Exception: {err}\n")
        raise