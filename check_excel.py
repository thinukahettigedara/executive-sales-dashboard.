import sys

try:
    import win32com.client
    xl = win32com.client.Dispatch("Excel.Application")
    print("Excel version:", xl.Version)
    print("Workbooks count:", xl.Workbooks.Count)
    for wb in xl.Workbooks:
        print(f"Workbook: {wb.Name} | FullName: {wb.FullName} | Saved: {wb.Saved}")
except Exception as e:
    print("Error:", e)
