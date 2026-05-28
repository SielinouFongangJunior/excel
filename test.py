import os
import win32com.client as win32

script_dir = os.path.dirname(os.path.abspath(__file__))
enterfile = os.path.join(script_dir, "v1.xlsx")
outputfile = os.path.join(script_dir, "v1.pdf")

print(os.path.exists(enterfile))

excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(enterfile)
    ws = wb.Worksheets(1) 

    ws.PageSetup.Orientation = 2
    ws.PageSetup.Zoom = False     
    ws.PageSetup.FitToPagesWide = 1  
    ws.PageSetup.FitToPagesTall = 1  

    wb.ExportAsFixedFormat(0, outputfile)

    wb.Close()

    print(f"PDF généré avec succès : {outputfile}")

except Exception as e:
    print("Erreur :", e)

finally:
    excel.Quit()
