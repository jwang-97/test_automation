import win32com.client as win32

def test_write_excel(path) :
    xl = win32.Dispatch("Excel.Application")
    # with xl.Workbooks.Open(Filename = path, ReadOnly = False) as xlbook:
    #     sheet = xlbook.Worksheets("Sheet1")
    #     sheet.Range("A1").Value = "test"
    xlbook = xl.Workbooks.Open(Filename = path, ReadOnly = False)
    xlbook.DisplayAlerts =False
    xlbook.Visible=False
    sheet = xlbook.Worksheets("Sheet1")
    sheet.Range("A1").Value = "test"
    # sheet.Range("A2").Value = "test1"
    write_excel_via_other_function(xlbook)
    xl.DisplayAlerts = False
    xl.ActiveWorkbook.Save()
    xl.ActiveWorkbook.Close()
    write_excel_in_other_function(path)

    # xl.Application.DisplayAlerts = False
    # xlbook.Close()

def write_excel_in_other_function(path):
    xl = win32.Dispatch("Excel.Application")
    # with xl.Workbooks.Open(Filename = path, ReadOnly = False) as xlbook:
    #     sheet = xlbook.Worksheets("Sheet1")
    #     sheet.Range("A1").Value = "test"
    xlbook = xl.Workbooks.Open(Filename = path, ReadOnly = False)
    xlbook.DisplayAlerts =False
    xlbook.Visible=False
    sheet = xlbook.Worksheets("Sheet1")
    # sheet.Range("A1").Value = "test"
    sheet.Range("A2").Value = "test1"
    # xl.Application.DisplayAlerts = False
    xl.DisplayAlerts = False
    xl.ActiveWorkbook.Save()
    xl.ActiveWorkbook.Close()
    
def write_excel_via_other_function(xlbook):
    sheet = xlbook.Worksheets("Sheet1")
    # sheet.Range("A1").Value = "test"
    sheet.Range("A3").Value = "test2"
    # xl.Application.DisplayAlerts = False
if __name__ == "__main__":
    test_write_excel(r"C:\Users\JWang294\Documents\projectfile\vba-files\test.xlsx")
