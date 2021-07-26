import win32com.client as win32
import time, os

def call_Multi_macro():
    xl=win32.Dispatch("Excel.Application")
    xl.Visible = True
    xl.Application.ScreenUpdating = False
    str = os.getcwd()
    Filename = str + "\\Multi_Bit_Compare_v2.52.xlsm"
    xl.Workbooks.Open(Filename)
    print (Filename)
    time.sleep(10)
    print('Open Multifile success')
    #Run Macro
    xl.Application.Run('Multi_Bit_Compare_v2.52.xlsm!Module1.Auto_Dynamic')
    print('Auto_Dynamic')
    xl.Application.ScreenUpdating = True
    
if __name__ == "__main__":
    call_Multi_macro()