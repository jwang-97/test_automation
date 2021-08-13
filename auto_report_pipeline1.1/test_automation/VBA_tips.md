# comment out function: "Application.Caller"
# private function: need write interface function
# two method to call macro: *call*(without parmeters) & *Application.Run*(with parmeters)
# Escape character should be watched out
# DisplayAlerts = False to aviod some popup windows
# ScreenUpdating = False to sppedup the function



# function : Private Sub cmdExport_Click():
to Make the instance hidden, add some code
change the ppt save-path to merge_template path(config.ini)
add title slide(type title) in frmExportProfile.ExportProfileCharts

# function : Private Sub ExportPlots(ppApp As PowerPoint.Application, blnStaticTitle As Boolean, blnTemplate As Boolean)
judge: if sting choose 0-4，6-7, else other static choose 0、1、3、4，6-7；
if sting, perc choose 50-95%

# function : UserForm_Activate
call *cmdClearAll_Click* to clearall the info which generate in the excel the last time 
call *Charts_Choose* to choose the case_type need tobe generate in the ppt
call *Templete_Choose* to decide the save path
call *cmdExport_Click* to export slides in ppt
Me.cmdExport.SetFocus & call *cmdCancel_Click* to close the showing form **frmExportProfile**


# function: Sub ReadBit
comment out the funtion *funcFilePicker*('Open the des/cfb file) to aviod popup windows
tips: all the plots generated in this function

# funtion : ProcessDes & ProcessCfb 
tips: read the info of bits in des&cfb file

# function : ProcessSummary
tips: read the summary file

# function : Auto_test
main function in SAM
*clear_template* del the merge_template dir in .ini
*add_slides* add title slides
*Auto_run_static* achieve the loop of clear & read & export



# frmExportProfile add-function:
**choose the static type to export**
Private Sub Charts_Choose()

    If Range("StingersPresent") Then
        Me.chkROP = True
        Me.chkROPChange = True
        Me.chkTorque = True
        Me.chkTqChange = True
        Me.chkROPTq = False
        Me.chkROPTqChange = False
        Me.chkDoC = True
        Me.chkTIF = True
        Me.chkBeta = True
        Me.chkRIF = True
        Me.chkCIF = True
        Me.chkRIF_CIF = False
        Me.chkRIF_CIF_Normal = False
    
    Else
        Me.chkROP = True
        Me.chkROPChange = True
        Me.chkTorque = True
        Me.chkTqChange = True
        Me.chkROPTq = True
        Me.chkROPTqChange = True
        Me.chkDoC = True
        Me.chkTIF = True
        Me.chkBeta = True
        Me.chkRIF = True
        Me.chkCIF = True
        Me.chkRIF_CIF = False
        Me.chkRIF_CIF_Normal = True

    End If
    
End Sub
**choose ppt path to export**
Private Sub Templete_Choose()
    Dim str As String
    Dim direct As String
    
    direct = ReadIniFileString("merge_template", "dirct")
    If direct = "" Then
        str = ReadIniFileString("templatefile_path", "template_path")
    Else
        str = direct
    End If
    If str = "" Then Exit Sub
    
    Range("ExTemplatePath") = str
    Me.txtTemplate = str
    Me.chkTemplate = True
    
End Sub

**read .ini config file**
Public Function IniFileName() As String
  IniFileName = Left$(Application.ActiveWorkbook.Path, InStrRev(Application.ActiveWorkbook.Path, "\")) & "docs\auto_test_config.ini"
End Function


Private Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String
Dim Worked As Long
Dim RetStr As String * 128
Dim StrSize As Long

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
  Else
    sProfileString = ""
    RetStr = Space(128)
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  ReadIniFileString = sIniString
End Function

# Multibit 

# function : Auto_Dynamic
main funtion of Multibit
*Run_dynamic* run process of dynamic(some click)

# function : subFindTemplate
write the ppt-path to generate slides

# function : fill_GEMs_Info
write the GEMs info under each slide