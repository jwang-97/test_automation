Sub Auto_Dynamic()
    Call Run_dynamic
    Call clear_template
    Call Close_Excel_Without_Save
End Sub

Private Sub Run_dynamic()
    Dim WorkbookName As String
    Dim MacroName1 As String
    Dim MacroName2 As String
    Dim MacroName3 As String
    Dim MacroName4 As String
    Dim MacroName5 As String
    Dim MacroName6 As String
    Dim MacroName7 As String


    WorkbookName = "test_Multi-Bit_Compare_v2.52.xlsm"
    Workbooks(WorkbookName).Activate
    MacroName1 = "Module1.subFindTemplate"
    MacroName2 = "Module1.BitBasePicker"
    MacroName3 = "Module1.StartOver"
    MacroName4 = "Module1.ToggleIDEASCompare"
    MacroName5 = "Module1.subMakeSlides"
    MacroName6 = "Module1.BitCountSelect"
    MacroName7 = "Module1.fill_GEMs_Info"


    Application.Run "'" & WorkbookName & "'!" & MacroName3
    Application.Run "'" & WorkbookName & "'!" & MacroName6
    Application.Run "'" & WorkbookName & "'!" & MacroName1
    Application.Run "'" & WorkbookName & "'!" & MacroName2
    Application.Run "'" & WorkbookName & "'!" & MacroName4
    Application.Run "'" & WorkbookName & "'!" & MacroName7
    Application.Run "'" & WorkbookName & "'!" & MacroName5



End Sub

Private Sub clear_template()
Dim TF As String
    TF = WritePrivateProfileString("merge_template", "dirct", "", IniFileName)
End Sub

Private Sub Close_Excel_Without_Save()
    Application.Visible = False
    Application.DisplayAlerts = False
    ActiveWorkbook.Close SaveChanges:=False
    Application.Quit
End Sub

