Option Explicit
'Public ppApp As Powerpoint.Application
'Public ppPres As Powerpoint.Presentation
'Public ActFileName As Variant
'Public totalslide As Object
'Global DESfile_path1 As String
'Global DESfile_path2 As String
'Global pptTemplatefile_path As String
'
'
'   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
'                 (ByVal lpApplicationName As String, _
'                   ByVal lpKeyName As String, _
'                   ByVal lpDefault As String, _
'                   ByVal lpReturnedString As String, _
'                   ByVal nSize As Long, _
'                   ByVal lpFileName As String) As Long
'Sub Read()
'    Dim Rec As String
'    Dim NC As Long
'    Dim Workbook_path As String
'
'    Rec = String(255, 0)
'    Workbook_path = Application.ActiveWorkbook.Path & "\auto_test_config.ini"
''    NC = GetPrivateProfileString("SAMfile_path", "file_path", "", SAMfile_path, 255, Workbook_path)
'    NC = GetPrivateProfileString("DESfile_path", "cases_path1", "", Rec, 255, Workbook_path)
'    DESfile_path1 = Trim(Rec)
'    NC = GetPrivateProfileString("DESfile_path", "cases_path2", "", Rec, 255, Workbook_path)
'    DESfile_path2 = Trim(Rec)
'
'End Sub

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function IniFileName() As String
  IniFileName = Application.ActiveWorkbook.Path & "\auto_test_config.ini"
End Function


Private Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String
Dim Worked As Long
Dim RetStr As String * 128
Dim StrSize As Long
Dim iNoOfCharInIni As String
Dim sIniString As String
Dim sProfileString As Variant

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

'Sub Read_ini()
'    Rec = ReadIniFileString("DESfile_path", "cases_path1")
'    MsgBox Rec
'End Sub

'Sub open_file_demo()
'
'    'declare variable
'    Dim pathname
'
'    ' assign a value
'    'open background
'    'hide excel
''    Application.Visible = False
'    'close screen update
''    Application.ScreenUpdating = False
'    'close alerts
''    Application.DisplayAlerts = False
''    ActiveWindow.Visible = False
'    pathname = Application.ActiveWorkbook.Path & "\test_SAM_v11.78.xlsm"
'    ' now open the file using the open statement
'    Workbooks.Open pathname
''    ActiveWindow.Visible = False
'    ThisWorkbook.Activate
''    Application.ScreenUpdating = True
''    Application.DisplayAlerts = True
''    Application.Visible = True
''    Excel.Application.Visible = False
'End Sub

Sub Auto_run_static()

    Dim WorkbookName As String
    Dim MacroName1 As String
    Dim MacroName2 As String
    Dim MacroName3 As String
    Dim MacroName4 As String
    Dim argument1 As String
    Dim argument2 As Integer
    Dim argument3 As Byte
    Dim argument4 As Integer
    Dim argument6 As Integer
    Dim path_array(20) As String
    Dim i As Integer

    WorkbookName = "test_SAM_v11.78.xlsm"
    Workbooks("test_SAM_v11.78.xlsm").Activate
    MacroName1 = "Module1.ReadBit"
    MacroName2 = "Module1.ExportProfileChartButton"
    MacroName3 = "Module1.Clear_All"
'    MacroName3 = "frmExportProfile.interface_outside"
'    MacroName4 = "Module1.new_title_slide"


'    Call Module1.Read(read ini file)
'    argument1 = ReadIniFileString("DESfile_path", "cases_path1")
    argument2 = False
    argument3 = 1
    argument4 = False
'    argument5 = ReadIniFileString("DESfile_path", "cases_path2")
    argument6 = 2

    path_array(0) = ReadIniFileString("DESfile_pdc_path", "cases_path1")
    path_array(1) = ReadIniFileString("DESfile_pdc_path", "cases_path2")
    path_array(2) = ReadIniFileString("DESfile_stinger_path", "cases_path1")
    path_array(3) = ReadIniFileString("DESfile_stinger_path", "cases_path2")
    path_array(4) = ReadIniFileString("DESfile_axe_path", "cases_path1")
    path_array(5) = ReadIniFileString("DESfile_axe_path", "cases_path2")
    path_array(6) = ReadIniFileString("DESfile_hyper_path", "cases_path1")
    path_array(7) = ReadIniFileString("DESfile_hyper_path", "cases_path2")
    path_array(8) = ReadIniFileString("DESfile_px_path", "cases_path1")
    path_array(9) = ReadIniFileString("DESfile_px_path", "cases_path2")
    path_array(10) = ReadIniFileString("DESfile_echo_path", "cases_path1")
    path_array(11) = ReadIniFileString("DESfile_echo_path", "cases_path2")
    path_array(12) = ReadIniFileString("DESfile_E616_path", "cases_path1")
    path_array(13) = ReadIniFileString("DESfile_E616_path", "cases_path2")

'    argument7 = ReadIniFileString("DESfile_path", "cases_path2")
'    argument7 = ReadIniFileString("DESfile_path", "cases_path2")

'   Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(10), argument2, argument3, argument4
'   Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(11), argument2, argument6, argument4


    For i = 0 To 6
        Application.Run "'" & WorkbookName & "'!" & MacroName3

        Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(i * 2), argument2, argument3, argument4
        Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(i * 2 + 1), argument2, argument6, argument4
        Application.Run "'" & WorkbookName & "'!" & MacroName2
    Next i


'    Application.Run "'" & WorkbookName & "'!" & MacroName4

    'Application.Run "'" & WorkbookName & "'!" & MacroName3
    'Worksheets("frmExportProfile").Activate
    'Application.Run "test_SAM_v11.78.xlsm!frmExportProfile.optColorNew_Click"
    'Application.Run "'" & WorkbookName & "'!" & MacroName4

End Sub
Sub Auto_static()

'    open_file_demo
    Auto_run_static
'    Call Close_Powerpoint_With_Save
    Close_Excel_Without_Save

End Sub
Sub Close_Excel_Without_Save()

    Application.Visible = False
    Application.DisplayAlerts = False
    ActiveWorkbook.Close SaveChanges:=False
    Application.Quit

End Sub

'Sub ppt()
'    Dim ppt As New PowerPoint.Application
'    Dim pres As PowerPoint.Presentation
'    ppt.Visible = True
'    Set pres = ppt.Presentations.Add
'    pres.SaveAs "C:\Users\xxx\Desktop\ppttest.pptx"
'    pres.Close
'    ppt.Quit
'    Set ppt = Nothing
'End Sub

'Sub Close_Powerpoint_With_Save()
'
'    Powerpoint.ActivePresentation.Save
''    PPT.ActivePresentation.SaveAs FileName
'    waiting (7) 'depends on size of presentation
''    PPT.Presentations(filename2).Close
'    Powerpoint.Presentation.Close
'    Powerpoint.Application.Quit
'    Set ppt = Nothing
'
'End Sub
'Sub waiting(tsecs As Single)
'
'    Dim sngsec As Single
'    sngsec = Timer + tsecs
'    Do While Timer < sngsec
'        DoEvents
'    Loop
'
'End Sub

'Sub ExceltoPPT()
'    'ppt template
'    ActFileName = Application.GetOpenFilename("")
'    Set ppApp = CreateObject("PowerPoint.Application")
'
'    ppApp.vsiable = True
'    ppApp.Active
'    Set ppPress = ppApp.Presentation.Open(ActFileName)
'    CreatePlan
'    Set ppPres = Nothing
'    Set ppApp = Nothing
'End Sub
'Sub CreatePlan()
'    ppApp.ActivePresentation.Slide.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutCustom
'End Sub

Sub open_Multiply_demo()
    Dim pathname
    pathname = Application.ActiveWorkbook.Path & "\test_Multi-Bit_Compare_v2.52.xlsm"
    Workbooks.Open pathname
End Sub
Sub Run_dynamic()
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
Sub clear_template()
Dim TF As String
    TF = WritePrivateProfileString("merge_template", "dirct", "", IniFileName)
End Sub
Sub Auto_dynamic_test()

    open_Multiply_demo
    Run_dynamic
'    Call Close_Excel_Without_Save

End Sub
Public Function GetLayout(ppPres As PowerPoint.Presentation, LayoutName As String, Optional ParentPresentation As Presentation = Nothing) As CustomLayout
'http://stackoverflow.com/questions/20997571/create-a-new-slide-in-vba-for-powerpoint-2010-with-custom-layout-using-master

    If ParentPresentation Is Nothing Then
        Set ParentPresentation = ppPres
    End If

    Dim oLayout As CustomLayout
    For Each oLayout In ParentPresentation.SlideMaster.CustomLayouts
        If oLayout.Name = LayoutName Then
            Set GetLayout = oLayout
            Exit For
        End If
    Next

End Function
Sub add_slides()
    Dim str As String, strDir As String, blnStaticTitle As Boolean
    Dim i As Integer
    Dim ppApp As PowerPoint.Application
    Dim direct As String
    Dim TF As Long
    Dim merge_template As String
    Dim ppSlide As Variant

'Create new instance if no instance exists
    If ppApp Is Nothing Then Set ppApp = New PowerPoint.Application
'Make the instance visible
    ppApp.Visible = True

    ppApp.Presentations.Open ReadIniFileString("merge_template", "dirct"), msoTrue

    strDir = "SAM_temp_" & Format(Now, "YYYY-MM-DD_hh.nn.ss") & ".pptx"
    merge_template = ppApp.Presentations(ppApp.Presentations.Count).Path & "\" & strDir
    TF = WritePrivateProfileString("merge_template", "dirct", merge_template, IniFileName)
    ppApp.Presentations(ppApp.Presentations.Count).SaveAs merge_template

    ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank") ' "2_Title Slide")
    Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
            'slide title
    ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Autoforce Comparison"
    '         'slide subtitle
    ' ppSlide.Shapes(1).TextFrame.TextRange.Text = "Build 20191231 vs Build 20200403"
            ' i = 0
            ' title_array = {"Important Build Notes", "Overview", "Design Tool Functionality Checks"}
    i = 1
    For i = 1 To 2
        ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Autoforce – PDC test"
        Next i
    ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank") ' "2_Title Slide")
    Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
    ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Contact Analysis Comparison"

    ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Section Title")
    Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
    ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Contact Analysis"
    ppSlide.Shapes(1).TextFrame.TextRange.Text = "Comparison of Baseline Bit and Proposed Bit MDOC Analysis"

    i = 1
    For i = 1 To 14
        ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Contact Analysis Comparison - Blade Top Contact"
            ' ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
            ' ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Design Tool Functionality Checks"
            ' ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank")
            ' ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Static Analysis Results"
        Next i
    ppApp.Presentations(ppApp.Presentations.Count).SaveAs merge_template
    ppApp.Presentations(ppApp.Presentations.Count).Close
'    ppApp.Quit
    Set ppApp = Nothing
End Sub
'Private Sub Auto_test()
'    Call clear_template
'    Call open_file_demo
'    Call Auto_run_static
'    Call add_slides
'    Call open_Multiply_demo
'    Call Run_dynamic
'    Call clear_template
'    Call Close_Excel_Without_Save
'End Sub
Private Sub Workbook_Open()
    clear_template
'    open_file_demo
    Auto_run_static
    add_slides
    open_Multiply_demo
    Run_dynamic
    clear_template
    Close_Excel_Without_Save
End Sub


