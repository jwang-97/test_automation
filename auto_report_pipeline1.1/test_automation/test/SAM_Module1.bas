Attribute VB_Name = "Module1"
'Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Dim Current_Workbook As String
Dim Bit_Load_Status As String
Dim Bit_Warn_Status As String
Global Debug_flag As Boolean
Global gblnZoIOverride As Boolean
Global gblnZoIShtOverride As Boolean

Global gstrPRTracker As String
Global giBit As Integer
Global gProc As String
Global giCount As Integer

Global sStart(5) As Single, sEnd(5) As Single
Public blnTitleFix As Boolean


Sub ShowHideBackups(Optional BitNum As Variant, Optional strType As String)

    gProc = "ShowHideBackups"
    
    gblnZoIShtOverride = True
    
    If IsMissing(BitNum) Then BitNum = Right(Application.Caller, 1)
    If strType = "" Then strType = Application.Caller
    
    If Left(strType, 4) = "Show" Then
        A = HideCutters(BitNum, False, False, False)
    ElseIf Left(strType, 8) = "Hide_Bit" Then
        A = HideCutters(BitNum, False, True, False)
    ElseIf Left(strType, 12) = "Hide_Backups" Then
        A = HideCutters(BitNum, True, True, False)
    ElseIf Left(strType, 16) = "Hide_Off_Backups" Then
        A = HideCutters(BitNum, True, True, True)
    End If
    Sheets("HOME").Activate
    
    gblnZoIShtOverride = False
    
End Sub
Sub ClearBit(Optional BitNum As Variant)
Dim i As Integer

    gProc = "ClearBit"
    
    On Error GoTo ErrHandler

    subActive True

    If Application.Caller = "cmdReset" Then
        If MsgBox("This will clear all bits and reset to Case 1. Are you sure you wish to do this?", vbYesNo, "Are you sure?") = vbNo Then
            Exit Sub
        End If
    End If
        
    
    Debug_flag = False
    subUpdate False
    
    If IsMissing(BitNum) Then BitNum = Right(Application.Caller, 1)
    
    If IsNumeric(BitNum) = False Then
        BitNum = 0
    Else
        BitNum = CInt(BitNum)
    End If
    
    If BitNum = 0 Then
        subClearRFPlot
        Sheets("SUMMARY").Activate
        Range("B6").Name = "CasesID"
    End If
    
    For i = 1 To 5
        If BitNum = 0 Or BitNum = i Then
            Clear_Cutter_Info (i)
            Clear_Summary (i)
            Range("BIT" & i & "_Name").Value = ""
            Range("Bit" & i & "_Status").Value = ""

            'Clear Info2
            Range(Range("Bit" & i & "_I2_ConeSize"), Range("Bit" & i & "_I2_GaugeSR")).ClearContents
            
            'Clear Drill_Scan
            Range("Bit" & i & "_DS_CS_CD").ClearContents
            Range("Bit" & i & "_DS_CS_Ht").ClearContents
            Range("Bit" & i & "_DS_CS_BR").ClearContents
            Range("Bit" & i & "_DS_AG_Ht").ClearContents
            Range("Bit" & i & "_DS_AG_BR").ClearContents
            
            'Clear Stingers Present
            Range("Bit" & i & "_StingersPresent") = False
            
            Sheets("HOME").Shapes("Hide_Bit" & i).ControlFormat.Value = 1
            A = HideCutters(i, False, True, False)
        End If
    Next i
    
    Sheets("HOME").Activate
    
    If BitNum = Range("CoreBit") Then
        Range("CoreDiam") = ""
        Range("StingerTip") = ""
    End If

    If IsRapidFire And BitNum = 0 Then
        subBaselineIteration 1, "BL"
        subBaselineIteration 2, "IT"
        subBaselineIteration 3, "IT"
        subBaselineIteration 4, "IT"
        subBaselineIteration 5, "IT"
        
'        Sheets("HOME").ListBoxes("RF").Value = 1
        Sheets("HOME").ListBoxes("RFPlotBox1").Value = 1
        Range("RFCaseID1") = 1
        
        subUpdateShowHide
    End If
    'hide zone
    'Range("Limits").Columns.EntireColumn.Hidden = True
    
    ZoI "Hide"
    gblnZoIShtOverride = False
    
    If BitNum = 0 Or BitNum = 1 Then SetCaseTo1
    
'    Sheets("HOME").OptionButtons("Option_Button_Hide_Zone").Value = Checked
'    Range("Limits").Columns.EntireColumn.Hidden = True
'    Range("Zone").Rows.EntireRow.Hidden = True
    
    Application.Calculation = xlCalculationAutomatic
    
    subActive False
    
    Exit Sub
    
ErrHandler:

    Sheets("HOME").Select
    
End Sub
Sub ReadBit(Optional Bit_File_Handler As String, Optional blnRefresh As Boolean, Optional iBit As Integer, Optional blnSkipIndices As Boolean)

    gProc = "ReadBit"
    
    Dim strVis As String, strTempWB As String, strSummaryFile As String, BitNum As Integer, A As String
    Dim Used_Range_Address As String, BFH_pos As Integer, i As Integer, strMsg As String, bln As Boolean
    
'Record the start "time"
    sStart(0) = Timer
    
    subActive True
  
'Begin ReadBit procedure
    Debug_flag = False
    Bit_Load_Status = ""
    Bit_Warn_Status = ""
    
    
    If iBit = 0 Then
        If Range("Bit1_Location") = "" Then
            BitNum = 1
        Else
            BitNum = 2
        End If
    Else
        BitNum = iBit
        Sheets("HOME").Select
    End If
    Current_Workbook = ActiveWorkbook.Name
    
    'hide zone
    ZoI "Hide"
    gblnZoIShtOverride = False
    
    'Open the des/cfb file
'    If Range("Bit" & BitNum & "_Location") = "" And Bit_File_Handler = "" Then
'        Bit_File_Handler = funcFilePicker("Select des/cfb file for Bit #" & BitNum, "des/cfb file", "*.des;*.cfb;*.cfa")
'    ElseIf Range("Bit" & BitNum & "_Location") <> "" And Not blnRefresh Then
'        Bit_File_Handler = funcFilePicker("Select des/cfb file for Bit #" & BitNum, "des/cfb file", "*.des;*.cfb;*.cfa", Mid(Range("Bit" & BitNum & "_Location"), InStr(1, Range("Bit" & BitNum & "_Location"), " ", vbTextCompare) + 1, InStrRev(Range("Bit" & BitNum & "_Location"), "\", , vbTextCompare) - InStr(1, Range("Bit" & BitNum & "_Location"), " ", vbTextCompare)), False)
'    End If
    
    If Bit_File_Handler = "False" Or Bit_File_Handler = "" Then
        subActive False
        Exit Sub
    End If
            
    If ActiveSheet.CheckBoxes("chkCheckBitIn").Value = 1 Then
        strVis = CheckBitIn(Bit_File_Handler, "Name")
        If strVis = "" Then
            'no bit.in found. Trying reamer.in.
            strVis = CheckReamerIn(Bit_File_Handler, "Name")
            If strVis = "" Then
                'no reamer.in found. Loading just Bit info.
            ElseIf strVis <> Right(Bit_File_Handler, Len(Bit_File_Handler) - InStrRev(Bit_File_Handler, "\")) Then
                If MsgBox("The active bit in this project is: " & strVis & vbCrLf & vbCrLf & "You've loaded: " & Right(Bit_File_Handler, Len(Bit_File_Handler) - InStrRev(Bit_File_Handler, "\")) & vbCrLf & vbCrLf & "Continue loading """ & Right(Bit_File_Handler, Len(Bit_File_Handler) - InStrRev(Bit_File_Handler, "\")) & """?", vbYesNo, "Not Active Bit") = vbNo Then
                    subActive False
                    Exit Sub
                End If
            End If
        ElseIf strVis <> Right(Bit_File_Handler, Len(Bit_File_Handler) - InStrRev(Bit_File_Handler, "\")) Then
            If MsgBox("The active bit in this project is: " & strVis & vbCrLf & vbCrLf & "You've loaded: " & Right(Bit_File_Handler, Len(Bit_File_Handler) - InStrRev(Bit_File_Handler, "\")) & vbCrLf & vbCrLf & "Continue loading """ & Right(Bit_File_Handler, Len(Bit_File_Handler) - InStrRev(Bit_File_Handler, "\")) & """?", vbYesNo, "Not Active Bit") = vbNo Then
                subActive False
                Exit Sub
            End If
        End If
    End If

    sStart(1) = Timer
    
    If blnRefresh Then
        If ActiveSheet.Shapes("Show_Bit" & BitNum).ControlFormat.Value = 1 Then
            strVis = "Show_Bit"
        ElseIf ActiveSheet.Shapes("Hide_Bit" & BitNum).ControlFormat.Value = 1 Then
            strVis = "Hide_Bit"
        ElseIf ActiveSheet.Shapes("Hide_Backups_Bit" & BitNum).ControlFormat.Value = 1 Then
            strVis = "Hide_Backups_Bit"
        ElseIf ActiveSheet.Shapes("Hide_Off_Backups_Bit" & BitNum).ControlFormat.Value = 1 Then
            strVis = "Hide_Off_Backups_Bit"
        End If
    Else
        strVis = "Show_Bit"
    End If
    
    ActiveSheet.Shapes("Show_Bit" & BitNum).ControlFormat.Value = 1
    On Error GoTo Err_Hide:
        A = HideCutters(BitNum, False, False, False)
    
    
    On Error GoTo Err_Clear:
        Clear_Cutter_Info (BitNum)
    Range("Bit" & BitNum & "_Status").Value = ""
    On Error GoTo Err_Open:
    
    
    
    Range("Bit" & BitNum & "_Location").Value = "[" & Format(Now(), "hh:nn") & "] " & Bit_File_Handler
    
    
    
    sStart(2) = Timer
' Introducing new method of reading in file to help make SAM compatible with Excel 2013. -A. Barr
    If Range("ReadVer") = "old" Then
        Workbooks.OpenText FileName:=Bit_File_Handler, Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=True, Other:=True, TrailingMinusNumbers:=True
        strTempWB = ActiveWorkbook.Name
    Else
        If LCase(Right(Bit_File_Handler, 3)) = "des" Then
            strTempWB = OpenText(Bit_File_Handler, Range("ReadVer"))
        Else
            strTempWB = OpenText(Bit_File_Handler, Range("ReadVer"), "cfb")
        End If
    End If
    sEnd(2) = Timer
    
    Used_Range_Address = ActiveSheet.UsedRange.Address
    ActiveSheet.UsedRange.Copy
    Workbooks(Current_Workbook).Activate
    Sheets("BIT" & BitNum & "_CUTTER_FILE").Activate
    Range(Used_Range_Address).PasteSpecial
    
    
     'Close the summary file
    Application.DisplayAlerts = False
    If Range("ReadVer") = "old" Then
        Windows(strTempWB).Close
        Workbooks(Current_Workbook).Activate
        
    Else
        Sheets(strTempWB).Delete
        Sheets("BIT" & BitNum & "_CUTTER_FILE").Activate
'        ClearQueryTables
    End If
    Application.DisplayAlerts = True


    If LCase(Right(Bit_File_Handler, 3)) = "des" Then
        On Error GoTo Err_Des:
            A = ProcessDes(BitNum, Bit_File_Handler)
    Else
        On Error GoTo Err_Cfb:
            A = ProcessCfb(BitNum, Bit_File_Handler)
            If Not IsEmpty(A) Then AddBitWarnStatus Bit_Warn_Status, CStr(A)
    End If
    
'Calculate James' Bit Indices
    If blnSkipIndices = False Then subBitIndices BitNum
    
    
    BFH_pos = InStrRev(Bit_File_Handler, "\")
    strSummaryFile = Mid(Bit_File_Handler, 1, BFH_pos) & "summary_regular" 'changed the extension from summary_regular.*
    On Error GoTo Err_Summary:
    
    
    
    sStart(3) = Timer
    A = ProcessSummary(BitNum, strSummaryFile)
    sEnd(3) = Timer
    
    
    
    If A = "True" Then
        AddBitWarnStatus Bit_Warn_Status, "Case Mismatch"
    ElseIf A = "No_Summary" Then
        AddBitWarnStatus Bit_Warn_Status, "No Summary File Found"
    ElseIf A = "Stinger_Static" Then
        AddBitWarnStatus Bit_Warn_Status, "Stinger Static Project not yet compatible with SAM."
    End If
    
    If Range("Bit" & BitNum & "_Name").Value = "" Or Sheets("HOME").Shapes("chkAutoName").ControlFormat.Value = 1 Then
        If ActiveSheet.Shapes("optStandardMode").ControlFormat.Value <> 1 And BitNum = 1 Then
            Range("Bit" & BitNum & "_Name").Value = "Baseline"
        Else
            If Sheets("HOME").Shapes("chkAutoName").ControlFormat.Value = 1 Then
                If IsNumeric(Mid(Bit_File_Handler, InStrRev(Bit_File_Handler, "\") + 1, InStrRev(Bit_File_Handler, ".") - InStrRev(Bit_File_Handler, "\") - 1)) Then
                    Range("Bit" & BitNum & "_Name").Value = "'" & Mid(Bit_File_Handler, InStrRev(Bit_File_Handler, "\") + 1, InStrRev(Bit_File_Handler, ".") - InStrRev(Bit_File_Handler, "\") - 1)
                Else
                    Range("Bit" & BitNum & "_Name").Value = Mid(Bit_File_Handler, InStrRev(Bit_File_Handler, "\") + 1, InStrRev(Bit_File_Handler, ".") - InStrRev(Bit_File_Handler, "\") - 1)
                End If
            Else
                bln = Application.ScreenUpdating
                Application.ScreenUpdating = False
                If IsReamer(BitNum) Then
                    Range("Bit" & BitNum & "_Name").Value = "Reamer " & BitNum
                Else
                    Range("Bit" & BitNum & "_Name").Value = "Bit " & BitNum
                End If
                Application.ScreenUpdating = bln
            End If
        End If
    End If
    If Bit_Load_Status = "" Then Bit_Load_Status = "Loading bit info complete"
        
    
    
    
    If BitNum = 1 Then subForceDefaults
    
Err_Hide:
    If Err.Number <> 0 Then
        Bit_Load_Status = "Unable to Hide/Show cutters"
        If Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then MsgBox "hide"
        Resume Next
    End If
    
Err_Clear:
    If Err.Number <> 0 Then
        Bit_Load_Status = "Unable to Clear Bit info"
        If Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then MsgBox "clear"
        Sheets("HOME").Activate
    End If
        
Err_Open:
    If Err.Number <> 0 Then
        Bit_Load_Status = "Unable to Open file"
        If Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then MsgBox "open"
        Sheets("HOME").Activate
    End If
    
Err_Des:
    If Err.Number <> 0 Then
        Bit_Load_Status = "Unable to process des file"
        If Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then MsgBox "des"
        Sheets("HOME").Activate
    End If
    
Err_Cfb:
    If Err.Number <> 0 Then
        Bit_Load_Status = "Unable to process cfb file"
        If Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then MsgBox "cfb"
        Sheets("HOME").Activate
    End If

Err_Summary:
    If Err.Number <> 0 Then
        Bit_Load_Status = "Unable to process summary file"
        Sheets("HOME").Activate
    End If


'Ensure just the bits with "Show Bit" flag are shown
    
    subUpdate False
    
    For i = 1 To 5
        Sheets("HOME").Activate
        If i = CInt(BitNum) Then
            Select Case strVis
                Case "Show_Bit": A = HideCutters(BitNum, False, False, False) ' Set to Show
                Case "Hide_Bit":
                    A = HideCutters(BitNum, False, False, False) ' Set to Show (it was hidden before, but user refreshed and should see the results!)
                    strVis = "Show_Bit"
                Case "Hide_Backups_Bit": A = HideCutters(BitNum, True, True, False) ' set to hide backups
                Case "Hide_Off_Backups_Bit": A = HideCutters(BitNum, True, True, True) ' set to hide off-profile backups
            End Select
            
            ActiveSheet.Shapes(strVis & i).ControlFormat.Value = 1
            
        ElseIf ActiveSheet.Shapes("Hide_Bit" & i).ControlFormat.Value = 1 Then
            ActiveSheet.Shapes("Hide_Bit" & i).ControlFormat.Value = 1
            A = HideCutters(i, False, True, False)
        End If
    Next i
    
    Sheets("HOME").Activate

    subUpdate True
    
'Record the finish "time"
    sEnd(0) = Timer
    sEnd(1) = Timer

'Status info updated for the bit
    strMsg = Bit_Load_Status
    If Bit_Warn_Status <> "" Then strMsg = strMsg & Chr(10) & Bit_Warn_Status
    strMsg = strMsg & Chr(10) & "{" & Range("ReadVer") & "} Loaded in [PostRead: " & Format(sEnd(1) - sStart(1), "0.00") & "s] [RBit: " & Format(sEnd(2) - sStart(2), "0.00") & "s] [RSumm: " & Format(sEnd(3) - sStart(3), "0.00") & "s]"
    
    Range("Bit" & BitNum & "_Status").Value = strMsg
    
    If Bit_Warn_Status <> "" And Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then
        MsgBox Bit_Warn_Status & " for Iteration " & BitNum - 1 & "." & vbCrLf & vbCrLf & "Please review SUMMARY sheet.", vbExclamation + vbOKOnly, "Warning"
    End If
    
    CoreDiamCutterUpdate CInt(BitNum)
    
    subActive False
    
End Sub
Function ProcessDes(BitNum, Bit_File_Handler)

    gProc = "ProcessDes"
    
    'Initialize data variables
    Dim radius() As Double
    Dim ang_arnd() As Double
    Dim height() As Double
    Dim prf_ang() As Double
    Dim bk_rake() As Double
    Dim sd_rake() As Double
    Dim cut_dia() As Variant
    Dim blade() As Double
    Dim size_type() As String
    Dim ProfL() As Double
    Dim exposure() As Double
    Dim Blade_Angle() As Double
    Dim Tag() As Integer
    Dim roll() As Double
    Dim hMax As Double
    Dim DeltaProfAng() As Double
    Dim Tag360() As Integer
    Dim TagCreo() As Integer
    
    On Error GoTo ErrHandler
    
    ID_Radius = "-"
    Cone_Angle = "-"
    Shoulder_Angle = "-"
    Nose_Radius = "-"
    Middle_Radius = "-"
    Shoulder_Radius = "-"
    Gauge_Length = "-"
    Profile_Height = "-"
    Nose_Location_Diameter = "-"
    
    
    'Identify the data  and store it in the variables
    With ActiveSheet.UsedRange
        Set find_BasicData = .Find("[BasicData]", LookIn:=xlValues)
        If Not find_BasicData Is Nothing Then
            Bit_Type = Range("A" + LTrim(str(find_BasicData.row + 1))).Value
            Bit_Size = RoundOff(CDbl(Range("A" + LTrim(str(find_BasicData.row + 2))).Value), 3)
            Blade_Count = Range("A" + LTrim(str(find_BasicData.row + 3))).Value
        End If
        Set find_Profile_Type = .Find("[profile_choice]", LookIn:=xlValues)
        If Not find_Profile_Type Is Nothing Then
            Bit_Profile = Range("A" + LTrim(str(find_Profile_Type.row + 1))).Value
        End If
        Set find_Profile_Info = .Find("[Profile" & Bit_Profile & "]", LookIn:=xlValues)
        If Not find_Profile_Info Is Nothing Then
            i = 1
            If Bit_Profile = "AAAG" Then
                ID_Radius = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 3)
                i = i + 1
            Else
                Cone_Angle = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 1)
                i = i + 1
            End If
            If Bit_Profile = "LALAG" Then
                Shoulder_Angle = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 1)
                i = i + 1
            End If
            If Bit_Profile = "LAAG" Or Bit_Profile = "LAAAG" Or Bit_Profile = "LALAG" Or Bit_Profile = "AAAG" Then
                Nose_Radius = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 3)
                i = i + 1
            End If
            If Bit_Profile = "LAAAG" Then
                Middle_Radius = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 3)
                i = i + 1
            End If
            Shoulder_Radius = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 3)
            i = i + 1
            Gauge_Length = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 3)
            i = i + 1
            If Bit_Profile = "LAAAG" Then
                Profile_Height = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 3)
                i = i + 1
            End If
            If Bit_Profile = "LAAG" Or Bit_Profile = "LAAAG" Or Bit_Profile = "LALAG" Or Bit_Profile = "AAAG" Then
                Nose_Location_Diameter = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 3)
                i = i + 1
            End If
            Tip_Grind = RoundOff(CDbl(Range("A" + LTrim(str(find_Profile_Info.row + i))).Value), 3) - Bit_Size
            If Bit_Profile = "LAG" Then
                Nose_Radius = Shoulder_Radius
                Nose_Location_Diameter = RoundOff(((Bit_Size + Tip_Grind) / 2 - Shoulder_Radius) * 2, 3)
            End If
        End If
        Set find_actual_positions = .Find("[positions]", LookIn:=xlValues)
        If Not find_actual_positions Is Nothing Then
             Position_index = Range("A" + LTrim(str(find_actual_positions.row + 1))).Value
        End If
        Position_length_index = find_actual_positions.row + 1
        Set find_positions = .Find("[cutters]", LookIn:=xlValues)
        If Not find_positions Is Nothing Then
             Position_Count = Range("A" + LTrim(str(find_positions.row + 1))).Value
        End If
        row_index = find_positions.row + 2
        pos_index = 1
        For i = 1 To Position_Count
            Temp_ProfL = Range("C" + LTrim(str(Position_length_index + i))).Value
            row_index = row_index + 1
            For j = 1 To Blade_Count
                Cutter_exists = CInt(Range("F" + LTrim(str(row_index))).Value)
                If Cutter_exists = 1 Then
                    ReDim Preserve radius(1 To pos_index) As Double
                    ReDim Preserve ang_arnd(1 To pos_index) As Double
                    ReDim Preserve height(1 To pos_index) As Double
                    ReDim Preserve prf_ang(1 To pos_index) As Double
                    ReDim Preserve bk_rake(1 To pos_index) As Double
                    ReDim Preserve sd_rake(1 To pos_index) As Double
                    ReDim Preserve cut_dia(1 To pos_index) As Variant
                    ReDim Preserve blade(1 To pos_index) As Double
                    ReDim Preserve size_type(1 To pos_index) As String
                    ReDim Preserve ProfL(1 To pos_index) As Double
                    ReDim Preserve exposure(1 To pos_index) As Double
                    ReDim Preserve Tag(1 To pos_index) As Integer
                    ReDim Preserve roll(1 To pos_index) As Double
                    ReDim Preserve DeltaProfAng(1 To pos_index) As Double
                    ReDim Preserve Tag360(1 To pos_index) As Integer
                    ReDim Preserve TagCreo(1 To pos_index) As Integer
                    ProfL(pos_index) = Temp_ProfL
                    radius(pos_index) = Range("H" + LTrim(str(row_index))).Value
                    ang_arnd(pos_index) = Range("J" + LTrim(str(row_index))).Value
                    height(pos_index) = Range("I" + LTrim(str(row_index))).Value
                    prf_ang(pos_index) = Range("K" + LTrim(str(row_index))).Value
                    bk_rake(pos_index) = Range("L" + LTrim(str(row_index))).Value
                    sd_rake(pos_index) = Range("M" + LTrim(str(row_index))).Value
                    blade(pos_index) = Range("D" + LTrim(str(row_index))).Value + 1
                    exposure(pos_index) = Range("R" + LTrim(str(row_index))).Value
                    cut_dia(pos_index) = CutterDia(Range("A" + LTrim(str(row_index + 1))).Value)
                    size_type(pos_index) = Range("A" + LTrim(str(row_index + 1))).Value
                    Tag(pos_index) = Range("C" + LTrim(str(row_index + 1))).Value
                    DeltaProfAng(pos_index) = Range("S" + LTrim(str(row_index))).Value
                    Tag360(pos_index) = Range("D" + LTrim(str(row_index + 1))).Value
                    TagCreo(pos_index) = Range("B" + LTrim(str(row_index + 1))).Value
                    If Range("E" + LTrim(str(row_index + 1))) <> "" Then
                        roll(pos_index) = Range("E" + LTrim(str(row_index + 1)))
                    Else
                        roll(pos_index) = 0
                    End If
                    pos_index = pos_index + 1
                End If
                row_index = row_index + 2
            Next j
        Next i

        GetGaugePadInfo CInt(BitNum), CSng(Bit_Size)

    End With
    ReDim Preserve Blade_Angle(1 To Blade_Count) As Double

    'Write the INFO page
    Sheets("INFO").Activate
    row_index = Range("Cutting_Structure_Info").row
    row_index = row_index + 2 + CInt(BitNum - 1) * 4
    Range(fXLNumberToLetter(Range("CS_Bit_Type").Column) & Trim(str(row_index))).Value = Bit_Type
    Range(fXLNumberToLetter(Range("CS_Bit_Size").Column) & Trim(str(row_index))).Value = Bit_Size
    Range(fXLNumberToLetter(Range("CS_Blade_Count").Column) & Trim(str(row_index))).Value = Blade_Count
    Range(fXLNumberToLetter(Range("Prof_Profile_Type").Column) & Trim(str(row_index))).Value = Bit_Profile
    Range(fXLNumberToLetter(Range("Prof_ID_Radius").Column) & Trim(str(row_index))).Value = ID_Radius
    Range(fXLNumberToLetter(Range("Prof_Cone_Ang").Column) & Trim(str(row_index))).Value = Cone_Angle
    Range(fXLNumberToLetter(Range("Prof_Shou_Ang").Column) & Trim(str(row_index))).Value = Shoulder_Angle
    Range(fXLNumberToLetter(Range("Prof_Nose_Rad").Column) & Trim(str(row_index))).Value = Nose_Radius
    Range(fXLNumberToLetter(Range("Prof_Mid_Rad").Column) & Trim(str(row_index))).Value = Middle_Radius
    Range(fXLNumberToLetter(Range("Prof_Shou_Rad").Column) & Trim(str(row_index))).Value = Shoulder_Radius
    Range(fXLNumberToLetter(Range("Prof_Gauge_Len").Column) & Trim(str(row_index))).Value = Gauge_Length
    Range(fXLNumberToLetter(Range("Prof_Prof_Ht").Column) & Trim(str(row_index))).Value = Profile_Height
    Range(fXLNumberToLetter(Range("Prof_Nose_Loc").Column) & Trim(str(row_index))).Value = Nose_Location_Diameter
    Range(fXLNumberToLetter(Range("Prof_Tip_Grind").Column) & Trim(str(row_index))).Value = Tip_Grind
    
    
    A = ProcessCutterInfo(BitNum, radius, ang_arnd, height, prf_ang, bk_rake, sd_rake, cut_dia, size_type, blade, exposure, ProfL, Tag, roll, hMax, DeltaProfAng, Tag360, TagCreo)
    
    If Range("Y" & Trim(str(row_index))).Value = "-" Then
        Application.Range("Y" & Trim(str(row_index))).Value = RoundOff(hMax, 3)
    End If
    
    Sheets("HOME").Activate
    
    Exit Function
    
ErrHandler:

Debug.Print Now(), Err, Err.Description

If Err = 13 Then
    Resume Next
Else
    Stop
    Resume
End If

    
End Function
Function ProcessCfb(BitNum, Bit_File_Handler)

    gProc = "ProcessCfb"
    
    Dim radius() As Double
    Dim ang_arnd() As Double
    Dim height() As Double
    Dim prf_ang() As Double
    Dim bk_rake() As Double
    Dim sd_rake() As Double
    Dim cut_dia() As Double
    Dim blade() As Double
    Dim size_type() As String
    Dim ProfL() As Double
    Dim exposure() As Double
    Dim Blade_Angle() As Double
    Dim Tag() As Integer
    Dim roll() As Double
    Dim hMax As Double
    Dim DeltaProfAng() As Double
    Dim Tag360() As Integer
    Dim TagCreo() As Integer
    Dim iColOffset As Integer
    Dim iRow As Integer, iCol As Integer
    
    On Error GoTo ErrHandler
    
    ID_Radius = "-"
    Cone_Angle = "-"
    Shoulder_Angle = "-"
    Nose_Radius = "-"
    Middle_Radius = "-"
    Shoulder_Radius = "-"
    Gauge_Length = "-"
    Profile_Height = "-"
    Nose_Location_Diameter = "-"
    
'find first column of data
    iColOffset = 0
    
    Do While Cells(3, iColOffset + 1) = ""
        iColOffset = iColOffset + 1
    Loop
    
    If Cells.Find(What:="[expan", After:=Range("A1"), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row > 0 Then
        'ExpandedTopFound
        LastRow = Cells.Find(What:="[expan", After:=Range("A1"), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row - 2
    Else
ExpandedTopNotFound:
        LastRow = Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    End If
    
' Test to make sure all columns are aligned
    iRow = 3
    iCol = 1
    Do While iRow <= LastRow
        If Cells(iRow, iColOffset + 1) = iRow - 2 Then
            ' next iRow
        Else
            If iColOffset - iCol >= 0 Then
                If Cells(iRow, iColOffset - iCol + 1) = iRow - 2 Then
                    Range("A" & iRow & ":" & fXLNumberToLetter(CLng(iCol)) & iRow).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                Else
                    iCol = iCol + 1
                End If
            Else
                
            End If
        End If
        iRow = iRow + 1
    Loop


    Cutter_Count = LastRow - 2
    ReDim Preserve radius(1 To Cutter_Count) As Double
    ReDim Preserve ang_arnd(1 To Cutter_Count) As Double
    ReDim Preserve height(1 To Cutter_Count) As Double
    ReDim Preserve prf_ang(1 To Cutter_Count) As Double
    ReDim Preserve bk_rake(1 To Cutter_Count) As Double
    ReDim Preserve sd_rake(1 To Cutter_Count) As Double
    ReDim Preserve cut_dia(1 To Cutter_Count) As Double
    ReDim Preserve blade(1 To Cutter_Count) As Double
    ReDim Preserve size_type(1 To Cutter_Count) As String
    ReDim Preserve ProfL(1 To Cutter_Count) As Double
    ReDim Preserve exposure(1 To Cutter_Count) As Double
    ReDim Preserve Tag(1 To Cutter_Count) As Integer
    ReDim Preserve roll(1 To Cutter_Count) As Double
    ReDim Preserve DeltaProfAng(1 To Cutter_Count) As Double
    ReDim Preserve Tag360(1 To Cutter_Count) As Integer
    ReDim Preserve TagCreo(1 To Cutter_Count) As Integer
    
    row_index = 2
    For i = 1 To Cutter_Count
        row_index = row_index + 1
        
'        radius(i) = Range("C" + LTrim(Str(row_index))).Value
'        ang_arnd(i) = Range("D" + LTrim(Str(row_index))).Value
'        height(i) = Range("E" + LTrim(Str(row_index))).Value
'        prf_ang(i) = Range("F" + LTrim(Str(row_index))).Value
'        bk_rake(i) = Range("G" + LTrim(Str(row_index))).Value
'        sd_rake(i) = -Range("H" + LTrim(Str(row_index))).Value
'        cut_dia(i) = CutterDia(Range("I" + LTrim(Str(row_index))).Value)
'        size_type(i) = Range("J" + LTrim(Str(row_index))).Value
'        blade(i) = Range("K" + LTrim(Str(row_index))).Value
        
        radius(i) = Cells(LTrim(str(row_index)), 2 + iColOffset).Value
        ang_arnd(i) = Cells(LTrim(str(row_index)), 3 + iColOffset).Value
        height(i) = Cells(LTrim(str(row_index)), 4 + iColOffset).Value
        prf_ang(i) = Cells(LTrim(str(row_index)), 5 + iColOffset).Value
        bk_rake(i) = Cells(LTrim(str(row_index)), 6 + iColOffset).Value
        sd_rake(i) = -Cells(LTrim(str(row_index)), 7 + iColOffset).Value
        cut_dia(i) = CutterDia(Cells(LTrim(str(row_index)), 8 + iColOffset).Value)
        
        If InStr(1, Cells(LTrim(str(row_index)), 9 + iColOffset).Value, ":", vbTextCompare) > 0 Then
            Cells(LTrim(str(row_index)), 9 + iColOffset).Value = Right(Cells(LTrim(str(row_index)), 9 + iColOffset).Value, Len(Cells(LTrim(str(row_index)), 9 + iColOffset).Value) - InStr(1, Cells(LTrim(str(row_index)), 9 + iColOffset).Value, ":", vbTextCompare))
        End If
        
        size_type(i) = Cells(LTrim(str(row_index)), 9 + iColOffset).Value
        blade(i) = Cells(LTrim(str(row_index)), 10 + iColOffset).Value
        
        If Cells(LTrim(str(row_index)), 11 + iColOffset).Value = "" Then
            roll(i) = 0
        Else
            roll(i) = Cells(LTrim(str(row_index)), 11 + iColOffset).Value
        End If
        
        ProfL(i) = 0
        exposure(i) = 0
        Tag(i) = 0
    Next i
    'Max_Height = WorksheetFunction.Max(Range(ActiveCell, ActiveCell.Offset(Cutter_Count - 1, 0)))
    'ReDim Preserve Blade_Angle(1 To Blade_Count) As Double
    'For i = 1 To Blade_Count
    '    j = Position_Count
    '    Do Until blade(j) = i
    '        j = j - 1
    '    Loop
    '    Blade_Angle(i) = ang_arnd(j)
    'Next i
    
    
    On Error GoTo ErrHandler:
    A = ProcessCutterInfo(BitNum, radius, ang_arnd, height, prf_ang, bk_rake, sd_rake, cut_dia, size_type, blade, exposure, ProfL, Tag, roll, hMax, DeltaProfAng, Tag360, TagCreo)
    B = GetProfileInfo(BitNum)
    
'    Stop
    If Not IsEmpty(A) Then ProcessCfb = A
    If Not IsEmpty(B) Then ProcessCfb = ProcessCfb & B
    
    
    
    'Stop 'investigate gauge pad info
    GetGaugePadInfo CInt(BitNum), 0
    
    Exit Function
    
ErrHandler:

    If Err = 91 Then
        'No Expanded Top section found in CFB
        Resume ExpandedTopNotFound
    Else
        Debug.Print Now(), gProc, Err, Err.Description
        
        Sheets("HOME").Activate
    End If

End Function
Function GetProfileInfo(BitNum)

    gProc = "GetProfileInfo"
    
    Dim fs As scripting.FileSystemObject
    Dim txtIn As scripting.TextStream ' scripting.textstream
    Dim strLine As String, strAtts() As String

    On Error GoTo ErrHandler

    Set fs = New scripting.FileSystemObject
    Set txtIn = fs.OpenTextFile(Right(Range("Bit" & BitNum & "_Location").Value, Len(Range("Bit" & BitNum & "_Location").Value) - InStr(1, Range("Bit" & BitNum & "_Location").Value, "]") - 1), 1) ' 1 ForReading
    strLine = txtIn.ReadLine
    txtIn.Close
    Set txtIn = Nothing
    Set fs = Nothing

    Blade_Count = "-"
    ID_Radius = "-"
    Cone_Angle = "-"
    Shoulder_Angle = "-"
    Nose_Radius = "-"
    Middle_Radius = "-"
    Shoulder_Radius = "-"
    Gauge_Length = "-"
    Profile_Height = "-"
    Nose_Location_Diameter = "-"
    Tip_Grind = "-"
    Bit_Size = "-"

    If Dir(Mid(Range("Bit" & BitNum & "_Location").Value, InStr(1, Range("Bit" & BitNum & "_Location").Value, "]") + 2, InStrRev(Range("Bit" & BitNum & "_Location").Value, "\") - InStr(1, Range("Bit" & BitNum & "_Location").Value, "]") - 1) & "\bit.in", vbNormal) <> "" Then
        Bit_Size = CSng(CheckBitIn(Mid(Range("Bit" & BitNum & "_Location").Value, InStr(1, Range("Bit" & BitNum & "_Location").Value, "]") + 2, InStrRev(Range("Bit" & BitNum & "_Location").Value, "\") - InStr(1, Range("Bit" & BitNum & "_Location").Value, "]") - 1) & "\bit.in", "Diameter"))
    End If
    
    If InStr(1, strLine, "n_u_blades") > 0 Then
        strLine = Right(strLine, Len(strLine) - InStr(1, strLine, "n_u_blades") - 10)
        
        Debug.Print strLine
        
        strAtts = Split(strLine, " ")

        Select Case strAtts(1)
            Case "AAAG":
            ' Example: 6 AAAG 2.000000 1.720000 2.900000 0.750000 5.270000 8.790000
                Blade_Count = RoundOff(strAtts(0), 3)
                Bit_Profile = strAtts(1)
                ID_Radius = RoundOff(strAtts(2), 3)
                Nose_Radius = RoundOff(strAtts(3), 3)
                Shoulder_Radius = RoundOff(strAtts(4), 3)
                Gauge_Length = RoundOff(strAtts(5), 3)
                Nose_Location_Diameter = RoundOff(strAtts(6), 3)
                If Bit_Size <> "-" Then
                    Tip_Grind = RoundOff(strAtts(7) - Bit_Size, 3)
                Else
                    Bit_Size = RoundOff(strAtts(7), 3)
                End If
                
            Case "LAG":
            ' Example: 6 LAG 75.000000 1.900000 1.500000 8.520000
                Blade_Count = RoundOff(strAtts(0), 3)
                Bit_Profile = strAtts(1)
                Cone_Angle = RoundOff(strAtts(2), 3)
                Nose_Radius = RoundOff(strAtts(3), 3)
                Shoulder_Radius = RoundOff(strAtts(3), 3)
                Gauge_Length = RoundOff(strAtts(4), 3)
                Nose_Location_Diameter = RoundOff(strAtts(5) - strAtts(3), 3)
                If Bit_Size <> "-" Then
                    Tip_Grind = RoundOff(strAtts(5) - Bit_Size, 3)
                Else
                    Bit_Size = RoundOff(strAtts(5), 3)
                End If

            Case "LAAG":
            ' Example: 6 LAAG 80.000000 1.095000 2.250000 0.527000 3.552000 6.190000
                Blade_Count = RoundOff(strAtts(0), 3)
                Bit_Profile = strAtts(1)
                Cone_Angle = RoundOff(strAtts(2), 3)
                Nose_Radius = RoundOff(strAtts(3), 3)
                Shoulder_Radius = RoundOff(strAtts(4), 3)
                Gauge_Length = RoundOff(strAtts(5), 3)
                Nose_Location_Diameter = RoundOff(strAtts(6), 3)
                If Bit_Size <> "-" Then
                    Tip_Grind = RoundOff(strAtts(7) - Bit_Size, 3)
                Else
                    Bit_Size = RoundOff(strAtts(7), 3)
                End If
                
            Case "LAAAG":
            ' Example: 6 LAAAG 67.500000 2.500000 4.000000 2.100000 1.301000 2.200000 7.500000 12.270000
                Blade_Count = RoundOff(strAtts(0), 3)
                Bit_Profile = strAtts(1)
                Cone_Angle = RoundOff(strAtts(2), 3)
                Nose_Radius = RoundOff(strAtts(3), 3)
                Middle_Radius = RoundOff(strAtts(4), 3)
                Shoulder_Radius = RoundOff(strAtts(5), 3)
                Gauge_Length = RoundOff(strAtts(6), 3)
                Profile_Height = RoundOff(strAtts(7), 3)
                Nose_Location_Diameter = RoundOff(strAtts(8), 3)
                If Bit_Size <> "-" Then
                    Tip_Grind = RoundOff(strAtts(9) - Bit_Size, 3)
                Else
                    Bit_Size = RoundOff(strAtts(9), 3)
                End If
                
            Case "LALAG":
            ' Example: 6 LALAG 72.000000 60.000000 2.000000 2.000000 1.458300 4.350000 8.835000
                Blade_Count = RoundOff(strAtts(0), 3)
                Bit_Profile = strAtts(1)
                Cone_Angle = RoundOff(strAtts(2), 3)
                Shoulder_Angle = RoundOff(strAtts(3), 3)
                Nose_Radius = RoundOff(strAtts(4), 3)
                Shoulder_Radius = RoundOff(strAtts(5), 3)
                Gauge_Length = RoundOff(strAtts(6), 3)
                Nose_Location_Diameter = RoundOff(strAtts(7), 3)
                If Bit_Size <> "-" Then
                    Tip_Grind = RoundOff(strAtts(8) - Bit_Size, 3)
                Else
                    Bit_Size = RoundOff(strAtts(8), 3)
                End If
            Case Else: Bit_Profile = "Custom"
        End Select
    Else
    'Estimate Nose Location Diameter
        Sheets("BIT" & BitNum).Activate

        'find last negative profile angle
        Pi = 4 * Atn(1)

        iRowB = 5
        Do While Cells(iRowB, Range("Bit" & BitNum & "_Profile_Angle").Column) < 0
            iRowB = iRowB + 1
        Loop
        
        If Cells(iRowB, Range("Bit" & BitNum & "_Profile_Angle").Column) = 0 Then
            Nose_Location_Diameter = Cells(iRow, Range("Bit" & BitNum & "_Radius").Column) * 2
            iRowB = iRowB + 1
        End If
        
        iRowA = iRowB - 1
        
        sRadA = Cells(iRowA, Range("Bit" & BitNum & "_Radius").Column)
        sRadB = Cells(iRowB, Range("Bit" & BitNum & "_Radius").Column)
        
        sAngA = Abs(Cells(iRowA, Range("Bit" & BitNum & "_Profile_Angle").Column))
        sAngB = Cells(iRowB, Range("Bit" & BitNum & "_Profile_Angle").Column)
        
        X = sRadB - sRadA
        sA = X / ((Sin(sAngB * Pi / 180) / Sin(sAngA * Pi / 180)) + 1)

        Nose_Location_Diameter = RoundOff((sRadA + sA) * 2, 3)
        Nose_Radius = RoundOff(sA / Sin(sAngA * Pi / 180), 3)
        
        sDia = Range("Bit" & BitNum & "_Radius").Item(Range("Bit" & BitNum & "_Radius").Count) * 2
        If (sDia - 0.085) * 8 = CInt((sDia - 0.085) * 8) Then
            Tip_Grind = 0.085
        ElseIf (sDia - 0.065) * 8 = CInt((sDia - 0.065) * 8) Then
            Tip_Grind = 0.065
        ElseIf (sDia - 0.04) * 8 = CInt((sDia - 0.04) * 8) Then
            Tip_Grind = 0.04
        Else
            Tip_Grind = 0
        End If
        Bit_Size = sDia - Tip_Grind
        
        
        
        Blade_Count = WorksheetFunction.Max(Range("BIT" & BitNum & "_Blade_Num"))
        
        Sheets("BIT" & BitNum).Activate
    End If
    
    Sheets("BIT" & BitNum).Activate
    Profile_Height = WorksheetFunction.Max(Range("BIT" & BitNum & "_Height3")) - WorksheetFunction.Min(Range("BIT" & BitNum & "_Height3"))
    
    Sheets("INFO").Activate
    row_index = Range("Cutting_Structure_Info").row
    row_index = row_index + 2 + CInt(BitNum - 1) * 4
    
'    Range("C" & Trim(str(row_index))).Value = Bit_Type
'    Range("E" & Trim(str(row_index))).Value = Bit_Size
'    Range("F" & Trim(str(row_index))).Value = Blade_Count
'    Range("Q" & Trim(str(row_index))).Value = Bit_Profile
'    Range("R" & Trim(str(row_index))).Value = ID_Radius
'    Range("S" & Trim(str(row_index))).Value = Cone_Angle
'    Range("T" & Trim(str(row_index))).Value = Shoulder_Angle
'    Range("U" & Trim(str(row_index))).Value = Nose_Radius
'    Range("V" & Trim(str(row_index))).Value = Middle_Radius
'    Range("W" & Trim(str(row_index))).Value = Shoulder_Radius
'    Range("X" & Trim(str(row_index))).Value = Gauge_Length
'    Range("Y" & Trim(str(row_index))).Value = Profile_Height
'    Range("Z" & Trim(str(row_index))).Value = Nose_Location_Diameter
'    Range("AA" & Trim(str(row_index))).Value = Tip_Grind

    Range(fXLNumberToLetter(Range("CS_Bit_Type").Column) & Trim(str(row_index))).Value = Bit_Type
    Range(fXLNumberToLetter(Range("CS_Bit_Size").Column) & Trim(str(row_index))).Value = Bit_Size
    Range(fXLNumberToLetter(Range("CS_Blade_Count").Column) & Trim(str(row_index))).Value = Blade_Count
    Range(fXLNumberToLetter(Range("Prof_Profile_Type").Column) & Trim(str(row_index))).Value = Bit_Profile
    Range(fXLNumberToLetter(Range("Prof_ID_Radius").Column) & Trim(str(row_index))).Value = ID_Radius
    Range(fXLNumberToLetter(Range("Prof_Cone_Ang").Column) & Trim(str(row_index))).Value = Cone_Angle
    Range(fXLNumberToLetter(Range("Prof_Shou_Ang").Column) & Trim(str(row_index))).Value = Shoulder_Angle
    Range(fXLNumberToLetter(Range("Prof_Nose_Rad").Column) & Trim(str(row_index))).Value = Nose_Radius
    Range(fXLNumberToLetter(Range("Prof_Mid_Rad").Column) & Trim(str(row_index))).Value = Middle_Radius
    Range(fXLNumberToLetter(Range("Prof_Shou_Rad").Column) & Trim(str(row_index))).Value = Shoulder_Radius
    Range(fXLNumberToLetter(Range("Prof_Gauge_Len").Column) & Trim(str(row_index))).Value = Gauge_Length
    Range(fXLNumberToLetter(Range("Prof_Prof_Ht").Column) & Trim(str(row_index))).Value = Profile_Height
    Range(fXLNumberToLetter(Range("Prof_Nose_Loc").Column) & Trim(str(row_index))).Value = Nose_Location_Diameter
    Range(fXLNumberToLetter(Range("Prof_Tip_Grind").Column) & Trim(str(row_index))).Value = Tip_Grind

    Exit Function
    
ErrHandler:

Debug.Print "ERROR:", Now(), Err, Err.Description
Stop
Resume

End Function

Function ProcessCutterInfo(BitNum, radius, ang_arnd, height, prf_ang, bk_rake, sd_rake, cut_dia, size_type, blade, exposure, ProfL, Tag, roll, hMax, DeltaProfAng, Tag360, TagCreo)

    gProc = "ProcessCutterInfo"
    
    Dim iRowLastInner As Integer, strLastCol As String, iStartConerow As Integer, iColTag As Integer, iTemp As Integer, i As Integer
    Dim blnHidden As Boolean
    On Error GoTo ErrHandler
    
    Pi = 4 * Atn(1)
    
    gblnCentSting = False
    
    total_cutter_count = UBound(radius)
    Sheets("BIT" & BitNum).Activate
    
    
    start_row_index = 4
    row_index = start_row_index + 1
    For i = 1 To total_cutter_count
        Range("A" + LTrim(str(row_index))).Value = i
        Range("B" + LTrim(str(row_index))).Value = RoundOff(CDbl(radius(i)), 4)
        Range("C" + LTrim(str(row_index))).Value = ang_arnd(i)
        Range("D" + LTrim(str(row_index))).Value = RoundOff(CDbl(height(i)), 4)
        Range("E" + LTrim(str(row_index))).Value = RoundOff(CDbl(prf_ang(i)), 2)
        If Range("E" + LTrim(str(row_index))).Value <= Range("ForceRatioSplitInc") Then iRowLastInner = row_index
        Range("F" + LTrim(str(row_index))).Value = RoundOff(CDbl(bk_rake(i)), 2)
        Range("G" + LTrim(str(row_index))).Value = RoundOff(CDbl(sd_rake(i)), 2)
        Range("H" + LTrim(str(row_index))).Value = cut_dia(i)
        Range("I" + LTrim(str(row_index))).Value = size_type(i)
        Range("J" + LTrim(str(row_index))).Value = blade(i)
        Range("K" + LTrim(str(row_index))).Value = RoundOff(CDbl(exposure(i)), 4)
        Range("L" + LTrim(str(row_index))).Value = RoundOff(CDbl(ProfL(i)), 3)
        Range("P" + LTrim(str(row_index))).Value = CInt(Tag(i))
        Range("Q" + LTrim(str(row_index))).Value = RoundOff(roll(i), 2)
        Range("R" + LTrim(str(row_index))).Value = RoundOff(DeltaProfAng(i), 4)
        Range("S" + LTrim(str(row_index))).Value = CInt(Tag360(i))
        Range("T" + LTrim(str(row_index))).Value = CInt(TagCreo(i))
        
        If i = 1 Then
            Range("M" + LTrim(str(row_index))).Value = RoundOff(CDbl(ProfL(i)), 3)
        Else
            Range("M" + LTrim(str(row_index))).Value = RoundOff(CDbl(ProfL(i) - ProfL(i - 1)), 3)
        End If
        If ProfL(i) = 0 Then   'Adding Cutter Distances if CFB is used
            If i = 1 Then
                Range("M" + LTrim(str(row_index))).Value = RoundOff(CDbl(Range("B" + LTrim(str(row_index))).Value / Cos(Range("E" + LTrim(str(row_index))).Value * Pi / 180)), 3)
                Range("L" + LTrim(str(row_index))).Value = Range("M" + LTrim(str(row_index))).Value
            ElseIf i > 1 Then
'                Range("M" + LTrim(str(row_index))).Value = Format(((Range("B" + LTrim(str(row_index))).Value - Range("B" + LTrim(str(row_index - 1))).Value) ^ 2 + (Range("D" + LTrim(str(row_index))).Value - Range("D" + LTrim(str(row_index - 1))).Value) ^ 2) ^ 0.5, "0.00")
                Range("M" + LTrim(str(row_index))).Value = RoundOff(((Range("B" + LTrim(str(row_index))).Value - Range("B" + LTrim(str(row_index - 1))).Value) ^ 2 + (Range("D" + LTrim(str(row_index))).Value - Range("D" + LTrim(str(row_index - 1))).Value) ^ 2) ^ 0.5, 3)
                ProcessCutterInfo = "Cutter Distances Approximated"
                Range("L" + LTrim(str(row_index))).Value = Range("M" + LTrim(str(row_index))).Value + Range("L" + LTrim(str(row_index - 1))).Value
            End If
        End If
'        Stop
        row_index = row_index + 1
    Next i


' Process the rest of the cutters info
    Dim Heading_Array As Variant
    Heading_Array = Array("Cutter_Num", "Radius", "Angle_Around", "Height", "Profile_Angle", "Back_Rake", "Side_Rake", "Cutter_Diameter", "Cutter_Type", "Blade_Num", "Exposure", "Profile_Position", "Cutter_Distance", "Height2", "Height3", "Tag", "Roll", "DeltaProfAng", "E360Tag", "CreoTag")
    Range("D" + LTrim(str(start_row_index + 1))).Select
    Max_Height = WorksheetFunction.Max(Range(ActiveCell, ActiveCell.Offset(total_cutter_count - 1, 0)))
    For i = 0 To UBound(Heading_Array)
        Cells(start_row_index, i + 1).Value = Heading_Array(i)
        Cells(start_row_index + 1, i + 1).Select
        Range(ActiveCell, ActiveCell.Offset(total_cutter_count - 1, 0)).Name = "Bit" & BitNum & "_" & Heading_Array(i)
    Next i
    i = 1
    Gauge_zero_bit_height = 0
    Do Until RoundOff(CDbl(prf_ang(i)), 2) >= 90 Or i >= total_cutter_count
        i = i + 1
    Loop
    Gauge_zero_bit_height = height(i)
    'MsgBox Gauge_zero_bit_height, vbInformation
    For i = 1 To total_cutter_count
        Cells(i + start_row_index, 14).Value = RoundOff(CDbl(height(i) - Max_Height), 4)
        Cells(i + start_row_index, 15).Value = RoundOff(CDbl(height(i) - Gauge_zero_bit_height), 4)
    Next i
    
    hMax = Application.WorksheetFunction.Max(Application.Range("Bit" & BitNum & "_Height3"))
    
    Dim Cutter_Type(1 To 4) As String
    Dim Cutter_Count(1 To 4) As String
    Dim BR_Low(1 To 4) As Double
    Dim BR_High(1 To 4) As Double
    Dim SR_Low(1 To 4) As Double
    Dim SR_High(1 To 4) As Double
    
    
'Adding new columns
    strLastCol = AddRadiusRatio(CInt(BitNum), CInt(total_cutter_count))
    strLastCol = AddIdentity(CInt(BitNum), CInt(total_cutter_count))
    strLastCol = AddXAxis(CInt(BitNum), CInt(total_cutter_count))
    strLastCol = AddProfileLocation(CInt(BitNum), CInt(total_cutter_count))
'    strLastCol = AddProfZone(CInt(BitNum), CInt(Total_Cutter_Count))

    Cells(4, 1).Select
    Range(Selection, Selection.End(xlEnd)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Last_Cell = Selection.Address
    
'    Last_Cell = ActiveSheet.UsedRange.Address
    Last_row = Range(Last_Cell).Rows.Count
    Last_Col = Range(Last_Cell).Columns.Count
    row_index = start_row_index + 1
    row_count = 0
    
    iColTag = Range("Bit" & BitNum & "_Tag").Column
    
    
    
    
'Getting # of cone cutters (based off of equal cone profile angles
    iStartConerow = row_index
'    Do While IsStinger(Cells(iStartConerow, Range("Bit" & BitNum & "_Cutter_Type").Column)) And Cells(row_index, Range("Bit" & BitNum & "_Radius").Column) < 0.005
'Determine if there's a central stinger
    Do While IsStinger(Cells(iStartConerow, Range("Bit" & BitNum & "_Cutter_Type").Column)) And Cells(iStartConerow, Range("Bit" & BitNum & "_Radius").Column) < 0.005
        iStartConerow = iStartConerow + 1
        row_count = row_count + 1
        Range("CentStingPresent") = True
    Loop
    
'Determine if a Stinger is cutting the core
    If Range("CentStingPresent") = True Then
    'Typically when a stinger cuts the core, it's profile angle is at -47deg
        If IsStinger(Cells(iStartConerow, Range("Bit" & BitNum & "_Cutter_Type").Column)) And Cells(iStartConerow, Range("Bit" & BitNum & "_Profile_Angle").Column) < -45 Then
            iStartConerow = iStartConerow + 1
            row_count = row_count + 1
        End If
    End If
    
    row_index = iStartConerow
'Determine if the profile is AAAG (first A cutters will have gradually increasing negative profile angles. These will be considered "Cone" cutters. When the profile angles start to decrease, the tangent where the 2nd A comes in has been passed. This is the start of the "Nose".
    If Range("Bit" & BitNum & "_Type") = "AAAG" Then
        Do While Range("E" + LTrim(str(row_index))).Value < Range("E" + LTrim(str(row_index - 1))).Value
            row_index = row_index + 1
            row_count = row_count + 1
        Loop
    Else
    '    Do Until Not Range("E" + LTrim(str(row_index))).Value = Range("E" + LTrim(str(start_row_index + 1))).Value

'Stop ' Test to make sure the stingers in the cone are not affecting "cone area" counts/regions
        Do
            
            If IsStinger(Cells(LTrim(str(row_index + 1)), Range("Bit" & BitNum & "_Cutter_Type").Column).Value) And Cells(LTrim(str(row_index + 1)), Range("Bit" & BitNum & "_Tag").Column).Value > 0 Then
                row_index = row_index + 1
                row_count = row_count + 1
            End If
            
            row_index = row_index + 1
            row_count = row_count + 1
        Loop Until Not Range("E" + LTrim(str(row_index))).Value = Range("E" + LTrim(str(iStartConerow))).Value
    End If
    
    Cutter_Count(1) = row_index - start_row_index - 1
    
    Cells(start_row_index + 1, 1).Select
'    Cells(iStartConerow, 1).Select
    Range(ActiveCell, ActiveCell.Offset(row_count - 1, Last_Col - 1)).Interior.ColorIndex = 35
    Range(ActiveCell.Offset(0, iColTag - 1), ActiveCell.Offset(row_count - 1, iColTag - 1)).Name = "Bit" & BitNum & "_Cone"
    
    For i = 1 To Range("Bit" & BitNum & "_Cone").Count
        Range("Bit" & BitNum & "_ProfileLocation").Item(i) = "C"
    Next i
    
    
    Cutter_Type(1) = FindCutters(start_row_index + 1, row_index - 1)
    Range("F" + LTrim(str(start_row_index + 1))).Select
    BR_Low(1) = WorksheetFunction.Min(Range(ActiveCell.Offset(iStartConerow - (start_row_index + 1), 0), ActiveCell.Offset(row_count - 1, 0)))
    BR_High(1) = WorksheetFunction.Max(Range(ActiveCell.Offset(iStartConerow - (start_row_index + 1), 0), ActiveCell.Offset(row_count - 1, 0)))
    Range("G" + LTrim(str(start_row_index + 1))).Select
    SR_Low(1) = WorksheetFunction.Min(Range(ActiveCell.Offset(iStartConerow - (start_row_index + 1), 0), ActiveCell.Offset(row_count - 1, 0)))
    SR_High(1) = WorksheetFunction.Max(Range(ActiveCell.Offset(iStartConerow - (start_row_index + 1), 0), ActiveCell.Offset(row_count - 1, 0)))
    pre_row_index = row_index
    Cells(row_index, 1).Select
    
    
    
    
' Getting Nose cutters
'    If Not Range("E" + LTrim(str(start_row_index + 1))).Value = 0 Then
    If Not Range("E" + LTrim(str(iStartConerow))).Value = 0 Then
        row_count = 0
        Do Until Range("E" + LTrim(str(row_index))).Value >= (Range("E" + LTrim(str(iStartConerow))).Value) * (-1)
            row_index = row_index + 1
            row_count = row_count + 1
        Loop
        
        If row_count = 0 Then
            'no cone cutters in design
'            Stop
            GoTo NoNoseCutters
        Else
            Cutter_Count(2) = row_index - pre_row_index
            Range(ActiveCell, ActiveCell.Offset(row_count - 1, Last_Col - 1)).Interior.ColorIndex = 36
            Range(ActiveCell.Offset(0, iColTag - 1), ActiveCell.Offset(row_count - 1, iColTag - 1)).Name = "Bit" & BitNum & "_Nose"
            
            Cutter_Type(2) = FindCutters(pre_row_index, row_index - 1)
            Range("F" & pre_row_index).Select
            BR_Low(2) = WorksheetFunction.Min(Range(ActiveCell, ActiveCell.Offset(row_count - 1, 0)))
            BR_High(2) = WorksheetFunction.Max(Range(ActiveCell, ActiveCell.Offset(row_count - 1, 0)))
            Range("G" & pre_row_index).Select
            SR_Low(2) = WorksheetFunction.Min(Range(ActiveCell, ActiveCell.Offset(row_count - 1, 0)))
            SR_High(2) = WorksheetFunction.Max(Range(ActiveCell, ActiveCell.Offset(row_count - 1, 0)))
            pre_row_index = row_index
            Cells(row_index, 1).Select
        End If
    Else
NoNoseCutters:
        Cutter_Type(2) = "-"
        Cutter_Count(2) = 0
        BR_Low(2) = 0
        BR_High(2) = 0
        SR_Low(2) = 0
        SR_High(2) = 0
        Range(Cells(1, iColTag), Cells(1, iColTag)).Name = "Bit" & BitNum & "_Nose"
    End If

    For i = 1 To Range("Bit" & BitNum & "_Nose").Count
        Range("Bit" & BitNum & "_ProfileLocation").Item(Range("Bit" & BitNum & "_Cone").Count + i) = "N"
    Next i
    
' getting Shoulder cutters
    row_count = 0
    Do Until Range("E" + LTrim(str(row_index))).Value >= 90 Or row_index = Last_row + start_row_index
        row_index = row_index + 1
        row_count = row_count + 1
    Loop
    Cutter_Count(3) = row_index - pre_row_index
    Range(ActiveCell, ActiveCell.Offset(row_count - 1, Last_Col - 1)).Interior.ColorIndex = 37
    Range(ActiveCell.Offset(0, iColTag - 1), ActiveCell.Offset(row_count - 1, iColTag - 1)).Name = "Bit" & BitNum & "_Shoulder"
    
    For i = 1 To Range("Bit" & BitNum & "_Shoulder").Count
        Range("Bit" & BitNum & "_ProfileLocation").Item(Range("Bit" & BitNum & "_Cone").Count + Range("Bit" & BitNum & "_Nose").Count + i) = "S"
    Next i
    
    Cutter_Type(3) = FindCutters(pre_row_index, row_index - 1)
    Range("F" & pre_row_index).Select
    BR_Low(3) = WorksheetFunction.Min(Range(ActiveCell, ActiveCell.Offset(row_count - 1, 0)))
    BR_High(3) = WorksheetFunction.Max(Range(ActiveCell, ActiveCell.Offset(row_count - 1, 0)))
    Range("G" & pre_row_index).Select
    SR_Low(3) = WorksheetFunction.Min(Range(ActiveCell, ActiveCell.Offset(row_count - 1, 0)))
    SR_High(3) = WorksheetFunction.Max(Range(ActiveCell, ActiveCell.Offset(row_count - 1, 0)))
    If row_index = Last_row + start_row_index Then
        Cutter_Type(4) = "-"
        Cutter_Count(4) = 0
        BR_Low(4) = 0
        BR_High(4) = 0
        SR_Low(4) = 0
        SR_High(4) = 0
    Else
        pre_row_index = row_index
        Cells(row_index, 1).Select
        Cutter_Count(4) = Last_row - pre_row_index + start_row_index
        Range(ActiveCell, ActiveCell.Offset(Last_row - row_index + start_row_index - 1, Last_Col - 1)).Interior.ColorIndex = 40
        Range(ActiveCell.Offset(0, iColTag - 1), ActiveCell.Offset(Last_row - row_index + start_row_index - 1, iColTag - 1)).Name = "Bit" & BitNum & "_Gauge"
    
        For i = 1 To Range("Bit" & BitNum & "_Gauge").Count
            Range("Bit" & BitNum & "_ProfileLocation").Item(Range("Bit" & BitNum & "_Cone").Count + Range("Bit" & BitNum & "_Nose").Count + Range("Bit" & BitNum & "_Shoulder").Count + i) = "G"
        Next i
        
        Cutter_Type(4) = FindCutters(pre_row_index, Last_row + start_row_index - 1)
        Range("F" & pre_row_index).Select
        BR_Low(4) = WorksheetFunction.Min(Range(ActiveCell, ActiveCell.Offset(Last_row - row_index + start_row_index - 1, 0)))
        BR_High(4) = WorksheetFunction.Max(Range(ActiveCell, ActiveCell.Offset(Last_row - row_index + start_row_index - 1, 0)))
        Range("G" & pre_row_index).Select
        SR_Low(4) = WorksheetFunction.Min(Range(ActiveCell, ActiveCell.Offset(Last_row - row_index + 1, 0)))
        SR_High(4) = WorksheetFunction.Max(Range(ActiveCell, ActiveCell.Offset(Last_row - row_index + 1, 0)))
    End If
    
'    Range("A" & start_row_index & ":" & strLastCol & start_row_index).AutoFilter
    

    
' Create Range for Inner/Outer cutters
    If iRowLastInner = 0 Then iRowLastInner = row_index - 1
    Range(Cells(start_row_index + 1, Range("Bit" & BitNum & "_Profile_Angle").Column), Cells(iRowLastInner, Range("Bit" & BitNum & "_Profile_Angle").Column)).Name = "Bit" & BitNum & "_InnerCutters"
    Range(Cells(iRowLastInner + 1, Range("Bit" & BitNum & "_Profile_Angle").Column), Cells(total_cutter_count + start_row_index, Range("Bit" & BitNum & "_Profile_Angle").Column)).Name = "Bit" & BitNum & "_OuterCutters"

    
' Autofilter
    Range("A4:" & fXLNumberToLetter(CLng(Range(ActiveSheet.UsedRange.Address).Columns.Count)) & "4").AutoFilter
    
'Info for INFO2
    Dim strArea As Variant
    Dim strType() As String
    Dim iCount() As Integer
    Dim iBR() As Integer
    Dim iSR() As Integer
    
    Range(Range("Bit" & BitNum & "_I2_ConeSize"), Range("Bit" & BitNum & "_I2_GaugeSR")).ClearContents
    
    strArea = Array("Cone", "Face", "Gauge")
    
    For i = 0 To 2
        CutterCount CStr(strArea(i)), CInt(BitNum), strType, iCount, iBR, iSR
        
        'for review only
        Sheets("INFO2").Select
        
        X = 0
        Do While X <= UBound(strType)
            Range("Bit" & BitNum & "_I2_" & strArea(i) & "Size") = Range("Bit" & BitNum & "_I2_" & strArea(i) & "Size") & strType(X) & "/"
            X = X + 1
        Loop
        Range("Bit" & BitNum & "_I2_" & strArea(i) & "Size") = Left(Range("Bit" & BitNum & "_I2_" & strArea(i) & "Size"), Len(Range("Bit" & BitNum & "_I2_" & strArea(i) & "Size")) - 1)
        
        Range("Bit" & BitNum & "_I2_" & strArea(i) & "Qty") = iCount(0) & "/" & iCount(1) & "/" & iCount(2)
'        If iCount(1) <> 0 Then Range("Bit" & BitNum & "_I2_" & strArea(i) & "Qty") = Range("Bit" & BitNum & "_I2_" & strArea(i) & "Qty") & "(" & iCount(1) & ")"
        Range("Bit" & BitNum & "_I2_" & strArea(i) & "BR") = iBR(0) & "/" & iBR(1)
        Range("Bit" & BitNum & "_I2_" & strArea(i) & "SR") = iSR(0) & "/" & iSR(1)
    
        Erase strType, iCount, iBR, iSR
    Next i
        

'Info for Drill_Scan
    Dim iColBR As Integer
    
    Sheets("Bit" & BitNum).Select
    
    iColBR = Range("Bit" & BitNum & "_Back_Rake").Column
    iColTag = Range("Bit" & BitNum & "_Gauge").Column
    
    'Cone Depth
    If Cells(Range("Bit" & BitNum & "_Cone").row, Range("Bit" & BitNum & "_Radius").Column) < 0.001 And IsStinger(Cells(Range("Bit" & BitNum & "_Cone").row, Range("Bit" & BitNum & "_Cutter_Type").Column)) Then
        i = 1
    Else
        i = 0
    End If
    Range("Bit" & BitNum & "_DS_CS_CD") = RoundOff(Cells(Range("Bit" & BitNum & "_Cone").row + i, Range("Bit" & BitNum & "_Height2").Column) * -1, 2)
    
    'Cutting Strux Ht
    Range("Bit" & BitNum & "_DS_CS_Ht") = RoundOff(Cells(Range("Bit" & BitNum & "_Gauge").row, Range("Bit" & BitNum & "_Height2").Column) * -1, 2)
    
    'Cutting Strux Avg BR
    Range("Bit" & BitNum & "_DS_CS_BR") = RoundOff(Application.WorksheetFunction.Average(Range(Range("Bit" & BitNum & "_Back_Rake").Item(1), Cells(Range("Bit" & BitNum & "_Gauge").row - 1, Range("Bit" & BitNum & "_Back_Rake").Column))), 2)
    
    'Active Gauge Height
    Range("Bit" & BitNum & "_DS_AG_Ht") = RoundOff(Range("Bit" & BitNum & "_Height3").Item(Range("Bit" & BitNum & "_Height3").Count) * -1, 2)
    
    'Active Gauge BR
    Range("Bit" & BitNum & "_DS_AG_BR") = RoundOff(Application.WorksheetFunction.Average(Range("Bit" & BitNum & "_Gauge").Offset(, iColBR - iColTag)), 2)
    
    
    
    
    
'End Extra Info stuff
        
    If blnHidden Then Sheets("BIT" & BitNum).Hidden = True
    
    Sheets("INFO").Activate
    row_index = Range("Cutting_Structure_Info").row
    row_index = row_index + 2 + CInt(BitNum - 1) * 4
    For i = 1 To 4
        Range("I" & Trim(str(row_index + i - 1))).Value = Cutter_Type(i)
        Range("L" & Trim(str(row_index + i - 1))).Value = Cutter_Count(i)
        Range("M" & Trim(str(row_index + i - 1))).Value = BR_Low(i)
        Range("N" & Trim(str(row_index + i - 1))).Value = BR_High(i)
        Range("O" & Trim(str(row_index + i - 1))).Value = SR_Low(i)
        Range("P" & Trim(str(row_index + i - 1))).Value = SR_High(i)
    Next i
    
'    GoTo continue
continue:

    Range("G" & Trim(str(row_index))).Value = total_cutter_count
    
    
    
    Exit Function
    
ErrHandler:
If Err = 1004 Then
    Resume Next
Else
    Debug.Print "ERROR:", Now(), Err, Err.Description
    Stop
    Resume
End If

End Function

Function ProcessCompTime(iBit, strFile As String) As String
    Dim fs As scripting.FileSystemObject
    Dim txtIn As scripting.TextStream ' scripting.textstream
    Dim strLine As String, sCost(3) As Single

    Set fs = New scripting.FileSystemObject
    
    If Dir(strFile) <> "" Then
        
        Set txtIn = fs.OpenTextFile(strFile, 1) ' 1 ForReading
    
        Do While Not txtIn.AtEndOfStream
            strLine = txtIn.ReadLine
            
            If Left(strLine, 10) = "Cost Hour=" Then
                sCost(0) = CSng(Mid(strLine, 11))
                
                If sCost(0) < 0 Then
                    Range("Bit" & iBit & "_CompTime") = "Err. Neg. #"
                End If
                
                sCost(1) = Int(sCost(0))                               'hours
                sCost(2) = Int((sCost(0) - Int(sCost(0))) * 60)        'minutes
                sCost(3) = CInt((sCost(0) - Int(sCost(0))) * 60 * 60)  'seconds
                
                Range("Bit" & iBit & "_CompTime") = Format(sCost(1), "00") & ":" & Format(sCost(2), "00") & ":" & Format(sCost(3), "00")
            
                Exit Do
            End If
        Loop
        
        txtIn.Close
        
        Set txtIn = Nothing
        Set fs = Nothing
    End If

End Function


Function ProcessSummary(BitNum, Summary_File_Handler As String, Optional blnRapid As Boolean) As String

    gProc = "ProcessSummary"
    
    Dim find_count As Integer, i As Integer, Contact As String, Contact_offset As Integer, CPU_offset As Integer
    Dim case_start_row() As String
    Dim cutter_start_row() As String
    Dim cutter_end_row() As String
    Dim case_name() As String
    Dim ROP() As Double
    Dim WOB() As Double
    Dim RPM() As Double
    Dim HRS() As Double
    Dim DOC() As Double
    Dim ROCK() As String
    Dim Case_Inputs() As String
    Dim Conf_Pressure() As Integer
    Dim FOOTAGE() As Double
    Dim TIF() As Double
    Dim RIF() As Double
    Dim CIF() As Double
    Dim RIF_CIF() As Double
    Dim RIF_CIF_Normal() As Double
    Dim BETA() As Double
    Dim TORQUE() As Double
    Dim ROP_TORQUE() As Double
    Dim blnCaseError As Boolean
    Dim strSummarySheet As String
    Dim blnStingerStatic As Boolean
    
    On Error GoTo ErrHandler
    
'    subUpdate True


    start_row_index = 4
    
    Range("Bit" & BitNum & "_StingersPresent") = False
    
    If blnRapid Then
        Sheets("BIT" & BitNum & "_SUMMARY").Activate
        Used_Range_Address = ActiveSheet.UsedRange.Address
        Last_row = Range(Used_Range_Address).Rows.Count

        GoTo ProcessBit
    End If
    
    'Current_Workbook = ActiveWorkbook.Name
    Clear_Summary (BitNum)
    
    If Dir(Summary_File_Handler, vbNormal) = "" Then
        ProcessSummary = "No_Summary"
        GoTo LeaveProcess
    End If
    
'Open Summary_regular file
    If Range("ReadVer") = "old" Then
        Workbooks.OpenText FileName:=Summary_File_Handler, Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=True, Other:=True, TrailingMinusNumbers:=True
        strSummarySheet = ActiveWorkbook.Name
    Else
        strSummarySheet = OpenText(Summary_File_Handler, Range("ReadVer"))
    End If
    
    Used_Range_Address = ActiveSheet.UsedRange.Address
    Last_row = Range(Used_Range_Address).Rows.Count
    ActiveSheet.UsedRange.Copy
    Workbooks(Current_Workbook).Activate
    Sheets("BIT" & BitNum & "_SUMMARY").Activate
    Range(Used_Range_Address).PasteSpecial

     'Close the summary file
    Application.DisplayAlerts = False
    If Range("ReadVer") = "old" Then
        Windows(strSummarySheet).Close
        Workbooks(Current_Workbook).Activate
        
    Else
        Sheets(strSummarySheet).Delete
        Sheets("BIT" & BitNum & "_SUMMARY").Activate
'        DeleteQueryTables
    End If
    Application.DisplayAlerts = True


ProcessBit:
    Contact_offset = 0
    With ActiveSheet.UsedRange
        Set find_info = .Find("Date", LookIn:=xlValues, LookAt:=xlWhole)
        If Not find_info Is Nothing Then
            Range(find_info.Address).Offset(0, 1).Activate
            Build_Date = ActiveCell
        End If
        
'Find Contact info
        Set find_info = .Find("contact:", LookIn:=xlValues)    'old way of doing contact
        If Not find_info Is Nothing Then
            
            Set find_info = .Find("body(top):", LookIn:=xlValues, LookAt:=xlWhole)
            i = 0
            Do While find_info.Offset(i + 1, 0) <> ""
                i = i + 1
            Loop
            
            Contact = "On"
            
            Contact_offset = 3 + 2 * (i + 1) '3: 1 blank + 2 header rows, 2: sections (contact, friction), i: # FNFs, +1: accounts for the section headers
        Else
            Contact = "Off"
        End If
        
'Find Wear info
        Set find_info = .Find("control,", LookIn:=xlValues, LookAt:=xlWhole)
        If Not find_info Is Nothing Then
            Range(find_info.Address).Offset(0, -1).Activate
            Control = ActiveCell
            Range(find_info.Address).Offset(0, 1).Activate
            If ActiveCell = "not" Then
                Wear_Analysis = "No"
            Else
                If ActiveCell = "calc" Then
                    Wear_Analysis = "Yes"
                Else
                    Wear_Analysis = "NA"
                End If
            End If
        End If
    End With
    
    
    
    
'Find # of CPUs used
    With ActiveSheet.UsedRange
        Set find_info = .Find("CPU(s):", LookIn:=xlValues, LookAt:=xlWhole)
        If Not find_info Is Nothing Then
            Range("Bit" & BitNum & "_StaticCores") = find_info.Offset(, 1)
            CPU_offset = 3
        Else
            Range("Bit" & BitNum & "_StaticCores") = "1"
            CPU_offset = 0
        End If
    End With
    
'Get Comp Times
    A = ProcessCompTime(BitNum, Left(Summary_File_Handler, InStrRev(Summary_File_Handler, "\")) & "summary_variables")
    
    
    
        
    
'Check if the project is a stinger Static project
    With ActiveSheet.UsedRange
        Set find_info = .Find("Output:", LookIn:=xlValues, LookAt:=xlWhole)
        If Not find_info Is Nothing Then
            If Range(find_info.Address).Offset(2, 2).Value = "max" And Range(find_info.Address).Offset(2, 3).Value = "min" And (Range(find_info.Address).Offset(2, 4).Value = "mean" Or Range(find_info.Address).Offset(2, 4).Value = "ave") Then
                'stinger static project
                
                ProcessStingerSummary CInt(BitNum), Summary_File_Handler, WOB(), RPM(), ROP(), TORQUE(), DOC(), ROCK(), Conf_Pressure(), TIF(), RIF(), CIF(), RIF_CIF(), RIF_CIF_Normal(), BETA(), find_count, blnRapid
                
                Range("Bit" & BitNum & "_StingerStatic") = "Yes"
                blnStingerStatic = True
                
                
                ProcessSummary = "Stinger_Static"
                GoTo EnterSummaryInfo
            
            Else
                Range("Bit" & BitNum & "_StingerStatic") = "No"
                Range("Bit" & BitNum & "_StaticRevs") = "-"
            End If
        Else
            Range("Bit" & BitNum & "_StingerStatic") = "No"
            Range("Bit" & BitNum & "_StaticRevs") = "-"
        End If
    End With
    
    
    

    
    
    
    
    
    
    
    
    find_count = 0
    cutterdata_row = " "
    With ActiveSheet.UsedRange
        Set find_case = .Find("percent", LookIn:=xlValues)
        If Not find_case Is Nothing Then
            firstAddress = find_case.Address
            Do
                find_count = find_count + 1
                Set find_case = .FindNext(find_case)
            Loop While Not find_case Is Nothing And find_case.Address <> firstAddress
        End If
    End With
    ReDim case_start_row(1 To find_count) As String
    ReDim cutter_start_row(1 To find_count) As String
    ReDim cutter_end_row(1 To find_count) As String
    ReDim case_name(1 To find_count) As String
    ReDim ROP(1 To find_count) As Double
    ReDim WOB(1 To find_count) As Double
    ReDim RPM(1 To find_count) As Double
    ReDim HRS(1 To find_count) As Double
    ReDim DOC(1 To find_count) As Double
    ReDim ROCK(1 To find_count) As String
    ReDim Case_Inputs(1 To find_count) As String
    ReDim Conf_Pressure(1 To find_count) As Integer
    ReDim FOOTAGE(1 To find_count) As Double
    ReDim TIF(1 To find_count) As Double
    ReDim CIF(1 To find_count) As Double
    ReDim RIF(1 To find_count) As Double
    ReDim RIF_CIF(1 To find_count) As Double
    ReDim RIF_CIF_Normal(1 To find_count) As Double
    ReDim BETA(1 To find_count) As Double
    ReDim TORQUE(1 To find_count) As Double
    ReDim ROP_TORQUE(1 To find_count) As Double
    i = 0
    With ActiveSheet.UsedRange
        Set find_case = .Find("percent", LookIn:=xlValues)
        If Not find_case Is Nothing Then
            firstAddress = find_case.Address
            Do
                i = i + 1
                case_start_row(i) = find_case.row
                Set find_case = .FindNext(find_case)
            Loop While Not find_case Is Nothing And find_case.Address <> firstAddress
        End If
    End With
    i = 0

    With ActiveSheet.UsedRange
        Set find_case = .Find("fn", LookIn:=xlValues, LookAt:=xlWhole)
        If Not find_case Is Nothing Then
            firstAddress = find_case.Address
            Do
                i = i + 1
                cutter_start_row(i) = find_case.row + 3
                Set find_case = .FindNext(find_case)
            Loop While Not find_case Is Nothing And find_case.Address <> firstAddress
        End If
    End With
    For i = 1 To find_count - 1
        cutter_end_row(i) = case_start_row(i + 1) - 2 - Contact_offset
    Next i
    cutter_end_row(i) = Last_row - 2 - Contact_offset - CPU_offset

    For i = 1 To find_count
        With Range("A" + case_start_row(i) + ":Z" + cutter_end_row(i))
            Set find_para = .Find("percent", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                TIF(i) = Range("D" & find_para.Row).Value
                TIF(i) = find_para.Offset(0, 3).Value
            End If
            Set find_para = .Find("rif/wob", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                RIF(i) = Range("M" & find_para.Row).Value
                RIF(i) = find_para.Offset(0, 1).Value
            End If
            Set find_para = .Find("cif/wob", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                CIF(i) = Range("N" & find_para.Row).Value
                CIF(i) = find_para.Offset(0, 1).Value
            End If
            Set find_para = .Find("rif/cif", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                RIF_CIF(i) = Range("H" & find_para.Row).Value
                RIF_CIF(i) = find_para.Offset(0, 1).Value
                If RIF_CIF(i) > 1 Then
                    RIF_CIF_Normal(i) = 1 / RIF_CIF(i)
                Else
                    RIF_CIF_Normal(i) = RIF_CIF(i)
                End If
            End If
            Set find_para = .Find("b", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)

            If Not find_para Is Nothing Then
'                BETA(i) = Range("I" & find_para.Row).Value
                Do While Not IsNumeric(find_para.Offset(0, 1))
                    Set find_para = .Find("b", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, After:=find_para)
                Loop
                BETA(i) = find_para.Offset(0, 1).Value
            End If
            
            Set find_para = .Find("weight-on-bit", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                WOB(i) = Range("C" & find_para.Row).Value
                WOB(i) = find_para.Offset(0, 2).Value
            End If
            Set find_para = .Find("torque", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                TORQUE(i) = Range("D" & find_para.Row).Value
                TORQUE(i) = find_para.Offset(0, 2).Value
            End If
            Set find_para = .Find("penetration", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                ROP(i) = Range("D" & find_para.Row).Value
                ROP(i) = find_para.Offset(0, 3).Value
                If TORQUE(i) = 0 Then
                    ROP_TORQUE(i) = 0
                Else
                    ROP_TORQUE(i) = RoundOff(ROP(i) / TORQUE(i), 4)
                End If
            End If
            Set find_para = .Find("rotary", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                RPM(i) = Range("D" & find_para.Row).Value
                RPM(i) = find_para.Offset(0, 3).Value
            End If
            Set find_para = .Find("rotating", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                HRS(i) = Range("D" & find_para.Row).Value
                HRS(i) = find_para.Offset(0, 3).Value
            End If
            Set find_para = .Find("advance", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
'                DOC(i) = Range("D" & find_para.Row).Value
                DOC(i) = find_para.Offset(0, 2).Value
            End If
            Set find_para = .Find("cutter/rock", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_para Is Nothing Then
                ROCK(i) = Range("A" & (find_para.row + 1)).Value
                ROCK(i) = Mid(ROCK(i), InStr(1, ROCK(i), "_") + 1)
                ROCK(i) = Mid(ROCK(i), InStr(1, ROCK(i), "_") + 1)
                Conf_Pressure(i) = Mid(ROCK(i), InStr(1, ROCK(i), "_") + 1)
                ROCK(i) = Mid(ROCK(i), 1, InStr(1, ROCK(i), "_") - 1)
            End If
            
            If i = 1 Then
                FOOTAGE(i) = ROP(i) * HRS(i)
            Else
                FOOTAGE(i) = FOOTAGE(i - 1) + (ROP(i) * (HRS(i) - HRS(i - 1)))
            End If
        End With
    Next i


EnterSummaryInfo:


    Sheets("INFO").Activate
    If InStr(1, Build_Date, "/") > 0 Then Build_Date = Format(Build_Date, "YYYYMMDD")
    Range("Bit" & BitNum & "_Build") = Mid(Build_Date, 5, 2) & "/" & Right(Build_Date, 2) & "/" & Left(Build_Date, 4)
    Range("Bit" & BitNum & "_Control") = Control
    Range("Bit" & BitNum & "_Cases") = find_count
    Range("Bit" & BitNum & "_Contact") = Contact
    Range("Bit" & BitNum & "_WearProj") = Wear_Analysis
    Range("Bit" & BitNum & "_Hours") = HRS(find_count)
    
    Sheets("SUMMARY").Activate
    If BitNum = 1 And Control = "wob" Then
        Range("Control_Info").Value = "Instantaneous ROP (ft/hr)"
        Case_details = "WOB_"
    ElseIf BitNum = 1 And Control = "rop" Then
        Range("Control_Info").Value = "WOB (Klb)"
        Case_details = "ROP_"
    End If
    
    If Wear_Analysis = "Yes" Then
        Case_details = Case_details & "HRS_"
    End If
    Case_details = Case_details & "RPM_ROCK_CONFINING PRESSURE"
    Range("Case_Info").Value = "Case Info (" & Case_details & ")"
    row_index = Range("Case_Info").row + 2
    case_column_index = Range("Case_Info").Column
    control_column_index = Range("Control_Info").Column
    ROP_column_index = Range("ROP_Compare").Column
    DOC_column_index = Range("DOC_info").Column
    Cum_Hrs_index = Range("Cumulative_Hours").Column
    Cum_Foot_index = Range("Cumulative_Footage").Column
    Torque_index = Range("TORQUE_info").Column
    TQ_Compare_index = Range("TQ_Compare").Column
    TIF_index = Range("TIF_info").Column
    RIF_index = Range("RIF_info").Column
    CIF_index = Range("CIF_info").Column
    RIF_CIF_index = Range("RIF_CIF_info").Column
    RIF_CIF_Normal_index = Range("RIF_CIF_Normal_info").Column
    BETA_index = Range("BETA_info").Column
    ROP_TORQUE_index = Range("ROP_TORQUE_info").Column
    ROP_TORQUE_Compare_index = Range("ROP_TORQUE_Compare").Column
    
    
    bit_offset = BitNum
    
    'If BitNum = 1 Then
    '    Cells(row_index + 1, case_column_index + 1).Select
    '    Range(ActiveCell, ActiveCell.Offset(0, 0)).Name = "Cases"
    'End If
    'If Not BitNum = 1 Then
    '    Cells(row_index + 1, ROP_column_index + BitNum - 1).Select
    '    Range(ActiveCell, ActiveCell.Offset(0, 0)).Name = "Bit" & BitNum & "_ROP"
    'End If
    
    For i = 1 To find_count
        Cells(row_index + i, case_column_index).Value = i
        CellFormat1 (Cells(row_index + i, case_column_index).Address)
        Cells(row_index + i, control_column_index).Value = i
        CellFormat1 (Cells(row_index + i, control_column_index).Address)
        Cells(row_index + i, ROP_column_index).Value = i
        CellFormat1 (Cells(row_index + i, ROP_column_index).Address)
        Cells(row_index + i, DOC_column_index).Value = i
        CellFormat1 (Cells(row_index + i, DOC_column_index).Address)
        Cells(row_index + i, Cum_Hrs_index).Value = i
        CellFormat1 (Cells(row_index + i, Cum_Hrs_index).Address)
        Cells(row_index + i, Cum_Foot_index).Value = i
        CellFormat1 (Cells(row_index + i, Cum_Foot_index).Address)
        Cells(row_index + i, Torque_index).Value = i
        CellFormat1 (Cells(row_index + i, Torque_index).Address)
        Cells(row_index + i, TQ_Compare_index).Value = i
        CellFormat1 (Cells(row_index + i, TQ_Compare_index).Address)
        Cells(row_index + i, TIF_index).Value = i
        CellFormat1 (Cells(row_index + i, TIF_index).Address)
        Cells(row_index + i, RIF_index).Value = i
        CellFormat1 (Cells(row_index + i, RIF_index).Address)
        Cells(row_index + i, CIF_index).Value = i
        CellFormat1 (Cells(row_index + i, CIF_index).Address)
        Cells(row_index + i, RIF_CIF_index).Value = i
        CellFormat1 (Cells(row_index + i, RIF_CIF_index).Address)
        Cells(row_index + i, RIF_CIF_Normal_index).Value = i
        CellFormat1 (Cells(row_index + i, RIF_CIF_Normal_index).Address)
        Cells(row_index + i, BETA_index).Value = i
        CellFormat1 (Cells(row_index + i, BETA_index).Address)
        Cells(row_index + i, ROP_TORQUE_index).Value = i
        CellFormat1 (Cells(row_index + i, ROP_TORQUE_index).Address)
        Cells(row_index + i, ROP_TORQUE_Compare_index).Value = i
        CellFormat1 (Cells(row_index + i, ROP_TORQUE_Compare_index).Address)
        
        If Control = "wob" Then
            'Cells(row_index + i, control_column_index + bit_offset).Value = Str(ROP(i)) & " (" & Str(DOC(i)) & ")"
            Cells(row_index + i, control_column_index + bit_offset).Value = ROP(i)
            Case_details = ConvertToK(WOB(i)) & "_"
'ROP Compare
            If Not BitNum = 1 And Not Cells(row_index + i, control_column_index + 1).Value = "" Then
                baseline_rop = Cells(row_index + i, control_column_index + 1).Value
                If baseline_rop = 0 Then
                    Cells(row_index + i, ROP_column_index + bit_offset - 1).Value = "-"
                Else
                    ROP_Change = (ROP(i) - baseline_rop) / baseline_rop * 100
                    Cells(row_index + i, ROP_column_index + bit_offset - 1).Value = RoundOff(ROP_Change, 2)
                End If
                CellFormat1 (Cells(row_index + i, ROP_column_index + bit_offset - 1).Address)
            ElseIf BitNum = 1 Then
                baseline_rop = ROP(i)
                For Z = 2 To 5
                    If Not Range("Bit" & Z & "_Location").Value = "" Then
                        Bit_ROP = Cells(row_index + i, control_column_index + Z).Value
                        If baseline_rop = 0 Then
                            Cells(row_index + i, ROP_column_index + bit_offset + Z - 2).Value = "-"
'                            Cells(row_index + i, ROP_column_index + bit_offset - 1).Value = "-"
                        Else
                            ROP_Change = (Bit_ROP - baseline_rop) / baseline_rop * 100
                            Cells(row_index + i, ROP_column_index + bit_offset + Z - 2).Value = RoundOff(ROP_Change, 2)
'                            Cells(row_index + i, ROP_column_index + bit_offset - 1).Value = RoundOff(ROP_Change, 2)
                        End If
                        CellFormat1 (Cells(row_index + i, ROP_column_index + Z - 1).Address)
                    End If
                Next Z
            End If
        Else
            Cells(row_index + i, control_column_index + bit_offset).Value = ConvertToK(WOB(i))
            Case_details = ROP(i) & "_"
        End If

        Cells(row_index + i, DOC_column_index + bit_offset).Value = DOC(i)
        Cells(row_index + i, Torque_index + bit_offset).Value = TORQUE(i)
        
'Tq Compare
        If Not BitNum = 1 And Not Cells(row_index + i, Torque_index + 1).Value = "" Then
            baseline_tq = Cells(row_index + i, Torque_index + 1).Value
            If baseline_tq = 0 Then
                Cells(row_index + i, TQ_Compare_index + bit_offset - 1).Value = "-"
            Else
                TQ_Change = (TORQUE(i) - baseline_tq) / baseline_tq * 100
                Cells(row_index + i, TQ_Compare_index + bit_offset - 1).Value = RoundOff(TQ_Change, 2)
            End If
            CellFormat1 (Cells(row_index + i, TQ_Compare_index + bit_offset - 1).Address)
        ElseIf BitNum = 1 Then
            baseline_tq = TORQUE(i)
            For Z = 2 To 5
                If Not Range("Bit" & Z & "_Location").Value = "" Then
                    Bit_TQ = Cells(row_index + i, Torque_index + Z).Value
                    If baseline_tq = 0 Then
                        Cells(row_index + i, TQ_Compare_index + bit_offset + Z - 2).Value = "-"
'                        Cells(row_index + i, TQ_Compare_index + bit_offset - 1).Value = "-"
                    Else
                        TQ_Change = (Bit_TQ - baseline_tq) / baseline_tq * 100
                        Cells(row_index + i, TQ_Compare_index + bit_offset + Z - 2).Value = RoundOff(TQ_Change, 2)
'                        Cells(row_index + i, TQ_Compare_index + bit_offset - 1).Value = RoundOff(TQ_Change, 2)
                    End If
                    CellFormat1 (Cells(row_index + i, TQ_Compare_index + Z - 1).Address)
                End If
            Next Z
        End If
        
        Cells(row_index + i, TIF_index + bit_offset).Value = TIF(i)
        Cells(row_index + i, RIF_index + bit_offset).Value = RIF(i)
        Cells(row_index + i, CIF_index + bit_offset).Value = CIF(i)
        Cells(row_index + i, RIF_CIF_index + bit_offset).Value = RIF_CIF(i)
        Cells(row_index + i, RIF_CIF_Normal_index + bit_offset).Value = RIF_CIF_Normal(i)
        Cells(row_index + i, BETA_index + bit_offset).Value = BETA(i)
        Cells(row_index + i, ROP_TORQUE_index + bit_offset).Value = ROP_TORQUE(i)
        
        
'ROP_TORQUE Compare
        If Not BitNum = 1 And Not Cells(row_index + i, ROP_TORQUE_index + 1).Value = "" Then
            baseline_ROP_TORQUE = Cells(row_index + i, ROP_TORQUE_index + 1).Value
            If baseline_ROP_TORQUE = 0 Then
                Cells(row_index + i, ROP_TORQUE_Compare_index + bit_offset - 1).Value = "-"
            Else
                ROP_TORQUE_Change = (ROP_TORQUE(i) - baseline_ROP_TORQUE) / baseline_ROP_TORQUE * 100
                Cells(row_index + i, ROP_TORQUE_Compare_index + bit_offset - 1).Value = RoundOff(ROP_TORQUE_Change, 2)
            End If
            CellFormat1 (Cells(row_index + i, ROP_TORQUE_Compare_index + bit_offset - 1).Address)
        ElseIf BitNum = 1 Then
            baseline_ROP_TORQUE = ROP_TORQUE(i)
            For Z = 2 To Range("MaxBits")
                If Not Range("Bit" & Z & "_Location").Value = "" Then
                    Bit_ROP_TORQUE = Cells(row_index + i, ROP_TORQUE_index + Z).Value
                    If baseline_ROP_TORQUE = 0 Then
                        Cells(row_index + i, ROP_TORQUE_Compare_index + bit_offset + Z - 2).Value = "-"
'                        Cells(row_index + i, ROP_TORQUE_Compare_index + bit_offset - 1).Value = "-"
                    Else
                        ROP_TORQUE_Change = (Bit_ROP_TORQUE - baseline_ROP_TORQUE) / baseline_ROP_TORQUE * 100
                        Cells(row_index + i, ROP_TORQUE_Compare_index + bit_offset + Z - 2).Value = RoundOff(ROP_TORQUE_Change, 2)
'                        Cells(row_index + i, ROP_TORQUE_Compare_index + bit_offset - 1).Value = RoundOff(ROP_TORQUE_Change, 2)
                    End If
                    CellFormat1 (Cells(row_index + i, ROP_TORQUE_Compare_index + Z - 1).Address)
                End If
            Next Z
        End If
        
        
        
        If Wear_Analysis = "Yes" Then
            Case_details = Case_details & HRS(i) & "_"
            Cells(row_index + i, Cum_Hrs_index + bit_offset).Value = HRS(i)
            CellFormat1 (Cells(row_index + i, Cum_Hrs_index + bit_offset).Address)
'            If Control = "wob" Then
                Cells(row_index + i, Cum_Foot_index + bit_offset).Value = FOOTAGE(i)
                CellFormat1 (Cells(row_index + i, Cum_Foot_index + bit_offset).Address)
'            End If
        End If
        Case_details = Case_details & RPM(i) & "_" & ROCK(i) & "_" & ConvertToK(Conf_Pressure(i))
        Cells(row_index + i, case_column_index + bit_offset).Value = Case_details
        If Cells(row_index + i, case_column_index + bit_offset).Value <> Cells(row_index + i, case_column_index + 1).Value Then blnCaseError = True
        
        Case_Inputs(i) = Case_details
        CellFormat1 (Cells(row_index + i, control_column_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, case_column_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, DOC_column_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, Torque_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, TIF_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, RIF_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, CIF_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, RIF_CIF_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, RIF_CIF_Normal_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, BETA_index + bit_offset).Address)
        CellFormat1 (Cells(row_index + i, ROP_TORQUE_index + bit_offset).Address)
    Next i
    
'Name Case Range
    If find_count > Range("CasesID").Count Then
'        Range(Cells(row_index + 1, case_column_index + BitNum), Cells(row_index + find_count, case_column_index + BitNum)).Select
'        Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Cases"
        
        Range(Cells(row_index + 1, case_column_index), Cells(row_index + find_count, case_column_index)).Name = "CasesID"
        If BitNum = 1 Then
            Range("Cases").ClearContents
            
            Cells(row_index + 1, case_column_index + 1).Select
            Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Copy
            Range("ConstCase").PasteSpecial xlPasteValues
            Range(Range("ConstCase"), Range("ConstCase").Offset(find_count - 1, 0)).Name = "Cases"
'            Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Cases"
        End If
    ElseIf BitNum = 1 Then
        Range("Cases").ClearContents
        
        Cells(row_index + 1, case_column_index + 1).Select
        Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Copy
        Range("ConstCase").PasteSpecial xlPasteValues
        Range(Range("ConstCase"), Range("ConstCase").Offset(find_count - 1, 0)).Name = "Cases"
'        Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Cases"
    Else
        If Control = "wob" Then
            Cells(row_index + 1, ROP_column_index + BitNum - 1).Select
            Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_ROP_Change"
        End If
    End If
    Cells(row_index + 1, control_column_index + bit_offset).Select
    
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_ControlList"
'    If Control = "wob" Then
'        Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).name = "Bit" & BitNum & "_ROP"
'    Else
'        Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).name = "Bit" & BitNum & "_WOB"
'    End If
    Cells(row_index + 1, DOC_column_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_DOC"
    Cells(row_index + 1, Cum_Hrs_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_CUM_HRS"
    Cells(row_index + 1, Cum_Foot_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_CUM_FOOT"
    Cells(row_index + 1, Torque_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_TORQUE"
    Cells(row_index + 1, TQ_Compare_index + BitNum - 1).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_TQ_Change"
    Cells(row_index + 1, TIF_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_TIF"
    Cells(row_index + 1, RIF_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_RIF"
    Cells(row_index + 1, CIF_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_CIF"
    Cells(row_index + 1, RIF_CIF_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_RIF_CIF"
    Cells(row_index + 1, RIF_CIF_Normal_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_RIF_CIF_Normal"
    Cells(row_index + 1, BETA_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_BETA"
    Cells(row_index + 1, ROP_TORQUE_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_ROP_TORQUE"
    Cells(row_index + 1, ROP_TORQUE_Compare_index + BitNum - 1).Select
    Range(ActiveCell, ActiveCell.Offset(find_count - 1, 0)).Name = "Bit" & BitNum & "_ROP_TORQUE_Change"
    
    
    
    If Not blnStingerStatic Then
        Dim Heading_Array1 As Variant
        Dim Heading_Array2 As Variant
        Dim Force_Angle, Normal_Force, Vertical_Force, Circumferential_Force, Side_Force, BR, SR, Eff_BR, Eff_SR As Double
        Heading_Array1 = Array("area", "vol", "force", "Circumferential Force (fx/fcut)", "Vertical Force (fy)", "Normal Force (fn)", "vel", "wf area", "wf area", "eff", "eff", "ProE fx", "ProE fy", "ProE fz", "Torque to ctr center", "Workrate (Circumferential Force)", "Workrate (Normal Force)", "Radial Force", "Side Force", "Resultant Force", "Workrate (Resultant Force)", "Delam Force", "Workrate (Delam Force)", "Torque to Bit Axis") ', "pro_fxy", "Fwear", "Fdelam_normal", "Fdelam_side", "Fdelam")
        Heading_Array2 = Array("cut", "cut", "angle", "lb", "lb", "lb", "ft/s", "in**2", "", "sr", "br", "lb", "lb", "lb", "lbf-in", "lb-ft/s", "lb-ft/s", "lb", "lb", "lb", "lb-ft/s", "lb", "lb-ft/s", "lbf-ft") ', "lb", "lb", "lb", "lb", "lb")
        For i = 1 To find_count
            
            Sheets("BIT" & BitNum & "_SUMMARY").Activate
            Range("C" + cutter_start_row(i) + ":Q" + cutter_end_row(i)).Select
            Selection.Copy
            
            Sheets("BIT" & BitNum).Activate
            info_last_col = Range(ActiveSheet.UsedRange.Address).Columns.Count
    '        If i = 1 Then
    '            Range(Cells(4, 1), Cells(, info_last_col)).AutoFilter
    ''            Range("A4:" & fXLNumberToLetter(CLng(info_last_col)) & "4").AutoFilter
    '        End If
            
            Cells(start_row_index - 3, info_last_col + 2).Select
            ActiveCell.Value = "CASE " & str(i)
            ActiveCell.Offset(1, 0).Value = Case_Inputs(i)
            Cells(start_row_index + 1, info_last_col + 2).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 10)).PasteSpecial
            For j = 0 To UBound(Heading_Array1)
                Cells(start_row_index - 1, j + info_last_col + 2).Value = Heading_Array1(j)
                Cells(start_row_index, j + info_last_col + 2).Value = Heading_Array2(j)
            Next j
            For j = 1 To (cutter_end_row(i) - cutter_start_row(i) + 1)
                Cells(start_row_index + j, info_last_col + 2).Offset(0, 12).Value = (-1) * Cells(start_row_index + j, info_last_col + 2).Offset(0, 12).Value
                Cells(start_row_index + j, info_last_col + 2).Offset(0, 15).Value = Cells(start_row_index + j, info_last_col + 2).Offset(0, 3).Value * Cells(start_row_index + j, info_last_col + 2).Offset(0, 6).Value
                Cells(start_row_index + j, info_last_col + 2).Offset(0, 16).Value = Cells(start_row_index + j, info_last_col + 2).Offset(0, 5).Value * Cells(start_row_index + j, info_last_col + 2).Offset(0, 6).Value
                Force_Angle = Cells(start_row_index + j, info_last_col + 2).Offset(0, 2).Value * Excel.WorksheetFunction.Pi() / 180
                Normal_Force = Cells(start_row_index + j, info_last_col + 2).Offset(0, 5).Value
                Vertical_Force = Cells(start_row_index + j, info_last_col + 2).Offset(0, 4).Value
                Circumferential_Force = Cells(start_row_index + j, info_last_col + 2).Offset(0, 3).Value
                BR = Cells(start_row_index + j, 1).Offset(0, 5).Value * Excel.WorksheetFunction.Pi() / 180
                SR = Cells(start_row_index + j, 1).Offset(0, 6).Value * Excel.WorksheetFunction.Pi() / 180
                If Cells(start_row_index + j, info_last_col + 2).Offset(0, 10).Value <> "NA" Then Eff_BR = Cells(start_row_index + j, info_last_col + 2).Offset(0, 10).Value * Excel.WorksheetFunction.Pi() / 180
                If Cells(start_row_index + j, info_last_col + 2).Offset(0, 9).Value <> "NA" Then Eff_SR = Cells(start_row_index + j, info_last_col + 2).Offset(0, 9).Value * Excel.WorksheetFunction.Pi() / 180
                If Force_Angle <> 0 Then
                    Cells(start_row_index + j, info_last_col + 2).Offset(0, 17).Value = (Vertical_Force * Cos(Force_Angle) - Normal_Force) / Sin(Force_Angle)
                    Cells(start_row_index + j, info_last_col + 2).Offset(0, 18).Value = (Normal_Force * Cos(Force_Angle) - Vertical_Force) / Sin(Force_Angle)
                Else
                    Cells(start_row_index + j, info_last_col + 2).Offset(0, 17).Value = 0
                    Cells(start_row_index + j, info_last_col + 2).Offset(0, 18).Value = 0
                End If
                Side_Force = Cells(start_row_index + j, info_last_col + 2).Offset(0, 18).Value
        
        'Resultant and Delam forces
                Cells(start_row_index + j, info_last_col + 2).Offset(0, 19).Value = ((Cells(start_row_index + j, info_last_col + 2).Offset(0, 11).Value) ^ 2 + (Cells(start_row_index + j, info_last_col + 2).Offset(0, 12).Value) ^ 2 + (Cells(start_row_index + j, info_last_col + 2).Offset(0, 13).Value) ^ 2) ^ 0.5
                Cells(start_row_index + j, info_last_col + 2).Offset(0, 20).Value = Cells(start_row_index + j, info_last_col + 2).Offset(0, 19).Value * Cells(start_row_index + j, info_last_col + 2).Offset(0, 6).Value
                Cells(start_row_index + j, info_last_col + 2).Offset(0, 21).Value = ((Cells(start_row_index + j, info_last_col + 2).Offset(0, 11).Value) ^ 2 + (Cells(start_row_index + j, info_last_col + 2).Offset(0, 12).Value) ^ 2) ^ 0.5
                Cells(start_row_index + j, info_last_col + 2).Offset(0, 22).Value = Cells(start_row_index + j, info_last_col + 2).Offset(0, 21).Value * Cells(start_row_index + j, info_last_col + 2).Offset(0, 6).Value
                Cells(start_row_index + j, info_last_col + 2).Offset(0, 23).Value = Cells(start_row_index + j, info_last_col + 2).Offset(0, 15).Value / (2 * 3.14 * RPM(i) / 60) 'WorkRate Circumferential/Angular Velocity
                
                'Cells(start_row_index + j, info_last_col + 2).Offset(0, 19).Value = Sqr(Cells(start_row_index + j, info_last_col + 2).Offset(0, 11).Value * Cells(start_row_index + j, info_last_col + 2).Offset(0, 11).Value + Cells(start_row_index + j, info_last_col + 2).Offset(0, 12).Value * Cells(start_row_index + j, info_last_col + 2).Offset(0, 12).Value)
                'Cells(start_row_index + j, info_last_col + 2).Offset(0, 20).Value = Normal_Force * Sin(Eff_BR) * Cos(Eff_SR) + Circumferential_Force * Cos(Eff_BR) * Cos(Eff_SR) + Side_Force * Sin(Eff_SR)
                'Cells(start_row_index + j, info_last_col + 2).Offset(0, 21).Value = Normal_Force * Cos(Eff_BR) - Circumferential_Force * Sin(Eff_BR)
                'Cells(start_row_index + j, info_last_col + 2).Offset(0, 22).Value = -Normal_Force * Sin(Eff_BR) * Sin(Eff_SR) - Circumferential_Force * Cos(Eff_BR) * Sin(Eff_SR) + Side_Force * Cos(Eff_SR)
                'Cells(start_row_index + j, info_last_col + 2).Offset(0, 23).Value = Sqr(Cells(start_row_index + j, info_last_col + 2).Offset(0, 21).Value * Cells(start_row_index + j, info_last_col + 2).Offset(0, 21).Value + Cells(start_row_index + j, info_last_col + 2).Offset(0, 22).Value * Cells(start_row_index + j, info_last_col + 2).Offset(0, 22).Value)
            Next j
            case_name(i) = "Bit" & BitNum & "_Case" & LTrim(str(i))
            
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 0).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_AreaCut"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 1).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_VolCut"
            
            
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 3).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Circumferential"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 4).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Vertical"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 5).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Normal"
            Range(ActiveCell, ActiveCell.Offset(Range("Bit" & BitNum & "_InnerCutters").Rows.Count - 1, 0)).Name = case_name(i) + "_NormalInner"
            Range(ActiveCell.Offset(Range("Bit" & BitNum & "_InnerCutters").Rows.Count, 0), ActiveCell.Offset(Range("Bit" & BitNum & "_InnerCutters").Rows.Count + Range("Bit" & BitNum & "_OuterCutters").Rows.Count - 1, 0)).Name = case_name(i) + "_NormalOuter"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 7).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_WearFlatArea"
            'stealth   fx, fy along the face of the cutter. fy acts through center (~normal)
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 11).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_ProE_fx"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 12).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_ProE_fy"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 13).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_ProE_fz"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 14).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Cutter_Torque"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 15).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Circumferential_Workrate"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 16).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Normal_Workrate"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 17).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Radial"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 18).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Side"
            
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 19).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Resultant"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 20).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Resultant_Workrate"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 21).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Delam"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 22).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_Delam_Workrate"
            Cells(start_row_index + 1, info_last_col + 2).Offset(0, 23).Select
            Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_TQ_BIT"
            'Cells(start_row_index + 1, info_last_col + 2).Offset(0, 19).Select
            'Range(ActiveCell, ActiveCell.Offset(cutter_end_row(i) - cutter_start_row(i), 0)).Name = case_name(i) + "_ProE_fxy"
        Next i
        
        
        If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
        
        Range("A4:" & fXLNumberToLetter(CLng(Range(ActiveSheet.UsedRange.Address).Columns.Count)) & "4").AutoFilter
    End If
    
'    Sheets("FORCES").Activate
'    For i = 1 To find_count
'        Range("A" & Trim(Str((2 * i - 1) * 11))).Value = "CASE " & Str(i)
'        Range("A" & Trim(Str((2 * i - 1) * 11 + 1))).Value = Case_Inputs(i)
'    Next i
'
'    'stealth
'    Sheets("CUTTER_FORCES").Activate
'    For i = 1 To find_count
'        Range("A" & Trim(Str((2 * i - 1) * 11))).Value = "CASE " & Str(i)
'        Range("A" & Trim(Str((2 * i - 1) * 11 + 1))).Value = Case_Inputs(i)
'    Next i
    
    'Sheets("WEAR").Activate
    'ActiveSheet.ChartObjects("WearFlatArea").Activate
    'ActiveChart.SeriesCollection(1).XValues = "=test2.xlsm!" & i
    'ActiveChart.SeriesCollection(CInt(Right(BitName, 1))).Values = "=" & Current_Workbook & "!" & BitName & "_case" & Trim(Str(i - 1)) & "_WearFlatArea"
    
    If Wear_Analysis = "Yes" Then
        Dim strCalcFile As String
        
        BFH_pos = InStrRev(Summary_File_Handler, "\")
        strCalcFile = Mid(Summary_File_Handler, 1, BFH_pos) & "calc.in"
        On Error GoTo ErrHandler:
            A = GetWearCoef(BitNum, strCalcFile)
    End If
    
    ProcessSummary = CStr(blnCaseError)
    
LeaveProcess:
    Sheets("HOME").Activate
    
    Exit Function
    
ErrHandler:

Debug.Print Now(), gProc, Err, Err.Description

    If Err = 13 Or Err = 1004 Then
        Resume
        Sheets("HOME").Activate
    ElseIf Err = 9 Then
        Resume Next
    
    Else
        
        Debug.Print Err, Err.Description
        Resume
    End If
    
End Function


Function ProcessStingerSummary(iBit As Integer, strFile As String, WOB() As Double, RPM() As Double, ROP() As Double, TORQUE() As Double, DOC() As Double, ROCK() As String, Conf_Pressure() As Integer, TIF() As Double, RIF() As Double, CIF() As Double, RIF_CIF() As Double, RIF_CIF_Normal() As Double, BETA() As Double, iCases As Integer, Optional blnRapid As Boolean) As String

    gProc = "ProcessStingerSummary"
    
    Dim strDir As String, strSheet As String
    Dim iFile As Integer, i As Integer, iSummCol As Integer, iCutters As Integer, iBitCol As Integer, k As Integer
    Dim fs As scripting.FileSystemObject
    Dim txtIn As scripting.TextStream ' scripting.textstream
    Dim strLine As String
    Dim strArray As Variant, strParams() As String
    Dim blnNameFound As Boolean
    Dim RockIndex() As String
    Dim iRockIndex As Integer, strRockName As String

    On Error GoTo ErrHandler
    
    strDir = Left(strFile, InStrRev(strFile, "\"))
    
    blnNameFound = True
    
'get case count
    Set fs = New scripting.FileSystemObject
    
    strFile = strDir & "summary_doc"
    
    If Dir(strFile) <> "" Then
        
        Set txtIn = fs.OpenTextFile(strFile, 1) ' 1 ForReading
        
        Range("Bit" & iBit & "_StingersPresent") = True
    
        Do While Not txtIn.AtEndOfStream
            strLine = txtIn.ReadLine
            strLine = txtIn.ReadLine
            
            iCases = Val(Trim(strLine))
            
            If iCases <> 0 And blnNameFound Then
                ReDim WOB(1 To iCases) As Double
                ReDim RPM(1 To iCases) As Double
                ReDim ROP(1 To iCases) As Double
                ReDim TORQUE(1 To iCases) As Double
                ReDim DOC(1 To iCases) As Double
                ReDim ROCK(1 To iCases) As String
                ReDim Conf_Pressure(1 To iCases) As Integer
                ReDim TIF(1 To iCases) As Double
                ReDim RIF(1 To iCases) As Double
                ReDim CIF(1 To iCases) As Double
                ReDim BETA(1 To iCases) As Double
                ReDim RIF_CIF(1 To iCases) As Double
                ReDim RIF_CIF_Normal(1 To iCases) As Double
                
                blnNameFound = False
            End If
            
            strLine = txtIn.ReadLine
            
            For i = 1 To iCases
                strLine = txtIn.ReadLine
                
            'The original import delimited on space and tab. I'm changing these to @s and will split line into an array at the @s. This may be an issue if the user uses an @ sign in their bit name or path.
                strLine = Replace(strLine, " ", "@")
                strLine = Replace(strLine, vbTab, "@")
                
            'Make sure there are no extra spaces between line data
                Do While InStr(1, strLine, "@@") > 0
                    strLine = Replace(strLine, "@@", "@")
                Loop
                    
            'Split into an array that will be entered into the blank sheet.
                strParams = Split(strLine, "@")
                
                WOB(i) = strParams(2)
                RPM(i) = strParams(3)
                ROP(i) = strParams(4)
                TORQUE(i) = strParams(5)
                DOC(i) = strParams(6)
            Next i
    
        Loop
    
        Set txtIn = Nothing
    End If
    
    
    
    strFile = strDir & "rock.in"
    
    If Dir(strFile) <> "" Then
        Set txtIn = fs.OpenTextFile(strFile, ForReading)
        
        Do While Not txtIn.AtEndOfStream
            strLine = txtIn.ReadLine
            
            If InStr(1, strLine, "RockIndex"": {") > 0 Then iRockIndex = iRockIndex + 1
        Loop
        
        txtIn.Close
        
        ReDim RockIndex(1 To iRockIndex) As String
        
        Set txtIn = fs.OpenTextFile(strDir & "rock.in", ForReading)
        
        Do While Not txtIn.AtEndOfStream
            strLine = txtIn.ReadLine
        
        'Get filename of a rockfile for this rock index
            If InStr(1, strLine, "FileName"": {") > 0 And blnNameFound = False Then
                strLine = txtIn.ReadLine
                strLine = txtIn.ReadLine
                
                If InStr(1, strLine, "BV_") > 0 Then
                    strRockName = Mid(strLine, InStrRev(strLine, "BV") + 3, Len(strLine) - InStrRev(strLine, "BV") - 3)
                ElseIf InStr(1, strLine, "_STINGER_") > 0 Then
                    If InStr(1, strLine, "_V3") > 0 Then strLine = Replace(strLine, "_V3", "")
                    If InStr(1, strLine, "_V2") > 0 Then strLine = Replace(strLine, "_V2", "")
                    strRockName = Mid(strLine, InStrRev(strLine, "STINGER") + 8, Len(strLine) - InStrRev(strLine, "STINGER") - 8)
                ElseIf InStr(1, strLine, "9694_") Then
                    strRockName = Mid(strLine, InStrRev(strLine, "0SR") + 4, Len(strLine) - InStrRev(strLine, "0SR") - 4)
                End If
                
                blnNameFound = True
                
        'Put the rock names in the RockIndex
            ElseIf InStr(1, strLine, "RockIndex"": {") > 0 Then
                strLine = txtIn.ReadLine
                strLine = txtIn.ReadLine
    
                iRockIndex = Trim(Right(strLine, Len(strLine) - InStrRev(strLine, ":")))
                
                RockIndex(iRockIndex) = strRockName
                
                blnNameFound = False
                
            End If
        Loop
        
        txtIn.Close
    End If
    
    
    strFile = strDir & "calc.in"
    
    If Dir(strFile) <> "" Then
        Set txtIn = fs.OpenTextFile(strFile, ForReading)
        i = 0
        blnNameFound = False
        
        Do While Not txtIn.AtEndOfStream
            strLine = txtIn.ReadLine
        
            If InStr(1, strLine, "Regular"": [") > 0 Then
                blnNameFound = True
            ElseIf InStr(1, strLine, "RockIndex"": {") > 0 And blnNameFound Then
                i = i + 1
                
                strLine = txtIn.ReadLine
                strLine = txtIn.ReadLine
                strLine = txtIn.ReadLine
                strLine = txtIn.ReadLine
                
            'Determine which rock is active for the case.
                X = Trim(Right(strLine, Len(strLine) - InStrRev(strLine, ":")))
                
                ROCK(i) = Left(RockIndex(X), InStr(1, RockIndex(X), "_") - 1)
                Conf_Pressure(i) = Right(RockIndex(X), Len(RockIndex(X)) - InStrRev(RockIndex(X), "_"))
                
            End If
    
        Loop
        
        Set txtIn = Nothing
    End If
    
    Set fs = Nothing
    
    
    Selection.ClearContents

    Sheets("BIT" & iBit & "_SUMMARY").Visible = True

    For i = 1 To iCases
        iSummCol = 1
        
        For iFile = 1 To Range("StingerFiles").Count
            strFile = strDir & "utility_files\NormalCase" & i & Range("StingerFiles").Rows(iFile)
            
            If Dir(strFile) = "" Then GoTo NextFile
            
            strSheet = OpenText(strFile, "query")
            
            If InStr(1, Range("StingerFiles").Rows(iFile), "cutter") Then
                iCol = 1
                
                Range("B2").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                
                Do While IsEmpty(Cells(3, iCol))
                    iCol = iCol + 1
                Loop
                
                iCutters = Cells(3, iCol)
                
                Range(Cells(3, iCol), Cells(2 + iCutters, iCol + 7)).Select
                
                With ActiveWorkbook.Worksheets(strSheet).Sort
                    .SortFields.Add Key:=Range("B3:B" & 2 + iCutters), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                    .SetRange Range("B3:I" & 2 + iCutters)
'                    .Header = xlGuess
'                    .MatchCase = False
'                    .Orientation = xlTopToBottom
'                    .SortMethod = xlPinYin
                    .Apply
                End With
                
                Range(Cells(1, 1), Cells(2 + iCutters, iCol + 7)).Copy
                

                Sheets("BIT" & iBit & "_SUMMARY").Select
                
                
                
                Range(fXLNumberToLetter(CLng(iSummCol)) & "1").PasteSpecial xlPasteValues
                

                
                iSummCol = iSummCol + Selection.Columns.Count + 1
                
            ElseIf InStr(1, Range("StingerFiles").Rows(iFile), "imbf") Then
                
                iCol = 1
                Do Until Cells(2, iCol) = "Revolution(Revs)"
                    iCol = iCol + 1
                Loop
                
                'Sort Data
                Range("B" & iCol).Select
                Range(Selection, Selection.End(xlToRight)).Select
                Range(Selection, Selection.End(xlDown)).Select
                
                strLine = Selection.Address
                
                ActiveWorkbook.Worksheets(strSheet).Sort.SortFields.Clear
                ActiveWorkbook.Worksheets(strSheet).Sort.SortFields.Add Key:=Range(fXLNumberToLetter(iCol + 3) & "3:" & fXLNumberToLetter(iCol + 3) & Mid(strLine, InStrRev(strLine, "$") + 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets(strSheet).Sort
                    .SetRange Range(Replace(strLine, "$", ""))
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                
                k = Selection.Rows.Count - 1

                If k Mod 2 = 0 Then
                    'Even

                    RIF(i) = Format((Cells(Int(k / 2) + 2, iCol + 1) + Cells(Int(k / 2) + 1 + 2, iCol + 1)) / 2, "0.0000")
                    CIF(i) = Format((Cells(Int(k / 2) + 2, iCol + 2) + Cells(Int(k / 2) + 1 + 2, iCol + 2)) / 2, "0.0000")
                    TIF(i) = Format((Cells(Int(k / 2) + 2, iCol + 3) + Cells(Int(k / 2) + 1 + 2, iCol + 3)) / 2, "0.0000")
                    BETA(i) = Format((Cells(Int(k / 2) + 2, iCol + 4) + Cells(Int(k / 2) + 1 + 2, iCol + 4)) / 2, "0.0000")
                    RIF_CIF(i) = Format(RIF(i) / CIF(i), "0.0000")
                    If RIF_CIF(i) > 1 Then
                        RIF_CIF_Normal(i) = Format(1 / RIF_CIF(i), "0.0000")
                    Else
                        RIF_CIF_Normal(i) = RIF_CIF(i)
                    End If

                Else
                    'Odd
                    
                    RIF(i) = Cells(Int(k / 2) + 1 + 2, iCol + 1)
                    CIF(i) = Cells(Int(k / 2) + 1 + 2, iCol + 2)
                    TIF(i) = Cells(Int(k / 2) + 1 + 2, iCol + 3)
                    BETA(i) = Cells(Int(k / 2) + 1 + 2, iCol + 4)
                    RIF_CIF(i) = Format(RIF(i) / CIF(i), "0.0000")
                    If RIF_CIF(i) > 1 Then
                        RIF_CIF_Normal(i) = Format(1 / RIF_CIF(i), "0.0000")
                    Else
                        RIF_CIF_Normal(i) = RIF_CIF(i)
                    End If
                    
                End If

            End If
            
            Application.DisplayAlerts = False
            Sheets(strSheet).Delete
            Application.DisplayAlerts = True
NextFile:
        Next iFile
        
        Range(Cells(1, 1), Cells(2 + iCutters, iSummCol)).Copy
        
        Sheets("BIT" & iBit).Select
        
        iBitCol = Range(ActiveSheet.UsedRange.Address).Columns.Count
        
        Cells(4, iBitCol + 2) = "FCirc"
        Cells(4, iBitCol + 3) = "FCircWR"
        Cells(4, iBitCol + 4) = "FN"
        Cells(4, iBitCol + 5) = "FNWR"
        
        Cells(3, iBitCol + 7).Select
        Selection.PasteSpecial xlPasteValues
        
        strArray = Array("Circumferential", "Circumferential_Workrate", "Normal", "Normal_Workrate")
        
        For iFile = 1 To 4
            Range(Cells(5, iBitCol + iFile + 1), Cells(4 + iCutters, iBitCol + iFile + 1)).Name = "Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1)
            
            Range(Cells(5, iBitCol + 9 + 10 * (iFile - 1)), Cells(4 + iCutters, iBitCol + 9 + 10 * (iFile - 1))).Name = "Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1) & "_ST0"
            Range(Cells(5, iBitCol + 10 + 10 * (iFile - 1)), Cells(4 + iCutters, iBitCol + 10 + 10 * (iFile - 1))).Name = "Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1) & "_ST5"
            Range(Cells(5, iBitCol + 11 + 10 * (iFile - 1)), Cells(4 + iCutters, iBitCol + 11 + 10 * (iFile - 1))).Name = "Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1) & "_ST25"
            Range(Cells(5, iBitCol + 12 + 10 * (iFile - 1)), Cells(4 + iCutters, iBitCol + 12 + 10 * (iFile - 1))).Name = "Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1) & "_ST50"
            Range(Cells(5, iBitCol + 13 + 10 * (iFile - 1)), Cells(4 + iCutters, iBitCol + 13 + 10 * (iFile - 1))).Name = "Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1) & "_ST75"
            Range(Cells(5, iBitCol + 14 + 10 * (iFile - 1)), Cells(4 + iCutters, iBitCol + 14 + 10 * (iFile - 1))).Name = "Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1) & "_ST95"
            Range(Cells(5, iBitCol + 15 + 10 * (iFile - 1)), Cells(4 + iCutters, iBitCol + 15 + 10 * (iFile - 1))).Name = "Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1) & "_ST100"
        
            For X = 1 To Range("Bit" & iBit & "_Cutter_Num").Count
                Range("Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1)).Cells(X) = Range("Bit" & iBit & "_Case" & i & "_" & strArray(iFile - 1) & "_ST" & Range("StingerPercentile")).Cells(X)
            Next X
        Next iFile
        
DoEvents

    Next i

    
    Sheets("BIT" & iBit & "_SUMMARY").Visible = False

    Exit Function
    
ErrHandler:

Debug.Print Now(), gProc, Err, Err.Description
Stop
Resume


End Function


Public Sub UpdateStingerStaticPercentile(Optional strPerc As String, Optional iPerc As Integer)

    gProc = "UpdateStingerStaticPercentile"
    
    Dim strArray As Variant
    Dim iBit As Integer, iCut As Integer, iVal As Integer, iCase As Integer
    Dim strActiveSheet As String
    
    
    Application.ScreenUpdating = False
    
    If strPerc = "" Then strPerc = Range("StingerPercentile")
    
    strActiveSheet = ActiveSheet.Name
        
    strArray = Array("Circumferential", "Circumferential_Workrate", "Normal", "Normal_Workrate")
        
    For iBit = 1 To Range("MaxBits")
        If Range("Bit" & iBit & "_StingerStatic") = "Yes" Then
        
            Sheets("BIT" & iBit).Activate
            
            For iCase = 1 To Range("Bit" & iBit & "_Cases")
                For iVal = 0 To UBound(strArray)
                    For iCut = 1 To Range("Bit" & iBit & "_Cutter_Num").Count
                        Range("Bit" & iBit & "_Case" & iCase & "_" & strArray(iVal)).Cells(iCut) = Range("Bit" & iBit & "_Case" & iCase & "_" & strArray(iVal) & "_ST" & strPerc).Cells(iCut)
                    Next iCut
                Next iVal
            Next iCase
        End If
    Next iBit
    
    Application.ScreenUpdating = True
    
    Sheets(strActiveSheet).Select
    
    If iPerc > 0 Then Range("StingerPercentileRow") = iPerc

    DoEvents
    
End Sub


Function GetWearCoef(BitNum, strCalcFile As String)
    Dim strTempWB As String, iCol As Integer
    
    On Error GoTo ErrHandler
    
    If Range("ReadVer") = "old" Then
        Workbooks.OpenText FileName:=strCalcFile, Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=True, Other:=True, TrailingMinusNumbers:=True
        strTempWB = ActiveWorkbook.Name
    Else
        strTempWB = OpenText(strCalcFile, Range("ReadVer"))
    End If
    
    If Range("A1") = "{" Then
        With ActiveSheet.UsedRange
'A1
            Set find_coeff = .Find("A1:", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_coeff Is Nothing Then A1 = find_coeff.Offset(4, 1)
'b1
            Set find_coeff = .Find("B1:", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_coeff Is Nothing Then b1 = find_coeff.Offset(4, 1)
'c0
            Set find_coeff = .Find("C0:", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_coeff Is Nothing Then c0 = find_coeff.Offset(4, 1)
'c1
            Set find_coeff = .Find("C1:", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not find_coeff Is Nothing Then c1 = find_coeff.Offset(4, 1)
        End With
    Else
        c0 = Range("A3").Value
        c1 = Range("B3").Value
        A1 = Range("C3").Value
        b1 = Range("D3").Value
    End If
    
    Application.DisplayAlerts = False
    Workbooks(Current_Workbook).Activate
    If Range("ReadVer") = "old" Then
        Windows(strTempWB).Close
        Workbooks(Current_Workbook).Activate
        
    Else
        Sheets(strTempWB).Delete
        ClearQueryTables
    End If
    Application.DisplayAlerts = True
    Sheets("INFO").Activate
    
    
    row_index = Range("Summary_Info").row
    row_index = row_index + 2 + CInt(BitNum - 1) * 2
    
    iCol = Range(Range("Summary_Info").row + 1 & ":" & Range("Summary_Info").row + 1).Find("c0").Column
    
    Cells(Trim(str(row_index)), iCol).Value = c0
    Cells(Trim(str(row_index)), iCol + 1).Value = c1
    Cells(Trim(str(row_index)), iCol + 2).Value = A1
    Cells(Trim(str(row_index)), iCol + 3).Value = b1
    
    Exit Function
    
ErrHandler:
    
    Debug.Print Now(), gProc, Err, Err.Description
    Stop
    Resume
    
End Function

Function CellFormat1(Cell_Address)
    Range(Cell_Address).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Function
Function Clear_Cutter_Info(BitNum)

    gProc = "Clear_Cutter_Info"
    
    Dim row_index As Integer
    'Clears BIT<BitNum>_CUTTER_FILE, BIT<BitNum> and INFO tabs
    Sheets("BIT" & BitNum & "_CUTTER_FILE").Activate
'    ActiveSheet.UsedRange.ClearContents
    Reset_UsedRange
    
    Sheets("BIT" & BitNum).Activate
    Reset_UsedRange
    ActiveSheet.UsedRange.Select
    'Dim Name_Range As Name
    'For Each Name_Range In ThisWorkbook.Names
    '    If Range(Name_Range).Parent.Name = "BIT" & BitNum Then
    '        Name_Range.Delete
    '    End If
    'Next Name_Range
'    ActiveSheet.UsedRange.ClearContents
'    ActiveSheet.UsedRange.ClearFormats

    Sheets("INFO").Activate
    row_index = Range("Cutting_Structure_Info").row
    row_index = row_index + 2 + CInt(BitNum - 1) * 4
    Range("C" & Trim(str(row_index)) & ":G" & Trim(str(row_index + 3))).ClearContents
    Range("I" & Trim(str(row_index)) & ":" & fXLNumberToLetter(Range("CS_SRHigh").Column) & Trim(str(row_index + 3))).ClearContents
    Range(fXLNumberToLetter(Range("PI_Agg").Column + 1) & Trim(str(row_index)) & ":" & fXLNumberToLetter(Range("Bit" & BitNum & "_GPBeta").Column) & Trim(str(row_index + 3))).ClearContents
    Range("Bit" & BitNum & "_Location").Value = ""
End Function
Function Clear_Summary(BitNum)

    gProc = "Clear_Summary"
    
    Dim lCol As Long

    Sheets("BIT" & BitNum & "_SUMMARY").Activate
    Reset_UsedRange

    Sheets("INFO").Activate
    row_index = Range("Summary_Info").row
    iCol = Range(Range("Summary_Info").row + 1 & ":" & Range("Summary_Info").row + 1).Find("b1").Column

    row_index = row_index + 2 + CInt(BitNum - 1) * 2
    Range("C" & Trim(str(row_index)) & ":" & fXLNumberToLetter(lCol) & Trim(str(row_index + 1))).ClearContents
    
    Sheets("SUMMARY").Activate
    row_index = Range("Case_Info").row + 3
    case_column_index = Range("Case_Info").Column
    control_column_index = Range("Control_Info").Column
    ROP_column_index = Range("ROP_Compare").Column
    Cum_Hrs_index = Range("Cumulative_Hours").Column
    Cum_Foot_index = Range("Cumulative_Footage").Column
    DOC_column_index = Range("DOC_info").Column
    Torque_index = Range("TORQUE_info").Column
    TIF_index = Range("TIF_info").Column
'    Torque_index = Range("TORQUE_info").Column
    TQ_Compare_index = Range("TQ_Compare").Column
    TIF_index = Range("TIF_info").Column
    RIF_index = Range("RIF_info").Column
    CIF_index = Range("CIF_info").Column
    RIF_CIF_index = Range("RIF_CIF_info").Column
    RIF_CIF_Normal_index = Range("RIF_CIF_Normal_info").Column
    BETA_index = Range("BETA_info").Column
    ROP_TORQUE_index = Range("ROP_TORQUE_info").Column
    ROP_TORQUE_Compare_index = Range("ROP_TORQUE_Compare").Column
    
    Interpolate_index = Range("Interpolate_Info").Column
    bit_offset = BitNum
    i = 1
    Do Until Cells(row_index + i - 1, case_column_index + bit_offset).Value = ""
        i = i + 1
    Loop
    If i > 1 Then
        Cells(row_index, case_column_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, control_column_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        If Not BitNum = 1 Then
            Cells(row_index, ROP_column_index + bit_offset - 1).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Else
            Cells(row_index, ROP_column_index + 1).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 3)).ClearContents
        End If
        Cells(row_index, DOC_column_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, Cum_Hrs_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, Cum_Foot_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, Torque_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        If Not BitNum = 1 Then
            Cells(row_index, TQ_Compare_index + bit_offset - 1).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Else
            Cells(row_index, TQ_Compare_index + 1).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 3)).ClearContents
        End If
        Cells(row_index, TIF_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, RIF_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, CIF_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, RIF_CIF_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, RIF_CIF_Normal_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, BETA_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Cells(row_index, ROP_TORQUE_index + bit_offset).Select
        Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        If Not BitNum = 1 Then
            Cells(row_index, ROP_TORQUE_Compare_index + bit_offset - 1).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        Else
            Cells(row_index, ROP_TORQUE_Compare_index + 1).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 3)).ClearContents
        End If
    End If
    
    Cells(row_index, Interpolate_index + bit_offset).Select
    Range(ActiveCell, ActiveCell.Offset(2, 0)).ClearContents
       
    If Range("Bit1_Location").Value = "" And Range("Bit2_Location").Value = "" And Range("Bit3_Location").Value = "" And Range("Bit4_Location").Value = "" And Range("Bit5_Location").Value = "" Then
        i = 1
        Do Until Cells(row_index + i - 1, case_column_index).Value = ""
            i = i + 1
        Loop
        If i > 1 Then
            Cells(row_index, case_column_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, control_column_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, ROP_column_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, DOC_column_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, Cum_Hrs_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, Cum_Foot_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, Torque_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, TQ_Compare_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, TIF_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, RIF_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, CIF_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, RIF_CIF_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, RIF_CIF_Normal_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, BETA_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, ROP_TORQUE_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
            Cells(row_index, ROP_TORQUE_Compare_index).Select
            Range(ActiveCell, ActiveCell.Offset(i - 2, 0)).ClearContents
        End If

'        Sheets("FORCES").Activate
'        Range("A1:A" & Trim(Str((2 * i - 1) * 11 + 1))).ClearContents
'
'        Sheets("CUTTER_FORCES").Activate
'        Range("A1:A" & Trim(Str((2 * i - 1) * 11 + 1))).ClearContents
    End If
    
    Sheets("HOME").Activate

End Function
Function CutterDia(Cutter As String, Optional blnSize As Boolean)
    Dim Cutter_Size As Integer
    
    gProc = "CutterDia"
    
    'Finds the cutter diameter based on the cutter size

    If Left(Cutter, 1) = "G" Then 'Axe Cutter
        Cutter_Size = CInt(Mid(Cutter, 2, 2))
    ElseIf Left(Cutter, 1) = "P" Then 'PX/G188/Axe TR Cutter
        Cutter_Size = CInt(Mid(Cutter, 2, 2))
    ElseIf Left(Cutter, 1) = "J" Or Left(Cutter, 1) = "K" Then 'PX/Axe Arrow Cutter
        Cutter_Size = CInt(Mid(Cutter, 2, 2))
    ElseIf Left(Cutter, 1) = "D" Or Left(Cutter, 1) = "Q" Then 'Delta/Hyper Cutter
        Cutter_Size = CInt(Mid(Cutter, 2, 2))
    ElseIf Left(Cutter, 1) = "M" Then 'Echo Cutter
        Cutter_Size = CInt(Mid(Cutter, 2, 2))
    Else
        If Len(Cutter) = 3 Then
            Cutter_Size = CInt(Left(Cutter, 1))
        Else
            Cutter_Size = CInt(Mid(Cutter, 1, 2))
        End If
    End If

    If blnSize Then
        CutterDia = Cutter_Size
        Exit Function
    End If
    
    Select Case Cutter_Size
        Case 6
            CutterDia = 0.236
        Case 9
            CutterDia = 0.371
        Case 11
            CutterDia = 0.44
        Case 13
            CutterDia = 0.529
        Case 16
            CutterDia = 0.625
        Case 19
            CutterDia = 0.75
        Case 22
            CutterDia = 0.866
        Case Else
            If Cutter = "" Then
                'if stinger
            Else
                CutterDia = Cutter
            End If
    End Select
    
    
End Function
Function RoundOff(Num, Dig)
    'Round off <Num> to <Dig> decimal places
    Dim temp As Double
    temp = 10 ^ Dig
    RoundOff = CLng(Num * temp) / temp
End Function
Function ConvertToK(Num)

    gProc = "ConvertToK"
    
    'Converts to Kilos
    Dim temp As Double
    temp = Num / 1000
    temp = RoundOff(temp, 1)
    Dim temp2 As String
    temp2 = Trim(str(temp))
    If InStr(1, temp2, ".") = 0 Then
        ConvertToK = temp2 ' & "K"
    Else
        ConvertToK = Mid(temp2, 1, (InStr(1, temp2, ".") + 1)) ' & "K"
    End If
End Function
Function FindCutters(Row_Start, Row_End)

    gProc = "FindCutters"
    
    
    'Finds different type of cutters, their quantities and puts them in a string
    Dim Cutter_Type() As String
    ReDim Preserve Cutter_Type(1 To (Row_End - Row_Start + 1)) As String
    For i = 1 To (Row_End - Row_Start + 1)
        Cutter_Type(i) = Range("I" & Trim(str(i + Row_Start - 1))).Value
    Next i
    On Error Resume Next
    For i = 1 To (UBound(Cutter_Type) - 1)
        For j = i + 1 To UBound(Cutter_Type)
            If CStr(Cutter_Type(j)) > CStr(Cutter_Type(i)) Then
                temp_str = Cutter_Type(j)
                Cutter_Type(j) = Cutter_Type(i)
                Cutter_Type(i) = temp_str
            End If
        Next j
    Next i
    i = 1
    temp_str = ""
    Do
        temp_cnt = 1
        For j = i + 1 To UBound(Cutter_Type)
            If Cutter_Type(j) = Cutter_Type(i) Then
                temp_cnt = temp_cnt + 1
            Else
                Exit For
            End If
        Next j
        temp_str = temp_str & Cutter_Type(i) & " (" & Trim(str(temp_cnt)) & "), "
        i = j
    Loop Until i > UBound(Cutter_Type)
    temp_str = Left(temp_str, Len(temp_str) - 2)
    FindCutters = temp_str
End Function
Function HideCutters(BitNum, Backups, hide, Off)

    gProc = "HideCutters"
    
    Dim iSeriesCount As Integer, iSeries As Integer, iBitNum As Integer, start_row_index As Integer, Last_row As Integer, i As Integer
    Dim strFormula(3) As String, strCol(4) As String, Cutting_Structure_index As Integer, Summary_index As Integer
    Dim blnHidden(4) As Boolean, hide_flag As Boolean
    Dim case_column_index As Integer, control_column_index As Integer, DOC_column_index As Integer, ROP_column_index As Integer, Cum_Hrs_index As Integer, Cum_Foot_index As Integer, Torque_index As Integer, TIF_index As Integer, RIF_index As Integer, CIF_index As Integer, RIF_CIF_index As Integer, RIF_CIF_Normal_index As Integer, BETA_index As Integer, Interpolate_index As Integer
    
    On Error GoTo ErrHandler
    
    subUpdate False
    
    Sheets("BIT" & BitNum).Activate
    ActiveSheet.Cells.EntireRow.Hidden = False
    start_row_index = 4
    Last_row = Range(ActiveSheet.UsedRange.Address).Rows.Count
    If hide Then
        If Backups Then
            If Last_row > start_row_index Then
                For i = start_row_index + 1 To Last_row
                    If CInt(Range("P" & CStr(i)).Value) = 1 Or CInt(Range("P" & CStr(i)).Value) = 3 Then
                        If Not Off Then
                            Range("P" & CStr(i)).EntireRow.Hidden = True
                        Else
                            If CDbl(Range("K" & CStr(i)).Value) < 0 Then
                                Range("P" & CStr(i)).EntireRow.Hidden = True
                            End If
                        End If
                    End If
                Next i
            Else
                HideCutters = False
            End If
        Else
            ActiveSheet.Cells.EntireRow.Hidden = True
        End If
    End If
    Sheets("SUMMARY").Activate
    'row_index = Range("Case_Info").Row + 3
    case_column_index = Range("Case_Info").Column
    control_column_index = Range("Control_Info").Column
    DOC_column_index = Range("DOC_Info").Column
    ROP_column_index = Range("ROP_Compare").Column
    Cum_Hrs_index = Range("Cumulative_Hours").Column
    Cum_Foot_index = Range("Cumulative_Footage").Column
    Torque_index = Range("TORQUE_info").Column
    TQ_Compare_index = Range("TQ_Compare").Column
    TIF_index = Range("TIF_info").Column
    RIF_index = Range("RIF_info").Column
    CIF_index = Range("CIF_info").Column
    RIF_CIF_index = Range("RIF_CIF_info").Column
    RIF_CIF_Normal_index = Range("RIF_CIF_Normal_info").Column
    BETA_index = Range("BETA_info").Column
    ROP_TORQUE_index = Range("ROP_TORQUE_info").Column
    ROP_TORQUE_Compare_index = Range("ROP_TORQUE_Compare").Column
    Interpolate_index = Range("Interpolate_Info").Column
    If hide And Not Backups And Not Off Then
        If Not BitNum = 1 Then
            Cells(1, ROP_column_index + BitNum - 1).EntireColumn.Hidden = True
        Else
            Cells(1, ROP_column_index + 1).EntireColumn.Hidden = True
            Cells(1, ROP_column_index + 2).EntireColumn.Hidden = True
            Cells(1, ROP_column_index + 3).EntireColumn.Hidden = True
            Cells(1, ROP_column_index + 4).EntireColumn.Hidden = True
        End If
        Cells(1, case_column_index + BitNum).EntireColumn.Hidden = True
        Cells(1, DOC_column_index + BitNum).EntireColumn.Hidden = True
        Cells(1, ROP_column_index + BitNum - 1).EntireColumn.Hidden = True
        Cells(1, control_column_index + BitNum).EntireColumn.Hidden = True
        Cells(1, Cum_Hrs_index + BitNum).EntireColumn.Hidden = True
        Cells(1, Cum_Foot_index + BitNum).EntireColumn.Hidden = True
        Cells(1, Torque_index + BitNum).EntireColumn.Hidden = True
        If Not BitNum = 1 Then
            Cells(1, TQ_Compare_index + BitNum - 1).EntireColumn.Hidden = True
        Else
            Cells(1, TQ_Compare_index + 1).EntireColumn.Hidden = True
            Cells(1, TQ_Compare_index + 2).EntireColumn.Hidden = True
            Cells(1, TQ_Compare_index + 3).EntireColumn.Hidden = True
            Cells(1, TQ_Compare_index + 4).EntireColumn.Hidden = True
        End If
        Cells(1, TIF_index + BitNum).EntireColumn.Hidden = True
        Cells(1, RIF_index + BitNum).EntireColumn.Hidden = True
        Cells(1, CIF_index + BitNum).EntireColumn.Hidden = True
        Cells(1, RIF_CIF_index + BitNum).EntireColumn.Hidden = True
        Cells(1, RIF_CIF_Normal_index + BitNum).EntireColumn.Hidden = True
        Cells(1, BETA_index + BitNum).EntireColumn.Hidden = True
        Cells(1, ROP_TORQUE_index + BitNum).EntireColumn.Hidden = True
        If Not BitNum = 1 Then
            Cells(1, ROP_TORQUE_Compare_index + BitNum - 1).EntireColumn.Hidden = True
        Else
            Cells(1, ROP_TORQUE_Compare_index + 1).EntireColumn.Hidden = True
            Cells(1, ROP_TORQUE_Compare_index + 2).EntireColumn.Hidden = True
            Cells(1, ROP_TORQUE_Compare_index + 3).EntireColumn.Hidden = True
            Cells(1, ROP_TORQUE_Compare_index + 4).EntireColumn.Hidden = True
        End If
        Cells(1, Interpolate_index + BitNum).EntireColumn.Hidden = True
    Else
        If Not BitNum = 1 Then
            Cells(1, ROP_column_index + BitNum - 1).EntireColumn.Hidden = False
        Else
            Cells(1, ROP_column_index + 1).EntireColumn.Hidden = False
            Cells(1, ROP_column_index + 2).EntireColumn.Hidden = False
            Cells(1, ROP_column_index + 3).EntireColumn.Hidden = False
            Cells(1, ROP_column_index + 4).EntireColumn.Hidden = False
        End If
        Cells(1, case_column_index + BitNum).EntireColumn.Hidden = False
        Cells(1, DOC_column_index + BitNum).EntireColumn.Hidden = False
        Cells(1, ROP_column_index + BitNum - 1).EntireColumn.Hidden = False
        Cells(1, control_column_index + BitNum).EntireColumn.Hidden = False
        Cells(1, Cum_Hrs_index + BitNum).EntireColumn.Hidden = False
        Cells(1, Cum_Foot_index + BitNum).EntireColumn.Hidden = False
        Cells(1, Torque_index + BitNum).EntireColumn.Hidden = False
        If Not BitNum = 1 Then
            Cells(1, TQ_Compare_index + BitNum - 1).EntireColumn.Hidden = False
        Else
            Cells(1, TQ_Compare_index + 1).EntireColumn.Hidden = False
            Cells(1, TQ_Compare_index + 2).EntireColumn.Hidden = False
            Cells(1, TQ_Compare_index + 3).EntireColumn.Hidden = False
            Cells(1, TQ_Compare_index + 4).EntireColumn.Hidden = False
        End If
        Cells(1, TIF_index + BitNum).EntireColumn.Hidden = False
        Cells(1, RIF_index + BitNum).EntireColumn.Hidden = False
        Cells(1, CIF_index + BitNum).EntireColumn.Hidden = False
        Cells(1, RIF_CIF_index + BitNum).EntireColumn.Hidden = False
        Cells(1, RIF_CIF_Normal_index + BitNum).EntireColumn.Hidden = False
        Cells(1, BETA_index + BitNum).EntireColumn.Hidden = False
        Cells(1, ROP_TORQUE_index + BitNum).EntireColumn.Hidden = False
        If Not BitNum = 1 Then
            Cells(1, ROP_TORQUE_Compare_index + BitNum - 1).EntireColumn.Hidden = False
        Else
            Cells(1, ROP_TORQUE_Compare_index + 1).EntireColumn.Hidden = False
            Cells(1, ROP_TORQUE_Compare_index + 2).EntireColumn.Hidden = False
            Cells(1, ROP_TORQUE_Compare_index + 3).EntireColumn.Hidden = False
            Cells(1, ROP_TORQUE_Compare_index + 4).EntireColumn.Hidden = False
        End If
        Cells(1, Interpolate_index + BitNum).EntireColumn.Hidden = False
    End If
    
    Sheets("INFO").Activate
    Cutting_Structure_index = Range("Cutting_Structure_Info").row
    Cutting_Structure_index = Cutting_Structure_index + 2 + CInt(BitNum - 1) * 4
    Summary_index = Range("Summary_Info").row
    Summary_index = Summary_index + 2 + CInt(BitNum - 1) * 2
    
    If hide And Not Backups And Not Off Then
        hide_flag = True
    Else
        hide_flag = False
    End If
    For i = 1 To 4
        Cells(Cutting_Structure_index + i - 1, 1).EntireRow.Hidden = hide_flag
    Next i
    For i = 1 To 2
        Cells(Summary_index + i - 1, 1).EntireRow.Hidden = hide_flag
    Next i
    'subUpdate True
    
    Sheets("Constants").Activate
    Range("Bit" & BitNum & "_PrimaryCutters").Columns.Hidden = hide_flag
    
    Sheets("FORCE_RATIOS").Activate
    Range("Bit" & BitNum & "_ForceRatios").Columns.Hidden = hide_flag
    
    'Gets turned off when you load a new bit. must be turned on elsewhere (calculations occur at this time). -AB 2018-10-25
'    Sheets("ZONE_OF_INTEREST").Activate
'    Range("Bit" & BitNum & "_Sum_ZNormal").Rows.Hidden = hide_flag
    
    Sheets("INFO2").Activate
    Range("Bit" & BitNum & "_I2_Name").Columns.Hidden = hide_flag
    
    Sheets("Drill_Scan").Activate
    Range("Bit" & BitNum & "_DS_Name").Columns.Hidden = hide_flag
    
    Sheets("Home").Activate
        
    Exit Function

ErrHandler:

    Stop

    Debug.Print Err, Err.Description

    Resume
    

End Function

Sub Wear_plots()

    gProc = "Wear_plots"
    
'    Dim Bit1() As Single
'    Dim Bit2() As Single
'    Dim Bit3() As Single
'    Dim Bit4() As Single
'    Dim Bit5() As Single
    Dim Bit() As Single
    Dim strRange(8) As String
    Dim iCutter As Integer, iPlot As Integer, info_last_col As Integer, info_last_row As Integer, i As Integer, strErr As String
    Dim Heading_Array1 As Variant
    Dim Heading_Array2 As Variant
    Dim str(7) As String
    Dim strChrt As Variant
    Dim strCase As String
    
    On Error GoTo ErrHandler

    If ActiveSheet.Shapes("Wear_Hours_Selection").ControlFormat.Value = 1 Then
        input_control = 1
    ElseIf ActiveSheet.Shapes("Wear_Footage_Selection").ControlFormat.Value = 1 Then
        input_control = 2
    ElseIf ActiveSheet.Shapes("Wear_ROP_Selection").ControlFormat.Value = 1 Then
        input_control = 3
    Else
        MsgBox "No Wear Plot option selected! Please define your selection.", vbCritical + vbOKOnly, "No option selected!"
        Exit Sub
    End If

    For i = 1 To Range("MaxBits")
        If Sheets("HOME").Shapes("Hide_Bit" & i).ControlFormat.Value <> 1 Then
            If Range("Bit" & i & "_WearProj") <> "Yes" Then
                strErr = strErr & "Bit " & i & ", "
            End If
        End If
    Next i
    
    If strErr <> "" Then
        MsgBox Left(strErr, Len(strErr) - 2) & " are not Wear Projects." & vbCrLf & "Please load wear projects to use this tool.", vbInformation + vbOKOnly, "Wear Projects Error"
        Exit Sub
    End If
      
    
    subUpdate False
    
    
    strRange(1) = "_AreaCut"
    strRange(2) = "_VolCut"
    strRange(3) = "_Circumferential"
    strRange(4) = "_Vertical"
    strRange(5) = "_Normal"
    strRange(6) = "_WearFlatArea"
    strRange(7) = "_Circumferential_Workrate"
    strRange(8) = "_Normal_Workrate"

    input_data = Range("Wear_Input").Value
    Sheets("SUMMARY").Activate
    row_index = Range("ROP_Compare").row + 2
    For BitNum = 1 To 5
        If Not Range("Bit" & BitNum & "_Name").Value = "" Then
            ROP_column_index = Range("Control_Info").Column + BitNum
            Cum_Hrs_index = Range("Cumulative_Hours").Column + BitNum
            Cum_Foot_index = Range("Cumulative_Footage").Column + BitNum
            Case_count = Range("Bit" & BitNum & "_CUM_FOOT").Rows.Count
            i = 1
            If input_control = 1 Then
                Do Until i > Case_count Or Cells(row_index + i, Cum_Hrs_index).Value >= input_data
                    i = i + 1
                Loop
            Else
                If input_control = 2 Then
                    Do Until i > Case_count Or Cells(row_index + i, Cum_Foot_index).Value >= input_data
                        i = i + 1
                    Loop
                Else
                    Do Until i > Case_count Or Cells(row_index + i, ROP_column_index).Value <= input_data
                        i = i + 1
                    Loop
                End If
            End If
            If i > Case_count Then i = Case_count
'            Debug.Print "Wear Info: Bit " & BitNum & " Case limits are " & i - 1 & " & " & i
            If i = 1 Then i = 2
            Interpolate_index = Range("Interpolate_Info").Column + BitNum
            Cells(row_index + input_control, Interpolate_index).Value = input_data
            
            Select Case input_control
                Case 1:
                    Cells(row_index + 2, Interpolate_index).Value = LinInterpolate(Cells(row_index + i - 1, Cum_Hrs_index).Value, Cells(row_index + i, Cum_Hrs_index).Value, input_data, Cells(row_index + i - 1, Cum_Foot_index).Value, Cells(row_index + i, Cum_Foot_index).Value)
                    Cells(row_index + 3, Interpolate_index).Value = LinInterpolate(Cells(row_index + i - 1, Cum_Hrs_index).Value, Cells(row_index + i, Cum_Hrs_index).Value, input_data, Cells(row_index + i - 1, ROP_column_index).Value, Cells(row_index + i, ROP_column_index).Value)
                Case 2:
                    Cells(row_index + 1, Interpolate_index).Value = LinInterpolate(Cells(row_index + i - 1, Cum_Foot_index).Value, Cells(row_index + i, Cum_Foot_index).Value, input_data, Cells(row_index + i - 1, Cum_Hrs_index).Value, Cells(row_index + i, Cum_Hrs_index).Value)
                    Cells(row_index + 3, Interpolate_index).Value = LinInterpolate(Cells(row_index + i - 1, Cum_Foot_index).Value, Cells(row_index + i, Cum_Foot_index).Value, input_data, Cells(row_index + i - 1, ROP_column_index).Value, Cells(row_index + i, ROP_column_index).Value)
                Case 3:
                    Cells(row_index + 1, Interpolate_index).Value = LinInterpolate(Cells(row_index + i - 1, ROP_column_index).Value, Cells(row_index + i, ROP_column_index).Value, input_data, Cells(row_index + i - 1, Cum_Hrs_index).Value, Cells(row_index + i, Cum_Hrs_index).Value)
                    Cells(row_index + 2, Interpolate_index).Value = LinInterpolate(Cells(row_index + i - 1, ROP_column_index).Value, Cells(row_index + i, ROP_column_index).Value, input_data, Cells(row_index + i - 1, Cum_Foot_index).Value, Cells(row_index + i, Cum_Foot_index).Value)
            End Select
            
            ReDim Bit(8, Range("Bit" & BitNum & "_Cutter_Num").Count)
            
            For iPlot = 1 To 8
                For iCutter = 1 To Range("Bit" & BitNum & "_Cutter_Num").Count
                    Bit(iPlot, iCutter) = LinInterpolate(Cells(row_index + i - 1, Cum_Foot_index).Value, Cells(row_index + i, Cum_Foot_index).Value, input_data, Range("Bit" & BitNum & "_Case" & i - 1 & strRange(iPlot)).Rows(iCutter), Range("Bit" & BitNum & "_Case" & i & strRange(iPlot)).Rows(iCutter))
                Next iCutter
            Next iPlot
            
            Sheets("BIT" & BitNum).Activate
            info_last_col = Range(ActiveSheet.UsedRange.Address).Columns.Count
            info_last_row = Range(ActiveSheet.UsedRange.Address).Rows.Count

            
'Find Previous CASE WEARs
            Cells.Find(What:="CASE WEAR", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Activate
            
            Range(ActiveCell, Cells(info_last_row, info_last_col)).ClearContents
            info_last_col = info_last_col - 9
            
'Setup Headers
NoPreviousWearPlot:
            Heading_Array1 = Array("area", "vol", "Circumferential Force (fx/fcut)", "Vertical Force (fy)", "Normal Force (fn)", "wf area", "Workrate (Circumferential Force)", "Workrate (Normal Force)")
            Heading_Array2 = Array("cut", "cut", "lb", "lb", "lb", "in**2", "lb-ft/s", "lb-ft/s")
            
            Cells(1, info_last_col + 2) = "CASE WEAR"
            For i = 0 To 7
                Cells(3, info_last_col + 2 + i) = Heading_Array1(i)
                Cells(4, info_last_col + 2 + i) = Heading_Array2(i)
                Range(Cells(5, info_last_col + 2 + i), Cells(4 + Range("Bit" & BitNum & "_Cutter_Num").Count, info_last_col + 2 + i)).Name = "Bit" & BitNum & "_CaseWear" & strRange(i + 1)
            Next i
            
            For iCutter = 1 To Range("Bit" & BitNum & "_Cutter_Num").Count
                For iPlot = 1 To 8
                    Cells(4 + iCutter, info_last_col + 1 + iPlot) = Bit(iPlot, iCutter)
                Next iPlot
            Next iCutter
            Sheets("SUMMARY").Activate
        End If
    Next BitNum
    Sheets("WEAR").Activate
    
    
    strCase = Range("Cases").Rows(1)
    strCase = Left(strCase, InStr(1, strCase, "_")) & Right(strCase, Len(strCase) - InStr(InStr(1, strCase, "_") + 1, strCase, "_")) & " @ " & Range("Wear_Input")
    
    If Sheets("WEAR").Shapes("Wear_Hours_Selection").ControlFormat.Value = xlOn Then
        strCase = strCase & " Hr"
    ElseIf Sheets("WEAR").Shapes("Wear_Footage_Selection").ControlFormat.Value = xlOn Then
        strCase = strCase & " ft"
    ElseIf Sheets("WEAR").Shapes("Wear_ROP_Selection").ControlFormat.Value = xlOn Then
        strCase = strCase & " ft/hr"
    End If
'    If Sheets("WEAR").OptionButtons("Wear_Hours_Selection").Value = xlOn Then
'        strCase = strCase & " Hr"
'    ElseIf Sheets("WEAR").OptionButtons("Wear_Footage_Selection").Value = xlOn Then
'        strCase = strCase & " ft"
'    ElseIf Sheets("WEAR").OptionButtons("Wear_ROP_Selection").Value = xlOn Then
'        strCase = strCase & " ft/hr"
'    End If
    strChrt = Array("NFwr_Chart", "WNFwr_Chart", "CFwr_Chart", "WCFwr_Chart", "VFwr_Chart", "WFAwr_Chart", "CAwr_Chart", "CVwr_Chart")
    
    str(0) = "Normal Forces " & Chr(13)
    str(1) = "Workrate of Normal Forces " & Chr(13)
    str(2) = "Circumferential Forces " & Chr(13)
    str(3) = "Workrate of Circumferential Forces " & Chr(13)
    str(4) = "Vertical Forces " & Chr(13)
    str(5) = "Wear Flat Area " & Chr(13)
    str(6) = "Cut Area " & Chr(13)
    str(7) = "Cut Volume " & Chr(13)
    
    For i = 0 To 7
        ActiveSheet.ChartObjects(strChrt(i)).Chart.ChartTitle.Text = str(i) & strCase
    Next i
    
    subUpdate True
    
    Exit Sub
    
ErrHandler:

    If Err = 91 Then 'Search for "CASE WEAR" came up empty
        Resume NoPreviousWearPlot
    Else
        Stop
        Resume
    End If
    
End Sub

Function LinInterpolate(X1, X2, X3, Y1, Y2)
    LinInterpolate = Y1 + ((Y2 - Y1) * (X3 - X1) / (X2 - X1))
End Function

Public Sub subModeSelect()

    If InStr(1, Application.Caller, "Standard", vbTextCompare) <> 0 Then
        
        subSetButtonCaps "Standard"

    ElseIf InStr(1, Application.Caller, "Rapid", vbTextCompare) <> 0 Then
    
'        If MsgBox("Initiating Rapid-Fire mode..." & vbCrLf & vbCrLf & "Would you like to clear all existing bits?", vbYesNo + vbQuestion, "Start Rapid-Fire Mode") = vbYes Then ClearBit 0
        
        subSetButtonCaps "Rapid"
        
        subUpdateShowHide
    End If
    
End Sub

Sub subRefreshBit()
    Dim strDir As String

    subActive True

    strDir = Range("Bit" & Right(Application.Caller, 1) & "_Location")
    strDir = Trim(Right(strDir, Len(strDir) - InStr(1, strDir, "]", vbTextCompare)))
    
    If strDir = "" Then
        MsgBox "Please load bit using the ""Read Bit"" buttons.", vbOKOnly, ""
        GoTo EndProcess
    End If
    
    subUpdate False
    
    If IsRapidFire Then subMoveIterationsDown Right(Application.Caller, 1)
    
    ReadBit strDir, True
    
    subUpdate True

EndProcess:
    subActive False
    
End Sub

Sub subSetButtonCaps(strMode As String)
    Dim i As Integer
    
    If strMode = "Standard" Then
        For i = 1 To Range("MaxBits")
            ActiveSheet.Shapes("Read_Des_Bit" & i).TextFrame.Characters.Text = "Read Bit " & i
            ActiveSheet.Shapes("Clear_Bit" & i).TextFrame.Characters.Text = "Clear Bit " & i
            
            ActiveSheet.Shapes("cmdRefreshBit" & i).TextFrame.Characters.Text = "Refresh Bit " & i
            ActiveSheet.Shapes("grpBLIT" & i).Visible = False
        Next i
        
        ActiveSheet.Shapes("lblRF").Visible = False
        ActiveSheet.Shapes("lstRF").Visible = False
        ActiveSheet.Shapes("RF_Chart").Visible = False
        ActiveSheet.Shapes("PlotItrComp").Visible = False
        ActiveSheet.ListBoxes("RFPlotBox1").Visible = False
        ActiveSheet.Shapes("cmdReset").Visible = False
        Range("RFCaseLabel").Font.ThemeColor = xlThemeColorDark1
        
        Range("Status").Columns.Hidden = False
'        Sheets("HOME").frmRFAxis.Visible = False
        Range("RFCols").Columns.Hidden = True
        
    ElseIf strMode = "Rapid" Then
        For i = 1 To 5
            ActiveSheet.Shapes("grpBLIT" & i).Visible = True
            
            SetBitButtons i
        Next i
        
        ActiveSheet.Shapes("lblRF").Visible = True
        ActiveSheet.Shapes("lstRF").Visible = True
        ActiveSheet.Shapes("RF_Chart").Visible = True
        ActiveSheet.ListBoxes("RFPlotBox1").Visible = True
        ActiveSheet.Shapes("cmdReset").Visible = True
        Range("RFCaseLabel").Font.ColorIndex = xlAutomatic
        
        Range("RFCols").Columns.Hidden = False
        Range("Status").Columns.Hidden = True
'        Sheets("HOME").frmRFAxis.Visible = True


    End If
    
End Sub

Sub subMoveIterationsDown(iBit As Integer)
    Dim i As Integer
    Dim iName As Integer
    Dim strArray As Variant
    Dim strDir As String, strName As String, strVis As String
    
    On Error GoTo ErrHandler

    If iBit = 1 Or iBit = Range("MaxBits") Then Exit Sub 'MaxBit rolls off sheet, first bit hardcoded as baseline

    i = Range("MaxBits") - 1
    
    Do While i >= iBit
        If Range("Bit" & i & "_Location") = "" Then GoTo EmptyIteration
        
        'Check to see if a summary file was loaded for this bit
        Sheets("BIT" & i & "_SUMMARY").Activate
        If ActiveSheet.UsedRange.Address = "$A$1" Then
            Sheets("HOME").Activate
            GoTo EmptyIteration
        End If
        
        strDir = Range("Bit" & i & "_Location")
        strName = Range("Bit" & i & "_Name")
        
        iName = 0
        Clear_Summary i + 1
        Clear_Cutter_Info i + 1
        
        Range("Bit" & i + 1 & "_Location") = strDir
        Range("Bit" & i + 1 & "_Name") = strName
        
        strArray = Array("", "_CUTTER_FILE", "_SUMMARY")
        
    'Check for i+1 visibility. Show for loading, then switch back to previous visibility
        If Sheets("HOME").Shapes("Show_Bit" & i + 1).ControlFormat.Value = 1 Then
            strVis = "Show_Bit"
        ElseIf Sheets("HOME").Shapes("Hide_Bit" & i + 1).ControlFormat.Value = 1 Then
            strVis = "Hide_Bit"
        ElseIf Sheets("HOME").Shapes("Hide_Backups_Bit" & i + 1).ControlFormat.Value = 1 Then
            strVis = "Hide_Backups_Bit"
        ElseIf Sheets("HOME").Shapes("Hide_Off_Backups_Bit" & i + 1).ControlFormat.Value = 1 Then
            strVis = "Hide_Off_Backups_Bit"
        End If
        
    'If not visibile, make bit visible to process info
        If strVis <> "Show_Bit" Then A = HideCutters(i + 1, False, False, False)
        
                
        Do While iName <= UBound(strArray)
        'Cut info from cutter i sheet and move it to next i #
            Sheets("BIT" & i & strArray(iName)).Activate
            ActiveSheet.UsedRange.Copy
            
            Sheets("BIT" & i + 1 & strArray(iName)).Activate
            Cells(1, 1).Select
            ActiveSheet.Paste
            
            iName = iName + 1
        Loop
        
        Sheets("BIT" & i + 1 & "_CUTTER_FILE").Activate

        If Right(strDir, 3) = "des" Then
            A = ProcessDes(i + 1, strDir)
        Else
            A = ProcessCfb(i + 1, strDir)
        End If

'must be on HOME sheet
        If Dir(Mid(strDir, InStr(1, strDir, "]", vbTextCompare) + 2, InStrRev(strDir, "\") - InStr(1, strDir, "]") - 1) & "summary_regular.*", vbNormal) <> "" Then
            A = ProcessSummary(i + 1, Mid(strDir, InStr(1, strDir, "]", vbTextCompare) + 2, InStrRev(strDir, "\") - InStr(1, strDir, "]") - 1) & "summary_regular.*", True)
        End If
        
    'revert visibility back to before this load
        Select Case strVis
            Case "Show_Bit": A = HideCutters(i + 1, False, False, False) ' Set to Show
            Case "Hide_Bit": A = HideCutters(i + 1, False, True, False) ' Set to Hidden
            Case "Hide_Backups_Bit": A = HideCutters(i + 1, True, True, False) ' set to hide backups
            Case "Hide_Off_Backups_Bit": A = HideCutters(i + 1, True, True, True) ' set to hide off-profile backups
        End Select

EmptyIteration:

        i = i - 1
    Loop
    
    Sheets("HOME").Select
    
    Exit Sub
    
ErrHandler:

'    Stop
'    Resume
'
'    If Err = 1004 Then
'        Resume Next
'    Else
        Sheets("HOME").Select
'    End If
End Sub

Sub subRFPlotData()
    Dim strPlot As String
    Dim i As Integer
    Dim blnVisible As Boolean
    
    
    
    strPlot = Right(Application.Caller, 2)
    
    If ActiveSheet.Shapes("chk" & strPlot).ControlFormat.Value = 1 Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
    i = 1
    
    ActiveSheet.Shapes("RF_Chart").Select
    
    Do While i <= ActiveChart.SeriesCollection.Count
        If InStr(1, ActiveChart.SeriesCollection(i).Name, strPlot) Then
            If blnVisible Then
                ActiveChart.SeriesCollection(i).bOrder.LineStyle = xlNone
                ActiveChart.SeriesCollection(i).MarkerStyle = xlNone
            Else
                If strPlot = "WR" Then
                    ActiveChart.SeriesCollection(i).MarkerStyle = xlSquare
                Else
                    ActiveChart.SeriesCollection(i).MarkerStyle = xlDiamond
                End If
                
                ActiveChart.SeriesCollection(i).bOrder.LineStyle = xlNone

            End If
        End If
    
        i = i + 1
    Loop

End Sub

Public Sub subChartSeriesSize(iBit As Integer, iKeep As Integer)
    Dim iSeries As Integer
    Dim iBits As Integer
    Dim i As Integer
    Dim iItr As Integer
    
    On Error GoTo ErrHandler
    
    With Sheets("HOME").ChartObjects("RF_Chart").Chart
        For i = 1 To Range("MaxBits")
            If Sheets("HOME").Shapes("Hide_Bit" & i).ControlFormat.Value <> xlOn Then iBits = iBits + 1
        Next i
        
        
        iSeries = .SeriesCollection.Count

        For i = 1 To iSeries
            strLine = .SeriesCollection(i).Formula
            iBit = CInt(Mid(strLine, InStr(1, strLine, "!Bit", vbTextCompare) + 4, 1))
            
            With .SeriesCollection(i)
            ' Test if series is for normal force or workrate
                If SeriesIsBaseline(iBit) Then
                ElseIf InStr(1, strLine, "_NF_", vbTextCompare) And SeriesIsBaseline(iBit) = False Then
                 ' Is Normal Force
                    .MarkerSize = (iKeep - i + 1) * 3
                ElseIf InStr(1, strLine, "_WR_", vbTextCompare) And SeriesIsBaseline(iBit) = False Then
                 ' Is Workrate
                    .MarkerSize = (iKeep - (i - iBits) + 1) * 3 - 1
                End If
            End With
        Next i
'
'
'
'        For i = 1 To iBits
'            If .SeriesCollection(i).Name <> Range("Bit1_Name") & "_NF" And SeriesIsBaseline(i) = False Then .SeriesCollection(i).MarkerSize = (iKeep - i + 1) * 3
'        Next i
'
'        For i = iBits + 1 To iSeries
'            If .SeriesCollection(i).Name <> Range("Bit1_Name") & "_WR" And SeriesIsBaseline(i) = False Then .SeriesCollection(i).MarkerSize = (iKeep - (i - iBits) + 1) * 3 - 1
'        Next i
    End With
    
    Exit Sub
    
ErrHandler:

Debug.Print Err, Err.Description

If Err = -2147467259 Then  'Method 'MarkerSize' of object 'Series' failed
    Resume Next
Else
    Stop
    Resume
End If

Stop
Resume

End Sub

Public Sub subUpdateShowHide()
    Dim i As Integer
    Dim iKeep As Integer
    Dim iShow(5) As Integer
    
    i = 1
    iKeep = Range("ItrToKeep") + 1
    
'Get Current Show Settings
    Do While i <= 5
        'Show_Bit1             = 1
        'Hide_Bit1             = 2
        'Hide_Backups_Bit1     = 3
        'Hide_Off_Backups_Bit1 = 4
        If ActiveSheet.Shapes("Show_Bit" & i).ControlFormat.Value = xlOn Then
            iShow(i) = 1
        ElseIf ActiveSheet.Shapes("Hide_Bit" & i).ControlFormat.Value = xlOn Then
            iShow(i) = 2
        ElseIf ActiveSheet.Shapes("Hide_Backups_Bit" & i).ControlFormat.Value = xlOn Then
            iShow(i) = 3
        ElseIf ActiveSheet.Shapes("Hide_Off_Backups_Bit" & i).ControlFormat.Value = xlOn Then
            iShow(i) = 4
        End If
'        If ActiveSheet.OptionButtons("Show_Bit" & i).Value = xlOn Then
'            iShow(i) = 1
'        ElseIf ActiveSheet.OptionButtons("Hide_Bit" & i).Value = xlOn Then
'            iShow(i) = 2
'        ElseIf ActiveSheet.OptionButtons("Hide_Backups_Bit" & i).Value = xlOn Then
'            iShow(i) = 3
'        ElseIf ActiveSheet.OptionButtons("Hide_Off_Backups_Bit" & i).Value = xlOn Then
'            iShow(i) = 4
'        End If
        
        i = i + 1
    Loop
    
    i = 1
    
    Do While i <= 5
        If i <= iKeep And iShow(i) = 2 Then
            ShowHideBackups i, "Show"
            
            ActiveSheet.Shapes("Show_Bit" & i).ControlFormat.Value = xlOn
            ActiveSheet.Shapes("Hide_Bit" & i).ControlFormat.Value = xlOff
        ElseIf i > iKeep Then
            ShowHideBackups i, "Hide_Bit"
            
            ActiveSheet.Shapes("Show_Bit" & i).ControlFormat.Value = xlOff
            ActiveSheet.Shapes("Hide_Bit" & i).ControlFormat.Value = xlOn
        End If
        
        i = i + 1
    Loop
    
    subChartSeriesSize i, iKeep

End Sub

Public Sub subForceDefaults()
    Dim i As Integer
    Dim iMax As Integer
    Dim blnHOME As Boolean
    
    If Range("Cases").Count < Range("MaxBits") Then
        iMax = Range("Cases").Count
    Else
        iMax = Range("MaxBits")
    End If
    
    If ActiveSheet.Name = "HOME" Then blnHOME = True
    
    For i = 1 To iMax
        If blnHOME Then
            Sheets("FORCES").ListBoxes("PlotBox" & i).Value = i
            Sheets("CUTTER_FORCES").ListBoxes("PlotBox" & i).Value = i
            subPlotBoxChange i, , "FORCES"
            subPlotBoxChange i, , "CUTTER_FORCES"
        Else
            ActiveSheet.ListBoxes("PlotBox" & i).Value = i
            subPlotBoxChange i
        End If
    Next i
    
    If iMax < Range("MaxBits") Then
        For i = iMax + 1 To Range("MaxBits")
            subUpdateCasePlots "FORCES", i, 0, True
            subUpdateCasePlots "CUTTER_FORCES", i, 0, True
        Next i
    End If
    
    
End Sub

Public Sub subRatioDefaults()
    Dim i As Integer
    Dim iMax As Integer
    Dim blnHOME As Boolean
    
    If Range("Cases").Count < Range("MaxBits") Then
        iMax = Range("Cases").Count
    Else
        iMax = Range("MaxBits")
    End If
    
    If ActiveSheet.Name = "HOME" Then blnHOME = True
    
    For i = 1 To iMax
        If blnHOME Then
            Sheets("FORCE_RATIOS").ListBoxes("PlotBox" & i).Value = i
            subRatioBoxChange i, , "FORCE_RATIOS"
        Else
            ActiveSheet.ListBoxes("PlotBox" & i).Value = i
            subRatioBoxChange i, 1, "FORCE_RATIOS"
        End If
    Next i
    
End Sub

Public Sub subPlotBoxChange(Optional iBox As Integer, Optional iCase As Integer, Optional strSheetName As String, Optional blnSmooth As Boolean)
    On Error GoTo ErrHandler
    
    If strSheetName = "" Then strSheetName = ActiveSheet.Name
    
'    If Application.Caller = "PlotItrComp" Then
'        iBox = 0
'        iCase = Sheets(strSheetName).ListBoxes(Application.Caller).Value
    If iBox = 0 Then iBox = 1
'    End If
    
    If strSheetName = "FORCES" Then
        If IsError(Range("ForceCase" & iBox)) Then Exit Sub
    ElseIf strSheetName = "CUTTER_FORCES" Then
        If IsError(Range("ProECase" & iBox)) Then Exit Sub
    End If
    

    If iCase = 0 Then
        iCase = Sheets(strSheetName).ListBoxes("PlotBox" & iBox).Value
'    ElseIf Application.Caller <> "PlotItrComp" Then
    Else
        Sheets(strSheetName).ListBoxes("PlotBox" & iBox) = iCase
    End If
    
'    subRewriteCasePlots strSheetName, iBox, iCase, False, blnSmooth
    subUpdateCasePlots strSheetName, iBox, iCase, False, blnSmooth
    
    Exit Sub
    
ErrHandler:

    If Err = 424 Then
        Exit Sub
    ElseIf Err = 13 Then
        Exit Sub
    Else
        Resume Next
    End If
    
End Sub

Public Sub subRatioBoxChange(Optional iBox As Integer, Optional iCase As Integer, Optional strSheetName As String)
    Dim iBit As Integer
    On Error GoTo ErrHandler
    
'    If strSheetName = "" Then strSheetName = ActiveSheet.Name
'
'    If Application.Caller = "PlotItrComp" Then
'        iBox = 0
'        iCase = Sheets(strSheetName).ListBoxes(Application.Caller).Value
'    ElseIf iBox = 0 Then
'        iBox = Right(Application.Caller, 1)
'    End If
    If iBox = 0 Then iBox = 1
    
    If strSheetName = "FORCE_RATIOS" Then
        If IsError(Range("ForceCase" & iBox)) Then Exit Sub
    End If
    
    If iCase = 0 Then
        iCase = Sheets(strSheetName).ListBoxes("PlotBox" & iBox).Value
    ElseIf Application.Caller <> "PlotItrComp" Then
        Sheets(strSheetName).ListBoxes("PlotBox" & iBox) = iCase
    End If
    
    For iBit = 1 To Range("MaxBits")
        Range("Bit" & iBit & "_OuterForces") = "=SUM(Bit" & iBit & "_Case" & iCase & "_NormalOuter)"
        Range("Bit" & iBit & "_InnerForces") = "=SUM(Bit" & iBit & "_Case" & iCase & "_NormalInner)"
    Next iBit
    
    Exit Sub
    
ErrHandler:

    If Err = 424 Then
        Exit Sub
    Else
        Resume Next
    End If
    
End Sub

Public Sub subRewriteCasePlots(strChartSheet As String, iBox As Integer, iCase As Integer, Optional blnClear As Boolean, Optional blnSmooth As Boolean)
    Dim strArray As Variant, strPlotArrayDesc As Variant, strPlotArrayTitle As Variant
    Dim i As Integer, iBitNum As Integer
    Dim iSeriesCount As Integer, iSeries As Integer
    Dim strFormula As String
    Dim strForces As String
    Dim strChart As String
    
    On Error GoTo ErrHandler
    
    strChart = "_Chart"
    
    If strChartSheet = "FORCES" Then
        strPlotArray = Array("CF", "NF", "VF", "WCF", "WNF", "WFA", "CA", "CV", "SF", "TQB")
        strPlotArrayDesc = Array("Circumferential", "Normal", "Vertical", "Circumferential_Workrate", "Normal_Workrate", "WearFlatArea", "AreaCut", "VolCut", "Side", "TQ_BIT")
        strPlotArrayTitle = Array("Circumferential Forces", "Normal Forces", "Vertical Forces", "Workrate of Circumferential Forces", "Workrate of Normal Forces", "Wear Flat Area", "Cut Area", "Cut Volume", "Side Forces", "Cutter Contribution to Bit Torque")
        strForces = "ForceCase"
        
    ElseIf strChartSheet = "CUTTER_FORCES" Then
        strPlotArray = Array("FX", "FY", "FZ", "CT", "RF", "WRF", "DL", "WDL")
        strPlotArrayDesc = Array("ProE_fx", "ProE_fy", "ProE_fz", "Cutter_Torque", "Resultant", "Resultant_Workrate", "Delam", "Delam_Workrate")
        strPlotArrayTitle = Array("ProE - fx", "ProE - fy", "ProE - fz", "Cutter Torque", "Resultant Forces", "Workrate of Resultant Forces", "Delam Forces", "Workrate of Delam Forces")
        strForces = "ProECase"
        
    ElseIf iBox = 0 Then
        strPlotArray = Array("RF")
        strPlotArrayDesc = Array("ProE_fx", "ProE_fy", "ProE_fz")
        strPlotArrayTitle = Array("ProE - fx", "ProE - fy", "ProE - fz")
        strForces = "ProECase"
        strChart = "Chart"

    Else
        If Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then MsgBox "Error. Not on a force sheet."
        Exit Sub
    End If
        
    Do While i <= UBound(strPlotArray)
        With Sheets(strChartSheet).ChartObjects(strPlotArray(i) & iBox & strChart).Chart
        
            If blnClear Then
                .ChartTitle.Text = strPlotArrayTitle(i)
                GoTo NextI
            Else
                .ChartTitle.Text = strPlotArrayTitle(i) & Chr(13) & Range(strForces & iBox)
            End If
            
            iSeriesCount = .SeriesCollection.Count

            For iSeries = 1 To iSeriesCount
                strFormula = .SeriesCollection(iSeries).Formula
                
                If InStr(1, strFormula, "!Bit1") > 0 Then
                    iBitNum = 1
                    .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_XAxis,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Case" & iCase & "_" & strPlotArrayDesc(i) & "," & iBitNum & ")"
                ElseIf InStr(1, strFormula, "!Bit2") > 0 Then
                    iBitNum = 2
                    .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_XAxis,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Case" & iCase & "_" & strPlotArrayDesc(i) & "," & iBitNum & ")"
                ElseIf InStr(1, strFormula, "!Bit3") > 0 Then
                    iBitNum = 3
                    .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_XAxis,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Case" & iCase & "_" & strPlotArrayDesc(i) & "," & iBitNum & ")"
                ElseIf InStr(1, strFormula, "!Bit4") > 0 Then
                    iBitNum = 4
                    .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_XAxis,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Case" & iCase & "_" & strPlotArrayDesc(i) & "," & iBitNum & ")"
                ElseIf InStr(1, strFormula, "!Bit5") > 0 Then
                    iBitNum = 5
                    .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_XAxis,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Case" & iCase & "_" & strPlotArrayDesc(i) & "," & iBitNum & ")"
                ElseIf iSeries = 7 Then
                    .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Zone_Outer,'" & ActiveSheet.Parent.Name & "'!Zone_Xouter,'" & ActiveSheet.Parent.Name & "'!Zone_Youter,7)"
                ElseIf InStr(1, strFormula, "Zone_Inner") > 0 Or InStr(1, strFormula, "$P$1") > 0 Then
                    .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Zone_Inner,'" & ActiveSheet.Parent.Name & "'!Zone_Xinner,'" & ActiveSheet.Parent.Name & "'!Zone_Yinner,6)"
                End If
            
            Next iSeries



'
'
'
'
'
'
'
'                For iBitNum = 1 To Range("MaxBits")
'                    strFormula = .SeriesCollection(iSeries).Formula
'                    If InStr(1, strFormula, "Bit" & iBitNum, vbTextCompare) > 0 Then
'                        Exit For
'                    ElseIf InStr(1, strFormula, "Zone_Inner", vbTextCompare) > 0 Then
'                        Exit For
'                    ElseIf InStr(1, strFormula, "Zone_Outer", vbTextCompare) > 0 Then
'                        Exit For
'                    End If
'                Next iBitNum
'
'                If iSeries <= ActiveBits Then 'Range("MaxBits") Then
'                    .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Radius,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Case" & iCase & "_" & strPlotArrayDesc(i) & "," & iSeries & ")"
'                    .SeriesCollection(iSeries).Smooth = blnSmooth
'                Else
''                    'to add zone of interest
'                    If iSeries = ActiveBits + 1 Then 'InStr(1, strFormula, "Inner", vbTextCompare) > 0 Then
'''                        'Outer
'                        .SeriesCollection(iSeries).Formula = "=SERIES(SAM_v10.00.xlsm!Zone_Inner,SAM_v10.00.xlsm!Zone_Xinner,SAM_v10.00.xlsm!Zone_Yinner," & iSeries & ")"
'                    ElseIf iSeries = ActiveBits + 2 Then
'''                        'Inner
'                        .SeriesCollection(iSeries).Formula = "=SERIES(SAM_v10.00.xlsm!Zone_Outer,SAM_v10.00.xlsm!Zone_Xouter,SAM_v10.00.xlsm!Zone_Youter," & iSeries & ")"
'                    End If
'                    .SeriesCollection(iSeries).Smooth = blnSmooth
'                End If
'            Next iSeries
        End With
NextI:
        i = i + 1
    Loop

    Exit Sub
    
ErrHandler:


If Err = 1004 Then 'desc
    Debug.Print Err, Err.Description
    Stop
    Resume
    MsgBox "A formula in this worksheet contains one or more invalid references." & vbCrLf & vbCrLf & "Please ensure all active bits have static results in the cases you're exporting.", vbCritical + vbOKOnly, "Formula Error"
    Debug_flag = True
    Exit Sub
Else
    Debug.Print Err, Err.Description
    Stop
    Resume
End If

End Sub


Public Sub subUpdateCasePlots(strChartSheet As String, iBox As Integer, iCase As Integer, Optional blnClear As Boolean, Optional blnSmooth As Boolean)
    Dim strArray As Variant, strPlotArray As Variant, strPlotArrayDesc As Variant, strPlotArrayTitle As Variant
    Dim i As Integer, iBitNum As Integer
    Dim iSeriesCount As Integer, iSeries As Integer
    Dim strFormula As String, strFormula2 As String
    Dim strForces As String
    Dim strChart As String
    Dim strZoIShowHide As String
    
    On Error GoTo ErrHandler
    
    'Option_Button_Hide_Zone_Click() = True
'    Range("Limits").Columns.EntireColumn.Hidden = True
'    Range("Zone").Rows.EntireRow.Hidden = True
    
    strZoIShowHide = Range("ZoIShowHide")
    
    ZoI "Hide"
    gblnZoIShtOverride = False
    
    strChart = "_Chart"
    
    If strChartSheet = "FORCES" Then
        strPlotArray = Array("CF", "NF", "VF", "WCF", "WNF", "WFA", "CA", "CV", "SF", "TQB")
        strPlotArrayDesc = Array("Circumferential", "Normal", "Vertical", "Circumferential_Workrate", "Normal_Workrate", "WearFlatArea", "AreaCut", "VolCut", "Side", "TQ_BIT")
        strPlotArrayTitle = Array("Circumferential Forces", "Normal Forces", "Vertical Forces", "Workrate of Circumferential Forces", "Workrate of Normal Forces", "Wear Flat Area", "Cut Area", "Cut Volume", "Side Forces", "Cutter Contribution to Bit Torque")
        strForces = "ForceCase"
        
    ElseIf strChartSheet = "CUTTER_FORCES" Then
        strPlotArray = Array("FX", "FY", "FZ", "CT", "RF", "WRF", "DL", "WDL")
        strPlotArrayDesc = Array("ProE_fx", "ProE_fy", "ProE_fz", "Cutter_Torque", "Resultant", "Resultant_Workrate", "Delam", "Delam_Workrate")
        strPlotArrayTitle = Array("ProE - fx", "ProE - fy", "ProE - fz", "Cutter Torque", "Resultant Forces", "Workrate of Resultant Forces", "Delam Forces", "Workrate of Delam Forces")
        strForces = "ProECase"
        
    ElseIf iBox = 0 Then
        strPlotArray = Array("RF")
        strPlotArrayDesc = Array("ProE_fx", "ProE_fy", "ProE_fz")
        strPlotArrayTitle = Array("ProE - fx", "ProE - fy", "ProE - fz")
        strForces = "ProECase"
        strChart = "Chart"

    Else
        If Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then MsgBox "Error. Not on a force sheet."
        Exit Sub
    End If
    
    Do While i <= UBound(strPlotArray)
        With Sheets(strChartSheet).ChartObjects(strPlotArray(i) & iBox & strChart).Chart
'            If strPlotArray(i) = "SF" Then Stop
            
            If blnClear Then
                .ChartTitle.Text = strPlotArrayTitle(i)
                GoTo NextI
            Else
                .ChartTitle.Text = strPlotArrayTitle(i) & Chr(13) & Range(strForces & iBox)
            End If
            
            iSeriesCount = .SeriesCollection.Count
            
            For iSeries = 1 To iSeriesCount
                For iBitNum = 1 To Range("MaxBits")
                    strFormula = .SeriesCollection(iSeries).Formula
                    If InStr(1, strFormula, "Bit" & iBitNum, vbTextCompare) > 0 Then
                        Exit For
                    ElseIf InStr(1, strFormula, "Zone_Inner", vbTextCompare) > 0 Then
                        Exit For
                    ElseIf InStr(1, strFormula, "Zone_Outer", vbTextCompare) > 0 Then
                        Exit For
                    End If
                Next iBitNum

                If iSeries <= ActiveBits Then 'Range("MaxBits") Then
                    .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_XAxis,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Case" & iCase & "_" & strPlotArrayDesc(i) & "," & iSeries & ")"
                    .SeriesCollection(iSeries).Smooth = blnSmooth
                End If
            Next iSeries
        End With
NextI:
        i = i + 1
    Loop

        
'    Do While i <= UBound(strPlotArray)
'        With Sheets(strChartSheet).ChartObjects(strPlotArray(i) & iBox & strChart).Chart
'
'            If blnClear Then
'                .ChartTitle.Text = strPlotArrayTitle(i)
'                GoTo NextI
'            Else
'                .ChartTitle.Text = strPlotArrayTitle(i) & Chr(13) & Range(strForces & iBox)
'            End If
'
'            iSeriesCount = .SeriesCollection.Count
'
'            For iSeries = 1 To iSeriesCount
'
'            If strPlotArray(i) = "SF" Or strPlotArray(i) = "TQB" Then GoTo RewriteFormula
'
'                For iBitNum = 1 To Range("MaxBits")
'                    strFormula = .SeriesCollection(iSeries).Formula
'                    If InStr(1, strFormula, "Bit" & iBitNum, vbTextCompare) > 0 Then
'                        Exit For
'                    ElseIf InStr(1, strFormula, "Zone_Inner", vbTextCompare) > 0 Then
'                        Exit For
'                    ElseIf InStr(1, strFormula, "Zone_Outer", vbTextCompare) > 0 Then
'                        Exit For
'                    End If
'                Next iBitNum
'RewriteFormula:
'                If iSeries <= ActiveBits Then 'Range("MaxBits") Then
'                    strFormula2 = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_XAxis,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Case" & iCase & "_" & strPlotArrayDesc(i) & "," & iSeries & ")"
'                    .SeriesCollection(iSeries).Formula = strFormula2
'                    .SeriesCollection(iSeries).Smooth = blnSmooth
'                ElseIf iSeries = ActiveBits + 1 Then
'                    strFormula2 = "=SERIES('" & ActiveSheet.Parent.Name & "'!Zone_Inner,'" & ActiveSheet.Parent.Name & "'!Zone_Xinner,'" & ActiveSheet.Parent.Name & "'!Zone_Yinner," & iSeries & ")"
'                    .SeriesCollection(iSeries).Formula = strFormula2
''                    .SeriesCollection(iSeries).Smooth = blnSmooth
'                ElseIf iSeries = ActiveBits + 2 Then
'                    strFormula2 = "=SERIES('" & ActiveSheet.Parent.Name & "'!Zone_Outer,'" & ActiveSheet.Parent.Name & "'!Zone_Xouter,'" & ActiveSheet.Parent.Name & "'!Zone_Youter," & iSeries & ")"
'                    .SeriesCollection(iSeries).Formula = strFormula2
''                    .SeriesCollection(iSeries).Smooth = blnSmooth
'                End If
'            Next iSeries
'        End With
'NextI:
'        i = i + 1
'    Loop
    
    If strZoIShowHide = "Show" Then
        ZoI "Show"
        gblnZoIShtOverride = False
    End If
    
    Range("A1").Select


    
    
    
    Exit Sub
    
ErrHandler:
Debug.Print Now(), Err, Err.Description

If Err = 1004 Then 'Parameter not valid
    Stop
    Resume
    MsgBox "A formula in this worksheet contains one or more invalid references." & vbCrLf & vbCrLf & "Please ensure all active bits have static results in the cases you're exporting.", vbCritical + vbOKOnly, "Formula Error"
    Debug_flag = True
    Exit Sub
Else
    Debug.Print Err, Err.Description
    Stop
    Resume
End If

End Sub


Public Sub subRewritePlotFormulas()
    Dim strChartSheet As Variant, strPlotArray As Variant, strPlotArrayDesc As Variant, strPlotArrayTitle As Variant
    Dim i As Integer, iBitNum As Integer, iSheet As Integer
    Dim iSeriesCount As Integer, iSeries As Integer
    Dim strFormula As String
    Dim strForces As String
    Dim strChart As String
    Dim strZoIShowHide As String
    
    On Error GoTo ErrHandler
    
    For i = 1 To Range("MaxBits")
        ShowHideBackups i, "Show"
    Next i
    
    ZoI "Show"
    
    strZoIShowHide = Range("ZoIShowHide")
    
    gblnZoIShtOverride = False
    
    strChart = "_Chart"
        
    strChartSheet = Array("FORCES", "CUTTER_FORCES")

    For iSheet = 0 To UBound(strChartSheet)
        If strChartSheet(iSheet) = "FORCES" Then
            strPlotArray = Array("CF", "NF", "VF", "WCF", "WNF", "WFA", "CA", "CV", "SF", "TQB")
            strPlotArrayDesc = Array("Circumferential", "Normal", "Vertical", "Circumferential_Workrate", "Normal_Workrate", "WearFlatArea", "AreaCut", "VolCut", "Side", "TQ_BIT")
            strPlotArrayTitle = Array("Circumferential Forces", "Normal Forces", "Vertical Forces", "Workrate of Circumferential Forces", "Workrate of Normal Forces", "Wear Flat Area", "Cut Area", "Cut Volume", "Side Forces", "Cutter Contribution to Bit Torque")
            strForces = "ForceCase"
        ElseIf strChartSheet(iSheet) = "CUTTER_FORCES" Then
            strPlotArray = Array("FX", "FY", "FZ", "CT", "RF", "WRF", "DL", "WDL")
            strPlotArrayDesc = Array("ProE_fx", "ProE_fy", "ProE_fz", "Cutter_Torque", "Resultant", "Resultant_Workrate", "Delam", "Delam_Workrate")
            strPlotArrayTitle = Array("ProE - fx", "ProE - fy", "ProE - fz", "Cutter Torque", "Resultant Forces", "Workrate of Resultant Forces", "Delam Forces", "Workrate of Delam Forces")
            strForces = "ProECase"
        Else
            If Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then MsgBox "Error. Not on a force sheet."
            Exit Sub
        End If
        
        Sheets(strChartSheet(iSheet)).Select
        
        For iPlot = 0 To UBound(strPlotArray)
            For iBox = 1 To 5
                Debug.Print strChartSheet(iSheet), iPlot, strPlotArray(iPlot), iBox
                
                With Sheets(strChartSheet(iSheet)).ChartObjects(strPlotArray(iPlot) & iBox & strChart).Chart
                
                    .ChartTitle.Text = strPlotArrayTitle(iPlot) & Chr(13) & Range(strForces & iBox)
                    
                    iSeriesCount = .SeriesCollection.Count
                    If iSeriesCount <> 7 Then Stop
                    
                    For iSeries = 1 To Range("MaxBits")
                        strFormula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iSeries & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iSeries & "_XAxis,'" & ActiveSheet.Parent.Name & "'!Bit" & iSeries & "_Case" & iBox & "_" & strPlotArrayDesc(iPlot) & "," & iSeries & ")"
                        .SeriesCollection(iSeries).Formula = strFormula
                        .SeriesCollection(iSeries).Smooth = False
                    Next iSeries
                    
                    iSeries = 6
                    strFormula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Zone_Inner,'" & ActiveSheet.Parent.Name & "'!Zone_Xinner,'" & ActiveSheet.Parent.Name & "'!Zone_Yinner," & iSeries & ")"
                    .SeriesCollection(iSeries).Formula = strFormula
                    .SeriesCollection(iSeries).Smooth = False
                    
                    iSeries = 7
                    strFormula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Zone_Outer,'" & ActiveSheet.Parent.Name & "'!Zone_Xouter,'" & ActiveSheet.Parent.Name & "'!Zone_Youter," & iSeries & ")"
                    .SeriesCollection(iSeries).Formula = strFormula
                    .SeriesCollection(iSeries).Smooth = False
                    
                End With
NextI:
            Next iBox
        Next iPlot
    Next iSheet


    If strZoIShowHide = "Show" Then
        ZoI "Show"
        gblnZoIShtOverride = False
    End If
    
    Range("A1").Select


    
    
    
    Exit Sub
    
ErrHandler:


If Err = 1004 Then 'desc
    Stop
    Resume
    MsgBox "A formula in this worksheet contains one or more invalid references." & vbCrLf & vbCrLf & "Please ensure all active bits have static results in the cases you're exporting.", vbCritical + vbOKOnly, "Formula Error"
    Debug_flag = True
    Exit Sub
Else
    Debug.Print Err, Err.Description
    Stop
    Resume
End If

End Sub
Public Sub subOpenExportForm()

    frmExportForces.Show
    
End Sub

Public Sub subSmoothCurves(Optional strSheet As String, Optional blnValue As Boolean)
    Dim i As Integer, k As Integer
    Dim iCaseF As Integer, iCaseC As Integer
    Dim iMax As Integer
    Dim blnSmooth As Boolean
    
    On Error GoTo ErrHandler

    If strSheet <> "" Then
        blnSmooth = blnValue
    Else
        If ActiveSheet.CheckBoxes("chkSmooth").Value = 1 Then
            blnSmooth = True
        Else
            blnSmooth = False
        End If
    End If
    
    If ActiveSheet.Name = "PROFILE" Or strSheet = "PROFILE" Then
        subUpdate False
        For i = 1 To ActiveSheet.ChartObjects.Count
            For k = 1 To Sheets("PROFILE").ChartObjects(i).Chart.SeriesCollection.Count
                Sheets("PROFILE").ChartObjects(i).Chart.SeriesCollection(k).Smooth = blnSmooth
            Next k
        Next i
        subUpdate True
'        Stop
    Else
        Sheets("FORCES").CheckBoxes("chkSmooth").Value = blnSmooth
        Sheets("CUTTER_FORCES").CheckBoxes("chkSmooth").Value = blnSmooth
        
        If Range("Cases").Count < Range("MaxBits") Then
            iMax = Range("Cases").Count
        Else
            iMax = Range("MaxBits")
        End If
        
        For i = 1 To iMax
            subPlotBoxChange i, Sheets("FORCES").ListBoxes("PlotBox" & i).Value, "FORCES", blnSmooth
            subPlotBoxChange i, Sheets("CUTTER_FORCES").ListBoxes("PlotBox" & i).Value, "CUTTER_FORCES", blnSmooth
        Next i
        
        If iMax < Range("MaxBits") Then
            For i = iMax + 1 To Range("MaxBits")
                subUpdateCasePlots "FORCES", i, 0, True
                subUpdateCasePlots "CUTTER_FORCES", i, 0, True
            Next i
        End If
    End If
    
    Exit Sub
    
ErrHandler:

If Err = 1004 Then
'Stop
Resume Next
End If

End Sub

Public Sub subRFPlotCase()
    Dim iSeriesCount As Integer, iSeries As Integer, iBox As Integer
    Dim strFormula As String
    Dim strForces As String
    Dim strChart As String, strChartSheet As String

    On Error GoTo ErrHandler
    
'    Stop
    
    If IsError(Range("RFCase1")) Then Exit Sub

    strChartSheet = ActiveSheet.Name
    strChart = "_Chart"
    strForces = "RFCase"
    iBox = 1
    
    With Sheets(strChartSheet).ChartObjects("RF" & strChart).Chart
    
        If blnClear Then
            .ChartTitle.Text = "Iteration Comparison"
'            GoTo NextI
        Else
            .ChartTitle.Text = "Iteration Comparison - " & Range(strForces & iBox)
        End If
        
        iSeriesCount = .SeriesCollection.Count
        
        For iSeries = 1 To iSeriesCount
            strFormula = .SeriesCollection(iSeries).Formula
'            strFormula = Replace(strFormula, "Radius", "XAxis")
'            For iCase = 1 To Range("Cases").Count
'                strFormula = .SeriesCollection(iSeries).Formula
'                If InStr(1, strFormula, "Case" & iCase, vbTextCompare) > 0 Then
'                    Stop
'                    Exit For
'                End If
'            Next iCase

'            strFormula = Replace(strFormula, "Case" & iCase, "Case" & Range("RFCaseID1"), , , vbTextCompare)
            
'            .SeriesCollection(iSeries).Formula = Replace(strFormula, "Case" & iCase, "Case" & Range("RFCaseID1"), , , vbTextCompare)
            .SeriesCollection(iSeries).Formula = Left(strFormula, InStr(1, strFormula, "_Case", vbTextCompare) + 4) & Range("RFCaseID1") & Right(strFormula, Len(strFormula) - InStr(InStr(1, strFormula, "_Case", vbTextCompare) + 4, strFormula, "_", vbTextCompare) + 1)
        Next iSeries
    End With
    
    If strChartSheet = "HOME" Then Range("Rest").Select
    
    Exit Sub
    
ErrHandler:



If Err = 1004 Then
    'this
    Resume Next
Else
    'that
    Stop
    Resume
End If

End Sub

Public Sub subResetRFPlot()

    subClearRFPlot
    

End Sub
Public Sub subClearRFPlot()

    strChartSheet = "HOME"
    
    With Sheets(strChartSheet).ChartObjects("RF_Chart").Chart
    
        .ChartTitle.Text = "Iteration Comparison"
        
        iSeriesCount = .SeriesCollection.Count
        
        For i = 1 To iSeriesCount
            strFormula = .SeriesCollection(i).Formula
            
'            Debug.Print strFormula & vbCrLf & Replace(strFormula, Mid(strFormula, InStr(1, strFormula, "Case", vbTextCompare), InStr(InStr(1, strFormula, "Case", vbTextCompare) + 1, strFormula, "_", vbTextCompare) - InStr(1, strFormula, "Case", vbTextCompare)), "Case1", , , vbTextCompare) & vbCrLf
            
            strFormula = Replace(strFormula, Mid(strFormula, InStr(1, strFormula, "Case", vbTextCompare), InStr(InStr(1, strFormula, "Case", vbTextCompare) + 1, strFormula, "_", vbTextCompare) - InStr(1, strFormula, "Case", vbTextCompare)), "Case1", , , vbTextCompare)
            .SeriesCollection(i).Formula = strFormula
        Next i
        
'        Sheets(strChartSheet).listboxes("
        
        
        
'            For iBitNum = 1 To Range("MaxBits")
'                strFormula = .SeriesCollection(iSeries).Formula
'                If InStr(1, strFormula, "Bit" & iBitNum, vbTextCompare) > 0 Then Exit For
'            Next iBitNum
'
'            .SeriesCollection(iSeries).Formula = "=SERIES('" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Name,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Radius,'" & ActiveSheet.Parent.Name & "'!Bit" & iBitNum & "_Case" & iCase & "_" & strPlotArrayDesc(i) & "," & iSeries & ")"

    End With

End Sub


Public Function IsRapidFire() As Boolean

    If Sheets("HOME").Shapes("optRapidFireMode").ControlFormat.Value = xlOn Then
        IsRapidFire = True
    Else
        IsRapidFire = False
    End If

End Function

Public Function IsZoI() As Boolean

    If Sheets("HOME").Shapes("Option_Button_Show_zone").ControlFormat.Value = xlOn Then
        IsZoI = True
    Else
        IsZoI = False
    End If

End Function

Private Function funcFilePicker(strTitle As String, strFilter As String, strExtension As String, Optional strStartLoc As String, Optional bMulti As Boolean) As String
    Dim strFileName As String
    Dim strFile As String
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    If IsNull(bMulti) Then bMulti = False
    
    If IsNull(strFile) Or strFile = "" Then
        With fd
            .Filters.Clear
            .Title = strTitle
            .AllowMultiSelect = bMulti
            .Filters.Add strFilter, strExtension, 1
            .InitialFileName = strStartLoc
            
            If .Show = -1 Then
                strFileName = .SelectedItems(1)
            Else
                Exit Function
            End If
            .Filters.Clear
        End With
        Set fd = Nothing
    Else
        strFileName = strFile
    End If
    
    funcFilePicker = strFileName

End Function

Private Function funcFolderPicker(strTitle As String, strFilter As String, Optional strExtension As String, Optional strStartLoc As String, Optional bMulti As Boolean) As String
    Dim strFileName As String
    Dim strFile As String
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    If IsNull(bMulti) Then bMulti = False
    
    If IsNull(strFile) Or strFile = "" Then
        With fd
            .Title = strTitle
            .AllowMultiSelect = bMulti
'            .Filters.Add strFilter, strExtension, 1
            .InitialFileName = strStartLoc
            
            If .Show = -1 Then
                strFileName = .SelectedItems(1)
            Else
                strFileName = ""
            End If
        End With
        Set fd = Nothing
    Else
        strFileName = strFile
    End If
    
    funcFolderPicker = strFileName
    
End Function


'Sub IT2BL()
Public Sub subBaselineIteration(Optional iBit As Integer, Optional strBLIT As String)
    Dim i As Integer, iSeries As Integer, iBitFirstItr As Integer
'    Dim iBit As Integer, strBLIT As String
    Dim strLine As String
    Dim strType As String
    
    If Left(Application.Caller, 3) <> "opt" Then
        Sheets("HOME").Shapes("opt" & strBLIT & "Bit" & iBit).ControlFormat.Value = xlOn
    Else
        iBit = Right(Application.Caller, 1)
        strBLIT = Mid(Application.Caller, 4, 2)
    End If

    SetBitButtons iBit
    iBitFirstItr = GetFirstBitIteration
    
    With Sheets("HOME").ChartObjects("RF_Chart").Chart
    
        iSeries = .SeriesCollection.Count
        
        For i = 1 To iSeries
            strLine = .SeriesCollection(i).Formula
            
            If InStr(1, strLine, "Bit" & iBit & "_", vbTextCompare) > 0 Then
                With .SeriesCollection(i)
                ' Test if series is for normal force or workrate
                    If InStr(1, strLine, "Bit" & iBit & "_NF_", vbTextCompare) Then
                        strType = "NF"
                    ElseIf InStr(1, strLine, "Bit" & iBit & "_WR_", vbTextCompare) Then
                        strType = "WR"
                    End If
                
                    If strBLIT = "BL" Then
                    'Found series to be a baseline
                        If strType = "NF" Then
                            .MarkerStyle = xlMarkerStyleDiamond
                            .MarkerSize = 6
                        ElseIf strType = "WR" Then
                            .MarkerStyle = xlMarkerStyleSquare
                            .MarkerSize = 3
                        End If
                        
                        .MarkerBackgroundColor = RGB(255, 255, 255)
                        .MarkerForegroundColor = RGB(GetBitColor(iBit, "R", strType), GetBitColor(iBit, "G", strType), GetBitColor(iBit, "B", strType))
                        .Format.Fill.Visible = msoFalse
                        
                        With .Format.Line
                            .Visible = msoTrue
                            .ForeColor.RGB = RGB(GetBitColor(iBit, "R", strType), GetBitColor(iBit, "G", strType), GetBitColor(iBit, "B", strType))
                            .Transparency = 0
                            .DashStyle = msoLineSysDot   'msoLineSolid
                            If iBit = 3 Then
                                .Weight = 1.5
                            Else
                                .Weight = 0.25
                            End If
                        End With
                    
                        .MarkerForegroundColor = RGB(GetBitColor(iBit, "R", strType), GetBitColor(iBit, "G", strType), GetBitColor(iBit, "B", strType))  'marker border
                        
                        'Stop
'                        MakePointsSolid i, .Points.Count, Sheets("HOME").ChartObjects("RF_Chart").Chart
                    
                    ElseIf strBLIT = "IT" Then
                    'Found series to be an iteration
                    
                        .MarkerBackgroundColor = RGB(GetBitColor(iBit, "R", strType), GetBitColor(iBit, "G", strType), GetBitColor(iBit, "B", strType)) 'marker fill
                        .MarkerForegroundColorIndex = xlColorIndexNone  'tranparent marker lines
    
                        
                        If iBitFirstItr > iBit Then
                            .Format.Line.Visible = msoFalse
                        Else
                            .Format.Line.Visible = msoTrue
                            .Format.Line.Weight = 1
                            .Format.Line.DashStyle = msoLineSolid
                            .Format.Line.ForeColor.RGB = RGB(GetBitColor(iBit, "R", strType), GetBitColor(iBit, "G", strType), GetBitColor(iBit, "B", strType))
                            
                        End If

                        .Format.Fill.Visible = msoTrue
                        .Format.Fill.ForeColor.RGB = RGB(GetBitColor(iBit, "R"), GetBitColor(iBit, "G"), GetBitColor(iBit, "B"))
                        .Format.Fill.Transparency = 0
                        .Format.Fill.Solid
    
                        subUpdateShowHide
                    End If
                End With
            End If
        Next i
    End With
    
End Sub

Public Function GetFirstBitIteration() As Integer
    Dim i As Integer, iLastBit As Integer
    
    iLastBit = Range("MaxBits")
    
    If Sheets("HOME").Shapes("optITBit1").ControlFormat.Value = xlOn Then
        GetFirstBitIteration = 1
    ElseIf Sheets("HOME").Shapes("optITBit2").ControlFormat.Value = xlOn Then
        GetFirstBitIteration = 2
    ElseIf Sheets("HOME").Shapes("optITBit3").ControlFormat.Value = xlOn Then
        GetFirstBitIteration = 3
    ElseIf Sheets("HOME").Shapes("optITBit4").ControlFormat.Value = xlOn Then
        GetFirstBitIteration = 4
    ElseIf Sheets("HOME").Shapes("optITBit5").ControlFormat.Value = xlOn Then
        GetFirstBitIteration = 5
    Else
        GetFirstBitIteration = 0  ' All Itrs set to Baseline
    End If

End Function

Public Function GetBitColor(iBit As Integer, strColor As String, Optional strForce As String) As Integer

    If strForce = "" Then strForce = "NF"
    
    Select Case strForce
        Case "NF":
            Select Case iBit
                Case 1:
                    Select Case strColor
                        Case "R": GetBitColor = 192
                        Case "G": GetBitColor = 0
                        Case "B": GetBitColor = 0
                    End Select
                Case 2:
                    Select Case strColor
                        Case "R": GetBitColor = 0
                        Case "G": GetBitColor = 112
                        Case "B": GetBitColor = 192
                    End Select
                Case 3:
                    Select Case strColor
                        Case "R": GetBitColor = 146
                        Case "G": GetBitColor = 208
                        Case "B": GetBitColor = 80
                    End Select
                Case 4:
                    Select Case strColor
                        Case "R": GetBitColor = 112
                        Case "G": GetBitColor = 48
                        Case "B": GetBitColor = 160
                    End Select
                Case 5:
                    Select Case strColor
                        Case "R": GetBitColor = 255
                        Case "G": GetBitColor = 192
                        Case "B": GetBitColor = 0
                    End Select
            End Select
        Case "WR"
            Select Case iBit
                Case 1:
                    Select Case strColor
                        Case "R": GetBitColor = 192
                        Case "G": GetBitColor = 0
                        Case "B": GetBitColor = 0
                    End Select
                Case 2:
                    Select Case strColor
                        Case "R": GetBitColor = 0
                        Case "G": GetBitColor = 146
                        Case "B": GetBitColor = 255
                    End Select
                Case 3:
                    Select Case strColor
                        Case "R": GetBitColor = 176
                        Case "G": GetBitColor = 221
                        Case "B": GetBitColor = 127
                    End Select
                Case 4:
                    Select Case strColor
                        Case "R": GetBitColor = 159
                        Case "G": GetBitColor = 95
                        Case "B": GetBitColor = 207
                    End Select
                Case 5:
                    Select Case strColor
                        Case "R": GetBitColor = 255
                        Case "G": GetBitColor = 218
                        Case "B": GetBitColor = 101
                    End Select
            End Select
    End Select


End Function

Sub makeBL()
Attribute makeBL.VB_ProcData.VB_Invoke_Func = " \n14"
'
' makeBL Macro
'

    ActiveSheet.ChartObjects("RF_Chart").Activate

    ActiveChart.SeriesCollection(2).Select
    
    Selection.Format.Fill.Visible = msoFalse
    
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(146, 208, 80)
        .Transparency = 0
        .DashStyle = msoLineSolid
        .Weight = 0.25
    End With
End Sub

Public Sub MakePointsSolid(iCollection As Integer, iPoints As Integer, oChart As Excel.Chart)
    Dim i As Integer
    
    With oChart.SeriesCollection(iCollection)
'        For i = 1 To iPoints
'            .Points(i).Marker.Format.Line.DashStyle = msoLineSolid
''            .Points(i).Format.Line.Weight = 1
'        Next i
        .Format.Line.DashStyle = msoLineSysDot
    End With
    
End Sub

Public Sub SetBitButtons(iBit As Integer)
    Dim str As String
    
    If Sheets("HOME").Shapes("optBLBit" & iBit).ControlFormat.Value = 1 Then
        str = "Baseline"
    Else
        str = "Iteration"
    End If
    
    ActiveSheet.Shapes("Read_Des_Bit" & iBit).TextFrame.Characters.Text = "Read " & str
    ActiveSheet.Shapes("Clear_Bit" & iBit).TextFrame.Characters.Text = "Clear " & str
    ActiveSheet.Shapes("cmdRefreshBit" & iBit).TextFrame.Characters.Text = "Refresh " & str
    
End Sub

Public Function SeriesIsBaseline(iBit As Integer) As Boolean

    If Sheets("HOME").Shapes("optBLBit" & iBit).ControlFormat.Value = xlOn Then SeriesIsBaseline = True
    
End Function

Public Sub AddBitWarnStatus(strCurrentWarning As String, strNewWarning As String)

    If strCurrentWarning = "" Then
        strCurrentWarning = strNewWarning
    Else
        strCurrentWarning = strCurrentWarning & ", " & strNewWarning
    End If
    
End Sub

Public Function fXLNumberToLetter(lNum As Long) As String
    Dim strLet As String
    Dim iPass As Integer
    Dim lRem As Long
    Dim iTest As Integer

    If lNum > 16384 Then
        MsgBox "lNum too large. Abort."
        Exit Function
    End If
    
    If lNum > 676 Then 'At least greater than ZZ
        iPass = 3
    ElseIf lNum > 26 Then 'At least greater than Z
        iPass = 2
    Else
        iPass = 1
    End If
    
    lRem = lNum
    
    Do While iPass > 0
        If lRem Mod 26 = 0 Then
            If iPass = 1 Then
                iTest = 26
            Else
                iTest = WorksheetFunction.Floor((lRem - 1) / (26 ^ (iPass - 1)), 1)
            End If
        Else
            iTest = WorksheetFunction.Floor(lRem / (26 ^ (iPass - 1)), 1)
        End If
        
        Select Case iTest
            Case 1: strLet = "A"
            Case 2: strLet = "B"
            Case 3: strLet = "C"
            Case 4: strLet = "D"
            Case 5: strLet = "E"
            Case 6: strLet = "F"
            Case 7: strLet = "G"
            Case 8: strLet = "H"
            Case 9: strLet = "I"
            Case 10: strLet = "J"
            Case 11: strLet = "K"
            Case 12: strLet = "L"
            Case 13: strLet = "M"
            Case 14: strLet = "N"
            Case 15: strLet = "O"
            Case 16: strLet = "P"
            Case 17: strLet = "Q"
            Case 18: strLet = "R"
            Case 19: strLet = "S"
            Case 20: strLet = "T"
            Case 21: strLet = "U"
            Case 22: strLet = "V"
            Case 23: strLet = "W"
            Case 24: strLet = "X"
            Case 25: strLet = "Y"
            Case 26: strLet = "Z"
        End Select
        
        fXLNumberToLetter = fXLNumberToLetter & strLet
        
        lRem = lRem - WorksheetFunction.Floor(lRem / (26 ^ (iPass - 1)), 1) * (26 ^ (iPass - 1))
        
        iPass = iPass - 1
    Loop

'    Debug.Print "fXLNumberToLetter: " & lNum & " = " & fXLNumberToLetter

End Function

Public Function fXLLetterToNumber(strLet As String) As Integer
    Dim lAlpha(2) As Long
    Dim iLen As Integer
    
' Example
' 321
' WGT

    iLen = 0
    
    lAlpha(0) = 0
    lAlpha(1) = 0
    lAlpha(2) = 0
    
    Do While iLen < Len(strLet)
    
        Select Case UCase(Mid(strLet, Len(strLet) - iLen, 1))
            Case "A": lAlpha(iLen) = 1
            Case "B": lAlpha(iLen) = 2
            Case "C": lAlpha(iLen) = 3
            Case "D": lAlpha(iLen) = 4
            Case "E": lAlpha(iLen) = 5
            Case "F": lAlpha(iLen) = 6
            Case "G": lAlpha(iLen) = 7
            Case "H": lAlpha(iLen) = 8
            Case "I": lAlpha(iLen) = 9
            Case "J": lAlpha(iLen) = 10
            Case "K": lAlpha(iLen) = 11
            Case "L": lAlpha(iLen) = 12
            Case "M": lAlpha(iLen) = 13
            Case "N": lAlpha(iLen) = 14
            Case "O": lAlpha(iLen) = 15
            Case "P": lAlpha(iLen) = 16
            Case "Q": lAlpha(iLen) = 17
            Case "R": lAlpha(iLen) = 18
            Case "S": lAlpha(iLen) = 19
            Case "T": lAlpha(iLen) = 20
            Case "U": lAlpha(iLen) = 21
            Case "V": lAlpha(iLen) = 22
            Case "W": lAlpha(iLen) = 23
            Case "X": lAlpha(iLen) = 24
            Case "Y": lAlpha(iLen) = 25
            Case "Z": lAlpha(iLen) = 26
        End Select
        
        iLen = iLen + 1
    Loop
    
    fXLLetterToNumber = 26 * 26 * lAlpha(2) + 26 * lAlpha(1) + lAlpha(0)
    
'    Debug.Print "fXLLetterToNumber: " & strLet & " = " & fXLLetterToNumber

End Function

Public Sub subOpenIEPage(strAddress As String, Optional strSaveAs As String)
' FROM http://www.bigresource.com/VB-Accessing-an-open-webpage-BvldZzSX9L.html

    Dim ie As Object
    
    Set ie = CreateObject("InternetExplorer.Application")
    
    ie.Visible = True
    ie.Navigate strAddress

End Sub

Public Sub ExportZoneInterestButton()

    MsgBox "Export not yet functioning!!", vbCritical + vbOKOnly, "Not yet!"

End Sub

Public Sub ExportProfileChartButton()

    frmExportProfile.Show

End Sub

Public Sub ExportProfileCharts(ppApp As PowerPoint.Application, blnTemplate As Boolean)

    gProc = "ExportProfileCharts"
    
    Dim iError As Integer
    Dim strName As String, strTitle As String, strSheet As String
    Dim ppSlide As PowerPoint.Slide
    Dim ppShape As PowerPoint.Shapes
    Dim oChart As Variant, i As Integer, X As Integer, blnProfTitle As Boolean

    
    On Error GoTo ErrHandler
    
    Sheets("PROFILE").Activate
    
    For X = 1 To 3
        
        If Not frmExportProfile("chkProfile" & X) Then GoTo NextProfile
            
        ActiveSheet.ChartObjects("Profile" & X & "_Chart").Activate
        
    ' Create new profile plot
        Set oChart = ActiveChart 'Application.Worksheets("PROFILE").ChartObjects("Profile2_Chart") '.Duplicate
        
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = 73 'oChart.ChartType ' Application.Worksheets("PROFILE").ChartObjects("Profile2_Chart").ChartType
        
        Do While ActiveChart.SeriesCollection.Count > 0
            ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count).Delete
        Loop
        i = 1
        Do While ActiveChart.SeriesCollection.Count < oChart.SeriesCollection.Count
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(i).Name = oChart.SeriesCollection(i).Name
            ActiveChart.SeriesCollection(i).XValues = oChart.SeriesCollection(i).XValues
            ActiveChart.SeriesCollection(i).Values = oChart.SeriesCollection(i).Values
            ActiveChart.SeriesCollection(i).AxisGroup = oChart.SeriesCollection(i).AxisGroup
            ActiveChart.SeriesCollection(i).Format.Line.Weight = oChart.SeriesCollection(i).Format.Line.Weight
            ActiveChart.SeriesCollection(i).Format.Line.ForeColor.RGB = oChart.SeriesCollection(i).Format.Line.ForeColor.RGB
            ActiveChart.SeriesCollection(i).Format.Line.DashStyle = oChart.SeriesCollection(i).Format.Line.DashStyle
            ActiveChart.SeriesCollection(i).Smooth = oChart.SeriesCollection(i).Smooth
    
            i = i + 1
        Loop
        
        ActiveChart.SetElement msoElementLegendBottom
        
        ActiveChart.SetElement msoElementChartTitleAboveChart
        ActiveChart.ChartTitle.Text = oChart.ChartTitle.Text
        ActiveChart.ChartTitle.Characters.Font.Bold = oChart.ChartTitle.Characters.Font.Bold
    
    'Sizing chart
        ActiveChart.ChartArea.Width = oChart.ChartArea.Width
        ActiveChart.ChartArea.height = oChart.ChartArea.height
        ActiveChart.PlotArea.Width = oChart.PlotArea.Width
        ActiveChart.PlotArea.height = oChart.PlotArea.height
        ActiveChart.PlotArea.Left = oChart.PlotArea.Left
        ActiveChart.PlotArea.Top = oChart.PlotArea.Top
            
    'Horizontal Axis
        ActiveChart.SetElement msoElementPrimaryCategoryAxisTitleAdjacentToAxis
        ActiveChart.Axes(xlCategory).MajorTickMark = oChart.Axes(xlCategory).MajorTickMark
        With ActiveChart.Axes(xlCategory, xlPrimary)
            .AxisTitle.Text = oChart.Axes(xlCategory, xlPrimary).AxisTitle.Text
            .MinimumScale = oChart.Axes(xlCategory, xlPrimary).MinimumScale
            .MaximumScale = oChart.Axes(xlCategory, xlPrimary).MaximumScale
            .MajorUnit = oChart.Axes(xlCategory, xlPrimary).MajorUnit
            .MinorUnit = oChart.Axes(xlCategory, xlPrimary).MinorUnit
            .HasMajorGridlines = oChart.Axes(xlCategory, xlPrimary).HasMajorGridlines
            .HasMinorGridlines = oChart.Axes(xlCategory, xlPrimary).HasMinorGridlines
            .TickLabelPosition = xlHigh 'oChart.Axes(xlCategory, xlPrimary).TickLabelPosition
            .AxisTitle.Left = oChart.Axes(xlCategory, xlPrimary).AxisTitle.Left
            .AxisTitle.Top = oChart.Axes(xlCategory, xlPrimary).AxisTitle.Top
            .MajorTickMark = oChart.Axes(xlCategory, xlPrimary).MajorTickMark
            .MinorTickMark = oChart.Axes(xlCategory, xlPrimary).MinorTickMark
        End With
        
    'Vertical Axis
        ActiveChart.SetElement msoElementPrimaryValueAxisTitleRotated
        With ActiveChart.Axes(xlValue, xlPrimary)
            .AxisTitle.Text = oChart.Axes(xlValue, xlPrimary).AxisTitle.Text
            .MinimumScale = oChart.Axes(xlValue, xlPrimary).MinimumScale
            .MaximumScale = oChart.Axes(xlValue, xlPrimary).MaximumScale
            .MajorUnit = oChart.Axes(xlValue, xlPrimary).MajorUnit
            .MinorUnit = oChart.Axes(xlValue, xlPrimary).MinorUnit
            .HasMajorGridlines = oChart.Axes(xlValue, xlPrimary).HasMajorGridlines
            .HasMinorGridlines = oChart.Axes(xlValue, xlPrimary).HasMinorGridlines
            .TickLabelPosition = oChart.Axes(xlValue, xlPrimary).TickLabelPosition
            .AxisTitle.Left = oChart.Axes(xlValue, xlPrimary).AxisTitle.Left
            .AxisTitle.Top = oChart.Axes(xlValue, xlPrimary).AxisTitle.Top
            .MajorTickMark = oChart.Axes(xlValue, xlPrimary).MajorTickMark
            .MinorTickMark = oChart.Axes(xlValue, xlPrimary).MinorTickMark
        End With
        
        If IsZoI Then
            With ActiveChart.Axes(xlValue, xlSecondary)
                .MinimumScale = oChart.Axes(xlValue, xlSecondary).MinimumScale
                .MaximumScale = oChart.Axes(xlValue, xlSecondary).MaximumScale
                .MajorUnit = oChart.Axes(xlValue, xlSecondary).MajorUnit
                .MinorUnit = oChart.Axes(xlValue, xlSecondary).MinorUnit
                .HasMajorGridlines = oChart.Axes(xlValue, xlSecondary).HasMajorGridlines
                .HasMinorGridlines = oChart.Axes(xlValue, xlSecondary).HasMinorGridlines
                .TickLabelPosition = oChart.Axes(xlValue, xlSecondary).TickLabelPosition
                .MajorTickMark = oChart.Axes(xlValue, xlSecondary).MajorTickMark
                .MinorTickMark = oChart.Axes(xlValue, xlSecondary).MinorTickMark
            End With
        End If
        
        Sheets("PROFILE").Activate
        strName = ActiveChart.Name
        strName = Trim(Right(strName, Len(strName) - Len("PROFILE")))
        RenameChart "Sanitized_Profile", strName
    
    'Matching Gridlines
        ActiveChart.Axes(xlValue).bOrder.Weight = oChart.Axes(xlValue).bOrder.Weight
        ActiveChart.Axes(xlValue).bOrder.LineStyle = oChart.Axes(xlValue).bOrder.LineStyle
        
        ActiveChart.Axes(xlValue).MajorTickMark = oChart.Axes(xlValue).MajorTickMark
        ActiveChart.Axes(xlValue).MinorTickMark = oChart.Axes(xlValue).MinorTickMark
        ActiveChart.Axes(xlValue).TickLabelPosition = oChart.Axes(xlValue).TickLabelPosition
        
        ActiveChart.Axes(xlCategory).bOrder.Weight = oChart.Axes(xlCategory).bOrder.Weight
        ActiveChart.Axes(xlCategory).bOrder.LineStyle = oChart.Axes(xlCategory).bOrder.LineStyle
        
        ActiveChart.Axes(xlCategory).MajorTickMark = oChart.Axes(xlCategory).MajorTickMark
        ActiveChart.Axes(xlCategory).MinorTickMark = oChart.Axes(xlCategory).MinorTickMark
        ActiveChart.Axes(xlCategory).TickLabelPosition = oChart.Axes(xlCategory).TickLabelPosition
    
        If IsZoI Then
            ActiveChart.Axes(xlSecondary).bOrder.Weight = oChart.Axes(xlSecondary).bOrder.Weight
            ActiveChart.Axes(xlSecondary).bOrder.LineStyle = oChart.Axes(xlSecondary).bOrder.LineStyle
            
            ActiveChart.Axes(xlSecondary).MajorTickMark = oChart.Axes(xlSecondary).MajorTickMark
            ActiveChart.Axes(xlSecondary).MinorTickMark = oChart.Axes(xlSecondary).MinorTickMark
            ActiveChart.Axes(xlSecondary).TickLabelPosition = oChart.Axes(xlSecondary).TickLabelPosition
        End If
        
    'Add Confidential Note
        ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 286.4999212598, -0.0004724409, 72.75, 15.75).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Confidential"
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 12).ParagraphFormat.FirstLineIndent = 0
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 12).Font
    '        .NameComplexScript = "+mn-cs"
    '        .NameFarEast = "+mn-ea"
            .Size = 11
    '        .Name = "+mn-lt"
        End With
        With Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
            .Solid
        End With
        Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
        Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
    
        
    'Created Plot is complete, now to sanitize or not.
        ActiveSheet.ChartObjects("Sanitized_Profile").Activate
        
        If ActiveBits > 1 Then
            strTitle = "Profile Comparison"
        Else
            strTitle = "Profile"
        End If
        
        ActiveChart.ChartTitle.Text = strTitle
         
        If Range("XAxis") = "Radius" Then MakePlotGridSquareOfActiveChart
        
    ' Sanitize plot
        If frmExportProfile.optSanitize Then
            Sheets("PROFILE").Activate
            strName = Right(ActiveChart.Name, Len(ActiveChart.Name) - Len(ActiveSheet.Name) - 1)
            strSheet = ActiveSheet.Name
            SanitizeChart strName, strSheet
            strName = ""
        End If
            
NoSani:
        subUpdate True
    
    
    ' Make plots smooth or not
        subSmoothCurves "PROFILE", frmExportProfile.chkSmooth
        
'rename bit
        Worksheets("Home").Range("Bit1_Name") = "build" & ReadIniFileString("profile_title", "title1")
        Worksheets("Home").Range("Bit2_Name") = "build" & ReadIniFileString("profile_title", "title2")
'Make static title slide


        'del the first slide(template slide)
        If blnTitleFix = False Then
            ppApp.ActivePresentation.Slides.Range(Array(1)).Delete
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank") ' "2_Title Slide")
    '        ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
    '        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
            Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
            'define the first slide title
            ppSlide.Shapes.Title.TextFrame.TextRange.Text = "IDEAS Test Build Validation"
            'define the first slide subtitle
            ppSlide.Shapes(1).TextFrame.TextRange.Text = "Build " & ReadIniFileString("profile_title", "title1") & " vs " & "Build " & ReadIniFileString("profile_title", "title2")
            'three blank slide
            ' i = 0
            ' title_array = {"Important Build Notes", "Overview", "Design Tool Functionality Checks"}
            ' For i = 1 To 3
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Important Build Notes"
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Overview"
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Design Tool Functionality Checks"
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank")
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Static Analysis Results"
            blnTitleFix = True
        End If
        
        
        frmExportProfile.optTitleInc = True
        blnProfTitle = False
        If frmExportProfile.optTitleInc And blnProfTitle = False Then
            If blnTemplate Then
                ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Section Title") ' "2_Title Slide")
            Else
                ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
            End If
            ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
            Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)

        'make type title
            Casetype = RegularCheck(Worksheets("Home").Range("Bit1_Location"))
            If Casetype = "pdc" Then
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "PDC Static"
            ElseIf Casetype = "stinger" Then
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "StingBlade Static"
            ElseIf Casetype = "axe" Then
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "AxeBlade Static"
            ElseIf Casetype = "hyper" Then
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "HyperBlade Static"
            ElseIf Casetype = "px" Then
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "PX Static"
            ElseIf Casetype = "E616" Then
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "'EverythingBlade' Static"
            ElseIf Casetype = "echo" Then
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "StrataBlade Static"
            Else
'                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Comparison of Baseline and Proposed Bits"
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = str(Casetype) & "Static"
            End If
            blnProfTitle = True
            
        End If
        
        
        
        
        
        
        
        
        Sheets("PROFILE").Shapes.Range("Sanitized_Profile").Chart.ChartArea.Copy
        
        If blnTemplate Then
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only") ' "1_Title Only")
        Else
            ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
        End If
        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
        
        ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        
'Center pasted object in the slide
        If blnTemplate Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 7.5 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 9.35 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 3.33 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 0.9 * 72
        Else
            ppApp.ActiveWindow.Selection.ShapeRange.height = 5.8 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.31 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 1.34 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.19 * 72
        End If
        
        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = strTitle
        
        Set ppSlide = Nothing
        
    ' Delete Sanitized plot
        Sheets("PROFILE").Activate
        
        ActiveSheet.ChartObjects("Sanitized_Profile").Activate
        Selection.Delete
        
NextProfile:
    Next X
        


' New slide for Backrake/Siderake
    If frmExportProfile.chkBR Or frmExportProfile.chkSR Then
        If blnTemplate Then
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
        Else
            ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
        End If
        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
        
    'Titling the slide
        If frmExportProfile.chkBR And frmExportProfile.chkSR Then
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Back Rake/Side Rake Information"                       '.AddTitle.TextFrame.TextRange.Text = "Profile Comparison"
            
            If frmExportProfile.optSanitize Then
                Sanitize_Plot "Backrake_Chart", True
                Sanitize_Plot "Siderake_Chart", True
            End If
            
        ElseIf frmExportProfile.chkBR And Not frmExportProfile.chkSR Then
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Back Rake Information"                       '.AddTitle.TextFrame.TextRange.Text = "Profile Comparison"
            
            If frmExportProfile.optSanitize Then
                Sanitize_Plot "Backrake_Chart", True
            End If
            
        ElseIf Not frmExportProfile.chkBR And frmExportProfile.chkSR Then
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Side Rake Information"                       '.AddTitle.TextFrame.TextRange.Text = "Profile Comparison"
            
            If frmExportProfile.optSanitize Then
                Sanitize_Plot "Siderake_Chart", True
            End If
        End If
        
    'Scenerio both BR & SR
        If frmExportProfile.chkBR Then
            Sheets("PROFILE").Shapes.Range("BackRake_Chart").Chart.ChartArea.Copy
            
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkBR And frmExportProfile.chkSR Then
            'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 6.2 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.21 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.43 * 72
        ElseIf frmExportProfile.chkBR Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 5.71 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.13 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.07 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
        End If
'        ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True
        
        
        If frmExportProfile.chkSR Then
            Sheets("PROFILE").Shapes.Range("SideRake_Chart").Chart.ChartArea.Copy
        
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkBR And frmExportProfile.chkSR Then
        'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 6.2 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 8.08 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.43 * 72
        ElseIf frmExportProfile.chkSR Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 5.71 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.13 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.07 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
        End If
'        ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True
        
        Set ppSlide = Nothing
        
            
        If frmExportProfile.optSanitize Then
            Sanitize_Plot "Backrake_Chart", False
            Sanitize_Plot "Siderake_Chart", False
        End If
        
    End If



' New slide for Exposure/Roll Angle
    If frmExportProfile.chkExposures Or frmExportProfile.chkRollAngles Then
        If blnTemplate Then
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
        Else
            ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
        End If
        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
        
    'Titling the slide
        If frmExportProfile.chkExposures Or frmExportProfile.chkRollAngles Then
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Cutter Exposure/Roll Angle Information"                       '.AddTitle.TextFrame.TextRange.Text = "Profile Comparison"
            
            If frmExportProfile.optSanitize Then
                Sanitize_Plot "Exposure_Chart", True
                Sanitize_Plot "Roll_Chart", True
            End If
            
        ElseIf frmExportProfile.chkExposures And Not frmExportProfile.chkRollAngles Then
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Cutter Exposure Information"                       '.AddTitle.TextFrame.TextRange.Text = "Profile Comparison"
            
            If frmExportProfile.optSanitize Then
                Sanitize_Plot "Exposure_Chart", True
            End If
            
        ElseIf Not frmExportProfile.chkExposures And frmExportProfile.chkRollAngles Then
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Cutter Roll Angle Information"                       '.AddTitle.TextFrame.TextRange.Text = "Profile Comparison"
            
            If frmExportProfile.optSanitize Then
                Sanitize_Plot "Roll_Chart", True
            End If
        End If
        
    'Scenerio both BR & SR
        If frmExportProfile.chkExposures Then
            Sheets("PROFILE").Shapes.Range("Exposure_Chart").Chart.ChartArea.Copy
            
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkExposures And frmExportProfile.chkSR Then
            'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 6.2 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.21 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.43 * 72
        ElseIf frmExportProfile.chkExposures Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 5.71 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.13 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.07 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
        End If
'        ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True
        
        
        If frmExportProfile.chkRollAngles Then
            Sheets("PROFILE").Shapes.Range("Roll_Chart").Chart.ChartArea.Copy
        
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkExposures And frmExportProfile.chkRollAngles Then
        'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 6.2 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 8.08 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.43 * 72
        ElseIf frmExportProfile.chkRollAngles Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 5.71 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.13 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.07 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
        End If
'        ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True
        
        Set ppSlide = Nothing
        
            
        If frmExportProfile.optSanitize Then
            Sanitize_Plot "Exposure_Chart", False
            Sanitize_Plot "Roll_Chart", False
        End If
        
    End If


'Cutter Summary output

    If frmExportProfile.chkCutOld Then
    ' New chart for cutter comparison
        If blnTemplate Then
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only") '"1_Title Only")
        Else
            ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
        End If
        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
        
        If Not ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.HasTitle Then ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Layout = ppLayoutTitleOnly
        
        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Cutter Information"                       '.AddTitle.TextFrame.TextRange.Text = "Profile Comparison"
        
    ' Add cutter distances
        Sheets("PROFILE").Shapes.Range("Cutter_Dist_Chart").Chart.ChartArea.Copy
        
        ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
                        
        'Move pasted object in the slide
        If blnTemplate Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 4.8 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 5.99 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 1.93 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 0.91 * 72
        Else
            ppApp.ActiveWindow.Selection.ShapeRange.height = 3.76 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 4.7 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.17 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.22 * 72
        End If

        
        
    ' new 5"  old 7.31"
        ActiveSheet.ChartObjects("Total_Cutter_Count").Activate
        strName = Right(ActiveChart.Name, Len(ActiveChart.Name) - Len(ActiveSheet.Name) - 1)
        ActiveSheet.Shapes(strName).ScaleWidth (5 / 7.31), msoFalse, msoScaleFromTopLeft
        
        Sheets("PROFILE").Shapes.Range("Total_Cutter_Count").Chart.ChartArea.Copy
        
        ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
                        
        'Center pasted object in the slide
        If blnTemplate Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 4.8 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 5.99 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 8.03 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 0.91 * 72
        Else
            ppApp.ActiveWindow.Selection.ShapeRange.height = 3.76 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 4.7 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 5.13 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.22 * 72
        End If
        
        
        ActiveSheet.Shapes(strName).ScaleWidth 7.31 / 5, msoFalse, msoScaleFromTopLeft
        
        
        Sheets("INFO").Activate
        Application.GoTo Reference:="CS_Info_All"
        Selection.Copy
        Range("A1").Select
        
    '    ppApp.ActiveWindow.View.PasteSpecial ppPasteEnhancedMetafile
        ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
                        
    ' Center pasted object in the slide
        ppApp.ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
        
        If ActiveBits > 1 Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 2.69 * 72
        Else
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 12.67 * 72
        End If
    
        If blnTemplate Then
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 5.81 * 72
        Else
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 5.07 * 72
        End If
        ppApp.ActiveWindow.Selection.ShapeRange.Left = ((16 - ppApp.ActiveWindow.Selection.ShapeRange.Width / 72) / 2) * 72
    End If
    
    If frmExportProfile.chkCutNew Then
    ' New chart for cutter comparison
        If blnTemplate Then
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only") '"1_Title Only")
        Else
            ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
        End If
        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
        
        If Not ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.HasTitle Then ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Layout = ppLayoutTitleOnly
        
        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Cutter Information"                       '.AddTitle.TextFrame.TextRange.Text = "Profile Comparison"
    
    
        Sheets("INFO2").Activate
        
        Range("INFO2_Table").Font.Size = 14
        Range("INFO2_Table").ShrinkToFit = True
        
        Range(Cells(1, 1), Cells(Range("Bit1_I2_GaugeSR").row, ActiveBits + 2)).Select
        
        Selection.Copy
        
        
        Range("A1").Select
        
    '    ppApp.ActiveWindow.View.PasteSpecial ppPasteEnhancedMetafile
        
        DoEvents
        
        ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        
        Range("INFO2_Table").ShrinkToFit = False
                       
    ' Center pasted object in the slide
        ppApp.ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoTrue
        
        ppApp.ActiveWindow.Selection.ShapeRange.height = 7.5 * 72
        
        If blnTemplate Then
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 0.95 * 72
        Else
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.42 * 72
        End If
        ppApp.ActiveWindow.Selection.ShapeRange.Left = ((16 - ppApp.ActiveWindow.Selection.ShapeRange.Width / 72) / 2) * 72

    End If
    
    Sheets("PROFILE").Activate
    
    Exit Sub
    
ErrHandler:

If Err = 1004 Then   'Paste method of Worksheet class failed
    iError = iError + 1
    Debug.Print "Errored " & iError & " times."
    If iError = 10 Then
        MsgBox "There has been an error. Please try again."
        Exit Sub
    End If
    Resume
ElseIf Err = 424 Then ' Object required
'        Stop
        Resume Next

ElseIf Err = -2147467259 Then   'Automation error

    Resume Next

Else
    Debug.Print Err, Err.Description
    
    Stop
    
    Resume
End If

End Sub



Public Sub ExportCharts(ppApp As PowerPoint.Application, blnStaticTitle As Boolean, blnTemplate As Boolean)

    gProc = "ExportCharts"
    
    Dim iError As Integer
    Dim strName As String, strTitle As String, strSheet As String
'    Dim ppApp As PowerPoint.Application
    Dim ppSlide As PowerPoint.Slide
    Dim oChart As Variant, i As Integer, X As Integer
    Dim strCharts(12, 1) As String
    
    On Error GoTo ErrHandler

    strCharts(0, 0) = "ROP"
    strCharts(1, 0) = "ROPChange"
    strCharts(2, 0) = "Torque"
    strCharts(3, 0) = "TqChange"
    strCharts(4, 0) = "ROPTq"
    strCharts(5, 0) = "ROPTqChange"
    strCharts(6, 0) = "DoC"
    strCharts(6, 1) = "Bit Advance Comparison"
    strCharts(7, 0) = "TIF"
    strCharts(7, 1) = "Total Imbalance Force Comparison"
    strCharts(8, 0) = "Beta"
    strCharts(8, 1) = "Beta Angle Comparison"
    strCharts(9, 0) = "CIF"
    strCharts(9, 1) = "Circumferential Imbalance Force Comparison"
    strCharts(10, 0) = "RIF"
    strCharts(10, 1) = "Radial Imbalance Force Comparison"
    strCharts(11, 0) = "RIF_CIF"
    strCharts(11, 1) = "Imbalance Force Ratio Comparison"
    strCharts(12, 0) = "RIF_CIF_Normal"
    strCharts(12, 1) = "Normalized Imbalance Force Ratio Comparison"

    For i = 0 To UBound(strCharts)
        If frmExportProfile("chk" & strCharts(i, 0)) Then GoTo ContinueExport
    Next i
    
    frmExportProfile.lblMsg.Caption = "No static charts are chosen for export."
    Exit Sub
    
ContinueExport:

'Make static title slide
    If frmExportProfile.optTitleInc And blnStaticTitle = False Then
        If blnTemplate Then
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Section Title") '"2_Title Slide")
        Else
            ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
        End If
        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
    
        ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Static Analysis Results"
    
        blnStaticTitle = True
    End If

    
' New slides for ROP/% Diff
    If frmExportProfile.chkROP Or frmExportProfile.chkROPChange Then
        If blnTemplate Then
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only") '"1_Title Only")
        Else
            ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
        End If
        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
        
        If ActiveBits > 1 Then
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = Range("Control") & " Comparison"
        Else
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = Range("Control")
        End If
        
        
        If frmExportProfile.chkROP Then
            Sheets("SUMMARY_PLOTS").Shapes.Range("ROP").Chart.ChartArea.Copy
            
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkROP And frmExportProfile.chkROPChange Then
            'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 3.93 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.85 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.11 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 2.68 * 72
        ElseIf frmExportProfile.chkROP Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 4.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 9.72 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.14 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.65 * 72
        End If
        
        If frmExportProfile.chkROPChange Then
            Sheets("SUMMARY_PLOTS").Shapes.Range("PercDiffROP").Chart.ChartArea.Copy
        
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkROP And frmExportProfile.chkROPChange Then
        'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 3.93 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.85 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 8.04 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 2.68 * 72
        ElseIf frmExportProfile.chkROPChange Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 4.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 9.72 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.14 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.65 * 72
        End If
        
        Set ppSlide = Nothing
    End If


' New slides for Torque/% Diff
    If frmExportProfile.chkTorque Or frmExportProfile.chkTqChange Then
        If blnTemplate Then
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only") '"1_Title Only")
        Else
            ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
        End If
        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
        
        If ActiveBits > 1 Then
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Torque Comparison"
        Else
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Torque"
        End If
        
        
        If frmExportProfile.chkTorque Then
            Sheets("SUMMARY_PLOTS").Shapes.Range("Torque").Chart.ChartArea.Copy
            
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkTorque And frmExportProfile.chkTqChange Then
            'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 3.93 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.85 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.11 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 2.68 * 72
        ElseIf frmExportProfile.chkROP Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 4.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 9.72 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.14 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.65 * 72
        End If
        
        If frmExportProfile.chkTqChange Then
            Sheets("SUMMARY_PLOTS").Shapes.Range("PercDiffTq").Chart.ChartArea.Copy
        
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkTorque And frmExportProfile.chkTqChange Then
        'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 3.93 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.85 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 8.04 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 2.68 * 72
        ElseIf frmExportProfile.chkROPChange Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 4.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 9.72 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.14 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.65 * 72
        End If
        
        Set ppSlide = Nothing
    End If


' New slides for ROP/Torque / % Diff
    If frmExportProfile.chkROPTq Or frmExportProfile.chkROPTqChange Then
        If blnTemplate Then
            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only") '"1_Title Only")
        Else
            ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
        End If
        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
        
        If ActiveBits > 1 Then
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "ROP/Torque Comparison"
        Else
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "ROP/Torque"
        End If
        
        
        If frmExportProfile.chkROPTq Then
            Sheets("SUMMARY_PLOTS").Shapes.Range("ROPTq").Chart.ChartArea.Copy
            
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkROPTq And frmExportProfile.chkROPTqChange Then
            'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 3.93 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.85 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.11 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 2.68 * 72
        ElseIf frmExportProfile.chkROPTq Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 4.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 9.72 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.14 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.65 * 72
        End If
        
        If frmExportProfile.chkROPTqChange Then
            Sheets("SUMMARY_PLOTS").Shapes.Range("PercDiffROPTq").Chart.ChartArea.Copy
        
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
        End If
        
        If frmExportProfile.chkROPTq And frmExportProfile.chkROPTqChange Then
        'Move pasted object in the slide
            ppApp.ActiveWindow.Selection.ShapeRange.height = 3.93 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 7.85 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 8.04 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 2.68 * 72
        ElseIf frmExportProfile.chkROPChange Then
            ppApp.ActiveWindow.Selection.ShapeRange.height = 4.75 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Width = 9.72 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.14 * 72
            ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.65 * 72
        End If
        
        Set ppSlide = Nothing
    End If



' New slides for Summary Plots

    For i = 6 To UBound(strCharts)
        If frmExportProfile("chk" & strCharts(i, 0)) Then
            If blnTemplate Then
                ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only") ' "1_Title Only")
            Else
                ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
            End If
            ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
            Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
            
            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = strCharts(i, 1)
            
            
            Sheets("SUMMARY_PLOTS").Shapes.Range(strCharts(i, 0)).Chart.ChartArea.Copy
            
            ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse).Select
            
            If blnTemplate Then
                ppApp.ActiveWindow.Selection.ShapeRange.height = 7 * 72
                ppApp.ActiveWindow.Selection.ShapeRange.Width = 14 * 72
                ppApp.ActiveWindow.Selection.ShapeRange.Left = 1 * 72
                ppApp.ActiveWindow.Selection.ShapeRange.Top = 1 * 72
            Else
                ppApp.ActiveWindow.Selection.ShapeRange.height = 4.86 * 72
                ppApp.ActiveWindow.Selection.ShapeRange.Width = 9.72 * 72
                ppApp.ActiveWindow.Selection.ShapeRange.Left = 0.14 * 72
                ppApp.ActiveWindow.Selection.ShapeRange.Top = 1.65 * 72
            End If
            Set ppSlide = Nothing
        End If
    
    Next i
    
    
    
    
    
    
    
    
    Sheets("PROFILE").Activate

    Exit Sub
    
ErrHandler:

If Err = 1004 Then   'Paste method of Worksheet class failed
    iError = iError + 1
    Debug.Print "Errored " & iError & " times."
    If iError = 10 Then
        MsgBox "There has been an error. Please try again."
        Exit Sub
    End If
    Resume
ElseIf Err = 424 Then ' Object required
'        Stop
        Resume Next

'ElseIf Err = -2147188160 Then
'    Resume
    
Else
    Stop
    
    Debug.Print Err, Err.Description
    
    Resume
End If

End Sub














Public Sub SanitizeChart(strChart As String, strSheet As String)

    gProc = "SanitizeChart"
    

Sheets(strSheet).ChartObjects(strChart).Activate

    For Each A In ActiveChart.Axes
        A.HasMajorGridlines = False
        A.HasMinorGridlines = False
    Next A
    
    With ActiveChart.Axes(xlValue)
        With .bOrder
            .Weight = xlHairline
            .LineStyle = xlNone
        End With
        
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNone
    End With
    
    
    With ActiveChart.Axes(xlCategory)
        With .bOrder
            .Weight = xlHairline
            .LineStyle = xlNone
        End With
        
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNone
    End With

End Sub

Public Function ActiveBits() As Integer
    Dim iBit As Integer
    
    For iBit = 1 To Range("MaxBits")
        If Sheets("HOME").Shapes("Hide_Bit" & iBit).ControlFormat.Value <> xlOn Then
            ActiveBits = ActiveBits + 1
        ElseIf Range("Bit" & iBit & "_Location") <> "" Then
            ActiveBits = ActiveBits + 1
        End If
    Next iBit
    
End Function

Public Function BitIsActive(iBit As Integer) As Boolean
    
    BitIsActive = False
    
    If Sheets("HOME").Shapes("Hide_Bit" & iBit).ControlFormat.Value <> xlOn Then
        BitIsActive = True
'    ElseIf Range("Bit" & iBit & "_Location") <> "" Then
'        BitIsActive = True
    End If
    
End Function

Public Sub SummaryColors(Optional strColor As String, Optional blnExportProfile As Boolean)

    gProc = "SummaryColors"
    
    Dim str As String
    Dim i As Integer
    
    If strColor = "" Then strColor = Application.Caller
    
    Sheets("INFO").Activate
    
    If Not blnExportProfile Then
        
    '    ActiveSheet.OptionButtons(strColor).Value = xlOn
        
        str = ActiveCell.Address
        
        subUpdate False
    End If
    
    If strColor = "optColorNew" Then
        Range("H4:P7").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 9408511
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Range("H8:P11").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 16758615
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Range("H12:P15").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10806983
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Range("H16:P19").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 14586557
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Range("H20:P23").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 9486586
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    ElseIf strColor = "optColorClassic" Then
        
        For i = 1 To 5
            Range("H" & 4 * i + 0 & ":P" & 4 * i + 0).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 13434828
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("H" & 4 * i + 1 & ":P" & 4 * i + 1).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 10092543
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("H" & 4 * i + 2 & ":P" & 4 * i + 2).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16764057
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("H" & 4 * i + 3 & ":P" & 4 * i + 3).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 10079487
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next i
        
    End If
    
    If blnExportProfile Then Exit Sub
    
    Range(str).Select
    subUpdate True
End Sub



Public Sub UpdateForceRatios()
    Dim iBit As Integer, iCase As Integer, iRow As Integer, iMaxBits As Integer
    Dim strTemp As String
    
    On Error GoTo ErrHandler

    subUpdate False
    
    For iBit = 1 To Application.Range("MaxBits")
      Sheets("Bit" & iBit).Activate
        For iRow = 0 To Application.Range("Bit" & iBit & "_Profile_Angle").Rows.Count
            If Range("ForceRatioSplitIO") = "Inner" Then
                If Application.Cells(Application.Range("Bit" & iBit & "_Profile_Angle").row + iRow, Application.Range("Bit" & iBit & "_Profile_Angle").Column).Value > Range("ForceRatioSplitInc") Then
'                        iRow = Right(Application.Cells(Application.Range("Bit" & iBit & "_Profile_Angle").Row + iRow, Application.Range("Bit" & iBit & "_Profile_Angle").Column).Address, Len(Application.Cells(Application.Range("Bit" & iBit & "_Profile_Angle").Row + iRow, Application.Range("Bit" & iBit & "_Profile_Angle").Column).Address) - InStrRev(Application.Cells(Application.Range("Bit" & iBit & "_Profile_Angle").Row + iRow, Application.Range("Bit" & iBit & "_Profile_Angle").Column).Address, "$"))
                    Exit For
                End If
            ElseIf Range("ForceRatioSplitIO") = "Outer" Then
                If Application.Cells(Application.Range("Bit" & iBit & "_Profile_Angle").row + iRow, Application.Range("Bit" & iBit & "_Profile_Angle").Column).Value >= Range("ForceRatioSplitInc") Then
                    Exit For
                End If
            End If
        Next iRow
        
        iRow = Right(Application.Cells(Application.Range("Bit" & iBit & "_Profile_Angle").row + iRow, Application.Range("Bit" & iBit & "_Profile_Angle").Column).Address, Len(Application.Cells(Application.Range("Bit" & iBit & "_Profile_Angle").row + iRow, Application.Range("Bit" & iBit & "_Profile_Angle").Column).Address) - InStrRev(Application.Cells(Application.Range("Bit" & iBit & "_Profile_Angle").row + iRow, Application.Range("Bit" & iBit & "_Profile_Angle").Column).Address, "$"))
    
    'iRow now equals first row of "outer" forces
        For iCase = 1 To Application.Range("Bit" & iBit & "_Cases")
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Normal").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Normal").Address, ":") - 1)
            Application.Range(strTemp, Left(strTemp, InStrRev(strTemp, "$")) & iRow - 1).Name = "Bit" & iBit & "_Case" & iCase & "_NormalInner"
            
            strTemp = Right(Application.Range("Bit" & iBit & "_Case" & iCase & "_Normal").Address, Len(Application.Range("Bit" & iBit & "_Case" & iCase & "_Normal").Address) - InStrRev(Application.Range("Bit" & iBit & "_Case" & iCase & "_Normal").Address, ":"))
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRow, strTemp).Name = "Bit" & iBit & "_Case" & iCase & "_NormalOuter"
            
        Next iCase
    Next iBit
    
    Sheets("FORCE_RATIOS").Activate
    
    subUpdate True
    
    Exit Sub

ErrHandler:

Debug.Print Err, Err.Description

Stop
Resume

End Sub

Public Sub CoreDiamCutterUpdate(Optional iBitNum As Integer)
    Dim ActiveRange As Range
    
    On Error Resume Next

    subUpdate False
    Set ActiveRange = Range(Selection.Address)
    
    CoreDiamColorCells Range("CoreBit")
    
    If Sheets("HOME").Shapes("chkCalcCore").ControlFormat.Value <> 1 Then
        Exit Sub
    ElseIf Range("Bit" & Range("CoreBit") & "_Location") = "" Then
        ActiveSheet.Shapes.Range(Array("cmbCoreCutter")).Select
        With Selection
            .ListFillRange = ""
            .LinkedCell = Range("CoreCutter").Address
            .DropDownLines = 0
            .Display3DShading = False
        End With
        
        Range("CoreDiam") = ""
        Range("StingerTip") = ""
        
        subUpdate True
        
        ActiveRange.Select


        Exit Sub
'        Stop
    End If

    If iBitNum > 0 And iBitNum <> Range("CoreBit") Then Exit Sub

'    With ActiveSheet.Shapes.Range(Array("cmbCoreCutter"))
    ActiveSheet.Shapes.Range(Array("cmbCoreCutter")).Select
    With Selection
        .ListFillRange = "Bit" & Range("CoreBit") & "_Cutter_Num"
        .LinkedCell = Range("CoreCutter").Address
        If Range("Bit" & Range("CoreBit") & "_Cutter_Num").Count < 15 Then
            .DropDownLines = Range("Bit" & Range("CoreBit") & "_Cutter_Num").Count
        Else
            .DropDownLines = 15
        End If
        .Display3DShading = False
    End With

    subUpdate True
    
    ActiveRange.Select
    
    CoreDiamCalculate

End Sub

Public Sub CoreDiamCalculate(Optional iForceBit As Integer, Optional iForceCutter As Integer, Optional sForceResult As Single)
    Dim iRow As Integer, iBit As Integer
    Dim dblAngle As Double ' Cone angle for specific cutter
    Dim sngRtip As Single  ' Radius for cutter tip
    Dim sngRcut As Single  ' Radius of cutter
    Dim sngRcor As Single  ' Core Radius
    
    On Error GoTo ErrHandler
    
    If iForceBit <> 0 And iForceCutter <> 0 Then
        iBit = iForceBit
        GoTo CalcCoreDiam
    Else
        iBit = Range("CoreBit")
    End If
    
    subUpdate False
    
    Sheets("BIT" & iBit).Select
    
    iRow = Range("Bit" & iBit & "_Cutter_Num").row
    
    Do While Cells(iRow, 1) <> Range("CoreCutter")
        iRow = iRow + 1
    Loop
    
    If IsStinger(Cells(iRow, Range("Bit" & iBit & "_Cutter_Type").Column)) Then
        Sheets("HOME").Select
    
        Range("CoreDiam") = "C" & Range("CoreCutter") & " is stinger"
        Range("StingerTip") = ""
        
    Else
    
CalcCoreDiam:
        If iForceBit <> 0 And iForceCutter <> 0 Then
            Sheets("BIT" & iForceBit).Select
            Do Until Range("Bit" & iForceBit & "_Cutter_Num").Item(iRow) = iForceCutter
                iRow = iRow + 1
            Loop
            iRow = Range("Bit" & iForceBit & "_Cutter_Num").Item(iRow).row
        End If
        Range("StingHelp_CutType") = CStr(Cells(iRow, Range("Bit" & iBit & "_Cutter_Type").Column))
        Range("StingHelp_ProfAng") = Cells(iRow, Range("Bit" & iBit & "_Profile_Angle").Column)
        Range("StingHelp_CutTipRad") = Cells(iRow, Range("Bit" & iBit & "_Radius").Column)
        Range("StingHelp_CutTipHt") = Cells(iRow, Range("Bit" & iBit & "_Height").Column)
        Range("StingHelp_BR") = Cells(iRow, Range("Bit" & iBit & "_Back_Rake").Column)
        Range("StingHelp_SR") = Cells(iRow, Range("Bit" & iBit & "_Side_Rake").Column)
        
        If iForceBit <> 0 And iForceCutter <> 0 Then
            sForceResult = Range("StingHelp_CoreDiam")
            Exit Sub
        End If
        
        If IsNumeric(Range("StingHelp_CoreDiam")) Then Range("CoreDiam") = Range("StingHelp_CoreDiam")
        If IsNumeric(Range("StingHelp_CoreHt")) Then Range("StingerTip") = Range("StingHelp_CoreHt") - 0.375
        
        Sheets("HOME").Select


' Old Measurements
'        Range("CoreDiam") = RoundOff(2 * (sngRtip - sngRcut * (1 - Cos((90 - dblAngle) * WorksheetFunction.Pi / 180))), 5)
'        Range("StingerTip") = RoundOff(sngHtip - sngRcut * Sin((90 - dblAngle) * WorksheetFunction.Pi / 180) - 0.375, 4)
    End If
    
    Exit Sub
    
ErrHandler:

    Debug.Print Now(), gProc, Err, Err.Description
    Stop
    Resume
    
End Sub

Public Sub CoreDiamActivate()

    If Sheets("HOME").Shapes("chkCalcCore").ControlFormat.Value = 1 Then
        'Active
        CoreDiamColorCells Range("CoreBit")
        
        ActiveSheet.Shapes.Range(Array("cmbCoreBit")).Select
        With Selection
            .ListFillRange = "NumBits"
            .LinkedCell = "CoreBit"
            .DropDownLines = Range("MaxBits")
            .Display3DShading = False
        End With
'        ActiveSheet.Shapes.Range(Array("cmbCoreCutter")).Select
'        With Selection
'            .ListFillRange = "Bit" & Range("CoreBit") & "_Cutter_Num"
'            .LinkedCell = "CoreCutter"
'            If Range("Bit" & Range("CoreBit") & "_Cutter_Num").Count < 15 Then
'                .DropDownLines = Range("Bit" & Range("CoreBit") & "_Cutter_Num").Count
'            Else
'                .DropDownLines = 15
'            End If
'            .Display3DShading = False
'        End With

    Else
        'Inactive
        CoreDiamColorCells 0
        
        ActiveSheet.Shapes.Range(Array("cmbCoreCutter")).Select
        With Selection
            .ListFillRange = ""
            .LinkedCell = "CoreCutter"
            .DropDownLines = 1
            .Display3DShading = False
        End With
        ActiveSheet.Shapes.Range(Array("cmbCoreBit")).Select
        With Selection
            .ListFillRange = ""
            .LinkedCell = "CoreBit"
            .DropDownLines = 1
            .Display3DShading = False
        End With
        
        Range("CoreDiam") = ""
        Range("StingerTip") = ""
    End If
    
    Range("Rest").Select

End Sub

Public Sub CoreDiamColorCells(iBit As Integer)

    With Range("QuickCoreCalc")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .PatternTintAndShade = 0
        End With
        With .Font
            .TintAndShade = 0
        End With
        
        Select Case iBit
            Case 0
                .Interior.Color = RGB(192, 192, 192)
                .Font.Color = RGB(128, 128, 128)
            Case 1
                .Interior.Color = RGB(255, 180, 180)
                .Font.Color = RGB(0, 0, 0)
            Case 2
                .Interior.Color = RGB(180, 222, 255)
                .Font.Color = RGB(0, 0, 0)
            Case 3
                .Interior.Color = RGB(198, 230, 162)
                .Font.Color = RGB(0, 0, 0)
            Case 4
                .Interior.Color = RGB(222, 200, 238)
                .Font.Color = RGB(0, 0, 0)
            Case 5
                .Interior.Color = RGB(252, 211, 152)
                .Font.Color = RGB(0, 0, 0)
        End Select

    End With

    With Range("QuickCoreTipHt")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .PatternTintAndShade = 0
        End With
        With .Font
            .TintAndShade = 0
        End With
        
        Select Case iBit
            Case 0
                .Interior.Color = RGB(192, 192, 192)
                .Font.Color = RGB(128, 128, 128)
            Case 1
                .Interior.Color = RGB(255, 217, 217)
                .Font.Color = RGB(0, 0, 0)
            Case 2
                .Interior.Color = RGB(217, 239, 255)
                .Font.Color = RGB(0, 0, 0)
            Case 3
                .Interior.Color = RGB(223, 241, 203)
                .Font.Color = RGB(0, 0, 0)
            Case 4
                .Interior.Color = RGB(239, 229, 247)
                .Font.Color = RGB(0, 0, 0)
            Case 5
                .Interior.Color = RGB(253, 228, 191)
                .Font.Color = RGB(0, 0, 0)
        End Select

    End With



End Sub

Sub Option_Button_Hide_Zone_Click()

    ZoI "Hide"
    gblnZoIShtOverride = False
    
    Sheets("HOME").Shapes("Option_Button_Hide_Zone").ControlFormat.Value = xlOn
    
End Sub

Sub Option_Button_Show_zone_Click()

    ZoI "Show"
    gblnZoIShtOverride = False
    
    Sheets("HOME").Shapes("Option_Button_Show_Zone").ControlFormat.Value = xlOn

End Sub

Public Sub CalculateZoneForces()
    Dim iBit As Integer, iCase As Integer, iRowI As Integer, iRowO As Integer, iMaxBits As Integer
    Dim strTemp As String
    
    On Error GoTo ErrHandler

    subUpdate False
    
    For iBit = 1 To Application.Range("MaxBits")
        Sheets("Bit" & iBit).Activate
        For iRowI = 0 To Application.Range("Bit" & iBit & "_Radius").Rows.Count
                If Application.Cells(Application.Range("Bit" & iBit & "_Radius").row + iRowI, Application.Range("Bit" & iBit & "_Radius").Column).Value > Range("Zinner") Then
                    Exit For
                End If
        Next iRowI
        
        For iRowO = 0 To Application.Range("Bit" & iBit & "_Radius").Rows.Count
                If Application.Cells(Application.Range("Bit" & iBit & "_Radius").row + iRowO, Application.Range("Bit" & iBit & "_Radius").Column).Value > Range("Zouter") Then
                    Exit For
                End If
        Next iRowO
        
        
        'iRow = Right(Application.Cells(Application.Range("Bit" & iBit & "_Radius").Row + iRow, Application.Range("Bit" & iBit & "_Radius").Column).Address, Len(Application.Cells(Application.Range("Bit" & iBit & "_Radius").Row + iRow, Application.Range("Bit" & iBit & "_Radius").Column).Address) - InStrRev(Application.Cells(Application.Range("Bit" & iBit & "_Radius").Row + iRow, Application.Range("Bit" & iBit & "_Radius").Column).Address, "$"))
    
    'iRowI equals -1 row of "outer" forces
        
        
        For iCase = 1 To Application.Range("Bit" & iBit & "_Cases")
            'for normal force
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Normal").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Normal").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZNormal"
            'for workrate-normal
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Normal_Workrate").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Normal_Workrate").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZNormal_Workrate"
            'for circumferential force
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Circumferential").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Circumferential").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZCircumferential"
            'for workrate-circumferential
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Circumferential_Workrate").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Circumferential_Workrate").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZCircumferential_Workrate"
            'for vertical force
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Vertical").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Vertical").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZVertical"
            'for side force
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Side").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Side").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZSide"
            'for WearFlatArea
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_WearFlatArea").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_WearFlatArea").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZWearFlatArea"
            'for AreaCut
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_AreaCut").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_AreaCut").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZAreaCut"
            'for VolCut
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_VolCut").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_VolCut").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZVolCut"
            'for ProE_fx
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_ProE_fx").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_ProE_fx").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZProE_fx"
            'for ProE_fy
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_ProE_fy").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_ProE_fy").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZProE_fy"
            'for ProE_fz
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_ProE_fz").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_ProE_fz").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZProE_fz"
            'for Cutter Torque
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Cutter_Torque").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Cutter_Torque").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZCutter_Torque"
            'for Resultant Force
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Resultant").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Resultant").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZResultant"
            'for workrate-Resultant Force
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Resultant_Workrate").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Resultant").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZResultant_Workrate"
            'for Delam Force
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Delam").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Resultant").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZDelam"
            'for workrate-Delam Force
            strTemp = Left(Application.Range("Bit" & iBit & "_Case" & iCase & "_Delam_Workrate").Address, InStr(1, Application.Range("Bit" & iBit & "_Case" & iCase & "_Resultant").Address, ":") - 1)
            Application.Range(Left(strTemp, InStrRev(strTemp, "$")) & iRowI + 5, Left(strTemp, InStrRev(strTemp, "$")) & iRowO + 4).Name = "Bit" & iBit & "_Case" & iCase & "_ZDelam_Workrate"
        Next iCase
    Next iBit

    subZoneBoxChangeUpdateListBox 1
    Sheets("ZONE_OF_INTEREST").ListBoxes("PlotBox1").Value = 1
    
    
    
    Exit Sub

ErrHandler:

Debug.Print Err, Err.Description

Stop
Resume

End Sub

Public Sub subZoneBoxChange(Optional iBox As Integer, Optional iCase As Integer, Optional strSheetName As String)
    Dim iBit As Integer
    Dim sSum As Single, i As Integer, X As Integer, strCol() As Variant
    Dim strForces As String
    Dim sAvg As Single
    

    On Error GoTo ErrHandler
    
    strForces = "ForceCase"
    If strSheetName = "" Then strSheetName = ActiveSheet.Name
    
'    If Application.Caller = "PlotItrComp" Then
'        iBox = 0
'        iCase = Sheets(strSheetName).ListBoxes(Application.Caller).Value
'    ElseIf iBox = 0 Then
'        iBox = Right(Application.Caller, 1)
    If iBox = 0 Then iBox = 1
'    End If
    
    If strSheetName = "ZONE_OF_INTEREST" Then
        If IsError(Range("ForceCase" & iBox)) Then Exit Sub
    End If
    
    If iCase = 0 Then iCase = Sheets(strSheetName).ListBoxes("PlotBox" & iBox).Value
'    If iCase = 0 Then
'        iCase = Sheets(strSheetName).ListBoxes("PlotBox" & iBox).Value
'    ElseIf Application.Caller <> "PlotItrComp" Then
'        Sheets(strSheetName).ListBoxes("PlotBox" & iBox) = iCase
'    End If
    
    subZoneBoxChangeUpdateListBox iCase
    
    
'    strCol() = Array("ZNormal", "ZNormal_Workrate", "ZCircumferential", "ZCircumferential_Workrate", "ZVertical", "ZWearFlatArea", "ZAreaCut", "ZVolCut", "ZProE_fx", "ZProE_fy", "ZProE_fz", "ZCutter_Torque", "ZResultant", "ZResultant_Workrate", "ZDelam", "ZDelam_Workrate")
'    x = 0
'
'    For iBit = 1 To Range("MaxBits")
'        sSum = 0
'        x = 0
'        Do While x <= UBound(strCol)
'            sSum = 0
'            j = 0
'            For i = 1 To Range("Bit" & iBit & "_Case" & iCase & "_" & strCol(x)).Rows.Count
'                If Range("Bit" & iBit & "_Case" & iCase & "_" & strCol(x)).Rows(i).Hidden = False Then
'                    sSum = sSum + Range("Bit" & iBit & "_Case" & iCase & "_" & strCol(x)).Rows(i)
'                    j = j + 1
'                End If
'            Next i
'
'            Range("Bit" & iBit & "_Sum_" & strCol(x)) = sSum 'Format(sSum, "0.0")
'            Range("Bit" & iBit & "_Avg_" & strCol(x)) = sSum / j
'            'Range("Bit" & iBit & "_Avg_" & strCol(x)) = Application.WorksheetFunction.Average("Bit" & iBit & "_Case" & iCase & "_" & strCol(x))
'            x = x + 1
'        Loop
'
'    Next iBit
    
    
    Exit Sub
    
ErrHandler:

    If Err = 424 Then
        Exit Sub
    Else
        Resume Next
    End If
    
End Sub

Public Sub subZoneBoxChangeUpdateListBox(Optional iCase1 As Integer)
    Dim iBit As Integer
    Dim sSum As Single, i As Integer, j As Integer, X As Integer, strCol() As Variant
    Dim strForces As String
    Dim sAvg As Single
    
    On Error GoTo ErrHandler
    
    strCol() = Array("ZNormal", "ZNormal_Workrate", "ZCircumferential", "ZCircumferential_Workrate", "ZVertical", "ZSide", "ZWearFlatArea", "ZAreaCut", "ZVolCut", "ZProE_fx", "ZProE_fy", "ZProE_fz", "ZCutter_Torque", "ZResultant", "ZResultant_Workrate", "ZDelam", "ZDelam_Workrate")
    X = 0
    
    For iBit = 1 To Range("MaxBits")
        If Not BitIsActive(iBit) Then GoTo NextBit
        sSum = 0
        X = 0
        Do While X <= UBound(strCol)
            sSum = 0
            j = 0
            For i = 1 To Range("Bit" & iBit & "_Case" & iCase1 & "_" & strCol(X)).Rows.Count
                If Range("Bit" & iBit & "_Case" & iCase1 & "_" & strCol(X)).Rows(i).Hidden = False Then
                    sSum = sSum + Range("Bit" & iBit & "_Case" & iCase1 & "_" & strCol(X)).Rows(i)
                    j = j + 1
                End If
            Next i
            
            Range("Bit" & iBit & "_Sum_" & strCol(X)) = sSum 'Format(sSum, "0.0")
            Range("Bit" & iBit & "_Avg_" & strCol(X)) = sSum / j
            'Range("Bit" & iBit & "_Avg_" & strCol(x)) = Application.WorksheetFunction.Average("Bit" & iBit & "_Case" & iCase & "_" & strCol(x))
            X = X + 1
        Loop
NextBit:
    Next iBit

    Exit Sub
    
ErrHandler:

    Stop
    Resume
    
End Sub

Sub Reset_UsedRange()

    gProc = "Reset_UsedRange"
    
    Range("A1").Select
    
    Cells.Select
    Cells.ClearContents
    Cells.ClearFormats

    Range("A1").Select
    
'    Debug.Print "Sheet: " & ActiveSheet.Name & "    UsedRange: " & ActiveSheet.UsedRange.Address
    
End Sub

Sub subZoIColor(strHideShow As String)

    If strHideShow = "Hide" Then
        With Range("ZoIInputValues1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12632256
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Range("ZoIInputValues2").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12632256
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Range("ZoIInput1").Font
            .Color = -8355712
            .TintAndShade = 0
        End With
        With Range("ZoIInputValues2").Font
            .Color = -8355712
            .TintAndShade = 0
        End With
    Else
        With Range("ZoIInputValues1").Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Range("ZoIInputValues2").Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Range("ZoIInput1").Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
        With Range("ZoIInputValues2").Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
    End If
    
End Sub

Public Sub ZoI(strShowHideCalc As String)
    Dim bit_num As Integer, hide_flag As Boolean, hide_backups_flag As Boolean, hide_offprofile_flag As Boolean
    Dim legend_name As String, A As String, X As Integer, legend_count As Integer
    Dim chart_names() As Variant
    Dim strSheet As String


    If gblnZoIShtOverride Then
        Exit Sub
    ElseIf ActiveBits < 1 Then
'        ActiveSheet.OptionButtons("Option_Button_Hide_Zone").Value = Checked
'        ActiveSheet.Shapes("Option_Button_Hide_Zone").Value = xlOn
        Sheets("HOME").Shapes("Option_Button_Hide_Zone").ControlFormat.Value = xlOn
        Exit Sub

    End If
    
    subUpdate False
    
    strSheet = ActiveSheet.Name

    If strShowHideCalc = "Hide" Then
        'ZoI Hide
        
        subZoIColor "Hide"
        Range("ZoIShowHide") = "Hide"
        
        Range("Limits").Columns.EntireColumn.Hidden = True
        Range("Zone").Rows.EntireRow.Hidden = True
        
        subZoIOptionSet strShowHideCalc

    ElseIf strShowHideCalc = "Show" Or strShowHideCalc = "Calc" Then
        'ZoI Show
        
        If Range("XAxis") <> "Radius" Then SetPlotsXAxis "Radius"
        
        subZoIOptionSet "Show"
        subZoIColor "Show"
        Range("ZoIShowHide") = "Show"
        
        Range("Limits").Columns.EntireColumn.Hidden = False
        
        Range("Zone").Rows.EntireRow.Hidden = False
        For bit_num = 1 To Application.Range("MaxBits")
            If Sheets("HOME").Shapes("Show_Bit" & bit_num).ControlFormat.Value = xlOn Then
                A = HideCutters(bit_num, False, False, False)
            ElseIf Sheets("HOME").Shapes("Hide_Bit" & bit_num).ControlFormat.Value = xlOn Then
                    A = HideCutters(bit_num, False, True, False)
                ElseIf Sheets("HOME").Shapes("Hide_Backups_Bit" & bit_num).ControlFormat.Value = xlOn Then
                        A = HideCutters(bit_num, True, True, False)
                    Else: A = HideCutters(bit_num, True, True, True)
                    'End If
                'End If
            End If
        Next bit_num
        
        Sheets("PROFILE").Select
        chart_names() = Array("Profile2_Chart", "Cutter_Dist_Chart", "BackRake_Chart", "SideRake_Chart", "Cutter_Exposure_Chart")
        X = 0
        Do While X <= UBound(chart_names)
            ActiveSheet.ChartObjects(chart_names(X)).Activate
            ActiveChart.Legend.Select
            legend_count = ActiveChart.Legend.LegendEntries.Count
            legend_name = ActiveChart.SeriesCollection(legend_count).Name
            If legend_name = "Inner Limit" Then
                ActiveChart.Legend.LegendEntries(legend_count).Select
                Selection.Delete
            End If
            legend_count = ActiveChart.Legend.LegendEntries.Count
            legend_name = ActiveChart.SeriesCollection(legend_count).Name
            If legend_name = "Outer Limit" Then
                ActiveChart.Legend.LegendEntries(legend_count).Select
                Selection.Delete
            End If
            X = X + 1
        Loop
        Erase chart_names
        ActiveSheet.Range("A1").Select
            
        Sheets("WEAR").Select
        chart_names() = Array("NFwr_Chart", "WNFwr_Chart", "CFwr_Chart", "WCFwr_Chart", "VFwr_Chart", "WFAwr_Chart", "CAwr_Chart", "CVwr_Chart")
        X = 0
        Do While X <= UBound(chart_names)
            ActiveSheet.ChartObjects(chart_names(X)).Activate
            ActiveChart.Legend.Select
            legend_count = ActiveChart.Legend.LegendEntries.Count
            legend_name = ActiveChart.SeriesCollection(legend_count).Name
            If legend_name = "Outer Limit" Then
                ActiveChart.Legend.LegendEntries(legend_count).Select
                Selection.Delete
            End If
            legend_count = ActiveChart.Legend.LegendEntries.Count
            legend_name = ActiveChart.SeriesCollection(legend_count).Name
            If legend_name = "Inner Limit" Then
                ActiveChart.Legend.LegendEntries(legend_count).Select
                Selection.Delete
            End If
            X = X + 1
        Loop
        Erase chart_names
        ActiveSheet.Range("A1").Select
        
        Sheets("FORCES").Select
        chart_names() = Array("NF1_Chart", "NF2_Chart", "NF3_Chart", "NF4_Chart", "NF5_Chart", "WNF1_Chart", "WNF2_Chart", "WNF3_Chart", "WNF4_Chart", "WNF5_Chart", "VF1_Chart", "VF2_Chart", "VF3_Chart", "VF4_Chart", "VF5_Chart", "CF1_Chart", "CF2_Chart", "CF3_Chart", "CF4_Chart", "CF5_Chart", "WCF1_Chart", "WCF2_Chart", "WCF3_Chart", "WCF4_Chart", "WCF5_Chart", "WFA1_Chart", "WFA2_Chart", "WFA3_Chart", "WFA4_Chart", "WFA5_Chart", "CA1_Chart", "CA2_Chart", "CA3_Chart", "CA4_Chart", "CA5_Chart", "CV1_Chart", "CV2_Chart", "CV3_Chart", "CV4_Chart", "CV5_Chart", "SF1_Chart", "SF2_Chart", "SF3_Chart", "SF4_Chart", "SF5_Chart", "TQB1_Chart", "TQB2_Chart", "TQB3_Chart", "TQB4_Chart", "TQB5_Chart")
        X = 0
        Do While X <= UBound(chart_names)
            ActiveSheet.ChartObjects(chart_names(X)).Activate
            ActiveChart.Legend.Select
            legend_count = ActiveChart.Legend.LegendEntries.Count
            legend_name = ActiveChart.SeriesCollection(legend_count).Name
            If legend_name = "Outer Limit" Then
                ActiveChart.Legend.LegendEntries(legend_count).Select
                Selection.Delete
            End If
            legend_count = ActiveChart.Legend.LegendEntries.Count
            legend_name = ActiveChart.SeriesCollection(legend_count).Name
            If legend_name = "Inner Limit" Then
                ActiveChart.Legend.LegendEntries(legend_count).Select
                Selection.Delete
            End If
            X = X + 1
        Loop
        Erase chart_names
        ActiveSheet.Range("A1").Select
        
        Sheets("CUTTER_FORCES").Select
        chart_names() = Array("FX1_Chart", "FX2_Chart", "FX3_Chart", "FX4_Chart", "FX5_Chart", "FY1_Chart", "FY2_Chart", "FY3_Chart", "FY4_Chart", "FY5_Chart", "FZ1_Chart", "FZ2_Chart", "FZ3_Chart", "FZ4_Chart", "FZ5_Chart", "CT1_Chart", "CT2_Chart", "CT3_Chart", "CT4_Chart", "CT5_Chart", "RF1_Chart", "RF2_Chart", "RF3_Chart", "RF4_Chart", "RF5_Chart", "DL1_Chart", "DL2_Chart", "DL3_Chart", "DL4_Chart", "DL5_Chart", "WRF1_Chart", "WRF2_Chart", "WRF3_Chart", "WRF4_Chart", "WRF5_Chart", "WDL1_Chart", "WDL2_Chart", "WDL3_Chart", "WDL4_Chart", "WDL5_Chart")
        X = 0
        Do While X <= UBound(chart_names)
            ActiveSheet.ChartObjects(chart_names(X)).Activate
            ActiveChart.Legend.Select
            legend_count = ActiveChart.Legend.LegendEntries.Count
            legend_name = ActiveChart.SeriesCollection(legend_count).Name
            If legend_name = "Outer Limit" Then
                ActiveChart.Legend.LegendEntries(legend_count).Select
                Selection.Delete
            End If
            legend_count = ActiveChart.Legend.LegendEntries.Count
            legend_name = ActiveChart.SeriesCollection(legend_count).Name
            If legend_name = "Inner Limit" Then
                ActiveChart.Legend.LegendEntries(legend_count).Select
                Selection.Delete
            End If
            X = X + 1
        Loop
        Erase chart_names
        ActiveSheet.Range("A1").Select

        subZoIOptionSet strShowHideCalc
        
'    ElseIf strShowHideCalc = "Calc" Then
        'Perform ZoI Calculations

'        If Range("ZoIShowHide") = "Hide" Then ZoI "Show"
        
        CalculateZoneForces
        
    End If
    
    subZoIRefreshMsg strShowHideCalc
    
    Sheets(strSheet).Select
    
    subUpdate True


End Sub

Public Sub ZoIRangeChange(strIO As String, strValue As String, iWhichZ As Integer)
    Dim i As Integer
    Dim iZ As Integer
    
    On Error GoTo ErrHandler
    
        
    Range("Z" & strIO) = Val(strValue)
    
    
    ZoI "Show"
    gblnZoIShtOverride = False
    
    i = 0
    
    Do
        i = i + 1
        Debug.Print "Z" & strIO & i & " Address: " & Range("Z" & strIO & i).Address
        iZ = i
    Loop
    
ResumeSub:
    gblnZoIOverride = True
    For i = 1 To iZ
        If i <> iWhichZ Then Range("Z" & strIO & i) = Range("Z" & strIO)
    Next i
    gblnZoIOverride = False
    
    Exit Sub
    
ErrHandler:

    If Err = 1004 Then
        GoTo ResumeSub
    End If
    
End Sub

Public Sub subXAxisOptionSet(strFormat As String)

    On Error GoTo ErrHandler

    For Each sh In ThisWorkbook.Sheets
        sh.Shapes("optXAxis" & strFormat).ControlFormat.Value = xlOn
    Next sh
    
    Exit Sub
    
ErrHandler:
   
    If Err = 1004 Then
'        Debug.Print "Sheet """ & sh.Name & """ does not contain Option_Button_" & strShowHide & "_Zone!"
        Resume Next
    ElseIf Err = -2147024809 Then
        Resume Next
    Else
        Debug.Print Err, Err.Description
        Stop
        Resume
    End If
    
End Sub
Public Sub subZoIOptionSet(strShowHide As String)
    Dim sh As Object
    
    On Error GoTo ErrHandler

    For Each sh In ThisWorkbook.Sheets
        'Option_Button_Hide_Zone
'        sh.OptionButtons("Option_Button_" & strShowHide & "_Zone").Value = Checked
'        Debug.Print sh.name
        sh.Shapes("Option_Button_" & strShowHide & "_Zone").ControlFormat.Value = xlOn
    Next sh
    
    Exit Sub
    
ErrHandler:
   
    If Err = 1004 Then
'        Debug.Print "Sheet """ & sh.Name & """ does not contain Option_Button_" & strShowHide & "_Zone!"
        Resume Next
    ElseIf Err = -2147024809 Then
        Resume Next
    Else
        Debug.Print Err, Err.Description
        Stop
        Resume
    End If
    
End Sub

Public Sub subZoIRefreshMsg(strShowHide As String)

Exit Sub
    With Range("ZoIRefreshMsg")
        If strShowHide = "Hide" Then
            With .Font
                .Name = "Calibri"
                .FontStyle = "Bold"
                .Size = 16
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.499984741
                .ThemeFont = xlThemeFontMinor
            End With
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = -16776961
                .TintAndShade = 0
                .Weight = xlThick
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = -16776961
                .TintAndShade = 0
                .Weight = xlThick
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = -16776961
                .TintAndShade = 0
                .Weight = xlThick
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = -16776961
                .TintAndShade = 0
                .Weight = xlThick
            End With
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 52479
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            .Value = "CLICK CALCULATE BUTTON TO REFRESH PAGE!"
        ElseIf strShowHide = "Show" Then
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            With .Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            .Value = ""
        End If
    End With
    

End Sub

Public Function AddProfileLocation(iBitNum As Integer, iCutterCount As Integer) As String
    Dim iCol As Integer, iEndRow As Integer

    iCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column    'fXLLetterToNumber(Mid(ActiveSheet.UsedRange.Address, InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2, InStrRev(ActiveSheet.UsedRange.Address, "$") - (InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2)))
    AddProfileLocation = fXLNumberToLetter(CLng(iCol))
    
    iEndRow = CInt(Right(ActiveSheet.UsedRange.Address, Len(ActiveSheet.UsedRange.Address) - InStrRev(ActiveSheet.UsedRange.Address, "$")))
    
    Cells(4, iCol) = "ProfileLocation"
    Cells(5, iCol).Select
    
    Range(Cells(5, iCol), Cells(iEndRow, iCol)).Name = "Bit" & iBitNum & "_ProfileLocation"
    
    Range("A1").Select
    
End Function

Public Function AddIdentity(iBitNum As Integer, iCutterCount As Integer) As String
    Dim iCol As Integer, i As Integer, iCount As Integer, iTag As Integer, iBlade As Integer, X As Integer
    Dim strLet As String, iColLet As Integer
    Dim iColTag As Integer
    Dim strLetIdent As String
    Dim iEndRow As Integer

    iCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column    'fXLLetterToNumber(Mid(ActiveSheet.UsedRange.Address, InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2, InStrRev(ActiveSheet.UsedRange.Address, "$") - (InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2)))
    AddIdentity = fXLNumberToLetter(CLng(iCol))
    
    iColTag = Range("Bit" & iBitNum & "_Tag").Column
    iColLet = Range("Bit" & iBitNum & "_Blade_Num").Column
    strLet = fXLNumberToLetter(CLng(iColLet))
    iEndRow = CInt(Right(ActiveSheet.UsedRange.Address, Len(ActiveSheet.UsedRange.Address) - InStrRev(ActiveSheet.UsedRange.Address, "$")))
        
    Cells(4, iCol) = "Identity"
    Cells(5, iCol).Select
    
    
    For i = 5 To iEndRow
        iCount = 0
        iBlade = Cells(i, iColLet)
        iTag = Cells(i, iColTag)
        For X = 5 To i
            If Cells(X, iColLet) = iBlade Then
                If Cells(X, iColTag) = iTag Then
                    iCount = iCount + 1
                End If
            End If
        Next X
    
        If iTag = 0 Then
            Cells(i, iCol) = "B" & iBlade & "C" & iCount
        ElseIf iTag = 1 Or iTag = 3 Then
            Cells(i, iCol) = "B" & iBlade & "BU" & iCount
        Else
            Cells(i, iCol) = "B" & iBlade & "SC" & iCount
        End If
    
    Next i
    
    Range(Cells(5, iCol), Cells(iEndRow, iCol)).Name = "Bit" & iBitNum & "_Identity"
    
    Range("A1").Select
    
End Function

Public Function AddRadiusRatio(iBitNum As Integer, iCutterCount As Integer) As String
    Dim iCol As Integer, i As Integer, iColRad As Integer
    Dim iEndRow As Integer
    Dim sngMaxRadius As Single

    iCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column    'fXLLetterToNumber(Mid(ActiveSheet.UsedRange.Address, InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2, InStrRev(ActiveSheet.UsedRange.Address, "$") - (InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2)))
    AddRadiusRatio = fXLNumberToLetter(CLng(iCol))
    
    iColRad = Range("Bit" & iBitNum & "_Radius").Column
    iEndRow = CInt(Right(ActiveSheet.UsedRange.Address, Len(ActiveSheet.UsedRange.Address) - InStrRev(ActiveSheet.UsedRange.Address, "$")))
    
    sngMaxRadius = Application.WorksheetFunction.Max(Range("Bit" & iBitNum & "_Radius"))
    
    Cells(4, iCol) = "Radius_Ratio"
    Cells(5, iCol).Select
    
    For i = 5 To iEndRow
        If sngMaxRadius = 0 Then
            Cells(i, iCol) = 0
        Else
            Cells(i, iCol) = Cells(i, iColRad) / sngMaxRadius
        End If
    Next i
    
    Range(Cells(5, iCol), Cells(iEndRow, iCol)).Name = "Bit" & iBitNum & "_Radius_Ratio"
    
    Range("A1").Select
    
End Function

Public Function AddXAxis(iBitNum As Integer, iCutterCount As Integer) As String
    Dim iCol As Integer, i As Integer, iColRad As Integer
    Dim iEndRow As Integer
    Dim sngMaxRadius As Single

    iCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column
    iEndRow = CInt(Right(ActiveSheet.UsedRange.Address, Len(ActiveSheet.UsedRange.Address) - InStrRev(ActiveSheet.UsedRange.Address, "$")))

    Cells(4, iCol) = "XAxis"
    Cells(5, iCol).Select
    
    For i = 5 To iEndRow
        Cells(i, iCol) = Range("Bit" & iBitNum & "_" & Range("XAxis")).Cells(i - 4)
    Next i
    
    Range(Cells(5, iCol), Cells(iEndRow, iCol)).Name = "Bit" & iBitNum & "_XAxis"
    
    Range("A1").Select

End Function

Public Function AddCutterRoll(iBitNum As Integer, iCutterCount As Integer) As String
    Dim iCol As Integer, i As Integer, iColRad As Integer
    Dim iEndRow As Integer
    Dim sngMaxRadius As Single

    iCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column
    iEndRow = CInt(Right(ActiveSheet.UsedRange.Address, Len(ActiveSheet.UsedRange.Address) - InStrRev(ActiveSheet.UsedRange.Address, "$")))

    Cells(4, iCol) = "Roll"
    Cells(5, iCol).Select
    
    For i = 5 To iEndRow
        Cells(i, iCol) = Range("Bit" & iBitNum & "_" & Range("XAxis")).Cells(i - 4)
    Next i
    
    Range(Cells(5, iCol), Cells(iEndRow, iCol)).Name = "Bit" & iBitNum & "_CutterRoll"
    
    Range("A1").Select

End Function

Public Function AddProfZone(iBitNum As Integer, iCutterCount As Integer) As String
    Dim iCol As Integer, i As Integer, iColRad As Integer
    Dim iEndRow As Integer
    Dim sngMaxRadius As Single

    iCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column
    iEndRow = CInt(Right(ActiveSheet.UsedRange.Address, Len(ActiveSheet.UsedRange.Address) - InStrRev(ActiveSheet.UsedRange.Address, "$")))

    Cells(4, iCol) = "ProfZone"
    Cells(5, iCol).Select
        
    Range(Cells(5, iCol), Cells(iEndRow, iCol)).Name = "Bit" & iBitNum & "_ProfZone"
    
    Range("A1").Select

End Function


Public Function GetIntegral(strRange As String, iBit As Integer, iCase As Integer) As Single
    Dim iCutterCount As Integer
    Dim sngInteg As Single, sngStep As Single
    Dim i As Integer
    Dim sngRadA As Single, sngRadB As Single, sngDataA As Single, sngDataB As Single
    
    iCutterCount = Range("Bit" & iBit & "_Case" & iCase & "_" & strRange).Count
    
    For i = 2 To iCutterCount
        sngRadA = Range("Bit" & iBit & "_Radius").Rows(i - 1)
        sngRadB = Range("Bit" & iBit & "_Radius").Rows(i)
        sngDataA = Range("Bit" & iBit & "_Case" & iCase & "_" & strRange).Rows(i - 1)
        sngDataB = Range("Bit" & iBit & "_Case" & iCase & "_" & strRange).Rows(i)
        
        If sngDataA = sngDataB Then
            sngStep = Abs(sngRadB - sngRadA) * sngDataA
        ElseIf sngDataA < sngDataB Then
            sngStep = Abs(sngRadB - sngRadA) * (sngDataA + (sngDataB - sngDataA) * 0.5)
        Else
            sngStep = Abs(sngRadB - sngRadA) * (sngDataB + (sngDataA - sngDataB) * 0.5)
        End If
        
        sngInteg = sngInteg + sngStep
    Next i

    GetIntegral = sngInteg
    
    Debug.Print "Bit" & iBit & "_Case" & iCase & "_" & strRange, Format(sngInteg, "0.000")
    
End Function


Public Sub ClearRanges(iBit As Integer)
    
    gProc = "ClearRanges"
    
    Dim iCase As Integer
    Dim iRange As Integer
    
    Exit Sub
    
    On Error GoTo ErrHandler

'  This subroutine is in the works still. Saved for later development. -A. Barr

    iCase = 10

    Do While iCase > 0  'Just keep going until you error out!
        For iRange = 1 To Range("CaseRanges").Count
            ActiveWorkbook.Names("Bit" & iBit & "_Case" & iCase & Range("CaseRanges").Rows(iRange)).Delete
        Next iRange
        iCase = iCase + 1
    Loop

    Exit Sub
    
ErrHandler:
If Err = 1004 Then Exit Sub
    Stop
    Debug.Print "Bit" & iBit & "_Case" & iCase & Range("CaseRanges").Rows(iRange), Err, Err.Description
    Resume
    

End Sub

Sub SetCaseTo1()
    Dim i As Integer

    For i = 2 To Range("Cases").Count
        Range("Cases").Rows(i).ClearContents
    Next i
    
    Range("ConstCase").Item(1).Name = "Cases"
    
'    Range("ConstCase").Select
'    Selection.name = "Cases"
    
End Sub

Sub RFPlotXAxis()
    Dim i As Integer, strSetToFormat As String
    Dim strXY As String
    
    strSetToFormat = Right(Application.Caller, Len(Application.Caller) - Len("optRFXAxis"))

    With Sheets("HOME").ChartObjects("RF_Chart").Chart
        For i = 1 To Range("MaxBits")
            If Sheets("HOME").Shapes("Hide_Bit" & i).ControlFormat.Value <> xlOn Then iBits = iBits + 1
        Next i
        
        
        iSeries = .SeriesCollection.Count

        For i = 1 To iSeries
            strLine = .SeriesCollection(i).Formula
            
            If strSetToFormat = "Radius" Then
                strLine = Replace(strLine, "Cutter_Num", "Radius")
            ElseIf strSetToFormat = "Cutter_Num" Then
                strLine = Replace(strLine, "Radius", "Cutter_Num")
            End If
'            If InStr(1, strLine, "Cutter_Num") > 0 Then
'                strLine = Replace(strLine, "Cutter_Num", "Radius")
'            ElseIf InStr(1, strLine, "Radius") > 0 Then
'                strLine = Replace(strLine, "Radius", "Cutter_Num")
'            End If
            
            .SeriesCollection(i).Formula = strLine

'            .Format.TextFrame2.TextRange.Characters.Text = strSetToFormat

            
        Next i
    
        If strSetToFormat = "Radius" Then
            strSetToFormat = "Radius (in)"
        ElseIf strSetToFormat = "Cutter_Num" Then
            strSetToFormat = "Cutter Number"
        End If

        ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = strSetToFormat
    End With



End Sub

Sub SetPlotsXAxis(Optional strFormat As String)
    Dim strSheet As String, strSheet2 As String
    Dim i As Integer, X As Integer
    
    On Error GoTo ErrHandler
    
    strSheet = ActiveSheet.Name
    
    If strFormat = "" Then
        Range("XAxis") = Right(Application.Caller, Len(Application.Caller) - Len("optXAxis"))
    ElseIf strFormat = "Radius" Or strFormat = "Cutter_Num" Or strFormat = "Radius_Ratio" Then
        Range("XAxis") = strFormat
    Else
        MsgBox "X-Axis format not appropriately set." & vbCrLf & vbCrLf & """" & strFormat & """ is not a recognizable format."
        Exit Sub
    End If
    
    Select Case Range("XAxis")
        Case "Radius":
            Range("XAxisText") = "Radius (in)"
            Sheets("HOME").Shapes("optXAxisRadius").ControlFormat.Value = xlOn
        Case "Cutter_Num":
            Range("XAxisText") = "Cutter Number"
            Sheets("HOME").Shapes("optXAxisCutter_Num").ControlFormat.Value = xlOn
            If IsZoI Then MsgBox ("Zone of Interest may not work correctly in this mode.")
        Case "Radius_Ratio":
            Range("XAxisText") = "Radius Ratio"
            Sheets("HOME").Shapes("optXAxisRadius_Ratio").ControlFormat.Value = xlOn
            If IsZoI Then MsgBox ("Zone of Interest may not work correctly in this mode.")
    End Select
    
    subUpdate False
    
    
    For i = 1 To Range("MaxBits")
        Sheets("BIT" & i).Activate
        
        For X = 1 To Range("Bit" & i & "_Cutter_Num").Count
            Range("Bit" & i & "_XAxis").Cells(X) = Range("Bit" & i & "_" & Range("XAxis")).Cells(X)
        Next X
    Next i
    
    
    For i = 1 To Range("PlotIDs").Count
        strSheet2 = Range("PlotSheets").Cells(i)
        Sheets(strSheet2).Activate
        
        If strSheet2 = "FORCES" Or strSheet2 = "CUTTER_FORCES" Then
            For X = 1 To 5
                Sheets(strSheet2).ChartObjects(Range("PlotIDs").Cells(i) & X & "_Chart").Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = Range("XAxisText")
            Next X
        Else
'            If strSheet2 = "WEAR" Then
'
'                Dim strtemp As String
'
'                For x = 1 To 5
'                strtemp = Sheets(strSheet2).ChartObjects(Range("PlotIDs").Cells(i) & "_Chart").Chart.SeriesCollection(x).Formula
'
'                strtemp = Replace(strtemp, "Radius", "XAxis")
'                Sheets(strSheet2).ChartObjects(Range("PlotIDs").Cells(i) & "_Chart").Chart.SeriesCollection(x).Formula = strtemp
'                Next x
'
'            End If
            
        
            Sheets(strSheet2).ChartObjects(Range("PlotIDs").Cells(i) & "_Chart").Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = Range("XAxisText")
        End If
    Next i
    
    subXAxisOptionSet Range("XAxis")
    
    Sheets(strSheet).Activate
    
    subUpdate True

    Exit Sub
    
ErrHandler:

Stop
Debug.Print Err, Err.Description

Resume


End Sub

'Function GetOldFormat(strFormula As String) As String
'    Dim i As Integer, x(2) As Integer
'
'    x(0) = InStr(1, strFormula, "!")
'    x(1) = InStr(x(0) + 1, strFormula, "!")
'    x(2) = InStr(x(1) + 1, strFormula, ",")
'    GetOldFormat = Mid(strFormula, x(1) + 1, x(2) - x(1) - 1)
'    GetOldFormat = Right(GetOldFormat, Len(GetOldFormat) - InStr(1, GetOldFormat, "_"))
'
'End Function
'
'Sub PlotsXAxis()
'    Dim i As Integer, strSetToFormat As String
'    Dim strXY As String
'
'    strSetToFormat = Right(Application.Caller, Len(Application.Caller) - Len("optXAxis"))
'
'    subUpdate False
'
'    With Sheets("HOME").ChartObjects("RF_Chart").Chart
'        For i = 1 To Range("MaxBits")
'            If Sheets("HOME").Shapes("Hide_Bit" & i).ControlFormat.Value <> xlOn Then iBits = iBits + 1
'
'        Next i
'
'
'        iSeries = .SeriesCollection.Count
'
'        For i = 1 To iSeries
'            strLine = .SeriesCollection(i).Formula
'
'            If strSetToFormat = "Radius" Then
'                strLine = Replace(strLine, "Cutter_Num", "Radius")
'            ElseIf strSetToFormat = "Cutter_Num" Then
'                strLine = Replace(strLine, "Radius", "Cutter_Num")
'            End If
'
'            .SeriesCollection(i).Formula = strLine
'
'        Next i
'
'        If strSetToFormat = "Radius" Then
'            strSetToFormat = "Radius (in)"
'        ElseIf strSetToFormat = "Cutter_Num" Then
'            strSetToFormat = "Cutter Number"
'        ElseIf strSetToFormat = "Ratio" Then
'            strSetToFormat = "Radius Ratio"
'        End If
'
'        ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = strSetToFormat
'    End With
'
'    subUpdate True
'
'End Sub

Public Sub subUpdate(blnTrueOrFalse As Boolean)

    If blnTrueOrFalse Then
        Application.ScreenUpdating = True
'        Application.Calculation = xlAutomatic
    Else
        Application.ScreenUpdating = False
'        Application.Calculation = xlManual
    End If
    
End Sub

Public Function IsStinger(strCutterType As String) As Boolean

    gProc = "IsStinger"
    
    Select Case strCutterType
        Case "565_790_85_90": IsStinger = True
        Case "565_690_85_90": IsStinger = True
        Case "566_790_85_90": IsStinger = True
        Case "566_690_85_90": IsStinger = True
        Case "440_540_85_90": IsStinger = True
        Case "440_537_85_90": IsStinger = True
        Case "625_826_96_94": IsStinger = True
        Case "562_7880_S": IsStinger = True
        Case "562_8590_S": IsStinger = True
        Case "562_7880_L": IsStinger = True
        Case "562_8590_L": IsStinger = True
        Case "562_9694_L": IsStinger = True
        Case "B565_690_441_900_80": IsStinger = True
        Case "B565_790_541_900_80": IsStinger = True
        Case "New_8590_L": IsStinger = True

        Case Else: IsStinger = False
    End Select

End Function

Public Sub TEST1()
    Dim strFile As String
    
    strFile = "D:\Documents\Linux_Xfers\064_MSi413_70530\Static\New_Design_Itrs\Itr001\summary_regular"
    
A = OpenText(strFile)
'Workbooks.OpenText FileName:=strFile, Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=True, Other:=True, TrailingMinusNumbers:=True

End Sub

Public Function OpenText(strDir As String, Optional strMethod As String, Optional strExtra As String) As String
    
    On Error GoTo ErrHandler
    
    gProc = "OpenText"
    
'The entire purpose of this function is to work around the random "Automation Error" crash that is seen on Excel 2013.
'I feel the reason this is occuring is sometimes in the native Excel.OpenText function, the program can get away from
'itself and it runs wild. I see this when stepping line to line. At the native function, it sometimes completes as if
'I pressed F5 instead of F8.
'This function will create the same output the native function outputs, but uses a FileSystemObject and array instead
'of the native Excel command.

If strMethod = "array" Then

    Dim fs As scripting.FileSystemObject
    Dim txtIn As Object ' scripting.textstream
    Dim strLine As String, strArray() As String
    Dim iRow As Integer, iCol As Integer
    
    iRow = 1
    
    ActiveWorkbook.Sheets.Add
    
    Set fs = New scripting.FileSystemObject
    Set txtIn = fs.OpenTextFile(strDir, 1) ' 1 ForReading
    Do While Not txtIn.AtEndOfStream
        strLine = txtIn.ReadLine
    
    'The original import delimited on space and tab. I'm changing these to @s and will split line into an array at the @s. This may be an issue if the user uses an @ sign in their bit name or path.
        strLine = Replace(strLine, " ", "@")
        strLine = Replace(strLine, vbTab, "@")
        
    'Make sure there are no extra spaces between line data
        Do While InStr(1, strLine, "@@") > 0
            strLine = Replace(strLine, "@@", "@")
        Loop
            
    'Split into an array that will be entered into the blank sheet.
        strArray = Split(strLine, "@")

        For iCol = 0 To UBound(strArray)
            Cells(iRow, iCol + 1) = strArray(iCol)
        Next iCol
        
        iRow = iRow + 1
    Loop
    
    OpenText = ActiveSheet.Name
    
    txtIn.Close
    Set txtIn = Nothing
    Set fs = Nothing
    
    
    
    
ElseIf strMethod = "query" Then
'Taken from http://stackoverflow.com/questions/11267459/vba-importing-text-file-into-excel-sheet on 2016-02-24

    ActiveWorkbook.Sheets.Add

    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & strDir, Destination:=Range("$A$1"))
        .Name = "Sample"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        If strExtra = "cfb" Then
            .TextFileColumnDataTypes = Array(xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlTextFormat, xlGeneralFormat)
        Else
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1)
        End If
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    
    ClearQueryTables
    
    
    OpenText = ActiveSheet.Name
    
End If
    
Exit Function

ErrHandler:
Debug.Print Now(), Err, Err.Description

Stop

Resume

End Function

Sub ClearQueryTables()

    gProc = "ClearQueryTables"
    
' taken from https://www.safaribooksonline.com/library/view/programming-excel-with/0596007663/re661.html
    Dim qt As QueryTable
    For Each qt In ActiveSheet.QueryTables
        If qt.Refreshing Then qt.CancelRefresh
        qt.Delete
    Next

'RemoveConnections()
'taken from https://social.technet.microsoft.com/Forums/office/en-US/a8ff7d1f-3278-4f0d-a080-b30d9a1cc065/vba-to-remove-all-data-connections-excel-2013?forum=excel on 2016-03-05
    Dim i As Long
    
    Do While ActiveWorkbook.Connections.Count > 0
        ActiveWorkbook.Connections.Item(ActiveWorkbook.Connections.Count).Delete
    Loop
    
    Exit Sub
    
    For i = ActiveWorkbook.Connections.Count To 1 Step -1
        Debug.Print "Removing Connection #" & i

        ActiveWorkbook.Connections.Item(i).Delete
    Next i
    
End Sub

Sub subActive(blnRunning As Boolean)

    gProc = "subActive"
    
    
    If blnRunning Then
        Range("Active") = "Running..."
    Else
        Range("Active") = "Ready"
    End If
    
End Sub

Sub Sanitize_Plot(strPlot As String, blnSanitize As Boolean)
    On Error GoTo ErrHandler
'    strPlot = "BackRake_Chart"
    
    ActiveSheet.ChartObjects(strPlot).Activate
    
    If blnSanitize Then
        ActiveSheet.ChartObjects(strPlot).Chart.Axes(xlCategory, xlPrimary).TickLabels.Font.Color = vbWhite
        ActiveChart.Axes(xlCategory).MajorTickMark = xlNone
        ActiveChart.Axes(xlCategory).MajorGridlines.Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
        End With

        ActiveSheet.ChartObjects(strPlot).Chart.Axes(xlValue, xlPrimary).TickLabels.Font.Color = vbWhite
        ActiveChart.Axes(xlValue).MajorTickMark = xlNone
        ActiveChart.Axes(xlValue).MajorGridlines.Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
        End With
        
    Else
        ActiveSheet.ChartObjects(strPlot).Chart.Axes(xlCategory, xlPrimary).TickLabels.Font.Color = vbBlack
        ActiveChart.Axes(xlCategory).MajorTickMark = xlOutside
        ActiveChart.Axes(xlCategory).MajorGridlines.Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    
        ActiveSheet.ChartObjects(strPlot).Chart.Axes(xlValue, xlPrimary).TickLabels.Font.Color = vbBlack
        ActiveChart.Axes(xlValue).MajorTickMark = xlOutside
        ActiveChart.Axes(xlValue).MajorGridlines.Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    
    End If
    
    Exit Sub
    
ErrHandler:

Debug.Print Now(), gProc, Err, Err.Description
Stop
Resume

End Sub

Public Function CountInStr(strString As String, strCountWhat As String) As Integer
    Dim i As Integer
    
    i = 1
    
    Do While InStr(i, strString, strCountWhat) > 0
        i = InStr(i, strString, strCountWhat) + 1
        CountInStr = CountInStr + 1
    Loop

End Function

Public Sub CutterCount(strRegion As String, iBitNum As Integer, strType() As String, iCount() As Integer, iBR() As Integer, iSR() As Integer)

    gProc = "CutterCount"
    
    Dim i As Integer, iCutter As Integer, blnFound As Boolean, blnStarted As Boolean
    Dim iRow As Integer, iA As Integer, iB As Integer
    
    On Error GoTo ErrHandler
    
    Sheets("Bit" & iBitNum).Select
    
    blnFound = False
    blnStarted = False
    
    ReDim strType(0)
    ReDim iCount(2)
    ReDim iBR(1)
    ReDim iSR(1)
    
'define row starts and stops (iA, iB) for region
    If strRegion = "Cone" Or strRegion = "Gauge" Then
        iA = Range("Bit" & iBitNum & "_" & strRegion).Item(1).row
        iB = Range("Bit" & iBitNum & "_" & strRegion).Item(Range("Bit" & iBitNum & "_" & strRegion).Count).row
    Else
        If Range("Bit" & iBitNum & "_Nose").Count <> 0 Then
            iA = Range("Bit" & iBitNum & "_Nose").Item(1).row
        Else
            iA = Range("Bit" & iBitNum & "_Shoulder").Item(1).row
        End If
        iB = Range("Bit" & iBitNum & "_Shoulder").Item(Range("Bit" & iBitNum & "_Shoulder").Count).row
    End If
        
    
    For iRow = iA To iB
        For i = 0 To UBound(strType)
            If strType(i) = Cells(iRow, Range("Bit" & iBitNum & "_Cutter_Type").Column) Then
                blnFound = True
                
                iCount(0) = iCount(0) + 1  'Total Cutter Count
                If Cells(iRow, Range("Bit" & iBitNum & "_Tag").Column) = 0 Then
                    iCount(1) = iCount(1) + 1  'Primary Cutter Count
                Else
                    iCount(2) = iCount(2) + 1  'Backup Cutter Count
                End If
                
                If iBR(0) > Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column) Then iBR(0) = Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column)
                If iBR(1) < Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column) Then iBR(1) = Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column)
                If iSR(0) > Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column) Then iSR(0) = Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column)
                If iSR(1) < Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column) Then iSR(1) = Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column)
                
                Exit For
            End If
        Next i
        
        If blnFound = False Then
            If strType(i - 1) = "" Then
                strType(i - 1) = Cells(iRow, Range("Bit" & iBitNum & "_Cutter_Type").Column)
            Else
                strType(i) = Cells(iRow, Range("Bit" & iBitNum & "_Cutter_Type").Column)
            End If
            
            iCount(0) = iCount(0) + 1  'Total Cutter Count
            If Cells(iRow, Range("Bit" & iBitNum & "_Tag").Column) = 0 Then
                iCount(1) = iCount(1) + 1  'Primary Cutter Count
            Else
                iCount(2) = iCount(2) + 1  'Backup Cutter Count
            End If
            
            If Not blnStarted Then
                iBR(0) = Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column)
                iBR(1) = Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column)
                iSR(0) = Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column)
                iSR(1) = Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column)
            Else
                If iBR(0) > Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column) Then iBR(0) = Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column)
                If iBR(1) < Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column) Then iBR(1) = Cells(iRow, Range("Bit" & iBitNum & "_Back_Rake").Column)
                If iSR(0) > Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column) Then iSR(0) = Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column)
                If iSR(1) < Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column) Then iSR(1) = Cells(iRow, Range("Bit" & iBitNum & "_Side_Rake").Column)
            End If
            blnStarted = True
        End If
        
        blnFound = False
        
    Next iRow
    
    Exit Sub

ErrHandler:

    If Err = 9 Then
        ReDim Preserve strType(i)
        Resume
    End If

End Sub

Public Sub ExcelErrors(blnOnOff As Boolean)
    
    With Application.ErrorCheckingOptions
        .EmptyCellReferences = blnOnOff
        .InconsistentFormula = blnOnOff
        .InconsistentTableFormula = blnOnOff
        .ListDataValidation = blnOnOff
        .NumberAsText = blnOnOff
        .OmittedCells = blnOnOff
        .TextDate = blnOnOff
    End With

End Sub

Public Sub OpenFile(strPath As String, Optional strWithProgram As String)

    If strPath = "" Then Exit Sub
    
    If Dir(strPath, vbArchive) = "" And Dir(strPath, vbDirectory) = "" And Dir(strPath, vbHidden) = "" And Dir(strPath, vbNormal) = "" And Dir(strPath, vbReadOnly) = "" And Dir(strPath, vbSystem) = "" Then Exit Sub
    
    If strWithProgram <> "" Then
        Shell strWithProgram & " " & strPath, vbNormalFocus
    Else
        Shell "explorer.exe " & strPath, vbNormalFocus
    End If
    
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

'
'Public Sub GetFromWindchill()
'    Dim iBit As Integer
'    Dim strLink As String, strTempFile As String, strDir As String
'    Dim lFilesize As Long
'    Dim fs As scripting.FileSystemObject
'
'
'    iBit = Right(Application.Caller, 1)
'
'    If Range("Bit" & iBit & "_Name") = "" Then
'        MsgBox "Please enter the bit BOM number into the Bit Name field.", vbOKOnly + vbExclamation, "BOM# Needed"
'        Range("Bit" & iBit & "_Name").Select
'
'        Exit Sub
'    End If
'
'    subActive True
'
'
'
'    Set fs = New scripting.FileSystemObject
'
'    If Not fs.FolderExists(Range("DESSaveLoc")) Then
'        If Range("DESSaveLoc") = "Press ""Browse…"" to Select a" & Chr(10) & "Save Location" Or Range("DESSaveLoc") = "" Then
'            DESSaveLocation
'        ElseIf Not fs.DriveExists(Range("DESSaveLoc")) Then
'            DESSaveLocation
'        ElseIf Dir(Range("DESSaveLoc"), vbDirectory) = "" Then
'            DESSaveLocation
'        End If
'    End If
'
'    strDir = Range("DESSaveLoc")
'
'    If strDir = "" Then Exit Sub
'
'
'    If Right(strDir, 1) = "\" Then
'        strTempFile = strDir & LCase(Range("Bit" & iBit & "_Name"))
'    Else
'        strTempFile = strDir & "\" & LCase(Range("Bit" & iBit & "_Name"))
'    End If
'
'    strLink = "http://pdmp.corp.smith.com:8080/Windchill/infoengine/jsp/SmithBits/ProductView.jsp?NUMBER=" & LCase(Range("Bit" & iBit & "_Name"))
'
'    If Download_File(strLink & ".des", strTempFile & ".des", lFilesize) = False Then
'        Range("Bit3_Status") = "Read via Windchill Aborted by user."
'        Stop
'        subActive False
'
'        Exit Sub
'    End If
'
'    If Not IsDesignFile(strTempFile & ".des") Then
'        Kill strTempFile & ".des"
'
'        Download_File strLink & ".CFB", strTempFile & ".cfb", lFilesize
'
'        If Not IsDesignFile(strTempFile & ".cfb") Then
'            Kill strTempFile & ".cfb"
'
'            If Len(Range("Bit" & iBit & "_Name")) = 5 Then
'                Download_File strLink & "a00.DES", strTempFile & "a00.des", lFilesize
'                If Not IsDesignFile(strTempFile & "a00.des") Then
'                    Kill strTempFile & "a00.des"
'
'                    Download_File strLink & "a00.CFB", strTempFile & "a00.cfb", lFilesize
'
'                    If Not IsDesignFile(strTempFile & "a00.cfb") Then
'                        Kill strTempFile & "a00.cfb"
'
'                        If MsgBox("""" & Range("Bit" & iBit & "_Name") & """ not found on Windchill. Would you like to open Windchill?", vbYesNo + vbExclamation, "File not found") = vbYes Then subOpenIEPage "http://pdmp.corp.smith.com:8080/Windchill/app/"
'
'                    Else
'                        Range("Bit" & iBit & "_Name") = Range("Bit" & iBit & "_Name") & "a00"
'                        strTempFile = strTempFile & "a00.cfb"
'                    End If
'                Else
'                    Range("Bit" & iBit & "_Name") = Range("Bit" & iBit & "_Name") & "a00"
'                    strTempFile = strTempFile & "a00.des"
'                End If
'            Else
'                If MsgBox("""" & Range("Bit" & iBit & "_Name") & """ not found on Windchill. Would you like to open Windchill?", vbYesNo + vbExclamation, "File not found") = vbYes Then subOpenIEPage "http://pdmp.corp.smith.com:8080/Windchill/app/"
'                Exit Sub
'            End If
'        Else
'            strTempFile = strTempFile & ".cfb"
'        End If
'    Else
'        strTempFile = strTempFile & ".des"
'    End If
'
'    ReadBit strTempFile, True
'
''    Kill strTempFile
'
'
'    subActive False
'
'
'
'
'
'End Sub

Function IsDesignFile(strFile As String) As Boolean
    Dim fs As scripting.FileSystemObject
    Dim txtIn As Object ' scripting.textstream
    Dim strLine As String, strArray() As String

    Set fs = New scripting.FileSystemObject
    Set txtIn = fs.OpenTextFile(strFile, 1) ' 1 ForReading
        
    strLine = txtIn.ReadLine
    
    If Left(strLine, 1) = "[" Then 'designtype]" Then
        IsDesignFile = True
    Else
        strLine = txtIn.ReadLine
        If InStr(1, strLine, "cut#") > 0 Then
            IsDesignFile = True
        End If
    End If
    
    Set txtIn = Nothing
    Set fs = Nothing

End Function

Function OnlineFileExists(strPath As String) As Boolean

    gProc = "OnlineFileExists"
    
' Code taken from http://www.mrexcel.com/forum/excel-questions/699272-how-check-if-file-exists-web.html
' on 2015/06/16
    Dim oXHTTP As Object
    On Error GoTo haveError
    
    Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    If Not UCase(strPath) Like "HTTP:*" Then
        strPath = "http://" & strPath
    End If
    
    oXHTTP.Open "HEAD", strPath, False
    oXHTTP.sEnd
    
    If oXHTTP.Status = 200 Then OnlineFileExists = True
    
    Exit Function
    
haveError:
    OnlineFileExists = False
End Function

Function Download_File(ByVal vWebFile As String, ByVal vLocalFile As String, lFilesize As Long) As Boolean

    gProc = "Download_File"
    
'   Code taken from: http://stackoverflow.com/questions/12815142/vba-download-and-save-in-c-user
'   on 2013-06-19
    Dim strHeader As Variant
    Dim oXMLHTTP As Object, i As Long, vFF As Long, oResp() As Byte
    
    On Error Resume Next

'Debug.Print "Download_File: Stage 1"
    'You can also set a ref. to Microsoft XML, and Dim oXMLHTTP as MSXML2.XMLHTTP
'    Me.lblStatus.Caption = "Stage 1: Request Connection": Me.Repaint: Debug.Print Time$: DoEvents
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP") ': Me.Repaint: Debug.Print Time$: DoEvents
    oXMLHTTP.Open "GET", vWebFile, False ': Me.Repaint: Debug.Print Time$: DoEvents 'Open socket to get the website
    oXMLHTTP.sEnd ': Me.Repaint: Debug.Print Time$: DoEvents 'send request

'Debug.Print "Download_File: Stage 2"
    'Wait for request to finish
'    Me.lblStatus.Caption = "Stage 2: Connecting...": Me.Repaint: DoEvents
    Do While oXMLHTTP.readyState <> 4
        DoEvents
    Loop
    
'Debug.Print "Download_File: Stage 3"
'    Me.lblStatus.Caption = "Stage 3: Download File Headers": Me.Repaint: DoEvents
    strHeader = oXMLHTTP.getAllResponseHeaders
'    Stop
    If InStr(1, strHeader, "Content-Length:") > 1 Then
        'Password aborted. Abort Windchill get.
        Set oXMLHTTP = Nothing
        Download_File = False
    ElseIf InStr(1, strHeader, "WWW-Authenticate:") > 1 Then
        lFilesize = CLng(Mid(strHeader, InStr(1, strHeader, "Content-Length:") + 16, InStr(1, strHeader, "Content-Type:") - (InStr(1, strHeader, "Content-Length:") + 16)))
    End If
    
'Debug.Print "Download_File: Stage 4"
'    Me.lblStatus.Caption = "Stage 4: Get file as byte array": Me.Repaint: DoEvents
    oResp = oXMLHTTP.responseBody 'Returns the results as a byte array

'Debug.Print "Download_File: Stage 5"
    'Create local file and save results to it
'    Me.lblStatus.Caption = "Stage 5: Create local file": Me.Repaint: DoEvents
    vFF = FreeFile
    If Dir(vLocalFile) <> "" Then Kill vLocalFile
    Open vLocalFile For Binary As #vFF
    Put #vFF, , oResp
    Close #vFF

'Debug.Print "Download_File: Stage 6"
    'Clear memory
'    Me.lblStatus.Caption = "Stage 6: Close connection/Download Success": Me.Repaint: DoEvents
    Set oXMLHTTP = Nothing
    Download_File = True
'Debug.Print "Download_File: Stage End"
    
End Function

Public Function CheckBitIn(strFile As String, strType As String) As String

    gProc = "CheckBitIn"
    
    'Check if bit is active in Bit.In file. True if it is, false if not.
    
    Dim fs As scripting.FileSystemObject
    Dim txtIn As scripting.TextStream
    Dim strLine As String
    
    On Error GoTo ErrHandler

    Set fs = New scripting.FileSystemObject
    
    If Dir(Left(strFile, InStrRev(strFile, "\")) & "bit.in") = "" Then
        CheckBitIn = ""
        Exit Function
    End If
    
    Set txtIn = fs.OpenTextFile(Left(strFile, InStrRev(strFile, "\")) & "bit.in", 1) ' 1 ForReading
        
    strLine = txtIn.ReadLine

    If strLine = "{" Then 'new file format
        Do While Not txtIn.AtEndOfStream
            strLine = txtIn.ReadLine
            If strType = "Name" Then
                If InStr(1, strLine, """DesName"": {") > 0 Then
                    strLine = txtIn.ReadLine
                    strLine = txtIn.ReadLine
                    
                    CheckBitIn = Mid(strLine, InStr(1, strLine, ":") + 1)
                    
                    CheckBitIn = Replace(CheckBitIn, """", "")
                    CheckBitIn = Replace(CheckBitIn, " ", "")
                    
                    GoTo EndCheckBitIn
                End If
            ElseIf strType = "Diameter" Then
                If InStr(1, strLine, """Diameter"": {") > 0 Then
                    strLine = txtIn.ReadLine
                    strLine = txtIn.ReadLine
                    
                    CheckBitIn = Mid(strLine, InStr(1, strLine, ":") + 1)
                    GoTo EndCheckBitIn
                End If
            End If
        Loop
        
    Else 'old file format
        Do While Not txtIn.AtEndOfStream
            strLine = txtIn.ReadLine
            
            If strType = "Name" Then
                If InStr(1, strLine, "[cfb]") > 0 Or InStr(1, strLine, "[des]") > 0 Then
                    strLine = txtIn.ReadLine
                    
                    CheckBitIn = strLine
                    GoTo EndCheckBitIn
                End If
            ElseIf strType = "Diameter" Then
                If InStr(1, strLine, "[diameter]") > 0 Then
                    strLine = txtIn.ReadLine
                    
                    CheckBitIn = strLine
                    GoTo EndCheckBitIn
                End If
            End If
        Loop
    End If

EndCheckBitIn:
    Set txtIn = Nothing
    Set fs = Nothing
    
    Exit Function
    
ErrHandler:

Stop
Resume

End Function

Public Function CheckReamerIn(strFile As String, strType As String) As String

    gProc = "CheckReamerIn"
    
    'Check if reamer is active in Bit.In file. True if it is, false if not.
    
    Dim fs As scripting.FileSystemObject
    Dim txtIn As scripting.TextStream
    Dim strLine As String
    
    On Error GoTo ErrHandler

    Set fs = New scripting.FileSystemObject
    
    If Dir(Left(strFile, InStrRev(strFile, "\")) & "reamer.in") = "" Then
        CheckReamerIn = ""
        Exit Function
    End If
    
    Set txtIn = fs.OpenTextFile(Left(strFile, InStrRev(strFile, "\")) & "reamer.in", 1) ' 1 ForReading
        
    strLine = txtIn.ReadLine

    If strLine = "{" Then 'new file format
        Do While Not txtIn.AtEndOfStream
            strLine = txtIn.ReadLine
            If strType = "Name" Then
                If InStr(1, strLine, """DrillFileName"": {") > 0 Then
                    strLine = txtIn.ReadLine
                    strLine = txtIn.ReadLine
                    
                    CheckReamerIn = Mid(strLine, InStr(1, strLine, ":") + 1)
                    
                    CheckReamerIn = Replace(CheckReamerIn, """", "")
                    CheckReamerIn = Replace(CheckReamerIn, " ", "")
                    
                    GoTo EndCheckReamerIn
                End If
            ElseIf strType = "Diameter" Then
                If InStr(1, strLine, """DrillDiameter"": {") > 0 Then
                    strLine = txtIn.ReadLine
                    strLine = txtIn.ReadLine
                    
                    CheckReamerIn = Mid(strLine, InStr(1, strLine, ":") + 1)
                    GoTo EndCheckReamerIn
                End If
            End If
        Loop
        
    End If

EndCheckReamerIn:
    Set txtIn = Nothing
    Set fs = Nothing
    
    Exit Function
    
ErrHandler:

Stop
Resume

End Function

Public Sub DESSaveLocation()
    Dim strDir As String
    
    MsgBox "Please select which folder would you like to save your DES files that are downloaded from Windchill.", vbOKOnly, "Pick DES Download Folder"
    
    strDir = funcFolderPicker("DES Download Folder", "", "", "")
    
    If strDir = "" Then
        MsgBox "Folder change cancelled by user."
        Exit Sub
    Else
        Range("DESSaveLoc") = strDir
        Application.Calculate
    End If
    
End Sub

Public Sub CPAReadBit()

    gProc = "CPAReadBit"
    
    Dim sCutVals(3, 2) As Single, sVals() As Single, strBlades() As String
    Dim strRegion As Variant
    Dim i As Integer, X As Integer, iRow As Integer, iPrimCt As Integer, m As Integer
    Dim sAngle As Single, sMin As Single, sMax As Single, sTemp As Single
    Dim strTemp As String, strTemp2 As String
    Dim blnPrim As Boolean, blnBU As Boolean, blnOffP As Boolean, blnOnP As Boolean, blnNormalize As Boolean
    
    On Error GoTo ErrHandler
    
    frmPRPicker.Show
    
    If giBit = 0 Then
        Exit Sub
    ElseIf Range("Bit" & giBit & "_Location") = "" Then
        MsgBox "Bit #" & giBit & " has not been loaded. Please try again."
        Exit Sub
    End If
    
    strRegion = Array("Cone", "Nose", "Shoulder", "Gauge")
    
    gstrPRTracker = Right(Application.Caller, Len(Application.Caller) - Len("cmdPRTracker"))
    
    ShowHideBackups giBit, "Show"
    
    Sheets("CPA-V&V-Risk Tracker").Activate
    
    'Clear cells
    Range(Range("PR_" & gstrPRTracker).Offset(1).Address, Range("PR_BottomRow").Offset(-1, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column + 1).Address).ClearContents
    With Range(Range("PR_Diff").Offset(1).Address, Range("PR_BottomRow").Offset(-1, Range("PR_Diff").Column - Range("PR_DesAtt").Column + 1).Address)
        .ClearContents
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    
'Bit Info
    Range("PR_BitDiam").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_Diam")
    Range("PR_BladeCount").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_BladeCt")
    
'Profile
    Range("PR_ProfileType").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_ProfType")
    Range("PR_ConeAng").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_ConeAngle")
    Range("PR_NoseRad").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_NoseRad")
    Range("PR_ShoulderRad").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_ShoulderRad")
    Range("PR_NoseLoc").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_NoseLocDia")
    Range("PR_ProfHt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_ProfHt")
    Range("PR_GaugeLen").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_GaugeLen")
    Range("PR_TipGrind").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_TipGrind")
    
'Substrates
    Sheets("Bit" & giBit).Activate
    strTemp = ""
    
    ReDim sVals(0) ' cutter quantity
    ReDim strBlades(0) ' cutter type
    
    For iRow = 1 To Range("Bit" & giBit & "_Cutter_Type").Count
        blnNormalize = False ' use as blnFound
        
        For i = 0 To UBound(sVals)
            If strBlades(i) = Range("Bit" & giBit & "_Cutter_Type").Item(iRow) Then
                sVals(i) = sVals(i) + 1
                blnNormalize = True
                Exit For
            End If
        Next i
        
        If blnNormalize = False Then
            ReDim Preserve sVals(i)
            ReDim Preserve strBlades(i)
            strBlades(i) = Range("Bit" & giBit & "_Cutter_Type").Item(iRow)
            sVals(i) = 1
        End If
    Next iRow

    For i = 1 To UBound(sVals)
        strTemp = strTemp & sVals(i) & "x " & strBlades(i) & ", "
    Next i
    
    strTemp = Left(strTemp, Len(strTemp) - 2)
    
    Sheets("CPA-V&V-Risk Tracker").Activate
    Range("PR_Substrates").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp

    
'Backups
    Sheets("Bit" & giBit).Activate

'BU Exposures
    For iRow = 1 To Range("Bit" & giBit & "_Tag").Count
        If Range("Bit" & giBit & "_Tag").Item(iRow) = "1" Or Range("Bit" & giBit & "_Tag").Item(iRow) = "3" Then
            If Abs(Cells(Range("Bit" & giBit & "_Tag").Item(iRow).row, Range("Bit" & giBit & "_Profile_Position").Column) - Cells(Range("Bit" & giBit & "_Tag").Item(iRow).row - 1, Range("Bit" & giBit & "_Profile_Position").Column)) < 0.005 Then
                If Cells(Range("Bit" & giBit & "_Tag").Item(iRow).row, Range("Bit" & giBit & "_Exposure").Column) = 0 Then
                    Cells(Range("Bit" & giBit & "_Tag").Item(iRow).row, Range("Bit" & giBit & "_Exposure").Column) = RoundOff(((Cells(Range("Bit" & giBit & "_Tag").Item(iRow).row - 1, Range("Bit" & giBit & "_Radius").Column) - Cells(Range("Bit" & giBit & "_Tag").Item(iRow).row, Range("Bit" & giBit & "_Radius").Column)) ^ 2 + (Cells(Range("Bit" & giBit & "_Tag").Item(iRow).row - 1, Range("Bit" & giBit & "_Height").Column) - Cells(Range("Bit" & giBit & "_Tag").Item(iRow).row, Range("Bit" & giBit & "_Height").Column)) ^ 2) ^ 0.5 * -1, 2)
                End If
            End If
        End If
    Next iRow
    
    
    For i = 0 To 3
        X = 0
        
        ReDim sVals(Range("Bit" & giBit & "_" & strRegion(i)).Count - 1)
        
        For iRow = 1 To Range("Bit" & giBit & "_" & strRegion(i)).Count
            If (Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow) = "1" Or Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow) = "3") And Not IsStinger(Cells(Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow).row, Range("Bit" & giBit & "_Cutter_Diameter").Column)) Then
                sVals(X) = Cells(Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow).row, Range("Bit" & giBit & "_Exposure").Column)
                X = X + 1
            End If
        Next iRow
        
        If X > 0 Then ReDim Preserve sVals(X - 1)
        
        sCutVals(i, 0) = X
        sCutVals(i, 1) = Application.WorksheetFunction.Average(sVals)
    
    Next i

'BU Cutter Count
    sTemp = sCutVals(0, 0) + sCutVals(1, 0) + sCutVals(2, 0) + sCutVals(3, 0)
    
    strTemp = "Total:    " & sTemp & "  "
    For i = 0 To 3
        If sCutVals(i, 0) = 0 Then
            strTemp = strTemp & " - "
        Else
            strTemp = strTemp & sCutVals(i, 0)
        End If
        
        strTemp = strTemp & "/"
    Next i
    strTemp = Left(strTemp, Len(strTemp) - 1)
    
    Sheets("CPA-V&V-Risk Tracker").Activate
    Range("PR_BUCutCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp
    
    
'BU Exp
    Sheets("Bit" & giBit).Activate
    
    strTemp = ""
    For i = 0 To 3
        If sCutVals(i, 0) = 0 Then
            strTemp = strTemp & " - "
        Else
            strTemp = strTemp & RoundOff(sCutVals(i, 1), 2)
        End If
        
        strTemp = strTemp & "/"
    Next i
    strTemp = Left(strTemp, Len(strTemp) - 1)
    
    Sheets("CPA-V&V-Risk Tracker").Activate
    Range("PR_BUCutExp").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp
    
'BU Cfg
    Sheets("Bit" & giBit).Activate
    
    ReDim sVals(1) '0=blade count, 1=BU blade count
    ReDim strBlades(Range("Bit" & giBit & "_BladeCt"))
    
    For i = 1 To Range("Bit" & giBit & "_BladeCt")
        blnPrim = False
        blnBU = False
        
        For iRow = 1 To Range("Bit" & giBit & "_Blade_Num").Count
            If Range("Bit" & giBit & "_Blade_Num").Item(iRow) = i Then
                If InStr(1, Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Identity").Column), "C") > 0 Then
                    If blnPrim = False Then
                        sVals(0) = sVals(0) + 1
                        blnPrim = True
                        
                        strBlades(i) = "Primary"
                    End If
                ElseIf InStr(1, Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Identity").Column), "BU") > 0 Then
                    If blnBU = False Then
                        sVals(1) = sVals(1) + 1
                        blnBU = True
                        
                        If strBlades(i) <> "Primary" Then strBlades(i) = "Backup"
                    End If
                    
                    If Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Exposure").Column) < 0 Then
                        blnOffP = True
                    Else
                        blnOnP = True
                    End If
                    
                End If
            End If
            
        Next iRow
    Next i
    
    iPrimCt = sVals(0)
    
    If blnOffP Then
        If blnOnP Then
            strTemp = "SMXP " & sVals(0) & "x" & sVals(1)
        Else
            strTemp = "SOFP " & sVals(0) & "x" & sVals(1)
        End If
    ElseIf blnOnP Then
        strTemp = "SONP " & sVals(0) & "x" & sVals(1)
    Else
        strTemp = "-"
    End If
    
    Sheets("CPA-V&V-Risk Tracker").Activate
    Range("PR_BUCutCfg").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp
    Range("PR_BladeCount").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = sVals(0)
        
    If sTemp = 0 Then 'No backups present in design.
        Sheets("CPA-V&V-Risk Tracker").Activate
        Range("PR_BUCutCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "-"
        Range("PR_BUCutExp").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "-"
        Range("PR_BUCutCfg").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "-"
    End If
    
'Cutter Count
    Range("PR_ConeCutCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_PrimaryCutters").Offset(1)
    Range("PR_NoseCutCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_PrimaryCutters").Offset(2)
    Range("PR_ShoulderCutCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_PrimaryCutters").Offset(3)
    Range("PR_GaugeCutCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_PrimaryCutters").Offset(4)
    
    If Range("PR_GaugeCutCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) < Range("PR_BladeCount").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) Then
        With Range("PR_GaugeCutCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    Else
        With Range("PR_GaugeCutCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End If
    
'Cutter Spacing
    Sheets("Bit" & giBit).Activate
    
    For i = 0 To 3
        X = 0
        
        If i = 1 And Range("Bit" & giBit & "_ConeAngle") >= 90 Then
            'skip this
        Else
            ReDim sVals(Range("Bit" & giBit & "_" & strRegion(i)).Count - 1)
            
            For iRow = 1 To Range("Bit" & giBit & "_" & strRegion(i)).Count
                If Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow) = "0" Then
                    sVals(X) = Cells(Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow).row, Range("Bit" & giBit & "_Cutter_Distance").Column)
                    X = X + 1
                End If
            Next iRow
            
            ReDim Preserve sVals(X - 1)
            
            sCutVals(i, 0) = Application.WorksheetFunction.Percentile(sVals, 0)
            sCutVals(i, 1) = Application.WorksheetFunction.Percentile(sVals, 0.5)
            sCutVals(i, 2) = Application.WorksheetFunction.Percentile(sVals, 1)
        End If
    Next i
    
    Sheets("CPA-V&V-Risk Tracker").Activate
    
    For i = 0 To 3
        Range("PR_" & strRegion(i) & "Spacing").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "Min: " & RoundOff(sCutVals(i, 0), 3) & ", Med: " & RoundOff(sCutVals(i, 1), 3) & ", Max: " & RoundOff(sCutVals(i, 2), 3)
    Next i
    
    
'Backrake
    Sheets("Bit" & giBit).Activate
    
    For i = 0 To 3
        X = 0
        
        If i = 1 And Range("Bit" & giBit & "_ConeAngle") >= 90 Then
            'skip this
        Else
            ReDim sVals(Range("Bit" & giBit & "_" & strRegion(i)).Count - 1)
            
            For iRow = 1 To Range("Bit" & giBit & "_" & strRegion(i)).Count
                If Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow) = "0" Then
                    sVals(X) = Cells(Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow).row, Range("Bit" & giBit & "_Back_Rake").Column)
                    X = X + 1
                End If
            Next iRow
            
            ReDim Preserve sVals(X - 1)
            
            sCutVals(i, 0) = Application.WorksheetFunction.Percentile(sVals, 0)
            sCutVals(i, 1) = Application.WorksheetFunction.Average(sVals)
            sCutVals(i, 2) = Application.WorksheetFunction.Percentile(sVals, 1)
        End If
    Next i
    
    Sheets("CPA-V&V-Risk Tracker").Activate

    For i = 0 To 3
        Range("PR_" & strRegion(i) & "BR").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "Min: " & RoundOff(sCutVals(i, 0), 3) & ", Avg: " & RoundOff(sCutVals(i, 1), 3) & ", Max: " & RoundOff(sCutVals(i, 2), 3)
    Next i
    
'Siderake
    Sheets("Bit" & giBit).Activate
    
    For i = 0 To 3
        X = 0
        
        If i = 1 And Range("Bit" & giBit & "_ConeAngle") >= 90 Then
            'skip this
        Else

            ReDim sVals(Range("Bit" & giBit & "_" & strRegion(i)).Count - 1)
            
            For iRow = 1 To Range("Bit" & giBit & "_" & strRegion(i)).Count
                If Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow) = "0" Then
                    sVals(X) = Cells(Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow).row, Range("Bit" & giBit & "_Side_Rake").Column)
                    X = X + 1
                End If
            Next iRow
            
            ReDim Preserve sVals(X - 1)
            
            sCutVals(i, 0) = Application.WorksheetFunction.Percentile(sVals, 0)
            sCutVals(i, 1) = Application.WorksheetFunction.Average(sVals)
            sCutVals(i, 2) = Application.WorksheetFunction.Percentile(sVals, 1)
        End If
    Next i
    
    Sheets("CPA-V&V-Risk Tracker").Activate

    For i = 0 To 3
        Range("PR_" & strRegion(i) & "SR").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "Min: " & RoundOff(sCutVals(i, 0), 3) & ", Avg: " & RoundOff(sCutVals(i, 1), 3) & ", Max: " & RoundOff(sCutVals(i, 2), 3)
    Next i
    
    

    
'ONYX360
    Sheets("Bit" & giBit).Activate

    ReDim sVals(5) '0=1317, 1=1320ONYX, 2=1620ONYX, 3=Profile Angle, 4=SRmin, 5=SRmax
    
    sVals(3) = 999
    sVals(4) = 999
    sVals(5) = -999
    
    For iRow = 1 To Range("Bit" & giBit & "_Cutter_Type").Count
        blnPrim = False
        Select Case Range("Bit" & giBit & "_Cutter_Type").Item(iRow)
            Case "1317": sVals(0) = sVals(0) + 1: blnPrim = True
            Case "1320ONYX360": sVals(1) = sVals(1) + 1: blnPrim = True
            Case "1620ONYX360": sVals(2) = sVals(2) + 1: blnPrim = True
        End Select
        
        If blnPrim Then
            If Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Profile_Angle").Column) < sVals(3) Then sVals(3) = Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Profile_Angle").Column)
            
            If Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Side_Rake").Column) < sVals(4) Then sVals(4) = Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Side_Rake").Column)
        
            If Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Side_Rake").Column) > sVals(5) Then sVals(5) = Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Side_Rake").Column)
        End If
    Next iRow
    
    strTemp = ""
    If sVals(0) > 0 Then strTemp = sVals(0) & "x 1317"
    If sVals(1) > 0 Then
        If strTemp <> "" Then strTemp = strTemp & ", "
        
        strTemp = strTemp & sVals(1) & "x 1320ONYX"
    End If
    If sVals(2) > 0 Then
        If strTemp <> "" Then strTemp = strTemp & ", "
        
        strTemp = strTemp & sVals(2) & "x 1620ONYX"
    End If
    
    Sheets("CPA-V&V-Risk Tracker").Activate
    
    If strTemp = "" Then ' not an ONYX 360 bit, post blanks
        Range("PR_O360Size").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "-"
        Range("PR_O360ProfAng").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "-"
        Range("PR_O360SR").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "-"
    Else 'is an ONYX 360 bit, show stats
        Range("PR_O360Size").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp
        Range("PR_O360ProfAng").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = RoundOff(sVals(3), 2)
        Range("PR_O360SR").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "Min: " & sVals(4) & ", Max: " & sVals(5)
    End If
    
'Stinger
    Sheets("Bit" & giBit).Activate

    For i = 0 To 3
        X = 0
        
        ReDim sVals(Range("Bit" & giBit & "_" & strRegion(i)).Count - 1)
        
        For iRow = 1 To Range("Bit" & giBit & "_" & strRegion(i)).Count
            If (Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow) = "1" Or Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow) = "3") And IsStinger(Cells(Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow).row, Range("Bit" & giBit & "_Cutter_Diameter").Column)) Then
                sVals(X) = Cells(Range("Bit" & giBit & "_" & strRegion(i)).Item(iRow).row, Range("Bit" & giBit & "_Exposure").Column)
                X = X + 1
            End If
        Next iRow
        
        If X > 0 Then ReDim Preserve sVals(X - 1)
        
        sCutVals(i, 0) = X
        sCutVals(i, 1) = Application.WorksheetFunction.Average(sVals)
    
    Next i

'Stinger Count
    If sCutVals(0, 0) + sCutVals(1, 0) + sCutVals(2, 0) + sCutVals(3, 0) > 0 Then 'Stingers are present in this design
        strTemp = "Total:    " & sCutVals(0, 0) + sCutVals(1, 0) + sCutVals(2, 0) + sCutVals(3, 0) & "  "
        For i = 0 To 3
            If sCutVals(i, 0) = 0 Then
                strTemp = strTemp & " - "
            Else
                strTemp = strTemp & sCutVals(i, 0)
            End If
            
            strTemp = strTemp & "/"
        Next i
        strTemp = Left(strTemp, Len(strTemp) - 1)
        
        Sheets("CPA-V&V-Risk Tracker").Activate
        Range("PR_StingCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp
        
        
    'Z Exp
        Sheets("Bit" & giBit).Activate
        
        strTemp = ""
        For i = 0 To 3
            If sCutVals(i, 0) = 0 Then
                strTemp = strTemp & " - "
            Else
                strTemp = strTemp & RoundOff(sCutVals(i, 1), 2)
            End If
            
            strTemp = strTemp & "/"
        Next i
        strTemp = Left(strTemp, Len(strTemp) - 1)
        
        Sheets("CPA-V&V-Risk Tracker").Activate
        Range("PR_StingExp").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp
        
    'Z Cfg
        Sheets("Bit" & giBit).Activate
        
        ReDim sVals(1) '0=blade count, 1=BU blade count
        
        blnOffP = False
        blnOnP = False
        
        For i = 1 To Range("Bit" & giBit & "_BladeCt")
            blnPrim = False
            blnBU = False
            
            For iRow = 1 To Range("Bit" & giBit & "_Blade_Num").Count
                If Range("Bit" & giBit & "_Blade_Num").Item(iRow) = i Then
    '                If IsStinger(Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Cutter_Diameter").Column)) Then
                    
                        If InStr(1, Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Identity").Column), "C") > 0 Then
                            If blnPrim = False Then
                                sVals(0) = sVals(0) + 1
                                blnPrim = True
                            End If
                        ElseIf InStr(1, Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Identity").Column), "BU") > 0 Then
                            If blnBU = False Then
                                sVals(1) = sVals(1) + 1
                                blnBU = True
                            End If
                            
                            If Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Exposure").Column) < -0.005 And Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Radius_Ratio").Column) < 0.99 Then
                                blnOffP = True
                            Else
                                blnOnP = True
                            End If
                            
                        End If
    '                End If
                End If
                
            Next iRow
        Next i
        
        If blnOffP Then
            If blnOnP Then
                strTemp = "ZMXP " & sVals(0) & "x" & sVals(1)
            Else
                strTemp = "ZOFP " & sVals(0) & "x" & sVals(1)
            End If
        ElseIf blnOnP Then
            strTemp = "ZONP " & sVals(0) & "x" & sVals(1)
        Else
            strTemp = "-"
        End If
        
        Sheets("CPA-V&V-Risk Tracker").Activate
        Range("PR_StingCfg").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp
    Else
        Sheets("CPA-V&V-Risk Tracker").Activate
        Range("PR_StingCt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "-"
        Range("PR_StingExp").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "-"
        Range("PR_StingCfg").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = "-"
    End If
    

    
'Central Stinger
    Sheets("Bit" & giBit).Activate
    strTemp = "-"
    
'CS insert size
    For iRow = 1 To Range("Bit" & giBit & "_Cutter_Type").Count
        '1: is it a stinger?
        If IsStinger(Range("Bit" & giBit & "_Cutter_Type").Item(iRow)) Then
            '2: is profile angle = 0?
            If Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Profile_Angle").Column) = 0 Then
                '3: is radius = 0?
                If Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Radius").Column) < 0.001 Then
                    'if all are true, the row is a central stinger
                    strTemp = Range("Bit" & giBit & "_Cutter_Type").Item(iRow)
                    Exit For
                End If
            End If
        End If
    Next iRow
    
    Sheets("CPA-V&V-Risk Tracker").Activate
    Range("PR_CentStingSize").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp
    
'CS Core diam
    If strTemp <> "-" Then
        strTemp = ""
        Sheets("Bit" & giBit).Activate
    
        For iRow = 1 To Range("Bit" & giBit & "_Cutter_Type").Count
            If Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Radius").Column) > 0 And Not IsStinger(Range("Bit" & giBit & "_Cutter_Type").Item(iRow)) Then
                CoreDiamCalculate giBit, Cells(Range("Bit" & giBit & "_Cutter_Type").Item(iRow).row, Range("Bit" & giBit & "_Cutter_Num").Column), sCutVals(0, 0)
                Exit For
            End If
        Next iRow
        
        Sheets("CPA-V&V-Risk Tracker").Activate
        Range("PR_CentStingCoreDiam").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = RoundOff(sCutVals(0, 0), 3)
    Else
        Range("PR_CentStingCoreDiam").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = strTemp
    End If
    

'Cutter Angle Around
    Sheets("Bit" & giBit).Activate
    
    X = 0
    For i = 1 To Range("Bit" & giBit & "_BladeCt")
        X = 0
        For iRow = 1 To Range("Bit" & giBit & "_Blade_Num").Count
            If Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Blade_Num").Column) = i And InStr(1, Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Identity").Column), "C") > 0 Then
                X = X + 1
            End If
        Next iRow
        
        If X > sTemp Then sTemp = X
    Next i
    
    X = CInt(sTemp)
    
    ReDim sVals(iPrimCt, X) ' 0=blade number, 1=cutter angle around
    m = 0 'Primary Blade #
    
    For i = 1 To Range("Bit" & giBit & "_BladeCt")
        X = 1
        If strBlades(i) = "Primary" Then
            m = m + 1
            For iRow = 1 To Range("Bit" & giBit & "_Blade_Num").Count
                If Range("Bit" & giBit & "_Blade_Num").Item(iRow) = i Then
                    
                    If InStr(1, Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Identity").Column), "C") > 0 Then
                        sVals(m, X) = Cells(Range("Bit" & giBit & "_Blade_Num").Item(iRow).row, Range("Bit" & giBit & "_Angle_Around").Column)
                        X = X + 1
                    End If
                End If
            Next iRow
            
            Do While X <= UBound(sVals, 2)
                sVals(m, X) = 999
                X = X + 1
            Loop
        End If
    Next i

'calculate helix/blade angle
    strTemp = ""
    strTemp2 = ""
    
    For i = 1 To iPrimCt
    
        blnNormalize = False
TryWithNormalizedAngles:
        sVals(0, 0) = 0
        sVals(0, 1) = 0
        sMin = 999
        sMax = -999
        
        sAngle = 999

        For X = 1 To UBound(sVals, 2)
            If sVals(i, X) = 999 Then Exit For
            
            If blnNormalize Then
                sTemp = sVals(i, X)
                sVals(i, X) = StandardizeAngle(sVals(i, X))
            End If
            
            If sVals(i, X) < sMin Then sMin = sVals(i, X)
            If sVals(i, X) > sMax Then sMax = sVals(i, X)
            
            If sAngle <> 999 Then sVals(0, 0) = sVals(0, 0) + sAngle - sVals(i, X) ' summing up blade angle changes (helix) to determine average blade helix
            sVals(0, 1) = sVals(0, 1) + sVals(i, X)  'summing up blade angles for each cutter to determine average blade angle
            sAngle = sVals(i, X)
            
        Next X
        
        If sMax - sMin > 300 Then
            blnNormalize = True
            GoTo TryWithNormalizedAngles
        End If
        
        blnNormalize = False
        
        strTemp = strTemp & "  B" & i & ": " & RoundOff(sVals(0, 0) / (X - 1), 1) ' Getting Average Blade Helix
        strTemp2 = strTemp2 & "  B" & i & ": " & RoundOff(sVals(0, 1) / (X - 1), 1) ' Getting Average Blade Angle

        
    Next i
      
    Sheets("CPA-V&V-Risk Tracker").Activate

    Range("PR_Helix").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Right(strTemp, Len(strTemp) - 2)
    Range("PR_BladeAngles").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Right(strTemp2, Len(strTemp2) - 2)
    
    
'Gauge Pad
    Range("PR_GaugePadRad").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_GPRad")
    Range("PR_GaugePadLen").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_GPLen")
    Range("PR_GaugePadNomLen").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_GPNomLen")
    Range("PR_GaugePadTapLen").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_GPTaperLen")
    Range("PR_GaugePadTapAng").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_GPTaperAng")
    Range("PR_GaugePadUcutLen").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_GPUcutLen")
    Range("PR_GaugePadUcutAmt").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_GPUcutAmt")
    Range("PR_GaugePadBeta").Offset(, Range("PR_" & gstrPRTracker).Column - Range("PR_DesAtt").Column) = Range("Bit" & giBit & "_GPBeta")



    
    Range("CPA_" & gstrPRTracker & "_Loaded") = True
    
'Blank
'    Range("PR_").Offset(, Range("PR_DesAtt").Column -  Range("PR_" & gstrPRTracker).Column) = Range("Bit" & giBit & "_")


    CalcCPADifferences


    Exit Sub
    
ErrHandler:

Debug.Print Now(), gProc, Err, Err.Description

Stop
Resume


End Sub

Function StandardizeAngle(sAngle As Single) As Single
' code taken from http://stackoverflow.com/questions/2320986/easy-way-to-keeping-angles-between-179-and-180-degrees on 2017-04-14

'this solution may take longer for large angles, but for this task works fine.
    Do While sAngle <= -180
        sAngle = sAngle + 360
    Loop
    Do While sAngle > 180
        sAngle = sAngle - 360
    Loop
    
    StandardizeAngle = sAngle
    
End Function

Public Sub GetGaugePadInfo(iBit As Integer, sBitSize As Single)

    gProc = "GetGaugePadInfo"
    
    Dim sRad1() As Single, sWidth(), sLen1() As Single, sBeta1() As Single, sTaper1() As Single, sRad2() As Single, sLen2() As Single, sBeta2() As Single, sTaper2() As Single
    Dim i As Integer, sVals(8) As Single, iRowEnd As Integer
    Dim find_GP As Range
    
    On Error GoTo ErrHandler
    
    Sheets("BIT" & iBit & "_CUTTER_FILE").Activate
    
    Set find_GP = ActiveSheet.UsedRange.Find(What:="#padNo(1..)", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If Not find_GP Is Nothing Then
        'Debug.Print find_GP.Address
        iRowEnd = 0
        
        If sBitSize > 0 Then
            Do Until find_GP.Offset(iRowEnd + 1) = "[end_section_simple_pad]"
                iRowEnd = iRowEnd + 1
            Loop
        Else
            Do Until find_GP.Offset(iRowEnd + 1, 1) = ""
                iRowEnd = iRowEnd + 1
            Loop
        End If
        
        If iRowEnd = 0 Then Exit Sub
        
        ReDim sRad1(1 To iRowEnd)
        ReDim sLen1(1 To iRowEnd)
        ReDim sBeta1(1 To iRowEnd)
        ReDim sTaper1(1 To iRowEnd)
        ReDim sRad2(1 To iRowEnd)
        ReDim sLen2(1 To iRowEnd)
        ReDim sBeta2(1 To iRowEnd)
        ReDim sTaper2(1 To iRowEnd)
        ReDim sWidth(1 To iRowEnd)
        
        For i = 1 To iRowEnd
            sRad1(i) = find_GP.Offset(i, 4)
            sLen1(i) = find_GP.Offset(i, 8)
            sBeta1(i) = find_GP.Offset(i, 9)
            sTaper1(i) = find_GP.Offset(i, 10)
            sRad2(i) = find_GP.Offset(i, 12)
            sLen2(i) = find_GP.Offset(i, 11)
            sBeta2(i) = find_GP.Offset(i, 13)
            sTaper2(i) = find_GP.Offset(i, 14)
            sWidth(i) = find_GP.Offset(i, 7)
        Next i
        
    'averaging all values for all blades
        For i = 1 To iRowEnd
            sVals(0) = sVals(0) + sRad1(i)
            sVals(1) = sVals(1) + sLen1(i)
            sVals(2) = sVals(2) + sBeta1(i)
            sVals(3) = sVals(3) + sTaper1(i)
            sVals(4) = sVals(4) + sRad2(i)
            sVals(5) = sVals(5) + sLen2(i)
            sVals(6) = sVals(6) + sBeta2(i)
            sVals(7) = sVals(7) + sTaper2(i)
            sVals(8) = sVals(8) + sWidth(i)
        Next i

        For i = 0 To UBound(sVals)
            sVals(i) = sVals(i) / iRowEnd
        Next i
        
        Sheets("INFO").Select
        Range("Bit" & iBit & "_GPRad") = RoundOff(sVals(0), 4)
        Range("Bit" & iBit & "_GPLen") = RoundOff(sVals(1) + sVals(5), 2)
        Range("Bit" & iBit & "_GPWidth") = RoundOff(sVals(8), 4)
        
        'Calc Nominal Length
        If sBitSize > 0 And GaugePadNom(sBitSize) = sVals(0) Then 'section 1 is nominal
            If sVals(3) = 0 Then
                If sVals(7) = 0 And sVals(5) > 0 Then
                    Range("Bit" & iBit & "_GPNomLen") = RoundOff(sVals(1) + sVals(5), 2)
                Else
                    Range("Bit" & iBit & "_GPNomLen") = RoundOff(sVals(1), 2)
                End If
            Else
                Range("Bit" & iBit & "_GPNomLen") = 0
            End If
        Else
            Range("Bit" & iBit & "_GPNomLen") = 0
        End If
        
        If sVals(3) > 0 Then 'section 1 is tapered
            If sVals(5) > 0 Then
                If sVals(7) > 0 Then
                    If sVals(3) = sVals(7) Then
                        Range("Bit" & iBit & "_GPTaperLen") = RoundOff(sVals(1) + sVals(5), 2)
                        If sVals(3) = sVals(7) Then
                            Range("Bit" & iBit & "_GPTaperAng") = RoundOff(sVals(3), 2)
                        Else
                            Range("Bit" & iBit & "_GPTaperAng") = "'" & RoundOff(sVals(3), 2) & " / " & RoundOff(sVals(7), 2)
                        End If
                    Else
                        Range("Bit" & iBit & "_GPTaperLen") = "'" & RoundOff(sVals(1), 2) & " / " & RoundOff(sVals(5), 2)
                        Range("Bit" & iBit & "_GPTaperAng") = "'" & RoundOff(sVals(3), 2) & " / " & RoundOff(sVals(7), 2)
                    End If
                Else
                    Range("Bit" & iBit & "_GPTaperLen") = RoundOff(sVals(1), 2)
                    Range("Bit" & iBit & "_GPTaperAng") = RoundOff(sVals(3), 2)
                End If
            Else
                Range("Bit" & iBit & "_GPTaperLen") = RoundOff(sVals(1), 2)
                Range("Bit" & iBit & "_GPTaperAng") = RoundOff(sVals(3), 2)
            End If
        Else
            If sVals(5) > 0 And sVals(7) > 0 Then 'section 2 has length and it's tapered
                Range("Bit" & iBit & "_GPTaperLen") = RoundOff(sVals(5), 2)
                Range("Bit" & iBit & "_GPTaperAng") = RoundOff(sVals(7), 2)
            Else ' section 2 eith has no length or it's not taper
                Range("Bit" & iBit & "_GPTaperLen") = 0
                Range("Bit" & iBit & "_GPTaperAng") = 0
            End If
        End If
        
        If sVals(4) = 0 Then ' no undercut from section 1
            If sBitSize > 0 And GaugePadNom(sBitSize) > sVals(0) Then
                Range("Bit" & iBit & "_GPUcutLen") = RoundOff(sVals(1) + sVals(5), 2)
                Range("Bit" & iBit & "_GPUcutAmt") = RoundOff(GaugePadNom(sBitSize) - sVals(0), 2)
            Else
                Range("Bit" & iBit & "_GPUcutLen") = 0
                Range("Bit" & iBit & "_GPUcutAmt") = 0
            End If
        Else 'Section 2 has an undercut
            Range("Bit" & iBit & "_GPUcutLen") = RoundOff(sVals(5), 2)
            Range("Bit" & iBit & "_GPUcutAmt") = RoundOff(sVals(4), 2)
        End If
        
        'Beta
        If sVals(5) = 0 Then 'section 2 has no length, only use beta from section 1
            Range("Bit" & iBit & "_GPBeta") = RoundOff(sVals(2), 2)
        ElseIf sVals(2) = sVals(6) Then
            Range("Bit" & iBit & "_GPBeta") = RoundOff(sVals(2), 2)
        Else
            Range("Bit" & iBit & "_GPBeta") = "'" & RoundOff(sVals(2), 2) & " / " & RoundOff(sVals(6), 2)
        End If
        
    Else 'If is a DES but has no GaugePad, or is a CFB
        Set find_BasicData = ActiveSheet.UsedRange.Find("[BasicData]", LookIn:=xlValues)
        If find_BasicData Is Nothing Then
            Range("Bit" & iBit & "_GPRad") = "-"
            Range("Bit" & iBit & "_GPLen") = "-"
            Range("Bit" & iBit & "_GPNomLen") = "-"
            Range("Bit" & iBit & "_GPTaperLen") = "-"
            Range("Bit" & iBit & "_GPTaperAng") = "-"
            Range("Bit" & iBit & "_GPUcutLen") = "-"
            Range("Bit" & iBit & "_GPUcutAmt") = "-"
            Range("Bit" & iBit & "_GPBeta") = "-"
            Range("Bit" & iBit & "_GPWidth") = "-"
        End If
    End If
    
    Exit Sub

ErrHandler:

Debug.Print Now(), gProc, Err, Err.Description
Stop
Resume

End Sub
'Public Sub GetGaugePadInfo(iBit As Integer, sBitSize As Single)
'
'    gProc = "GetGaugePadInfo"
'
'    Dim sRad1() As Single, sWidth(), sLen1() As Single, sBeta1() As Single, sTaper1() As Single, sRad2() As Single, sLen2() As Single, sBeta2() As Single, sTaper2() As Single
'    Dim i As Integer, sVals(8) As Single, iRowEnd As Integer
'    Dim find_GP As Range
'
'    On Error GoTo ErrHandler
'
'    Sheets("BIT" & iBit & "_CUTTER_FILE").Activate
'
'    Set find_GP = ActiveSheet.UsedRange.Find(What:="#padNo(1..)", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
'    If Not find_GP Is Nothing Then
'        'Debug.Print find_GP.Address
'        iRowEnd = 0
'
'        If sBitSize > 0 Then
'            Do Until find_GP.Offset(iRowEnd + 1) = "[end_section_simple_pad]"
'                iRowEnd = iRowEnd + 1
'            Loop
'        Else
'            Do Until find_GP.Offset(iRowEnd + 1, 1) = ""
'                iRowEnd = iRowEnd + 1
'            Loop
'        End If
'
'        If iRowEnd = 0 Then Exit Sub
'
'        Stop
'
'        ReDim sRad1(1 To iRowEnd)
'        ReDim sLen1(1 To iRowEnd)
'        ReDim sBeta1(1 To iRowEnd)
'        ReDim sTaper1(1 To iRowEnd)
'        ReDim sRad2(1 To iRowEnd)
'        ReDim sLen2(1 To iRowEnd)
'        ReDim sBeta2(1 To iRowEnd)
'        ReDim sTaper2(1 To iRowEnd)
'        ReDim sWidth(1 To iRowEnd)
'
'        For i = 1 To iRowEnd
'            sRad1(i) = find_GP.Offset(i, 4)
'            sLen1(i) = find_GP.Offset(i, 8)
'            sBeta1(i) = find_GP.Offset(i, 9)
'            sTaper1(i) = find_GP.Offset(i, 10)
'            sRad2(i) = find_GP.Offset(i, 12)
'            sLen2(i) = find_GP.Offset(i, 11)
'            sBeta2(i) = find_GP.Offset(i, 13)
'            sTaper2(i) = find_GP.Offset(i, 14)
'            sWidth(i) = find_GP.Offset(i, 7)
'        Next i
'
'    'averaging all values for all blades
'        For i = 1 To iRowEnd
'            sVals(0) = sVals(0) + sRad1(i)
'            sVals(1) = sVals(1) + sLen1(i)
'            sVals(2) = sVals(2) + sBeta1(i)
'            sVals(3) = sVals(3) + sTaper1(i)
'            sVals(4) = sVals(4) + sRad2(i)
'            sVals(5) = sVals(5) + sLen2(i)
'            sVals(6) = sVals(6) + sBeta2(i)
'            sVals(7) = sVals(7) + sTaper2(i)
'            sVals(8) = sVals(8) + sWidth(i)
'        Next i
'
'        For i = 0 To UBound(sVals)
'            sVals(i) = sVals(i) / iRowEnd
'        Next i
'
'        Sheets("INFO").Select
'        Range("Bit" & iBit & "_GPRad") = RoundOff(sVals(0), 4)
'        Range("Bit" & iBit & "_GPLen") = RoundOff(sVals(1) + sVals(5), 2)
'        Range("Bit" & iBit & "_GPWidth") = RoundOff(sVals(8), 4)
'
'        'Calc Nominal Length
'        If sBitSize > 0 And GaugePadNom(sBitSize) = sVals(0) Then 'section 1 is nominal
'            If sVals(3) = 0 Then
'                If sVals(7) = 0 And sVals(5) > 0 Then
'                    Range("Bit" & iBit & "_GPNomLen") = RoundOff(sVals(1) + sVals(5), 2)
'                Else
'                    Range("Bit" & iBit & "_GPNomLen") = RoundOff(sVals(1), 2)
'                End If
'            Else
'                Range("Bit" & iBit & "_GPNomLen") = 0
'            End If
'        Else
'            Range("Bit" & iBit & "_GPNomLen") = 0
'        End If
'
'        If sVals(3) > 0 Then 'section 1 is tapered
'            If sVals(5) > 0 Then
'                If sVals(7) > 0 Then
'                    If sVals(3) = sVals(7) Then
'                        Range("Bit" & iBit & "_GPTaperLen") = RoundOff(sVals(1) + sVals(5), 2)
'                        If sVals(3) = sVals(7) Then
'                            Range("Bit" & iBit & "_GPTaperAng") = RoundOff(sVals(3), 2)
'                        Else
'                            Range("Bit" & iBit & "_GPTaperAng") = "'" & RoundOff(sVals(3), 2) & " / " & RoundOff(sVals(7), 2)
'                        End If
'                    Else
'                        Range("Bit" & iBit & "_GPTaperLen") = "'" & RoundOff(sVals(1), 2) & " / " & RoundOff(sVals(5), 2)
'                        Range("Bit" & iBit & "_GPTaperAng") = "'" & RoundOff(sVals(3), 2) & " / " & RoundOff(sVals(7), 2)
'                    End If
'                Else
'                    Range("Bit" & iBit & "_GPTaperLen") = RoundOff(sVals(1), 2)
'                    Range("Bit" & iBit & "_GPTaperAng") = RoundOff(sVals(3), 2)
'                End If
'            Else
'                Range("Bit" & iBit & "_GPTaperLen") = RoundOff(sVals(1), 2)
'                Range("Bit" & iBit & "_GPTaperAng") = RoundOff(sVals(3), 2)
'            End If
'        Else
'            If sVals(5) > 0 And sVals(7) > 0 Then 'section 2 has length and it's tapered
'                Range("Bit" & iBit & "_GPTaperLen") = RoundOff(sVals(5), 2)
'                Range("Bit" & iBit & "_GPTaperAng") = RoundOff(sVals(7), 2)
'            Else ' section 2 eith has no length or it's not taper
'                Range("Bit" & iBit & "_GPTaperLen") = 0
'                Range("Bit" & iBit & "_GPTaperAng") = 0
'            End If
'        End If
'
'        If sVals(4) = 0 Then ' no undercut from section 1
'            If sBitSize > 0 And GaugePadNom(sBitSize) > sVals(0) Then
'                Range("Bit" & iBit & "_GPUcutLen") = RoundOff(sVals(1) + sVals(5), 2)
'                Range("Bit" & iBit & "_GPUcutAmt") = RoundOff(GaugePadNom(sBitSize) - sVals(0), 2)
'            Else
'                Range("Bit" & iBit & "_GPUcutLen") = 0
'                Range("Bit" & iBit & "_GPUcutAmt") = 0
'            End If
'        Else 'Section 2 has an undercut
'            Range("Bit" & iBit & "_GPUcutLen") = RoundOff(sVals(5), 2)
'            Range("Bit" & iBit & "_GPUcutAmt") = RoundOff(sVals(4), 2)
'        End If
'
'        'Beta
'        If sVals(5) = 0 Then 'section 2 has no length, only use beta from section 1
'            Range("Bit" & iBit & "_GPBeta") = RoundOff(sVals(2), 2)
'        ElseIf sVals(2) = sVals(6) Then
'            Range("Bit" & iBit & "_GPBeta") = RoundOff(sVals(2), 2)
'        Else
'            Range("Bit" & iBit & "_GPBeta") = "'" & RoundOff(sVals(2), 2) & " / " & RoundOff(sVals(6), 2)
'        End If
'
'    Else 'If is a DES but has no GaugePad, or is a CFB
'        Set find_BasicData = ActiveSheet.UsedRange.Find("[BasicData]", LookIn:=xlValues)
'        If find_BasicData Is Nothing Then
'            Range("Bit" & iBit & "_GPRad") = "-"
'            Range("Bit" & iBit & "_GPLen") = "-"
'            Range("Bit" & iBit & "_GPNomLen") = "-"
'            Range("Bit" & iBit & "_GPTaperLen") = "-"
'            Range("Bit" & iBit & "_GPTaperAng") = "-"
'            Range("Bit" & iBit & "_GPUcutLen") = "-"
'            Range("Bit" & iBit & "_GPUcutAmt") = "-"
'            Range("Bit" & iBit & "_GPBeta") = "-"
'            Range("Bit" & iBit & "_GPWidth") = "-"
'        End If
'    End If
'
'    Exit Sub
'
'ErrHandler:
'
'Debug.Print Now(), gProc, Err, Err.Description
'Stop
'Resume
'
'End Sub

Public Function GaugePadNom(sBitSize As Single) As Single
    'Returns the nominal gauge pad radius for a given bit size. Table taken from design book 60088589 Rev- on 2017-04-17
    Dim sFinishGP As Single
    
    If sBitSize <= 6.75 Then
        'Bit OD - 0.015
        sFinishGP = 0.015
        
    ElseIf sBitSize <= 9 Then
        'Bit OD - 0.015
        sFinishGP = 0.015
        
    ElseIf sBitSize <= 13.75 Then
        'Bit OD - 0.025
        sFinishGP = 0.025
        
    ElseIf sBitSize <= 17.5 Then
        'Bit OD - 0.025
        sFinishGP = 0.025
        
    Else '17-17/32 and larger
        'Bit OD - 0.035
        sFinishGP = 0.035
    End If
        
    GaugePadNom = (sBitSize - sFinishGP) / 2
    
End Function

Public Sub CalcCPADifferences()

    gProc = "CalcCPADifferences"
    
    Dim strRegion As Variant, strArray As Variant
    Dim i As Integer, X As Integer, m As Integer, n As Integer
    Dim iComma1(1 To 2) As Integer, iComma2(1 To 2) As Integer, iColon(1 To 2, 2) As Integer
    Dim sVal(1 To 2, 2) As Single
    Dim rng As Range
    Dim blnTest As Boolean
    Dim strTemp As String, strTemp2 As String, strVal As String
    
    On Error GoTo ErrHandler
    
    If Range("CPA_BL_Loaded") = False Or Range("CPA_New_Loaded") = False Then Exit Sub
    
    strRegion = Array("Cone", "Nose", "Shoulder", "Gauge")

    strArray = Array("PR_BitDiam", "PR_BladeCount", "PR_ConeAng", "PR_NoseRad", "PR_ShoulderRad", "PR_NoseLoc", "PR_ProfHt", "PR_GaugeLen", "PR_TipGrind", "PR_ConeCutCt", "PR_NoseCutCt", "PR_ShoulderCutCt", "PR_GaugeCutCt", "PR_GaugePadRad", "PR_GaugePadLen", "PR_GaugePadNomLen", "PR_GaugePadTapLen", "PR_GaugePadTapAng", "PR_GaugePadUcutLen", "PR_GaugePadUcutAmt", "PR_GaugePadBeta", "PR_CentStingCoreDiam")

'Straight value compare
    For i = 0 To UBound(strArray)
        blnTest = False
        Set rng = Range(strArray(i)).Offset(, 3)
        
        rng = RoundOff(Range(strArray(i)).Offset(, 2) - Range(strArray(i)).Offset(, 1), 3)
        
        If rng <> 0 Or blnTest Then
            With rng.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        End If
        Set rng = Nothing
    Next i
    
'More complex compare
'3 values (Min: 0, Avg: 0, Max: 0)

    strArray = Array("Spacing", "BR", "SR")
    
    For i = 0 To UBound(strArray)
        For X = 0 To UBound(strRegion)
            'First Value
            For n = 1 To 2
                iComma1(n) = InStr(1, Range("PR_" & strRegion(X) & strArray(i)).Offset(, n), ",")
                iComma2(n) = InStr(iComma1(n) + 1, Range("PR_" & strRegion(X) & strArray(i)).Offset(, n), ",")
                iColon(n, 0) = InStr(1, Range("PR_" & strRegion(X) & strArray(i)).Offset(, n), ":")
                iColon(n, 1) = InStr(iColon(n, 0) + 1, Range("PR_" & strRegion(X) & strArray(i)).Offset(, n), ":")
                iColon(n, 2) = InStr(iColon(n, 1) + 1, Range("PR_" & strRegion(X) & strArray(i)).Offset(, n), ":")
                
                sVal(n, 0) = Mid(Range("PR_" & strRegion(X) & strArray(i)).Offset(, n), iColon(n, 0) + 1, iComma1(n) - iColon(n, 0) - 1)
                sVal(n, 1) = Mid(Range("PR_" & strRegion(X) & strArray(i)).Offset(, n), iColon(n, 1) + 1, iComma2(n) - iColon(n, 1) - 1)
                sVal(n, 2) = Right(Range("PR_" & strRegion(X) & strArray(i)).Offset(, n), Len(Range("PR_" & strRegion(X) & strArray(i)).Offset(, n)) - iColon(n, 2))
            Next n
            
            strTemp = RoundOff(sVal(2, 0) - sVal(1, 0), 2) & " / " & RoundOff(sVal(2, 1) - sVal(1, 1), 2) & " / " & RoundOff(sVal(2, 2) - sVal(1, 2), 2)
            Range("PR_" & strRegion(X) & strArray(i)).Offset(, 3) = strTemp
            
            If strTemp <> "0 / 0 / 0" Then
                With Range("PR_" & strRegion(X) & strArray(i)).Offset(, 3).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            End If
        Next X
    Next i
    
'Blade Helix/Angle
    If Range("PR_BladeCount").Offset(, 1) <> Range("PR_BladeCount").Offset(, 2) Then
        Range("PR_BladeAngles").Offset(, 3) = "Blade Ct Diff"
        Range("PR_Helix").Offset(, 3) = "Blade Ct Diff"
        
        With Range("PR_BladeAngles").Offset(, 3).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        With Range("PR_Helix").Offset(, 3).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    Else
        strTemp = ""
        strTemp2 = ""
        
        iComma1(1) = 0
        iComma1(2) = 0
        iComma2(1) = 0
        iComma2(2) = 0
        m = 0
        X = 0
        
        For i = 1 To Range("PR_BladeCount").Offset(, 1) '- 1
            For n = 1 To 2
                strVal = Range("PR_BladeAngles").Offset(, n)
                
                iComma1(n) = InStr(iComma1(n) + 1, Range("PR_BladeAngles").Offset(, n), ":")
                iComma2(n) = InStr(iComma2(n) + 1, Range("PR_Helix").Offset(, n), ":")
                
                If i < Range("PR_BladeCount").Offset(, 1) Then
                    sVal(n, 0) = Mid(Range("PR_BladeAngles").Offset(, n), iComma1(n) + 1, InStr(iComma1(n) + 2, Range("PR_BladeAngles").Offset(, n), " ") - iComma1(n) - 1)
                    sVal(n, 1) = Mid(Range("PR_Helix").Offset(, n), iComma2(n) + 1, InStr(iComma2(n) + 2, Range("PR_Helix").Offset(, n), " ") - iComma2(n) - 1)
                Else
                    sVal(n, 0) = Right(Range("PR_BladeAngles").Offset(, n), Len(Range("PR_BladeAngles").Offset(, n)) - iComma1(n))
                    sVal(n, 1) = Right(Range("PR_Helix").Offset(, n), Len(Range("PR_Helix").Offset(, n)) - iComma2(n))
                    
                End If
            Next n
            
            If Abs(sVal(2, 0) - sVal(1, 0)) > 300 Then
                sVal(2, 0) = StandardizeAngle(sVal(2, 0))
                sVal(1, 0) = StandardizeAngle(sVal(1, 0))
            End If
            
            strTemp = strTemp & "B" & i & ": " & RoundOff(sVal(2, 0) - sVal(1, 0), 1) & "  "
            strTemp2 = strTemp2 & "B" & i & ": " & RoundOff(sVal(2, 1) - sVal(1, 1), 1) & "  "
            
            If RoundOff(sVal(2, 0) - sVal(1, 0), 1) <> 0 Then m = 1
            If RoundOff(sVal(2, 1) - sVal(1, 1), 1) <> 0 Then X = 1
        Next i
        
        Range("PR_BladeAngles").Offset(, 3) = Trim(strTemp)
        Range("PR_Helix").Offset(, 3) = Trim(strTemp2)
        
        If m = 1 Then
            With Range("PR_BladeAngles").Offset(, 3).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        End If
        
        If X = 1 Then
            With Range("PR_Helix").Offset(, 3).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        End If
        
    End If
    
    
'Just different
    strArray = Array("PR_ProfileType", "PR_BUCutCfg", "PR_CentStingSize", "PR_Substrates", "PR_BUCutCt", "PR_BUCutExp", "PR_O360Size", "PR_O360ProfAng", "PR_O360SR", "PR_StingCt", "PR_StingExp", "PR_StingCfg")
    
    For i = 0 To UBound(strArray)
        If Range(strArray(i)).Offset(, 1) <> Range(strArray(i)).Offset(, 2) Then
            With Range(strArray(i)).Offset(, 3).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        Else
            Range(strArray(i)).Offset(, 3) = "-"
        End If
    Next i
    
    Exit Sub
            


ErrHandler:

Debug.Print Now(), gProc, Err, Err.Description

If Err = 13 Then 'Type mismatch
'    Stop
    blnTest = True
    Resume Next
Else
    Stop
    Resume
End If
End Sub

Public Sub ClearCPA()

    With Range(Range("PR_BL").Offset(1).Address, Range("PR_BottomRow").Offset(-1, Range("PR_Diff").Column - Range("PR_DesAtt").Column + 1).Address)
        .ClearContents
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    
    Range("CPA_BL_Loaded") = False
    Range("CPA_New_Loaded") = False
    
    Range("PR_BL").Offset(1).Select

End Sub

Public Sub ExportProjRep()
    Dim strName As String, strNewName As String
    Dim strDir As String
    
'60077984_Rev-_Project Report_66821.xlsx
    If Range("GEMSDoc") = "" Or Range("GEMSRev") = "" Then
        MsgBox "Enter GEMS Doc Info to the Project Report sheet."
        Range("GEMSDoc").Select
        Exit Sub
    ElseIf Range("NewBOM") = "" Then
        MsgBox "Enter New 5-digit BOM# to the Project Report sheet."
        Range("NewBOM").Select
        Exit Sub
    End If
    
    MsgBox "Please select which folder would you like to export the Project Report sheets.", vbOKOnly, "Pick Project Report Export Folder"
    
    strDir = funcFolderPicker("Pick Project Report Export Folder", "", "", "")

    If strDir = "" Then
        MsgBox "Export aborted."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    strName = ActiveWorkbook.Name

    Sheets("Knowledge_Capture").Visible = True
    Sheets("Knowledge_Capture").Copy
    strNewName = ActiveWorkbook.Name
    Windows(strName).Activate
    Sheets("Knowledge_Capture").Visible = False
    
    Windows(strName).Activate
    Sheets("Project Closure").Copy Before:=Workbooks(strNewName).Sheets(1)
    
    Windows(strName).Activate
    Sheets("CPA-V&V-Risk Tracker").Copy Before:=Workbooks(strNewName).Sheets(1)
    
    Windows(strName).Activate
    Sheets("Requirements Tracker").Copy Before:=Workbooks(strNewName).Sheets(1)
    
    Windows(strName).Activate
    Sheets("Project Report").Copy Before:=Workbooks(strNewName).Sheets(1)
    
    If Dir(strDir & "\" & Range("GEMSDoc") & "_Rev" & Range("GEMSRev") & "_Project Report_" & Range("NewBOM") & ".xlsx") <> "" Then
        If MsgBox("""" & strDir & "\" & Range("GEMSDoc") & "_Rev" & Range("GEMSRev") & "_Project Report_" & Range("NewBOM") & ".xlsx"" already exists in this directory." & vbCrLf & vbCrLf & "Would you like to overwrite the file?", vbExclamation + vbYesNo, "Overwrite Existing File?") = vbNo Then
            Exit Sub
        Else
            Kill strDir & "\" & Range("GEMSDoc") & "_Rev" & Range("GEMSRev") & "_Project Report_" & Range("NewBOM") & ".xlsx"
        End If
    End If
    
    ActiveWorkbook.SaveAs strDir & "\" & Range("GEMSDoc") & "_Rev" & Range("GEMSRev") & "_Project Report_" & Range("NewBOM") & ".xlsx"
    ActiveWorkbook.Close

    Application.ScreenUpdating = True
    
    MsgBox "Project Report successfully exported.", vbOKOnly, "Export Success"
    
End Sub

Public Sub subBitIndices(iBit As Integer)

    gProc = "subBitIndices"
    
' These Performance Indices were created by James Masdea for his G11 project in 2017
    Dim iA As Integer, iC As Integer
    Dim sB As Single, sD As Single, sE As Single, sF As Single, sG As Single, sh As Single, sI As Single
    Dim iFirstInner As Integer, iLastInner As Integer, iFirstOuter As Integer, iLastOuter As Integer
    Dim sIA As Single, sOA As Single, sOLp As Single, sOAp As Single, sOLA As Single
    Dim sROPr As Single, sROPs As Single, sROPn As Single
    Dim sWEARr As Single, sWEARs As Single, sWEARn As Single
    Dim sIMPACTr As Single, sIMPACTn As Single, sROPScale As Single, sWEARScale As Single
    Dim blnHidden As Boolean
        
    On Error GoTo ErrHandler
    
    If IsReamer(iBit) Then Exit Sub
    
    Pi = 4 * Atn(1)

    Sheets("Bit" & iBit).Activate
    
    blnHidden = False
    
'Hide backups, stingers
    For i = 1 To Range("Bit" & iBit & "_Cutter_Type").Count
        If IsStinger(Range("Bit" & iBit & "_Cutter_Type").Item(i)) Then
            Range("Bit" & iBit & "_Cutter_Type").Item(i).Rows.Hidden = True
            blnHidden = True
        ElseIf Range("Bit" & iBit & "_Tag").Item(i) > 0 Then
            Range("Bit" & iBit & "_Tag").Item(i).Rows.Hidden = True
            blnHidden = True
        End If
    Next i
'    Stop
    
    iFirstInner = Range("Bit" & iBit & "_InnerCutters").row
    Do While IsStinger(Cells(iFirstInner, Range("Bit" & iBit & "_Cutter_Diameter").Column))
        iFirstInner = iFirstInner + 1
    Loop
    
    iLastInner = iFirstInner
    
    Do Until Cells(iLastInner, Range("Bit" & iBit & "_Height2").Column) = 0
        iLastInner = iLastInner + 1
    Loop
    
'Test if multiple cutters are at Height2 = 0
    Do While Cells(iLastInner + 1, Range("Bit" & iBit & "_Height2").Column) = 0
        iLastInner = iLastInner + 1
        
        If iLastInner > (Range("Bit" & iBit & "_Cutter_Type").row + Range("Bit" & iBit & "_Cutter_Type").Count) Then GoTo EndSub
    Loop

    
    iFirstOuter = iLastInner + 1
    iLastOuter = Range("Bit" & iBit & "_Shoulder").row + Range("Bit" & iBit & "_Shoulder").Rows.Count - 1
    
'Inner cutters
    'Inner count
    iA = Application.WorksheetFunction.Aggregate(2, 5, Range(Cells(iFirstInner, Range("Bit" & iBit & "_Back_Rake").Column), Cells(iLastInner, Range("Bit" & iBit & "_Back_Rake").Column)))
    'Inner average backrake
    sB = Application.WorksheetFunction.Aggregate(1, 5, Range(Cells(iFirstInner, Range("Bit" & iBit & "_Back_Rake").Column), Cells(iLastInner, Range("Bit" & iBit & "_Back_Rake").Column)))
    
'Outer Cutters
    'Outer count
    iC = Application.WorksheetFunction.Aggregate(2, 5, Range(Cells(iFirstOuter, Range("Bit" & iBit & "_Back_Rake").Column), Cells(iLastOuter, Range("Bit" & iBit & "_Back_Rake").Column)))
    'Outer average backrake
    sE = Application.WorksheetFunction.Aggregate(1, 5, Range(Cells(iFirstOuter, Range("Bit" & iBit & "_Back_Rake").Column), Cells(iLastOuter, Range("Bit" & iBit & "_Back_Rake").Column)))
    'Outer average cutter diameter
    sD = Application.WorksheetFunction.Aggregate(1, 5, Range(Cells(iFirstOuter, Range("Bit" & iBit & "_Cutter_Diameter").Column), Cells(iLastOuter, Range("Bit" & iBit & "_Cutter_Diameter").Column)))
    sD2 = sD ^ 2 * Pi / 4
    
'Profile Information
    'Profile Length
    sF = Cells(iLastOuter + 1, Range("Bit" & iBit & "_Profile_Position").Column).Value
    'Profile Length to Apex
    sG = Cells(iLastInner, Range("Bit" & iBit & "_Profile_Position").Column).Value

'General Information
    'Nose Diameter
    sh = Range("Bit" & iBit & "_NoseLocDia")
    'Bit Diameter
    sI = Range("Bit" & iBit & "_Diam")
    
'    Debug.Print "InCt: " & iA & "   " & "InAvgBR: " & sB & "   " & "OutCt: " & iC & "   " & "OutAvgCutDiam: " & sD & "   " & "OutAvgBR: " & sE & "   " & "ProfLen: " & sF & "   " & "ProfLen2Apex: " & sG & "   " & "NoseDiam: " & sh & "   " & "BitDiam: " & sI
'    Debug.Print "Inner Count: " & iA & "  " & "Inner Avg BR: " & sB & "  " & "Outer Count: " & iC & "  " & "Outer Avg Cutter Diam: " & sD & "  " & "Outer Avg BR: " & sE & "  " & "Profile Len: " & sF & "  " & "Prof Len to Apex: " & sG & "  " & "Nose Diam: " & sH & "  " & "Bit Diam: " & sI
'    Stop
    
'Base Calculations
    'inner area
    sIA = Pi * sh ^ 2 / 4
    'outer area
    sOA = Pi * sI ^ 2 / 4 - sIA
    'outer length%
    sOLp = (sF - sG) / sF
    'outer area%
    sOAp = sOA / (sIA + sOA)
    'outer length/area
    sOLA = sOLp / sOAp
    
'Raw Indices
    sROPr = (((iA ^ 0.065) * (sIA ^ 0.35) * (sB ^ 0.17)) ^ 2 + ((iC ^ 0.065) * (sOA ^ 0.35) * (sE ^ 0.17)) ^ 2) ^ 0.5
    sWEARr = (iC ^ 0.66) * (sD2 ^ 0.46) * (sE ^ 0.161) * (sOLp ^ -0.395) * (sOLA ^ -0.28)
    sIMPACTr = ((((iA ^ 1) / (sIA ^ 1.96)) * (sB ^ 0.88)) ^ 1.02 + (((iC ^ 2.16) * ((sD ^ 2 * Pi / 4) ^ 1.51) / (sOA ^ 1.14)) * (sE ^ -0.06)) ^ 0.96) ^ 0.36
    
    sROPScale = 0.73
    sWEARScale = -0.898
    
'Scaled Indices
    sROPs = sROPr / (sI ^ sROPScale)
    sWEARs = sWEARr * (sI ^ sWEARScale)
    
'Normalized Indices
    sROPn = (((1.99) - sROPs) / ((1.99) - (sI * 0.021 + 1.5))) * 100
    sWEARn = (sWEARs - 0.81) / (1.675 - 0.81) * 100
    sIMPACTn = (sIMPACTr - 0.8) / (1.58 - 0.8) * 100
    
    Range("Bit" & iBit & "_RatingROP") = RoundOff(CDbl(sROPn), 1)
    Range("Bit" & iBit & "_RatingWear") = RoundOff(CDbl(sWEARn), 1)
    Range("Bit" & iBit & "_RatingImpact") = RoundOff(CDbl(sIMPACTn), 1)
    
'End James' indices
    
    
    
' These Performance Indices were created by Katrina Hergenreter for her G11 project in 2017
    Dim strLoc As Variant, d As Integer, A As Single, B As Single, c As Single, sLen(3) As Single, sAgg(3) As Single, X As Integer
    Dim iColTag As Integer, iCol As Integer, strTemp As String, iTemp As Integer, sTemp As Single
    
    strLoc = Array("Cone", "Nose", "Shoulder")
    
'Stop

    For i = 0 To 2
        If Range("Bit" & iBit & "_" & strLoc(i)).Address = "$P$1" Then GoTo nextRegion
        
        A = 0
        B = 0
        c = 0
        
        d = Range("Bit" & iBit & "_" & strLoc(i)).Count
        
        iColTag = Range("Bit" & iBit & "_" & strLoc(i)).Column
        
        iCol = Range("Bit" & iBit & "_Cutter_Distance").Column
        
        sLen(i) = Application.WorksheetFunction.Sum(Range("Bit" & iBit & "_" & strLoc(i)).Offset(, iCol - iColTag))
        
    'Average Cutter Size normalized
        iCol = Range("Bit" & iBit & "_Cutter_Type").Column
        For X = 1 To d
            iTemp = CutterDia(Range("Bit" & iBit & "_" & strLoc(i)).Offset(, iCol - iColTag).Item(X), True)
            A = A + (1 + ((iTemp - 9) * 0.9))
        Next X
        A = A / d
        
    'Average Back rake normalized
        iCol = Range("Bit" & iBit & "_Back_Rake").Column
        For X = 1 To d
            iTemp = Range("Bit" & iBit & "_" & strLoc(i)).Offset(, iCol - iColTag).Item(X)
            B = B + (1 + ((iTemp - 45) * 9 / -40))
        Next X
        B = B / d
    
    'Average Cutter Spacing normalized
        iCol = Range("Bit" & iBit & "_Cutter_Distance").Column
        For X = 1 To d
            sTemp = Range("Bit" & iBit & "_" & strLoc(i)).Offset(, iCol - iColTag).Item(X)
            c = c + (1 + ((sTemp - 0) * 9 / 0.75))
        Next X
        c = c / d
    
    'Aggressiveness for region
        If A = 0 Or d = 0 Then
            sAgg(i) = 0
        Else
            sAgg(i) = (B * c) / (A * d)
        End If
    
        Range("Bit" & iBit & "_PI_Agg_" & Left(strLoc(i), 1)) = sAgg(i)
nextRegion:
    Next i

'Overall Indices
    sLen(i) = Application.WorksheetFunction.Sum(Range("Bit" & iBit & "_Cutter_Distance"))
    
    For i = 0 To 2
        sAgg(3) = sAgg(3) + sAgg(i) * sLen(i) / sLen(3)
    Next i
    
    'Overall Aggressiveness
    Range("Bit" & iBit & "_PI_Agg_O") = sAgg(3)
    
    'Side Cutting Capability
    If (sAgg(0) * sLen(0) / sLen(3)) = 0 Then
        Range("Bit" & iBit & "_PI_SCC") = 0
    Else
        Range("Bit" & iBit & "_PI_SCC") = (sAgg(2) * sLen(2) / sLen(3)) / (sAgg(0) * sLen(0) / sLen(3))
    End If
    
    'Tool Face Control
    If sLen(3) = 0 Then
        Range("Bit" & iBit & "_PI_TFC") = 0
    Else
        Range("Bit" & iBit & "_PI_TFC") = 1 / sAgg(3)
    End If
    
'End Katrina's indices
    
    
    
    
    
    
    
EndSub:
    
    If blnHidden Then Rows.Hidden = False
    
    Sheets("HOME").Activate

    Exit Sub
    
ErrHandler:

Debug.Print Now(), gProc, Err, Err.Description

Stop
Resume

End Sub

Sub BladePlot() 'Optional iBit As Integer)
    Dim i As Integer, iStart As Integer, iBit As Integer, blnPlots As Boolean
    Dim oChart As Variant, X As Integer, blnProfTitle As Boolean
    
    On Error GoTo ErrHandler
    
    Sheets("Blades").Activate
    
    iBit = 2
    
    Pi = 4 * Atn(1)
    blnPlots = False

'calculate in X and Y values for each cutter based on Radius and Angle Around
    For i = 1 To Range("Bit" & iBit & "_Radius").Count
        Range("Bit" & iBit & "_Blades_X").Offset(i) = Range("Bit" & iBit & "_Radius").Item(i) * Cos(Range("Bit" & iBit & "_Angle_Around").Item(i) * Pi / 180)
        Range("Bit" & iBit & "_Blades_Y").Offset(i) = Range("Bit" & iBit & "_Radius").Item(i) * Sin(Range("Bit" & iBit & "_Angle_Around").Item(i) * Pi / 180)
        Range("Bit" & iBit & "_Blades").Offset(i) = Range("Bit" & iBit & "_Blade_Num").Item(i)
    Next i
    
'Sort by Cutter #
    Range(fXLNumberToLetter(Range("Bit" & iBit & "_Blades_X").Column) & "2:" & fXLNumberToLetter(Range("Bit" & iBit & "_Blades").Column) & i).Select
    ActiveWorkbook.Worksheets("Blades").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Blades").Sort.SortFields.Add Key:=Range(fXLNumberToLetter(Range("Bit" & iBit & "_Blades").Column) & "2:" & fXLNumberToLetter(Range("Bit" & iBit & "_Blades").Column) & i), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Blades").Sort
        .SetRange Range(fXLNumberToLetter(Range("Bit" & iBit & "_Blades_X").Column) & "1:" & fXLNumberToLetter(Range("Bit" & iBit & "_Blades").Column) & i)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
'Name ranges based on blade #
    iStart = 1
    For i = 1 To Range("Bit" & iBit & "_Radius").Count
        Do While Range("Bit" & iBit & "_Blades").Offset(i) = Range("Bit" & iBit & "_Blades").Offset(iStart)
            i = i + 1
        Loop
        
        Range(Range("Bit" & iBit & "_Blades_X").Offset(iStart).Address & ":" & Range("Bit" & iBit & "_Blades_X").Offset(i - 1).Address).Name = "Bit" & iBit & "_Blades_X" & Range("Bit" & iBit & "_Blades").Offset(iStart)
        Range(Range("Bit" & iBit & "_Blades_Y").Offset(iStart).Address & ":" & Range("Bit" & iBit & "_Blades_Y").Offset(i - 1).Address).Name = "Bit" & iBit & "_Blades_Y" & Range("Bit" & iBit & "_Blades").Offset(iStart)
        
        iStart = i
        i = i - 1
        
    Next i
    
'Remove erroneous blade ranges
    i = Range("Bit" & iBit & "_Blades").Offset(i - 1)
    Do While 1 = 1
        i = i + 1
        ThisWorkbook.Names("Bit" & iBit & "_Blades_X" & i).Delete
        ThisWorkbook.Names("Bit" & iBit & "_Blades_Y" & i).Delete
    Loop
    
BuildPlots:
        
        blnPlots = True
        Sheets("INFO").Activate
        
        Sheets("INFO").ChartObjects("Bit" & iBit & "_Blades_Chart").Select
'        ActiveSheet.Shapes.AddChart.Select
'        ActiveChart.ChartType = 73 'oChart.ChartType ' Application.Worksheets("PROFILE").ChartObjects("Profile2_Chart").ChartType
    
        Do While ActiveChart.SeriesCollection.Count > 0
            ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count).Delete
        Loop
    
        For i = 1 To Range("Bit" & iBit & "_BladeCt")
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(i).Name = "Blade" & i
            ActiveChart.SeriesCollection(i).XValues = Range("Bit" & iBit & "_Blades_X" & i)
            ActiveChart.SeriesCollection(i).Values = Range("Bit" & iBit & "_Blades_Y" & i)
'            ActiveChart.SeriesCollection(i).AxisGroup = oChart.SeriesCollection(i).AxisGroup
'            ActiveChart.SeriesCollection(i).Format.Line.Weight = oChart.SeriesCollection(i).Format.Line.Weight
'            ActiveChart.SeriesCollection(i).Format.Line.ForeColor.RGB = oChart.SeriesCollection(i).Format.Line.ForeColor.RGB
'            ActiveChart.SeriesCollection(i).Format.Line.DashStyle = oChart.SeriesCollection(i).Format.Line.DashStyle
            ActiveChart.SeriesCollection(i).Smooth = False
            
        Next i
        
FinishPlots:

'        ActiveChart.SetElement msoElementLegendBottom
        
'        ActiveChart.SetElement msoElementChartTitleAboveChart
        ActiveChart.ChartTitle.Text = Range("Bit" & iBit & "_Name") & " Blades"
'        ActiveChart.ChartTitle.Characters.Font.Bold = True

'        ActiveSheet.ChartObjects("Bit" & iBit & "Blades_Chart").Activate
'        ActiveChart.Axes(xlValue).Select
        ActiveChart.Axes(xlValue).MinimumScale = -(Range("Bit" & iBit & "_Diam") / 2)
        ActiveChart.Axes(xlValue).MaximumScale = (Range("Bit" & iBit & "_Diam") / 2)
'        ActiveChart.Axes(xlCategory).Select
        ActiveChart.Axes(xlCategory).MinimumScale = -(Range("Bit" & iBit & "_Diam") / 2)
        ActiveChart.Axes(xlCategory).MaximumScale = (Range("Bit" & iBit & "_Diam") / 2)

    'Sizing chart
'        ActiveChart.ChartArea.Width = oChart.ChartArea.Width
'        ActiveChart.ChartArea.height = oChart.ChartArea.height
'        ActiveChart.PlotArea.Width = oChart.PlotArea.Width
'        ActiveChart.PlotArea.height = oChart.PlotArea.height
'        ActiveChart.PlotArea.Left = oChart.PlotArea.Left
'        ActiveChart.PlotArea.Top = oChart.PlotArea.Top

        MakePlotGridSquareOfActiveChart
    
    
    
    
    
    Exit Sub

ErrHandler:

If Err = 1004 Then
    If blnPlots Then
        GoTo FinishPlots
    Else
        GoTo BuildPlots
    End If
Else
    Debug.Print Now(), Err, Err.Description
    
    Stop
    
    Resume
End If

End Sub

Public Function IsReamer(iBit As Integer) As Boolean
    Dim strSheet As String
    
    strSheet = ActiveSheet.Name
    
    Sheets("Bit" & iBit & "_Cutter_File").Activate
    
    Select Case Range("A2")
        Case "": IsReamer = False
        Case "0": IsReamer = False
        Case "1": IsReamer = True
        Case Else:  IsReamer = False
    End Select
    
    Sheets(strSheet).Activate
    
End Function
        
Public Sub ReadBatch()
    Dim iRow As Integer, iCut As Integer
    Dim strLoc As String
    
    strLoc = "C:\work\HattoriHanzo\"
    iRow = 1
    
    Do While Range("_5dig").Offset(iRow) <> ""
        If Range("_blade1").Offset(iRow) = "" Then
            ReadBit strLoc & Range("_5dig").Offset(iRow, 1), True, 1, True
    
            For iCut = 1 To Range("Bit1_BladeCt").Count
                Range("_blade1").Offset(iRow, iCut - 1) = Range("Bit1_Back_Rake").Item(iCut)
            Next iCut
            
            If iRow Mod 50 = 0 Then ActiveWorkbook.Save
        
        End If
                
        Debug.Print "iRow = " & iRow, Format(iRow / 685, "00.00%")
        Application.StatusBar = "iRow = " & iRow & "    -    " & Format(iRow / 685, "00.00%")
        
        DoEvents
        
        iRow = iRow + 1
    Loop
    
    MsgBox "done"

End Sub

Sub Clear_All(Optional BitNum As Variant)

    Application.ScreenUpdating = False

    gProc = "ClearBit"
    
    On Error GoTo ErrHandler

    subActive True
    

'    If Application.Caller = "cmdReset" Then
'        If MsgBox("This will clear all bits and reset to Case 1. Are you sure you wish to do this?", vbYesNo, "Are you sure?") = vbNo Then
'            Exit Sub
'        End If
'    End If
        
    
    Debug_flag = False
    subUpdate False
    
'    If IsMissing(BitNum) Then BitNum = Right(Application.Caller, 1)
'
'    If IsNumeric(BitNum) = False Then
'        BitNum = 0
'    Else
'        BitNum = CInt(BitNum)
'    End If
    
    'define bitnum as 0, clear all info
    BitNum = 0
    If BitNum = 0 Then
        subClearRFPlot
        Sheets("SUMMARY").Activate
        Range("B6").Name = "CasesID"
    End If
    
    For i = 1 To 5
        If BitNum = 0 Or BitNum = i Then
            Clear_Cutter_Info (i)
            Clear_Summary (i)
            Range("BIT" & i & "_Name").Value = ""
            Range("Bit" & i & "_Status").Value = ""

            'Clear Info2
            Range(Range("Bit" & i & "_I2_ConeSize"), Range("Bit" & i & "_I2_GaugeSR")).ClearContents
            
            'Clear Drill_Scan
            Range("Bit" & i & "_DS_CS_CD").ClearContents
            Range("Bit" & i & "_DS_CS_Ht").ClearContents
            Range("Bit" & i & "_DS_CS_BR").ClearContents
            Range("Bit" & i & "_DS_AG_Ht").ClearContents
            Range("Bit" & i & "_DS_AG_BR").ClearContents
            
            'Clear Stingers Present
            Range("Bit" & i & "_StingersPresent") = False
            
            Sheets("HOME").Shapes("Hide_Bit" & i).ControlFormat.Value = 1
            A = HideCutters(i, False, True, False)
        End If
    Next i
    
    Sheets("HOME").Activate
    
    If BitNum = Range("CoreBit") Then
        Range("CoreDiam") = ""
        Range("StingerTip") = ""
    End If

    If IsRapidFire And BitNum = 0 Then
        subBaselineIteration 1, "BL"
        subBaselineIteration 2, "IT"
        subBaselineIteration 3, "IT"
        subBaselineIteration 4, "IT"
        subBaselineIteration 5, "IT"
        
'        Sheets("HOME").ListBoxes("RF").Value = 1
        Sheets("HOME").ListBoxes("RFPlotBox1").Value = 1
        Range("RFCaseID1") = 1
        
        subUpdateShowHide
    End If
    'hide zone
    'Range("Limits").Columns.EntireColumn.Hidden = True
    
    ZoI "Hide"
    gblnZoIShtOverride = False
    
    If BitNum = 0 Or BitNum = 1 Then SetCaseTo1
    
'    Sheets("HOME").OptionButtons("Option_Button_Hide_Zone").Value = Checked
'    Range("Limits").Columns.EntireColumn.Hidden = True
'    Range("Zone").Rows.EntireRow.Hidden = True
    
    Application.Calculation = xlCalculationAutomatic
    
    subActive False
    
    Exit Sub
    
ErrHandler:

    Sheets("HOME").Select
    Application.ScreenUpdating = True
    
End Sub
''É¾³ýºÍÐÞ¸ÄÊ×Ò³µÄslides
'Sub new_title_slide()
'Dim ppApp As PowerPoint.Application
'        'É¾³ýÄ£°åÒ³µÄµÚÒ»Ò³ppt
'        ppApp.ActivePresentation.Slides.Range(Array(1)).Delete
'        ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank") ' "2_Title Slide")
''        ppApp.ActivePresentation.Slides.Add ppApp.ActivePresentation.Slides.Count + 1, ppLayoutTitle
''        ppApp.ActiveWindow.View.GotoSlide ppApp.ActivePresentation.Slides.Count
'        Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
'        '¶¨Òå±êÌâ
'        ppSlide.Shapes.Title.TextFrame.TextRange.Text = "IDEAS Test Build Validation"
'        '¶¨Òå¸±±êÌâ£¬´Ë´¦¸±±êÌâÐèÒªÐÞ¸Ä
'        ppSlide.Shapes(1).TextFrame.TextRange.Text = "Build 20191231 vs Build 20200403"
'        'Éú³ÉÈýÒ³ppt£¬´Ë´¦µÄÄÚÈÝ¿ÉÄÜÐèÒªºóÃæÔÙ²¹³ä
'        ' i = 0
'        ' title_array = {"Important Build Notes", "Overview", "Design Tool Functionality Checks"}
'        ' For i = 1 To 3
'        ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
'        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Important Build Notes"
'        ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
'        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Overview"
'        ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
'        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Design Tool Functionality Checks"
'
'End Sub

'Sub test621()
'Dim Casetype As String
'    Casetype = RegularCheck(Worksheets("Home").Range("Bit1_Location"))
'End Sub
Function RegularCheck(path_name As String) As String
Dim path_string1 As Variant
Dim path_string2 As Variant
Dim lNumElements As Long
Dim Case_type As String
' Dim stringOne As String
' Dim regexOne As Object

' Set regexOne = New RegExp

'     regexOne.Pattern = "\-\A.C\-"
'     regexOne.Global = False
'     regexOne.IgnoreCase = IgnoreCase
'match the sub-title of cases
    lNumElements = 0
    path_string = Split(path_name, "\")
    lNumElements = UBound(path_string) - LBound(path_string) + 1
    Case_type = path_string(lNumElements - 2)
    RegularCheck = Split(Case_type, "_")(1)
End Function
Public Function IniFileName() As String
  IniFileName = Left$(Application.ActiveWorkbook.Path, InStrRev(Application.ActiveWorkbook.Path, "\")) & "docs\auto_test_config.ini"
End Function


Private Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String
Dim Worked As Long
Dim RetStr As String * 256
Dim StrSize As Long
Dim iNoOfCharInIni As Variant

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
  Else
    sProfileString = ""
    RetStr = Space(256)
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  ReadIniFileString = sIniString
End Function

'Option Explicit
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

'Public Function IniFileName() As String
'Dim IniFileName As String
'  IniFileName = Application.ActiveWorkbook.Path & "\auto_test_config.ini"
'End Function


'Private Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String
'Dim Worked As Long
'Dim RetStr As String * 128
'Dim StrSize As Long
'Dim iNoOfCharInIni As String
'Dim sIniString As String
'Dim sProfileString As Variant
'
'  iNoOfCharInIni = 0
'  sIniString = ""
'  If Sect = "" Or Keyname = "" Then
'    MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
'  Else
'    sProfileString = ""
'    RetStr = Space(128)
'    StrSize = Len(RetStr)
'    Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, IniFileName)
'    If Worked Then
'      iNoOfCharInIni = Worked
'      sIniString = Left$(RetStr, Worked)
'    End If
'  End If
'  ReadIniFileString = sIniString
'End Function

'Sub Read_ini()
'    Rec = ReadIniFileString("DESfile_path", "cases_path1")
'    MsgBox Rec
'End Sub

'Sub open_file_demo()
'    'declare variable
'Dim pathname As String
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
    Dim cases_type As String
    Dim case_name As Variant

    WorkbookName = "SAM_v11.78.xlsm"
    Workbooks("SAM_v11.78.xlsm").Activate
    MacroName1 = "Module1.ReadBit"
    MacroName2 = "Module1.ExportProfileChartButton"
    MacroName3 = "Module1.Clear_All"
'    MacroName4 = "Module1.AutoRefreshBit"
'    MacroName3 = "frmExportProfile.interface_outside"
'    MacroName4 = "Module1.new_title_slide"


'    Call Module1.Read(read ini file)
'    argument1 = ReadIniFileString("DESfile_path", "cases_path1")
    argument2 = False
    argument3 = 1
    argument4 = True
'    argument5 = ReadIniFileString("DESfile_path", "cases_path2")
    argument6 = 2
    cases_type = ReadIniFileString("static_case_type", "type_list")
'    For Each oLayout In ParentPresentation.SlideMaster.CustomLayouts
    case_list = Split(cases_type, " ")
    For Each case_name In case_list
        path_array(0) = ReadIniFileString(case_name, "cases_path1")
        path_array(1) = ReadIniFileString(case_name, "cases_path2")
        Application.Run "'" & WorkbookName & "'!" & MacroName3
        Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(0), argument4, argument3, argument2
        Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(1), argument4, argument6, argument2
        Application.Run "'" & WorkbookName & "'!" & MacroName2
        Application.Run "'" & WorkbookName & "'!" & MacroName3
    Next
'    path_array(0) = ReadIniFileString("DESfile_pdc_path", "cases_path1")
'    path_array(1) = ReadIniFileString("DESfile_pdc_path", "cases_path2")
'    path_array(2) = ReadIniFileString("DESfile_stinger_path", "cases_path1")
'    path_array(3) = ReadIniFileString("DESfile_stinger_path", "cases_path2")
'    path_array(4) = ReadIniFileString("DESfile_axe_path", "cases_path1")
'    path_array(5) = ReadIniFileString("DESfile_axe_path", "cases_path2")
'    path_array(6) = ReadIniFileString("DESfile_hyper_path", "cases_path1")
'    path_array(7) = ReadIniFileString("DESfile_hyper_path", "cases_path2")
'    path_array(8) = ReadIniFileString("DESfile_px_path", "cases_path1")
'    path_array(9) = ReadIniFileString("DESfile_px_path", "cases_path2")
'    path_array(10) = ReadIniFileString("DESfile_echo_path", "cases_path1")
'    path_array(11) = ReadIniFileString("DESfile_echo_path", "cases_path2")
'    path_array(12) = ReadIniFileString("DESfile_E616_path", "cases_path1")
'    path_array(13) = ReadIniFileString("DESfile_E616_path", "cases_path2")

'    argument7 = ReadIniFileString("DESfile_path", "cases_path2")
'    argument7 = ReadIniFileString("DESfile_path", "cases_path2")

'   Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(10), argument2, argument3, argument4
'   Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(11), argument2, argument6, argument4


'    For i = 0 To 6
'        Application.Run "'" & WorkbookName & "'!" & MacroName3
'        Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(i * 2), argument4, argument3, argument2
''        Application.Run "'" & WorkbookName & "'!" & MacroName4, argument3
'        Application.Run "'" & WorkbookName & "'!" & MacroName1, path_array(i * 2 + 1), argument4, argument6, argument2
''        Application.Run "'" & WorkbookName & "'!" & MacroName4, argument6
'        Application.Run "'" & WorkbookName & "'!" & MacroName2
'        Application.Run "'" & WorkbookName & "'!" & MacroName3
'    Next i


'    Application.Run "'" & WorkbookName & "'!" & MacroName4

    'Application.Run "'" & WorkbookName & "'!" & MacroName3
    'Worksheets("frmExportProfile").Activate
    'Application.Run "test_SAM_v11.78.xlsm!frmExportProfile.optColorNew_Click"
    'Application.Run "'" & WorkbookName & "'!" & MacroName4

End Sub
'Sub AutoRefreshBit(BitNum)
'    Dim strDir As String
'
'    subActive True
'
''    strDir = Range("Bit" & Right(Application.Caller, 1) & "_Location")
''    strDir = Trim(Right(strDir, Len(strDir) - InStr(1, strDir, "]", vbTextCompare)))
'    strDir = Range("Bit" & BitNum & "_Location")
'    strDir = Trim(Right(strDir, Len(strDir) - InStr(1, strDir, "]", vbTextCompare)))
'
'    If strDir = "" Then
'        MsgBox "Please load bit using the ""Read Bit"" buttons.", vbOKOnly, ""
'        GoTo EndProcess
'    End If
'
'    subUpdate False
'
''    If IsRapidFire Then subMoveIterationsDown Right(Application.Caller, 1)
'
'    ReadBit strDir, True
'
'    subUpdate True
'
'EndProcess:
'    subActive False
'
'End Sub
'Sub Auto_static()
'
''    open_file_demo
'    Auto_run_static
''    Call Close_Powerpoint_With_Save
'    Close_Excel_Without_Save
'
'End Sub
'Sub Close_Excel_Without_Save()
'
'    Application.Visible = False
'    Application.DisplayAlerts = False
'    ActiveWorkbook.Close SaveChanges:=False
'    Application.Quit
'
'End Sub

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

'Sub open_Multiply_demo()
'    Dim pathname
'    pathname = Application.ActiveWorkbook.Path & "\test_Multi-Bit_Compare_v2.52.xlsm"
'    Workbooks.Open pathname
'End Sub
'Sub Run_dynamic()
'    Dim WorkbookName As String
'    Dim MacroName1 As String
'    Dim MacroName2 As String
'    Dim MacroName3 As String
'    Dim MacroName4 As String
'    Dim MacroName5 As String
'    Dim MacroName6 As String
'    Dim MacroName7 As String
'
'
'    WorkbookName = "test_Multi-Bit_Compare_v2.52.xlsm"
'    Workbooks(WorkbookName).Activate
'    MacroName1 = "Module1.subFindTemplate"
'    MacroName2 = "Module1.BitBasePicker"
'    MacroName3 = "Module1.StartOver"
'    MacroName4 = "Module1.ToggleIDEASCompare"
'    MacroName5 = "Module1.subMakeSlides"
'    MacroName6 = "Module1.BitCountSelect"
'    MacroName7 = "Module1.fill_GEMs_Info"
'
'
'    Application.Run "'" & WorkbookName & "'!" & MacroName3
'    Application.Run "'" & WorkbookName & "'!" & MacroName6
'    Application.Run "'" & WorkbookName & "'!" & MacroName1
'    Application.Run "'" & WorkbookName & "'!" & MacroName2
'    Application.Run "'" & WorkbookName & "'!" & MacroName4
'    Application.Run "'" & WorkbookName & "'!" & MacroName7
'    Application.Run "'" & WorkbookName & "'!" & MacroName5
'
'
'
'End Sub
Sub clear_template()
Dim TF As String
    TF = WritePrivateProfileString("merge_template", "dirct", "", IniFileName)
End Sub
'Sub Auto_dynamic_test()
'
'    open_Multiply_demo
'    Run_dynamic
''    Call Close_Excel_Without_Save
'
'End Sub
'Public Function GetLayout(ppPres As PowerPoint.Presentation, LayoutName As String, Optional ParentPresentation As Presentation = Nothing) As CustomLayout
''http://stackoverflow.com/questions/20997571/create-a-new-slide-in-vba-for-powerpoint-2010-with-custom-layout-using-master
'
'    If ParentPresentation Is Nothing Then
'        Set ParentPresentation = ppPres
'    End If
'
'    Dim oLayout As CustomLayout
'    For Each oLayout In ParentPresentation.SlideMaster.CustomLayouts
'        If oLayout.Name = LayoutName Then
'            Set GetLayout = oLayout
'            Exit For
'        End If
'    Next
'
'End Function
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
'    For i = 1 To 2
'        ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
'        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Autoforce ¨C PDC test"
'        Next i
'    ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank") ' "2_Title Slide")
'    Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
'    ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Contact Analysis Comparison"
'
'    ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Section Title")
'    Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
'    ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Contact Analysis"
'    ppSlide.Shapes(1).TextFrame.TextRange.Text = "Comparison of Baseline Bit and Proposed Bit MDOC Analysis"

'    i = 1
'    For i = 1 To 14
'        ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
'        ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Contact Analysis Comparison - Blade Top Contact"
'            ' ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
'            ' ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Design Tool Functionality Checks"
'            ' ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank")
'            ' ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Static Analysis Results"
'        Next i
    ppApp.Presentations(ppApp.Presentations.Count).SaveAs merge_template
    ppApp.Presentations(ppApp.Presentations.Count).Close
'    ppApp.Quit
    Set ppApp = Nothing
End Sub
Sub Auto_test()
    clear_template
'    Call open_file_demo
    Auto_run_static
    add_slides
'    Call open_Multiply_demo
'    Call Run_dynamic
'    Call clear_template
'    Close_Excel_Without_Save
End Sub
'Private Sub Workbook_Open()
'    clear_template
''    open_file_demo
'    Auto_run_static
'    add_slides
'    open_Multiply_demo
'    Run_dynamic
'    clear_template
'    Close_Excel_Without_Save
'End Sub



