Sub subMakeSlides(Optional blnEndOfSub As Boolean)
    Dim ppApp As PowerPoint.Application
    Dim ppSlide As PowerPoint.Slide
    Dim ppShpRng As PowerPoint.ShapeRange
'    Dim ppPres As PowerPoint.Presentation
    Dim strBase As String
    Dim iSub As Integer
    Dim iBit As Integer, iNumBits As Integer, iTotalSlides As Integer, iTotalParams As Integer, i As Integer, x As Integer
    Dim iBHA As Integer, strBHA As String, strLastBHA As String
    Dim iWell As Integer, strWell As String, strLastWell As String
    Dim iWOB As Integer
    Dim iRPM As Integer
    Dim iMS As Integer
    Dim iRock As Integer
    Dim iStartRow As Integer
    Dim iPlot As Integer
    Dim strPlot As String
    Dim strArray(6) As String
    Dim iColStart As Integer
    Dim iSide As Integer, sngTop As Single, sngLeft As Single, iSlides As Integer, sngWidth As Single, sngHeight As Single, iCycleSlides As Integer
    Dim fs As Scripting.FileSystemObject
    Dim txtIn As Scripting.TextStream
    Dim strLine As String, strDesc() As String
    Dim sngSSP As Single  'SummaryInfo() As Single,
    Dim ppPic As PowerPoint.Shape
    Dim blnBitFound As Boolean, strSheet As String, strDES As String, strFile As String, strFolder As String, strDAMPlots As String, strDAMPercs As String
    Dim blnGEMS As Boolean, blnIDEASCompare As Boolean
    
    On Error GoTo ErrHandler
    
    gProc = "subMakeSlides"

' 'Check for empty Template field
'     If Range("Template") = "" Then
'         MsgBox "Multi-Bit Template not found. Please use the ""Find Template"" button and retry.", vbOKOnly + vbCritical, "Error"
'         Exit Sub
'     ElseIf Dir(Range("Template"), vbNormal) = "" Then
'         MsgBox "Multi-Bit Template not found. Please use the ""Find Template"" button and retry.", vbOKOnly + vbCritical, "Error"
'         Exit Sub
'     End If

' 'Check for GEMS Ref Doc Info
'     If Range("IncludeGEMSNum") = 2 Then
'         If MsgBox("Do you wish to hide the GEMS info on your slides?" & vbCrLf & vbCrLf & "FYI, to comply with CLMS, the GEMS Doc # and Rev MUST be shown on all slides.", vbYesNo) = vbYes Then
'             If MsgBox("GEMS Ref Doc Info will be hidden on your slides.", vbOKCancel, ":'(") = vbCancel Then
'                 Range("GEMSDoc").Select
'                 Exit Sub
'             End If
'             blnGEMS = False
'         Else
'             Range("IncludeGEMSNum") = 1
'             Range("GEMSDoc").Select
'             Exit Sub
'         End If
  
'     Else
'         If Range("GEMSDoc") = "60..." Or Range("GEMSRev") = "" Or Range("GEMSDoc") = "60xxxxxx" Or Range("GEMSRev") = ":(" Then
'             If MsgBox("GEMS Document Info has not been completed. Do you wish to enter the info now?", vbCritical + vbYesNo, "GEMS Info Missing") = vbYes Then
'                 Range("GEMSDoc").Select
'                 Exit Sub
'             Else
'                 If MsgBox("To comply with CLMS, the GEMS Doc # and Rev MUST be shown on all slides." & vbCrLf & vbCrLf & "Do you wish to hide the GEMS info on your slides?", vbYesNo) = vbYes Then
'                     If MsgBox("GEMS Ref Doc Info will be hidden on your slides.", vbOKCancel, ":(") = vbCancel Then
'                         Range("GEMSDoc").Select
'                         Exit Sub
'                     End If
'                     blnGEMS = False
'                 Else
'                     Range("GEMSDoc").Select
'                     Exit Sub
'                 End If
'             End If
'         Else
'             blnGEMS = True
'         End If
'     End If
    

    ' gblnAbort = False
    ' frmStatus.cmdAbort.Enabled = True

    
    ' UpdateStatus "", , , , , , "Initializing..."
    ' frmStatus.Show False
    
    
    Set fs = New Scripting.FileSystemObject
    

'Find number of comparisons
    iNumBits = Range("CompBits")
    
    ReDim strDesc(iNumBits - 1)
        
    For i = 1 To iNumBits
        If Range("BitName" & i) = "" Then
            MsgBox "Make sure all bits have a name entered!", vbExclamation + vbOKOnly, "Bit Names"
            Range("BitName" & i).Select
            
            frmStatus.hide
            Exit Sub
        End If
    Next i
    
    
    If Sheets("Bit_Compare").Shapes("chkIDEASCompare").ControlFormat.Value = 1 Then
        blnIDEASCompare = True
    Else
        blnIDEASCompare = False
    End If
    
    iTotalParams = WriteOrder(blnIDEASCompare)
    
    
'    GoTo HERE
'Import DES info
    If Sheets("Bit_Compare").Shapes("chkDAM").ControlFormat.Value = 1 Then
'Stop
        Application.ScreenUpdating = False
        BuildDAMBW
        
        strDAMPlots = ""
        With Sheets("Bit_Compare").ListBoxes("lstDAMPlot")
            For x = 1 To .ListCount
                If .Selected(x) Then
                    strDAMPlots = strDAMPlots & "<" & .List(x) & ">"
                    Debug.Print strDAMPlots
                End If
                
            Next x
        End With
        
        strDAMPercs = ""
        With Sheets("Bit_Compare").ListBoxes("lstDAMPerc")
            For x = 1 To .ListCount
                If .Selected(x) Then
                    strDAMPercs = strDAMPercs & "<" & .List(x) & ">"
                    Debug.Print strDAMPercs
                End If
                
            Next x
        End With
        

        For iBit = 1 To iNumBits

            If gblnAbort Then GoTo EndOfSub

            UpdateStatus "", , , , , , "Getting Bit" & iBit & " Design File Info..."
            blnBitFound = False
            Do While blnBitFound = False
                For iSub = 1 To Range("SubBase").Count
                    If Range("BitDES" & iBit) = "" Then
                        Range("BitDES" & iBit) = GetBitName(iBit, 1)
                    End If

                    strDES = Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\" & Range("BitDES" & iBit)

                    If fs.FileExists(strDES) Then
                        blnBitFound = True
                        Exit For
                    End If
                Next iSub
            Loop

            strSheet = OpenText(strDES, "query")

            Reset_UsedRange "DES_Bit" & iBit, ActiveSheet.Name

            Worksheets(strSheet).Select
            Cells.Select
            Cells.Copy

            Worksheets("DES_Bit" & iBit).Activate
            Cells.PasteSpecial xlPasteAll

            Application.DisplayAlerts = False
            Worksheets(strSheet).Delete
            Application.DisplayAlerts = True

            ProcessDes iBit

            For iSub = 1 To Range("SubBase").Count
                UpdateStatus "Getting Bit" & iBit & " DAM Info...", , , (iBit - 1) * Range("SubBase").Count + iSub, iNumBits * Range("SubBase").Count

                If Range("SubBase").Item(iSub, 2) = "" Then Exit For

                If fs.FileExists(Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\" & Summary) Then
                    blnBitFound = True
                    Exit For
                End If

                strFolder = Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub)

'                Call organizedata("DES_Bit" & iBit, strFolder, iBit)
                GetForces "Bit" & iBit, strFolder, iBit, iSub, strDAMPlots
                
                If iBit = 1 Then BuildDAMBWCase iSub
                
                GetDAM "Bit" & iBit, strFolder, iBit, iSub, strDAMPlots
                'GetTorque "Bit" & iBit, strFolder, iBit, iSub
'                Call GetWOB("DES_Bit" & iBit, strFolder, iBit)
'                Call GetRPM("DES_Bit" & iBit, strFolder, iBit)
'                Call Getax("DES_Bit" & iBit, strFolder, iBit)
'                Call Getay("DES_Bit" & iBit, strFolder, iBit)
'                Call Getayz("DES_Bit" & iBit, strFolder, iBit)

            Next iSub

        Next iBit

        HideShowCutters
'        SetGraphFormulas
        subUpdateCasePlots "Graphs", 1, strDAMPlots, False, False

        Sheets("Bit_Compare").Select
    End If
'    Stop
HERE:
    iColStart = Range("BHA").Column - 2
    
    'Confirm info entered into tables
    strBase = Cells(3, 3)
    If Right(strBase, 1) <> "\" Then strBase = strBase & "\"

    iStartRow = Range("Order").Row
    
    iSub = 1
    
    iPlot = 1
    
'    GoTo HERE2
    
    If GatherInfoTest Then
        UpdateStatus "", , , , , , "Gathering Summary Info..."
        GatherSummaryInfo
'        SummaryInfo = GatherSummaryInfo
    End If
HERE2:
    If gblnAbort Then GoTo EndOfSub
    
    
' Open PPT and template
    If funcOpenPPTApp(ppApp, blnGEMS) = False Then Exit Sub ', ppSlide, ppPres


'Test if GEMS info is filled out
    ppApp.Presentations(ppApp.Presentations.Count).Windows(1).ViewType = ppViewNormal
    
    If blnGEMS Then
    'Stop
        ppApp.Presentations(ppApp.Presentations.Count).Designs(1).SlideMaster.Shapes("txtGEMS").Visible = msoTrue
        ppApp.Presentations(ppApp.Presentations.Count).Designs(1).SlideMaster.Shapes("txtGEMS").TextFrame.TextRange.Text = "GEMS #: " & Range("GEMSDoc") & " Rev. " & UCase(Range("GEMSRev"))
    
        GEMSDocInfo ppApp.Presentations(ppApp.Presentations.Count), blnGEMS
    
    Else
        ppApp.Presentations(ppApp.Presentations.Count).Designs(1).SlideMaster.Shapes("txtGEMS").Visible = msoFalse
        GEMSDocInfo ppApp.Presentations(ppApp.Presentations.Count), blnGEMS
    End If


'Add Dynamic Analysis Results Section slide
'    ppApp.ActivePresentation.Slides.Range(Array(1)).Delete
    ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank")
    Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
'define title slide
    ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Dynamic Analysis Results"
'    ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count + 1, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Section Title")
'    Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count)
'    ppSlide.Select
'    ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Dynamic Analysis Results"



    UpdateStatus "", , , , , , "Generating Template..."

' Set height and width of images
    Select Case Range("SlideFormat")
        Case 4: sngWidth = 348
        Case 16:
            If iNumBits <= 2 Then
                sngWidth = 6.29
            Else
                sngWidth = 5.33 '312
            End If
    End Select
    

    Select Case iNumBits
        Case 1: sngHeight = 2.9
        Case 2: sngHeight = 2.9
        Case 3: sngHeight = 2.2 '144
        Case 4: sngHeight = 1.65 '108
        Case 5: sngHeight = 1.45 '86.4
        Case 6: sngHeight = 2.2 '144
        Case 8: sngHeight = 1.65 '108
        Case 9: sngHeight = 2.2 '144
        Case 10: sngHeight = 1.45 '86.4
        Case 12: sngHeight = 1.65 '108
        Case 15: sngHeight = 1.45 '86.4
    End Select
    
    ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
    
    BuildTemplate ppApp, iNumBits, Range("SlideFormat")
    
    
    

    
    i = 1
    Do While i <= Range("CompBits")
        
        If gblnAbort Then GoTo EndOfSub
        
        sngLeft = funcFindLeft(iNumBits, i, 1, 1)
        sngTop = funcFindTop(iNumBits, i, 1, 1)

        MakeBGBox ppApp, "Bit_" & i & "_BG1", sngLeft, sngTop, sngWidth, sngHeight, Range("BitColor" & i).Interior.Color
        
        If Range("NumBits") <= 5 Then
            sngLeft = funcFindLeft(iNumBits, i, 2, 2)
            sngTop = funcFindTop(iNumBits, i, 2, 2)
            MakeBGBox ppApp, "Bit_" & i & "_BG2", sngLeft, sngTop, sngWidth, sngHeight, Range("BitColor" & i).Interior.Color
        End If
        
        If Range("NumBits") = 1 Then
            sngLeft = funcFindLeft(iNumBits, i, 3, 3)
            sngTop = funcFindTop(iNumBits, i, 3, 3)
            MakeBGBox ppApp, "Bit_" & i & "_BG3", sngLeft, sngTop, sngWidth, sngHeight, Range("BitColor" & i).Interior.Color

            sngLeft = funcFindLeft(iNumBits, i, 4, 4)
            sngTop = funcFindTop(iNumBits, i, 4, 4)
            MakeBGBox ppApp, "Bit_" & i & "_BG0", sngLeft, sngTop, sngWidth, sngHeight, Range("BitColor" & i).Interior.Color
'        ElseIf Range("NumBits") <= 5 Then
'            sngLeft = funcFindLeft(iNumBits, i, 2)
'            sngTop = funcFindTop(iNumBits, i, 2)
'            MakeBGBox ppApp, "Bit_" & i & "_BG2", sngLeft, sngTop, sngWidth, sngHeight, Range("BitColor" & i).Interior.Color
        End If
        
        i = i + 1
    Loop

    
'Add Version info into template slide
    ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
    ReplaceText "MultiBit_version", "Created Using Multi-Bit Compare Rev" & Range("ThisVersion") & " & Template: " & Mid(Range("Template"), InStrRev(Range("Template"), "\") + 1), ppApp    ' {MultiBitVers}
    
    
    
    iTotalSlides = 0
    
    Do While iSub <= Range("Order").Count And Range("Order").Item(iSub) <> ""
        
        subParseProject strArray, Range("SubBase").Item(iSub)
        
        
'Add BHA/Well Slide if checkbox is checked
        If Sheets("Bit_Compare").Shapes("chkBHASlides").ControlFormat.Value = 1 Then

'Check if pics are there:
TryNextiSub:
            strPlot = Range("BitPath1") & Range("BitID1") & Range("SubBase").Item(iSub) & "\utility_files\bha.png"
            If Dir(strPlot, vbNormal) = "" Then
            ' Can't find pic, try a different extension
                If Dir(Left(strPlot, Len(strPlot) - 3) & "png", vbNormal) <> "" Then
                    strPlot = Left(strPlot, Len(strPlot) - 3) & "png"
                ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "tif", vbNormal) <> "" Then
                    strPlot = Left(strPlot, Len(strPlot) - 3) & "tif"
                ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "gif", vbNormal) <> "" Then
                    strPlot = Left(strPlot, Len(strPlot) - 3) & "gif"
                ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "jpg", vbNormal) <> "" Then
                    strPlot = Left(strPlot, Len(strPlot) - 3) & "jpg"
                Else
                    GoTo EndBHAWellSlide
                End If
            End If
            

            strPlot = Range("BitPath1") & Range("BitID1") & Range("SubBase").Item(iSub) & "\utility_files\well.png"
            If Dir(strPlot, vbNormal) = "" Then
            ' Can't find pic, try a different extension
                If Dir(Left(strPlot, Len(strPlot) - 3) & "png", vbNormal) <> "" Then
                    strPlot = Left(strPlot, Len(strPlot) - 3) & "png"
                ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "tif", vbNormal) <> "" Then
                    strPlot = Left(strPlot, Len(strPlot) - 3) & "tif"
                ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "gif", vbNormal) <> "" Then
                    strPlot = Left(strPlot, Len(strPlot) - 3) & "gif"
                ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "jpg", vbNormal) <> "" Then
                    strPlot = Left(strPlot, Len(strPlot) - 3) & "jpg"
                Else
                    GoTo EndBHAWellSlide
                End If
            End If

        'both bha and well pics are present. make slide if this is a new BHA/Well
            'Testing if new BHA
            If Len(strArray(1)) >= 2 Then
                i = 2
                strBHA = strArray(1)
                Do While i <= Len(strBHA)
                    If Mid(strBHA, 1, 1) = "B" And IsNumeric(Mid(strBHA, 2, 1)) Then
                        'normal, continue
                        If IsNumeric(Mid(strBHA, i, 1)) Then
                            Debug.Print Left(strBHA, i)
                        Else
                            strBHA = Left(strBHA, i - 1)
                        End If
                    Else
                        'not normal Project Nomenclature. Do not add BHA/Well slides
                        Debug.Print "Not normal Project Nomenclature. Do not add BHA/Well slides."
                        Sheets("Bit_Compare").Shapes("chkBHASlides").ControlFormat.Value = 0
                        Exit Do
                    End If
                    
                    i = i + 1
                Loop
            End If
                
            'Testing if new Well
            If Len(strArray(2)) >= 2 Then
                i = 2
                strWell = strArray(2)
                Do While i <= Len(strWell)
                    If Mid(strWell, 1, 1) = "W" And IsNumeric(Mid(strWell, 2, 1)) Then
                        'normal, continue

                    Else
                        'not normal Project Nomenclature. Do not add BHA/Well slides
                        Debug.Print "Not normal Project Nomenclature. Do not add BHA/Well slides."
                        Sheets("Bit_Compare").Shapes("chkBHASlides").ControlFormat.Value = 0
                        Exit Do
                    End If
                    
                    i = i + 1
                Loop
            End If
                
            If strBHA <> strLastBHA Or strWell <> strLastWell Then
                ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Title Only")
                Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count - 1)
                ppSlide.Select
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = funcFindName(strArray(1), "BHA") & vbCrLf & "Well " & funcFindName(strArray(2), "Well")
                ppSlide.Shapes.Title.TextFrame.TextRange.Font.Size = 26
                
                strPlot = Range("BitPath1") & Range("BitID1") & Range("SubBase").Item(iSub) & "\utility_files\bha.png"
                
                If Dir(strPlot, vbNormal) = "" Then
    ' Can't find pic, try a different extension
                    If Dir(Left(strPlot, Len(strPlot) - 3) & "png", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "png"
                    ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "tif", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "tif"
                    ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "gif", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "gif"
                    ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "jpg", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "jpg"
                    Else
                        strPlot = "NOBHAPIC" 'NextPlot
                    End If
                End If

                If strPlot <> "NOBHAPIC" Then
                    Set ppPic = ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Selection.SlideRange(ppSlide.Name).Shapes.AddPicture(filename:=strPlot, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=0.22 * 72, Top:=72 * 1.09, Width:=8.16 * 72, height:=7.2 * 72)
                    If Sheets("Bit_Compare").Shapes("chkTrans").ControlFormat.Value = 1 Then ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count - 1).Shapes(ppPic.Name).PictureFormat.TransparencyColor = RGB(255, 255, 255)
                    Set ppPic = Nothing
                End If
                
                
                strPlot = Range("BitPath1") & Range("BitID1") & Range("SubBase").Item(iSub) & "\utility_files\well.png"
                
                If Dir(strPlot, vbNormal) = "" Then
    ' Can't find pic, try a different extension
                    If Dir(Left(strPlot, Len(strPlot) - 3) & "png", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "png"
                    ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "tif", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "tif"
                    ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "gif", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "gif"
                    ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "jpg", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "jpg"
                    Else
                        strPlot = "NOWELLPIC"   'NextPlot
                    End If
                End If
            
                If strPlot <> "NOWELLPIC" Then
                    Set ppPic = ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Selection.SlideRange(ppSlide.Name).Shapes.AddPicture(filename:=strPlot, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=8.6 * 72, Top:=1.09 * 72, Width:=7.2 * 72, height:=7.2 * 72)
                    If Sheets("Bit_Compare").Shapes("chkTrans").ControlFormat.Value = 1 Then ppSlide.Shapes(ppPic.Name).PictureFormat.TransparencyColor = RGB(255, 255, 255)
                    Set ppPic = Nothing
                End If
                
                strLastBHA = strBHA
                strLastWell = strWell
            End If
    
        End If
EndBHAWellSlide:
        
' Test to make sure all parameters are defined in table, else go to next project
'        If Not funcTesProject(strArray) Then GoTo NextProject
'        If iTotalSlides = 0 Or (iNumBits <= 5 And iNumBits > 1) Then
'            ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
'            ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Selection.SlideRange.Duplicate
'            ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count - 1
'        End If

        
'Start slide creation
        iSlides = Range("NumSlides")
        
        If iNumBits > 5 Then
            iCycleSlides = iSlides
        ElseIf iNumBits > 1 Then
            iCycleSlides = 2 * iSlides
        Else
            iCycleSlides = 4 * iSlides
        End If
        
        Do While iPlot <= iCycleSlides
            
            If (iNumBits > 1 And iPlot Mod 2 = 1 And iPlot <= 2 * iSlides Or iNumBits > 5) Or (iNumBits = 1 And iPlot Mod 4 = 1) Then
                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Selection.SlideRange.Duplicate
                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count - 1
            End If
            
            If gblnAbort Then GoTo EndOfSub
        
            If iNumBits > 5 Then
                iSide = 1
            ElseIf iPlot Mod 2 = 1 Then
                iSide = 1
            Else
                iSide = 2
            End If
            
            If iSide = 1 Then
                ReplaceText "SlideTitleTop", funcFindName(strArray(2), "Well", Range("BitPath1"), Range("BitID1") & Range("SubBase").Item(iSub)) & "; " & funcFindName(strArray(1), "BHA", Range("BitPath1"), Range("BitID1") & Range("SubBase").Item(iSub)), ppApp  ' {WellDESC}; {BHADESC}
                ReplaceText "SlideTitleBot", funcFindName(strArray(6), "Rock", Range("BitPath1"), Range("BitID1") & Range("SubBase").Item(iSub)) & "; WOB: " & funcFindName(strArray(3), "WOB", Range("BitPath1"), Range("BitID1") & Range("SubBase").Item(iSub)) & "K lb; RPM: " & funcFindName(strArray(4), "RPM", Range("BitPath1"), Range("BitID1") & Range("SubBase").Item(iSub)) & "(R) + " & funcFindName(strArray(5), "MS", Range("BitPath1"), Range("BitID1") & Range("SubBase").Item(iSub)) & "(M)", ppApp     ' {RockGenericDesc} - {RockDesc}; WOB: {WOBDESC}K lb; RPM: {SurfaceRPM}(R) + {MotorRPM}(M)
            End If

            If iNumBits > 1 Then
                ReplaceText "ImgGroupHeader" & iSide, Range("PlotName" & iPlot), ppApp
            Else
                ReplaceText "ImgGroupHeader" & (iPlot Mod 4), Range("PlotName" & iPlot), ppApp
            End If
            
            iBit = 1
            
            Do While iBit <= iNumBits
                If gblnAbort Then GoTo EndOfSub
                
                If iNumBits > 5 Then
                    UpdateStatus "Generating Slide ", Format((iTotalSlides + (iBit / iNumBits)), "0"), iTotalParams * iSlides, iTotalSlides + (iBit / iNumBits), iTotalParams * iCycleSlides
                ElseIf iNumBits > 1 Then
                    UpdateStatus "Generating Slide ", Format((iTotalSlides + (iBit / iNumBits)) / 2, "0"), iTotalParams * iSlides, iTotalSlides + (iBit / iNumBits), iTotalParams * iCycleSlides
                Else
                    UpdateStatus "Generating Slide ", Format(iTotalSlides / 4, "0"), iTotalParams * iSlides, CSng(iTotalSlides / 4), iTotalParams * iSlides
                    Debug.Print iTotalSlides, iTotalParams, iSlides, iCycleSlides
                End If

                
                strPlot = Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\utility_files\" & Range("Plot" & iPlot) & "." & Range("PlotExt" & iPlot)
                If Dir(strPlot, vbNormal) = "" Then
    ' Can't find pic, try a different extension
                    If Dir(Left(strPlot, Len(strPlot) - 3) & "png", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "png"
                    ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "tif", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "tif"
                    ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "gif", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "gif"
                    ElseIf Dir(Left(strPlot, Len(strPlot) - 3) & "jpg", vbNormal) <> "" Then
                        strPlot = Left(strPlot, Len(strPlot) - 3) & "jpg"
                    Else
                        GoTo NextBit   'NextPlot
                    End If
                End If
                
                sngLeft = funcFindLeft(iNumBits, iBit, iSide, iPlot)
                sngTop = funcFindTop(iNumBits, iBit, iSide, iPlot)
                
                
                
'                Dim ppPic As PowerPoint.Shape
                Set ppPic = ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Selection.SlideRange.Shapes.AddPicture(filename:=strPlot, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=sngLeft * 72, Top:=sngTop * 72, Width:=sngWidth * 72, height:=sngHeight * 72)
                
                If Sheets("Bit_Compare").Shapes("chkTrans").ControlFormat.Value = 1 Then
                    ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count - 1).Shapes(ppPic.Name).PictureFormat.TransparencyColor = RGB(CInt(Range("PlotTransR" & iPlot)), CInt(Range("PlotTransG" & iPlot)), CInt(Range("PlotTransB" & iPlot)))
                End If
                
                Set ppPic = Nothing
                
                If iSide = 2 And iPlot = 2 Or iNumBits > 5 Then
                    strDesc(iBit - 1) = Range("BitName" & iBit)
                    
                    If Sheets("Bit_Compare").Shapes("chkCR").ControlFormat.Value = 1 And Dir(Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\summary", vbNormal) <> "" Then
                        strDesc(iBit - 1) = strDesc(iBit - 1) & Chr(13) & "CR=" & Format(Range("BitCR" & iBit).Offset(iSub), "0.00") 'funcGrabText(Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\summary", "           Convergence ratio")
                    End If
                    
'                    If Sheets("Bit_Compare").Shapes("chkROP").ControlFormat.Value = 1 And dir(Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\summary", vbNormal) <> "" Then
'                        strDesc(iBit - 1) = strDesc(iBit - 1) & Chr(13) & "ROP=" & Format(Range("BitROP" & iBit).Offset(iSub), "0.00") 'funcGrabText(Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\summary", "          Average ROP(ft/hr)")
'                    End If
                    
                    If Sheets("Bit_Compare").Shapes("chkROP").ControlFormat.Value = 1 And Dir(Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\summary", vbNormal) <> "" Then
                        If Sheets("Bit_Compare").Shapes("chkROPComp").ControlFormat.Value = 1 And iBit > 1 Then
                            strDesc(iBit - 1) = strDesc(iBit - 1) & Chr(13) & "ROP=" & Format(Range("BitROP" & iBit).Offset(iSub), "0.0") & " (" & Format(Range("BitROPComp" & iBit).Offset(iSub), "0.0") & "%" & ChrW(916) & ")"
                        Else
                            strDesc(iBit - 1) = strDesc(iBit - 1) & Chr(13) & "ROP=" & Format(Range("BitROP" & iBit).Offset(iSub), "0.0")
                        End If
                    ElseIf Sheets("Bit_Compare").Shapes("chkROPComp").ControlFormat.Value = 1 And iBit > 1 And Dir(Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\summary", vbNormal) <> "" Then
                        strDesc(iBit - 1) = strDesc(iBit - 1) & Chr(13) & "ROP=" & Format(Range("BitROPComp" & iBit).Offset(iSub), "0.0") & "%" & ChrW(916)
                    End If
                    
                    If Sheets("Bit_Compare").Shapes("chkTq").ControlFormat.Value = 1 And Dir(Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\summary", vbNormal) <> "" Then
                        If Sheets("Bit_Compare").Shapes("chkTqComp").ControlFormat.Value = 1 And iBit > 1 Then
                            strDesc(iBit - 1) = strDesc(iBit - 1) & Chr(13) & "Tq=" & Format(Range("BitTq" & iBit).Offset(iSub), "0.0") & " (" & Format(Range("BitTqComp" & iBit).Offset(iSub), "0.0") & "%" & ChrW(916) & ")"
                        Else
                            strDesc(iBit - 1) = strDesc(iBit - 1) & Chr(13) & "Tq=" & Format(Range("BitTq" & iBit).Offset(iSub), "0.0")
                        End If
                    ElseIf Sheets("Bit_Compare").Shapes("chkTqComp").ControlFormat.Value = 1 And iBit > 1 And Dir(Range("BitPath" & iBit) & Range("BitID" & iBit) & Range("SubBase").Item(iSub) & "\summary", vbNormal) <> "" Then
                        strDesc(iBit - 1) = strDesc(iBit - 1) & Chr(13) & "Tq=" & Format(Range("BitTqComp" & iBit).Offset(iSub), "0.0") & "%" & ChrW(916)
                    End If
                    
                    If Sheets("Bit_Compare").Shapes("chkSSP").ControlFormat.Value = 1 Then
                        If Range("BitSS" & iBit).Offset(iSub) <> "" Then strDesc(iBit - 1) = strDesc(iBit - 1) & Chr(13) & "SS%=" & Format(Range("BitSS" & iBit).Offset(iSub), "0.00")
                    End If
                
                End If
                
                If iSide = 2 Or iNumBits > 5 Then
                    If Sheets("Bit_Compare").Shapes("chkCR").ControlFormat.Value <> 1 Then ReplaceText "CR", "", ppApp
                    
                    ReplaceText "Bit_" & iBit & "_Desc", strDesc(iBit - 1), ppApp
                    BringToFront "Bit_" & iBit & "_Desc", ppApp
                End If
NextBit:
                iBit = iBit + 1
            Loop
            
            
        
NextPlot:
            iPlot = iPlot + 1
            'Debug.Print "iPlot: " & iPlot, "2S-1: " & 2 * iSlides - 1
            iTotalSlides = iTotalSlides + 1
            
'            If iNumBits > 1 And iPlot Mod 2 = 1 And iPlot <= 2 * iSlides Or iNumBits > 5 Or iNumBits = 1 And iPlot Mod 4 = 1 Then
'                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
'                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Selection.SlideRange.Duplicate
'                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count - 1
'            End If
        Loop

NextProject:

    ' Update and export DAM plots
        If Sheets("Bit_Compare").Shapes("chkDAM").ControlFormat.Value = 1 Then
            subUpdateCasePlots "Graphs", iSub, strDAMPlots

            ExportDAMPlots ppApp, strDAMPlots, strDAMPercs
        End If
        
        iPlot = 1
    
        iSub = iSub + 1
    Loop













'Slides have been created, Delete Template Slide
    ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
    ppApp.Presentations(ppApp.Presentations.Count).Slides.Item(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count).Delete
'    If iNumBits > 5 Or iNumBits = 1 Then ppApp.Presentations(ppApp.Presentations.Count).Slides.Item(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count).Delete
    
    
    
    UpdateStatus "Generating Slide ", CSng(iTotalSlides), iTotalParams * iSlides * 2, CSng(iTotalSlides), iTotalParams * iCycleSlides, True



    
    If Sheets("Bit_Compare").Shapes("chkROP").ControlFormat.Value = 1 Or Sheets("Bit_Compare").Shapes("chkSSP").ControlFormat.Value = 1 Or Sheets("Bit_Compare").Shapes("chkCPU").ControlFormat.Value = 1 Then PlotLayoutChange Range("PlotLayout"), iNumBits
    
    Application.Calculate
    
    
'ROP Slides
    If Sheets("Bit_Compare").Shapes("chkROP").ControlFormat.Value = 1 Then
        UpdateStatus "", , , , , , "Adding Dynamic ROP Slides"
        
'        ppApp.Activate
        ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Activate
        
        
        Sheets("DynROP").Select
        Sheets("DynROP").Activate
        
        ActiveChart.ChartArea.Select
        ActiveChart.ChartArea.Copy

        
        ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count + 1, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Title Only")
        ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
        Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count)
        ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Dynamic ROP"
        
        
        Set ppShpRng = ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse)

        If Range("PlotLayout") = "One" Or iNumBits = 1 Then
            'Center pasted object in the slide
            ppShpRng.height = 7.22 * 72
            ppShpRng.Width = 13.08 * 72
            ppShpRng.Align msoAlignCenters, msoTrue
            ppShpRng.Top = 0.94 * 72
        ElseIf Range("PlotLayout") = "Side" Then
            ppShpRng.height = 6.15 * 72
            ppShpRng.Width = 7.9 * 72
            ppShpRng.Left = 0.07 * 72
            ppShpRng.Top = 1.62 * 72
        End If
        
    
    
        If iNumBits > 1 Then
            Sheets("DynROPComp").Select
            Sheets("DynROPComp").Activate
            
    'Update Compare slides to reflect name of bit 1
            ActiveChart.ChartTitle.Select
            ActiveChart.ChartTitle.Text = "% change in Dynamic ROP compared to " & Range("BitName1")
            
            strLine = ""
            'Test values range
            For iBit = 2 To Range("CompBits")
                For i = 1 To Range("ROPComp_" & iBit).Count
                    If Range("ROPComp_" & iBit).Item(i) > 0 Then 'value is positive
                        If strLine = "" Or strLine = "+" Then
                            strLine = "+"
                        ElseIf strLine = "-" Then
                            strLine = "both"
                            Exit For
                        End If
                    ElseIf Range("ROPComp_" & iBit).Item(i) < 0 Then 'value is negative
                        If strLine = "" Or strLine = "-" Then
                            strLine = "-"
                        ElseIf strLine = "+" Then
                            strLine = "both"
                            Exit For
                        End If
                    Else
                        strLine = ""
                    End If
                Next i
                If strLine = "both" Then Exit For
            Next iBit
            
            If strLine = "+" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MinimumScale = 0
                ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
            ElseIf strLine = "-" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MaximumScale = 0
                ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
            ElseIf strLine = "both" Or strLine = "" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
                ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
            End If
            
            
            ActiveChart.ChartArea.Select
            ActiveChart.ChartArea.Copy
            
            
            If Range("PlotLayout") = "One" Then
                ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count + 1, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Title Only")
                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
                Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count)
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Dynamic ROP Comparison"

            End If
            
            Set ppShpRng = ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse)
    
            If Range("PlotLayout") = "One" Then
                'Center pasted object in the slide
                ppShpRng.height = 7.22 * 72
                ppShpRng.Width = 13.08 * 72
                ppShpRng.Align msoAlignCenters, msoTrue
                ppShpRng.Top = 0.94 * 72
            ElseIf Range("PlotLayout") = "Side" Then
                ppShpRng.height = 6.15 * 72
                ppShpRng.Width = 7.9 * 72
                ppShpRng.Left = 8.02 * 72
                ppShpRng.Top = 1.62 * 72
            End If
            
'            Set ppSlide = Nothing
'            Set ppShpRng = Nothing
        End If
        
        Set ppSlide = Nothing
        Set ppShpRng = Nothing
    End If
    
    
'Torque slides
    If Sheets("Bit_Compare").Shapes("chkTq").ControlFormat.Value = 1 Then
        UpdateStatus "", , , , , , "Adding Dynamic Torque Slides"
        
'        ppApp.Activate
        ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Activate
        
        
        Sheets("DynTq").Select
        Sheets("DynTq").Activate
        
        ActiveChart.ChartArea.Select
        ActiveChart.ChartArea.Copy

        
        ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count + 1, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Title Only")
        ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
        Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count)
        ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Dynamic Torque"
        
        
        Set ppShpRng = ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse)

        If Range("PlotLayout") = "One" Or iNumBits = 1 Then
            'Center pasted object in the slide
            ppShpRng.height = 7.22 * 72
            ppShpRng.Width = 13.08 * 72
            ppShpRng.Align msoAlignCenters, msoTrue
            ppShpRng.Top = 0.94 * 72
        ElseIf Range("PlotLayout") = "Side" Then
            ppShpRng.height = 6.15 * 72
            ppShpRng.Width = 7.9 * 72
            ppShpRng.Left = 0.07 * 72
            ppShpRng.Top = 1.62 * 72
        End If
        
    
    
        If iNumBits > 1 Then
            Sheets("DynTqComp").Select
            Sheets("DynTqComp").Activate
            
    'Update Compare slides to reflect name of bit 1
            ActiveChart.ChartTitle.Select
            ActiveChart.ChartTitle.Text = "% change in Dynamic Torque compared to " & Range("BitName1")
            
            strLine = ""
            'Test values range
            For iBit = 2 To Range("CompBits")
                For i = 1 To Range("TqComp_" & iBit).Count
                    If Range("TqComp_" & iBit).Item(i) > 0 Then 'value is positive
                        If strLine = "" Or strLine = "+" Then
                            strLine = "+"
                        ElseIf strLine = "-" Then
                            strLine = "both"
                            Exit For
                        End If
                    ElseIf Range("TqComp_" & iBit).Item(i) < 0 Then 'value is negative
                        If strLine = "" Or strLine = "-" Then
                            strLine = "-"
                        ElseIf strLine = "+" Then
                            strLine = "both"
                            Exit For
                        End If
                    Else
                        strLine = ""
                    End If
                Next i
                If strLine = "both" Then Exit For
            Next iBit
            
            If strLine = "+" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MinimumScale = 0
                ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
            ElseIf strLine = "-" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MaximumScale = 0
                ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
            ElseIf strLine = "both" Or strLine = "" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
                ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
            End If
            
            
            ActiveChart.ChartArea.Select
            ActiveChart.ChartArea.Copy
            
            
            If Range("PlotLayout") = "One" Then
                ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count + 1, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Title Only")
                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
                Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count)
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Dynamic Torque Comparison"

            End If
            
            Set ppShpRng = ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse)
    
            If Range("PlotLayout") = "One" Then
                'Center pasted object in the slide
                ppShpRng.height = 7.22 * 72
                ppShpRng.Width = 13.08 * 72
                ppShpRng.Align msoAlignCenters, msoTrue
                ppShpRng.Top = 0.94 * 72
            ElseIf Range("PlotLayout") = "Side" Then
                ppShpRng.height = 6.15 * 72
                ppShpRng.Width = 7.9 * 72
                ppShpRng.Left = 8.02 * 72
                ppShpRng.Top = 1.62 * 72
            End If
            
'            Set ppSlide = Nothing
'            Set ppShpRng = Nothing
        End If
        
        Set ppSlide = Nothing
        Set ppShpRng = Nothing
    End If
    

'Stick Slip slides
    If Sheets("Bit_Compare").Shapes("chkSSP").ControlFormat.Value = 1 Then
        UpdateStatus "", , , , , , "Adding Dynamic StickSlip % Slides"
        
'        ppApp.Activate
'        ppApp.WindowState = ppWindowNormal
        ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Activate
        
        Sheets("DynSS").Select
        Sheets("DynSS").Activate
        
        ActiveChart.ChartArea.Select
        ActiveChart.ChartArea.Copy

        ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count + 1, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Title Only")
        ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
        Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count)
        ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Dynamic Stick-Slip Percentage"

        Set ppShpRng = ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse)

        If Range("PlotLayout") = "One" Or iNumBits = 1 Then
            'Center pasted object in the slide
            ppShpRng.height = 7.22 * 72
            ppShpRng.Width = 13.08 * 72
            ppShpRng.Align msoAlignCenters, msoTrue
            ppShpRng.Top = 0.94 * 72
        ElseIf Range("PlotLayout") = "Side" Then
            ppShpRng.height = 6.15 * 72
            ppShpRng.Width = 7.9 * 72
            ppShpRng.Left = 0.07 * 72
            ppShpRng.Top = 1.62 * 72
        End If
        
    
        
        If iNumBits > 1 Then
            Sheets("DynSSComp").Select
            Sheets("DynSSComp").Activate
    
            ActiveChart.ChartTitle.Select
            ActiveChart.ChartTitle.Text = "% change in Dynamic SS% compared to " & Range("BitName1")
    
    
            strLine = ""
            'Test values range
            For iBit = 2 To Range("CompBits")
                For i = 1 To Range("SSComp_" & iBit).Count
                    If Range("SSComp_" & iBit).Item(i) > 0 Then 'value is positive
                        If strLine = "" Or strLine = "+" Then
                            strLine = "+"
                        ElseIf strLine = "-" Then
                            strLine = "both"
                            Exit For
                        End If
                    ElseIf Range("SSComp_" & iBit).Item(i) < 0 Then 'value is negative
                        If strLine = "" Or strLine = "-" Then
                            strLine = "-"
                        ElseIf strLine = "+" Then
                            strLine = "both"
                            Exit For
                        End If
                    Else
                        strLine = ""
                    End If
                Next i
                If strLine = "both" Then Exit For
            Next iBit
            
            If strLine = "+" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MinimumScale = 0
                ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
            ElseIf strLine = "-" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MaximumScale = 0
                ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
            ElseIf strLine = "both" Or strLine = "" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
                ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
            End If
    
            
            ActiveChart.ChartArea.Select
            ActiveChart.ChartArea.Copy
    
            If Range("PlotLayout") = "One" Then
                ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count + 1, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Title Only")
                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
                Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count)
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Dynamic SS% Comparison"
                
            End If
            
            Set ppShpRng = ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse)
    
            If Range("PlotLayout") = "One" Then
                'Center pasted object in the slide
                ppShpRng.height = 7.22 * 72
                ppShpRng.Width = 13.08 * 72
                ppShpRng.Align msoAlignCenters, msoTrue
                ppShpRng.Top = 0.94 * 72
            ElseIf Range("PlotLayout") = "Side" Then
                ppShpRng.height = 6.15 * 72
                ppShpRng.Width = 7.9 * 72
                ppShpRng.Left = 8.02 * 72
                ppShpRng.Top = 1.62 * 72
            End If
            
            Set ppSlide = Nothing
            Set ppShpRng = Nothing
        End If
        
        Set ppSlide = Nothing
        Set ppShpRng = Nothing
    End If
    
    If Sheets("Bit_Compare").Shapes("chkCPU").ControlFormat.Value = 1 Then
        UpdateStatus "", , , , , , "Adding Simulation CPU Time Slides"
        
'        ppApp.Activate
'        ppApp.WindowState = ppWindowNormal
        ppApp.Presentations(ppApp.Presentations.Count).Windows(1).Activate
        
        Sheets("DynCPU").Select
        Sheets("DynCPU").Activate
        
        ActiveChart.ChartArea.Select
        ActiveChart.ChartArea.Copy

        ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count + 1, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Title Only")
        ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
        Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count)
        ppSlide.Shapes.Title.TextFrame.TextRange.Text = "CPU Simulation Times"
        
        Set ppShpRng = ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse)

        If Range("PlotLayout") = "One" Or iNumBits = 1 Then
            'Center pasted object in the slide
            ppShpRng.height = 7.22 * 72
            ppShpRng.Width = 13.08 * 72
            ppShpRng.Align msoAlignCenters, msoTrue
            ppShpRng.Top = 0.94 * 72
        ElseIf Range("PlotLayout") = "Side" Then
            ppShpRng.height = 6.15 * 72
            ppShpRng.Width = 7.9 * 72
            ppShpRng.Left = 0.07 * 72
            ppShpRng.Top = 1.62 * 72
        End If
        
    
        
        If iNumBits > 1 Then
            Sheets("DynCPUComp").Select
            Sheets("DynCPUComp").Activate
    
            ActiveChart.ChartTitle.Select
            ActiveChart.ChartTitle.Text = "% change in Simulation CPU Time compared to " & Range("BitName1")
    
    
            strLine = ""
            'Test values range
            For iBit = 2 To Range("CompBits")
                For i = 1 To Range("CPUComp_" & iBit).Count
                    If Range("CPUComp_" & iBit).Item(i) > 0 Then 'value is positive
                        If strLine = "" Or strLine = "+" Then
                            strLine = "+"
                        ElseIf strLine = "-" Then
                            strLine = "both"
                            Exit For
                        End If
                    ElseIf Range("CPUComp_" & iBit).Item(i) < 0 Then 'value is negative
                        If strLine = "" Or strLine = "-" Then
                            strLine = "-"
                        ElseIf strLine = "+" Then
                            strLine = "both"
                            Exit For
                        End If
                    Else
                        strLine = ""
                    End If
                Next i
                If strLine = "both" Then Exit For
            Next iBit
            
            If strLine = "+" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MinimumScale = 0
                ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
            ElseIf strLine = "-" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MaximumScale = 0
                ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
            ElseIf strLine = "both" Or strLine = "" Then
                ActiveChart.ChartArea.Select
                ActiveChart.Axes(xlValue).Select
                ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
                ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
            End If
    
            
            ActiveChart.ChartArea.Select
            ActiveChart.ChartArea.Copy
    
            If Range("PlotLayout") = "One" Then
                ppApp.Presentations(ppApp.Presentations.Count).Slides.AddSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count + 1, GetLayout(ppApp.Presentations(ppApp.Presentations.Count), "Title Only")
                ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide ppApp.Presentations(ppApp.Presentations.Count).Slides.Count
                Set ppSlide = ppApp.Presentations(ppApp.Presentations.Count).Slides(ppApp.Presentations(ppApp.Presentations.Count).Slides.Count)
                ppSlide.Shapes.Title.TextFrame.TextRange.Text = "CPU Simulation Time Comparisons"
                
            End If
            
            Set ppShpRng = ppSlide.Shapes.PasteSpecial(ppPasteEnhancedMetafile, msoFalse, , , , msoFalse)
    
            If Range("PlotLayout") = "One" Then
                'Center pasted object in the slide
                ppShpRng.height = 7.22 * 72
                ppShpRng.Width = 13.08 * 72
                ppShpRng.Align msoAlignCenters, msoTrue
                ppShpRng.Top = 0.94 * 72
            ElseIf Range("PlotLayout") = "Side" Then
                ppShpRng.height = 6.15 * 72
                ppShpRng.Width = 7.9 * 72
                ppShpRng.Left = 8.02 * 72
                ppShpRng.Top = 1.62 * 72
            End If
'            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank")
'            Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
'            ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Conclusion"
'            'write into the configuration file later
'            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
'            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "StrataBlade Changes Build Notes 2020.3.0 on Slide 2 1598771"
'            ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
'            ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Conclusion"
'            ppApp.Presentations(ppApp.Presentations.Count).SaveAs ReadIniFileString("merge_template", "dirct")
''            ppApp.Presentations(ppApp.Presentations.Count).Close
''            ppApp.Quit
            Set ppSlide = Nothing
            Set ppShpRng = Nothing
        End If
        Set ppSlide = Nothing
        Set ppShpRng = Nothing
    End If
    

    

    
    UpdateStatus "", , , , , , "Multi-Bit Compare Slides Generated."
    
    
    
EndOfSub:

    Sheets(1).Select
    
    ppApp.WindowState = ppWindowNormal
    ppApp.Presentations(ppApp.Presentations.Count).Windows(1).View.GotoSlide 1
    
    frmStatus.cmdAbort.Enabled = False
    frmStatus.cmdClose.Enabled = True
    
    If Sheets("Bit_Compare").Shapes("chkShow").ControlFormat.Value <> 1 Then ppApp.Presentations(ppApp.Presentations.Count).Windows(1).WindowState = ppWindowNormal     'ppApp.WindowState = ppWindowNormal
    
    If gblnAbort Then UpdateStatus "", , , , , , "Process Aborted."
        
        
    'ActiveSheet.Protect
    ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Slide Blank")
    Set ppSlide = ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count)
    ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Conclusion"
    'write into the configuration file later
    ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
    ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "StrataBlade Changes Build Notes 2020.3.0 on Slide 2 1598771"
    ppApp.ActivePresentation.Slides.AddSlide ppApp.ActivePresentation.Slides.Count + 1, GetLayout(ppApp.ActivePresentation, "Title Only")
    ppApp.ActivePresentation.Slides(ppApp.ActivePresentation.Slides.Count).Shapes.Title.TextFrame.TextRange.Text = "Conclusion"
'    strDir = strDir & "Multi-Bit_" & Format(Now, "YYYY-MM-DD_hh.nn.ss") & ".pptx"
    ppApp.Presentations(ppApp.Presentations.Count).SaveAs ReadIniFileString("merge_template", "dirct")
    ppApp.Presentations(ppApp.Presentations.Count).Close
    ppApp.Quit
    gblnAbort = False

    Exit Sub
    
ErrHandler:

    If Err = -2147188160 Then 'ShapeRange (unknown member) : Integer out of range. 0 is not in the valid range of 1 to 1.
        Debug.Print Err, Err.Description, Now()
'        Stop
'        Resume
        Resume Next
        
    ElseIf Err = 13 Then ' Type mismath
        Stop
        Resume
    ElseIf Err = 9 Then ' Subscript out of range
        Exit Sub
    ElseIf Err = -2147467259 Then '
        MsgBox "Multi-Bit Compare Slides Aborted.", vbInformation + vbOKOnly, "Aborted"
        GoTo EndOfSub
    ElseIf Err = 91 Then
        Debug.Print Err, Err.Description, Now()
        Resume Next
    ElseIf Err = 424 Then
        Debug.Print Err, Err.Description, Now()
        Resume Next
    ElseIf Err = 52 Then
        Debug.Print Err, Err.Description, Now()
        iSub = iSub + 1
        GoTo TryNextiSub
    Else
        Debug.Print Err, Err.Description, Now()
        Stop
        Resume
    End If
    

End Sub
