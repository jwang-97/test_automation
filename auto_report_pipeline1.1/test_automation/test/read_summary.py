import os, re, logging, time, argparse, pptx, pptx.util, csv, io, unicodedata, csv, math
from pickle import LONG
from goto import with_goto
import configparser
import win32com.client as win32
import xlwings as xw
from pptx import Presentation
import matplotlib.pyplot as plt
import numpy as np
from pptx.util import Cm
import pandas as pd

def RoundOff(Num, Dig):
    #Round off <Num> to <Dig> decimal places
    temp = math.pow(10, Dig)
    RoundOff = int(Num * temp) / temp
    return RoundOff

def ProcessCompTime(iBit, strFile, key, value) :
    sCost = []
    # Dim fs As scripting.FileSystemObject
    # Dim txtIn As scripting.TextStream ' scripting.textstream
    # Dim strLine As String, sCost(3) As Single
    if os.path.exists(strFile) != False:
        with open(strFile,'r',encoding='utf-8') as txtIn:
        # lines = des_file.readlines()
            lines = txtIn.read().splitlines()
        find_info = [ string for string in lines if "Cost Hour=" in string] # old way of doing contact
        if find_info != None:
            sCost.append(float(find_info[10:]))
            if sCost[0] < 0:
                sCost.append(float(find_info[10:]))
                key.append("Bit" + str(iBit) + "_CompTime")
                value.append("Err.Neg.#")
            sCost.append(int(sCost[0]))                               #hours
            sCost.append(int((sCost[0] - int(sCost[0])) * 60))        #minutes
            sCost.append(int((sCost[0] - int(sCost[0])) * 60 * 60))  #seconds
            key.append("Bit" + str(iBit) + "_CompTime")
            value.append(str("{:0>2d}".format(sCost[0]))+":"+str("{:0>2d}".format(sCost[1]))+":"+str("{:0>2d}".format(sCost[2])))
            # find_info = [ string for string in lines if "body(top):" in string]       
            # find_info = lines.index(find_info[0])
            # i = 0
            # while lines[find_info + i +1].split()[0] != "":    

    return ProcessCompTime, key, value
    
def ConvertToK(Num):
    
    #Converts to Kilos
    temp = float(Num) / 1000
    temp = RoundOff(temp, 1)
    temp2 = str(temp).strip()
    if temp2.find(".") == None :
        ConvertToK = temp2 # & "K"
    else:
        ConvertToK = temp2.split(".")[0] # & "K"
    return ConvertToK

def ProcessSummary(BitNum, Summary_File_Handler,blnRapid = False):#todo
    key = []
    value = []
    start_row_index = 4

    # If blnRapid Then
    #     Sheets("BIT" & BitNum & "_SUMMARY").Activate
    #     Used_Range_Address = ActiveSheet.UsedRange.Address
    #     Last_row = Range(Used_Range_Address).Rows.Count

    #     GoTo ProcessBit
    # End If
    
    #Current_Workbook = ActiveWorkbook.Name
    #clear summary.csv
    with open(r'test.csv','a+',encoding='utf-8') as test:
        test.truncate(0)
    if os.path.exists(Summary_File_Handler) == False:
        logging.warning('Invalid summary path:' + Summary_File_Handler)
        return None
    
#Open Summary_regular file
    # If Range("ReadVer") = "old" Then
    #     Workbooks.OpenText FileName:=Summary_File_Handler, Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=True, Other:=True, TrailingMinusNumbers:=True
    #     strSummarySheet = ActiveWorkbook.Name
    # Else
    #     strSummarySheet = OpenText(Summary_File_Handler, Range("ReadVer"))
    # End If
    
    # Used_Range_Address = ActiveSheet.UsedRange.Address
    # Last_row = Range(Used_Range_Address).Rows.Count
    # ActiveSheet.UsedRange.Copy
    # Workbooks(Current_Workbook).Activate
    # Sheets("BIT" & BitNum & "_SUMMARY").Activate
    # Range(Used_Range_Address).PasteSpecial

#Close the summary file
    # Application.DisplayAlerts = False
    # If Range("ReadVer") = "old" Then
    #     Windows(strSummarySheet).Close
    #     Workbooks(Current_Workbook).Activate
        
    # Else
    #     Sheets(strSummarySheet).Delete
    #     Sheets("BIT" & BitNum & "_SUMMARY").Activate
#DeleteQueryTables
    # End If
    # Application.DisplayAlerts = True


    Contact_offset = 0
        
#Find Contact info
    with open(Summary_File_Handler,'r',encoding='utf-8') as summary_file:
        # lines = des_file.readlines()
        lines = summary_file.read().splitlines()
        find_info = [ string for string in lines if "contact:" in string] # old way of doing contact
        if find_info != None:
            find_info = [ string for string in lines if "body(top):" in string]       
            find_info = lines.index(find_info[0])
            i = 0
            while lines[find_info + i +1].split()[0] != "":
                i = i + 1            
            Contact = "On"
            Contact_offset = 3 + 2 * (i + 1) #3: 1 blank + 2 header rows, 2: sections (contact, friction), i: # FNFs, +1: accounts for the section headers
        else:
            Contact = "Off"
        
#Find Wear info
        find_info = [ string for string in lines if "control" in string]
        if find_info != None:
            find_info_num = lines.index(find_info[0])
            Control_num = lines[find_info_num].split().index("control")
            Control = lines[find_info_num].split()[Control_num-1]
        Wear_Analysis = "No"
    #     Set find_info = .Find("control,", LookIn:=xlValues, LookAt:=xlWhole)
    #     If Not find_info Is Nothing Then
    #         Range(find_info.Address).Offset(0, -1).Activate
    #         Control = ActiveCell
    #         Range(find_info.Address).Offset(0, 1).Activate
    #         If ActiveCell = "not" Then
    #             Wear_Analysis = "No"
    #         Else
    #             If ActiveCell = "calc" Then
    #                 Wear_Analysis = "Yes"
    #             Else
    #                 Wear_Analysis = "NA"
    #             End If
    #         End If
    #     End If
    # End With
    
    
    
    
#Find # of CPUs used
        find_info = [ string for string in lines if "CPU(s):" in string]
        if find_info != None:
            find_info = lines.index(find_info[0])
            key.append("Bit" + str(BitNum) + "_StaticCores")
            value.append(lines[find_info].split()[1])
            CPU_offset = 3
        else:
            key.append("Bit" + str(BitNum) + "_StaticCores")
            value.append("1")
            CPU_offset = 0
        
    #Get Comp Times
        A = ProcessCompTime(BitNum, Summary_File_Handler.replace(Summary_File_Handler.split('\\')[-1], "summary_variables"), key, value)
        key = A[1]
        value = A[2]
        a = A[0]
        
#Check if the project is a stinger Static project
        find_info = [ string for string in lines if "Output:" in string]
        if find_info != None:
            find_info = lines.index(find_info[0])
            if lines.index(find_info+2).split()[2] == "max" and lines.index(find_info+2).split()[3] == "min" and lines.index(find_info+2).split()[4] == "mean" or lines.index(find_info+2).split()[4] == "ave":
                #stinger static project
                ProcessStingerSummary(int(BitNum), Summary_File_Handler, WOB, RPM, ROP, TORQUE, DOC, ROCK, Conf_Pressure, TIF, RIF, CIF, RIF_CIF, RIF_CIF_Normal, BETA, find_count, blnRapid)#todo
                key.append("Bit" + str(BitNum) + "_StingerStatic")
                value.append("Yes")
                blnStingerStatic = True
                
                ProcessSummary = "Stinger_Static"
                GoTo EnterSummaryInfo#todo
            
            else:
                key.append("Bit" + str(BitNum) + "_StingerStatic")
                value.append("No")
                key.append("Bit" + str(BitNum) + "_StaticRevs")
                value.append("-")
        else:
            key.append("Bit" + str(BitNum) + "_StingerStatic")
            value.append("No")
            key.append("Bit" + str(BitNum) + "_StaticRevs")
            value.append("-")
    
    
    
    
    
    
    
    
        find_count = 0
        cutterdata_row = " "
        find_case = [ string for string in lines if "Output:" in string]
        if find_info != None:
            find_count = len(find_case)
        case_start_row = [""] * find_count
        cutter_start_row = [""] * find_count
        TIF = [""] * find_count
        cutter_end_row = [""] * find_count
        
        case_name = [""] * find_count
        ROP = [""] * find_count
        WOB = [""] * find_count
        RPM = [""] * find_count
        HRS = [""] * find_count
        DOC = [""] * find_count
        ROCK = [""] * find_count
        Case_Inputs = [""] * find_count
        Conf_Pressure = [""] * find_count
        CIF = [""] * find_count
        RIF_CIF = [""] * find_count
        RIF_CIF_Normal = [""] * find_count
        BETA = [""] * find_count
        TORQUE = [""] * find_count
        FOOTAGE = [""] * find_count
        RIF = [""] * find_count
        ROP_TORQUE = [""] * find_count
        
        
    # ReDim case_start_row(1 To find_count) As String
    # ReDim cutter_start_row(1 To find_count) As String
    # ReDim cutter_end_row(1 To find_count) As String
    # ReDim case_name(1 To find_count) As String
    # ReDim ROP(1 To find_count) As Double
    # ReDim WOB(1 To find_count) As Double
    # ReDim RPM(1 To find_count) As Double
    # ReDim HRS(1 To find_count) As Double
    # ReDim DOC(1 To find_count) As Double
    # ReDim ROCK(1 To find_count) As String
    # ReDim Case_Inputs(1 To find_count) As String
    # ReDim Conf_Pressure(1 To find_count) As Integer
    # ReDim FOOTAGE(1 To find_count) As Double
    # ReDim TIF(1 To find_count) As Double
    # ReDim CIF(1 To find_count) As Double
    # ReDim RIF(1 To find_count) As Double
    # ReDim RIF_CIF(1 To find_count) As Double
    # ReDim RIF_CIF_Normal(1 To find_count) As Double
    # ReDim BETA(1 To find_count) As Double
    # ReDim TORQUE(1 To find_count) As Double
    # ReDim ROP_TORQUE(1 To find_count) As Double
        i = 0
        find_case = [ string for string in lines if "percent" in string]
        if find_info != None:
            firstAddress = lines.index(find_info[0])
            while i < find_count:
                case_start_row[i] = lines.index(find_info[i])
                i = i + 1

        i = 0
        find_case = [ string for string in lines if "fn" in string]
        if find_info != None:
            firstAddress = lines.index(find_info[0])
            while i < find_count:
                cutter_start_row[i] = lines.index(find_info[0]) + 3
                i = i + 1
        i = 1
        while i <= find_count - 1:
            cutter_end_row[i] = case_start_row[i + 1] - 2 - Contact_offset
            i = i + 1
        cutter_end_row[i] = len(lines) - 2 - Contact_offset - CPU_offset
        
        i=1
        while i <= find_count:
            # key.append("A" + str(case_start_row[i]) + ":Z" + str(cutter_end_row[i]))
            find_para = [ string for string in lines if "percent" in string]
            if find_para != None:
                find_para = lines.index(find_para[0])
                TIF[i] = lines[find_para].split()[3]
            find_para = [ string for string in lines if "rif/wob" in string]
            if find_para != None:
                find_para = lines.index(find_para[0])
                RIF[i] = lines[find_para].split()[1]
            find_para = [ string for string in lines if "cif/wob" in string]
            if find_para != None:
                find_para = lines.index(find_para[0])
                CIF[i] = lines[find_para].split()[1]
            find_para = [ string for string in lines if "rif/cif" in string]
            if find_para != None:
                find_para = lines.index(find_para[0])
                RIF_CIF[i] = lines[find_para].split()[1]
                if RIF_CIF[i] > 1:
                    RIF_CIF_Normal[i] = 1 / float(RIF_CIF[i])
                else:
                    RIF_CIF_Normal[i] = RIF_CIF[i]
            find_para = [ string for string in lines if "b" in string]
            para_inum = 0
            if find_para != None:#todo confirm the logic
                while lines[find_para_num].split()[1].isdigit() == False and para_inum < len(find_para):
                    find_para_num = lines.index(find_para[para_inum])
                    para_inum += 1
                BETA[i] = lines[find_para_num].split()[1]
            find_para = [ string for string in lines if "weight-on-bit" in string]
            if find_para != None:
                find_para_num = lines.index(find_para[0])
                WOB[i] = lines[find_para_num].split()[1]
            find_para = [ string for string in lines if "torque" in string]
            if find_para != None:
                find_para_num = lines.index(find_para[0])
                TORQUE[i] = lines[find_para_num].split()[2]
            find_para = [ string for string in lines if "penetration" in string]
            if find_para != None:
                find_para_num = lines.index(find_para[0])
                ROP[i] = lines[find_para_num].split()[3]
                if TORQUE[i] == 0 :
                    ROP_TORQUE[i] = 0
                else:
                    ROP_TORQUE[i] = RoundOff(ROP[i] / TORQUE[i], 4)
            find_para = [ string for string in lines if "rotary" in string]
            if find_para != None:
                find_para_num = lines.index(find_para[0])
                RPM[i] = lines[find_para_num].split()[3]
            find_para = [ string for string in lines if "rotating" in string]
            if find_para != None:
                find_para_num = lines.index(find_para[0])
                HRS[i] = lines[find_para_num].split()[3]
            find_para = [ string for string in lines if "advance" in string]
            if find_para != None:
                find_para_num = lines.index(find_para[0])
                DOC[i] = lines[find_para_num].split()[2]
            find_para = [ string for string in lines if "cutter/rock" in string]
            if find_para != None:
                find_para_num = lines.index(find_para[0])
                ROCK[i] = lines[find_para_num+1].split()[0]
                ROCK[i] = ROCK[i][ROCK[i].find("_"): ]
                ROCK[i] = ROCK[i][ROCK[i].find("_"): ]
                Conf_Pressure[i] = ROCK[i][ROCK[i].find("_"): ]
                ROCK[i] = ROCK[i].split("_")[0]
            if i == 1:
                FOOTAGE[i] = ROP[i]*HRS[i]
            else:
                FOOTAGE[i] = FOOTAGE[i-1] + (ROP[i]*(HRS[i] - HRS[i-1]))
            i+=1

# EnterSummaryInfo:

        key.append("Bit" + str(BitNum) + "_Cases")
        value.append(find_count)
        key.append("Bit" + str(BitNum) + "_Contact")
        value.append(Contact)
        key.append("Bit" + str(BitNum) + "_WearProj")
        value.append(Wear_Analysis)
        key.append("Bit" + str(BitNum) + "_Hours")
        value.append(HRS[find_count])
        # Sheets("INFO").Activate
        # If InStr(1, Build_Date, "/") > 0 Then Build_Date = Format(Build_Date, "YYYYMMDD")
        # Range("Bit" & BitNum & "_Build") = Mid(Build_Date, 5, 2) & "/" & Right(Build_Date, 2) & "/" & Left(Build_Date, 4)
        # Range("Bit" & BitNum & "_Control") = Control
        # Sheets("SUMMARY").Activate
        if BitNum == 1 and Control == "wob" :
            key.append("Control_Info")
            value.append("Instantaneous ROP (ft/hr)")
            Case_details = "WOB_"
        elif BitNum == 1 and Control == "rop" :
            key.append("Control_Info")
            value.append("WOB (Klb)")
            Case_details = "ROP_"
        
        if Wear_Analysis == "Yes" :
            Case_details = Case_details + "HRS_"
        Case_details = Case_details + "RPM_ROCK_CONFINING PRESSURE"
        key.append("Case_Info")
        value.append("Case Info (" + Case_details + ")")
        # the row_index means the start lineup of data
        # row_index = key.index("Case_Info") + 2
        case_column_index=key.index("Case_Info")
        # row_index = Range("Case_Info").row + 2#todo
        control_column_index = key.index("Control_Info")
        ROP_column_index = key.index("ROP_Compare")
        DOC_column_index = key.index("DOC_info")
        Cum_Hrs_index = key.index("Cumulative_Hours")
        Cum_Foot_index = key.index("Cumulative_Footage")
        Torque_index = key.index("TORQUE_info")
        TQ_Compare_index = key.index("TQ_Compare")
        TIF_index = key.index("TIF_info")
        RIF_index = key.index("RIF_info")
        CIF_index = key.index("CIF_info")
        RIF_CIF_index = key.index("RIF_CIF_info")
        RIF_CIF_Normal_index = key.index("RIF_CIF_Normal_info")
        BETA_index = key.index("BETA_info")
        ROP_TORQUE_index = key.index("ROP_TORQUE_info")
        ROP_TORQUE_Compare_index = key.index("ROP_TORQUE_Compare")
        
    
        bit_offset = BitNum
    
    # 'If BitNum = 1 Then
    # '    Cells(row_index + 1, case_column_index + 1).Select
    # '    Range(ActiveCell, ActiveCell.Offset(0, 0)).Name = "Cases"
    # 'End If
    # 'If Not BitNum = 1 Then
    # '    Cells(row_index + 1, ROP_column_index + BitNum - 1).Select
    # '    Range(ActiveCell, ActiveCell.Offset(0, 0)).Name = "Bit" & BitNum & "_ROP"
    # 'End If
        
        row_index = 0
        i = 0
        summmary_case_title = ["Case Info (" + Case_details + ")","Instantaneous ROP (ft/hr)", "Bit Advance (in/rev)", r"% change in Instantaneous ROP",
                              "Cumulative Time (Hours)", "Cumulative Footage (ft)", "Torque (ft-lbf)", r"% change in Instantaneous Torque",
                              r"Total imbalance force (% of WOB)", r"Circumferential imbalance force (% of WOB)", r"Radial imbalance force (% of WOB)",
                              "Beta Angle", "RIF/CIF", "Normalized IF Ratio", "ROP/Torque", r"% Change in ROP/Torque"]#todo
        summmary_case_list = []

        while i <= find_count:
        # build the count_list in summary sheet
        #     Cells(row_index + i, case_column_index).Value = i
        #     CellFormat1 (Cells(row_index + i, case_column_index).Address)
        #     Cells(row_index + i, control_column_index).Value = i
        #     CellFormat1 (Cells(row_index + i, control_column_index).Address)
        #     Cells(row_index + i, ROP_column_index).Value = i
        #     CellFormat1 (Cells(row_index + i, ROP_column_index).Address)
        #     Cells(row_index + i, DOC_column_index).Value = i
        #     CellFormat1 (Cells(row_index + i, DOC_column_index).Address)
        #     Cells(row_index + i, Cum_Hrs_index).Value = i
        #     CellFormat1 (Cells(row_index + i, Cum_Hrs_index).Address)
        #     Cells(row_index + i, Cum_Foot_index).Value = i
        #     CellFormat1 (Cells(row_index + i, Cum_Foot_index).Address)
        #     Cells(row_index + i, Torque_index).Value = i
        #     CellFormat1 (Cells(row_index + i, Torque_index).Address)
        #     Cells(row_index + i, TQ_Compare_index).Value = i
        #     CellFormat1 (Cells(row_index + i, TQ_Compare_index).Address)
        #     Cells(row_index + i, TIF_index).Value = i
        #     CellFormat1 (Cells(row_index + i, TIF_index).Address)
        #     Cells(row_index + i, RIF_index).Value = i
        #     CellFormat1 (Cells(row_index + i, RIF_index).Address)
        #     Cells(row_index + i, CIF_index).Value = i
        #     CellFormat1 (Cells(row_index + i, CIF_index).Address)
        #     Cells(row_index + i, RIF_CIF_index).Value = i
        #     CellFormat1 (Cells(row_index + i, RIF_CIF_index).Address)
        #     Cells(row_index + i, RIF_CIF_Normal_index).Value = i
        #     CellFormat1 (Cells(row_index + i, RIF_CIF_Normal_index).Address)
        #     Cells(row_index + i, BETA_index).Value = i
        #     CellFormat1 (Cells(row_index + i, BETA_index).Address)
        #     Cells(row_index + i, ROP_TORQUE_index).Value = i
        #     CellFormat1 (Cells(row_index + i, ROP_TORQUE_index).Address)
        #     Cells(row_index + i, ROP_TORQUE_Compare_index).Value = i
        #     CellFormat1 (Cells(row_index + i, ROP_TORQUE_Compare_index).Address)
        
            if Control == "wob" :
                #Cells(row_index + i, control_column_index + bit_offset).Value = Str(ROP(i)) & " (" & Str(DOC(i)) & ")"
                summmary_case_list[1].append(ROP[i])
                Case_details = ConvertToK(WOB(i)) + "_"
    #ROP Compare
                if BitNum != 1 and summmary_case_list[1][0] != "" :
                    baseline_rop = summmary_case_list[1][0]
                    if baseline_rop == 0 :
                        summmary_case_list[3] = "-"
                    else:
                        ROP_Change = (ROP[i] - baseline_rop) / baseline_rop * 100
                        
                        Cells(row_index + i, ROP_column_index + bit_offset - 1).Value = RoundOff(ROP_Change, 2)
                    End If
                    CellFormat1 (Cells(row_index + i, ROP_column_index + bit_offset - 1).Address)
                elif BitNum == 1 :
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

Function ProcessStingerSummary(iBit As Integer, strFile As String, WOB() As Double, RPM() As Double, ROP() As Double, TORQUE() As Double, DOC() As Double, ROCK() As String, Conf_Pressure() As Integer, TIF() As Double, RIF() As Double, CIF() As Double, RIF_CIF() As Double, RIF_CIF_Normal() As Double, BETA() As Double, iCases As Integer, Optional blnRapid As Boolean) As String#todo

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
