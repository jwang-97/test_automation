import os, re, logging, time, argparse, pptx, pptx.util, csv, io, unicodedata, csv, math
from pickle import LONG
import configparser
import win32com.client as win32
# import xlwings as xw
from pptx import Presentation
import matplotlib.pyplot as plt
# import numpy as np
from pptx.util import Cm
import pandas as pd

logging.basicConfig(format='%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s',
                    level=logging.DEBUG)

def check_bit_in(strFile, strType):
    logging.debug(os.path.dirname(strFile) + r"\bit.in")
    if os.path.exists(os.path.dirname(strFile) + r"\bit.in") == False:
        return None
    with open(os.path.dirname(strFile) + r"\bit.in", "r", encoding="gbk", errors="ignore") as text_file:
        # lines = text_file.readlines()
        # new way to ignore newlines
        lines = text_file.read().splitlines()
        first_line= lines[0]
        # last_line= lines[-1]
        # txtIn = csv.reader(text_file)
        if first_line == "{": #new file format
            for line in lines:
                line_num = lines.index(line)
                if strType == "Name":
                    if re.search('"DesName": {', line) != None:
                        str_line = lines[line_num + 2]
                        check_bit_in = str(str_line.split(":")[1]).strip()
                        check_bit_in = check_bit_in.replace( '"' , '')
                        return check_bit_in
                elif strType == "Diameter":
                    if re.search('"Diameter": {', line) != None:
                        str_line = lines[line_num + 2]
                        check_bit_in = str(str_line.split(":")[1]).strip()
                        return check_bit_in
        else: #old file format
            for line in lines:
                line_num = lines.index(line)
                line = line
                if strType == "Name":
                    if re.search('[cfb]', line) != None or re.search('[des]', line) != None:
                        str_line = lines[line_num + 1]
                        return str_line
                elif strType == "Diameter":
                    if re.search('[diameter]', line) != None:
                        str_line = lines[line_num + 1]
                        return str_line

def check_ramer_in(strFile, strType):
    if os.path.exists(os.path.dirname(strFile) + "reamer.in") == False:
        return None
    with open(os.path.dirname(strFile) + "reamer.in", "r", encoding="gbk", errors="ignore") as text_file:
        lines = text_file.read().splitlines()
        first_line= lines[0]
        if first_line == "{": #new file format
            for line in lines:
                line_num = lines.index(line)
                if strType == "Name":
                    if re.search('"DrillFileName": {', line) != None:
                        str_line = lines[line_num + 2]
                        check_bit_in = str(str_line.split(":")[1]).strip()
                        check_bit_in = check_bit_in.replace( '"' , '')
                        return check_bit_in
                elif strType == "Diameter":
                    if re.search('"DrillDiameter": {', line) != None:
                        str_line = lines[line_num + 2]
                        check_bit_in = str(str_line.split(":")[1]).strip()
                        return check_bit_in

def RoundOff(Num, Dig):
    #Round off <Num> to <Dig> decimal places
    temp = math.pow(10, Dig)
    RoundOff = int(Num * temp) / temp
    return RoundOff

def CutterDia(Cutter, blnSize = False):
    
    #Finds the cutter diameter based on the cutter size
    if Cutter[0] == "G": #Axe Cutter
        Cutter_Size = int(Cutter[1: 4])
    elif Cutter[0] == "P": #PX/G188/Axe TR Cutter
        Cutter_Size = int(Cutter[1: 4])
    elif Cutter[0] == "J" or Cutter[0] == "K": #PX/Axe Arrow Cutter
        Cutter_Size = int(Cutter[1: 4])
    elif Cutter[0] == "D" or Cutter[0] == "Q": #Delta/Hyper Cutter
        Cutter_Size = int(Cutter[1: 4])
    elif Cutter[0] == "M": #Echo Cutter
        Cutter_Size = int(Cutter[1: 4])
    else:
        if len(Cutter) == 3:
            Cutter_Size = int(Cutter[0])
        else:
            # print(Cutter[0:2])
            Cutter_Size = int(Cutter[0:2])

    if blnSize:
        CutterDia = Cutter_Size
        return CutterDia
    
    if Cutter_Size == 6:
        CutterDia = 0.236
    elif Cutter_Size == 9:
        CutterDia = 0.371
    elif Cutter_Size == 11:
        CutterDia = 0.44
    elif Cutter_Size == 13:
        CutterDia = 0.529
    elif Cutter_Size == 16:
        CutterDia = 0.625
    elif Cutter_Size == 19:
        CutterDia = 0.75
    elif Cutter_Size == 22:
        CutterDia = 0.866
    else:
        if Cutter == "" :
            logging.debug('if stinger')
        else:
            CutterDia = Cutter
    return CutterDia

def GaugePadNom(sBitSize):
    #Returns the nominal gauge pad radius for a given bit size. Table taken from design book 60088589 Rev- on 2017-04-17
    
    if sBitSize <= 6.75 :
        #Bit OD - 0.015
        sFinishGP = 0.015
        
    elif sBitSize <= 9 :
        #Bit OD - 0.015
        sFinishGP = 0.015
        
    elif sBitSize <= 13.75 :
        #Bit OD - 0.025
        sFinishGP = 0.025
        
    elif sBitSize <= 17.5 :
        #Bit OD - 0.025
        sFinishGP = 0.025
        
    else: #17-17/32 and larger
        #Bit OD - 0.035
        sFinishGP = 0.035
        
    GaugePadNom = (sBitSize - sFinishGP) / 2
    return float(GaugePadNom)


def GetGaugePadInfo(iBit, sBitSize, Bit_File_Handler):

    xl = win32.Dispatch("Excel.Application")
    # xl.Visible = 0
    str_path = os.getcwd()
    try:
        with open(Bit_File_Handler, "r", encoding="gbk", errors="ignore") as des_file:
            # xl = win32.Dispatch("Excel.Application")
            xlbook = xl.Workbooks.Open(Filename = str_path + r"\test_automation\test\Book3.xlsx", ReadOnly = False)
            # sheet = xlbook.Worksheets("BIT" + str(BitNum))
            sheet = xlbook.Worksheets("INFO")
            # lines = des_file.readlines()
            lines = des_file.read().splitlines()
            find_BasicData = [ string for string in lines if "#padNo(1..)" in string]
            # fileHeader = []
            # d1 = []
            # find_GP = lines.index(find_BasicData[0])
            if find_BasicData != None:
                find_GP = lines.index(find_BasicData[0])
                iRowEnd = 0
            # for line in lines:
            #     line_num = lines.index(line)
            #     if re.search('#padNo(1..)', line) != None:
            #         find_GP = line_num
            #         break
        # if find_GP != None:
        #     iRowEnd = 0
                if sBitSize > 0:#todo wrong
                    while lines[iRowEnd + 1] != "[end_section_simple_pad]":
                        iRowEnd = iRowEnd + 1
                else:
                    while lines[iRowEnd + 1].split()[1] != "":
                        iRowEnd = iRowEnd + 1
                if iRowEnd == 0:
                    des_file.close()
                    return None
                # iRow: times of loop, equal to rows between "#padNo(1..)" and "[end_section_simple_pad]"
                iRow = iRowEnd - find_GP
                sRad1 = []
                sBeta1 = []
                sLen1 = []
                sTaper1 = []
                sRad2 = []
                sLen2 = []
                sBeta2 = []
                sTaper2 = []
                sWidth = []
                i = 1
                while i <= iRow:
                    sRad1.append(lines[find_GP + i].split()[3])
                    sLen1.append(lines[find_GP + i].split()[7])
                    sBeta1.append(lines[find_GP + i].split()[8])
                    sTaper1.append(lines[find_GP + i].split()[9])
                    sRad2.append(lines[find_GP + i].split()[11])
                    sLen2.append(lines[find_GP + i].split()[10])
                    sBeta2.append(lines[find_GP + i].split()[12])
                    sTaper2.append(lines[find_GP + i].split()[13])
                    sWidth.append(lines[find_GP + i].split()[6])
                    i += 1
            #averaging all values for all blades
                # i = 1
                sVals = [0]*9
                i = 1
                while i <= iRow:
                    sVals[0] = sVals[0] + float(sRad1[i -1])
                    sVals[1] = sVals[1] + float(sLen1[i -1])
                    sVals[2] = sVals[2] + float(sBeta1[i -1])
                    sVals[3] = sVals[3] + float(sTaper1[i -1])
                    sVals[4] = sVals[4] + float(sRad2[i -1])
                    sVals[5] = sVals[5] + float(sLen2[i -1])
                    sVals[6] = sVals[6] + float(sBeta2[i -1])
                    sVals[7] = sVals[7] + float(sTaper2[i -1])
                    sVals[8] = sVals[8] + float(sWidth[i -1])
                    i += 1
                
                i = 0
                while i <= len(sVals) - 1:
                    sVals[i] = float(sVals[i]) / iRow
                    i += 1
                des_file.close()
                # xl = win32.Dispatch("Excel.Application")
                # xlbook = xl.Workbooks.Open(Filename = str + r"\test_automation\test\Book3.xlsx", ReadOnly = False)
                # # sheet = xlbook.Worksheets("BIT" + str(BitNum))
                # sheet = xlbook.Worksheets("INFO")
                sheet.Range("Bit" + str(iBit) + "_GPRad").Value = RoundOff(sVals[0], 4)
                sheet.Range("Bit" + str(iBit) + "_GPLen").Value = RoundOff(sVals[1] + sVals[5], 2)
                sheet.Range("Bit" + str(iBit) + "_GPWidth").Value = RoundOff(sVals[8], 4)
                
                
                #Calc Nominal Length
                if sBitSize > 0 and GaugePadNom(sBitSize) == sVals[0] :#section 1 is nominal(todo)
                    if sVals[3] == 0 :
                        if sVals[7] == 0 and sVals[5] > 0 :
                            sheet.Range("Bit" + str(iBit) + "_GPNomLen").Value = RoundOff(sVals[1] + sVals[5], 2)
                        else:
                            sheet.Range("Bit" + str(iBit) + "_GPNomLen").Value = RoundOff(sVals[1], 2)
                    else:
                        sheet.Range("Bit" + str(iBit) + "_GPNomLen").Value = 0
                else:
                    sheet.Range("Bit" + str(iBit) + "_GPNomLen").Value = 0
                if sVals[3] > 0 : #section 1 is tapered
                    if sVals[5] > 0 :
                        if sVals[7] > 0 :
                            if sVals[3] == sVals[7] :
                                sheet.Range("Bit" + str(iBit) + "_GPTaperLen").Value = RoundOff(sVals(1) + sVals(5), 2)
                                if sVals[3] == sVals[7] :
                                    sheet.Range("Bit" + str(iBit) + "_GPTaperAng").Value = RoundOff(sVals[3], 2)
                                else:
                                    sheet.Range("Bit" + str(iBit) + "_GPTaperAng").Value = "'" + str(RoundOff(sVals[3], 2)) + " / " + str(RoundOff(sVals[7], 2))
                            else:
                                sheet.Range("Bit" + str(iBit) + "_GPTaperLen").Value = "'" + str(RoundOff(sVals[1], 2)) + " / " + str(RoundOff(sVals[5], 2))
                                sheet.Range("Bit" + str(iBit) + "_GPTaperAng").Value = "'" + str(RoundOff(sVals[3], 2)) + " / " + str(RoundOff(sVals[7], 2))
                        else:
                            sheet.Range("Bit" + str(iBit) + "_GPTaperLen").Value = RoundOff(sVals[1], 2)
                            sheet.Range("Bit" + str(iBit) + "_GPTaperAng").Value = RoundOff(sVals[3], 2)
                    else:
                        sheet.Range("Bit" + str(iBit) + "_GPTaperLen").Value = RoundOff(sVals[1], 2)
                        sheet.Range("Bit" + str(iBit) + "_GPTaperAng").Value = RoundOff(sVals[3], 2)
                else:
                    if sVals[5] > 0 and sVals[7] > 0 : #section 2 has length and it's tapered
                        sheet.Range("Bit" + str(iBit) + "_GPTaperLen").Value = RoundOff(sVals[5], 2)
                        sheet.Range("Bit" + str(iBit) + "_GPTaperAng").Value = RoundOff(sVals[7], 2)
                    else: # section 2 eith has no length or it's not taper
                        sheet.Range("Bit" + str(iBit) + "_GPTaperLen").Value = 0
                        sheet.Range("Bit" + str(iBit) + "_GPTaperAng").Value = 0

                if sVals[4] == 0 : #no undercut from section 1
                    if sBitSize > 0 and GaugePadNom(sBitSize) > sVals[0] :
                        sheet.Range("Bit" + str(iBit) + "_GPUcutLen").Value = RoundOff(sVals[1] + sVals[5], 2)
                        sheet.Range("Bit" + str(iBit) + "_GPUcutAmt").Value = RoundOff(GaugePadNom(sBitSize) - sVals[0], 2)
                    else:
                        sheet.Range("Bit" + str(iBit) + "_GPUcutLen").Value = 0
                        sheet.Range("Bit" + str(iBit) + "_GPUcutAmt").Value = 0
                    
                else: #Section 2 has an undercut
                    sheet.Range("Bit" + str(iBit) + "_GPUcutLen").Value = RoundOff(sVals[5], 2)
                    sheet.Range("Bit" + str(iBit) + "_GPUcutAmt").Value = RoundOff(sVals[4], 2)

                #Beta
                if sVals[5] == 0 : #section 2 has no length, only use beta from section 1
                    sheet.Range("Bit" + str(iBit) + "_GPBeta").Value = RoundOff(sVals[2], 2)
                elif sVals[2] == sVals[6] :
                    sheet.Range("Bit" + str(iBit) + "_GPBeta").Value = RoundOff(sVals[2], 2)
                else:
                    sheet.Range("Bit" + str(iBit) + "_GPBeta").Value = "'" + str(RoundOff(sVals[2], 2)) + " / " + str(RoundOff(sVals[6], 2))

            else: #if is a DES but has no GaugePad, or is a CFB
                with open(Bit_File_Handler, "r", encoding="gbk", errors="ignore") as des_file:
                    lines = des_file.read().splitlines()
                    find_BasicData = [ string for string in lines if "[BasicData]" in string]
                    if find_BasicData != None:
                        # find_GP = lines.index(find_BasicData[0])
                        sheet.Range("Bit" + str(iBit) + "_GPBeta").Value = "-"
                        sheet.Range("Bit" + str(iBit) + "_GPLen").Value = "-"
                        sheet.Range("Bit" + str(iBit) + "_GPNomLen").Value = "-"
                        sheet.Range("Bit" + str(iBit) + "_GPTaperLen").Value = "-"
                        sheet.Range("Bit" + str(iBit) + "_GPTaperAng").Value = "-"
                        sheet.Range("Bit" + str(iBit) + "_GPUcutLen").Value = "-"
                        sheet.Range("Bit" + str(iBit) + "_GPUcutAmt").Value = "-"
                        sheet.Range("Bit" + str(iBit) + "_GPBeta").Value = "-"
                        sheet.Range("Bit" + str(iBit) + "_GPWidth").Value = "-"

        # xlbook = xl.Workbooks.Open(Filename = str + r"\test_automation\test\Book3.xlsx", ReadOnly = False)
        # sheet = xlbook.Worksheets("BIT" + str(BitNum))
        # xl.Sheets("BIT" + iBit + "_CUTTER_FILE").Activate
            xl.DisplayAlerts = False
            xl.ActiveWorkbook.Save()
            xl.ActiveWorkbook.Close()
            # xl.Quit()
        return None
    except ValueError:
        logging.warning('Open failed: ' + Bit_File_Handler)
        return None
    
    
    
def process_des(BitNum, Bit_File_Handler):
    # Initialize data variables
    # Dim radius() As Double
    # Dim ang_arnd() As Double
    # Dim height() As Double
    # Dim prf_ang() As Double
    # Dim bk_rake() As Double
    # Dim sd_rake() As Double
    # Dim cut_dia() As Variant
    # Dim blade() As Double
    # Dim size_type() As String
    # Dim ProfL() As Double
    # Dim exposure() As Double
    # Dim Blade_Angle() As Double
    # Dim Tag() As Integer
    # Dim roll() As Double
    # Dim hMax As Double
    hMax = ""
    # Dim DeltaProfAng() As Double
    # Dim Tag360() As Integer
    # Dim TagCreo() As Integer
    
    # On Error GoTo ErrHandler
    
    ID_Radius = "-"
    Cone_Angle = "-"
    Shoulder_Angle = "-"
    Nose_Radius = "-"
    Middle_Radius = "-"
    Shoulder_Radius = "-"
    Gauge_Length = "-"
    Profile_Height = "-"
    Nose_Location_Diameter = "-"
    try:
        with open(Bit_File_Handler, "r", encoding="gbk", errors="ignore") as des_file:
            lines = des_file.read().splitlines()
            find_BasicData = [ string for string in lines if "[BasicData]" in string]
            line_num = lines.index(find_BasicData[0])
            if find_BasicData != None:
                Bit_Type = lines[line_num + 1]
                Bit_Size = lines[line_num + 2]
                Blade_Count = lines[line_num + 3]
            find_Profile_Type = [ string for string in lines if "[profile_choice]" in string]
            line_num = lines.index(find_Profile_Type[0])
            if find_Profile_Type != None:
                    Bit_Profile = lines[line_num + 1]
            find_Profile_Info = [ string for string in lines if "[Profile"+ Bit_Profile +']' in string]
            line_num = lines.index(find_Profile_Info[0])
            if find_Profile_Info != None:
                i = 1
                if Bit_Profile == "AAAG":
                    ID_Radius = RoundOff(float(lines[line_num + i]), 3)
                    i = i + 1
                else:
                    Cone_Angle = RoundOff(float(lines[line_num + i]), 1)
                    i = i + 1
                if Bit_Profile == "LALAG":
                    Shoulder_Angle = RoundOff(float(lines[line_num + i]), 1)
                    i = i + 1
                if Bit_Profile == "LAAG" or Bit_Profile == "LAAAG" or Bit_Profile == "LALAG" or Bit_Profile == "AAAG":
                    Nose_Radius = RoundOff(float(lines[line_num + i]), 3)
                    i = i + 1
                if Bit_Profile == "LAAAG":
                    Middle_Radius = RoundOff(float(lines[line_num + i]), 3)
                    i = i + 1
                Shoulder_Radius = RoundOff(float(lines[line_num + i]), 3)
                i = i + 1
                Gauge_Length = RoundOff(float(lines[line_num + i]), 3)
                i = i + 1
                if Bit_Profile == "LAAAG":
                    Profile_Height = RoundOff(float(lines[line_num + i]), 3)
                    i = i + 1
                if Bit_Profile == "LAAG" or Bit_Profile == "LAAAG" or Bit_Profile == "LALAG" or Bit_Profile == "AAAG":
                    Nose_Location_Diameter = RoundOff(float(lines[line_num + i]), 3)
                    i = i + 1
                Tip_Grind = RoundOff(float(lines[line_num + i]), 3) - float(Bit_Size)
                if Bit_Profile == "LAG":
                    Nose_Radius = Shoulder_Radius
                    Nose_Location_Diameter = RoundOff(((Bit_Size + Tip_Grind) / 2 - Shoulder_Radius) * 2, 3)
            find_actual_positions = [ string for string in lines if '[positions]' in string]
            line_num = lines.index(find_actual_positions[0])
            if find_actual_positions != None:
                Position_index = lines[line_num + 1]
            Position_length_index = line_num + 1
            find_positions = [ string for string in lines if '[cutters]' in string]
            line_num = lines.index(find_positions[0])
            if find_positions != None:
                Position_Count = lines[line_num + 1]
                # elif re.match('[positions]', line) != None:
                #     find_actual_positions = line_num
                #     Position_index = lines[line_num + 1]
                #     Position_length_index = line_num + 1
                # elif re.match('[cutters]', line) != None:
                #     find_positions = line_num
                #     Position_Count = lines[line_num + 1]
            row_index = line_num + 2
            pos_index = 1
            i = 1
            ProfL = []
            radius = []
            ang_arnd = []
            height = []
            prf_ang = []
            bk_rake = []
            sd_rake = []
            blade = []
            cut_dia = []
            size_type = []
            Tag = []
            DeltaProfAng = []
            Tag360 = []
            TagCreo = []
            roll = []
            exposure = []
            while i <= int(Position_Count):
                Temp_ProfL = lines[Position_length_index + i].split()[2]
                row_index = row_index + 1
                j = 1
                while j <= int(Blade_Count):
                    Cutter_exists = int(lines[row_index].split()[5])
                    if Cutter_exists == 1:
                        ProfL.append(float(Temp_ProfL))
                        radius.append(float(lines[row_index].split()[7]))
                        ang_arnd.append(float(lines[row_index].split()[9]))
                        height.append(float(lines[row_index].split()[8]))
                        prf_ang.append(float(lines[row_index].split()[10]))
                        bk_rake.append(float(lines[row_index].split()[11]))
                        sd_rake.append(float(lines[row_index].split()[12]))
                        blade.append(float(lines[row_index].split()[3]) + 1)
                        exposure.append(float(lines[row_index].split()[17]))
                        cut_dia.append(CutterDia(lines[row_index+1].split()[0]))
                        size_type.append(lines[row_index+1].split()[0])
                        Tag.append(int(lines[row_index+1].split()[2]))
                        DeltaProfAng.append(float(lines[row_index].split()[18]))
                        Tag360.append(int(lines[row_index+1].split()[3]))
                        TagCreo.append(int(lines[row_index+1].split()[1]))
                        # print(len(lines[row_index+1].split()))
                        if len(lines[row_index+1].split()) > 4:
                            roll.append(lines[row_index+1].split()[4])
                        else:
                            roll.append(0)
                        pos_index = pos_index + 1
                    row_index = row_index + 2
                    j += 1
                i += 1
        des_file.close()
        GetGaugePadInfo(int(BitNum), float(Bit_Size), Bit_File_Handler)
    except ValueError:
        logging.warning('Open failed: ' + Bit_File_Handler)
        return None
    
    #Write the INFO page(todo)
    xl = win32.Dispatch("Excel.Application")
    str_path = os.getcwd()
    xlbook = xl.Workbooks.Open(Filename = str_path + r"\test_automation\test\Book3.xlsx", ReadOnly = False)
    sheet = xlbook.Worksheets("INFO")
    xl.Sheets("INFO").Activate
    row_index = sheet.Range("Cutting_Structure_Info").Row
    row_index = row_index + 2 + int(BitNum - 1) * 4
    sheet.Range(fXLNumberToLetter(sheet.Range("CS_Bit_Type").Column) + str(row_index).strip()).Value = Bit_Type
    sheet.Range(fXLNumberToLetter(sheet.Range("CS_Bit_Size").Column) + str(row_index).strip()).Value = Bit_Size
    sheet.Range(fXLNumberToLetter(sheet.Range("CS_Blade_Count").Column) + str(row_index).strip()).Value = Blade_Count
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Profile_Type").Column) + str(row_index).strip()).Value = Bit_Profile
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_ID_Radius").Column) + str(row_index).strip()).Value = ID_Radius
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Cone_Ang").Column) + str(row_index).strip()).Value = Cone_Angle
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Shou_Ang").Column) + str(row_index).strip()).Value = Shoulder_Angle
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Nose_Rad").Column) + str(row_index).strip()).Value = Nose_Radius
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Mid_Rad").Column) + str(row_index).strip()).Value = Middle_Radius
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Shou_Rad").Column) + str(row_index).strip()).Value = Shoulder_Radius
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Gauge_Len").Column) + str(row_index).strip()).Value = Gauge_Length
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Prof_Ht").Column) + str(row_index).strip()).Value = Profile_Height
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Nose_Loc").Column) + str(row_index).strip()).Value = Nose_Location_Diameter
    sheet.Range(fXLNumberToLetter(sheet.Range("Prof_Tip_Grind").Column) + str(row_index).strip()).Value = Tip_Grind
    
    A = ProcessCutterInfo(BitNum, radius, ang_arnd, height, prf_ang, bk_rake, sd_rake, cut_dia, size_type, blade, exposure, ProfL, Tag, roll, hMax, DeltaProfAng, Tag360, TagCreo, xlbook, xl)
    
    if sheet.Range("Y" + str(row_index).strip()).Value == "-" :
        sheet.Range("Y" + str(row_index).strip()).Value = RoundOff(hMax, 3)
    # xl.Sheets("HOME").Activate
    xl.DisplayAlerts = False
    xl.ActiveWorkbook.Save()
    xlbook.Save()
    xl.ActiveWorkbook.Close()
    # xl.Quit()
    
def fXLNumberToLetter(lNum):

    strLet = ''
    fXLNumberToLetter = ""
    # xl = win32.Dispatch("Excel.Application")
    if lNum > 16384 :
        logging.warning( "lNum too large. Abort.")
        return None
    
    if lNum > 676 : #At least greater than ZZ
        iPass = 3
    elif lNum > 26 : #At least greater than Z
        iPass = 2
    else:
        iPass = 1
    
    lRem = lNum
    
    while iPass > 0:
        if lRem % 26 == 0 :
            if iPass == 1 :
                iTest = 26
            else:
                # iTest = xl.WorksheetFunction.Floor((lRem - 1) / (26 ^ (iPass - 1)), 1)
                iTest = math.floor(float(lRem) / (math.pow(26, (iPass - 1))))
        else:
            # iTest = xl.WorksheetFunction.Floor(float(lRem) / (26 ^ (iPass - 1)), 1)
            iTest = math.floor(float(lRem) / (math.pow(26, (iPass - 1))))
        
        #Select Case iTest
        if iTest==1: 
            strLet = "A"
        elif iTest==2: 
            strLet = "B"
        elif iTest== 3: 
            strLet = "C"
        elif iTest== 4: 
            strLet = "D"
        elif iTest== 5: 
            strLet = "E"
        elif iTest== 6: 
            strLet = "F"
        elif iTest== 7: 
            strLet = "G"
        elif iTest== 8: 
            strLet = "H"
        elif iTest== 9: 
            strLet = "I"
        elif iTest== 10: 
            strLet = "J"
        elif iTest== 11: 
            strLet = "K"
        elif iTest== 12: 
            strLet = "L"
        elif iTest== 13: 
            strLet = "M"
        elif iTest== 14: 
            strLet = "N"
        elif iTest== 15: 
            strLet = "O"
        elif iTest== 16: 
            strLet = "P"
        elif iTest== 17: 
            strLet = "Q"
        elif iTest== 18: 
            strLet = "R"
        elif iTest== 19: 
            strLet = "S"
        elif iTest== 20: 
            strLet = "T"
        elif iTest== 21: 
            strLet = "U"
        elif iTest== 22: 
            strLet = "V"
        elif iTest== 23: 
            strLet = "W"
        elif iTest== 24: 
            strLet = "X"
        elif iTest== 25: 
            strLet = "Y"
        elif iTest== 26: 
            strLet = "Z"
        
        fXLNumberToLetter = fXLNumberToLetter + strLet
        
        # lRem = lRem - xl.WorksheetFunction.Floor(lRem / (26 ^ (iPass - 1)), 1) * (26 ^ (iPass - 1))
        lRem = lRem - math.floor(lRem / math.pow( 26, (iPass - 1) )) * math.pow( 26, (iPass - 1) )
        
        iPass = iPass - 1
        
        return fXLNumberToLetter

#    Debug.Print "fXLNumberToLetter: " + lNum + " = " + fXLNumberToLetter


def AddRadiusRatio(iBitNum, iCutterCount, xl):

    # xl = win32.Dispatch("Excel.Application")
    # xlbook = xl.Workbooks.Open(Filename = str + r"\test_automation\test\Book3.xlsx", ReadOnly = False)
    # sheet = xlbook.Worksheets("BIT" + str(BitNum))
    iCol = xl.ActiveSheet.UsedRange.Columns(xl.ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column
    AddRadiusRatio = fXLNumberToLetter(int(iCol))
    
    iColRad = xl.ActiveSheet.Range("Bit" + iBitNum + "_Radius").Column
    #iEndRow = int(Right(xl.ActiveSheet.UsedRange.Address, len(xl.ActiveSheet.UsedRange.Address) - InStrRev(xl.ActiveSheet.UsedRange.Address, "$")))#
    iEndRow = int(str(xl.ActiveSheet.UsedRange.Address)[len(xl.ActiveSheet.UsedRange.Address) - re.search("$", xl.ActiveSheet.UsedRange.Address)])
    
    sngMaxRadius = xl.WorksheetFunction.Max(xl.ActiveSheet.Range("Bit" + iBitNum + "_Radius"))
    
    xl.ActiveSheet.Cells(4, iCol).Value = "Radius_Ratio"
    xl.ActiveSheet.Cells(5, iCol).Select
    
    i =5
    while i <= iEndRow:
        if sngMaxRadius == 0 :
            xl.ActiveSheet.Cells(i, iCol).Value = 0
        else:
            xl.ActiveSheet.Cells(i, iCol).Value = xl.ActiveSheet.Cells(i, iColRad) / sngMaxRadius
        i += 1
    
    xl.ActiveSheet.Range(xl.ActiveSheet.Cells(5, iCol), xl.ActiveSheet.Cells(iEndRow, iCol)).Name = "Bit" + str(iBitNum) + "_Radius_Ratio"
    
    xl.ActiveSheet.Range("A1").Select
    return AddRadiusRatio
    
def AddIdentity(iBitNum, iCutterCount, xl):

    # xl = win32.Dispatch("Excel.Application")
    iCol = xl.ActiveSheet.UsedRange.Columns(xl.ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column    #fXLLetterToNumber(Mid(ActiveSheet.UsedRange.Address, InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2, InStrRev(ActiveSheet.UsedRange.Address, "$") - (InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2)))
    AddIdentity = fXLNumberToLetter(int(iCol))
    
    iColTag = xl.ActiveSheet.Range("Bit" + str(iBitNum) + "_Tag").Column
    iColLet = xl.ActiveSheet.Range("Bit" + str(iBitNum) + "_Blade_Num").Column
    strLet = fXLNumberToLetter(int(iColLet))
    # iEndRow = int(Right(ActiveSheet.UsedRange.Address, Len(ActiveSheet.UsedRange.Address) - InStrRev(ActiveSheet.UsedRange.Address, "$")))
    iEndRow = int(str(xl.ActiveSheet.UsedRange.Address)[len(xl.ActiveSheet.UsedRange.Address) - re.search("$", xl.ActiveSheet.UsedRange.Address)])
        
    xl.ActiveSheet.Cells(4, iCol).Value = "Identity"
    xl.ActiveSheet.Cells(5, iCol).Select
    
    
    i =5
    while i < iEndRow:
        iCount = 0
        iBlade = xl.ActiveSheet.Cells(i, iColLet).Value
        iTag = xl.ActiveSheet.Cells(i, iColTag).Value
        X = 5
        while X <= i:
            if xl.ActiveSheet.Cells(X, iColLet).Value == iBlade:
                if xl.ActiveSheet.Cells(X, iColTag).Value == iTag :
                    iCount = iCount + 1
            X += 1
        i += 1
    
        if iTag == 0 :
            xl.ActiveSheet.Cells(i, iCol).Value = "B" + iBlade + "C" + iCount
        elif iTag == 1 or iTag == 3 :
            xl.ActiveSheet.Cells(i, iCol).Value = "B" + iBlade + "BU" + iCount
        else:
            xl.ActiveSheet.Cells(i, iCol).Value = "B" + iBlade + "SC" + iCount    
        i += 1
    
    xl.ActiveSheet.Range(xl.ActiveSheet.Cells(5, iCol).Value, xl.ActiveSheet.Cells(iEndRow, iCol).Value).Name = "Bit" + str(iBitNum) + "_Identity"
    
    xl.ActiveSheet.Range("A1").Select
    
def AddXAxis(iBitNum, iCutterCount, xl) :
    # xl = win32.Dispatch("Excel.Application")
    iCol = xl.ActiveSheet.UsedRange.Columns(xl.ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column
    iEndRow = int(str(xl.ActiveSheet.UsedRange.Address)[len(xl.ActiveSheet.UsedRange.Address) - re.search("$", xl.ActiveSheet.UsedRange.Address)])

    xl.ActiveSheet.Cells(4, iCol).Value = "XAxis"
    xl.ActiveSheet.Cells(5, iCol).Select
    
    i =5
    while i < iEndRow:
        xl.ActiveSheet.Cells(i, iCol).Value = xl.ActiveSheet.Range("Bit" + str(iBitNum) + "_" + xl.ActiveSheet.Range("XAxis").Value).Cells(i - 4).Value
        i += 1
    
    xl.ActiveSheet.Range(xl.ActiveSheet.Cells(5, iCol), xl.ActiveSheet.Cells(iEndRow, iCol)).Name = "Bit" + str(iBitNum) + "_XAxis"
    
    xl.ActiveSheet.Range("A1").Select

def AddProfileLocation(iBitNum, iCutterCount, xl):
    # xl = win32.Dispatch("Excel.Application")
    iCol = xl.ActiveSheet.UsedRange.Columns(xl.ActiveSheet.UsedRange.Columns.Count).Offset(0, 1).Column    #fXLLetterToNumber(Mid(ActiveSheet.UsedRange.Address, InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2, InStrRev(ActiveSheet.UsedRange.Address, "$") - (InStr(1, ActiveSheet.UsedRange.Address, ":$") + 2)))
    AddProfileLocation = fXLNumberToLetter(int(iCol))
    
    iEndRow = int(str(xl.ActiveSheet.UsedRange.Address)[len(xl.ActiveSheet.UsedRange.Address) - re.search("$", xl.ActiveSheet.UsedRange.Address)])
    
    xl.ActiveSheet.Cells(4, iCol).Value = "ProfileLocation"
    xl.ActiveSheet.Cells(5, iCol).Select
    
    xl.ActiveSheet.Range(xl.ActiveSheet.Cells(5, iCol).Value, xl.ActiveSheet.Cells(iEndRow, iCol).Value).Name = "Bit" + iBitNum + "_ProfileLocation"
    
    xl.ActiveSheet.Range("A1").Select
    
def IsStinger(strCutterType):
    # Select Case strCutterType
    if strCutterType == "565_790_85_90": 
        IsStinger = True
    elif strCutterType == "565_690_85_90": 
        IsStinger = True
    elif strCutterType == "566_790_85_90": 
        IsStinger = True
    elif strCutterType == "566_690_85_90": 
        IsStinger = True
    elif strCutterType == "440_540_85_90": 
        IsStinger = True
    elif strCutterType == "440_537_85_90": 
        IsStinger = True
    elif strCutterType == "625_826_96_94": 
        IsStinger = True
    elif strCutterType == "562_7880_S": 
        IsStinger = True
    elif strCutterType == "562_8590_S": 
        IsStinger = True
    elif strCutterType == "562_7880_L": 
        IsStinger = True
    elif strCutterType == "562_8590_L": 
        IsStinger = True
    elif strCutterType == "562_9694_L": 
        IsStinger = True
    elif strCutterType == "B565_690_441_900_80": 
        IsStinger = True
    elif strCutterType == "B565_790_541_900_80": 
        IsStinger = True
    elif strCutterType == "New_8590_L": 
        IsStinger = True
    else:
        IsStinger = False
    return IsStinger

def FindCutters(Row_Start, Row_End, xl):

    #Finds different type of cutters, their quantities and puts them in a string
    # Dim Cutter_Type() As String
    # ReDim Preserve Cutter_Type(1 To (Row_End - Row_Start + 1)) As String
    Cutter_Type = [""]*(Row_End - Row_Start + 1)
    # xl = win32.Dispatch("Excel.Application")
    i = 1
    while i <= (Row_End - Row_Start + 1):
        Cutter_Type[i] = xl.ActiveSheet.Range("I" + str(i + Row_Start - 1).strip()).Value
        i += 1
    # On Error Resume Next
    i = 1
    while i <= (len(Cutter_Type) - 1):
        j = i + 1
        while j<= len(Cutter_Type):
            if str(Cutter_Type[j]) == str(Cutter_Type[i]) :
                temp_str = Cutter_Type[j]
                Cutter_Type[j] = Cutter_Type[i]
                Cutter_Type[i] = temp_str
            j += 1
        i += 1
    i = 1
    temp_str = ""
    # Do
    while i <= len(Cutter_Type):
        temp_cnt = 1
        j = i + 1
        while j <= len(Cutter_Type):
            if Cutter_Type[j] == Cutter_Type[i] :
                temp_cnt = temp_cnt + 1
            else:
                break
            j += 1
        temp_str = temp_str + Cutter_Type[i] + " (" + str(temp_cnt).strip() + "), "
        i = j
    # Loop Until i > UBound(Cutter_Type)
    # temp_str = Left(temp_str, Len(temp_str) - 2)
    temp_str = temp_str[0:len(temp_str) - 2]
    FindCutters = temp_str
    return FindCutters

def CutterCount(strRegion, iBitNum, strType, iCount, iBR, iSR):#todo
    
    xl = win32.Dispatch("Excel.Application")
    xl.Sheets("Bit" + iBitNum).Select
    str_path = os.getcwd()
    xlbook = xl.Workbooks.Open(Filename = str_path + r"\test_automation\test\Book3.xlsx", ReadOnly = True)
    sheet = xlbook.Worksheets("BIT" + str(iBitNum))
    
    blnFound = False
    blnStarted = False
    
    # ReDim strType(0)
    # ReDim iCount(2)
    # ReDim iBR(1)
    # ReDim iSR(1)
    
#define row starts and stops (iA, iB) for region
    if strRegion == "Cone" or strRegion == "Gauge" :
        iA = sheet.Range("Bit" + str(iBitNum) + "_" + strRegion).Item[1].row
        iB = sheet.Range("Bit" + str(iBitNum) + "_" + strRegion).Item[sheet.Range("Bit" + str(iBitNum) + "_" + strRegion).Count].row
    else:
        if sheet.Range("Bit" + str(iBitNum) + "_Nose").Count != 0 :
            iA = sheet.Range("Bit" + str(iBitNum) + "_Nose").Item[1].row
        else:
            iA = sheet.Range("Bit" + str(iBitNum) + "_Shoulder").Item[1].row
        iB = sheet.Range("Bit" + str(iBitNum) + "_Shoulder").Item(sheet.Range("Bit" + str(iBitNum) + "_Shoulder").Count).row
        
    iRow = iA
    while iRow <= iB:
        i = 0
        while i <= len(strType) - 1:
            if strType[i] == sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Cutter_Type").Column) :
                blnFound = True
                
                iCount[0] = iCount[0] + 1  #Total Cutter Count
                if sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Tag").Column) == 0 :
                    iCount[1] = iCount[1] + 1  #Primary Cutter Count
                else:
                    iCount[2] = iCount[2] + 1  #Backup Cutter Count
                
                if iBR[0] > sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column) :
                    iBR[0] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column)
                if iBR[1] < sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column) :
                    iBR[1] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column)
                if iSR[0] > sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column) :
                    iSR[0] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column)
                if iSR[1] < sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column) :
                    iSR[1] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column)
                
                break
            i += 1
        
        if blnFound == False :
            if strType[i - 1] == "" :
                strType[i - 1] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Cutter_Type").Column).Value
            else:
                strType[i] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Cutter_Type").Column).Value
            
            iCount[0] = iCount[0] + 1  #Total Cutter Count
            if sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Tag").Column) == 0 :
                iCount[1] = iCount[1] + 1  #Primary Cutter Count
            else:
                iCount[2] = iCount[2] + 1  #Backup Cutter Count
            
            # If Not blnStarted Then
            if blnStarted != True:
                iBR[0] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column).Value
                iBR[1] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column).Value
                iSR[0] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column).Value
                iSR[1] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column).Value
            else:
                if iBR[0] > sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column).Value :
                    iBR[0] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column).Value
                if iBR[1] < sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column).Value :
                    iBR[1] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Back_Rake").Column).Value
                if iSR[0] > sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column).Value :
                    iSR[0] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column).Value
                if iSR[1] < sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column).Value :
                    iSR[1] = sheet.Cells(iRow, sheet.Range("Bit" + str(iBitNum) + "_Side_Rake").Column).Value
            blnStarted = True
        
        blnFound = False
        
        iRow += 1
    xl.DisplayAlerts = False
    xl.ActiveWorkbook.Save()
    xl.ActiveWorkbook.Close()
    return None

def ProcessCutterInfo(BitNum, radius, ang_arnd, height, prf_ang, bk_rake, sd_rake, cut_dia, size_type, blade, exposure, ProfL, Tag, roll, hMax, DeltaProfAng, Tag360, TagCreo, xlbook, xl):

    # Pi = 4 * Atn(1)
    pi = math.pi
    blnHidden = ""
    # gblnCentSting = False
    
    total_cutter_count = len(radius)
    # xl = win32.Dispatch("Excel.Application")
    # xl.Visible = 0
    # xl.DisplayAlerts = 0
    # str_path = os.getcwd()
    # xlbook = xl.Workbooks.Open(Filename = str_path + r"\test_automation\test\Book3.xlsx", ReadOnly = False)
    sheet = xlbook.Worksheets("BIT" + str(BitNum))
    
    start_row_index = 4
    row_index = start_row_index + 1
    # For i = 1 To total_cutter_count
    i = 1
    while i <= total_cutter_count:
        sheet.Range("A" + str(row_index).strip()).Value = i
        sheet.Range("B" + str(row_index).strip()).Value = RoundOff(float(radius[i]), 4)
        sheet.Range("C" + str(row_index).strip()).Value = ang_arnd[i]
        sheet.Range("D" + str(row_index).strip()).Value = RoundOff(float(height[i]), 4)
        sheet.Range("E" + str(row_index).strip()).Value = RoundOff(float(prf_ang[i]), 2)
        # print(sheet.Range("E" + str(row_index).strip()).Value)
        # print(sheet.Range("ForceRatioSplitInc").Value)
        if sheet.Range("E" + str(row_index).strip()).Value <= 0:  #todo
            iRowLastInner = row_index
        sheet.Range("F" + str(row_index).strip()).Value = RoundOff(float(bk_rake[i]), 2)
        sheet.Range("G" + str(row_index).strip()).Value = RoundOff(float(sd_rake[i]), 2)
        sheet.Range("H" + str(row_index).strip()).Value = cut_dia[i]
        sheet.Range("I" + str(row_index).strip()).Value = size_type[i]
        sheet.Range("J" + str(row_index).strip()).Value = blade[i]
        sheet.Range("K" + str(row_index).strip()).Value = RoundOff(float(exposure[i]), 4)
        sheet.Range("L" + str(row_index).strip()).Value = RoundOff(float(ProfL[i]), 3)
        sheet.Range("P" + str(row_index).strip()).Value = int(Tag[i])
        sheet.Range("Q" + str(row_index).strip()).Value = RoundOff(roll[i], 2)
        sheet.Range("R" + str(row_index).strip()).Value = RoundOff(DeltaProfAng[i], 4)
        sheet.Range("S" + str(row_index).strip()).Value = int(Tag360[i])
        sheet.Range("T" + str(row_index).strip()).Value = int(TagCreo[i])
        
        if i == 1 :
            sheet.Range("M" + str(row_index).strip()).Value = RoundOff(float(ProfL[i]), 3)
        else:
            sheet.Range("M" + str(row_index).strip()).Value = RoundOff(float(ProfL[i] - ProfL(i - 1)), 3)
        if ProfL[i] == 0 :   #Adding Cutter Distances if CFB is used
            if i == 1 :
                sheet.Range("M" + str(row_index).strip()).Value = RoundOff(float(sheet.Range("B" + str(row_index).strip()).Value / math.cos(sheet.Range("E" + str(row_index).strip()).Value * pi / 180)), 3)
                sheet.Range("L" + str(row_index).strip()).Value = sheet.Range("M" + str(row_index).strip()).Value
            elif i > 1 :
#                Range("M" + LTrim(str(row_index))).Value = Format(((Range("B" + LTrim(str(row_index))).Value - Range("B" + LTrim(str(row_index - 1))).Value) ^ 2 + (Range("D" + LTrim(str(row_index))).Value - Range("D" + LTrim(str(row_index - 1))).Value) ^ 2) ^ 0.5, "0.00")
                sheet.Range("M" + str(row_index).strip()).Value = RoundOff((sheet.Range("B" + str(row_index).strip()).Value - math.pow(sheet.Range("B" + str(row_index - 1).strip()).Value) , 2) + math.pow(math.pow((sheet.Range("D" + str(row_index).strip()).Value - sheet.Range("D" + str(row_index - 1).strip()).Value), 2), 0.5), 3)
                ProcessCutterInfo = "Cutter Distances Approximated"
                sheet.Range("L" + str(row_index).strip()).Value = sheet.Range("M" + str(row_index).strip()).Value + sheet.Range("L" + str(row_index - 1).strip()).Value
#        Stop
        row_index = row_index + 1
        i += 1


# Process the rest of the cutters info
    # Dim Heading_Array As Variant
    Heading_Array = ["Cutter_Num", "Radius", "Angle_Around", "Height", "Profile_Angle", "Back_Rake", "Side_Rake", "Cutter_Diameter", "Cutter_Type", "Blade_Num", "Exposure", "Profile_Position", "Cutter_Distance", "Height2", "Height3", "Tag", "Roll", "DeltaProfAng", "E360Tag", "CreoTag"]
    sheet.Range("D" + str(start_row_index + 1).strip()).Select
    Max_Height = xl.WorksheetFunction.Max(sheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(total_cutter_count - 1, 0)))
    for heading_name in sheet.Range(Heading_Array):
        sheet.Cells(start_row_index, i + 1).Value = heading_name
        sheet.Cells(start_row_index + 1, i + 1).Select
        sheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(total_cutter_count - 1, 0)).Name = "Bit" + BitNum + "_" + heading_name
    i = 1
    Gauge_zero_bit_height = 0
    while RoundOff(float(prf_ang[i]), 2) < 90 and i < total_cutter_count:
        i = i + 1

    Gauge_zero_bit_height = height[i]
    #MsgBox Gauge_zero_bit_height, vbInformation
    i = 1
    while i <= total_cutter_count:
        sheet.Cells(i + start_row_index, 14).Value = RoundOff(float(height[i] - Max_Height), 4)
        sheet.Cells(i + start_row_index, 15).Value = RoundOff(float(height[i] - Gauge_zero_bit_height), 4)
        i += 1
    
    hMax = xl.WorksheetFunction.Max(sheet.Range("Bit" + str(BitNum) + "_Height3"))
    
    # Dim Cutter_Type(1 To 4) As String
    # Dim Cutter_Count(1 To 4) As String
    # Dim BR_Low(1 To 4) As Double
    # Dim BR_High(1 To 4) As Double
    # Dim SR_Low(1 To 4) As Double
    # Dim SR_High(1 To 4) As Double
    Cutter_Type = []
    Cutter_Count = []
    BR_Low = []
    BR_High = []
    SR_Low = []
    SR_High = []
    
    
#Adding new columns
    strLastCol = AddRadiusRatio(int(BitNum), int(total_cutter_count), xl)
    strLastCol = AddIdentity(int(BitNum), int(total_cutter_count), xl)
    strLastCol = AddXAxis(int(BitNum), int(total_cutter_count), xl)
    strLastCol = AddProfileLocation(int(BitNum), int(total_cutter_count), xl)
#    strLastCol = AddProfZone(CInt(BitNum), CInt(Total_Cutter_Count))

    xl.ActiveSheet.Cells(4, 1).Select
    # xl.ActiveSheet.Range(xl.Selection, xl.Selection.End(xlEnd)).Select #todo
    # xl.ActiveSheet.Range(xl.Selection, xl.Selection.End(xlDown)).Select
    Last_Cell = xl.Selection.Address
    
#    Last_Cell = ActiveSheet.UsedRange.Address
    Last_row = xl.ActiveSheet.Range(Last_Cell).Rows.Count
    Last_Col = xl.ActiveSheet.Range(Last_Cell).Columns.Count
    row_index = start_row_index + 1
    row_count = 0
    
    iColTag = xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Tag").Column
    
    
    
    
#Getting # of cone cutters (based off of equal cone profile angles
    iStartConerow = row_index
#    Do While IsStinger(Cells(iStartConerow, Range("Bit" + BitNum + "_Cutter_Type").Column)) And Cells(row_index, Range("Bit" + BitNum + "_Radius").Column) < 0.005
#Determine if there's a central stinger
    while IsStinger(xl.ActiveSheet.Cells(iStartConerow, xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Cutter_Type").Column)) and xl.ActiveSheet.Cells(iStartConerow, xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Radius").Column) < 0.005:
        iStartConerow = iStartConerow + 1
        row_count = row_count + 1
        xl.ActiveSheet.Range("CentStingPresent").Value = True
    
#Determine if a Stinger is cutting the core
    if xl.ActiveSheet.Range("CentStingPresent") == True :
    #Typically when a stinger cuts the core, it's profile angle is at -47deg
        if IsStinger(xl.ActiveSheet.Cells(iStartConerow, xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Cutter_Type").Column)) and xl.ActiveSheet.Cells(iStartConerow, xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Profile_Angle").Column) < -45 :
            iStartConerow = iStartConerow + 1
            row_count = row_count + 1
    
    row_index = iStartConerow
#Determine if the profile is AAAG (first A cutters will have gradually increasing negative profile angles. These will be considered "Cone" cutters. When the profile angles start to decrease, the tangent where the 2nd A comes in has been passed. This is the start of the "Nose".
    if xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Type").Value == "AAAG" :
        while xl.ActiveSheet.Range("E" + str(row_index).strip()).Value < xl.ActiveSheet.Range("E" + str(row_index - 1).strip()).Value:
            row_index = row_index + 1
            row_count = row_count + 1
    else:
    #    Do Until Not Range("E" + LTrim(str(row_index))).Value = Range("E" + str(start_row_index + 1).strip()).Value

#Stop ' Test to make sure the stingers in the cone are not affecting "cone area" counts/regions
        if xl.ActiveSheet.Range("E" + str(row_index).strip()).Value == xl.ActiveSheet.Range("E" + str(iStartConerow).strip()).Value:
            if IsStinger(xl.ActiveSheet.Cells(str(row_index + 1).strip(), xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Cutter_Type").Column).Value) and xl.ActiveSheet.Cells(str(row_index + 1).strip(), xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Tag").Column).Value > 0 :
                row_index = row_index + 1
                row_count = row_count + 1
            
            row_index = row_index + 1
            row_count = row_count + 1
        # Loop Until Not Range("E" + LTrim(str(row_index))).Value = Range("E" + LTrim(str(iStartConerow))).Value
    
    Cutter_Count[1] = row_index - start_row_index - 1
    
    xl.ActiveSheet.Cells(start_row_index + 1, 1).Select
#    Cells(iStartConerow, 1).Select
    xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, Last_Col - 1)).Interior.ColorIndex = 35
    xl.ActiveSheet.Range(xl.ActiveCell.Offset(0, iColTag - 1), xl.ActiveCell.Offset(row_count - 1, iColTag - 1)).Name = "Bit" + str(BitNum) + "_Cone"
    
    # For i = 1 To Range("Bit" + BitNum + "_Cone").Count
    i = 1
    while i <= xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Cone").Count:
        xl.ActiveSheet.Range("Bit" + str(BitNum) + "_ProfileLocation").Item[i] = "C"
        i += 1
    
    
    Cutter_Type[1] = FindCutters(start_row_index + 1, row_index - 1, xl)
    xl.ActiveSheet.Range("F" + str(start_row_index + 1).strip()).Select
    BR_Low[1] = xl.WorksheetFunction.Min(xl.ActiveSheet.Range(xl.ActiveCell.Offset(iStartConerow - (start_row_index + 1), 0), xl.ActiveCell.Offset(row_count - 1, 0)))
    BR_High[1] = xl.WorksheetFunction.Max(xl.ActiveSheet.Range(xl.ActiveCell.Offset(iStartConerow - (start_row_index + 1), 0), xl.ActiveCell.Offset(row_count - 1, 0)))
    xl.ActiveSheet.Range("G" + str(start_row_index + 1).strip()).Select
    SR_Low[1] = xl.WorksheetFunction.Min(xl.ActiveSheet.Range(xl.ActiveCell.Offset(iStartConerow - (start_row_index + 1), 0), xl.ActiveCell.Offset(row_count - 1, 0)))
    SR_High[1] = xl.WorksheetFunction.Max(xl.ActiveSheet.Range(xl.ActiveCell.Offset(iStartConerow - (start_row_index + 1), 0), xl.ActiveCell.Offset(row_count - 1, 0)))
    pre_row_index = row_index
    xl.ActiveSheet.Cells(row_index, 1).Select
    
    
    
    
# Getting Nose cutters
#    If Not Range("E" + LTrim(str(start_row_index + 1))).Value = 0 :
    if  xl.ActiveSheet.Range("E" + str(iStartConerow).strip()).Value != 0 :
        row_count = 0
        # Do Until Range("E" + LTrim(str(row_index))).Value >= (Range("E" + LTrim(str(iStartConerow))).Value) * (-1)
        while xl.ActiveSheet.Range("E" + str(row_index).strip()).Value < (xl.ActiveSheet.Range("E" + str(iStartConerow).strip()).Value) * (-1):
            row_index = row_index + 1
            row_count = row_count + 1
        
        if row_count == 0 :
            #no cone cutters in design
#Stop
            Cutter_Type[2] = "-"
            Cutter_Count[2] = 0
            BR_Low[2] = 0
            BR_High[2] = 0
            SR_Low[2] = 0
            SR_High[2] = 0
            xl.ActiveSheet.Range(xl.ActiveSheet.Cells(1, iColTag), xl.ActiveSheet.Cells(1, iColTag)).Name = "Bit" + str(BitNum) + "_Nose"
        else:
            Cutter_Count[2] = row_index - pre_row_index
            xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, Last_Col - 1)).Interior.ColorIndex = 36
            xl.ActiveSheet.Range(xl.ActiveCell.Offset(0, iColTag - 1), xl.ActiveCell.Offset(row_count - 1, iColTag - 1)).Name = "Bit" + str(BitNum) + "_Nose"
            
            Cutter_Type[2] = FindCutters(pre_row_index, row_index - 1, xl)
            xl.ActiveSheet.Range("F" + pre_row_index).Select
            BR_Low[2] = xl.WorksheetFunction.Min(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, 0)))
            BR_High[2] = xl.WorksheetFunction.Max(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, 0)))
            xl.ActiveSheet.Range("G" + pre_row_index).Select
            SR_Low[2] = xl.WorksheetFunction.Min(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, 0)))
            SR_High[2] = xl.WorksheetFunction.Max(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, 0)))
            pre_row_index = row_index
            xl.ActiveSheet.Cells(row_index, 1).Select
    else:
        Cutter_Type[2] = "-"
        Cutter_Count[2] = 0
        BR_Low[2] = 0
        BR_High[2] = 0
        SR_Low[2] = 0
        SR_High[2] = 0
        xl.ActiveSheet.Range(xl.ActiveSheet.Cells(1, iColTag), xl.ActiveSheet.Cells(1, iColTag)).Name = "Bit" + str(BitNum) + "_Nose"

    i = 1
    while i <= xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Nose").Count:
        xl.ActiveSheet.Range("Bit" + str(BitNum) + "_ProfileLocation").Item[xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Cone").Count + i] = "N"
        i += 1
    
# getting Shoulder cutters
    row_count = 0
    # Do Until Range("E" + LTrim(str(row_index))).Value >= 90 Or row_index = Last_row + start_row_index
    while xl.ActiveSheet.Range("E" + str(row_index).strip()).Value < 90 and row_index != Last_row + start_row_index:
        row_index = row_index + 1
        row_count = row_count + 1
    Cutter_Count[3] = row_index - pre_row_index
    xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, Last_Col - 1)).Interior.ColorIndex = 37
    xl.ActiveSheet.Range(xl.ActiveCell.Offset(0, iColTag - 1), xl.ActiveCell.Offset(row_count - 1, iColTag - 1)).Name = "Bit" + str(BitNum) + "_Shoulder"
    
    i = 1
    while i <= xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Shoulder").Count:
        xl.ActiveSheet.Range("Bit" + str(BitNum) + "_ProfileLocation").Item[xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Cone").Count + xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Nose").Count + i] = "S"
        i += 1
    
    Cutter_Type[3] = FindCutters(pre_row_index, row_index - 1, xl)
    xl.ActiveSheet.Range("F" + pre_row_index).Select
    BR_Low[3] = xl.WorksheetFunction.Min(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, 0)))
    BR_High[3] = xl.WorksheetFunction.Max(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, 0)))
    xl.ActiveSheet.Range("G" + pre_row_index).Select
    SR_Low[3] = xl.WorksheetFunction.Min(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, 0)))
    SR_High[3] = xl.WorksheetFunction.Max(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(row_count - 1, 0)))
    if row_index == Last_row + start_row_index :
        Cutter_Type[4] = "-"
        Cutter_Count[4] = 0
        BR_Low[4] = 0
        BR_High[4] = 0
        SR_Low[4] = 0
        SR_High[4] = 0
    else:
        pre_row_index = row_index
        xl.ActiveSheet.Cells(row_index, 1).Select
        Cutter_Count[4] = Last_row - pre_row_index + start_row_index
        xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(Last_row - row_index + start_row_index - 1, Last_Col - 1)).Interior.ColorIndex = 40
        xl.ActiveSheet.Range(xl.ActiveCell.Offset(0, iColTag - 1), xl.ActiveCell.Offset(Last_row - row_index + start_row_index - 1, iColTag - 1)).Name = "Bit" + str(BitNum) + "_Gauge"
    
        i = 1
        while i <= xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Gauge").Count:
            xl.ActiveSheet.Range("Bit" + str(BitNum) + "_ProfileLocation").Item[xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Cone").Count + xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Nose").Count + xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Shoulder").Count + i] = "G"
            i += 1
        
        Cutter_Type[4] = FindCutters(pre_row_index, Last_row + start_row_index - 1, xl)
        xl.ActiveSheet.Range("F" + pre_row_index).Select
        BR_Low[4] = xl.WorksheetFunction.Min(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(Last_row - row_index + start_row_index - 1, 0)))
        BR_High[4] = xl.WorksheetFunction.Max(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(Last_row - row_index + start_row_index - 1, 0)))
        xl.ActiveSheet.Range("G" + pre_row_index).Select
        SR_Low[4] = xl.WorksheetFunction.Min(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(Last_row - row_index + 1, 0)))
        SR_High[4] = xl.WorksheetFunction.Max(xl.ActiveSheet.Range(xl.ActiveCell, xl.ActiveCell.Offset(Last_row - row_index + 1, 0)))
    
#    Range("A" + start_row_index + ":" + strLastCol + start_row_index).AutoFilter
    

    
# Create Range for Inner/Outer cutters
    if iRowLastInner == 0 : 
        iRowLastInner = row_index - 1
    xl.ActiveSheet.Range(xl.ActiveSheet.Cells(start_row_index + 1, xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Profile_Angle").Column), xl.ActiveSheet.Cells(iRowLastInner, xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Profile_Angle").Column)).Name = "Bit" + str(BitNum) + "_InnerCutters"
    xl.ActiveSheet.Range(xl.ActiveSheet.Cells(iRowLastInner + 1, xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Profile_Angle").Column), xl.ActiveSheet.Cells(total_cutter_count + start_row_index, xl.ActiveSheet.Range("Bit" + str(BitNum) + "_Profile_Angle").Column)).Name = "Bit" + str(BitNum) + "_OuterCutters"

    
# Autofilter
    xl.ActiveSheet.Range("A4:" + fXLNumberToLetter(int(xl.ActiveSheet.Range(xl.ActiveSheet.UsedRange.Address).Columns.Count)) + "4").AutoFilter
    
#Info for INFO2
    # Dim strArea As Variant
    # Dim strType() As String
    # Dim iCount() As Integer
    # Dim iBR() As Integer
    # Dim iSR() As Integer
    
    xl.ActiveSheet.Range(xl.ActiveSheet.Range("Bit" + str(BitNum) + "_I2_ConeSize"), xl.ActiveSheet.Range("Bit" + str(BitNum) + "_I2_GaugeSR")).ClearContents
    
    strArea = ["Cone", "Face", "Gauge"]
    strType = []
    iCount = []
    iBR = []
    iSR = []
    i = 0#todo
    while i <= 2:
        CutterCount(str(strArea[i]), int(BitNum), strType, iCount, iBR, iSR)
        
        #for review only
        sheet = xlbook.Worksheets("INFO2")
        
        X = 0
        while X <= len(strType) - 1:
            sheet.Range("Bit" + str(BitNum) + "_I2_" + strArea[i] + "Size").Value = sheet.Range("Bit" + str(BitNum) + "_I2_" + strArea[i] + "Size").Value + strType[X] + "/"
            X = X + 1
        sheet.Range("Bit" + BitNum + "_I2_" + strArea[i] + "Size").Value = str(sheet.Range("Bit" + BitNum + "_I2_" + strArea[i] + "Size").Value)[0:len(sheet.Range("Bit" + BitNum + "_I2_" + strArea[i] + "Size")) - 1]
        
        sheet.Range("Bit" + BitNum + "_I2_" + strArea[i] + "Qty").Value = iCount[0] + "/" + iCount[1] + "/" + iCount[2]
#        If iCount(1) <> 0 Then Range("Bit" + BitNum + "_I2_" + strArea[i] + "Qty") = Range("Bit" + BitNum + "_I2_" + strArea[i] + "Qty") + "(" + iCount(1) + ")"
        sheet.Range("Bit" + BitNum + "_I2_" + strArea[i] + "BR").Value = iBR[0] + "/" + iBR[1]
        sheet.Range("Bit" + BitNum + "_I2_" + strArea[i] + "SR").Value = iSR[0] + "/" + iSR[1]
    
        strType = []
        iCount = []
        iBR = []
        iSR = []
        i += 1
        

#Info for Drill_Scan
    # Dim iColBR As Integer
    
    sheet = xlbook.Worksheets("INFO2")
    
    iColBR = sheet.Range("Bit" + BitNum + "_Back_Rake").Column
    iColTag = sheet.Range("Bit" + BitNum + "_Gauge").Column
    
    #Cone Depth
    if sheet.Cells(sheet.Range("Bit" + BitNum + "_Cone").row, sheet.Range("Bit" + BitNum + "_Radius").Column) < 0.001 and IsStinger(sheet.Cells(sheet.Range("Bit" + BitNum + "_Cone").row, sheet.Range("Bit" + BitNum + "_Cutter_Type").Column)) :
        i = 1
    else:
        i = 0
    sheet.Range("Bit" + BitNum + "_DS_CS_CD").Value = RoundOff(sheet.Cells(sheet.Range("Bit" + BitNum + "_Cone").row + i, sheet.Range("Bit" + BitNum + "_Height2").Column) * -1, 2)
    
    #Cutting Strux Ht
    sheet.Range("Bit" + BitNum + "_DS_CS_Ht").Value = RoundOff(sheet.Cells(sheet.Range("Bit" + BitNum + "_Gauge").row, sheet.Range("Bit" + BitNum + "_Height2").Column) * -1, 2)
    
    #Cutting Strux Avg BR
    sheet.Range("Bit" + BitNum + "_DS_CS_BR").Value = RoundOff(xl.WorksheetFunction.Average(sheet.Range(sheet.Range("Bit" + BitNum + "_Back_Rake").Item(1), sheet.Cells(sheet.Range("Bit" + BitNum + "_Gauge").row - 1, sheet.Range("Bit" + BitNum + "_Back_Rake").Column))), 2)
    
    #Active Gauge Height
    sheet.Range("Bit" + BitNum + "_DS_AG_Ht").Value = RoundOff(sheet.Range("Bit" + BitNum + "_Height3").Item(sheet.Range("Bit" + BitNum + "_Height3").Count) * -1, 2)
    
    #Active Gauge BR
    sheet.Range("Bit" + BitNum + "_DS_AG_BR").Value = RoundOff(xl.WorksheetFunction.Average(sheet.Range("Bit" + BitNum + "_Gauge").Offset(0, iColBR - iColTag)), 2)
    
    
    
    
    
#End Extra Info stuff
        
    if blnHidden : 
        xlbook.Sheets("BIT" + str(BitNum)).Hidden = True
    
    sheet = xlbook.Worksheets("INFO")
    row_index = sheet.Range("Cutting_Structure_Info").row
    row_index = row_index + 2 + int(BitNum - 1) * 4
    i = 1
    while i <= 4:
        sheet.Range("I" + str(row_index + i - 1).strip()).Value = Cutter_Type[i]
        sheet.Range("L" + str(row_index + i - 1).strip()).Value = Cutter_Count[i]
        sheet.Range("M" + str(row_index + i - 1).strip()).Value = BR_Low[i]
        sheet.Range("N" + str(row_index + i - 1).strip()).Value = BR_High[i]
        sheet.Range("O" + str(row_index + i - 1).strip()).Value = SR_Low[i]
        sheet.Range("P" + str(row_index + i - 1).strip()).Value = SR_High[i]
        i += 1
    
#    GoTo continue
# continue:

    sheet.Range("G" + str(row_index).strip()).Value = total_cutter_count
    # xl.DisplayAlerts = False
    # xl.ActiveWorkbook.Save()
    # xl.ActiveWorkbook.Close()
    return ProcessCutterInfo
    
    

#(todo)
def read_bit(Bit_File_Handler, iBit, blnRefresh = True, blnSkipIndices = True):
    if os.path.exists(Bit_File_Handler) == False:
        logging.warning('Invalid bit path: ' + Bit_File_Handler)
        return None
    str_vis = check_bit_in(Bit_File_Handler, "Name")
    if str_vis == "":
        logging.debug('no bit.in found. Trying reamer.in.')
        str_vis = check_ramer_in(Bit_File_Handler, "Name")
        if str_vis == "":
            logging.debug('no reamer.in found. Loading just Bit info.')
    if str_vis != Bit_File_Handler.split('\\')[-1]:
        logging.warning('The active bit in this project is: ' + str_vis + ", You've loaded: " + Bit_File_Handler.split('\\')[-1])
        return None
    a = process_des(iBit ,Bit_File_Handler)
    if a == "" :
        logging.warning('load .des file fail')

    # BFH_pos = InStrRev(Bit_File_Handler, "\")
    BFH_pos = Bit_File_Handler.split('\\')[-1]
    # strSummaryFile = Mid(Bit_File_Handler, 1, BFH_pos) & "summary_regular" 'changed the extension from summary_regular.*
    strSummaryFile = Bit_File_Handler.replace(BFH_pos, "summary_regular")
    # A = ProcessSummary(BitNum, strSummaryFile)

    
    
  
    
    
    
    # if BitNum == 1 :
    #     subForceDefaults
    

#Ensure just the bits with "Show Bit" flag are shown
    
#     subUpdate False
    
#     For i = 1 To 5
#         Sheets("HOME").Activate
#         If i = CInt(BitNum) Then
#             Select Case strVis
#                 Case "Show_Bit": A = HideCutters(BitNum, False, False, False) ' Set to Show
#                 Case "Hide_Bit":
#                     A = HideCutters(BitNum, False, False, False) ' Set to Show (it was hidden before, but user refreshed and should see the results!)
#                     strVis = "Show_Bit"
#                 Case "Hide_Backups_Bit": A = HideCutters(BitNum, True, True, False) ' set to hide backups
#                 Case "Hide_Off_Backups_Bit": A = HideCutters(BitNum, True, True, True) ' set to hide off-profile backups
#             End Select
            
#             ActiveSheet.Shapes(strVis & i).ControlFormat.Value = 1
            
#         ElseIf ActiveSheet.Shapes("Hide_Bit" & i).ControlFormat.Value = 1 Then
#             ActiveSheet.Shapes("Hide_Bit" & i).ControlFormat.Value = 1
#             A = HideCutters(i, False, True, False)
#         End If
#     Next i
    
#     Sheets("HOME").Activate

#     subUpdate True
    
# #Record the finish "time"
#     sEnd(0) = Timer
#     sEnd(1) = Timer

# #Status info updated for the bit
#     strMsg = Bit_Load_Status
#     If Bit_Warn_Status <> "" Then strMsg = strMsg & Chr(10) & Bit_Warn_Status
#     strMsg = strMsg & Chr(10) & "{" & Range("ReadVer") & "} Loaded in [PostRead: " & Format(sEnd(1) - sStart(1), "0.00") & "s] [RBit: " & Format(sEnd(2) - sStart(2), "0.00") & "s] [RSumm: " & Format(sEnd(3) - sStart(3), "0.00") & "s]"
    
#     Range("Bit" & BitNum & "_Status").Value = strMsg
    
#     If Bit_Warn_Status <> "" And Sheets("HOME").Shapes("chkWarnings").ControlFormat.Value <> 1 Then
#         MsgBox Bit_Warn_Status & " for Iteration " & BitNum - 1 & "." & vbCrLf & vbCrLf & "Please review SUMMARY sheet.", vbExclamation + vbOKOnly, "Warning"
#     End If
    
#     CoreDiamCutterUpdate CInt(BitNum)
    
#     subActive False
    
    
    
if __name__ == "__main__":
    read_bit(r"C:\Users\JWang294\Documents\projectfile\test_config_cases\test-20210416173550-lbo-0416cpc\Static\66048_stinger\66048a00.des", 1)