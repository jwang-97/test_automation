import os, re, logging, time, argparse, pptx, pptx.util, csv, io, unicodedata, csv
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
    temp = 10 ^ Dig
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
    
def GetGaugePadInfo(iBit, sBitSize, Bit_File_Handler):
    try:
        with open(Bit_File_Handler, "r", encoding="gbk", errors="ignore") as des_file:
            # lines = des_file.readlines()
            lines = des_file.read().splitlines()
            find_BasicData = [ string for string in lines if "#padNo(1..)" in string]
            fileHeader = []
            d1 = []
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
                if sBitSize > 0:
                    while lines[iRowEnd + 1] != "[end_section_simple_pad]":
                        iRowEnd = iRowEnd + 1
                else:
                    while lines[iRowEnd + 1].split()[1] != "":
                        iRowEnd = iRowEnd + 1
                if iRowEnd == 0:
                    des_file.close()
                    return None
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
                while i <= iRowEnd:
                    sRad1.append(lines[find_GP + i].split()[4])
                    sLen1.append(lines[find_GP + i].split()[8])
                    sBeta1.append(lines[find_GP + i].split()[9])
                    sTaper1.append(lines[find_GP + i].split()[10])
                    sRad2.append(lines[find_GP + i].split()[12])
                    sLen2.append(lines[find_GP + i].split()[11])
                    sBeta2.append(lines[find_GP + i].split()[13])
                    sTaper2.append(lines[find_GP + i].split()[14])
                    sWidth.append(lines[find_GP + i].split()[7])
                    i += 1
            #averaging all values for all blades
                # i = 1
                sVals = []*8
                while i <= iRowEnd:
                    sVals[0] = sVals[0] + sRad1[i]
                    sVals[1] = sVals[1] + sLen1[i]
                    sVals[2] = sVals[2] + sBeta1[i]
                    sVals[3] = sVals[3] + sTaper1[i]
                    sVals[4] = sVals[4] + sRad2[i]
                    sVals[5] = sVals[5] + sLen2[i]
                    sVals[6] = sVals[6] + sBeta2[i]
                    sVals[7] = sVals[7] + sTaper2[i]
                    sVals[8] = sVals[8] + sWidth[i]
                    i += 1
                
                i = 0
                while i <= len(sVals):
                    sVals[i] = sVals[i] / iRowEnd
                i += 1
                des_file.close()

                fileHeader.append("Bit" + str(iBit) + "_GPRad")
                fileHeader.append("Bit" + str(iBit) + "_GPLen") 
                fileHeader.append("Bit" + str(iBit) + "_GPWidth")
                d1.append(RoundOff(sVals[0], 4))
                d1.append(RoundOff(sVals(1) + sVals[5], 2))
                d1.append(RoundOff(sVals[8], 4))
                # d2 = ["Li", "80"]
                
                #Calc Nominal Length
                if sBitSize > 0 and GaugePadNom(sBitSize) == sVals[0] :#section 1 is nominal(todo)
                    if sVals[3] == 0 :
                        if sVals[7] == 0 and sVals[5] > 0 :
                            fileHeader.append("Bit" + str(iBit) + "_GPNomLen")
                            d1.append(RoundOff(sVals[1] + sVals[5], 2))
                        else:
                            fileHeader.append("Bit" + str(iBit) + "_GPNomLen")
                            d1.append(RoundOff(sVals[1], 2))
                    else:
                        fileHeader.append("Bit" + str(iBit) + "_GPNomLen")
                        d1.append(0)
                else:
                    fileHeader.append("Bit" + str(iBit) + "_GPNomLen")
                    d1.append(0)

                if sVals[3] > 0 : #section 1 is tapered
                    if sVals[5] > 0 :
                        if sVals[7] > 0 :
                            if sVals[3] == sVals[7] :
                                fileHeader.append("Bit" + str(iBit) + "_GPTaperLen")
                                d1.append(RoundOff(sVals(1) + sVals(5), 2))
                                if sVals[3] == sVals[7] :
                                    fileHeader.append("Bit" + str(iBit) + "_GPTaperAng")
                                    d1.append(RoundOff(sVals[3], 2))
                                else:
                                    fileHeader.append("Bit" + str(iBit) + "_GPTaperAng")
                                    d1.append("'" + str(RoundOff(sVals[3], 2)) + " / " + str(RoundOff(sVals[7], 2)))
                            else:
                                fileHeader.append("Bit" + str(iBit) + "_GPTaperLen")
                                d1.append("'" + str(RoundOff(sVals[1], 2)) + " / " + str(RoundOff(sVals[5], 2)))
                                fileHeader.append("Bit" + str(iBit) + "_GPTaperAng")
                                d1.append("'" + str(RoundOff(sVals[3], 2)) + " / " + str(RoundOff(sVals[7], 2)))
                        else:
                            fileHeader.append("Bit" + str(iBit) + "_GPTaperLen")
                            d1.append(RoundOff(sVals[1], 2))
                            fileHeader.append("Bit" + str(iBit) + "_GPTaperAng")
                            d1.append(RoundOff(sVals[3], 2))
                    else:
                        fileHeader.append("Bit" + str(iBit) + "_GPTaperLen")
                        d1.append(RoundOff(sVals[1], 2))
                        fileHeader.append("Bit" + str(iBit) + "_GPTaperAng")
                        d1.append(RoundOff(sVals[3], 2))
                else:
                    if sVals[5] > 0 and sVals[7] > 0 : #section 2 has length and it's tapered
                        fileHeader.append("Bit" + str(iBit) + "_GPTaperLen")
                        d1.append(RoundOff(sVals[5], 2))
                        fileHeader.append("Bit" + str(iBit) + "_GPTaperAng")
                        d1.append(RoundOff(sVals[7], 2))
                    else: # section 2 eith has no length or it's not taper
                        fileHeader.append("Bit" + str(iBit) + "_GPTaperLen")
                        d1.append(0)
                        fileHeader.append("Bit" + str(iBit) + "_GPTaperAng")
                        d1.append(0)

                if sVals[4] == 0 : #no undercut from section 1
                    if sBitSize > 0 and GaugePadNom(sBitSize) > sVals[0] :
                        fileHeader.append("Bit" + str(iBit) + "_GPUcutLen")
                        d1.append(RoundOff(sVals[1] + sVals[5], 2))
                        fileHeader.append("Bit" + str(iBit) + "_GPUcutAmt")
                        d1.append(RoundOff(GaugePadNom(sBitSize) - sVals[0], 2))
                    else:
                        fileHeader.append("Bit" + str(iBit) + "_GPUcutLen")
                        d1.append(0)
                        fileHeader.append("Bit" + str(iBit) + "_GPUcutAmt")
                        d1.append(0)
                    
                else: #Section 2 has an undercut
                    fileHeader.append("Bit" + str(iBit) + "_GPUcutLen")
                    d1.append(RoundOff(sVals[5], 2))
                    fileHeader.append("Bit" + str(iBit) + "_GPUcutAmt")
                    d1.append(RoundOff(sVals[4], 2))

                #Beta
                if sVals[5] == 0 : #section 2 has no length, only use beta from section 1
                    fileHeader.append("Bit" + str(iBit) + "_GPBeta")
                    d1.append(RoundOff(sVals[2], 2))
                elif sVals[2] == sVals[6] :
                    fileHeader.append("Bit" + str(iBit) + "_GPBeta")
                    d1.append(RoundOff(sVals[2], 2))
                else:
                    fileHeader.append("Bit" + str(iBit) + "_GPBeta")
                    d1.append("'" + str(RoundOff(sVals[2], 2)) + " / " + str(RoundOff(sVals[6], 2)))

            else: #if is a DES but has no GaugePad, or is a CFB
                with open(Bit_File_Handler, "r", encoding="gbk", errors="ignore") as des_file:
                    lines = des_file.read().splitlines()
                    find_BasicData = [ string for string in lines if "[BasicData]" in string]
                    if find_BasicData != None:
                        # find_GP = lines.index(find_BasicData[0])
                        fileHeader.append("Bit" + str(iBit) + "_GPBeta")
                        fileHeader.append("Bit" + str(iBit) + "_GPLen")
                        fileHeader.append("Bit" + str(iBit) + "_GPNomLen")
                        fileHeader.append("Bit" + str(iBit) + "_GPTaperLen")
                        fileHeader.append("Bit" + str(iBit) + "_GPTaperAng")
                        fileHeader.append("Bit" + str(iBit) + "_GPUcutLen")
                        fileHeader.append("Bit" + str(iBit) + "_GPUcutAmt")
                        fileHeader.append("Bit" + str(iBit) + "_GPBeta")
                        fileHeader.append("Bit" + str(iBit) + "_GPWidth")
                        i = 1
                        while i <= 9:
                            d1.append("-")
                            i+=1
                
        with open(des_csv_file, "w", encoding="gbk", errors="ignore") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(fileHeader)
            writer.writerow(d1)
            # writer.writerow(d2)
            csv_file.close()
        
        
        
    except ValueError:
        logging.warning('Open failed: ' + Bit_File_Handler)
        return None

    
    
    
    
def process_des(BitNum, Bit_File_Handler):
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
                        print(len(lines[row_index+1].split()))
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
    
if __name__ == "__main__":
    read_bit(r"C:\Users\JWang294\Documents\projectfile\test_config_cases\test-20210416173550-lbo-0416cpc\Static\66048_stinger\66048a00.des", 1)