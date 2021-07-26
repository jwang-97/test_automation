import os, sys, re, logging, time, argparse, pptx, pptx.util, glob, re, csv, cv2, io
import configparser
import win32com.client as win32
# import xlwings as xw
from pptx import Presentation
import matplotlib.pyplot as plt
# import numpy as np
from pptx.util import Cm
import pandas as pd

def cases_types_config(config_file_path):
    cf = configparser.ConfigParser()
    #with open(config_file_path, 'rb') as f:
    #cf.read(config_file_path, encoding='utf-8', errors='ignore')
    #cf.read(config_file_path, "rb")
    cf.read(config_file_path, encoding='utf-8')
    # print(cf.sections())
    type_list = []
    cases_types = cf.sections()
    for type in cases_types:
        if re.search('DESfile', type) != None and re.search('path', type) != None:
            type_list.append(type)
    # print(type_list)
    str_type = ' '.join(str(type) for type in type_list)
    cf.set("static_case_type", "type_list", str_type)
    cf.write(open(config_file_path, "w+", encoding='utf-8'))
if __name__ == "__main__" :
    config_dir = os.getcwd() + r"\auto_test_config.ini"
    cases_types_config(config_dir)