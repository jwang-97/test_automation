#import os, sys
#import configParser

##From https://www.pythonf.cn/read/97851
#def ini_write(org_import_dir, config_file_path):
#    cf = configParser.ConfigParser()
#    cf.read(config_file_path)

#    #s = cf.sections()
#    #print 'section:', s

#    #o = cf.options("baseconf")
#    #print 'options:', o

#    #v = cf.items("baseconf")
#    #print 'db:', v
#    #dynamic_path = config_file_path + "\\Dynammic\\links"
#    #if exists(dynamic_path):
#    #    for root,filenames in os.walk(dynamic_path):
#    #        if name == "build":
#    #            dynamic_path = os.path.join(root, name)
#    #            cf.set("bit_base_directory", "bit1_directory", dynamic_path)
#    #            cf.set("bit_base_directory", "bit2_directory", dynamic_path)
#    g = os.walk(org_import_dir)
#    for path, dir_list, file_list in g:
#        for dir_name in dir_list:
#            if dir_name == "E616":
#                for sub_path, sub_dir_list, sub_file_list in os.walk(os.path.join(path, dir_name)):
#                    for file_name in sub_file_list:
#                        if file_name == "Wiley.des":
#                            cf.set("DESfile_E616_path", "cases_path1", os.path.join(sub_pathh, file_name))
#            elif dir_name == "stinger" or "axe" or "pdc" or "px" or "hyper" or "echo":
#                for sub_path, sub_dir_list, sub_file_list in os.walk(os.path.join(path, dir_name)):
#                    for file_name in sub_file_list:
#                        if file_name == ".des":
#                            type = dir_name.split('_')(-1)
#                            cf.set("DESfile_" + type + "_path", "cases_path1", os.path.join(sub_path, file_name))
#        #for file_name in file_list:
#        #    if name  == "pdc":
#        #        for root,filename in os.walk(name):

#            #if name == "stinger:
#            #elif name == "axe":
#            #elif name == "pdc":
#            #elif name == "px":
#            #elif name == "hyper":
#            #elif name == "E616":
#            #elif name == "echo":
            
    
#    config.write(open(config_file_path, "w"))
#    #db_host = cf.get("baseconf", "host")
#    #db_port = cf.getint("baseconf", "port")
#    #db_user = cf.get("baseconf", "user")
#    #db_pwd = cf.get("baseconf", "password")

#    #print db_host, db_port, db_user, db_pwd

#    #cf.set("baseconf", "db_pass", "123456")
#    #cf.write(open("config_file_path", "w"))
#if __name__ == "__main__":
#    ini_write(r'C:\Users\JWang294\Documents\projectfile\static-cases\Static', r'C:\Users\JWang294\Documents\projectfile\auto_test_config.ini')



import os, sys, re, logging, time, argparse
import configparser
import win32com.client as win32
import xlwings as xw


logging.basicConfig(format='%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s',
                    level=logging.DEBUG)

def ini_write(org_import_dir1, org_import_dir2, config_file_path):
    cf = configparser.ConfigParser()
    #with open(config_file_path, 'rb') as f:
    #cf.read(config_file_path, encoding='utf-8', errors='ignore')
    #cf.read(config_file_path, "rb")
    cf.read(config_file_path, encoding='utf-8')
    org_import_dir = (org_import_dir1, org_import_dir2)
    i = 1
    while (i<3):
        g = os.walk(org_import_dir[i-1])
        for path, dir_list, file_list in g:
            for dir_name in dir_list:
                if re.search("E616", dir_name) != None:
                    for sub_path, sub_dir_list, sub_file_list in os.walk(os.path.join(path, dir_name)):
                        for file_name in sub_file_list:
                            if re.search("Wiley.des", file_name) != None:
                                cf.set("DESfile_E616_path", "cases_path"+ str(i), os.path.join(sub_path, file_name))
                                #cf.write(open(config_file_path, "w+", encoding='utf-8'))
                                break
                elif re.search("stinger", dir_name) or re.search("axe", dir_name) or re.search("pdc", dir_name) or re.search("px", dir_name) or  re.search("hyper", dir_name) or  re.search("echo", dir_name) != None:
                    for sub_path, sub_dir_list, sub_file_list in os.walk(os.path.join(path, dir_name)):
                        for file_name in sub_file_list:
                            if re.search('.des', file_name) != None:
                                type_name = dir_name.split('_')
                                cf.set("DESfile_" + type_name[1] + "_path", "cases_path"+ str(i), os.path.join(sub_path, file_name))
                                #cf.write(open(config_file_path, "w+", encoding='utf-8'))
                elif re.search("links", dir_name) != None:
                    for sub_path, sub_dir_list, sub_file_list in os.walk(os.path.join(path, dir_name)):
                        for sub_dir_name in sub_dir_list:
                            while re.match('build', sub_dir_name) != None:
                                #type_name = dir_name.split('_')
                                cf.set("bit_base_directory", "bit"+ str(i) +"_directory", os.path.join(sub_path, sub_dir_name))
                                #cf.write(open(config_file_path, "w+", encoding='utf-8'))
                                bln_build = True
                                break
                            if bln_build:
                                break
                        if bln_build:
                            break
                else:
                    logging.debug(dir_name + ' incompatible')
                    #print(dir_name + ' incompatible')
                    continue
        cf.write(open(config_file_path, "w+", encoding='utf-8'))
        i += 1

def orgpath_sequence(directory1, directory2):
    cf = configparser.ConfigParser()
    cf.read(os.getcwd() + r'\auto_test_config.ini', encoding='utf-8')
    date1 = directory1.split('\\')[-1]
    date2 = directory2.split('\\')[-1]
    if date1.split("-")[1] > date2.split("-")[1]:
        #print(date1.split("-")[1])
        cf.set("profile_title", "title1", str(date2.split("-")[1])[0:8])
        cf.set("profile_title", "title2", str(date1.split("-")[1])[0:8])
        cf.write(open(os.getcwd() + r'\auto_test_config.ini', "w+", encoding='utf-8'))
        return directory2, directory1
    else:
        cf.set("profile_title", "title1", str(date1.split("-")[1])[0:8])
        #print(str(date1.split("-")[1])[0:8])
        cf.set("profile_title", "title2", str(date2.split("-")[1])[0:8])
        cf.write(open(os.getcwd() + r'\auto_test_config.ini', "w+", encoding='utf-8'))
        return directory1, directory2

def write_template_path_ini(tpl_path):
    cf = configparser.ConfigParser()
    cf.read(os.getcwd() + r'\auto_test_config.ini', encoding='utf-8')
    if os.path.exists(tpl_path) == False:
        template_file = os.getcwd() + r'\60095160_RevA_CPC_Template_v4.04.pptx'
        cf.set("templatefile_path", "template_path", template_file)
        logging.debug('Invalid template path, default read template under the current path')
        logging.debug('ppttemplate-path:' + template_file)
    else:
        cf.set("templatefile_path", "template_path", tpl_path)
        logging.debug('ppttemplate-path:' + tpl_path)
    cf.write(open(os.getcwd() + r'\auto_test_config.ini', "w+", encoding='utf-8'))
    return None
    
def call_SAM_macro():
    #try:
        #Launch Excel and Open Wrkbook
        xl=win32.Dispatch("Excel.Application")
        xl.Visible = True
        xl.Application.ScreenUpdating = False
        str = os.getcwd()
        xl.Workbooks.Open(Filename = str + "\\test_SAM_v11.78.xlsm")
        time.sleep(5)
        #xl.ThisWorkbook.Activate
        # xl.Workbooks("Sheet1").Activate
        #xl.Application.ScreenUpdating = True
        logging.debug('Open SAMfile success')
        #Run Macro
        xl.Application.Run('test_SAM_v11.78.xlsm!Module1.Auto_test')
        logging.debug('Auto_test')
        xl.Application.ScreenUpdating = True
      
def call_Multi_macro():
    xl=win32.Dispatch("Excel.Application")
    xl.Visible = True
    xl.Application.ScreenUpdating = False
    str = os.getcwd()
    xl.Workbooks.Open(Filename = str + "\\Multi_Bit_Compare_v2.52.xlsm")
    time.sleep(5)
    logging.debug('Open Multifile success')
    #Run Macro
    xl.Application.Run('Multi_Bit_Compare_v2.52.xlsm!Module1.Auto_Dynamic')
    logging.debug('Auto_Dynamic')
    xl.Application.ScreenUpdating = True
        
if __name__ == "__main__":
    cf = configparser.ConfigParser()
    cf.read(os.getcwd() + r'\auto_test_config.ini', encoding='utf-8')
    logging.debug('config file path:' + os.getcwd() + r'\auto_test_config.ini')
    # org_result_path1, org_result_path2, template_path = input("enter cases-results path1 & cases-result path2 & ppt_template-path separated with ',':").split(',')
    parser = argparse.ArgumentParser(description='input path argparse')
    parser.add_argument('--path1', '-f', help='path1 property, non-essential parameters, have default values', default= "U:\result-backup\test-2021\test-20210416173550-lbo-0416cpc")
    parser.add_argument('--path2', '-s', help='path2 property, non-essential parameters, have default values', default= "U:\result-backup\test-2021\test-20210702174422-lbo-0630cpc")
    parser.add_argument('--template_path', '-t', help='template_path property, non-essential parameters, have default values', default="")
    args = parser.parse_args()
    org_result_path1 = args.path1
    org_result_path2 = args.path2
    template_path = args.template_path
    # org_result_path1 = input("plz enter the first path contain static & dynamic cases results:")
    # org_result_path2 = input("plz enter the second path contain static & dynamic cases results:")
    # template_path = input("plz enter the path of the ppt template:")
    if os.path.exists(org_result_path1) == False:
        org_result_path1 = cf.get("original_path", "org_path1")
        logging.debug('Invalid path1, default read original_path in auto_test_config.ini')

    if os.path.exists(org_result_path2) == False:
        org_result_path2 = cf.get("original_path", "org_path2")
        logging.debug('Invalid path2, default read original_path in auto_test_config.ini')

    logging.debug('path1:' + org_result_path1)
    logging.debug('path2:' + org_result_path2)
    write_template_path_ini(template_path)
    # config_dir = orgpath_sequence(cf.get("original_path", "org_path1"), cf.get("original_path", "org_path2"))
    config_dir = orgpath_sequence(org_result_path1, org_result_path2)
    ini_write(config_dir[0], config_dir[1], os.getcwd() + r'\auto_test_config.ini')
    
    call_SAM_macro()
    call_Multi_macro()
    
    # run_command = "start " + os.getcwd() + r'\test_SAM_v11.78.xlsm'
    # os.system(run_command)
    #ini_write(r'\\bgc-nas002\ideas_engine_folder\Share\Drilling_software_development\Work\Work-2021\release_testing_for_cpc\testing-cases', r'C:\Users\JWang294\Documents\projectfile\auto_test_config.ini')
    #config_sequence(r'\\bgc-nas002\ideas_engine_folder\Share\Drilling_software_development\Work\Work-2021\release_testing_for_cpc\testing-cases', r'\\bgc-nas002\ideas_engine_folder\Share\Drilling_software_development\Work\Work-2021\release_testing_for_cpc\testing-cases')
    
    # C:\Users\JWang294\Documents\projectfile\test_config_cases\test-20210416173550-lbo-0416cpc
    # C:\Users\JWang294\Documents\projectfile\test_config_cases\test-20210702174422-lbo-0630cpc