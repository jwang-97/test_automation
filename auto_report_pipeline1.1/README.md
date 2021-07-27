# Automated_test_report instruction

The Automated_test_report is used to integrate data to generate report ppt.

Its function is mainly to automatically generate a ppt test report based on the calculation results (Static, Dynamic, and Contact)

# Installation and operating Requirements

1. Operation precondition:
* Under Windows
* Python environment
* Office software
* Excel's *macro trust* option is turned on
* Python libraries(TODO:setup.py)
  
2. Project operation steps:
* command run: **python setup.py install** or run: **pip install -r requirements.txt**.
* The two versions of static , dynamic , and Contact calculation results are placed in a general folder.
* The **make_links** script needs to be run when calculating dynamic results.
* Run the main.py script through the command line, etc., and pass in the *three parameters*: cases result path 1, cases result path 2, and ppt template path. You can also pass in *no parameters*, the *default parameters* are written in the first three paths of the .ini file.
* Note that in the current version of the script, the clipboard is occupied due to the call of the vba macro function. Please *do not use the EXCEL, POWERPOINT, and clipboard* on the machine while the script is running.

# Development requirements

* Under Windows
* Office software
* Python environment
* Excel's *macro trust* option is turned on.(https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)

# Code structure description

├── README.md                   // help
├── requirements.txt            // python libraries
├── steup.py                    // python libraries required
├── 60095160_RevA_CPC_Template_v4.04                //ppt template
├── test_automation             // 
│   ├── main.py                 // python main script
│   ├── Multi_Bit_Compare_v2.52.xlsm                // vba script(dynamic)
│   ├── SAM_v11.78.xlsm         // vba script(static)
│   └── test               // test enviroment
├── logs                        //
├── doc                         // config file

# Common problem description

1. Adding the type of static calculation: directly add, the naming rule is the same as the previous one, and it can be overwritten
2. Added type of dynamic calculation example: can include
3. Adding other types of main directories (such as connection, automatic force): Depending on the situation, you may need to modify the code.
4. Required python library (os, sys, re, logging, time, argparse, pptx, pptx.util, glob, re, csv, cv2, io, configparser, win32com.client ,pptx, matplotlib.pyplot, pptx.util, pandas)
5. The section "static_case_type" lists the types of cases which will be generated into the ppt report.
6. If more python libraries needed after upgrated, please list in the requirements.txt, and run: **pip install -r requirements.txt** before test.