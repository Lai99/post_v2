#-------------------------------------------------------------------------------
# Name:        main
# Purpose:     Open template excel file and read folder to find all data .csv.
#              Then start to call function to post value and when it finish save the report
#
# Author:      Lai
#
# Created:     12/15/2015
#-------------------------------------------------------------------------------
import os, sys
import time
import datetime
from workbook_template_xls import Workbook_Template_Xls
from workbook_template_xlsx import Workbook_Template_Xlsx

def get_folder_filenames(path):
    """
    Get all file names with folder under the path
    Input:  string:file_path
    Output: dict:{folder,file_names}
    """
    result = {}
    for parent, dirnames, filenames in os.walk(path):
        if not dirnames:
            result[os.path.split(parent)[-1]] = filenames
    return result

def make_folder(path,folder_names):
    """
    Make folders under the path if the folder dosen't exist
    """
    for folder in folder_names:
        if not os.path.exists(os.path.join(path,folder)):
            os.makedirs(os.path.join(path,folder))

def post(wb, data_path, data_name):
    """
    Get data path and pass to post function to post value
    """
    #initial setup
    standard_anchor = "standard"
##    channel_anchor = "Ch"
    band = find_band(data_name)
    tx_or_rx = find_tx_rx(data_name)
    if not (band and tx_or_rx):
        print data_path + " is NOT a legal data file."
        return 1

    wb.post(data_path, tx_or_rx, band, standard_anchor)

def save_report(wb,report_path,date,name):
    """
    Save workbook at report_path. If folder not exist, it will make
    """
    path = os.path.join(report_path,date)
    if not os.path.exists(path):
        os.makedirs(path)
    print "Save Result " + os.path.join(path,name)
    wb.save(os.path.join(path,name))

def open_workbook(path):
    """
    Open file
    """
    name = path.split(".")[-1].lower()
    if name == "xls":
        return Workbook_Template_Xls(path)
    elif name == "xlsx":
        return Workbook_Template_Xlsx(path)
    else:
        print "Not support this file type. Only *.xls and *xlsx"
        return None

def find_band(data_name):
    """
    Find it is "2.4G" or "5G"
    """
    data_name = data_name.lower()
    if "2g" in data_name:
        return "2G"
    elif "5g" in data_name:
        return "5G"
    return None

def find_tx_rx(data_name):
    """
    Find it is "TX" or "RX"
    """
    data_name = data_name.lower()
    if "tx" in data_name:
        return "TX"
    elif "rx" in data_name:
        return "RX"
    return None

def main():
    # the program path
##    rootdir = os.path.dirname(__file__)
    rootdir = os.path.dirname(os.path.abspath(sys.argv[0]))
    # log folder path
    log_path = os.path.join(rootdir,"Log")
    # report folder path
    report_path = os.path.join(rootdir,"Report")
    # Get all files with dict (folder:file_name)
    folder_file_names = get_folder_filenames(log_path)
    t = time.time()
    # today in year/month/day
    date = datetime.datetime.fromtimestamp(t).strftime(r"%Y%m%d")

    # Need get template path
    if len(sys.argv) == 1:
        sys.exit(0)

    template_path = sys.argv[1]
##    template_path = os.path.dirname(os.path.abspath(sys.argv[0]))
##    template_path = os.path.join(template_path,"NS.xls")

    wb = open_workbook(template_path)
    if not wb:
        return 1

    print "Start !"
    t1 = time.time()

    for folder in folder_file_names:
        try:
            for data_name in folder_file_names[folder]:
                data_path = os.path.join(log_path,folder,data_name)
                print data_path
                post(wb, data_path,data_name)
        except:
            print "Unexpected error:", sys.exc_info()
        save_report(wb,report_path,date,folder)

    print "Finish !"
    print time.time() - t1
    os.system("pause")

if __name__ == '__main__':
    main()