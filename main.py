#-------------------------------------------------------------------------------
# Name:        main
# Purpose:     Open template excel file and read folder to find all data .csv.
#              Then start to call function to post value and when it finish save the report
#
# Author:      Lai
#
# Created:     14/12/2015
#-------------------------------------------------------------------------------
import os, sys
import time
import datetime
from workbook_template import Workbook_Template

# save template sheet setup to avoid setup re-construct. This is for speed up
sheet_setup = {}

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
    standard_anchor = "Standard"
    channel_anchor = "Ch"
    band = find_band(data_name)
    tx_or_rx = find_tx_rx(data_name)
    if not (band and tx_or_rx):
        print data_path + " is NOT a legal data file."
        return 1

    wb.post(data_path, tx_or_rx, band, standard_anchor, channel_anchor)

##    sheet_post.post(data_path,sheet,sheet_setup[tx_or_rx + band], channel_anchor)

def save_report(wb,report_path,date,name):
    """
    Save workbook at report_path. If folder not exist, it will make
    """
    path = os.path.join(report_path,date)
    if not os.path.exists(path):
        os.makedirs(path)
    print os.path.join(path,name)
    wb.save(os.path.join(path,name))

def open_workbook(path):
    """
    Open file
    """
    return Workbook_Template(path)

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
##    if len(sys.argv) == 1:
##        sys.exit(0)

##    template_path = sys.argv[1]
    template_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    template_path = os.path.join(template_path,"tmp.xls")
    wb = open_workbook(template_path)

    print "Start !"
    t1 = time.time()

    for folder in folder_file_names:
        for data_name in folder_file_names[folder]:
            data_path = os.path.join(log_path,folder,data_name)
            print data_path
            post(wb, data_path,data_name)
        save_report(wb,report_path,date,folder)
    ##        Application(wb).screen_updating = True

    print "Finish !"
    print time.time() - t1
    # Not through quit by python itself will let excel file alive in process
##    os.system("pause")

if __name__ == '__main__':
    main()