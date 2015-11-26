#-------------------------------------------------------------------------------
# Name:        sheet_pos
# Purpose:     Get data and post to workbook sheet
#
# Author:      Lai
#
# Created:     28/10/2015
#-------------------------------------------------------------------------------

from xlwings import Workbook, Sheet, Range, Chart
import template_search
import data_manage
import time

# if data item name changed, it can modify the value reflection
item_ref = {"standard":"standard",
            "rate":"rate",
            "channel":"channel",
            "stream":"stream",
            "BW":"BW"
           }

# key: item name in sheet, value: item name in data
sheet_item_ref = {"Tx Power":"power",
                  "EVM":"EVM",
                  "Mask":"Mask",
                  "Freq error":"F_ER",
                  "SC error":"CR_ER",
                  "Flatness":"Flatness",
                  "Rx Power":"SENS"
                 }

rx_item = ["SENS"]

def post_power(data):
    if "power" in data:
        if data["power"]:
            return data["power"].split(",")
        return data["power"]
    return None

def post_evm(data):
    if "EVM" in data:
        if data["EVM"]:
            return data["EVM"].split(",")
        return data["EVM"]
    return None

def post_mask(data):
    if "Mask" in data:
        if data["Mask"]:
            return data["Mask"].split(",")
        return data["Mask"]
    return None

def post_freq_err(data):
    if "F_ER" in data:
        if data["F_ER"]:
            return list([data["F_ER"]])
        return data["F_ER"]
    return None

def post_cr_err(data):
    if "CR_ER" in data:
        if data["CR_ER"]:
            return list([data["CR_ER"]])
        return data["CR_ER"]
    return None

def post_flatness(data):
    if "Flatness" in data:
        if data["Flatness"]:
            return data["Flatness"].split(":")
        return data["Flatness"]
    return None

def post_sens(data):
    if "sens" in data:
        if data["sens"]:
            return data["sens"].split(",")
        return data["sens"]
    return None

post_func= {# TX
            "power":post_power,
            "EVM":post_evm,
            "Mask":post_mask,
            "F_ER":post_freq_err,
            "CR_ER":post_cr_err,
            "Flatness":post_flatness,
            # RX
            "SENS":post_sens
            }

def post(data_path, sheet, sheet_setup, channel_anchor):
    """
    Get template setup to find post position then post value
    """
    fill_pos, all_anchor_row = sheet_setup[0], sheet_setup[1]
##    print fill_pos.keys()
##    print all_anchor_row
    last_data_conf = None
    need_pos = None
    case_num = 0
    ch_start = None
    ch_pos = None
    ch_now = None
    last_ch = 0
    Sheet(sheet).activate()

    for data in data_manage.load_data(data_path):

        if not check_same_row(data, last_data_conf):
##            print data
##            print fill_pos.keys()
            # find "standard" position
            standard_pos = meet_standard(data, fill_pos)
            if not standard_pos:
                continue

            # find "rate" position
            rate_pos = meet_rate(data, standard_pos)
            if rate_pos:
                need_pos, case_num = rate_pos
                last_data_conf = get_data_conf(data)
            else:
                continue
##            print need_pos, case_num
            # Get post start position
            try:
                ch_start = template_search.get_channel_start(sheet,need_pos,all_anchor_row)
            except:
                time.sleep(3)
                ch_start = template_search.get_channel_start(sheet,need_pos,all_anchor_row)
            ch_now = ch_start
##        print ch_start, "ch start"
        # Get value post position
        try:
##            print data[item_ref["channel"]]
            if int(data[item_ref["channel"]]) > last_ch:
                ch_pos = template_search.get_channel_pos(sheet,ch_now,data[item_ref["channel"]])
            else:
                ch_pos = template_search.get_channel_pos(sheet,ch_start,data[item_ref["channel"]])
            last_ch = int(data[item_ref["channel"]])
        except:
            time.sleep(3)
            if int(data[item_ref["channel"]]) > last_ch:
                ch_pos = template_search.get_channel_pos(sheet,ch_now,data[item_ref["channel"]])
            else:
                ch_pos = template_search.get_channel_pos(sheet,ch_start,data[item_ref["channel"]])
            last_ch = int(data[item_ref["channel"]])
##        print ch_pos, "ch pos"
        if ch_pos:
            try:
                post_value(sheet,data,need_pos,ch_pos,case_num)
            except:
                time.sleep(3)
                post_value(sheet,data,need_pos,ch_pos,case_num)
            # if value appear in ch by ch will lose
            ch_now = ch_pos
        else:
            continue

def check_same_row(data, last_data_conf):
    """
    Check if the same sheet row. It will check that "standard", "rate", "band width", and "stream" are all meet
    """
    if not last_data_conf:
        return False

    conf = get_data_conf(data)
    if conf == last_data_conf:
        return True
    return False

def get_data_conf(data):
    return (data["standard"],data["rate"],data["BW"],data["stream"])

def post_value(sheet,data,start,ch_pos,case_num):
    """
    Post value in sheet explicit position
    """
    for i in range(case_num):
        # Get item name
        case = Range(sheet,(start[0]+i,start[1]-1)).value
##        print case
        # if valid item name
        if case in sheet_item_ref:
            value = post_func[sheet_item_ref[case]](data)
##            print data
##            print value
            antennas = data["antenna"].split(",")
##            print antennas, len(antennas)
            if value:
                # value > 1 means multiple streams
                if len(value) > 1:
                    for idx in range(int(data["stream"])):
##                    for idx in range(len(antennas)):
                        post_pos = (start[0]+i,ch_pos[1]+int(antennas[idx]))
                        Range(sheet,post_pos).value = value[idx]
                else:
                    # antennas = 1 and value = 1 will be 1 stream
                    if len(antennas) > 1:
###################### for RX 11ac MIMO and SIMO post #######################
                        if sheet_item_ref[case] in rx_item:
                            move = template_search.find_ch_sum(sheet,ch_pos)
##                            print move
                            post_pos = (start[0]+i,ch_pos[1]+move)
#############################################################################
                        else:
                            # TX no need move
                            post_pos = (start[0]+i,ch_pos[1])
                    else:
                        post_pos = (start[0]+i,ch_pos[1]+int(antennas[0]))
                    Range(sheet,post_pos).value = value[0]

def meet_standard(data,fill_pos):
    """
    Get "standard" in sheet position
    """
############ RX 11n need this to let "stream" to meet really config ###################
############ MIMO will always get stream '1'
    if data[item_ref["standard"]] == "11n":
        ch = int(data[item_ref["rate"]][3:])
        if ch < 8:
            stream = '1'
        elif ch < 16:
            stream = '2'
        elif ch < 24:
            stream = '3'
        else:
            stream = '4'
        k = (data[item_ref["standard"]], data[item_ref["BW"]], stream)
#######################################################################################
    else:
        k = (data[item_ref["standard"]], data[item_ref["BW"]], data[item_ref["stream"]])
##    print k
##    print fill_pos.keys()
    #
    if k in fill_pos:
        return fill_pos[k]

    # MCSx
    if k[0] in fill_pos:
        return fill_pos[k[0]]
    print "Can't find this channel " + data[item_ref["channel"]] + "," + data[item_ref["standard"]]
    return None

def meet_rate(data,fill_pos):
    """
    Get "rate" in sheet position
    """
##    print fill_pos.keys()
##    print data[item_ref["rate"]]
    for k in fill_pos.keys():
        if "-" in data[item_ref["rate"]]:
            if data[item_ref["rate"]] == k:
                return fill_pos[k]
        else:
            if "-" in k:
                #  data[item_ref["rate"]] with modulation
                if data[item_ref["rate"]] == k.split("-")[0]:
                    return fill_pos[k]
                # data[item_ref["rate"]] only have rate
                elif data[item_ref["rate"]] == k.split("-")[1]:
                    return fill_pos[k]
            else:
                if data[item_ref["rate"]] == k:
                    return fill_pos[k]
    print "Can't find this modulation " + data[item_ref["rate"]] + "," + data[item_ref["standard"]]
    return None
