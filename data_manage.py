#-------------------------------------------------------------------------------
# Name:        data_manage
# Purpose:     Load format ".csv" data which have explicit arrangement
#
# Author:      Lai
#
# Created:     27/10/2015
#-------------------------------------------------------------------------------
import csv

def get_standard(line):
    """
    "standard" explicit position is in list position "1"
    """
    if len(line) > 1:
        return line[1].split(".")[1]
    return None

def get_channel(line):
    """
    "channel" explicit position is in list position "2"
    """
    if len(line) > 2:
        return line[2]
    return None

def get_rate(line):
    """
    "rate" explicit position is in list position "3"
    """
    if len(line) > 3:
        return line[3]
    return None

def get_bw(line):
    """
    "band width" explicit position is in list position "4"
    """
    if len(line) > 4:
        return line[4].split("-")[1]
    return None

def get_stream(line):
    """
    "stream" explicit position is in list position "5"
    """
    if len(line) > 5:
        return line[5]
    return None

def get_antenna(line):
    """
    "antenna" explicit position is in list position "6"
    """
    if len(line) > 6:
        return line[6]
    return None

def get_power(line):
    """
    "tx power" explicit title position is in list position "7" and value position is in list position "8"
    """
    if len(line) > 8:
        if line[7] == "Power":
            return line[8]
    return None

def get_sens(line):
    """
    "rx power" explicit title position is in list position "7" and value position is in list position "8"
    """
    if len(line) > 8:
        if line[7] == "SENS":
            return line[8]
    return None

# item name corresponds to the value management function
get_func = {"standard":get_standard,
            "channel":get_channel,
            "rate":get_rate,
            "BW":get_bw,
            "stream":get_stream,
            "antenna":get_antenna,
            "power":get_power,
            "sens":get_sens
           }

def load_data(path):
    """
    Load .csv
    """
    data = {}
    with open(path, 'rb') as f:  # from python doc. it should open with para. "b"
        reader = csv.reader(f)
        reader.next()   # first line is data front that don't need
        for line in reader:
            line = [i for i in line if i!='']
            if line:
##                print line
                if len(line) > 2:  #line with config
                    if get_func["standard"](line):
                        data["standard"] = get_func["standard"](line)
                    if get_func["channel"](line):
                        data["channel"] = get_func["channel"](line)
                    if get_func["rate"](line):
                        data["rate"] = get_func["rate"](line)
                    if get_func["BW"](line):
                        data["BW"] = get_func["BW"](line)
                    if get_func["stream"](line):
                        data["stream"] = get_func["stream"](line)
                    if get_func["antenna"](line):
                        data["antenna"] = get_func["antenna"](line)
                    # TX power
                    if get_func["power"](line):
                        data["power"] = get_func["power"](line)

                    # RX SENS
                    if get_func["sens"](line):
                        data["sens"] = get_func["sens"](line)

                else: # line with item and value
                    if len(line) > 1:
                        data[line[0]] = line[1]
                    else:
                        data[line[0]] = None
            else:
                yield data
                data = {}

