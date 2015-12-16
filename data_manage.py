#-------------------------------------------------------------------------------
# Name:        data_manage
# Purpose:     Load format ".csv" data which have explicit arrangement
#
# Author:      Lai
#
# Created:     27/10/2015
#-------------------------------------------------------------------------------
import csv

class Data:
    def __init__(self,path):
        self._path = path
        # item name corresponds to the value management function
        self._get_func = {"standard":self._get_standard,
                    "channel":self._get_channel,
                    "rate":self._get_rate,
                    "BW":self._get_bw,
                    "stream":self._get_stream,
                    "antenna":self._get_antenna,
                    "power":self._get_power,
                    "sens":self._get_sens
                   }

    def load_data(self):
        """
        Load .csv
        """
        with open(self._path, 'rb') as f:  # from python doc. it should open with para. "b"
            reader = csv.reader(f)
            reader.next()   # first line is data front that don't need
            data = {}
            for line in reader:
                line = [i for i in line if i!='']
                if line:
    ##                print line
                    if len(line) > 2:  #line with config
                        if self._get_func["standard"](line):
                            data["standard"] = self._get_func["standard"](line)
                        if self._get_func["channel"](line):
                            data["channel"] = self._get_func["channel"](line)
                        if self._get_func["rate"](line):
                            data["rate"] = self._get_func["rate"](line)
                        if self._get_func["BW"](line):
                            data["BW"] = self._get_func["BW"](line)
                        if self._get_func["stream"](line):
                            data["stream"] = self._get_func["stream"](line)
                        if self._get_func["antenna"](line):
                            data["antenna"] = self._get_func["antenna"](line)
                        # TX power
                        if self._get_func["power"](line):
                            data["power"] = self._get_func["power"](line)

                        # RX SENS
                        if self._get_func["sens"](line):
                            data["sens"] = self._get_func["sens"](line)

                    else: # line with item and value
                        if len(line) > 1:
                            data[line[0]] = line[1]
                        else:
                            data[line[0]] = None
                else:
                    yield data
                    data = {}

    def _get_standard(self,line):
        """
        "standard" explicit position is in list position "1"
        """
        if len(line) > 1:
            return line[1].split(".")[1]
        return None

    def _get_channel(self,line):
        """
        "channel" explicit position is in list position "2"
        """
        if len(line) > 2:
            return line[2]
        return None

    def _get_rate(self,line):
        """
        "rate" explicit position is in list position "3"
        """
        if len(line) > 3:
            return line[3]
        return None

    def _get_bw(self,line):
        """
        "band width" explicit position is in list position "4"
        """
        if len(line) > 4:
            return line[4].split("-")[1]
        return None

    def _get_stream(self,line):
        """
        "stream" explicit position is in list position "5"
        """
        if len(line) > 5:
            return line[5]
        return None

    def _get_antenna(self,line):
        """
        "antenna" explicit position is in list position "6"
        """
        if len(line) > 6:
            return line[6]
        return None

    def _get_power(self,line):
        """
        "tx power" explicit title position is in list position "7" and value position is in list position "8"
        """
        if len(line) > 8:
            if line[7] == "Power":
                return line[8]
        return None

    def _get_sens(self,line):
        """
        "rx power" explicit title position is in list position "7" and value position is in list position "8"
        """
        if len(line) > 8:
            if line[7] == "SENS":
                return line[8]
        return None

if __name__ == '__main__':
    pass

