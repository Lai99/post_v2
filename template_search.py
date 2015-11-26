#-------------------------------------------------------------------------------
# Name:        template_search
# Purpose:
#
# Author:      Lai
#
# Created:     11/26/2015
#-------------------------------------------------------------------------------
import xlwt, xlrd, openpyxl
from xlutils.copy import copy

##def manage_standard_5G(sheet,pos):
##    """
##    Draw "5G standard" express from template
##    """
##    s = Range(sheet,pos).value
##    if "\n" in s:
##        standard_rate, stream = s.split("\n")
##        standard_rate = standard_rate.replace(" ","")
##        standard, rate = standard_rate.split("-")
####        print standard_rate
##        stream = stream.split(" ")[0]
##        rate = rate.split("T")[-1]
####        print (standard,rate,stream)
##        return (standard,rate,stream)
##    else:
####        print s
####************************************************
#### for match data "11a" actually got "11ag
##        if s == "11a":
##            s = "11ag"
####************************************************
##        return s
##
##def manage_standard_2G(sheet,pos):
##    """
##    Draw "2.4G standard" express from template
##    """
##    s = Range(sheet,pos).value
##    if "\n" in s:
##        standard_rate, stream = s.split("\n")
##        standard_rate = standard_rate.strip()
##        standard, rate = standard_rate.split(" ")
##        stream = stream.split(" ")[0]
##        rate = rate.split("M")[0]
####        print (standard,rate,stream)
####************************************************
#### from sheet will get "11gac" but data is "11ac"
##        if standard == "11gac":
##            standard = "11ac"
####************************************************
##        return (standard,rate,stream)
##    else:
####        print s
####************************************************
#### for match data "11g" actually got "11ag
##        if s == "11g":
##            s = "11ag"
####************************************************
##        return s
##
##def manage_modulation(sheet,pos):
##    """
##    Draw "modulation" express from template
##    """
####************************************************
#### for match 2.4G 11b "DSSS-CCK". data will get "CCK" instead of "DSSS"
##    if "CCK" in Range(sheet,pos).value:
##        modulation = Range(sheet,pos).value.split("-")[1]
####************************************************
##    else:
##        modulation = Range(sheet,pos).value.split("-")[0]
##    return modulation.strip()
##
##def mange_rate(sheet,pos):
##    pass
##
##def make_module_item_key(sheet, pos, offset):
##    """
##    Add modulation and rate to a string
##    """
##    if Range(sheet,(pos[0],offset)).value:
##        rate = str(Range(sheet,(pos[0],offset)).value)
##        # the rate from sheet will be float
##        if rate.split(".")[1] == '0':
##            rate = rate.split(".")[0]
##        else:
##            rate = rate.replace(".","_")
##        return manage_modulation(sheet,pos) + "-" + rate
##    else:
##        return manage_modulation(sheet,pos)
##
##standard_manage_func = {"2G":manage_standard_2G,
##                        "5G":manage_standard_5G
##                        }
##
##def get_fill_pos(sheet,anchor,band,standard_x = 1,module_x = 2,rate_x = 3, case_x = 5, start_x = 6):
##    """
##    Get all value can be filled position in a sheet
##    Input: int:specified sheet, string:anchor which used to split data block, int:band, int:spec column position, int: modulation column position,
##           int:data rate column position, int:test items column position
##    Output:dict:whole sheet value can be filled position, all anchors row loocation
##    """
##    start = 0
##    all_anchor_row = []
##    #Don't need sheet front content. Use anchor to go to standard start position
##    for row in range(1,50):
##        if Range(sheet,(row,standard_x)).value == anchor:
##            start = row
##            all_anchor_row.append(row)
##            break
##    last_standard = (0,0)
##    last_module = (0,0)
##    items = {}
##    module_items = {}
##    case_count = 0
##    #Don't no when to end. Need a value as bound ex:1500
##    for row in range(start+1,1500):
##        #Use standard between standard to split data block, need to add last_standard one in the end
##        if Range(sheet,(row,module_x)).value != None:    #Collect Modulations in a standard
##            if case_count != 0:
##                #Add "module and rate" with "value start position and case numbers"
##                k = make_module_item_key(sheet, last_module, rate_x)
##                module_items[k] = ((last_module[0], start_x),case_count)
##                case_count = 0
##            last_module = (row,module_x)
##
##        if Range(sheet,(row,standard_x)).value != None:
##            if  Range(sheet,(row,standard_x)).value != anchor:
##                if module_items:   #if true means it has a modulation collection
####************************************************************************************************
####  5G sheet has a sheet tail. When reach this, stop search and record "standard"
##                    if Range(sheet,(row,standard_x)).value == "Info":
##                        break
####************************************************************************************************
##                    items[standard_manage_func[band](sheet,last_standard)] = module_items
####                    Range(sheet,(last_standard[0],12)).value = module_items.keys()
####                    print module_items.values()
##                    module_items = {}
##                last_standard = (row,standard_x)  #A spec start position
##            else:
##                all_anchor_row.append(row)
##                continue   #Not include row which has anchor
##
##        if Range(sheet,(row,case_x)).value != None:  #Count how many test case
##            case_count += 1
##
##    #Don't forget last one have no end point
##    if case_count != 0:
##        #Add "module and rate" with "value start position and case numbers"
##        k = make_module_item_key(sheet, last_module, rate_x)
##        module_items[k] = ((last_module[0], start_x),case_count)
##
##    if module_items:
##        items[standard_manage_func[band](sheet,last_standard)] = module_items
####        Range(sheet,(last_standard[0],12)).value = module_items.keys()
####        print module_items.values()
##    return items, all_anchor_row
##
##def get_channel_start(sheet, pos, all_anchor_row = None):
##    """
##    Search row in all_anchor_row that closest to pos. The row have the channel information
##    """
##    row, col = pos[0], pos[1]
##    if all_anchor_row:
##        if row - all_anchor_row[0] > 0:
##            if len(all_anchor_row) == 1:
##                return (all_anchor_row[0],col)
##            else:
##                for i in range(len(all_anchor_row)):
##                    if row - all_anchor_row[i] < 0:
##                        return (all_anchor_row[i-1],col)
##                return (all_anchor_row[-1],col)
##    print "Can't find channel form location in sheet"
##    return None
##
##def get_channel_pos(sheet, pos, ch):
##    """
##    Search column to find the channel title position
##    """
##    row, col = pos[0], pos[1]
##    count = 31
##    # beacause it might have blank, need to pass
##    while count > 0:
##        while Range(sheet,(row,col)).value:
##            if ch in Range(sheet,(row,col)).value:
##                return (row, col)
##            col += 1
##            count = 31
##        col += 1
##        count -= 1
##    print "Can't find this channel in channel form " + str(ch) + " , "+ str(pos)
##    return None
##
##def find_ch_sum(sheet,ch_pos):
##    """
##    Search column to find the last block value meet value in ch_pos
##    """
##    count = 0
##    match = (Range(sheet, ch_pos).value).replace(" ","")
##    while Range(sheet, (ch_pos[0],ch_pos[1]+count)).value and (Range(sheet, (ch_pos[0],ch_pos[1]+count)).value).replace(" ","") == match:
##        count += 1
##    return count - 1

class _Abstract_Workbook_Template:
    _sheet_setup = {}

    def __init__(self, path):
        self._rb = xlrd.open_workbook(path,formatting_info=True)
        self._sheet_arrange = self._get_sheet_arrange(self._rb)
        self._wb = copy(self._rb)

    def __get__(self):
        return self._sheet_arrange

    def _get_sheet_arrange(self):
        pass

    def _get_items_pos(self,sheet,anchor,items):
        pass

    def get_fill_pos(self, mode, band, anchor):
        pass

class Workbook_Template(_Abstract_Workbook_Template):
    def _get_sheet_arrange(self,workbook):
        """
        To find sheet name "TX / RX" and "2.4G / 5G" and make a dict (name:sheet_pos)
        """
        sheet_names = [i.lower() for i in workbook.sheet_names()]
        sheet_ref = {}
        for idx in range(len(sheet_names)):
            if "2.4ghz" in sheet_names[idx]:
                if "tx" in sheet_names[idx]:
                    sheet_ref["TX2G"] = idx
                elif "sensitivity" in sheet_names[idx]:
                    sheet_ref["RX2G"] = idx
            elif "5ghz" in sheet_names[idx]:
                if "tx" in sheet_names[idx]:
                    sheet_ref["TX5G"] = idx
                elif "sensitivity" in sheet_names[idx]:
                    sheet_ref["RX5G"] = idx
        return sheet_ref

    def _get_items_pos(self,mode):
        """
        Recall template setup if it exist. If template setup isn't exist, call "get_fill_pos" to make.
        """
        if mode == "TX":
            # standard_x = 1,module_x = 2,rate_x = 3, case_x = 5, start_x = 6
            return 1,2,3,5,6
        else:
            # standard_x = 1,module_x = 2,rate_x = 3, case_x = 6, start_x = 7
            return 1,2,3,5,7

    def get_fill_pos(self, mode, band, anchor):
        """
        Get all value can be filled position in a sheet
        Input: int:specified sheet, string:anchor which used to split data block, int:band
        Output:dict:key:whole sheet value can be filled position, value:all anchors row loocation
        """
        sheet = self._rb.sheet_by_index(self._sheet_arrange[mode+band])
        standard_x ,module_x ,rate_x , case_x, start_x = _get_items_pos(mode)

        start = 0
        all_anchor_row = []
        #Don't need sheet front content. Use anchor to go to standard start position
        for row in range(1,50):
            if Range(sheet,(row,standard_x)).value == anchor:
                start = row
                all_anchor_row.append(row)
                break
        last_standard = (0,0)
        last_module = (0,0)
        items = {}
        module_items = {}
        case_count = 0
        #Don't no when to end. Need a value as bound ex:1500
        for row in range(start+1,1500):
            #Use standard between standard to split data block, need to add last_standard one in the end
            if Range(sheet,(row,module_x)).value != None:    #Collect Modulations in a standard
                if case_count != 0:
                    #Add "module and rate" with "value start position and case numbers"
                    k = make_module_item_key(sheet, last_module, rate_x)
                    module_items[k] = ((last_module[0], start_x),case_count)
                    case_count = 0
                last_module = (row,module_x)

            if Range(sheet,(row,standard_x)).value != None:
                if  Range(sheet,(row,standard_x)).value != anchor:
                    if module_items:   #if true means it has a modulation collection
    ##************************************************************************************************
    ##  5G sheet has a sheet tail. When reach this, stop search and record "standard"
                        if Range(sheet,(row,standard_x)).value == "Info":
                            break
    ##************************************************************************************************
                        items[standard_manage_func[band](sheet,last_standard)] = module_items
    ##                    Range(sheet,(last_standard[0],12)).value = module_items.keys()
    ##                    print module_items.values()
                        module_items = {}
                    last_standard = (row,standard_x)  #A spec start position
                else:
                    all_anchor_row.append(row)
                    continue   #Not include row which has anchor

            if Range(sheet,(row,case_x)).value != None:  #Count how many test case
                case_count += 1

        #Don't forget last one have no end point
        if case_count != 0:
            #Add "module and rate" with "value start position and case numbers"
            k = make_module_item_key(sheet, last_module, rate_x)
            module_items[k] = ((last_module[0], start_x),case_count)

        if module_items:
            items[standard_manage_func[band](sheet,last_standard)] = module_items
    ##        Range(sheet,(last_standard[0],12)).value = module_items.keys()
    ##        print module_items.values()
        return items, all_anchor_row