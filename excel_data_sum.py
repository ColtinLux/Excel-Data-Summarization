# Coltin Lappin-Lux (Data Analyst)
# lux.coltin@gmail.com
# Boys and Girls Club of Hawaii
# 1/25/18

from __future__ import division
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl import load_workbook


#===============================================================================
# Attribute List (Both Duplicated and Unduplicated) - List without blank cells
#===============================================================================

def att_list(sheet, col, row):
    #Where to find attribute (Column and Row) Ex. 'B', 2
    duplicate_list = []
    unduplicated_list = []

    index = row
    i = col + str(index)
    while sheet[i].value != None:
        #print(sheet[i].value)
        i = col + str(index)
        if sheet[i].value != None:
            duplicate_list.append(sheet[i].value)
        index += 1

    for i in duplicate_list:
        if unduplicated_list.count(i) == 0:
            unduplicated_list.append(i)

    unduplicated_list.sort()

    return unduplicated_list, duplicate_list

#==================================================================================
# Attribute List 2 (Both Duplicated and Unduplicated) - List with blank cells
#==================================================================================

def att_list_2(sheet, list_length, col, row):
    #Where to find attribute (Column and Row) Ex. 'B', 2
    duplicate_list = []
    unduplicated_list = []

    #attribute counter
    att_index = row
    att_i = col + str(att_index)

    #primary attribute counter
    pri_index = 2
    pri_i = 'B' + str(pri_index)

    for x in range(list_length):
        att_i = col + str(att_index)
        pri_i = 'B' + str(pri_index)
        if sheet[att_i].value == None:
            duplicate_list.append([sheet[pri_i].value,"Unknown"])
        else:
            duplicate_list.append([sheet[pri_i].value,sheet[att_i].value])
        att_index += 1
        pri_index += 1


    #attribute counter
    att_index = row
    att_i = col + str(att_index)

    for z in range(list_length):
        att_i = col + str(att_index)
        if unduplicated_list.count(sheet[att_i].value) == 0:
            if sheet[att_i].value == None:
                if unduplicated_list.count("Unknown") == 0:
                    unduplicated_list.append("Unknown")
            else:
                unduplicated_list.append(sheet[att_i].value)
        att_index += 1

    unduplicated_list.sort()

    #print(unduplicated_list)
    return unduplicated_list, duplicate_list

#===============================================================================
# Count
#===============================================================================

def count(what, where):
    #what = what to count (List)
    #where = where to find it (Location)

    results = []

    for x in what:
        results.append([x,where.count(x)])

    return results

#===============================================================================
# Attribute Count
#===============================================================================

def attribute_count(what, where):
    #what = what to count (List)
    #where = where to find it (Location)

    results = []
    count = 0

    for x in what:
        count = 0
        for y in where:
            if x == y[1]:
                count += 1
        results.append([x,count])

    return results

#===============================================================================
# Attribute Percentage
#===============================================================================

def attribute_percentage(att_count):
    #what = what to count (List)
    #where = where to find it (Location)

    results = []
    total = 0
    temp = 0
    for i in range(len(att_count)):
        total = total + att_count[i][1]
    for i in range(len(att_count)):
        temp = att_count[i][1]/total
        results.append([att_count[i][0],temp])

    return results

#===============================================================================
# Count By Club
#===============================================================================

def countByClub(bgch_list, undup, dup):
    #what = what to count (List)
    #where = where to find it (Location)
    #[clubhouse, attribute, count]

    results = []

    for x in bgch_list:
        for y in undup:
            count = 0
            for z in dup:
                if x == z[0] and y == z[1]:
                    count += 1
            results.append([x,y,count])

    return results

#===============================================================================
# Total Count
#===============================================================================

def total(what):
    #what = what to count (List)

    results = len(what)

    return results

#==================================================================
# Main Function (Initialize Variables and Call Functions)
#==================================================================

def main():
    """
    This function executes when this file is run as a script.
    """

    #==================================================================
    # Open Workbooks (Input and Output Excel Sheets)
    #==================================================================
    #Reading in existing Workbook
    title = raw_input("Input Excel Title: ")
    title = title + '.xlsx'

    wb = load_workbook(filename = title)

    #Creating Revised Workbook
    rwb = Workbook()
    dest_filename = raw_input("Output Excel Title: ")
    dest_filename = dest_filename + '.xlsx'

    #==================================================================
    # While loop in case user wants to summarize multiple excel sheets
    #==================================================================
    #Data Calculation
    continuetill = 'yes'
    while continuetill == 'yes':
        sheetname = raw_input("Input Sheet Name/Number: ")
        sheet = wb[sheetname]
        revisedSheet = rwb.active
        revisedSheet = rwb.create_sheet(title=sheetname)

        #==================================================================
        # Summarize Membership Data
        #==================================================================
        #Primary Attribute - Clubhouse
        bgch_clubhouse_list, duplicated_bgch_clubhouse_list = att_list(sheet, 'B', 2)
        #print(bgch_clubhouse_list)

        #Count how many members per clubhouse
        membership_count = count(bgch_clubhouse_list, duplicated_bgch_clubhouse_list)
        membership_total = total(duplicated_bgch_clubhouse_list)
        list_length = len(duplicated_bgch_clubhouse_list)
        membership_count.append(["Total", membership_total])

        #==================================================================
        # Print Title and other document details to new Excel
        #==================================================================
        #Title
        title_index = 1
        title_loc = 'A' + str(title_index)
        revisedSheet[title_loc] = 'Boys & Girls Club of Hawaii'
        #Report Type
        title_index = 2
        title_loc = 'A' + str(title_index)
        revisedSheet[title_loc] = 'Report Type:'
        #Period
        title_index = 3
        title_loc = 'A' + str(title_index)
        revisedSheet[title_loc] = 'Period:'
        #Description
        title_index = 4
        title_loc = 'A' + str(title_index)
        revisedSheet[title_loc] = 'Preparation Date: '

        #==================================================================
        # Print Membership Summary to new Excel
        #==================================================================
        #Membership Data
        club_index = 6
        club_i = 'A' + str(club_index)
        revisedSheet[club_i] = 'Site'

        count_index = 6
        count_i = 'B' + str(count_index)
        revisedSheet[count_i] = 'Membership Count'

        perc_index = 6
        perc_i = 'C' + str(perc_index)
        revisedSheet[perc_i] = 'Membership Percentage'

        for x in membership_count:
            for i in range(len(x)):
                if i == 0:
                    club_index += 1
                    club_i = 'A' + str(club_index)
                    revisedSheet[club_i] = x[i]
                elif i == 1:
                    count_index += 1
                    count_i = 'B' + str(count_index)
                    revisedSheet[count_i] = x[i]
                    perc_index += 1
                    perc_i = 'C' + str(perc_index)
                    revisedSheet[perc_i] = x[i]/membership_total

        #==================================================================
        # Summarizing Attributes
        #==================================================================
        att_continuetill = raw_input("Is there another attribute to summarize? (yes/no) ")
        while att_continuetill == 'yes':

            att_col = raw_input("Attribute Column (Letter): ")
            att_row = 2

            #==================================================================
            # Summarize 1 Attribute
            #==================================================================
            att, dupAtt = att_list_2(sheet, list_length, str(att_col), int(att_row))
            #totals by primary att
            att_countbyclub = countByClub(bgch_clubhouse_list, att, dupAtt)
            #totals
            att_count = attribute_count(att, dupAtt)
            att_percentage = attribute_percentage(att_count)

            #==================================================================
            # Print Attribute Results to New Excel Sheet
            #==================================================================
            #Attribute Title
            club_index += 2
            att_index = club_index + 1
            club_i = 'A' + str(club_index)
            revisedSheet[club_i] = sheet[att_col + str(1)].value

            #==================================================================
            # Hard coded and Calculated by Primary Attribute (Site)
            #==================================================================
            club_index += 1
            club_i = 'A' + str(club_index)
            revisedSheet[club_i] = 'Site'

            for x in bgch_clubhouse_list:
                club_index += 1
                club_i = 'A' + str(club_index)
                revisedSheet[club_i] = x
            
            #Print Total(#)
            counter = 1
            att_revised_row = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ']
            att_i = att_revised_row[counter] + str(att_index)

            club_index += 1
            club_i = 'A' + str(club_index)
            revisedSheet[club_i] = "Total (#)"
            for i in range(len(att_count)):
                att_i = att_revised_row[counter] + str(club_index)
                revisedSheet[att_i] = att_count[i][1]
                counter += 1

            #Print Total(%)
            counter = 1
            att_revised_row = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ']
            att_i = att_revised_row[counter] + str(att_index)

            club_index += 1
            club_i = 'A' + str(club_index)
            revisedSheet[club_i] = "Total (%)"
            for i in range(len(att_percentage)):
                att_i = att_revised_row[counter] + str(club_index)
                revisedSheet[att_i] = att_percentage[i][1]
                counter += 1

            #Print Attribute Titles
            counter = 1
            att_revised_row = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ']
            att_i = att_revised_row[counter] + str(att_index)

            for x in att:
                att_i = att_revised_row[counter] + str(att_index)
                revisedSheet[att_i] = x
                counter += 1

            #Print Attribute Summary Data
            att_revised_row = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ']

            att_printer_counter = 0
            for x in range(len(bgch_clubhouse_list)):
                counter = 1
                att_index += 1
                att_i = att_revised_row[counter] + str(att_index)
                for y in range(len(att)):
                    revisedSheet[att_i] = att_countbyclub[att_printer_counter][2]
                    #print(att_countbyclub[att_printer_counter][2])
                    att_printer_counter += 1
                    counter += 1
                    att_i = att_revised_row[counter] + str(att_index)






            att_continuetill = raw_input("Is there another attribute to summarize? (yes/no) ")

        continuetill = raw_input("Is there another sheet? (yes/no) ")

    #Done, Save Revised Workbook
    rwb.save(filename = dest_filename)


    
if __name__ == "__main__":
    main()