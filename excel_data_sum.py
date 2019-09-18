# Coltin Lappin-Lux
# lux.coltin@gmail.com
# 9/18/2019

from __future__ import division
from datetime import date
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl import load_workbook

#==================================================================
# Main Function
#==================================================================
class Report:
    title = ''
    preparationDate = ''
    reportType = ''
    reportPeriod = ''

    def __init__(self):
        self.title = 'Boys & Girls Club of Hawaii'
        self.preparationDate = "Report Preparation Date: " + date.today().strftime("%m/%d/%Y")

    def setReportType():
        reportType = "Report Type: " + raw_input("Report Type: ")
        self.reportType = reportType

    def setReportPeriod():
        reportPeriod = "Report Period: " + raw_input("Report Start Date: ") + " - " + raw_input("Report End Date: ")
        self.reportPeriod = reportPeriod

#==================================================================
# Main Function
#==================================================================

def main():
    #==================================================================
    # Load Workbook
    #==================================================================
    #Reading in existing Workbook
    print '\nLoading Workbook ...'
    wbTitle = raw_input("Excel Title: ")
    wbTitle += '.xlsx'
    workBook = load_workbook(filename = wbTitle)

    #==================================================================
    # Load Sheet & Create Summary Sheet
    #==================================================================
    print '\nLoading Sheet ...'
    dataSheetName = raw_input("Sheet Name: ")
    dataSheet = workBook[dataSheetName]
    summarySheetName = dataSheetName + 'Sum'
    summarySheet = workBook.create_sheet(title=summarySheetName)

    #==================================================================
    # Print Title
    #==================================================================
    print '\nPrinting Title ...'
    report = Report()
    report.setReportType()
    report.setReportPeriod()
    print report.reportType
    print report.reportPeriod

    #==================================================================
    # Get User Info
    #==================================================================
    print '\nCollecting Metadata ...'
    reportGroupingVar = raw_input("Grouping Column: ")
    report

    #==================================================================
    # Summarize Data
    #==================================================================
    #==================================================================
    # Print Summary
    #==================================================================
    workBook.save(filename = wbTitle)

    
if __name__ == "__main__":
    main()