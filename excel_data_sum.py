# Coltin Lappin-Lux
# lux.coltin@gmail.com
# 9/18/2019

from __future__ import division
from datetime import date
#import numpy as np
#import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl import load_workbook

#==================================================================
# ReportHeader
#==================================================================
class ReportHeader:
    def __init__(self):
        self.companyTitle = ''
        self.reportType = ''
        self.reportPeriod = ''
        self.preparationDate = "Report Preparation Date: " + date.today().strftime("%m/%d/%Y")
    
    def setCompanyTitle(self, companyTitle):
        self.companyTitle = companyTitle
    
    def setReportType(self, reportType):
        self.reportType = reportType
    
    def setReportPeriod(self, reportPeriod):
        self.reportPeriod = reportPeriod

#==================================================================
# Metadata
#==================================================================
class Metadata:
    def __init__(self):
        self.primaryAttribute = ''
        self.summaryAttributes = ''
    
    def setPrimaryAttribute(self, column):
        self.primaryAttribute = column.upper()

    def setSummaryColumns(self, columns):
        if(columns.isupper()):
            self.summaryColumns = columns.split()
        else:
            columnsUpper = columns.upper()
            self.summaryColumns = columnsUpper.split()
    
#==================================================================
# ReportContent
#==================================================================
class ReportContent:
    def __init__(self, primaryColumn, summaryColumn, dataSheet):
        self.primaryColumn = primaryColumn
        self.summaryColumn = summaryColumn
        self.dataSheet = dataSheet        
    
    def generateResults(self):
        if(self.primaryColumn == self.summaryColumn):
            return 'Same'
        else:
            self.primaryColumnTitle = self.dataSheet[self.primaryColumn + '1'].value
            self.summaryColumnTitle = self.dataSheet[self.summaryColumn + '1'].value
            return 'Not Same'

#==================================================================
# Report
#==================================================================
class Report:
    def __init__(self):
        self.header = ReportHeader()
        self.metadata = Metadata()
        self.content = []
    
    def setHeader(self):
        self.header.setCompanyTitle(raw_input("Company: "))
        self.header.setReportType("Report Type: " + raw_input("Report Type: "))
        self.header.setReportPeriod("Report Period: " + raw_input("Report Start Date: ") + " - " + raw_input("Report End Date: "))
    
    def setMetadata(self):
        self.metadata.setPrimaryAttribute(raw_input("Primary Attribute Column: "))
        self.metadata.setSummaryColumns(raw_input("Columns to Summarize: "))
    
    def setDataSheet(self, dataSheet):
        self.dataSheet = dataSheet
    
    def setSummarySheet(self, summarySheet):
        self.summarySheet = summarySheet

    def generateContent(self):
        for col in self.metadata.summaryColumns:
            columnContent = ReportContent(self.metadata.primaryAttribute, col, self.dataSheet)
            results = columnContent.generateResults()
            self.content.append(results)
        print self.content



#==================================================================
# Main Function
#==================================================================

def main():
    #==================================================================
    # Load Excel Workbook
    #==================================================================
    #Reading in existing Workbook
    print '\nLoading Excel Workbook ...'
    wbTitle = raw_input("Excel Title: ")
    wbTitle += '.xlsx'
    workBook = load_workbook(filename = wbTitle)

    #==================================================================
    # Load Excel Sheet & Create Summary Sheet
    #==================================================================
    print '\nLoading Excel Sheet ...'
    dataSheetName = raw_input("Sheet Name: ")
    dataSheet = workBook[dataSheetName]
    summarySheetName = dataSheetName + 'Sum'
    summarySheet = workBook.create_sheet(title=summarySheetName)

    #==================================================================
    # Create Report Object
    #==================================================================
    print '\nLoading Report ...'
    report = Report()
    report.setDataSheet(dataSheet)
    report.setSummarySheet(summarySheet)

    #==================================================================
    # Create Report Object & Load Report Header
    #==================================================================
    print '\nLoading Report Header ...'
    #report.setHeader()

    #==================================================================
    # Load Report Metadata
    #==================================================================
    print '\nLoading Report Metadata ...'
    report.setMetadata()

    #==================================================================
    # Summarize Data
    #==================================================================
    print '\nGenerating Report ...'
    report.generateContent()

    #==================================================================
    # Save Workbook
    #==================================================================
    workBook.save(filename = wbTitle)

    
if __name__ == "__main__":
    main()