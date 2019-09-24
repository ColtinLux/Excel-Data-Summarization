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
        self.preparationDate = date.today().strftime("%m/%d/%Y")
    
    def setCompanyTitle(self, companyTitle):
        self.companyTitle = companyTitle
    
    def setReportType(self, reportType):
        self.reportType = reportType
    
    def setReportPeriod(self, reportPeriodStart, reportPeriodEnd):
        self.reportPeriod = reportPeriodStart + ' - ' + reportPeriodEnd
    
    def getHeader(self):
        result = []
        if self.companyTitle != '':
            result.append(self.companyTitle)
        if self.reportType != '':
            result.append(self.reportType)
        if self.reportPeriod != ' - ':
            result.append(self.reportPeriod)
        result.append(self.preparationDate)
        return result

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
            self.primaryColumnTitle = self.dataSheet[self.primaryColumn + '1'].value
            values = []
            unduplicateValues = []
            index = 2
            primaryIndex = self.primaryColumn + str(index)
            while self.dataSheet[primaryIndex].value != None:
                value = self.dataSheet[primaryIndex].value
                values.append(value)
                if unduplicateValues.count(value) == 0:
                    unduplicateValues.append(value)
                index += 1
                primaryIndex = self.primaryColumn + str(index)
            self.result = [[self.primaryColumnTitle]]
            totalCount = len(values)
            totalPercentage = 0.00
            for value in unduplicateValues:
                count = values.count(value)
                percentage = round((count/totalCount),2)
                totalPercentage += percentage
                self.result.append([[value,'Count'],count])
                self.result.append([[value,'Percentage'],percentage])
            self.result.append([['Total','Count'],totalCount])
            self.result.append([['Total','Percentage'],totalPercentage])
            return self.result
        else:
            self.primaryColumnTitle = self.dataSheet[self.primaryColumn + '1'].value
            self.summaryColumnTitle = self.dataSheet[self.summaryColumn + '1'].value
            valuePairs = []
            unduplicateValuePairs = []
            unduplicateKeyValues = []
            summaryValues = []
            unduplicateSummaryValues = []
            index = 2
            primaryIndex = self.primaryColumn + str(index)
            summaryIndex = self.summaryColumn + str(index)
            while self.dataSheet[primaryIndex].value != None:
                key = self.dataSheet[primaryIndex].value
                value =  self.dataSheet[summaryIndex].value
                valuePairs.append([key,value])
                if unduplicateKeyValues.count(key) == 0:
                    unduplicateKeyValues.append(key)
                summaryValues.append(value)
                if unduplicateSummaryValues.count(value) == 0:
                    unduplicateSummaryValues.append(value)
                index += 1
                primaryIndex = self.primaryColumn + str(index)
                summaryIndex = self.summaryColumn + str(index)
            self.result = [[self.summaryColumnTitle, self.primaryColumnTitle]]
            for thisKey in unduplicateKeyValues:
                for thisValue in unduplicateSummaryValues:
                    if unduplicateValuePairs.count([thisKey,thisValue]) == 0:
                        unduplicateValuePairs.append([thisKey,thisValue])
            for pair in unduplicateValuePairs:
                count = valuePairs.count(pair)
                self.result.append([pair,count])
            totalCount = len(valuePairs)
            for value in unduplicateSummaryValues:
                count = summaryValues.count(value)
                percentage = round((count/totalCount),2)
                self.result.append([['Total (#)',value],count])
                self.result.append([['Total (%)',value],percentage])
            return self.result

#==================================================================
# Report
#==================================================================
class Report:
    def __init__(self):
        self.header = []
        self.metadata = Metadata()
        self.content = []
    
    def setHeader(self):
        header = ReportHeader()
        header.setCompanyTitle(raw_input("Company: "))
        header.setReportType(raw_input("Report Type: "))
        header.setReportPeriod(raw_input("Report Start Date: "), raw_input("Report End Date: "))
        self.header = header.getHeader()
    
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
    
    def printReportToConsole(self, withHeader, withContent):
        if withHeader:
            print self.header
        if withContent:
            for content in self.content:
                print ' '
                for data in content:
                    print data
    
    def printReport(self, withHeader, withContent):
        index = 1
        loc = 'A' + str(index)
        if withHeader:
            for item in self.header:
                self.summarySheet[loc] = item
                index += 1
                loc = 'A' + str(index)
            index += 1
            loc = 'A' + str(index)
        number = index
        loc = 'A' + str(number)
        if withContent:
            for contentItem in self.content:
                contentTitle = contentItem.pop(0)
                loc = 'A' + str(number)
                for item in contentTitle:
                    self.summarySheet[loc] = item
                    number += 1
                    loc = 'A' + str(number)
                
                #START
                letter = 'A'
                number = number - 1

                #ROW
                rowNumber = number + 1
                rowLoc = letter + str(rowNumber)

                #COLUMN
                colLetter = 'B'
                colLoc = colLetter + str(number)

                columns = []
                columnIndexList = []
                rows = []
                rowIndexList = []
                
                for data in contentItem:
                    if data[0][0] not in rows:
                        rows.append(data[0][0])
                        rowIndexList.append(rowNumber)
                        self.summarySheet[rowLoc] = data[0][0]
                        rowNumber += 1
                        rowLoc = letter + str(rowNumber)
                    if data[0][1] not in columns:
                        columns.append(data[0][1])
                        columnIndexList.append(colLetter)
                        self.summarySheet[colLoc] = data[0][1]
                        colLetter = chr(ord(colLetter) + 1)
                        colLoc = colLetter + str(number)
                    printNum = rowIndexList[rows.index(data[0][0])]
                    printCol = columnIndexList[columns.index(data[0][1])]
                    printLoc = printCol + str(printNum)
                    self.summarySheet[printLoc] = data[1]
                number = rowNumber + 1


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
    report.setHeader()

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
    report.printReport(True,True)
    
    workBook.save(filename = wbTitle)

    
if __name__ == "__main__":
    main()