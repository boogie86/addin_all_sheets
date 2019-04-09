import ctypes
import time
import datetime
import math
import os.path
import glob
import addin_all_sheets.listLoader
from collections import Counter



#Function that checks whether a date is bigger than another one
def checkDate1BiggerOrEqualDate2(sheet, date1str, date2str, display_name_date_1, display_name_date_2, date_format, idx, display_popup):
    try:
        if datetime.datetime.strptime(date1str,date_format) < datetime.datetime.strptime(date2str,date_format):
            line = idx + 1 #because first index is always 0
            errorIndex = line + 5 #because first 5 rows in excel are not actual data 
            message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + display_name_date_1 + " must be bigger or equal to " + display_name_date_2 + "! Check your data."
            
            if display_popup == True: 
                
                RaisePopup(message) 
                #logValidationError(message)
                
        else:
            message = ""
            
        return message
    
    except ValueError:
        line = idx + 1 #because first index is always 0
        errorIndex = line + 5 #because first 5 rows in excel are not actual data 
        
        #error = "line " + str(errorIndex) + ": " + " Date must be in YYYY-MM-DD format! Check your data."
        #RaisePopup(error)
      
#Function that checks column against forbidden values
#The 2nd input parameter - forbiddenValue - defines your forbidden value 
def containsForrbidenValue(sheet, column, forbiddenValue, display_name, idx, display_popup):
    if  column == str(forbiddenValue):
                line = idx + 1 #because first index is always 0
                errorIndex = line + 5 #because first 5 rows in excel are not actual data 
                message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + str(display_name) + " cannot contain " + "'" + str(forbiddenValue) + "'" + "! Check your data."
                
                if display_popup == True: 
                    
                    RaisePopup(message)  
                    #logValidationError(message)
                    
                return message
    else:
        message = ""
        
    return message

#Function that checks column against desired values
#The 2nd input parameter - wantedValue - defines your wanted value 
def containsWantedValue(sheet, column, wantedValue, display_name, idx, display_popup):
    if  column != str(wantedValue):
                line = idx + 1 #because first index is always 0
                errorIndex = line + 5 #because first 5 rows in excel are not actual data 
                message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + str(display_name) + " has to be set to " + "'" + str(wantedValue) + "'" + "! Check your data."
                
                if display_popup == True: 
                    
                    RaisePopup(message)  
                    #logValidationError(message)
                
    else:
        message = ""
        
    return message

#Function that sets a newline between log entries from different executions
def delimitLogEntries(wb):
    
    GenInf = wb["General Information"]
    strOutputDir = str(GenInf['C8'].value)
    outputfileName = "log_alfin_cashflow_validation_errors.txt"
    path =  strOutputDir + '\\' + outputfileName
    
    fileExists = os.path.isfile(path)
    
    if fileExists:
    
        with open(path, 'r+') as f:
            f.seek(0, 0)
            f.write('\n')
            
    else:
        open(path,"w+")
        
#Function that checks if there are duplicate elements in a list
def findDuplicates(wb, sheet, input_list, column_index, display_name, display_popup):
        
    message = ""
    
    initial = list()
    
    for val in input_list[5:]:
        
        if str(val[column_index]) != "None" and str(val[column_index]) != "":
            
            initial.append(str(val[column_index]))
        
        ##print("initial is: " + str(initial))    
             
    d =  Counter(initial)
        
    res = [k for k, v in d.items() if v > 1]
            
    if len(res) != 0:
    
        message = "Sheet:" + str(sheet) + " - " + str(display_name) + " contains duplicate values! " + str(res) + " Check your data."
          
    if display_popup == True and len(res) != 0: 
                               
        RaisePopup(message)  
             
    return message

#Function that checks if all elements in one list are found in another list
def findList1InList2(wb, sheet, list1, index_list1, list2, index_list2, display_name_1, display_name_2, display_popup):
        
    for val in list1[index_list1][5:]:

        if val in list2[index_list2][5:]:
            
            message = "Sheet: " + str(sheet) + " There are values in " + str(display_name_1) + " that are not part of " + str(display_name_2) + "! Check your data."
            if display_popup == True:
                
                RaisePopup(message)
                logValidationError(wb, message)
                
        else:
            
            message = ""
        
    return message
        
        
#Function that writes validation errors to a log file
def getCashFlowUnitList(wb):
    GenInf = wb["Cash Flow Units"]
    cashFlowUnitList = list()
       
    return cashFlowUnitList
        

#Function that retrieves the Windows Active Directory Full User Name
def get_display_name():
    GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
    NameDisplay = 3

    size = ctypes.pointer(ctypes.c_ulong(0))
    GetUserNameEx(NameDisplay, None, size)

    nameBuffer = ctypes.create_unicode_buffer(size.contents.value)
    GetUserNameEx(NameDisplay, nameBuffer, size)
    
    return nameBuffer.value

#Function that calculates total number of rows with data in a given sheet
def getTotalSheetRows(wb, sheet):
    ExcelSheet = addin_all_sheets.helper.ReadDDTSheet(wb, str(sheet)) 
    lastrow = 0

    for idx, row in enumerate(ExcelSheet[5:]):
        
        if  ((str(row[1]) != "None" and str(row[1]) != "" and str(row[1]) != "\n") and
             (str(row[2]) != "None" and str(row[2]) != "" and str(row[2]) != "\n") and 
             (str(row[3]) != "None" and str(row[3]) != "" and str(row[3]) != "\n")):
            
            lastrow = idx + 1
            #print("last row: " + str(lastrow))
  
    return lastrow
        
        
#Function that calculates the aggregate Cash Flow Amount from the Cash Flows sheet    
def getTotalCashFlowAmount(wb):
    ExcelSheet = addin_all_sheets.helper.ReadDDTSheet(wb, "Cash Flows") 
    totalAmount = 0
    
    for idx, row in enumerate(ExcelSheet[5:]):
        
        
        if  ((str(row[1]) != "None" and str(row[1]) != "" and str(row[1]) != "\n") and
             (str(row[2]) != "None" and str(row[2]) != "" and str(row[2]) != "\n") and 
             (str(row[3]) != "None" and str(row[3]) != "" and str(row[3]) != "\n")):
            
            print("Cash Flow Amount is: : " + str(row[16]))
            totalAmount = totalAmount + float(row[16])
            
    print("Total Amount is: : " + str(totalAmount))
  
    return totalAmount

#Function that calculates the aggregate CSM Run Off Coverage from the CSM Run Off Profiles sheet    
def getTotalCSMCoverage(wb):
    ExcelSheet = addin_all_sheets.helper.ReadDDTSheet(wb, "CSM Run Off Profiles") 
    totalCSMCoverage = 0
    
    for idx, row in enumerate(ExcelSheet[5:]):
        
        if  ((str(row[1]) != "None" and str(row[1]) != "" and str(row[1]) != "\n") and
             (str(row[2]) != "None" and str(row[2]) != "" and str(row[2]) != "\n") and 
             (str(row[3]) != "None" and str(row[3]) != "" and str(row[3]) != "\n")):
            
            print("CSM Coverage is: : " + str(row[3]))
            totalCSMCoverage = totalCSMCoverage + float(row[3])
            
    print("Total CSM Coverage is: : " + str(totalCSMCoverage))
  
    return totalCSMCoverage

#Function that returns current timestamp       
def getTimestamp():
    
    ts = time.gmtime()
    strtimestamp = str(time.strftime("%Y-%m-%d %H:%M:%S", ts))
    currentTimestamp = strtimestamp
    
    return currentTimestamp

#Function that returns current timestamp - formatted for usage in output csv file names      
def getTimestampLogFormatted():
    
    ts = time.gmtime()
    strtimestamp = str(time.strftime("%Y-%m-%d %H:%M:%S", ts))
    strDateTime = strtimestamp
    LogFormatDateTime = ((strDateTime.replace('-', '')).replace(':', '')).replace(' ','')
    
    return LogFormatDateTime


# Function that handles how null values 
# are treated before writing to CSV
def handleNulls(value):
    if value != 'None':
        return str(value) + ']'
    else: 
        return ']'

#Function that checks if column contains negative numbers
def hasNegativeValues(sheet, column, display_name, idx, display_popup):
    
    if  column.isnumeric():
        if int(column) < 0:
        
            line = idx + 1 #because first index is always 0
            errorIndex = line + 5 #because first 5 rows in excel are not actual data 
            message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + str(display_name) + " contains negative values! Check your data."
            
            if display_popup == True: 
                
                RaisePopup(message) 
                #logValidationError(message)
                    
        else:
            message = ""
     
    else:
        line = idx + 1 #because first index is always 0
        errorIndex = line + 5 #because first 5 rows in excel are not actual data 
        message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + str(display_name) + " is not a number! Check your data."
            
        if display_popup == True: 
                
                RaisePopup(message) 
                #logValidationError(message)
        
    return message


#Function that checks if column contains space characters
def hasSpaces(sheet, column, display_name, idx, display_popup):
    
    if  " " in column:
        
        line = idx + 1 #because first index is always 0
        errorIndex = line + 5 #because first 5 rows in excel are not actual data 
        message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + str(display_name) + " contains spaces! Check your data."
        
        if display_popup == True: 
            
            RaisePopup(message) 
            #logValidationError(message)
            
            return message

    else:
        message = ""
        
    return message

#Function that checks if column contains special characters
def hasSpecialChars(sheet, column, display_name, idx, display_popup):
    
    if  " " in column or "/" in column  or "\\" in column or "*" in column  or "%" in column or "&" in column:
        
        line = idx + 1 #because first index is always 0
        errorIndex = line + 5 #because first 5 rows in excel are not actual data 
        message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + str(display_name) + " contains special characters! Check your data."
        
        if display_popup == True: 
            
            RaisePopup(message) 
            #logValidationError(message)
            
            return message

    else:
        message = ""
        
    return message



#Function that writes validation errors to a log file
def logValidationError(wb, log_entry):
    GenInf = wb["General Information"]
    currentTimestamp = getTimestamp()
    outputfileName = "log_alfin_cashflow_validation_errors.txt"
    strOutputDir = str(GenInf['C8'].value)
    path =  strOutputDir + '\\' + outputfileName
    line = str(currentTimestamp) + ": " + str(log_entry)
    
    with open(path, 'r+') as f:
        content = f.read()
        f.seek(0, 0)
        f.write(line.rstrip('\r\n') + '\n' + content)
        
    return 0


#Function that checks if a mandatory column is filled
def MandatoryColumnIsFilled(sheet, column, display_name, idx, display_popup):
    if  column == "" or column == "None":
        line = idx + 1 # because first index is always 0
        errorIndex = line + 5 # because first 5 rows in excel are not actual data 
        message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + str(display_name) + " is mandatory! Check your data."
        
        if display_popup == True: 
            
            RaisePopup(message) 
            #logValidationError(message)
                    
    else:
        message = ""
        
    return message



#Function that prints all the elements of a given list
def printListElements(input_list): 
    line = ""
    
    for val in enumerate(input_list):
        line = line + " " + str(val)
        
    return line

#Function that raises a popup window for the display of errors.
def RaisePopup(message):
    ctypes.windll.user32.MessageBoxW(0, message, "Error!", 0x00001000)


#Function that reads the DDT (by worksheet). Returns the result as a list
def ReadDDTSheet(workbook_name, worksheet_name):
    worksheet_data = workbook_name[worksheet_name]
    
    output_data = list()
    # putting data from excel DDT into a list structure
    # iterating over the rows and
    # getting value from each cell in row
    
    for row in worksheet_data.iter_rows():
        
            row_data = list()
            for cell in row:
                row_data.append(str(cell.value))
            output_data.append(row_data) 
    
    clean_output = removeEmptyRows(output_data)
    
    return clean_output



#Function that removes rows with no data from list
def removeEmptyRows(input_list):
    cleanList = list() 
    for idx, row in enumerate(input_list):   #skip first 5 rows, they don't contain actual data
        
        row_data = ""
             
        if ((str(row[1]) == "None") and
            (str(row[2]) == "None") and
            (str(row[3]) == "None") and
            (str(row[4]) == "None")):
            
            pass
        
        else:
            #print("clean row " + str(idx + 1) + " : " + str(row))
            cleanList.append(row)
            
    return cleanList

#Function that removes output CSV files from location
def removeCSVFiles(wb):
    
    GenInf = wb["General Information"]
    strOutputDir = str(GenInf['C8'].value)
    PATH =  strOutputDir + '\\'
    try:
        for f in glob.glob(PATH + "*alfin_cf_*.csv"):
            os.remove(f)
    except: OSError


#Function for setting number of decimals
def truncate(number, decimals) -> float:
    stepper = pow(10.0, decimals)
    return math.trunc(stepper * number) / stepper

#Function that validates data entered in the General Information sheet 
def validateGeneralInformation(wb, display_popups):
    
    isValidData = True
    
    #Capturing Measurement Date and Movement Step values from the General Information tab
    GenInf = wb["General Information"]
    
    strMeasurementDate = str(GenInf['C4'].value)[:10]
    strReportingWindow = str(GenInf['C5'].value)
    strMovementStep = str(GenInf['C7'].value)
    
    ##print("Measurement Date is " + strMeasurementDate)
        
    if strMovementStep == "" or strMovementStep == "None":
        
        isValidData = False
        message = "Movement Step must be filled (General Information tab)! Check your data."
        
        if display_popups == True:
            RaisePopup(message) 
        
        logValidationError(message)
    else:
        pass
         
        #===================================================================
        # if valueExistsInList(strMovementStep, "Movement_Step_Desc_LoV", idx, display_popups):
        #     pass
        # else:
        #     isValidData = False
        #===================================================================

    if strReportingWindow != "2018Q1" or strReportingWindow == "" or strReportingWindow == "None":
        
        isValidData = False
        message = "Reporting Window must be set to 2018Q1 (General Information tab)! Check your data."
        
        if display_popups == True:
            RaisePopup(message) 
            
        logValidationError(message)
        
    else:
        
        pass
    
    
    if strMeasurementDate != "2018-01-01" or strMeasurementDate == "" or strMeasurementDate == "None":
        
        isValidData = False
        message = "Measurement Date must be set to 2018-01-01 (General Information tab)! Check your data."
        
        if display_popups == True:
            RaisePopup(message) 
            
        logValidationError(message)
        
    else:
        
        pass
    
    return isValidData

#Function that checks if a date is in the required YYYY-MM-DD format
def ValidDateFormat(sheet, column, display_name, idx, display_popup):
    try:
        valid_date = datetime.datetime.strptime(column,'%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d %H:%M:%S')
        message = ""
    except ValueError:
        line = idx + 1 #because first index is always 0
        errorIndex = line + 5 #because first 5 rows in excel are not actual data 
        message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + str(display_name) + " must be in YYYY-MM-DD format! Check your data."
        
        if display_popup == True: 
            
            RaisePopup(message) 
            #logValidationError(message)
            
    return message


#Function that checks column value against a fixed list of values
def valueExistsInList(sheet, value, input_list, idx, display_popup):
    
    inputL = addin_all_sheets.listLoader.LoadLoV(str(input_list))
    
    if str(value) in inputL or str(value) == "" or str(value) == "None": 
        message = "" 
    else:
        line = idx + 1 #because first index is always 0
        errorIndex = line + 5 #because first 5 rows in excel are not actual data 
        message = "Sheet: " + str(sheet) + " - line " + str(errorIndex) + ": " + str(value) + " is not in the list of values: '" + str(input_list) + "'. Check your data."
        
        if display_popup == True: 
            
            RaisePopup(message)  
            #logValidationError(message)
            
    return message


#Function that checks that the workbook has all the necessary sheets
def workbookCheck(workbook):
    intCheckSheet = 0
    blnContinue = False
                                
    if "General Information" in workbook.sheetnames: 
        intCheckSheet = intCheckSheet + 1
    if "Cash Flows" in workbook.sheetnames: 
        intCheckSheet = intCheckSheet + 1
    if "Cash Flow Units" in workbook.sheetnames: 
        intCheckSheet = intCheckSheet + 1
    if "Composition" in workbook.sheetnames: 
        intCheckSheet = intCheckSheet + 1
    if "CSM Run Off Profiles" in workbook.sheetnames: 
        intCheckSheet = intCheckSheet + 1
    if "Insurance Portfolio" in workbook.sheetnames: 
        intCheckSheet = intCheckSheet + 1
    if "Portfolio Group" in workbook.sheetnames: 
        intCheckSheet = intCheckSheet + 1
    
    if intCheckSheet == 7:
        blnContinue = True
    elif intCheckSheet != 7:
        blnContinue = False
    
    return blnContinue;

            





    



