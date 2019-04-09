from django.shortcuts import render
import addin_all_sheets.helper
import addin_all_sheets.WriteToCSV
import openpyxl
import time
import zipfile
import sys

global ExcelWorkbookName
ExcelWorkbookName = ""
    
def index(request):
    
    if "GET" == request.method:
        
        return render(request, 'addin_all_sheets/index.html', {})
    
    else:
        
        try:
                        
            #workbook to load based on user selection
            excel_file = request.FILES["excel_file"]
            ExcelWorkbookName = request.FILES['excel_file'].name
            print("ExcelWorkbookName is: " + str(ExcelWorkbookName))
            elapsedTimeMsg = ""
            print("Views - Starting timer..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
            start = time.time()
            
            print("Views - Started loading excel file..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

            #load selected workbook in memory...
            wb = openpyxl.load_workbook(excel_file)
    
            print("Views - Finished loading excel file..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

            print("Views - Start Checking workbook..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))    
            if addin_all_sheets.helper.workbookCheck(wb):  
                print("Views - Finished Checking workbook..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))    
                print("Views - Start writing CSV files..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))    

                if (
                    addin_all_sheets.WriteToCSV.WriteCashflowsToCSV(wb) and
                    addin_all_sheets.WriteToCSV.WriteCashflowUnitToCSV(wb) and 
                    addin_all_sheets.WriteToCSV.WriteCompositionToCSV(wb) and
                    addin_all_sheets.WriteToCSV.WriteCSMRunoffToCSV(wb) and 
                    addin_all_sheets.WriteToCSV.WriteInsurancePortfolioToCSV(wb) and 
                    addin_all_sheets.WriteToCSV.WritePortfolioGroupToCSV(wb) and 
                    addin_all_sheets.WriteToCSV.WriteControlFileToCSV(wb)):
                     
                    print("Views - Data written to CSV files..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
                    end = time.time()
                    print("Views - Ended timer..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
                    elapsedTime = str(end - start)[:5]
                    elapsedTimeMsg = "CSV files generated in " + elapsedTime + " seconds. "
                        
                else: 
                    elapsedTimeMsg = ""
                    addin_all_sheets.helper.removeCSVFiles(wb)
                        
            else:
                    
                addin_all_sheets.helper.RaisePopup('Error!Sheets are missing from Workbook!')      
                
            print("elapsed: " + elapsedTimeMsg)
        
        except (OSError, zipfile.BadZipfile, KeyError):
            
            addin_all_sheets.helper.RaisePopup("Error! Not a valid file!")
            sys.exit("Error! Not a valid file!")
        
    return render(request, 'addin_all_sheets/index.html', {"elapsedTimeMsg":elapsedTimeMsg})


       
        
def ExcelWBName():
    
    return ExcelWorkbookName


