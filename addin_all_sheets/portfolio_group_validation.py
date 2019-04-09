import addin_all_sheets.helper
import addin_all_sheets.listLoader
import time

def ValidatePortfolioGroup(wb, display_popups):
                        
######################################################################################
        
        
    #Capturing Measurement Date and Movement Step values from the General Information tab
    GenInf = wb["General Information"]
    strsheetName = "Portfolio Group"
                        
    print("Potfolio Group Validation - Started reading DDT Sheet..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

    #reading data from Portfolio Group sheet...
    excel_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Portfolio Group")               
    print("Potfolio Group Validation - Finished reading DDT Sheet..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
    addin_all_sheets.helper.delimitLogEntries(wb)
    isValidData = True
    
#####################Begin Validations################################################
    print("Potfolio Group - Beginning of validations..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))           
    #skipping first 5 rows, as they don't contain actual data
    for idx, val in enumerate(excel_data[5:]):
        
        if ((str(val[1]) != "None") and
        (str(val[2]) != "None") and
        (str(val[3]) != "None") and
        (str(val[4]) != "None")):
        
            message = ""
            strPortfolioGroupId = str(val[1])  
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strPortfolioGroupId, "Portfolio Group Id", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
######################################################################################

            strInsurancePortfolioCode = str(val[2])  
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strInsurancePortfolioCode, "Insurance Portfolio Code", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
               
######################################################################################
            
            strExpectedProfitability = str(val[3])  
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strExpectedProfitability, "Expected Profitability", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
            #message = ""
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strExpectedProfitability, "Expected_Profitability_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
                
######################################################################################
            
            strStatus = str(val[4])  
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strStatus, "Status", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
            #message = ""
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strStatus, "Status_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
                
######################################################################################
            
            strCohort = str(val[5])  
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCohort, "Cohort", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
            #message = ""
            message = addin_all_sheets.helper.containsWantedValue(strsheetName, strCohort, "2018", "Cohort", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False     
                
######################################################################################
            
            strStartDate = str(val[7])  
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strStartDate, "Start Date", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
              
            message = addin_all_sheets.helper.containsWantedValue(strsheetName, strStartDate[:10], "2018-01-01", "Start Date", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 
              
            message = addin_all_sheets.helper.ValidDateFormat(strsheetName, strStartDate, "Start Date", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 
                
######################################################################################
            
            strEndDate = str(val[8])  
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strEndDate, "End Date", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
            message = addin_all_sheets.helper.containsWantedValue(strsheetName, strEndDate[:10], "2018-12-31", "End Date", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 
              
            message = addin_all_sheets.helper.ValidDateFormat(strsheetName, strEndDate, "End Date", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 
                
    print("Potfolio Group - End of validations..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))            
                
    return isValidData
    
######################End of Validations###############################################
