import addin_all_sheets.helper
import addin_all_sheets.listLoader
import time

def ValidateCSMRunoff(wb, display_popups):
                        
######################################################################################
        
        
        #Capturing Measurement Date and Movement Step values from the General Information tab
        GenInf = wb["General Information"]
        strMeasurementDate = str(GenInf['C4'].value)
        strMovementStep = str(GenInf['C7'].value)
        strsheetName = "CSM Run Off Profiles"
        
        #reading data from CSM Run Off Profiles sheet...
        excel_data = addin_all_sheets.helper.ReadDDTSheet(wb, "CSM Run Off Profiles")               
        
        addin_all_sheets.helper.delimitLogEntries(wb)
        isValidData = True
        
#####################Begin Validations################################################
                    
        print("CSM Runoff - Start validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
      
        #skipping first 5 rows, as they don't contain actual data
        for idx, val in enumerate(excel_data[5:]):
            
            if ((str(val[1]) != "None") and
            (str(val[2]) != "None") and
            (str(val[3]) != "None") and
            (str(val[4]) != "None")):
            
                message = ""
                strPortfolioGroupID = str(val[1])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strPortfolioGroupID, "Portfolio Group ID", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                     
    ######################################################################################
    
                strProfileDate = str(val[2])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strProfileDate, "Profile Date", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
    
                #message = ""
                message = addin_all_sheets.helper.ValidDateFormat(strsheetName, strProfileDate, "Profile Date", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                    
    ######################################################################################
    
                strTotalCoverage = str(val[3])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strTotalCoverage, "Total Coverage", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
    
                #message = ""
                message = addin_all_sheets.helper.hasNegativeValues(strsheetName, strTotalCoverage, "Total Coverage", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                    
    ######################################################################################
    
                strCoverageServiced = str(val[4])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCoverageServiced, "Coverage Serviced", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                    
            
            print("CSM Runoff - End validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
        
            return isValidData
        
    ######################End of Validations###############################################
