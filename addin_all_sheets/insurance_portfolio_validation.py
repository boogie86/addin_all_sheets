import addin_all_sheets.helper
import addin_all_sheets.listLoader
import time

def ValidateInsurancePortfolio(wb, display_popups):
                        
######################################################################################
        
        
        #Capturing Measurement Date and Movement Step values from the General Information tab
        GenInf = wb["General Information"]
        strMeasurementDate = str(GenInf['C4'].value)
        strMovementStep = str(GenInf['C7'].value)
        strsheetName = "Insurance Portfolio"
        
        #reading data from Insurance Portfolio sheet...
        excel_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Insurance Portfolio")               
        
        addin_all_sheets.helper.delimitLogEntries(wb)
        isValidData = True
        
#####################Begin Validations################################################
                    
        print("Insurance Portfolio - Start validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
                    
        #skipping first 5 rows, as they don't contain actual data
        for idx, val in enumerate(excel_data[5:]):
            
            if ((str(val[1]) != "None") and
            (str(val[2]) != "None") and
            (str(val[3]) != "None") and
            (str(val[4]) != "None")):
            
                message = ""
                strInsurancePortfolioCode = str(val[1])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strInsurancePortfolioCode, "Insurance Portfolio Code", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                    
    ######################################################################################
    
                strSIILob = str(val[3])  
                
                #message = ""
                message = addin_all_sheets.helper.valueExistsInList(strsheetName, strSIILob, "SII_Lob_LoV", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                   
    ######################################################################################
                
                strMeasurementModelType = str(val[4])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strMeasurementModelType, "Measurement Model Type", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                   
                   
                #message = ""
                message = addin_all_sheets.helper.valueExistsInList(strsheetName, strMeasurementModelType, "Measurement_Model_Type_LoV", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                    
    ######################################################################################
                
                strMeasurementFrequency = str(val[5])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strMeasurementFrequency, "Measurement Frequency", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                   
                   
                #message = ""
                message = addin_all_sheets.helper.valueExistsInList(strsheetName, strMeasurementFrequency, "Measurement_Frequency_LoV", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False          
                    
    ######################################################################################
                
                strPortfolioGroupDuration = str(val[6])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strPortfolioGroupDuration, "Portfolio Group Duration", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
             
    ######################################################################################
                
                strStatus = str(val[7])  
                
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
                
                strIndCollIndicator = str(val[8])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strIndCollIndicator, "Ind Coll Indicator", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                   
                   
                #message = ""
                message = addin_all_sheets.helper.valueExistsInList(strsheetName, strIndCollIndicator, "Ind_Coll_Indicator_LoV", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False 
                    
    ######################################################################################
                
                strBaseCurrency = str(val[9])  
                
                #message = ""
                message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strBaseCurrency, "Base Currency", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False
                   
                   
                #message = ""
                message = addin_all_sheets.helper.valueExistsInList(strsheetName, strBaseCurrency, "Base_Currency_LoV", idx, display_popups)
                if message == "":
                    pass
                else:
                    addin_all_sheets.helper.logValidationError(wb, message)
                    isValidData = False 
                    
            
            print("Insurance Portfolio - End validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
        
            return isValidData
        
    ######################End of Validations###############################################
