import addin_all_sheets.helper
import addin_all_sheets.listLoader
import time 
def ValidateComposition(wb, display_popups):
                        
######################################################################################
        
        
    #Capturing Measurement Date and Movement Step values from the General Information tab
    GenInf = wb["General Information"]
    strMeasurementDate = str(GenInf['C4'].value)
    strMovementStep = str(GenInf['C7'].value)
    strsheetName = "Composition"
    
    #reading data from Cash Flows sheet...
    excel_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Composition")               
    
    addin_all_sheets.helper.delimitLogEntries(wb)
    isValidData = True
    
#####################Begin Validations################################################
    
    print("Composition - Start validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

    #skipping first 5 rows, as they don't contain actual data
    for idx, val in enumerate(excel_data[5:]):
        
        message = ""
        strCashFlowUnitID = str(val[1])  
        
        #message = ""
        message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCashFlowUnitID, "Cash Flow Unit ID", idx, display_popups)
        if message == "":
            pass
        else:
            addin_all_sheets.helper.logValidationError(wb, message)
            isValidData = False
             
        #message = ""
        message = addin_all_sheets.helper.hasSpaces(strsheetName, strCashFlowUnitID, "Cash Flow Unit ID", idx, display_popups)
        if message == "":
            pass
        else:
            addin_all_sheets.helper.logValidationError(wb, message)
            isValidData = False

    print("Composition - End validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

            
    return isValidData
        
######################End of Validations###############################################
