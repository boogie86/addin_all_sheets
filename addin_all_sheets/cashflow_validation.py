import addin_all_sheets.helper
import time
    
def ValidateCashflows(wb, display_popups):
                        
######################################################################################
        
        isValidData = True
        print("Cash Flow Validation - Start setting initial values..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
        #Capturing Measurement Date and Movement Step values from the General Information tab
        GenInf = wb["General Information"]
        strMeasurementDate = str(GenInf['C4'].value)
        strMovementStep = str(GenInf['C7'].value)
        strsheetName = "Cash Flows"
        print("Cash Flow Validation - delimiting log entries..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
        addin_all_sheets.helper.delimitLogEntries(wb)
        #print("Last row is: " + str(addin_all_sheets.getLastSheetRow(wb, strsheetName)))
        print("Cash Flow Validation - End setting initial values..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
        #reading data from Cash Flows sheet...
        print("Cash Flow Validation - Start reading from Excel DDT Sheet..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
        excel_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Cash Flows")               
        print("Cash Flow Validation - End reading from Excel DDT Sheet..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
        print("Cash Flow Validation - Start duplicates validation..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

        #message = ""
        message = addin_all_sheets.helper.findDuplicates(wb, strsheetName, excel_data, 1, "Cash Flow Unit ID", display_popups)
        
        if message == "":
            pass
        else:
            addin_all_sheets.helper.logValidationError(wb, message)
            isValidData = False
        print("Cash Flow Validation - End duplicates validation..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))
        print("Cash Flows - Start validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

        #skipping first 5 rows, as they don't contain actual data
        for idx, val in enumerate(excel_data[5:]):
                
            
#####################Begin Validations################################################
        
            strCashFlowUnitID = str(val[1])  
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCashFlowUnitID, "Cash Flow Unit ID", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                 
            #message = ""
            message = addin_all_sheets.helper.hasSpecialChars(strsheetName, strCashFlowUnitID, "Cash Flow Unit ID", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
######################################################################################

            strCashFlowDate = str(val[2])
            
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCashFlowDate, "Cash Flow Date", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                    
                                
            message = addin_all_sheets.helper.ValidDateFormat(strsheetName, strCashFlowDate, "Cash Flow Date", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
                
            message = addin_all_sheets.helper.checkDate1BiggerOrEqualDate2(strsheetName, strCashFlowDate, strMeasurementDate, 'Cash Flow Date', 'Measurement Date', '%Y-%m-%d %H:%M:%S', idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False

######################################################################################

            strFinancialFactStatus = str(val[3])

            
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strFinancialFactStatus, "Financial Fact Status", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
            
                        
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strFinancialFactStatus, "Cash_Flow_Status_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False

######################################################################################
                
            strFinancialFactType = str(val[4])
            
            
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strFinancialFactType, "Financial Fact Type", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
#                

            message = addin_all_sheets.helper.containsForrbidenValue(strsheetName, strFinancialFactType, "Guarantee", "Financial Fact Type", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strFinancialFactType, "Financial_Fact_Type_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
######################################################################################

            strInvestmentComponent = str(val[6])

            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strInvestmentComponent, "Investment_Component_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                    
######################################################################################

            strServicePeriod = str(val[7])

            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strServicePeriod, "Service_Period_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                    
######################################################################################
                    
            strCashFlowPurpose = str(val[8])
            
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCashFlowPurpose, "Cash Flow Purpose", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
                
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strCashFlowPurpose, "Cash_Flow_Purpose_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
######################################################################################

            strCurrency = str(val[10])

            
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCurrency, "Currency", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
                
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strCurrency, "Currency_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
######################################################################################

            strLRCLICIndicator = str(val[11])
            strDiscountFlag = str(val[15])


            if strLRCLICIndicator != "" and (strDiscountFlag == "Y" or strDiscountFlag == "" or strDiscountFlag == "None"):
                pass
            else:
                val[11] = ""
                
                
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strLRCLICIndicator, "LRC_LIC_Indicator_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
######################################################################################

            strRateIndicator = str(val[12])
            strDiscountFlag = str(val[15])


            if strRateIndicator != "" and (strDiscountFlag == "Y" or strDiscountFlag == "" or strDiscountFlag == "None"):
                pass
            else:
                val[12] = ""
                
                
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strRateIndicator, "Rate_Indicator_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
######################################################################################

            strPVAmountT0 = str(val[13])
            strDiscountFlag = str(val[15])

            if strPVAmountT0 != "" and (strDiscountFlag == "Y" or strDiscountFlag == "" or strDiscountFlag == "None"):
                pass
            else:
                val[13] = ""
                
######################################################################################

            strPVAmountT1 = str(val[14])
            strDiscountFlag = str(val[15])

            if strPVAmountT1 != "" and (strDiscountFlag == "Y" or strDiscountFlag == "" or strDiscountFlag == "None"):
                pass
            else:
                val[14] = ""
                
######################################################################################

            strIRSMIndicator = str(val[19])
            
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strIRSMIndicator, "IR/SM Indicator", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
                
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strIRSMIndicator, "IR_SM_Indicator_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
                
######################################################################################
                    
            if strMovementStep == "" or strMovementStep == "None":
                isValidData = False
                message = "Movement Step must be filled (General Information tab)! Check your data."
                addin_all_sheets.helper.RaisePopup(message) 
            else:
                pass
                
                
#             message = addin_all_sheets.helper.valueExistsInList(strsheetName, strMovementStep, "Movement_Step_Desc_LoV", idx, display_popups)
#                 pass
#             else:
#                 isValidData = False

            print("Cash Flows - End validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

        return isValidData
    
######################End of Validations###############################################
