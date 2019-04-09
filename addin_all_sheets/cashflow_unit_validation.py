import addin_all_sheets.helper
import addin_all_sheets.listLoader
import time

def ValidateCashflowUnit(wb, display_popups):
                        
######################################################################################
    
    
    #Capturing Measurement Date and Movement Step values from the General Information tab
    GenInf = wb["General Information"]
    strMeasurementDate = str(GenInf['C4'].value)
    strMovementStep = str(GenInf['C7'].value)
    strsheetName = "Cash Flow Units"
    isValidData = True
    
    #reading data from Cash Flows sheet...
    excel_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Cash Flow Units")               
    
    addin_all_sheets.helper.delimitLogEntries(wb)
    
#####################Begin Validations################################################
    print("Cash Flow Units - Start validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

    message = addin_all_sheets.helper.findDuplicates(wb, strsheetName, excel_data, 2, "Cash Flow Unit ID", display_popups)
     
    if message == "":
        pass
    else:
        addin_all_sheets.helper.logValidationError(wb, message)
        isValidData = False
            
            
#         message = addin_all_sheets.helper.findList1InList2(wb, strsheetName, excel_data, 4, excel_data, 2, "Cash Flow Unit Parent", "Cash Flow Unit", display_popups)
#         if message == "":
#             pass
#         else:
#             addin_all_sheets.helper.logValidationError(wb, message)
#             isValidData = False
 
#######################################################################################
            
    #skipping first 5 rows, as they don't contain actual data
    for idx, val in enumerate(excel_data[5:]):
        
            message = ""
            strCashFlowUnitID = str(val[2])  
            
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
        
#######################################################################################

            strCashFlowUnitApplicationLevel = str(val[3]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCashFlowUnitApplicationLevel, "Cash Flow Unit Application Level", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
        
#######################################################################################
        
#             strCashFlowUnitParent = str(val[4]) 
#         
#             message = addin_all_sheets.valueExistsInList(strsheetName, strCashFlowUnitParent, "Cash_Flow_Unit_LoV", idx, display_popups)
#             if message == "":
#                 pass
#             else:
#                 addin_all_sheets.logValidationError(wb, message)
#                 isValidData = False
        
#######################################################################################

            strCashFlowUnitLabel = str(val[5]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCashFlowUnitLabel, "Cash Flow Unit Label", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False

#######################################################################################
        
            strCashFlowUnitLevel = str(val[6]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCashFlowUnitLevel, "Cash Flow Unit Level", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
        
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strCashFlowUnitLevel, "Cash_Flow_Unit_Level_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 
        
#######################################################################################
        
            strInsurancePortfolioCode = str(val[7]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strInsurancePortfolioCode, "Insurance Portfolio Code", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
        
#             message = addin_all_sheets.valueExistsInList(strsheetName, strInsurancePortfolioCode, "Insurance_Portfolio_Code_LoV", idx, display_popups)
#             if message == "":
#                 pass
#             else:
#                 addin_all_sheets.logValidationError(wb, message)
#                 isValidData = False 
        
#######################################################################################
        
            strMeasurementModelType = str(val[9]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strMeasurementModelType, "Measurement Model Type", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False
        
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strMeasurementModelType, "Measurement_Model_Type_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 
        
#######################################################################################
        
            strCohort = str(val[10]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCohort, "Cohort", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False


            message = addin_all_sheets.helper.containsWantedValue(strsheetName, strCohort, "2018", "Cohort", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False

#######################################################################################
        
            strExpectedProfitability = str(val[11]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strExpectedProfitability, "Expected Profitability", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False


            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strExpectedProfitability, "Expected_Profitability_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 

#######################################################################################
        
            strCurveID = str(val[12]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strCurveID, "Curve ID", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False

#######################################################################################
        
            strTransitionMethod = str(val[14]) 
            
            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strTransitionMethod, "Transition_Method_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 

#######################################################################################
        
            strReinsuranceFlag = str(val[15]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strReinsuranceFlag, "Reinsurance Flag", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False


            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strReinsuranceFlag, "Reinsurance_Flag_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 

#######################################################################################
        
            strRiderMainCoverage = str(val[17]) 

            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strRiderMainCoverage, "Rider_Main_Coverage_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 

#######################################################################################
        
            strChannelID= str(val[18]) 

            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strChannelID, "Channel_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 

#######################################################################################
        
            strInterCompanyFlag = str(val[19]) 
            
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strInterCompanyFlag, "Inter Company Flag", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False


            message = addin_all_sheets.helper.valueExistsInList(strsheetName, strInterCompanyFlag, "Inter_Company_Flag_LoV", idx, display_popups)
            if message == "":
                pass
            else:
                addin_all_sheets.helper.logValidationError(wb, message)
                isValidData = False 

#######################################################################################
        
            strInterCompanyEntity = str(val[20]) 
                        
            #message = ""
            message = addin_all_sheets.helper.MandatoryColumnIsFilled(strsheetName, strInterCompanyEntity, "Inter Company Entity", idx, False)
            
            if message != "" and strInterCompanyFlag =="Y":
                
                addendum = "Inter Company Entity must be filled (only applies if Inter Company Flag is set to Y)"
                addin_all_sheets.helper.logValidationError(wb, addendum)
                addin_all_sheets.helper.RaisePopup(addendum)
                isValidData = False
                
            else:
                pass

#             message = addin_all_sheets.valueExistsInList(strsheetName, strInterCompanyEntity, "Inter_Company_Entity_LoV", idx, display_popups)
#             if message == "":
#                 pass
#             else:
#                 addin_all_sheets.helper.logValidationError(wb, message)
#                 isValidData = False 

#######################################################################################
        
            strStartDate = str(val[21]) 
            
            message = addin_all_sheets.helper.ValidDateFormat(strsheetName, strStartDate, "Start Date", idx, display_popups)
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
                
#######################################################################################
        
            strEndDate = str(val[22]) 
            #print("End Date: " + strEndDate)
            
            message = addin_all_sheets.helper.ValidDateFormat(strsheetName, strEndDate, "End Date", idx, display_popups)
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
                
#             message = addin_all_sheets.checkDate1BiggerOrEqualDate2(strsheetName, strStartDate, strEndDate, "End Date", "Start Date", "%Y-%m-%d %H:%M:%S", idx, display_popups)
#             if message == "":
#                 pass
#             else:
#                 addin_all_sheets.logValidationError(wb, message)
#                 isValidData = False 
                
            print("Cash Flow Units - End validation process..." + str(time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())))

    return isValidData
    
######################End of Validations###############################################
