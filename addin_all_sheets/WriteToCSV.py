import addin_all_sheets.helper
import addin_all_sheets.views
import addin_all_sheets.headers
import addin_all_sheets.cashflow_validation
import addin_all_sheets.cashflow_unit_validation
import addin_all_sheets.composition_validation
import addin_all_sheets.csm_runoff_validation
import addin_all_sheets.insurance_portfolio_validation
import addin_all_sheets.portfolio_group_validation
import time
import os
import getpass

global outputfileNameCashFlows
global outputfileNameCashFlowUnits
global outputfileNameComposition
global outputfileNameCSMRunoff
global outputfileNameInsurancePortfolio
global outputfileNamePortfolioGroup
global sumTotalCashFlowAmount
global sumTotalCSMCoverage

outputfileNameCashFlows          = ""
outputfileNameCashFlowUnits      = ""
outputfileNameComposition        = ""
outputfileNameCSMRunoff          = ""
outputfileNameInsurancePortfolio = ""
outputfileNamePortfolioGroup     = ""


#Function that writes the Cash Flow data to CSV file
def WriteCashflowsToCSV(wb):  
  
    GenInf = wb["General Information"]
    #Calculation date processing...
    ts = time.gmtime()
    strtimestamp = str(time.strftime("%Y-%m-%d %H:%M:%S", ts))
    strCalculatingDateTime = strtimestamp
    strFormattedCalculatingDateTime = ((strCalculatingDateTime.replace('-', '')).replace(':', '')).replace(' ','')
     
    #retrieving some values from the General Information sheet...
    strCalculatingBUCode = str(GenInf['C2'].value)
    strMeasurementDate = str(GenInf['C4'].value)
    strMovementStep = str(GenInf['C7'].value)
    strOutputDir = str(GenInf['C8'].value)
    sumTotalCashFlowAmount = 0
    #setting name and path for output file...
    outputfileNameCashFlows = "alfin_cf_cashflow_" + strCalculatingBUCode + "_" + strFormattedCalculatingDateTime + ".csv"
    path =  strOutputDir + '\\' + outputfileNameCashFlows
     
    strfixedPreHeader = addin_all_sheets.helper.handleNulls(strCalculatingBUCode) + (addin_all_sheets.helper.handleNulls(strCalculatingDateTime)) + (addin_all_sheets.helper.handleNulls(strMeasurementDate)[:10]) + ']'
    ##print("Pre-header: " + strfixedPreHeader)
     
    #reading cashflow lines from excel DDT...   
    cashflow_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Cash Flows") 
    ##print(cashflow_data)
     
    #validating cashflow data...
    if addin_all_sheets.cashflow_validation.ValidateCashflows(wb, True):     
        #print("Entered Cash Flow validation function...")
        #Open output CSV File for writing
        resultFile = open(path,'w')
        #print("Opened Cash Flow CSV file for writing...")
        #Write Data to CSV file 
        #start with header...
        resultFile.write(addin_all_sheets.headers.HeaderCashFlow())
        resultFile.write('\n')
        #print("Wrote Cash Flow header...")
        #print(cashflow_data)
         
        #iterating through the data and constructing 
        #the output - cell by cell, row by row
        for idx, row in enumerate(cashflow_data[5:]):   #skip first 5 rows, they don't contain actual data
             
            row_data = ""
             
            if ((str(row[1]) != "None") and
            (str(row[2]) != "None") and
            (str(row[3]) != "None") and
            (str(row[4]) != "None")):
             
                #retrieving value of each field...
                 
                CashFlowUnitId         = addin_all_sheets.helper.handleNulls(row[1])
                CashFlowDate           = addin_all_sheets.helper.handleNulls(row[2])[:10] + ']'
                FinancialFactStatus    = addin_all_sheets.helper.handleNulls(row[3])
                FinancialFactType      = addin_all_sheets.helper.handleNulls(row[4])
                FinancialFactType      = addin_all_sheets.helper.handleNulls(row[5])
                InvestmentComponent    = addin_all_sheets.helper.handleNulls(row[6])
                ServicePeriod          = addin_all_sheets.helper.handleNulls(row[7])
                CashFlowPurpose        = addin_all_sheets.helper.handleNulls(row[8])
                FinancialFactID        = addin_all_sheets.helper.handleNulls(row[9])
                Currency               = addin_all_sheets.helper.handleNulls(row[10])
                LRCLICIndicator        = addin_all_sheets.helper.handleNulls(row[11])
                RateIndicator          = addin_all_sheets.helper.handleNulls(row[12])
                PVAmountT0             = addin_all_sheets.helper.handleNulls(addin_all_sheets.helper.truncate(float(row[13]), 15)).replace('.',',') if str(row[13]) !="None" else "]"
                PVAmountT1             = addin_all_sheets.helper.handleNulls(addin_all_sheets.helper.truncate(float(row[14]), 15)).replace('.',',') if str(row[14]) !="None" else "]"
                InterestAccretion      = addin_all_sheets.helper.handleNulls(addin_all_sheets.helper.truncate(float(row[15]), 15)).replace('.',',') if str(row[15]) !="None" else "]"
                CashFlowAmount         = addin_all_sheets.helper.handleNulls(addin_all_sheets.helper.truncate(float(row[16]), 15)).replace('.',',') if str(row[16]) !="None" else "]"
                DiscountFlag           = addin_all_sheets.helper.handleNulls(row[17])
                AggregatedFlag         = addin_all_sheets.helper.handleNulls(row[18])
                IRSMIndicator          = addin_all_sheets.helper.handleNulls(row[19])
                DiscountingFrequency   = addin_all_sheets.helper.handleNulls(row[20])
                ProcessingIndicator    = addin_all_sheets.helper.handleNulls(row[21])
                MovementStep           = strMovementStep + "]"
                                
                #constructing output row...
                row_data = (strfixedPreHeader 
                + CashFlowUnitId      
                + CashFlowDate        
                + FinancialFactStatus 
                + FinancialFactType   
                + FinancialFactType   
                + InvestmentComponent 
                + ServicePeriod       
                + CashFlowPurpose     
                + FinancialFactID     
                + Currency            
                + LRCLICIndicator     
                + RateIndicator       
                + PVAmountT0          
                + PVAmountT1          
                + InterestAccretion   
                + CashFlowAmount      
                + DiscountFlag        
                + AggregatedFlag      
                + IRSMIndicator       
                + DiscountingFrequency
                + ProcessingIndicator 
                + MovementStep) 
                
                #print("row data: " + row_data)
                
                #writing row to the Cashflows CSV file
                resultFile.write(row_data)
                resultFile.write('\n')
             
        #print("wrote rows to Cash Flow file...")    
        resultFile.close()
             
        #print("Closed Cash Flow file...")
        isOk = True
         
    else:      #if validation goes boom...
        #print("Cash Flow validation returned False...")
        addin_all_sheets.helper.RaisePopup('Cash Flows CSV file not generated! Please try again after correcting the errors!')
        isOk = False
         
    return isOk
            
   
            
#Function that writes the Cash Flow Unit data to CSV file
def WriteCashflowUnitToCSV(wb):  
 
    GenInf = wb["General Information"]
    #Calculation date processing...
    ts = time.gmtime()
    strtimestamp = str(time.strftime("%Y-%m-%d %H:%M:%S", ts))
    strCalculatingDateTime = strtimestamp
    strFormattedCalculatingDateTime = ((strCalculatingDateTime.replace('-', '')).replace(':', '')).replace(' ','')
    
    #retrieving some values from the General Information sheet...
    strCalculatingBUCode = str(GenInf['C2'].value)
    strMeasurementDate = str(GenInf['C4'].value)
    strMovementStep = str(GenInf['C7'].value)
    strOutputDir = str(GenInf['C8'].value)

    #setting name and path for output file...
    outputfileNameCashFlowUnits = "alfin_cf_cashflowunit_" + strCalculatingBUCode + "_" + strFormattedCalculatingDateTime + ".csv"
    path =  strOutputDir + '\\' + outputfileNameCashFlowUnits
    
    strfixedPreHeader = addin_all_sheets.helper.handleNulls(strCalculatingBUCode) + (addin_all_sheets.helper.handleNulls(strCalculatingDateTime)) + (addin_all_sheets.helper.handleNulls(strMeasurementDate)[:10]) + ']'
    ##print("Pre-header: " + strfixedPreHeader)
    
    #reading cashflow lines from excel DDT...   
    cashflow_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Cash Flow Units") 
    
    #validating cashflow data...
    if addin_all_sheets.cashflow_unit_validation.ValidateCashflowUnit(wb, True):     
        #print("Entered Cash Flow Unit validation function...")
        #Open output CSV File for writing
        resultFile = open(path,'w')
        #print("Opened Cash Flow Unit CSV file for writing...")
        #Write Data to CSV file 
        #start with header...
        resultFile.write(addin_all_sheets.headers.HeaderCashFlowUnit())
        resultFile.write('\n')
        #print("Wrote Cash Flow Unit header...")
        
        #iterating through the data and constructing 
        #the output - cell by cell, row by row
        for idx, row in enumerate(cashflow_data[5:]):   #skip first 5 rows, they don't contain actual data
            
            row_data = ""
            
            if ((str(row[1]) != "None") and
            (str(row[2]) != "None") and
            (str(row[3]) != "None") and
            (str(row[4]) != "None")):
            
                #retrieving value of each field...
                
                CashFlowUnitId         = addin_all_sheets.helper.handleNulls(row[2])
                CashFlowUnitAppLevel   = addin_all_sheets.helper.handleNulls(row[3])
                CashFlowUnitIdParent   = addin_all_sheets.helper.handleNulls(row[4])
                CashFlowUnitLabel      = addin_all_sheets.helper.handleNulls(row[5])
                CashFlowUnitLevel      = addin_all_sheets.helper.handleNulls(row[6])
                InsurancePortfolioCode = addin_all_sheets.helper.handleNulls(row[7])
                InsuranceProductTypeCode = addin_all_sheets.helper.handleNulls(row[8])
                MeasurementModelType   = addin_all_sheets.helper.handleNulls(row[9])
                Cohort                 = addin_all_sheets.helper.handleNulls(row[10])
                ExpectedProfitability  = addin_all_sheets.helper.handleNulls(row[11])
                CurveID                = addin_all_sheets.helper.handleNulls(row[12])
                LockedInRate           = addin_all_sheets.helper.handleNulls(row[13])
                TransitionMethod       = addin_all_sheets.helper.handleNulls(row[14])
                ReinsuranceFlag        = addin_all_sheets.helper.handleNulls(row[15])
                SourcePolicySystem     = addin_all_sheets.helper.handleNulls(row[16])
                RiderMainCoverage      = addin_all_sheets.helper.handleNulls(row[17])
                ChannelID              = addin_all_sheets.helper.handleNulls(row[18])
                InterCompanyFlag       = addin_all_sheets.helper.handleNulls(row[19])
                InterCompanyEntity     = addin_all_sheets.helper.handleNulls(row[20])
                StartDate              = addin_all_sheets.helper.handleNulls(row[21])
                EndDate                = addin_all_sheets.helper.handleNulls(row[22])
               
                #constructing output row...
                row_data = (strfixedPreHeader 
                            + CashFlowUnitId
                            + CashFlowUnitAppLevel
                            + CashFlowUnitIdParent
                            + CashFlowUnitLabel
                            + CashFlowUnitLevel
                            + InsurancePortfolioCode
                            + InsuranceProductTypeCode
                            + MeasurementModelType
                            + Cohort
                            + ExpectedProfitability
                            + CurveID
                            + LockedInRate
                            + TransitionMethod
                            + ReinsuranceFlag
                            + SourcePolicySystem
                            + RiderMainCoverage
                            + ChannelID
                            + InterCompanyFlag
                            + InterCompanyEntity
                            + StartDate
                            + EndDate) 
               
                #writing row to the Cashflows Unit CSV file
                resultFile.write(row_data)
                resultFile.write('\n')
                
        #print("wrote rows to Cashflows Unit file...")    
        resultFile.close()
            
        #print("Closed Cashflows Unit file...")
        isOk = True
        
    else:      #if validation goes boom...
        #print("Cash Flow Unit validation returned False...")
        addin_all_sheets.helper.RaisePopup('Cash Flow Unit CSV file not generated! Please try again after correcting the errors!')
        isOk = False
    
    return isOk
    
    
    
#Function that writes the Composition data to CSV file
def WriteCompositionToCSV(wb):  
 
    GenInf = wb["General Information"]
    #Calculation date processing...
    ts = time.gmtime()
    strtimestamp = str(time.strftime("%Y-%m-%d %H:%M:%S", ts))
    strCalculatingDateTime = strtimestamp
    strFormattedCalculatingDateTime = ((strCalculatingDateTime.replace('-', '')).replace(':', '')).replace(' ','')
    
    #retrieving some values from the General Information sheet...
    strCalculatingBUCode = str(GenInf['C2'].value)
    strMeasurementDate = str(GenInf['C4'].value)
    strMovementStep = str(GenInf['C7'].value)
    strOutputDir = str(GenInf['C8'].value)

    #setting name and path for output file...
    outputfileNameComposition = "alfin_cf_composition_" + strCalculatingBUCode + "_" + strFormattedCalculatingDateTime + ".csv"
    path =  strOutputDir + '\\' + outputfileNameComposition
    
    strfixedPreHeader = addin_all_sheets.helper.handleNulls(strCalculatingBUCode) + (addin_all_sheets.helper.handleNulls(strCalculatingDateTime)) + (addin_all_sheets.helper.handleNulls(strMeasurementDate)[:10]) + ']'
    ##print("Pre-header: " + strfixedPreHeader)
    
    #reading Composition lines from excel DDT...   
    cashflow_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Composition") 
    
    #validating composition data...
    if addin_all_sheets.composition_validation.ValidateComposition(wb, True):     
        #print("Entered Composition validation function...")
        #Open output CSV File for writing
        resultFile = open(path,'w')
        #print("Opened Composition CSV file for writing...")
        #Write Data to CSV file 
        #start with header...
        resultFile.write(addin_all_sheets.headers.HeaderComposition())
        resultFile.write('\n')
        #print("Wrote Composition header...")
        
        #iterating through the data and constructing 
        #the output - cell by cell, row by row
        for idx, row in enumerate(cashflow_data[5:]):   #skip first 5 rows, they don't contain actual data
            
            row_data = ""
            
            if ((str(row[1]) != "None") and
            (str(row[2]) != "None") and
            (str(row[3]) != "None") and
            (str(row[4]) != "None")):
            
                #retrieving value of each field...
                CashFlowUnitId         = addin_all_sheets.helper.handleNulls(row[1])
                MasterPolicyId         = addin_all_sheets.helper.handleNulls(row[2])
                PolicyId               = addin_all_sheets.helper.handleNulls(row[3])
                PolicyStatus           = addin_all_sheets.helper.handleNulls(row[4])
                CoverageId             = addin_all_sheets.helper.handleNulls(row[5])
                SliceId                = addin_all_sheets.helper.handleNulls(row[6])
               
                #constructing output row...
                row_data = (strfixedPreHeader 
                            + CashFlowUnitId
                            + MasterPolicyId
                            + PolicyId
                            + PolicyStatus
                            + CoverageId
                            + SliceId
                           ) 
               
                #writing row to the Composition CSV file
                resultFile.write(row_data)
                resultFile.write('\n')
                
        #print("wrote Composition rows to file...")    
        resultFile.close()
            
        #print("Closed Composition file...")
        isOk = True
        
    else:      #if validation goes boom...
        #print("Composition validation returned False...")
        addin_all_sheets.helper.RaisePopup('Composition CSV file not generated! Please try again after correcting the errors!')
        isOk = False
        
    return isOk
    
    
    
#Function that writes the CSM Runoff data to CSV file
def WriteCSMRunoffToCSV(wb):  
 
    GenInf = wb["General Information"]
    #Calculation date processing...
    ts = time.gmtime()
    strtimestamp = str(time.strftime("%Y-%m-%d %H:%M:%S", ts))
    strCalculatingDateTime = strtimestamp
    strFormattedCalculatingDateTime = ((strCalculatingDateTime.replace('-', '')).replace(':', '')).replace(' ','')
    sumTotalCSMCoverage = 0
    
    #retrieving some values from the General Information sheet...
    strCalculatingBUCode = str(GenInf['C2'].value)
    strMeasurementDate = str(GenInf['C4'].value)
    strMovementStep = str(GenInf['C7'].value)
    strOutputDir = str(GenInf['C8'].value)

    #setting name and path for output file...
    outputfileNameCSMRunoff = "alfin_cf_csmrunoff_" + strCalculatingBUCode + "_" + strFormattedCalculatingDateTime + ".csv"
    path =  strOutputDir + '\\' + outputfileNameCSMRunoff
    
    strfixedPreHeader = addin_all_sheets.helper.handleNulls(strCalculatingBUCode) + (addin_all_sheets.helper.handleNulls(strCalculatingDateTime)) + (addin_all_sheets.helper.handleNulls(strMeasurementDate)[:10]) + ']'
    ##print("Pre-header: " + strfixedPreHeader)
    
    #reading CSM Runoff lines from excel DDT...   
    cashflow_data = addin_all_sheets.helper.ReadDDTSheet(wb, "CSM Run Off Profiles") 
    
    #validating CSM Run Off data...
    if addin_all_sheets.csm_runoff_validation.ValidateCSMRunoff(wb, True):     
        #print("Entered CSM Run Off validation function...")
        #Open output CSV File for writing
        resultFile = open(path,'w')
        #print("Opened CSM Run Off CSV file for writing...")
        #Write Data to CSV file 
        #start with header...
        resultFile.write(addin_all_sheets.headers.HeaderCSMRunoff())
        resultFile.write('\n')
        #print("Wrote CSM Run Off header...")
        
        #iterating through the data and constructing 
        #the output - cell by cell, row by row
        for idx, row in enumerate(cashflow_data[5:]):   #skip first 5 rows, they don't contain actual data
            
            row_data = ""
            
            if ((str(row[1]) != "None") and
            (str(row[2]) != "None") and
            (str(row[3]) != "None") and
            (str(row[4]) != "None")):
            
                #retrieving value of each field...
                PortfolioGroupId       = addin_all_sheets.helper.handleNulls(row[1])
                ProfileDate            = addin_all_sheets.helper.handleNulls(row[2])
                TotalCoverage          = addin_all_sheets.helper.handleNulls(row[3])
                CoverageServiced       = addin_all_sheets.helper.handleNulls(row[4])
               
                #constructing output row...
                row_data = (strfixedPreHeader 
                            + PortfolioGroupId
                            + ProfileDate
                            + TotalCoverage
                            + CoverageServiced
                           ) 
               
                #writing row to the CSM Runoff CSV file
                resultFile.write(row_data)
                resultFile.write('\n')
                
        #print("wrote CSM Runoff rows to file...")    
        resultFile.close()
            
        #print("Closed CSM Runoff file...")
        isOk = True
        
    else:      #if validation goes boom...
        #print("CSM Runoff validation returned False...")
        addin_all_sheets.helper.RaisePopup('CSM Runoff CSV file not generated! Please try again after correcting the errors!')
        isOk = False
    
    return isOk
    
    
    
    
#Function that writes the Insurance Portfolio data to CSV file
def WriteInsurancePortfolioToCSV(wb):  
 
    GenInf = wb["General Information"]
    #Calculation date processing...
    ts = time.gmtime()
    strtimestamp = str(time.strftime("%Y-%m-%d %H:%M:%S", ts))
    strCalculatingDateTime = strtimestamp
    strFormattedCalculatingDateTime = ((strCalculatingDateTime.replace('-', '')).replace(':', '')).replace(' ','')
    
    #retrieving some values from the General Information sheet...
    strCalculatingBUCode = str(GenInf['C2'].value)
    strMeasurementDate = str(GenInf['C4'].value)
    strMovementStep = str(GenInf['C7'].value)
    strOutputDir = str(GenInf['C8'].value)

    #setting name and path for output file...
    outputfileNameInsurancePortfolio = "alfin_cf_portfolio_" + strCalculatingBUCode + "_" + strFormattedCalculatingDateTime + ".csv"
    path =  strOutputDir + '\\' + outputfileNameInsurancePortfolio
    
    strfixedPreHeader = addin_all_sheets.helper.handleNulls(strCalculatingBUCode) + (addin_all_sheets.helper.handleNulls(strCalculatingDateTime)) + (addin_all_sheets.helper.handleNulls(strMeasurementDate)[:10]) + ']'
    ##print("Pre-header: " + strfixedPreHeader)
    
    #reading Insurance Portfolio lines from excel DDT...   
    cashflow_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Insurance Portfolio") 
    
    #validating Insurance Portfolio data...
    if addin_all_sheets.insurance_portfolio_validation.ValidateInsurancePortfolio(wb, True):     
        #print("Entered Insurance Portfolio validation function...")
        #Open output CSV File for writing
        resultFile = open(path,'w')
        #print("Opened Insurance Portfolio CSV file for writing...")
        #Write Data to CSV file 
        #start with header...
        resultFile.write(addin_all_sheets.headers.HeaderInsurancePortfolio())
        resultFile.write('\n')
        #print("Wrote Insurance Portfolio header...")
        
        #iterating through the data and constructing 
        #the output - cell by cell, row by row
        for idx, row in enumerate(cashflow_data[5:]):   #skip first 5 rows, they don't contain actual data
            
            row_data = ""
            
            if ((str(row[1]) != "None") and
            (str(row[2]) != "None") and
            (str(row[3]) != "None") and
            (str(row[4]) != "None")):
            
                #retrieving value of each field...
                InsurancePortfolioCode        = addin_all_sheets.helper.handleNulls(row[1])
                InsurancePortfolioDescription = addin_all_sheets.helper.handleNulls(row[2])
                SIILob                        = addin_all_sheets.helper.handleNulls(row[3])
                MeasurementModelType          = addin_all_sheets.helper.handleNulls(row[4])
                MeasurementFrequency          = addin_all_sheets.helper.handleNulls(row[5])
                PortfolioGroupDuration        = addin_all_sheets.helper.handleNulls(row[6])
                Status                        = addin_all_sheets.helper.handleNulls(row[7])
                IndCollIndicator              = addin_all_sheets.helper.handleNulls(row[8])
                BaseCurrency                  = addin_all_sheets.helper.handleNulls(row[9])
               
                #constructing output row...
                row_data = (strfixedPreHeader 
                            + InsurancePortfolioCode
                            + InsurancePortfolioDescription
                            + SIILob
                            + MeasurementModelType
                            + MeasurementFrequency
                            + PortfolioGroupDuration
                            + Status
                            + IndCollIndicator
                            + BaseCurrency
                           ) 
               
                #writing row to the Insurance Portfolio CSV file
                resultFile.write(row_data)
                resultFile.write('\n')
                
        #print("wrote Insurance Portfolio rows to file...")    
        resultFile.close()
            
        #print("Closed Insurance Portfolio file...")
        isOk = True
        
    else:      #if validation goes boom...
        #print("Insurance Portfolio returned False...")
        addin_all_sheets.helper.RaisePopup('Insurance Portfolio CSV file not generated! Please try again after correcting the errors!')
        isOk = False
    
    return isOk
    
    
    
    
#Function that writes the Portfolio Group data to CSV file
def WritePortfolioGroupToCSV(wb):  
    isOk = True
    GenInf = wb["General Information"]
    #Calculation date processing...
    ts = time.gmtime()
    strtimestamp = str(time.strftime("%Y-%m-%d %H:%M:%S", ts))
    strCalculatingDateTime = strtimestamp
    strFormattedCalculatingDateTime = ((strCalculatingDateTime.replace('-', '')).replace(':', '')).replace(' ','')
     
    #retrieving some values from the General Information sheet...
    strCalculatingBUCode = str(GenInf['C2'].value)
    strMeasurementDate = str(GenInf['C4'].value)
    strMovementStep = str(GenInf['C7'].value)
    strOutputDir = str(GenInf['C8'].value)
 
    #setting name and path for output file...
    outputfileNamePortfolioGroup = "alfin_cf_portfoliogroup_" + strCalculatingBUCode + "_" + strFormattedCalculatingDateTime + ".csv"
    path =  strOutputDir + '\\' + outputfileNamePortfolioGroup
     
    strfixedPreHeader = addin_all_sheets.helper.handleNulls(strCalculatingBUCode) + (addin_all_sheets.helper.handleNulls(strCalculatingDateTime)) + (addin_all_sheets.helper.handleNulls(strMeasurementDate)[:10]) + ']'
    ##print("Pre-header: " + strfixedPreHeader)
     
    #reading Portfolio Group lines from excel DDT...   
    cashflow_data = addin_all_sheets.helper.ReadDDTSheet(wb, "Portfolio Group") 
     
    #validating Insurance Portfolio data...
    if addin_all_sheets.portfolio_group_validation.ValidatePortfolioGroup(wb, True):     
        #print("Entered Portfolio Group validation function...")
        #Open output CSV File for writing
        resultFile = open(path,'w')
        #print("Opened Portfolio Group CSV file for writing...")
        #Write Data to CSV file 
        #start with header...
        resultFile.write(addin_all_sheets.headers.HeaderPortfolioGroup())
        resultFile.write('\n')
        #print("Wrote Portfolio Group header...")
         
        #iterating through the data and constructing 
        #the output - cell by cell, row by row
        for idx, row in enumerate(cashflow_data[5:]):   #skip first 5 rows, they don't contain actual data
             
            row_data = ""
             
            if ((str(row[1]) != "None") and
            (str(row[2]) != "None") and
            (str(row[3]) != "None") and
            (str(row[4]) != "None")):
             
                #retrieving value of each field...
                PortfolioGroupID        = addin_all_sheets.helper.handleNulls(row[1])
                InsurancePortfolioCode  = addin_all_sheets.helper.handleNulls(row[2])
                ExpectedProfitability   = addin_all_sheets.helper.handleNulls(row[3])
                Status                  = addin_all_sheets.helper.handleNulls(row[4])
                Cohort                  = addin_all_sheets.helper.handleNulls(row[5])
                LockedInRate            = addin_all_sheets.helper.handleNulls(row[6])
                StartDate               = addin_all_sheets.helper.handleNulls(row[7])
                EndDate                 = addin_all_sheets.helper.handleNulls(row[8])
                
                #constructing output row...
                row_data = (strfixedPreHeader 
                            + PortfolioGroupID
                            + InsurancePortfolioCode
                            + ExpectedProfitability
                            + Status
                            + Cohort
                            + LockedInRate
                            + StartDate
                            + EndDate
                           ) 
                
                #writing row to the Portfolio Group CSV file
                resultFile.write(row_data)
                resultFile.write('\n')
                 
        #print("wrote Portfolio Group rows to file...")    
        resultFile.close()
             
        #print("Closed Portfolio Group file...")
        isOk = True
         
    else:      #if validation goes boom...
        #print("Portfolio Group returned False...")
        addin_all_sheets.helper.RaisePopup('Portfolio Group CSV file not generated! Please try again after correcting the errors!')
        isOk = False
            
    return isOk




#Function that writes the Portfolio Group data to CSV file
def WriteControlFileToCSV(wb):  
    isOk = True
    GenInf = wb["General Information"]
    #Calculation date processing...
    ts = time.gmtime()
    strtimestamp = str(time.strftime("%Y-%m-%d %H:%M:%S", ts))
    strCalculatingDateTime = strtimestamp
    strFormattedCalculatingDateTime = ((strCalculatingDateTime.replace('-', '')).replace(':', '')).replace(' ','')
     
    #retrieving some values from the General Information sheet...
    strCalculatingBUCode = str(GenInf['C2'].value)
    strMeasurementDate = str(GenInf['C4'].value)
    strMovementStep = str(GenInf['C7'].value)
    strOutputDir = str(GenInf['C8'].value)
    strAppLevel = str(GenInf['C9'].value)
 
    #setting name and path for output file...
    outputfileNameControlFile = "alfin_cf_control_" + strCalculatingBUCode + "_" + strFormattedCalculatingDateTime + ".csv"
    path =  strOutputDir + '\\' + outputfileNameControlFile
     
    strfixedPreHeader = addin_all_sheets.helper.handleNulls(strCalculatingBUCode) + (addin_all_sheets.helper.handleNulls(strCalculatingDateTime)) + (addin_all_sheets.helper.handleNulls(strMeasurementDate)[:10]) + ']'
    ##print("Pre-header: " + strfixedPreHeader)
    
    
    #Open output CSV File for writing
    resultFile = open(path,'w')
    #print("Opened Portfolio Group CSV file for writing...")
    #Write Data to CSV file 
    #start with header...
    resultFile.write(addin_all_sheets.headers.HeaderControlFile())
    resultFile.write('\n')
    #print("Wrote Control File header...")
         
    row_data = ""
         
         
    #retrieving value of each field...
    CalculatingBUCcode                   = strCalculatingBUCode
    CalculationDateTime                  = strCalculatingDateTime
    MeasurementDate                      = strMeasurementDate[:10]
    ALCF_CFUFileName                     = outputfileNameCashFlowUnits
    ALCF_CFLTotalLines                   = addin_all_sheets.helper.getTotalSheetRows(wb, "Cash Flows")
    ALCF_CFUTotalLines                   = addin_all_sheets.helper.getTotalSheetRows(wb, "Cash Flow Units") 
    ALCF_CFCTotalLines                   = addin_all_sheets.helper.getTotalSheetRows(wb, "Composition") 
    ALCF_CSMTotalLines                   = addin_all_sheets.helper.getTotalSheetRows(wb, "CSM Run Off Profiles") 
    ALCF_IPFTotalLines                   = addin_all_sheets.helper.getTotalSheetRows(wb, "Insurance Portfolio") 
    ALCF_PFGTotalLines                   = addin_all_sheets.helper.getTotalSheetRows(wb, "Portfolio Group") 
    SumCashFlow                          = str(addin_all_sheets.helper.getTotalCashFlowAmount(wb))
    SumCSMCoverage                       = str(addin_all_sheets.helper.getTotalCSMCoverage(wb))
    UserName                             = str(addin_all_sheets.helper.get_display_name())
    VersionAddin                         = "5.0.9"
    VersionTemplate                      = "5.0.9"
    wbkSourceName                        = addin_all_sheets.views.ExcelWBName()
    ALCF_CTLDeliveryApplicationLevel     = strAppLevel
    ALCF_CTLMovementStep                 = strMovementStep
    
#     print("Total rows Cash Flows sheet: " + str(addin_all_sheets.helper.getTotalSheetRows(wb, "Cash Flows")))
#     print("Total rows Cash Flow Units sheet: " + str(addin_all_sheets.helper.getTotalSheetRows(wb, "Cash Flow Units")))
#     print("Total rows Composition sheet: " + str(addin_all_sheets.helper.getTotalSheetRows(wb, "Composition")))
#     print("Total rows CSM Runoff sheet: " + str(addin_all_sheets.helper.getTotalSheetRows(wb, "CSM Run Off Profiles")))
#     print("Total rows Insurance Portfolio sheet: " + str(addin_all_sheets.helper.getTotalSheetRows(wb, "Insurance Portfolio")))
#     print("Total rows Portfolio Group sheet: " + str(addin_all_sheets.helper.getTotalSheetRows(wb, "Portfolio Group")))
#     print("Sum Total Cash Flow Amount: " + str(sumTotalCashFlowAmount))
#     print("Sum Total CSM Coverage: " + str(sumTotalCSMCoverage))

    #constructing output row...
    row_data = (str(CalculatingBUCcode)               + "]" +
                str(CalculationDateTime)              + "]" +
                str(MeasurementDate)                  + "]" +
                str(ALCF_CFUFileName)                 + "]" +
                str(ALCF_CFLTotalLines)               + "]" +
                str(ALCF_CFUTotalLines)               + "]" +
                str(ALCF_CFCTotalLines)               + "]" +
                str(ALCF_CSMTotalLines)               + "]" +
                str(ALCF_IPFTotalLines)               + "]" +
                str(ALCF_PFGTotalLines)               + "]" +
                str(SumCashFlow)                      + "]" +
                str(SumCSMCoverage)                   + "]" +
                str(UserName)                         + "]" +
                str(VersionAddin)                     + "]" +
                str(VersionTemplate)                  + "]" +
                str(wbkSourceName)                    + "]" +
                str(ALCF_CTLDeliveryApplicationLevel) + "]" +
                str(ALCF_CTLMovementStep)             + "]"
               ) 
    print("Control File Row Data: " + row_data)
    #writing row to the Portfolio Group CSV file
    resultFile.write(row_data)
    resultFile.write('\n')
             
    #print("wrote Portfolio Group rows to file...")    
    resultFile.close()
         
    #print("Closed Portfolio Group file...")
            
    return isOk