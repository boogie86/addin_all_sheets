#Function that defines and returns Cash Flow Header line  
def HeaderCashFlow():  
    header = (
          "ALCF_CALCULATING_BU_CODE]"
        + "ALCF_CALCULATING_DATE_TIME]"
        + "ALCF_MEASUREMENT_DATE]"
        + "ALCF_CASHFLOW_UNIT_ID]"
        + "ALCF_CASHFLOW_DATE]"
        + "ALCF_FINANCIAL_FACT_STATUS]"
        + "ALCF_FINANCIAL_FACT_TYPE]"
        + "ALCF_FINANCIAL_FACT_SUBTYPE]"
        + "ALCF_INVESTMENT_COMPONENT_IND]"
        + "ALCF_SERVICE_PERIOD]"
        + "ALCF_CASHFLOW_PURPOSE]"
        + "ALCF_FINANCIAL_FACT_ID]"
        + "ALCF_CURRENCY]"
        + "ALCF_LRC_LIC_INDICATOR]"
        + "ALCF_RATE_INDICATOR]"
        + "ALCF_PV_AMOUNT_T0]"
        + "ALCF_PV_AMOUNT_T1]"
        + "ALCF_INTEREST_ACCRETION]"
        + "ALCF_CASHFLOW_AMOUNT]"
        + "ALCF_DISCOUNT_FLAG]"
        + "ALCF_AGGREGATED_FLAG]"
        + "ALCF_IR_SM_INDICATOR]"
        + "ALCF_DISCOUNTING_FREQUENCY]"
        + "ALCF_PROCESSING_INDICATOR]" 
        + "ALCF_MOVEMENT_STEP"
        )
    
    return str(header)

#Function that defines and returns Cash Flow Unit Header line  
def HeaderCashFlowUnit():  
    header = (
          "ALCF_CALCULATING_BU_CODE]"
        + "ALCF_CALCULATING_DATE_TIME]"
        + "ALCF_MEASUREMENT_DATE]"
        + "ALCF_CASH_FLOW_UNIT_ID]"
        + "ALCF_CASH_FLOW_UNIT_APPLICATION_LEVEL]"
        + "ALCF_CASH_FLOW_UNIT_ID_PARENT]"
        + "ALCF_CASH_FLOW_UNIT_LABEL]"
        + "ALCF_CASH_FLOW_UNIT_LEVEL]"
        + "ALCF_INSURANCE_PORTFOLIO_CODE]"
        + "ALCF_INSURANCE_PRODUCT_TYPE_CODE]"
        + "ALCF_MEASUREMENT_MODEL_TYPE]"
        + "ALCF_COHORT]"
        + "ALCF_EXPECTED_PROFITABILITY]"
        + "ALCF_CURVE_ID]"
        + "ALCF_LOCKED_IN_RATE]"
        + "ALCF_TRANSITION_METHOD]"
        + "ALCF_REINSURANCE_FLAG]"
        + "ALCF_SOURCE_POLICY_SYSTEM]"
        + "ALCF_RIDER_MAIN_COVERAGE]"
        + "ALCF_CHANNEL_ID]"
        + "ALCF_INTER_COMPANY_FLAG]"
        + "ALCF_INTER_COMPANY_ENTITY]"
        + "ALCF_START_DATE]"
        + "ALCF_END_DATE]"
        )
    
    return str(header)



#Function that defines and returns Composition Header line  
def HeaderComposition():  
    header = (
          "ALCF_CALCULATING_BU_CODE]"
        + "ALCF_CALCULATING_DATE_TIME]"
        + "ALCF_MEASUREMENT_DATE]"
        + "ALCF_CASH_FLOW_UNIT_ID]"
        + "ALCF_MASTER_POLICY_ID]"
        + "ALCF_POLICY_ID]"
        + "ALCF_POLICY_STATUS]"
        + "ALCF_COVERAGE_ID]"
        + "ALCF_SLICE_ID]"
        )
    
    return str(header)


#Function that defines and returns CSM Runoff header line  
def HeaderCSMRunoff():  
    header = (
          "ALCF_CALCULATING_BU_CODE]"
        + "ALCF_CALCULATING_DATE_TIME]"
        + "ALCF_MEASUREMENT_DATE]"
        + "ALCF_PORTFOLIO_GROUP_ID]"
        + "ALCF_PROFILE_DATE]"
        + "ALCF_TOTAL_COVERAGE]"
        + "ALCF_COVERAGE_SERVICED]"
        )
    
    return str(header)


#Function that defines and returns Insurance Portfolio header line  
def HeaderInsurancePortfolio():  
    header = (
          "ALCF_CALCULATING_BU_CODE]"
        + "ALCF_CALCULATING_DATE_TIME]"
        + "ALCF_MEASUREMENT_DATE]"
        + "ALCF_INSURANCE_PORTFOLIO_CODE]"
        + "ALCF_INSURANCE_PORTFOLIO_DESCRIPTION]"
        + "ALCF_SII_LOB]"
        + "ALCF_MEASUREMENT_MODEL_TYPE]"
        + "ALCF_MEASUREMENT_FREQUENCY]"
        + "ALCF_PORTFOLIO_GROUP_DURATION]"
        + "ALCF_STATUS]"
        + "ALCF_IND_COLL_INDICATOR]"
        + "ALCF_BASE_CURRENCY]"
        )
    
    return str(header)


#Function that defines and returns Portfolio Group header line  
def HeaderPortfolioGroup():  
    header = (
          "ALCF_CALCULATING_BU_CODE]"
        + "ALCF_CALCULATING_DATE_TIME]"
        + "ALCF_MEASUREMENT_DATE]"
        + "ALCF_PORTFOLIO_GROUP_ID]"
        + "ALCF_INSURANCE_PORTFOLIO_CODE]"
        + "ALCF_EXPECTED_PROFITABILITY]"
        + "ALCF_STATUS]"
        + "ALCF_COHORT]"
        + "ALCF_LOCKED_IN_RATE]"
        + "ALCF_STATUS]"
        + "ALCF_START_DATE]"
        + "ALCF_END_DATE]"
        )
    
    return str(header)



#Function that defines and returns Control file header line  
def HeaderControlFile():  
    header = (
          "ALCF_CALCULATING_BU_CODE]"
        + "ALCF_CALCULATING_DATE_TIME]"
        + "ALCF_MEASUREMENT_DATE]"
        + "ALCF_NAME_CASHFLOWUNITS]"
        + "ALCF_NR_LINES_CASHFLOWS]"
        + "ALCF_NR_LINES_CASHFLOWUNITS]"
        + "ALCF_NR_LINES_CASHFLOWCOMPOSITION]"
        + "ALCF_NR_LINES_CSMRUNOFF]"
        + "ALCF_NR_LINES_INSURANCEPORTFOLIO]"
        + "ALCF_NR_LINES_PORTFOLIOGROUP]"
        + "ALCF_SUM_TOTAL_CASHFLOWS]"
        + "ALCF_SUM_TOTAL_CSMCOVERAGE]"
        + "ALCF_USER_ID_APPROVAL]"
        + "ALCF_VERSION_NR_ADDIN]"
        + "ALCF_VERSION_NR_TEMPLATE]"
        + "ALCF_FILENAME_TEMPLATE]"
        + "ALCF_DELIVERY_APPLICATION_LEVEL]"
        + "ALCF_MOVEMENT_STEP]"
        )
    
    return str(header)
    
    
