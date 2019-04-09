#Function that loads the master list of values against which 
#the values in the DDT columns will be checked

def LoadLoV(input_value):
    
    if str(input_value) == "Cash_Flow_Status_LoV":
        output_list = ["Future","Due","Incurred"]
    
    elif str(input_value) == "Financial_Fact_Type_LoV":
        output_list = [
        "Premium",
        "Acquisition Costs",
        "Expenses",
        "TVOG",
        "Risk Adjustment",
        "Claim - Generic",
        "Claim - Death",
        "Claim - Morbidity",
        "Claim - Annuity",
        "Claim - Maturity",
        "Claim - Surrender",
        "Claim - Profit sharing"]
    
    elif str(input_value) == "Investment_Component_LoV":
        output_list = ["Y","N"]
    
    elif str(input_value) == "Service_Period_LoV":
        output_list = ["Future", "Current", "Past"]
    
    elif str(input_value) == "Cash_Flow_Unit_Level_LoV":
        output_list = ["Portfolio", "Master Policy", "Policy", "Coverage", "Slice"]
    
    elif str(input_value) == "Measurement_Model_Type_LoV":
        output_list = ["GMM", "PAA", "VFA"]
    
    elif str(input_value) == "Expected_Profitability_LoV":
        output_list = ["NSPO", "ONRS", "REMC"]
    
    elif str(input_value) == "Reporting_Unit_LoV":
        output_list = [
        "FAR-I0008",
        "FAR-I0094",
        "FAR-I0002",
        "FAR-I0308",
        "FAR-I0022",
        "X0482",
        "BG101",
        "TR001",
        "SK001",
        "CODA-20",
        "CODA-25",
        "EXACT-010",
        "EXACT-020",
        "EXACT-030",
        "EXACT-040",
        "EXACT-050",
        "ES002",
        "ES001",
        "GR001",
        "HU001",
        "CZ001",
        "PL001",
        "RO001",
        "X0339",
        "X2706",
        "X2995"]
    
    elif str(input_value) == "Policy_Status_LoV":
        output_list = [
        "In Force",
        "Lapsed",
        "Matured",
        "Cancelled",
        "Premium Holiday",
        "Claim Lump Sum",
        "Claim Annuity",
        "Claims WOP",
        "Paid UP",
        "Claims Riders",
        "Reinstatement"]
    
    elif str(input_value) == "Cash_Flow_Unit_Application_Level_LoV":
        output_list = ["Actual", "Future", "Both"]
    
    elif str(input_value) == "Rider_Main_Coverage_LoV":
        output_list = ["Rider", "Main Coverage"]
    
    elif str(input_value) == "Discounting_Frequency_LoV":
        output_list = ["Monthly", "Quarterly", "Annually"]
    
    elif str(input_value) == "Processing_Indicator_LoV":
        output_list = ["Advance", "Arrear", "Mid"]
    
    elif str(input_value) == "Movement_Step_LoV":
        output_list = ["0a",  "0b",  "1a",  "1b", "1", "2", "3", "4", "5a", "5b", "5c", "5", "6a", "6b", "6", "7", "8", "9", "10", "999"]
    
    elif str(input_value) == "Movement_Step_Desc_LoV":
        output_list = [
        "0b. Initial Recognition",
        "1a. Assumption Update - CSM",
        "1b. Assumption Update - P&L",
        "1. Assumption Update - Total",
        "2. Interest Accretion",
        "3. Exp Adj - New service",
        "4. Transfer LRC to LIC",
        "5a. Exp Adj - Direct (lapse)",
        "5b. Exp Adj - Direct (P&C - past)",
        "5c. Exp Adj - Direct (P&C - current)",
        "5. Exp Adj - Direct (Total)",
        "6a. Release Cash Flows (past)",
        "6b. Release Cash Flows (current)",
        "6. Release Cash Flows (Total)",
        "7. Exp Adj - Indirect Actuarial (lapse)",
        "8. Exp Adj - Indirect new service",
        "9. Economic Variance",
        "10. Unwind CSM / EoP",
        "None (Actuals)"]
    
    elif str(input_value) == "Cash_Flow_Purpose_LoV":
        output_list = ["Reporting only", "Forecasting only", "Reporting & Forecasting"]
    
    elif str(input_value) == "IR_SM_Indicator_LoV":
        output_list = ["SM", "IR"]
    
    elif str(input_value) == "Currency_LoV":
        output_list = [
        "EUR",
        "AUD",
        "BGN",
        "CAD",
        "CHF",
        "CNY",
        "CZK",
        "DKK",
        "GBP",
        "HKD",
        "HUF",
        "INR",
        "JPY",
        "KRW",
        "MYR",
        "NOK",
        "NZD",
        "PLN",
        "RON",
        "RUB",
        "SEK",
        "THB",
        "TRY",
        "TWD",
        "USD",
        "ZEU"
        ]
    
    elif str(input_value) == "Reporting_Window_LoV":
        output_list = [
        "2017Q1",
        "2017Q2",
        "2017Q3",
        "2017Q4",
        "2017YE",
        "2018Q1",
        "2018Q2",
        "2018Q3",
        "2018Q4",
        "2018YE"
        ]
    
    elif str(input_value) == "Channel_LoV":
        output_list = [
        "FAR_10",
        "FAR_15",
        "FAR_18",
        "FAR_20",
        "FAR_30",
        "FAR_41",
        "FAR_42",
        "FAR_43",
        "FAR_50",
        "FAR_55",
        "FAR_99",
        "PS_BA",
        "PS_BK",
        "PS_BR",
        "PS_CS",
        "PS_DR",
        "PS_OT",
        "PS_TA",
        "PS_UN",
        "PS_UNDIV"
        ]
    
    elif str(input_value) == "Measurement_Frequency_LoV":
        output_list = ["Annually", "Quarterly", "Monthly"]
    
    elif str(input_value) == "Status_LoV":
        output_list = ["A", "I"]
    
    elif str(input_value) == "Ind_Coll_Indicator_LoV":
        output_list = ["IND", "COLL"]
        
    elif str(input_value) == "Transition_Method_LoV":
        output_list = ["FRA", "MRA", "MVA"]
    
    elif str(input_value) == "Base_Currency_LoV":
        output_list = [
        "EUR",
        "AUD",
        "BGN",
        "CAD",
        "CHF",
        "CNY",
        "CZK",
        "DKK",
        "GBP",
        "HKD",
        "HUF",
        "INR",
        "JPY",
        "KRW",
        "MYR",
        "NOK",
        "NZD",
        "PLN",
        "RON",
        "RUB",
        "SEK",
        "THB",
        "TRY",
        "TWD",
        "USD",
        "ZEU"
        ]
    
    elif str(input_value) == "SII_Lob_LoV":
        output_list = [
        "Liabilities",
        "Medical expense insurance",
        "Income protection insurance",
        "Workers' compensation insurance",
        "Motor vehicle liability insurance",
        "Other motor insurance",
        "Marine, aviation and transport insurance",
        "Fire and other damage to property insurance",
        "General liability insurance",
        "Credit and surety ship insurance",
        "Legal expenses insurance",
        "Assistance insurance",
        "Miscellaneous financial loss insurance",
        "Proportional non-life reinsurance medical expense insurance",
        "Proportional non-life reinsurance income protection insurance",
        "Proportional non-life reinsurance workers' compensation insurance",
        "Proportional non-life reinsurance motor vehicle liability insurance",
        "Proportional non-life reinsurance other motor insurance",
        "Proportional non-life reinsurance marine, aviation and transport insurance",
        "Proportional non-life reinsurance fire and other damage to property insurance",
        "Proportional non-life reinsurance general liability insurance",
        "Proportional non-life reinsurance credit and suretyship insurance",
        "Proportional non-life reinsurance legal expenses insurance",
        "Proportional non-life reinsurance assistance insurance",
        "Proportional non-life reinsurance miscellaneous financial loss insurance",
        "Non-proportional health reinsurance",
        "Non-proportional casualty reinsurance",
        "Non-proportional marine, aviation and transport reinsurance",
        "Non-proportional property reinsurance",
        "Health insurance",
        "Insurance with profit participation",
        "Index-linked and unit-linked insurance - Contracts without options and guarantees",
        "Index-linked and unit-linked insurance - Contracts with options and guarantees",
        "Other life insurance - Contracts without options and guarantees",
        "Other life insurance - Contracts with options and guarantees",
        "Annuities stemming from non-life insurance contracts and relating to health insurance obligations",
        "Annuities stemming from non-life insurance contracts and relating to insurance obligations other than health",   "insurance obligations",
        "Health reinsurance",
        "Life reinsurance",
        "Assets"
        ]
    
    elif str(input_value) == "LRC_LIC_Indicator_LoV":
        output_list = ["LRC", "LIC"]
    
    elif str(input_value) == "Rate_Indicator_LoV":
        output_list = ["Locked-in", "Current"]
    
    elif str(input_value) == "Reinsurance_Flag_LoV":
        output_list = ["D", "R", "W"]
        
    elif str(input_value) == "Inter_Company_Flag_LoV":
        output_list = ["Y", "N"]
        
        
        
    return output_list

