Attribute VB_Name = "myArrays"
Dim arry_FilterFlags(1 To 10) As String
'Dim arry_FilterFlags_LightGrey(1 To 5) As String
Dim LOB_Array(1 To 6) As String
Dim arry_LER_Codes(1 To 7) As String

Option Explicit
Public Function Get_FilterFlag_Array_DarkGrey() As Variant
    
' Change Log:
'       3/29/2021: "Audited Financials Not Due Yet" was added to account for the new value added by Eric

' ****************************************************************************
    
    arry_FilterFlags(1) = "Fund Guaranteed Holdco Loan"
    arry_FilterFlags(2) = "LC Only"
    arry_FilterFlags(3) = "Limited Monitoring"
    arry_FilterFlags(4) = "Liquidation - Remaining Balance"
    arry_FilterFlags(5) = "MM HC Construction"
    arry_FilterFlags(6) = "MM No Debt"
    arry_FilterFlags(7) = "MM HC Fill-Up"
    arry_FilterFlags(8) = "MM Capital Call Line"
    arry_FilterFlags(9) = "New Deal in Quarter (Spreads Not Req)"
    arry_FilterFlags(10) = "WCF Broker"
    
    Get_FilterFlag_Array_DarkGrey = arry_FilterFlags

End Function
Public Function Get_FilterFlag_Array_LightGrey() As Variant
    
' Change Log:
'       3/29/2022: Initial Creation
'                  Added 'Annual / Semi Financials Only' and 'Financials Not Received (Late / Extended)'

' ****************************************************************************
    
    arry_FilterFlags(1) = "Other"
    arry_FilterFlags(2) = "No Historical"
    arry_FilterFlags(3) = "Audited Financials Not Due Yet"
    arry_FilterFlags(4) = "Annual / Semi Financials Only"
    arry_FilterFlags(5) = "Financials Not Received (Late / Extended)"
    
    'Annual / Semi Financials Only
    'Audited Financials Not Due Yet
    'Financials Not Received (Late / Extended)
    
    Get_FilterFlag_Array_LightGrey = arry_FilterFlags

End Function
Public Function Get_LOB_Array() As Variant
    
    LOB_Array(1) = "Middle Market Banking"
    LOB_Array(2) = "Sponsor And Specialty Finance"
    LOB_Array(3) = "Commercial Real Estate"
    LOB_Array(4) = "Asset Based Lending"
    LOB_Array(5) = "Public Sector Finance" ' Added 3/21/2023
    LOB_Array(6) = "Commercial Workout"

    Get_LOB_Array = LOB_Array
    
End Function
Public Function Get_LER_Codes() As Variant
    
' Change Log:
'       12/14/2020: Removed LFT P, as per convo w/ Eric R.
'       1/27/2021:  Went back to 8CRAP + WF as being LER, contrary to the Blue Book definition

' ****************************************************************************
    
    'Get all of the LER codes for determining the flag for the LTV field

    arry_LER_Codes(1) = "8 - LER at close based on actual funded debt"
    arry_LER_Codes(2) = "C - LER due to Committed Debt"
    arry_LER_Codes(3) = "R - LER due to Committed Debt, restricted"
    arry_LER_Codes(4) = "A - LER due to Performance"
    arry_LER_Codes(5) = "P - Performance - No Longer Leveraged"
    
    arry_LER_Codes(6) = "F - Indirect Leveraged"
    arry_LER_Codes(7) = "W - ABL Leveraged"

    Get_LER_Codes = arry_LER_Codes

End Function
