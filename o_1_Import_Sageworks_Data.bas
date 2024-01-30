Attribute VB_Name = "o_1_Import_Sageworks_Data"
' Declare Workbooks
    Dim wbSource As Workbook

' Declare Worksheets
    Dim wsValidation As Worksheet
    Dim wsData As Worksheet
    Dim wsDetailData As Worksheet
    Dim wsLists As Worksheet
    Dim wsFormulas As Worksheet
    
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet

' Declare Integers
    Dim intLastRow As Long
    Dim intLastCol As Long
    
    Dim intLastCol_wsLists As Long
    Dim intLastCol_wsSource As Long

    Dim intCurRowValidation As Long
    
    Dim intHeaderRow_wsSource As Long
    
' Declare Ranges / "Ranges"
    Dim arryHeader_wsData() As Variant
    
    Dim col_LOB As Long
    Dim col_LOBUpdated As Long
    Dim col_Region As Long
    Dim col_Team As Long
    Dim col_Borrower As Long
    
    Dim col_RM As Long
    Dim col_PM As Long
    Dim col_AM As Long
    
    Dim col_CREFlag As Long
    Dim col_SIC As Long
    Dim col_EndMarket As Long
    Dim col_BRG As Long
    Dim col_FRG As Long
    Dim col_CCRP As Long
    Dim col_LFT As Long
    Dim col_Bright_Line_Waiver As Long
    Dim col_Basis_of_Financials As Long
    Dim col_Role As Long
    Dim col_Sponsor As Long
                
    Dim col_LTV_LER As Long
    
    Dim col_FilterFlag As Long
    Dim col_CovCompl As Long

' Declare List "Ranges"

    Dim arryHeader_wsLists() As Variant
    
    Dim col_Region2_Lists As Long
    Dim col_Team2_Lists As Long

    Dim col_LOB3_Lists As Long
    Dim col_SIC3_Lists As Long

' Declare Arrays
    Dim arry_LER_Codes

' Declare Booleans
    Dim bolSingleSheet As Boolean

' Declare Colors
    Dim clrOrange As Long
    Dim clrLightGray As Long
    Dim clrDarkGray As Long
    
Option Explicit
Sub o_01_MAIN_PROCEDURE()

' ****************************************************************************
'
' Author:       James Rinaldi
' Created Date: 11/3/2020
' Last Updated: 10/6/2022
'
' ----------------------------------------------------------------------------
'
' Purpose:  To import the data from the Sageworks data dump, clean it, and manipulate it for the PMs.
'
' Trigger: uf_Sageworks_Regular
'
' Change Log:
'       11/6/2020:  I continue to add additional sections to manipulate the data.
'       11/24/2020: I copied over the Refresh_wsLists macro to the import process
'       1/27/2021:  Added in data validation through out.
'       6/3/2021:   Removed some fields, added eight fields for CRE, renamed some things
'       6/16/2021:  Created the calculation for the Bright Line Waiver.
'       6/24/2021:  Added the wb.Close at the end
'       7/26/2022:  Created the new code to copy a complete version of the Dashboard before removing LOBs
'       10/6/2022:  Moved the o_5_Create_Workout_LOB code to after breaking out the "Complete" vs normal Dashboard
'
' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency
Application.EnableEvents = False

' ----------------------
' Run the update process
' ----------------------
        
    Call o_02_DIM_GLOBAL_VARIABLES
    
    Call o_11_Import_Sageworks_Customer_Data
    Call o_12_Import_Sageworks_Loan_Data
    Call o_13_Clean_Sageworks_Customer_Data
    Call o_14_Manipulate_Sageworks_Customer_Data
    Call o_15_Update_SIC_Codes
    Call o_16_Import_EBITDA_Data
    Call o_17_Calculate_Bright_Line_Waiver

    Call o_21_Flag_Basic_Anomalies
    Call o_22_Flag_Unique_Anomalies
    Call o_23_Unique_Edits
    Call o_24_Apply_Grey_Cell_Fill
    Call o_25_Apply_Formulas

    Call o_31_Validate_Control_Totals

    Call o_41_Create_Complete_Dashboard
    Call o_42_Remove_Wealth_and_Business_Banking
    
    Call o_5_Create_Workout_LOB

    Application.GoTo wsData.Range("A1"), False
    
    wbSource.Close savechanges:=False
    
    Application.EnableEvents = True
    
Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_DIM_GLOBAL_VARIABLES()

' Purpose: To set the global variables used by the Main Sub.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 7/2/2023

' Change Log:
'       11/3/2020:  Intial Creation
'       2/24/2021:  Added code to unhide all columns
'       6/3/2021:   Changed the name of Bright Line Waiver to 'BLW (Y/N)'
'       10/6/2022:  Removed all of the wsLists lookup for Customer / LOB
'       7/2/2023:   Added the 'CRE Flag' field

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------
        
    ' Assign Worksheets
        Set wsData = ThisWorkbook.Sheets("Dashboard Review")
            wsData.Cells.EntireColumn.Hidden = False ' Unhide everything
        Set wsDetailData = ThisWorkbook.Sheets("Detailed Dashboard")
        Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
        Set wsLists = ThisWorkbook.Sheets("LISTS")
        Set wsFormulas = ThisWorkbook.Sheets("FORMULAS")
    
    ' Assign Integers
        intLastCol = wsData.Cells(1, Columns.count).End(xlToLeft).Column
        intLastCol_wsLists = wsLists.Cells(1, Columns.count).End(xlToLeft).Column
    
    ' Assign "Ranges"
        arryHeader_wsData = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))
    
        col_LOB = fx_Create_Headers("LOB", arryHeader_wsData)
        col_LOBUpdated = fx_Create_Headers("Updated LOB", arryHeader_wsData)
        col_Region = fx_Create_Headers("Region", arryHeader_wsData)
        col_Team = fx_Create_Headers("Team", arryHeader_wsData)
        col_Borrower = fx_Create_Headers("Customer", arryHeader_wsData)
        
        col_RM = fx_Create_Headers("RM", arryHeader_wsData)
        col_PM = fx_Create_Headers("PM", arryHeader_wsData)
        col_AM = fx_Create_Headers("Account Manager (PM & RM)", arryHeader_wsData)
            
        col_CREFlag = fx_Create_Headers("CRE Flag", arryHeader_wsData)
        col_SIC = fx_Create_Headers("SIC Code", arryHeader_wsData)
        col_EndMarket = fx_Create_Headers("End Market", arryHeader_wsData)
        col_BRG = fx_Create_Headers("BRG", arryHeader_wsData)
        col_FRG = fx_Create_Headers("FRG", arryHeader_wsData)
        col_CCRP = fx_Create_Headers("CCRP", arryHeader_wsData)
        col_LFT = fx_Create_Headers("LFT Code", arryHeader_wsData)
        col_Bright_Line_Waiver = fx_Create_Headers("BLW (Y/N)", arryHeader_wsData)
        col_Basis_of_Financials = fx_Create_Headers("Basis of Financials", arryHeader_wsData)
        col_Role = fx_Create_Headers("Role", arryHeader_wsData)
        col_Sponsor = fx_Create_Headers("Sponsor", arryHeader_wsData)
        
        col_LTV_LER = fx_Create_Headers("Loan to Value (LER only)", arryHeader_wsData)
        
        col_FilterFlag = fx_Create_Headers("Filter Flag", arryHeader_wsData)
        col_CovCompl = fx_Create_Headers("Covenant Compliance", arryHeader_wsData)
    
    ' Assign List "Ranges"
        arryHeader_wsLists = Application.Transpose(wsLists.Range(wsLists.Cells(1, 1), wsLists.Cells(1, intLastCol_wsLists)))
        
        col_Region2_Lists = fx_Create_Headers("Region2", arryHeader_wsLists)
        col_Team2_Lists = fx_Create_Headers("Market2", arryHeader_wsLists)
    
        col_LOB3_Lists = fx_Create_Headers("LOB3", arryHeader_wsLists)
        col_SIC3_Lists = fx_Create_Headers("SIC_Code3", arryHeader_wsLists)
    
    ' Assign Arrays
        arry_LER_Codes = Get_LER_Codes
    
    ' Assign Colors
        clrOrange = RGB(253, 223, 199)
        clrLightGray = RGB(217, 217, 217)
        clrDarkGray = RGB(64, 64, 64)

End Sub
Sub o_11_Import_Sageworks_Customer_Data()

' Purpose: To import all of the data from the Sageworks data dump at the Customer level.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 11/9/2022

' Change Log:
'       11/3/2020:  Intial Creation
'       11/13/2020: Turned off the warning to update the workbook on import.
'       11/20/2020: Updated the wsSource code to be dynamic if the ws name changes
'       11/20/2020: Updated the intHeaderRow to not throw an error if the Customer field moves
'       12/29/2020: Added the code to wipe out the old validation data
'       2/4/2021:   Added the code to End if no workbook was selected
'       2/12/2021:  Punted the import to the fx_Open_Workbook function
'       2/19/2021:  Added the code to remove the LOBs that are not currently in the Dashboard process
'       2/25/2021:  Added a sort before deleting the inapplicable LOBs which reduced time to run that code by 85%
'       3/9/2021:   Added code to capture if the source workbook is missing sheets
'       6/3/2021:   Change the wsSource name from 'Commercial Dashboard (Current D' to 'Commercial Dashboard 3.0'
'       8/12/2021:  Added code to point to the 'Source Data' folder, if it exists
'       8/12/2021:  Updated the code for intLastRow_Import to not double count 1435 Rail
'       8/12/2021:  Move the code to remove inapplicable LOBs to o_13_Clean_Sageworks_Customer_Data
'       8/20/2021:  Updated 'Dam it Eric' code to allow the import to run
'       11/9/2022:  Added to close the source workbook if an error occurs
'                   Moved to the latest version of 'fx_Copy_in_Data_for_Matching_Fields'

' ****************************************************************************

'On Error GoTo ErrorHandler

Call fx_Update_Default_Directory

' -----------------
' Declare Variables
' -----------------

    ' Assign Workbooks

        Set wbSource = myFunctions.fx_Open_Workbook(strPromptTitle:="Select the current Sageworks (RAW) data dump", bolCloseIfOpen:=True)
        
        If wbSource.Sheets.count = 1 Then
            MsgBox Title:="Dam it Eric", _
                Prompt:="Dam it Eric, there was only one sheet in the Sageworks data dump.  It will still upload, but you'll be missing EBITDA data."
    
            bolSingleSheet = True

        End If
        
    ' Assign Worksheets
        
        If Evaluate("ISREF('" & "Commercial Dashboard 5.0 Q22023" & "'!A1)") = True Then
            Set wsSource = wbSource.Sheets("Commercial Dashboard 5.0 Q22023")
        Else
            Set wsSource = wbSource.Sheets(3)
        End If
            
        Set wsDest = wsData
            
    ' Assign Integers
    
        If Not wsSource.Range("B:B").Find("Customer") Is Nothing Then
            intHeaderRow_wsSource = wsSource.Range("B:B").Find("Customer").Row
        Else
            intHeaderRow_wsSource = Application.InputBox(Title:="Header Row", Prompt:="Type the row that has the header data", Type:=1, Default:=8)
        End If
    
        intLastRow = WorksheetFunction.Max(wsSource.Cells(Rows.count, "B").End(xlUp).Row, wsSource.Cells(Rows.count, "D").End(xlUp).Row)
    
    ' Dim Strings (for function)
            
    Dim strControlTotalFieldName As String
        strControlTotalFieldName = "Webster Commitment (000's) - Gross Exposure"

    Dim strModuleName As String
        strModuleName = "o_11_Import_Sageworks_Customer_Data"

' --------------------------------
' Wipe out the old validation data
' --------------------------------
    
    wsValidation.Range("A2:E5").ClearContents

' ---------------------------------
' Remove the Header and Footer Rows
' ---------------------------------

    If intHeaderRow_wsSource > 1 Then wsSource.Rows("1:" & intHeaderRow_wsSource - 1).Delete
        intLastCol_wsSource = wsSource.Cells(1, Columns.count).End(xlToLeft).Column
    
' --------------------------------------------------------------------------------
' Sort the data based on the Footer field to just pull in the Customer Footer data
' --------------------------------------------------------------------------------
    
    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(intLastRow, intLastCol_wsSource)).Sort _
        Key1:=wsSource.Range("A1"), Order1:=xlAscending, Header:=xlYes
        
    Dim intLastRow_Import As Long
        intLastRow_Import = wsSource.Range("A:A").Find("Loan Footer").Row - 1
        
' ---------------
' Update the data
' ---------------

    Call fx_Copy_in_Data_for_Matching_Fields( _
        wsSource:=wsSource, _
        wsDest:=wsData, _
        strModuleName:=strModuleName, _
        strControlTotalField:="Webster Outstanding (000's) - Book Balance", _
        intLastRowtoImport:=intLastRow_Import, _
        bolCloseSourceWb:=False, _
        intCurRow_wsValidation:=2)

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_11_Import_Sageworks_Customer_Data", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
wbSource.Close savechanges:=False

End Sub
Sub o_12_Import_Sageworks_Loan_Data()

' Purpose: To import all of the data from the Sageworks data dump at the Loan level.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 11/9/2022

' Change Log:
'       2/12/2021:  Intial Creation, based on o_11_Import_Sageworks_Customer_Data
'       2/16/2021:  Added in the code to clean the data and reset the formatting
'       11/9/2022:  Moved to the latest version of 'fx_Copy_in_Data_for_Matching_Fields'

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    ' Assign Worksheets

    Set wsDest = wsDetailData
    
    ' Dim Strings (for function)
            
    Dim strControlTotalFieldName As String
        strControlTotalFieldName = "Webster Commitment (000's) - Gross Exposure"

    Dim strModuleName As String
        strModuleName = "o_12_Import_Sageworks_Loan_Data"

' ----------------------------------------------------------------------------
' Sort the data based on the Footer field to just pull in the Loan Footer data
' ----------------------------------------------------------------------------
    
    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(intLastRow, intLastCol_wsSource)).Sort _
        Key1:=wsSource.Range("A2"), Order1:=xlDescending, Header:=xlYes
        
    Dim intLastRow_Import As Long
        intLastRow_Import = wsSource.Range("A:A").Find("Customer Footer").Row
        
' ---------------
' Update the data
' ---------------

    Call fx_Copy_in_Data_for_Matching_Fields( _
        wsSource:=wsSource, _
        wsDest:=wsDest, _
        strModuleName:="o_11_Import_Sageworks_Customer_Data", _
        strControlTotalField:=strControlTotalFieldName, _
        intLastRowtoImport:=intLastRow_Import, _
        bolCloseSourceWb:=False, _
        intCurRow_wsValidation:=4)

ThisWorkbook.Activate

' -----------------------
' Clean the Detailed Data
' -----------------------

    ' Dim Formatting Ranges
    
    With wsDetailData
        Dim rngFormat As Range
            Set rngFormat = .Range(.Cells(2, 1), .Cells(2, intLastCol_wsSource))
        
        Dim rngTarget As Range
            Set rngTarget = .Range(.Cells(3, 1), .Cells(intLastRow_Import, intLastCol_wsSource))
    End With

    ' Apply formatting

    rngFormat.Interior.Color = xlNone
    
    rngFormat.Copy: rngTarget.PasteSpecial xlPasteFormats
        Application.CutCopyMode = False

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_12_Import_Sageworks_Loan_Data", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_13_Clean_Sageworks_Customer_Data()

' Purpose: To clean the data that was imported from the Sageworks data dump.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 12/21/2022

' Change Log:
'       11/3/2020:  Intial Creation
'       11/9/2020:  Added code to select no fill on the rngFormat.
'       11/13/2020: Added the code for the paid off accounts
'       3/8/2021:   Added the code to paste the validations from the 2nd row
'       3/29/2022:  Updated to move Lepage to BB on ACBS
'                   Added code to remove the WBS Internal Accounts
'       4/25/2022:  Updated to remove the EF loans based on Region, not LOB due to the move to CRE
'       12/21/2022: Added the Error Handling for an issue with EF Region not existing

' ****************************************************************************

'On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    'Dim intLastRow As Long
        intLastRow = wsData.Cells(Rows.count, "A").End(xlUp).Row
           
    ' Dim "Ranges"
    
    Dim col_PaidOff As Long
        col_PaidOff = fx_Create_Headers("Paid Off", arryHeader_wsData)
    
    Dim col_Outstanding As Long
        col_Outstanding = fx_Create_Headers("Webster Outstanding (000's) - Book Balance", arryHeader_wsData)

    ' Dim Formatting Ranges
    
    With wsData
    
        Dim rngFormat As Range
            Set rngFormat = .Range(.Cells(2, 1), .Cells(2, intLastCol))
        
        Dim rngTarget As Range
            Set rngTarget = .Range(.Cells(3, 1), .Cells(intLastRow, intLastCol))
            
    End With
        
    ' Dim Loop Variables
        
    Dim i As Long
    
' ------------------------------------------------
' Copy the formatting from the 2nd row to the rest
' ------------------------------------------------
    
    rngFormat.Interior.Color = xlNone
    
    rngFormat.Copy: rngTarget.PasteSpecial xlPasteFormats
    rngTarget.PasteSpecial xlPasteValidation
        Application.CutCopyMode = False

' -----------------------------------------
' Remove the data for the inapplicable LOBs
' -----------------------------------------
    
With wsData
    
On Error Resume Next
    
    .Range(.Cells(1, 1), .Cells(intLastRow, intLastCol)).Sort _
        Key1:=.Range("B1"), Order1:=xlAscending, Header:=xlYes

    ' Remove EF Loans based on LOB and Region
    .Range("A1").AutoFilter Field:=col_LOB, Criteria1:=Array("Equipment Finance", ""), Operator:=xlFilterValues
        .Range("A2:A" & intLastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .Range("A1").AutoFilter Field:=col_LOB
        
    .Range("A1").AutoFilter Field:=col_Region, Criteria1:="*Equip*"
        .Range("A2:A" & intLastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .Range("A1").AutoFilter Field:=col_Region
        
    ' Remove BB / SB Loans where PM = "Craig Harper"
    .Range("A1").AutoFilter Field:=col_LOB, Criteria1:=Array("Small Business", "Business Banking"), Operator:=xlFilterValues
    .Range("A1").AutoFilter Field:=col_PM, Criteria1:="Craig Harper"
        .Range("A2:A" & intLastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .Range("A1").AutoFilter Field:=col_LOB
            .Range("A1").AutoFilter Field:=col_PM
                    
    ' Remove BB / SB Loans where Region = "Workout'
    .Range("A1").AutoFilter Field:=col_LOB, Criteria1:=Array("Small Business", "Business Banking"), Operator:=xlFilterValues
    .Range("A1").AutoFilter Field:=col_Region, Criteria1:="*Workout*"
        .Range("A2:A" & intLastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        .Range("A1").AutoFilter Field:=col_LOB
        .Range("A1").AutoFilter Field:=col_Region
        
On Error GoTo 0
        
End With
            
' -------------------------------------------
' Update the BB on ALS loans based on PM name
' -------------------------------------------
    
With wsData
    
    For i = 2 To intLastRow
        If .Cells(i, col_LOB) = "Middle Market Banking" Then
            If .Cells(i, col_PM) = "Katherine Alibozak" Or .Cells(i, col_PM) = "Kim Schierholz" Then
                .Cells(i, col_LOB) = "Business Banking"
                .Cells(i, col_Team) = "BB on ACBS"
            End If
        ElseIf .Cells(i, col_Borrower) = "LePage Homes, Inc. (Group)" Then
            .Cells(i, col_LOB) = "Business Banking"
            .Cells(i, col_Team) = "BB on ACBS"
        End If
    Next i

End With

' ---------------
' Remove Accounts
' ---------------

With wsData

    For i = intLastRow To 2 Step -1
        If .Cells(i, col_PaidOff) = "Yes" And .Cells(i, col_Outstanding) = 0 Then .Rows(i).Delete
        If .Cells(i, col_FilterFlag) = "Liquidation - Zero Balance" And .Cells(i, col_Outstanding) = 0 Then .Rows(i).Delete
        If .Cells(i, col_Borrower) = "W J Connell (Group)" Then .Rows(i).Delete
        If .Cells(i, col_Borrower) = "Webster Capital Finance" Then .Rows(i).Delete
        If .Cells(i, col_Borrower) = "Webster Bank-Sponsor & Specialty" Then .Rows(i).Delete
    Next i

End With

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_13_Clean_Sageworks_Customer_Data", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_14_Manipulate_Sageworks_Customer_Data()

' Purpose: To manipulate the data that was imported from the Sageworks data dump.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 7/2/2023

' Change Log:
'       11/6/2020:  Intial Creation
'       12/7/2020:  Added in the Name Reverse code to fix the RM name
'       12/14/2020: Added in the code to determine the AM Name (PM / RM)
'       12/23/2020: Added code to resolve the issue with the 9900 SIC Codes for PPP Accounts
'       2/1/2021:   Added code to update the Org Type for Col F from '501 (c) (3)' -> 'NFP'
'       3/9/2021:   Added code to Auto-Fill the Covenant Compliance Attestation w/ "N/A - Limited Monitoring" for Limited Monitoring accounts
'       6/24/2021:  If SIC Code = "Unknown" AND Org Type = BLANK change SIC Code to "n/a", as per Eric
'       6/24/2021:  If 'Basel III / HVCRE Compliant (CRE)' is BLANK AND LOB = CRE THEN = Yes, as per Eric
'       6/24/2021:  If Collateral Property State = "" AND COllateral Description <> "" AND LOB = CRE then Collateral Property State = "N/A", as per Eric
'       8/12/2021:  Added code to manipulate the 'BB on ACBS' customers and 'Wealth' customers
'       3/16/2022:  Updated to have a blank org type be N/A
'                   Updated to remove the LER field
'       3/23/2022:  Added code to bypass the "BB on ACBS" if the market is already there
'       3/29/2022   Added the bypass for MM Healthcare CRE deals to not apply a CRE rule
'       12/20/2022: Added code to flag the L-SNB borrowers to apply the requested formatting changes
'       3/23/2023:  Updated to remove the 'MM' from 'MM - Healthcare'
'       6/29/2023:  Added the code to create the 'Team (Updated)' field
'                   Added the code to create the 'Segment (Updated)' field
'                   Added the code for the CRE LOB and related manipulation
'       7/2/2023:   Moved the Segment / Team / CRE code to the top of the sub
'                   Updated to use the 'CRE Flag' instead of LOB = 'Commercial Real Estate'
'       7/10/2023:  Updated the import to include the message box for missing data for Team and Segment lookups

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    'Dim Integers
    
    Dim intLastRow_Lists As Long
        intLastRow_Lists = wsLists.Cells(Rows.count, col_Team2_Lists).End(xlUp).Row
    
    ' Reset intLastRow after deleting some records
        intLastRow = WorksheetFunction.Max(wsData.Cells(Rows.count, "A").End(xlUp).Row, wsData.Cells(Rows.count, "B").End(xlUp).Row)
    
    'Dim "Ranges"
        
    Dim col_Region
        col_Region = fx_Create_Headers("Region", arryHeader_wsData)
    
    Dim col_RegionUpdated
        col_RegionUpdated = fx_Create_Headers("Region (Updated)", arryHeader_wsData)
    
    Dim col_OrgType
        col_OrgType = fx_Create_Headers("Org. Type", arryHeader_wsData)
        
    Dim col_Basel
        col_Basel = fx_Create_Headers("Basel III / HVCRE Compliant (CRE)", arryHeader_wsData)
        
    Dim col_CollatPropState As Long
        col_CollatPropState = fx_Create_Headers("Collateral Property State", arryHeader_wsData)

    Dim col_CollatDesc As Long
        col_CollatDesc = fx_Create_Headers("Collateral Description", arryHeader_wsData)
        
    Dim col_Legacy_Bank As Long
        col_Legacy_Bank = fx_Create_Headers("Legacy Bank", arryHeader_wsData)
        
    'Dim Loop Variables
    
    Dim i As Long
    
    Dim strRegionName As String
    
    'Dim Dictionaries / Arrays
    
    Dim dict_Market As Scripting.Dictionary
        Set dict_Market = New Scripting.Dictionary
        
    Dim arry_RM_Name
        arry_RM_Name = WorksheetFunction.Transpose(wsData.Range(wsData.Cells(1, col_RM), wsData.Cells(intLastRow, col_RM)))
        
    Dim arry_PM_Name
        arry_PM_Name = WorksheetFunction.Transpose(wsData.Range(wsData.Cells(1, col_PM), wsData.Cells(intLastRow, col_PM)))
        
    Dim arry_AM_Name
        arry_AM_Name = WorksheetFunction.Transpose(wsData.Range(wsData.Cells(1, col_AM), wsData.Cells(intLastRow, col_AM)))
    
    Dim arry_SNB_LOBs
        arry_SNB_LOBs = Array("C&I", "Closed", "Consumer Banking", "CRE", "Other", "Public Finance", "REIT", "Specialty Finance", "Support", "Verticals")
    
    
' -------------------------------
' Create the Updated Region field
' -------------------------------

    For i = 2 To intLastRow
        strRegionName = wsData.Cells(i, col_Region).Value2
          
        wsData.Cells(i, col_RegionUpdated) = Mid(strRegionName, InStr(1, strRegionName, " ") + 1, Len(strRegionName))
            
    Next i

' -----------------------------
' Pull in the updated Team data
' -----------------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=wsLists, wsDest:=wsData, _
        str_Source_TargetField:="4. Team", _
        str_Source_MatchField:="4. Business Unit Name", _
        str_Dest_TargetField:="Team (Updated)", _
        str_Dest_MatchField:="Region (Updated)", _
        bol_MissingLookupData_MsgBox:=True, _
        strWsNameLookup:="wsLists - 4. Business Unit Name / 4. Team")
    
' --------------------------------
' Pull in the updated Segment data
' --------------------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=wsLists, wsDest:=wsData, _
        str_Source_TargetField:="4. Segment", _
        str_Source_MatchField:="4. Business Unit Name", _
        str_Dest_TargetField:="Updated LOB", _
        str_Dest_MatchField:="Region (Updated)", _
        bol_MissingLookupData_MsgBox:=True, _
        strWsNameLookup:="wsLists - 4. Business Unit Name / 4. Segment")

' -------------------
' Create the CRE Flag
' -------------------

    Call fx_Update_Single_Field( _
        wsSource:=wsDetailData, wsDest:=wsData, _
        str_Source_TargetField:="Loan Type Code", _
        str_Source_MatchField:="Customer", _
        str_Dest_TargetField:="CRE Flag", _
        str_Dest_MatchField:="Customer", _
        arry_OnlyUseMultipleValues:=Array("Co-Op", "CRE Loans", "Cre Noo", "Cre Oo", "Multifamily"), _
        bol_MultipleOnlyUseValues:=True)

    Call fx_Update_Single_Field( _
        wsSource:=wsData, wsDest:=wsData, _
        str_Source_TargetField:="Updated LOB", _
        str_Source_MatchField:="Customer", _
        str_Dest_TargetField:="CRE Flag", _
        str_Dest_MatchField:="Customer", _
        str_OnlyUseValue:=("Commercial Real Estate"))

' -----------------------
' Manipulate the CRE Flag
' -----------------------

    ' Remove the single '-' from being included as a CRE field
    
    Call fx_Delete_Unused_Data( _
        ws:=wsData, _
        str_Target_Field:="CRE Flag", _
        str_Value_To_Delete:="-", _
        bol_DeleteDataOnly:=True)
        
    ' Update to convert the remaining values to 'Yes'
    
        Call fx_Convert_Values_In_Range( _
        wsData:=wsData, _
        strTargetField:="CRE Flag", _
        strNewValue:="Yes")
    
' -------------------------------------------------------------------------------------
' Loop through and manipulate the various fields before applying the highlighting rules
' -------------------------------------------------------------------------------------
    
    With wsData
        For i = 2 To intLastRow
        
            ' For any customer with a 'Unknown' in the Sponsor field replace it with "None"
            If .Cells(i, col_Sponsor).Value2 = "Unknown" Then .Cells(i, col_Sponsor).Value2 = "None"
            
            ' For any customers with a Filter Flag of 'LC Only' change the LFT Code to be 'LC Only'
            If .Cells(i, col_FilterFlag).Value2 = "LC Only" Then
                .Cells(i, col_LFT).Value2 = "LC Only"
            End If
        
            ' Convert a CCRP to be a 6W for those BRGs w/ a watch flag and FRG >=4
            If .Cells(i, col_BRG) = "6W" And .Cells(i, col_FRG) >= 4 And .Cells(i, col_CCRP) = 6 Then
                .Cells(i, col_CCRP) = "6W"
            End If
        
            ' Convert the RM name to be [FIRST] [LAST] from [LAST], [FIRST]
            arry_RM_Name(i) = fx_Reverse_Given_Name(CStr(arry_RM_Name(i)))
            
            ' Create the AM Name (PM / RM)
            arry_AM_Name(i) = arry_PM_Name(i) & " / " & arry_RM_Name(i)
        
            ' Update the NFP for NFP > update the Org Type for Col F from '501 (c) (3)' -> 'NFP'
            If .Cells(i, col_OrgType) = "501 (c) (3)" Then
                .Cells(i, col_OrgType) = "NFP"
            End If
            
            ' Auto-Fill the Covenant Compliance Attestation w/ "N/A - Limited Monitoring" for Limited Monitoring accounts
            If .Cells(i, col_FilterFlag) = "Limited Monitoring" Then
                .Cells(i, col_CovCompl) = "N/A - Limited Monitoring"
            End If
            
            ' If SIC Code = "Unknown" AND Org Type = BLANK change SIC Code and Org Type to "N/A"
            If .Cells(i, col_SIC) = "Unknown" And .Cells(i, col_OrgType) = "" Then
                .Cells(i, col_SIC) = "N/A"
                .Cells(i, col_OrgType) = "N/A"
            End If
            
            ' If 'Basel III / HVCRE Compliant (CRE)' is BLANK AND LOB = CRE THEN = Yes
            If .Cells(i, col_CREFlag) = "Yes" And .Cells(i, col_Basel) = "" Then
                If Left(.Cells(i, col_Team), 15) <> "Healthcare" Then
                    .Cells(i, col_Basel) = "Yes"
                End If
            End If

            ' If Collateral Property State = "" AND Collateral Description <> "" AND LOB = CRE then Collateral Property State = "N/A"
            If .Cells(i, col_CREFlag) = "Yes" And _
            .Cells(i, col_CollatPropState) = "" And _
            .Cells(i, col_CollatDesc) <> "" And _
            Left(.Cells(i, col_Team), 15) <> "Healthcare" Then
                .Cells(i, col_CollatPropState).Value2 = "N/A"
            End If
            
        Next i
                
    End With

    ' Output the values from the RM and AM arrays into their columns
    wsData.Range(wsData.Cells(1, col_RM), wsData.Cells(intLastRow, col_RM)) = WorksheetFunction.Transpose(arry_RM_Name)
    wsData.Range(wsData.Cells(1, col_AM), wsData.Cells(intLastRow, col_AM)) = WorksheetFunction.Transpose(arry_AM_Name)

' --------------------------------------------------
' Fill the Dictionary with the Region / Market names
' --------------------------------------------------
        
    With wsLists
        For i = 2 To intLastRow_Lists
            dict_Market.Add Key:=.Cells(i, col_Region2_Lists).Value2, Item:=.Cells(i, col_Team2_Lists).Value2
        Next i
    End With
    
' ----------------------
' Create the Market name
' ----------------------

    With wsData
        For i = 2 To intLastRow
            If dict_Market.Exists(.Cells(i, col_Region).Value2) Then
                If .Cells(i, col_Team) <> "BB on ACBS" Then ' Added 3/23/2022
                    .Cells(i, col_Team) = dict_Market.Item(.Cells(i, col_Region).Value2)
                End If
            End If
        Next i
    End With

' -----------------------------
' 'Move' BB and Wealth Accounts
' -----------------------------

    With wsData
        For i = 2 To intLastRow
            If .Cells(i, col_Team) = "BB on ACBS" Then
                .Cells(i, col_LOB) = "Business Banking"
            ElseIf .Cells(i, col_LOB) = "Webster Private Bank" Then
                .Cells(i, col_LOB) = "Wealth"
            End If
        Next i
    End With

' ------------------------
' Flag the L-SNB Borrowers
' ------------------------

    With wsData
        For i = 2 To intLastRow
            If Not IsNumeric(Application.Match(.Cells(i, col_LOB), arry_SNB_LOBs, 0)) Then
                .Cells(i, col_Legacy_Bank) = "WBS"
            Else
                .Cells(i, col_Legacy_Bank) = "SNB"
            End If
        Next i
    End With
    

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_14_Manipulate_Sageworks_Customer_Data", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_15_Update_SIC_Codes()

' Purpose: To resolve the issue with the 9900 SIC Codes for customers with PPP Accounts.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 12/23/2020

' Change Log:
'       12/23/2020: Intial Creation

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    ' Dim Integers

    Dim intLastRow_SIC3 As Long
        intLastRow_SIC3 = wsLists.Cells(Rows.count, col_SIC3_Lists).End(xlUp).Row

    ' Dim Arrays
    
    Dim arryCurData() As Variant
        arryCurData = wsData.Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, col_SIC))
        
    Dim arryTargetData() As Variant
        arryTargetData = wsLists.Range(wsLists.Cells(1, col_LOB3_Lists), wsLists.Cells(intLastRow_SIC3, col_SIC3_Lists))
    
' -------------
' Run your code
' -------------

    wsData.Range(wsData.Cells(1, col_SIC), wsData.Cells(intLastRow, col_SIC)) = _
    Application.Transpose(fx_Update_Data_SIC( _
        strMatchField_Cur:="Customer", _
        strTargetField_Cur:="SIC Code", _
        strMatchField_Lookup:="Customer3", _
        strTargetField_Lookup:="SIC_Code3", _
        arryCurData:=arryCurData, _
        arryLookupData:=arryTargetData))

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_15_Update_SIC_Codes", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_16_Import_EBITDA_Data()

' Purpose: To import the EBITDA fields from the Sageworks Data Dump.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 11/10/2022

' Change Log:
'       2/25/2021:  Intial Creation, based on o_17_Import_CRE_Data
'       3/12/2021:  Replaced the fx_Update_Single_Field with more streamlined v2
'       8/20/2021:  Added a bypass for this update if the EBITDA data wasn't imported.
'       11/10/2022: Updated the ws names for the 'Dashboard 4.0' naming convention

' ****************************************************************************

' If the EBITDA worksheets weren't included in the Sageworks Raw data, skip this part
If bolSingleSheet = True Then Exit Sub

' -----------------
' Declare Variables
' -----------------

    ' Dim Worksheets
    
    Dim wsYTD As Worksheet
    Set wsYTD = wbSource.Worksheets("Dashboard 4.0 - YTD - Groups")
    
        wsYTD.Rows("1:3").Delete ' Delete the header fields
        
    Dim wsPYTD As Worksheet
    Set wsPYTD = wbSource.Worksheets("Dashboard 4.0 - PYTD - Groups")

        wsPYTD.Rows("1:3").Delete ' Delete the header fields

' -----------
' Pull in the YTD EBITDA
' -----------

    Call fx_Update_Single_Field( _
        wsSource:=wsYTD, wsDest:=wsData, _
        str_Source_TargetField:="YTD Adj. EBITDA", _
        str_Source_MatchField:="Company Name", _
        str_Dest_TargetField:="YTD Adj. EBITDA (000's)", _
        str_Dest_MatchField:="Customer")

' -----------
' Pull in the PYTD EBITDA
' -----------

    Call fx_Update_Single_Field( _
        wsSource:=wsPYTD, wsDest:=wsData, _
        str_Source_TargetField:="PYTD Adj. EBITDA", _
        str_Source_MatchField:="Company Name", _
        str_Dest_TargetField:="PYTD Adj. EBITDA (000's)", _
        str_Dest_MatchField:="Customer")

End Sub
Sub o_17_Calculate_Bright_Line_Waiver()

' Purpose: To calculate the Bright Line Waiver and compare that to what the PM/RM selected.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 6/16/2021

' Change Log:
'       6/16/2021: Intial Creation
'       6/16/2021: Switched to using two dictionaries
'       6/24/2021: Moved the copying of the Calculated BLW out of the If statement to apply the Nos to blanks, as per Eric

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    ' Assign
    
        intLastRow = WorksheetFunction.Max(wsData.Cells(Rows.count, "A").End(xlUp).Row, wsData.Cells(Rows.count, "B").End(xlUp).Row)
    
    ' Dim Loop Variables
    
    Dim i As Long
    
    ' Dim Dictionary
    
    Dim dict_BLW_Original As Scripting.Dictionary
        Set dict_BLW_Original = New Scripting.Dictionary
        
    Dim dict_BLW_Calc As Scripting.Dictionary
        Set dict_BLW_Calc = New Scripting.Dictionary
        
' -----------
' Loop through and manipulate the various fields before applying the highlighting rules
' -----------
    
    With wsData
        
        For i = 2 To intLastRow
        
        ' Load the existing BLW values into the dictionary
        If .Cells(i, col_Bright_Line_Waiver).Value2 = "" Then
            dict_BLW_Original.Add Key:=i, Item:="N"
        Else
            dict_BLW_Original.Add Key:=i, Item:=.Cells(i, col_Bright_Line_Waiver).Value2
        End If
        
        ' Calculate the true BLW values
        If .Cells(i, col_BRG) >= 8 And .Cells(i, col_BRG) <> "6W" And .Cells(i, col_FRG) >= 4 Then
            If InStr(1, .Cells(i, col_Team), "Remediation", vbTextCompare) = 0 Then
                dict_BLW_Calc.Add Key:=i, Item:="Y"
            Else
                dict_BLW_Calc.Add Key:=i, Item:="N"
            End If
        Else
            dict_BLW_Calc.Add Key:=i, Item:="N"
        End If
        
        Next i
                
    End With

' -----------
' Output the calculated Bright Line Waiver and change to red font if values are different
' -----------

    With wsData
        For i = 2 To intLastRow
            
            .Cells(i, col_Bright_Line_Waiver).Value2 = dict_BLW_Calc.Item(i) '6/24: Moved out of the If statement
            
            If dict_BLW_Original.Item(i) <> dict_BLW_Calc.Item(i) Then
                .Cells(i, col_Bright_Line_Waiver).Font.Color = vbRed
            End If
        Next i
    End With

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_17_Calculate_Bright_Line_Waiver", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_21_Flag_Basic_Anomalies()

' Purpose: To flag accounts that needs to be addressed by the PMs.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 12/29/2020

' Change Log:
'       11/5/2020: Intial Creation
'       11/6/2020: Added the code to flag blanks
'       11/9/2020: Added code to flag the "red" Unacceptable exceptions
'       11/9/2020: Added code to handle unique exceptions, such as a blank End Market for CRE
'       11/12/2020: Added the code to ONLY make the updates where there wasn't color already
'       11/20/2020: Added in the more complex rules around financials
'       12/1/2020: Disabled the QC Flags to add to a button in the UserForm
'       12/21/2020: Updated the Stale Dated Financials rule to exclude Public Sector Finance
'       12/22/2020: Added additional code to reduce the # of missing LTV: If .Cells(cell.Row, col_LFT) <> "W - ABL Leveraged" And .Cells(cell.Row, col_LFT) <> "F - Indirect Leveraged" Then
'       12/22/2020: Added additional code to reduce the # of missing LTV: If .Cells(cell.Row, col_FilterFlag) <> "Liquidation - Remaining Balance" Then
'       12/23/2020: Broke out the section for clearing flags and added to clear if R&R accoutn w/ 0 Spreads Rating Outlook
'       12/29/2020: Split out the Unique anomalies and Unique Edits
'       9/22/2021: Updated to be more explicit with the col_Last_DataFlags

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------
        
    'Dim List "Ranges"
    
    Dim col_First_DataFlags As Long
        col_First_DataFlags = fx_Create_Headers("RM", arryHeader_wsLists)

    Dim col_Last_DataFlags As Long
        col_Last_DataFlags = fx_Create_Headers("BLW (Y/N)", arryHeader_wsLists)
        
    'Dim Integers
        
    intLastRow = wsData.Cells(Rows.count, "A").End(xlUp).Row
        
    'Dim Loop Integers
    
    Dim x As Long
    
    Dim z As Long
        z = col_First_DataFlags

    Dim intLastRow_DataFlags As Long
    
    'Dim Loop Arrays / Objects

    Dim aryAcceptableValues As Variant 'Temporarily house whats OK for the field
    
    Dim aryUnacceptableValues As Variant 'Temporarily house whats NOT OK for the field
    
    'Dim Colors
    
    Dim intGreen1 As Long
        intGreen1 = RGB(235, 241, 222)
    
    Dim intRed1 As Long
        intRed1 = RGB(242, 220, 219)
        
    'Dim Loop Variables
    
    Dim cell As Variant

' -----------
' Loop through the columns to flag the Basic Anomalies (Green Fields - Includes All Acceptable Values)
' -----------

    For x = 1 To intLastCol
        For z = col_First_DataFlags To col_Last_DataFlags
            If wsData.Cells(1, x).Value2 = wsLists.Cells(1, z).Value2 And wsLists.Cells(1, z).Interior.Color = intGreen1 Then
                intLastRow_DataFlags = wsLists.Cells(Rows.count, z).End(xlUp).Row: If intLastRow_DataFlags = 1 Then intLastRow_DataFlags = 2
                aryAcceptableValues = wsLists.Range(wsLists.Cells(2, z), wsLists.Cells(intLastRow_DataFlags, z))
                
                With wsData
                    For Each cell In .Range(.Cells(2, x), .Cells(intLastRow, x))
                    
                        If Not IsError(Application.Match(cell, aryAcceptableValues, 0)) And cell.Value2 <> "" Then
                            'Do Nothing if the value is in the array
                        Else
                            cell.Interior.Color = clrOrange
                        End If
                    
                    Next cell
                End With
                
            End If
        Next z
    Next x

' -----------
' Loop through the columns to flag the Basic Anomalies (Red Fields - Includes All Non-Acceptable Values)
' -----------

    For x = 1 To intLastCol
        For z = col_First_DataFlags To col_Last_DataFlags
            If wsData.Cells(1, x).Value2 = wsLists.Cells(1, z).Value2 And wsLists.Cells(1, z).Interior.Color = intRed1 Then
                intLastRow_DataFlags = wsLists.Cells(Rows.count, z).End(xlUp).Row: If intLastRow_DataFlags = 1 Then intLastRow_DataFlags = 2
                aryUnacceptableValues = wsLists.Range(wsLists.Cells(2, z), wsLists.Cells(intLastRow_DataFlags, z))
                
                With wsData
                    For Each cell In .Range(.Cells(2, x), .Cells(intLastRow, x))
                    
                        If Not IsError(Application.Match(cell, aryUnacceptableValues, 0)) And cell.Value2 <> "" Then
                            cell.Interior.Color = clrOrange
                        ElseIf cell = aryUnacceptableValues Then
                            cell.Interior.Color = clrOrange
                        ElseIf cell.Value2 = "" Then
                            cell.Interior.Color = clrOrange
                        End If
                    
                    Next cell
                End With
                
            End If
        Next z
    Next x

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_21_Flag_Basic_Anomalies", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_22_Flag_Unique_Anomalies()

' Purpose: To flag accounts that needs to be addressed by the PMs.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 7/2/2023

' Change Log:
'       12/29/2020: Intial Creation, split off from o_21_Flag_Basic_Anomalies
'       1/27/2021:  Updated the LER anomaly flag to only apply to 8CRA, P + FW are excluded
'       8/12/2021:  Switched from 'For Each Cell' to 'For i' to clarify / normalize
'       8/12/2021:  Updated to include CRE flagging for $4MM+
'       8/12/2021:  Added a flag for a BRG that is <> 6, but is CRE Limited Monitoring
'       9/22/2021:  Added code to NOT flag a BRG that is <> 6 AND LImited Monitoring, if Market = PSF
'       12/28/2021: Added the If .Cells(i, col_LOB) = "Middle Market Banking" Then for the LImited Monitoring orange highlight
'       3/16/2022:  Updated the code for Limited Monitring for PSF for $ threshold to flag and CCRP
'                   Added CRE to the BRG <> 6 rule for Limited Monitoring
'       3/29/2022:  Removed the 'Stale Financials' file
'                   Added the bypass for MM Healthcare CRE deals to not apply a CRE rule
'       3/31/2023:  Added the code for 'Loan to Value (LER only)' to be applied to the entire SS&F book (not just LER)
'       4/3/2023:   Temporarily disabled the 'Loan to Value (LER only)' code around SS&F
'       7/2/2023:   Updated to use the 'CRE Flag' instead of LOB = 'Commercial Real Estate'

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------
        
    'Dim Dates
    
    Dim dtCurQuarter As Date
        dtCurQuarter = fx_Return_Quarter(Date)
    
    'Dim "Ranges"
        
    Dim col_GrossExp As Long
        col_GrossExp = fx_Create_Headers("Webster Commitment (000's) - Gross Exposure", arryHeader_wsData)
        
    Dim col_Reporting_Date As Long
        col_Reporting_Date = fx_Create_Headers("Reporting Date (Latest Financials Received)", arryHeader_wsData)
        
    'Dim Loop Variables
    
    Dim i As Long

' -------------------------
' Flag the Unique Anomalies
' -------------------------

    With wsData
        For i = 2 To intLastRow
        
        'Flag if 'Loan to Value (LER only)' is blank but the customer is leveraged (based on LFT code)
            If Not IsError(Application.Match(.Cells(i, col_LFT), arry_LER_Codes, 0)) And .Cells(i, col_LTV_LER) = "" Then
                If .Cells(i, col_LFT) <> "W - ABL Leveraged" _
                And .Cells(i, col_LFT) <> "F - Indirect Leveraged" _
                And .Cells(i, col_LFT) <> "P - Performance - No Longer Leveraged" Then
                    If .Cells(i, col_FilterFlag) <> "Liquidation - Remaining Balance" Then
                        .Cells(i, col_LTV_LER).Interior.Color = clrOrange
                    End If
                End If
            End If
                
        ' Flag if 'Loan to Value (LER only)' is blank or 0% and the customer is in SS&F - 3/31/2023
'            If .Cells(i, col_LOB) = "Sponsor And Specialty Finance" And (.Cells(i, col_LTV_LER) = "" Or .Cells(i, col_LTV_LER).Value = 0) Then
'                If .Cells(i, col_FilterFlag) <> "Liquidation - Remaining Balance" Then
'                    .Cells(i, col_LTV_LER).Interior.Color = clrOrange
'                End If
'            End If
                
        ' Apply the Flagging for the "Limited Monitoring" Filter Flag
            If .Cells(i, col_FilterFlag).Value2 = "Limited Monitoring" _
            And .Cells(i, col_CREFlag) <> "Yes" _
            And .Cells(i, col_LOB) <> "Asset Based Lending" _
            And .Cells(i, col_Team) <> "Public Sector Finance" _
            And .Cells(i, col_GrossExp) > 3000 Then  'In Thousands
                .Cells(i, col_FilterFlag).Interior.Color = clrOrange
            
            ' Rule for CRE > MM Healthcare Only
            ElseIf .Cells(i, col_FilterFlag).Value2 = "Limited Monitoring" _
            And Left(.Cells(i, col_Team), 15) = "Healthcare" _
            And .Cells(i, col_GrossExp) > 3000 Then  'In Thousands
                .Cells(i, col_FilterFlag).Interior.Color = clrOrange
            
            ' Rule for CRE Only
            ElseIf .Cells(i, col_FilterFlag).Value2 = "Limited Monitoring" _
            And .Cells(i, col_CREFlag) = "Yes" _
            And Left(.Cells(i, col_Team), 15) <> "Healthcare" _
            And .Cells(i, col_GrossExp) > 4000 Then  'In Thousands
                .Cells(i, col_FilterFlag).Interior.Color = clrOrange
            
            ' Rule for PSF by $
            ElseIf .Cells(i, col_FilterFlag).Value2 = "Limited Monitoring" _
            And .Cells(i, col_Team) = "Public Sector Finance" _
            And .Cells(i, col_GrossExp) > 5000 Then  'In Thousands
                .Cells(i, col_FilterFlag).Interior.Color = clrOrange
            
            ' Rule for PSF by CCRP
            ElseIf .Cells(i, col_FilterFlag).Value2 = "Limited Monitoring" _
            And .Cells(i, col_Team) = "Public Sector Finance" _
            And Left(.Cells(i, col_CCRP), 1) > 3 Then '3/16/2022: Need to test this
                .Cells(i, col_FilterFlag).Interior.Color = clrOrange
                
            End If
        
        ' Apply the Flagging to BRG if <> 6 but using the "Limited Monitoring" Filter Flag for MM and CRE
            If .Cells(i, col_FilterFlag) = "Limited Monitoring" And .Cells(i, col_BRG) <> 6 Then
                If .Cells(i, col_LOB) = "Middle Market Banking" Then
                    If .Cells(i, col_Team) <> "Public Sector Finance" Then
                        .Cells(i, col_BRG).Interior.Color = clrOrange
                    End If
                ElseIf .Cells(i, col_CREFlag) = "Yes" Then
                    .Cells(i, col_BRG).Interior.Color = clrOrange
                End If
            End If

        'Stale Financials: 'Reporting Date (Latest Financials Received)' is more than 4 quarters old
'            If .Cells(i, col_Reporting_Date).Value <> "" And .Cells(i, col_Team) <> "Public Sector Finance" Then 'Abort if it's blank to avoid an error, and abort if PSF
'                If DateValue(.Cells(i, col_Reporting_Date)) <= DateAdd("q", -5, dtCurQuarter) Then
'                    .Cells(i, col_Reporting_Date).Interior.Color = clrOrange
'                End If
'            End If
        
        'If "Financials Not Received" Filter Flag AND the Reporting Date is the most recently closed quarter
            If .Cells(i, col_FilterFlag).Value2 = "Financials Not Received (Late / Extended)" Then
                If .Cells(i, col_Reporting_Date) <> "" Then 'Abort if it's blank to avoid an error
                    If fx_Return_Quarter(.Cells(i, col_Reporting_Date).Value) = DateAdd("q", -1, dtCurQuarter) Then
                        .Cells(i, col_FilterFlag).Interior.Color = clrOrange
                    End If
                End If
            End If
        
        Next i
    End With

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_22_Flag_Unique_Anomalies", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_23_Unique_Edits()

' Purpose: To apply the edits to the data that are more unique in nature.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 7/2/2023

' Change Log:
'       12/29/2020: Intial Creation, split off from o_21_Flag_Basic_Anomalies
'       1/27/2021:  Added error handling for George Harte
'       2/11/2021:  Added in the removal of all CRE flags
'       2/19/2021:  Turned off the code to remove the flags for CRE
'       2/19/2021:  Added code to wipe the Anomaly highlight for Role for CRE customers
'       6/3/2021:   Removed the code related to George Harte
'       6/15/2021:  Removed the Spreads Ratings Outlook, as per Eric
'       7/2/2023:   Updated to use the 'CRE Flag' instead of LOB = 'Commercial Real Estate'

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------
        
    Dim cell As Variant

' ---------------------------------
' Clear flags for unique situations
' ---------------------------------

    With wsData
        
        ' Clear out the flags for Role for CRE Accounts
        'For Each cell In .Range(.Cells(2, col_LOB), .Cells(intLastRow, col_LOB))
        For Each cell In .Range(.Cells(2, col_CREFlag), .Cells(intLastRow, col_CREFlag))
            If cell.Value2 = "Yes" Then
                If Left(.Cells(cell.Row, col_Team), 15) <> "Healthcare" Then
                    .Cells(cell.Row, col_Role).Interior.Color = xlNone
                End If
            End If
        Next cell
        
    End With

Exit Sub

ErrorHandler:
Global_Error_Handling SubName:="o_23_Unique_Edits", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_24_Apply_Grey_Cell_Fill()

' Purpose: To grey out the fields that don't need to be completed, based on a Filter Flag being added.
' Trigger: Called by o_1_Import_Sageworks_Data > o_01_MAIN_PROCEDURE
' Updated: 7/2/2023

' Change Log:
'       11/5/2020: Intial Creation
'       12/17/2020: Added in the dark grey for Public Sector Finance.
'       12/22/2020: Added code to grey out the two Leverage at Close fields for any LC Only
'       12/28/2020: Updated Public Sector Finance to grey out until col_TTM_EBITDA
'       12/28/2020: Added the "Liquidation - Remaining Balance" to the "LC Only" flags
'       12/29/2020: Added the code for George Hart
'       1/4/20201: Update the formatting for George Harte to grey out Org. Type / End Market / Role / BLW.
'       1/27/2021: Added error handling for George Harte
'       2/2/2021: Updated to remove the values from a dark grey filled cell for Spreads Ratings Outlook
'       2/11/2021: Added the grey fill for CRE deals
'       2/19/2021: Added additional grey fills for CRE
'       2/19/2021: Added to grey out the CRE fields for non-CRE customers
'       4/1/2021: Mimic same logic applied to Filter Flag = "New Detail in Quarter (Spreads Not Req)" records for 'BB on ACBS'
'       6/3/2021: Updated the fields and related code to include the eight new CRE fields
'       6/3/2021: Removed the code related to George Harte
'       6/3/2021: Changed the CRE fill to use a range to do all eight fields at the same time
'       6/15/2021: Commented out the Spreads Ratings Outlook field, as per Eric
'       6/15/2021: Commented out the grey fill for Bright Line Waiver for CRE deals, as per Eric
'       6/24/2021: Expanded the CRE dark grey to include TTM Revenue and TTM Adj EBITDA, as per Eric
'       6/24/2021: Added code to grey out any Roles that are Blank for CRE deals, as per Eric
'       6/24/2021: Added code to grey out 'Reporting Date (Latest Financials Received)' and 'Basis of Financials' for CRE, as per Eric
'       6/24/2021: Added code to grey out the Green CRE Fields if Limited Monitoring or LC Only, as per Eric
'       8/12/2021: Updated to move the "Leverage Structure" and to update the "MM - BB on ACBS" to just "BB on ACBS"
'       8/12/2021: Clarified / Organized things
'       8/12/2021: Added the code to grey out the Wealth LOB
'       9/22/2021: Added the code to grey out the Asset Based Lending LOB
'       9/22/2021: Grey out "Leverage Structure" if Market = PSF
'       12/28/2021: Added the grey for LOB = "Business Banking" OR "Small Business" => Dark Gray
'                   Updated CRE filtering to only go to col_LeverageStructure instead of Basis of Financials
'                   Add the grey filtering for Subscription Lines, based on End Market
'       3/16/2022:  Removed the LER and CRG fields
'       3/23/2022:  Added back in the CRG field
'       3/29/2022:  Updated to remove the "Other" light grey and create the new array with 5 filter flags
'                   Updated the CRE LOB so that 'Reporting Date (Latest Financials Received)' and 'Basis of Financials' are also dark grey
'                   Added the LOB / Filter Flag Combo section for applying the grey fill for CRE
'                   Added the bypass for MM Healthcare CRE deals to not apply a CRE rule
'       4/25/2022:  Updated so that the % Leased is the last field for CRE
'       3/31/2023:  Added the code for 'Loan to Value (LER only)' to not grey out for the entire SS&F book
'       4/3/2023:   Temporarily disabled the 'Loan to Value (LER only)' code around SS&F
'       7/2/2023:   Updated to use the 'CRE Flag' instead of LOB = 'Commercial Real Estate'

' ****************************************************************************

'On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    'Dim "Ranges"
    
    Dim col_Start_YTD_Adj_EBITDA As Long
        col_Start_YTD_Adj_EBITDA = fx_Create_Headers("YTD Adj. EBITDA (000's)", arryHeader_wsData)
        
    Dim col_End_LTV As Long
        col_End_LTV = col_LTV_LER
        
    Dim col_CRG As Long
        col_CRG = fx_Create_Headers("CRG", arryHeader_wsData)
        
    Dim col_SeniorDebt As Long
        col_SeniorDebt = fx_Create_Headers("Senior Debt / Adj. EBITDA", arryHeader_wsData)
        
    Dim col_TotalDebt_AtClose As Long
        col_TotalDebt_AtClose = fx_Create_Headers("Total Debt / Adj. EBITDA (at Initial Close)", arryHeader_wsData)
        
    Dim col_Chng_YTD_PYTD As Long
        col_Chng_YTD_PYTD = fx_Create_Headers("% Chng. (YTD vs. PYTD) (calculated)", arryHeader_wsData)
        
    Dim col_TTM_EBITDA As Long
        col_TTM_EBITDA = fx_Create_Headers("TTM EBITDA Addbacks ($000's)", arryHeader_wsData)
        
    Dim col_TTM_EBITDA_Calc As Long
        col_TTM_EBITDA_Calc = fx_Create_Headers("TTM % EBITDA Addbacks (calculated)", arryHeader_wsData)
        
    Dim col_SeniorDebt_AtClose As Long
        col_SeniorDebt_AtClose = fx_Create_Headers("Senior Debt / Adj. EBITDA (at Initial Close)", arryHeader_wsData)
        
    Dim col_ARR As Long
        col_ARR = fx_Create_Headers("ARR (000's)", arryHeader_wsData)

    Dim col_DebtYield As Long
        col_DebtYield = fx_Create_Headers("Debt Yield (UW)", arryHeader_wsData)

    'Dim col_CollatPropState as Long
    '    col_CollatPropState = fx_Create_Headers("Collateral Property State", arryHeader_wsData)
        
    Dim col_LeverageStructure As Long
        col_LeverageStructure = fx_Create_Headers("Leverage Structure", arryHeader_wsData)
        
    Dim col_CRE_Leased As Long
        col_CRE_Leased = fx_Create_Headers("% Leased", arryHeader_wsData)
                
    'Dim Loop Values
    
    Dim i As Long
        
    Dim strFilterField As String
    
' --------------------
' Apply Grey Cell Fill
' --------------------
    
With wsData
    For i = 2 To intLastRow
        
        strFilterField = .Cells(i, col_FilterFlag).Value2
        
    ' ------------------------------------
    ' Apply Grey Fill based on filter flag
    ' ------------------------------------
        
        ' If the customer has a Filter Flag then fill dark grey
        If UBound(Filter(Get_FilterFlag_Array_DarkGrey, strFilterField)) > -1 = True And strFilterField <> "" Then
            .Range(.Cells(i, col_Start_YTD_Adj_EBITDA), .Cells(i, col_End_LTV)).Interior.Color = clrDarkGray
        End If
        
        ' If the customer has a Filter Flag then fill light grey
        If UBound(Filter(Get_FilterFlag_Array_LightGrey, strFilterField)) > -1 = True And strFilterField <> "" Then
            .Range(.Cells(i, col_Start_YTD_Adj_EBITDA), .Cells(i, col_End_LTV)).Interior.Color = clrLightGray
        End If
        
        ' If the value is "Other" then fill light grey '3/29/22 Replaced by light gray filter
        'If strFilterField = "Other" Then
            '.Range(.Cells(i, col_Start_YTD_Adj_EBITDA), .Cells(i, col_End_LTV)).Interior.Color = clrLightGray
       ' End If
        
        ' Grey out the two Leverage at Close fields for any LC Only or Liquidation - Remaining Balance
        If .Cells(i, col_FilterFlag) = "LC Only" Or .Cells(i, col_FilterFlag) = "Liquidation - Remaining Balance" Then
            .Cells(i, col_TotalDebt_AtClose).Interior.Color = clrDarkGray
            .Cells(i, col_SeniorDebt_AtClose).Interior.Color = clrDarkGray
        End If

        ' Grey out CRG for the Liquidation Accounts
        If .Cells(i, col_FilterFlag) = "Liquidation Accounts" Then
            .Cells(i, col_CRG).Interior.Color = clrDarkGray
        End If
   
    ' ---------------------------------------
    ' Apply Grey Fill based on other criteria
    ' ---------------------------------------

        ' Grey out the fields for Subscription Lines
        If .Cells(i, col_EndMarket) = "Finance - Subscription Lines" Then
            .Range(.Cells(i, col_SeniorDebt_AtClose), .Cells(i, col_LTV_LER)).Interior.Color = clrDarkGray
        End If
        
    ' -------------------------------------
    ' Apply Grey Fill based on LOB / Market
    ' -------------------------------------
       
        ' Market = 'Public Sector Finance' => Dark Gray
        If .Cells(i, col_Team) = "Public Sector Finance" Then
            .Cells(i, col_CRG).Interior.Color = clrDarkGray
            .Cells(i, col_TTM_EBITDA).Interior.Color = clrDarkGray
            .Cells(i, col_TTM_EBITDA_Calc).Interior.Color = clrDarkGray
            .Cells(i, col_ARR).Interior.Color = clrDarkGray
            .Cells(i, col_LeverageStructure).Interior.Color = clrDarkGray
            
            .Range(.Cells(i, col_Bright_Line_Waiver), .Cells(i, col_Chng_YTD_PYTD)).Interior.Color = clrDarkGray
            .Range(.Cells(i, col_SeniorDebt), .Cells(i, col_End_LTV)).Interior.Color = clrDarkGray
        End If
        
        ' LOB = 'Commercial Real Estate' => Dark Gray
        If .Cells(i, col_CREFlag) = "Yes" And Left(.Cells(i, col_Team), 15) <> "Healthcare" Then
            .Cells(i, col_EndMarket).Interior.Color = clrDarkGray
            .Cells(i, col_CRG).Interior.Color = clrDarkGray
            
            .Range(.Cells(i, col_SeniorDebt_AtClose), .Cells(i, col_LTV_LER)).Interior.Color = clrDarkGray '3/29/2022 Update
            
            If .Cells(i, col_Role).Value2 = "" Then .Cells(i, col_Role).Interior.Color = clrDarkGray
            
            If .Cells(i, col_FilterFlag).Value2 = "Limited Monitoring" Or .Cells(i, col_FilterFlag).Value2 = "LC Only" Then
                .Range(.Cells(i, col_DebtYield), .Cells(i, col_CRE_Leased)).Interior.Color = clrDarkGray
            End If
            
        End If
        
        ' LOB = 'Asset Based Lending' => Dark Gray
        If .Cells(i, col_LOB) = "Asset Based Lending" Then
            .Cells(i, col_SeniorDebt_AtClose).Interior.Color = clrDarkGray
            .Cells(i, col_TotalDebt_AtClose).Interior.Color = clrDarkGray
            .Cells(i, col_ARR).Interior.Color = clrDarkGray
            .Cells(i, col_LeverageStructure).Interior.Color = clrDarkGray
            .Cells(i, col_LTV_LER).Interior.Color = clrDarkGray
        End If

        ' LOB = 'Wealth' => Dark Gray
        If .Cells(i, col_LOB) = "Wealth" Then
            .Range(.Cells(i, col_SeniorDebt_AtClose), .Cells(i, col_CRE_Leased)).Interior.Color = clrDarkGray
            
            .Cells(i, col_EndMarket).Interior.Color = clrDarkGray
            .Cells(i, col_CRG).Interior.Color = clrDarkGray
            .Cells(i, col_Role).Interior.Color = clrDarkGray
        End If
        
        If .Cells(i, col_LOB) = "Business Banking" Or .Cells(i, col_LOB) = "Small Business" Then
            .Cells(i, col_Sponsor).Interior.Color = clrDarkGray
            .Cells(i, col_EndMarket).Interior.Color = clrDarkGray
            .Cells(i, col_CRG).Interior.Color = clrDarkGray
            .Cells(i, col_Role).Interior.Color = clrDarkGray
        End If
        
        ' LOB <> 'Commercial Real Estate' => Dark Gray
        If .Cells(i, col_CREFlag) <> "Yes" Or Left(.Cells(i, col_Team), 15) = "Healthcare" Then
            .Range(.Cells(i, col_DebtYield), .Cells(i, col_CRE_Leased)).Interior.Color = clrDarkGray
        End If
                
    ' -----------------------------------------------------
    ' Apply Grey Fill based on combo of LOB and Filter Flag
    ' -----------------------------------------------------
        
        If .Cells(i, col_CREFlag) = "Yes" And Left(.Cells(i, col_Team), 15) <> "Healthcare" Then
            If .Cells(i, col_FilterFlag).Value2 = "Limited Monitoring" Or _
            .Cells(i, col_FilterFlag).Value2 = "LC Only" Or _
            .Cells(i, col_FilterFlag).Value2 = "MM HC Construction" Or _
            .Cells(i, col_FilterFlag).Value2 = "MM HC Fill-Up" Or _
            .Cells(i, col_FilterFlag).Value2 = "New Deal in Quarter (Spreads Not Req)" Then
                .Range(.Cells(i, col_DebtYield), .Cells(i, col_CRE_Leased)).Interior.Color = clrDarkGray
            End If
            
            If .Cells(i, col_FilterFlag).Value2 = "No Historical" Then
                .Range(.Cells(i, col_DebtYield), .Cells(i, col_CRE_Leased)).Interior.Color = clrLightGray
            End If
            
        End If
        
    Next i
        
End With
        
            
' ----------------------------------------------------------------
' Remove the Grey Cell fill for 'Loan to Value' if the LOB is SS&F
' ----------------------------------------------------------------
    
'With wsData 'Added 3/31/23
'    For i = 2 To intLastRow
'
'        If .Cells(i, col_LOB) = "Sponsor And Specialty Finance" Then
'            If .Cells(i, col_LTV_LER).Interior.Color = clrLightGray Or .Cells(i, col_LTV_LER).Interior.Color = clrDarkGray Then
'                If .Cells(i, col_FilterFlag) <> "Liquidation - Remaining Balance" Then
'                    '.Cells(i, col_LTV_LER).Interior.Color = RGB(255, 255, 255)
'                    .Cells(i, col_LTV_LER).Interior.Color = xlNone
'                End If
'            End If
'        End If
'
'        ' Flag if 'Loan to Value (LER only)' is blank or 0% and the customer is in SS&F - 3/31/2023
'        If .Cells(i, col_LOB) = "Sponsor And Specialty Finance" And (.Cells(i, col_LTV_LER) = "" Or .Cells(i, col_LTV_LER).Value = 0) Then
'            If .Cells(i, col_FilterFlag) <> "Liquidation - Remaining Balance" Then
'                .Cells(i, col_LTV_LER).Interior.Color = clrOrange
'            End If
'        End If
'
'    Next i
'
'End With
                
Exit Sub
    
ErrorHandler:

Global_Error_Handling SubName:="o_24_Apply_Grey_Cell_Fill", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
End Sub
Sub o_25_Apply_Formulas()

' Purpose: To copy the formulas from the FORMULAS ws to the Data tab.
' Trigger: Called by o_01_MAIN_PROCEDURE
' Updated: 2/17/2021

' Change Log:
'       11/6/2020: Intial Creation
'       2/17/2021: Added the code to wipe the change flag for the Weekly file

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------
   
    'Dim "Ranges"
        
    Dim intFirstCol_wsData As Long
        intFirstCol_wsData = 1
        
    Dim intLastCol_wsData As Long
        intLastCol_wsData = fx_Create_Headers("Change Flag", arryHeader_wsData)
        
    Dim col_ChangeFlag As Long
        col_ChangeFlag = intLastCol_wsData
        
    'Dim Loop Variables
    
    Dim x As Long
    
    Dim i As Long
    
    Dim strFormula As String
    
' -----------
' Copy the Formulas into wsData
' -----------

    For x = intFirstCol_wsData To intLastCol_wsData
        i = 2
        
        Do Until wsFormulas.Range("A" & i) <> wsData.Name
        
            If wsData.Cells(1, x) = wsFormulas.Cells(i, 3) Then
                strFormula = wsFormulas.Cells(i, 4).Value2
                wsData.Range(wsData.Cells(2, x), wsData.Cells(intLastRow, x)).Formula = strFormula
                'Exit For
            End If
                            
            i = i + 1

        Loop
    Next x

' -----------
' Copy back values only for "regular" dashboard and remove the Change Flag
' -----------

#If Quarterly = 0 Then

    wsData.Range(wsData.Cells(2, 1), wsData.Cells(intLastRow, intLastCol)).Value2 = _
    wsData.Range(wsData.Cells(2, 1), wsData.Cells(intLastRow, intLastCol)).Value2
    
    wsData.Range(wsData.Cells(2, col_ChangeFlag), wsData.Cells(intLastRow, col_ChangeFlag)).Value2 = ""
    
#End If

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_25_Apply_Formulas", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_31_Validate_Control_Totals()

' Purpose: To validate that the control totals for the data imported match.
' Trigger: Called by uf_Sageworks_Regular
' Updated: 2/12/2021

' Change Log:
'       2/12/2021: Intial Creation

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------
    
    ' Dim Integers
    
    Dim int1stTotal As Long
        int1stTotal = wsValidation.Range("D2").Value
    
    Dim int2ndTotal As Long
        int2ndTotal = wsValidation.Range("D3").Value
    
    Dim int1stCount As Long
        int1stCount = wsValidation.Range("E2").Value
    
    Dim int2ndCount As Long
        int2ndCount = wsValidation.Range("E3").Value
        
    ' Dim Booleans
    
    Dim bolTotalsMatch As Boolean
        If int1stTotal = int2ndTotal And int1stCount = int2ndCount Then
            bolTotalsMatch = True
        Else
            bolTotalsMatch = False
        End If
    
' -----------
' Output the messagebox with the results
' -----------
   
    If bolTotalsMatch = True Then
    MsgBox Title:="You're Golden boiiiii!", _
        Buttons:=vbOKOnly, _
        Prompt:="The validation totals match between the source and output." & Chr(10) & Chr(10) _
        & "Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "Validation Count: " & Format(int1stCount, "0,0")
       
    ElseIf bolTotalsMatch = False Then
    MsgBox Title:="Validation Totals Don't Match", _
        Buttons:=vbCritical, _
        Prompt:="The validation totals from the Sageworks Dashboard Dump don't match what was imported. " _
        & "Please review the totals in the Validation worksheet to determine what went awry. " _
        & "Once the issues has been identifed talk to James to fix / reimport the data." & Chr(10) & Chr(10) _
        & "1st Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "1st Validation Total Variance: " & Format(int1stTotal - int2ndTotal, "$#,##0") & Chr(10) & Chr(10) _
        & "1st Validation Count: " & Format(int1stCount, "0,0") & Chr(10) _
        & "1st Validation Count Variance: " & Format(int1stCount - int2ndCount, "0,0")
    
    End If

Exit Sub
    
ErrorHandler:

myPrivateMacros.DisableForEfficiencyOff
    
End Sub
Sub o_41_Create_Complete_Dashboard()

' Purpose: To create a copy of the dashboard that includes all of the borrowers.
' Trigger: Called by uf_Sageworks_Regular
' Updated: 7/26/2022

' Change Log:
'       7/26/2022: Intial Creation

' ****************************************************************************

' -----------------
' Copy the worksheet
' -----------------
    
    Call fx_Sheet_Exists(ThisWorkbook.Name, "Complete Dashboard", True)
        ThisWorkbook.Sheets("Dashboard Review").Copy After:=ThisWorkbook.Sheets("Dashboard Review")
        ActiveSheet.Name = "Complete Dashboard"
    
End Sub
Sub o_42_Remove_Wealth_and_Business_Banking()

' Purpose: To remove the Wealth and Business Banking customers from the population in the 'Dashboard Review' ws.
' Trigger: Called by uf_Sageworks_Regular
' Updated: 7/26/2022

' Change Log:
'       7/26/2022: Intial Creation

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

   'Dim intLastRow As Long
        intLastRow = wsData.Cells(Rows.count, "A").End(xlUp).Row

' -----------------------------------------
' Remove the data for BB/SB and Wealth LOBs
' -----------------------------------------
    
On Error Resume Next
    
    With wsData
                
        ' Remove Business Banking / Small Business Borrowers
        .Range("A1").AutoFilter Field:=col_LOB, Criteria1:=Array("Small Business", "Business Banking", "Wealth"), Operator:=xlFilterValues
            .Range("A2:A" & intLastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                .Range("A1").AutoFilter Field:=col_LOB
    
        ' Remove Wealth Borrowers
        .Range("A1").AutoFilter Field:=col_LOB, Criteria1:=Array("Wealth"), Operator:=xlFilterValues
            .Range("A2:A" & intLastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                .Range("A1").AutoFilter Field:=col_LOB

    End With

On Error GoTo 0
    
End Sub
Sub o_5_Create_Workout_LOB()

' Purpose: To refresh the wsLists sheet with updated data from the wsData.
' Trigger: Called by o_01_MAIN_PROCEDURE
' Updated: 6/29/2023

' Change Log:
'       11/4/2020:  Intial Creation
'       11/24/2020: Refreshed to replace the hard coded fields with "ranges".
'       11/25/2020: Brokeout the refresh of the data from the creation of the arrays.
'       3/29/2022:  Updated the code for intLastRow to reset based on wsLists
'       10/6/2022:  Broke out the intLastRow for the wsLists and wsData
'                   Moved all of the output to the Dashboard, created the Updated LOB column
'       6/29/2023:  Removed the code for non-Workout borrowers, now that we pull those in via lookup in o_14_Manipulate_Sageworks_Customer_Data

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------
    
    intLastRow = fx_Find_LastRow(wsData)
    
    'Dim Loop Values
    
    Dim i As Long
    
' --------------------------------
' Update the LOB for R&R Customers
' --------------------------------

    For i = 2 To intLastRow
        If InStr(1, wsData.Cells(i, col_Team), "Rem") > 0 Then
            wsData.Cells(i, col_LOBUpdated).Value2 = "Commercial Workout"
        End If
    Next i

Exit Sub
    
ErrorHandler:

Global_Error_Handling SubName:="o_5_Create_Workout_LOB", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
End Sub


