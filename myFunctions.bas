Attribute VB_Name = "myFunctions"
Option Explicit
Function Global_Error_Handling(SubName, ErrSource, ErrNum, ErrDesc)

ThisWorkbook.Activate

    Dim strTempVer As String
       strTempVer = [V_Dashboard_Version]

If Err.Number <> 0 Then MsgBox _
    Title:="I am Error", _
    Buttons:=vbCritical, _
    Prompt:="Something went awry with the Dashboard, try to hit OK and redo the last step. " _
    & "If that doesn't resolve it then reach out to James Rinaldi in Credit Analytics for a fix. " _
    & "This tool has a growth mindset, with each issue addressed we itterate to a better version." & Chr(10) & Chr(10) _
    & "Please take a screenshot of this message, and send it to James." & Chr(10) _
    & "Include a brief description of what you were doing when it occurred." & Chr(10) _
    & "If you get this message after hitting 'Email Credit Risk' please save the dashboard, and also send that to James." & Chr(10) & Chr(10) _
    & "Error Source: " & ErrSource & " " & strTempVer & Chr(10) _
    & "Subroutine: " & SubName & Chr(10) _
    & "Error Desc.: #" & ErrNum & " - " & ErrDesc & Chr(10)

myPrivateMacros.DisableForEfficiencyOff

End

'Or include all of the details in an auto email to me and just prompt them for what happened.

End Function
Function fx_Create_Headers(strHeaderTitle As String, arryHeader As Variant)

' Purpose: To determine the column number for a specific title in the header.
' Trigger: Called
' Updated: 12/11/2020

' Change Log:
'       5/1/2020: Intial Creation
'       12/11/2020: Updated to use an array instead of the range, reducing the time to run by 75%.

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim i As Long
    
    Dim arryTEMP()

' -----------
' Loop through the array
' -----------

    For i = LBound(arryHeader) To UBound(arryHeader)
        If arryHeader(i, 1) = strHeaderTitle Then
            fx_Create_Headers = i
            Exit Function
        End If
    Next i

End Function
Function fx_Create_Unique_List(rngListValues As Range)

' Purpose: To create a unique list of values based on the passed range.
' Trigger: Called
' Updated: 5/5/2020

' Change Log:
'          5/5/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
           
    Dim strsUniqueValues As New Collection
    
    Dim strValue As Variant
    
    Dim arryTempData()
           
' -----------
' Copy in the data selected for rngListValues into the array, then into the collection
' -----------

    arryTempData = Application.Transpose(rngListValues)

On Error Resume Next 'When a duplicate is found skip it, instead of erroring
    
    For Each strValue In arryTempData
        strsUniqueValues.Add strValue, strValue
    Next

On Error GoTo 0

' -----------
' Pass the collection of unique values
' -----------

Set fx_Create_Unique_List = strsUniqueValues
    Debug.Print fx_Create_Unique_List.count

End Function
Function fx_Name_Reverse()

' Purpose: This function splits and reverses a user name from LAST, FIRST to FIRST LAST
' Trigger: Called
' Updated: 4/1/2020

' Change Log:
'       3/23/2020: Fixed an issue with people with middle names (ex."Elias, Richard J.")
'       4/1/2020: Fixed an issue with people with unique name formatting (First Last) that was breaking due to the missing comma.

' ****************************************************************************

On Error GoTo ErrorHandler

Dim str_User_Name As String
    str_User_Name = Application.UserName

If InStr(str_User_Name, ",") = False Then 'If they have a unique name then abort
    fx_Name_Reverse = Replace(str_User_Name, ".", "")
    Exit Function
End If

Dim str_First_Name As String
    str_First_Name = Right(str_User_Name, Len(str_User_Name) - InStrRev(str_User_Name, ",") - 1)

Dim str_Last_Name As String
    str_Last_Name = Left(str_User_Name, InStrRev(str_User_Name, ",") - 1)

Dim str_Full_Name As String
    str_Full_Name = str_First_Name & " " & str_Last_Name

fx_Name_Reverse = Replace(str_Full_Name, ".", "") 'Output the new user name, removes any periods after a middle initial

Exit Function

ErrorHandler:

MsgBox "There was an error with your username, please let James Rinaldi know and he'll fix it."

End Function
Function fx_Privileged_User()

' Purpose: To output if the user is on the Privileged User list or not.
' Trigger: Called
' Updated: 9/26/2022

' Change Log:
'       9/23/2020:  Intial Creation
'       12/17/2020: Added the conditional compiler constant to determine if DebugMode was on, if so make Priviledged User false.
'       9/26/2022:  Added Allison Basili
    
' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strUserID As String
        strUserID = Application.UserName

    Dim bolPrivilegedUser As Boolean

' ---------------------------------
' Determine if they are on the list
' ---------------------------------

    If _
        strUserID = "Rinaldi, James" Or _
        strUserID = "Rauckhorst, Eric W." Or _
        strUserID = "Basili, Allison" Or _
        strUserID = "Duarte Espinoza, Axcel S." Or _
        strUserID = "Barcikowski, Melissa H." Or _
        strUserID = "DeLuca, Kayla R." Or _
        strUserID = "Renzulli, Scott W." _
    Then
        bolPrivilegedUser = True
    Else
        bolPrivilegedUser = False
    End If

    #If DebugMode = 1 Then
        bolPrivilegedUser = False
    #End If

    fx_Privileged_User = bolPrivilegedUser

End Function
Function fx_Data_Validation_Control_Totals(wsDataSource As Worksheet, strModuleName As String, strSourceName As String, intHeaderRow As Long, strControlTotalField As String, Optional intLastRowtoImport As Long)

' Purpose: To output the data validation control totals to the wsValidation, if it exists.
' Trigger: Called
' Updated: 2/12/2021

' Change Log:
'       9/26/2020: Intial Creation
'       11/3/2020: Updated to activate ThisWorkbook before checking for the Validation ws
'       2/12/2021: Added the code for intLastRowtoImport
    
' ****************************************************************************

ThisWorkbook.Activate

    ' Only run of the VALIDATION ws exists
    If Evaluate("ISREF(" & "'VALIDATION'" & "!A1)") = False Then
        Debug.Print "fx_Data_Validation_Control_Totals failed becuase there is no ws called 'VALIDATION' in the Workbook"
        Exit Function
    End If

' -----------
' Declare your variables
' -----------

    'Dim Worksheets

    Dim wsValidation As Worksheet
        Set wsValidation = ThisWorkbook.Sheets("VALIDATION")

    Dim wsSource As Worksheet
        Set wsSource = wsDataSource

    ' Dim Cell References

    Dim intLastCol As Long
        intLastCol = wsSource.Cells(1, Columns.count).End(xlToLeft).Column
      
    Dim intCurRow As Long
        intCurRow = wsValidation.Cells(Rows.count, "A").End(xlUp).Row + 1

    'Dim "Ranges"
    
    Dim arryHeader() As Variant
        arryHeader = Application.Transpose(wsSource.Range(wsSource.Cells(intHeaderRow, 1), wsSource.Cells(intHeaderRow, intLastCol)))
        
    Dim intColTotals As Long
        intColTotals = fx_Create_Headers(strControlTotalField, arryHeader)
    
    'Dim Integers

    Dim intRecordCount As Long
        If intLastRowtoImport > 0 Then
            intRecordCount = intLastRowtoImport - intHeaderRow
        Else
            intRecordCount = WorksheetFunction.Max( _
            wsSource.Cells(Rows.count, "A").End(xlUp).Row, _
            wsSource.Cells(Rows.count, "B").End(xlUp).Row) - intHeaderRow
        End If
    'Dim Other Variables

    Dim strCol_Totals As String
        strCol_Totals = Split(Cells(1, intColTotals).Address, "$")(1)
    
    Dim strRng_Totals As String
        strRng_Totals = strCol_Totals & "1:" & strCol_Totals & intLastRowtoImport
        
    Dim dblTotals As Double
        If intLastRowtoImport > 0 Then
            dblTotals = Round(Application.WorksheetFunction.Sum(wsSource.Range(strRng_Totals)), 2)
        Else
            dblTotals = Round(Application.WorksheetFunction.Sum(wsSource.Range(strCol_Totals & ":" & strCol_Totals)), 2)
        End If
' -----------
' Output the validation totals from the passed variables
' -----------

    With wsValidation
        .Range("A" & intCurRow) = Format(Now, "m/d/yyyy hh:mm")   'Date / Time
        .Range("B" & intCurRow) = strModuleName                   'Code Module
        .Range("C" & intCurRow) = strSourceName                   'Source
        .Range("D" & intCurRow) = Format(dblTotals, "$#,##0")     'Total
        .Range("E" & intCurRow) = Format(intRecordCount, "0,0")   'Count
    End With

End Function

Function fx_Reverse_Given_Name(strNametoReverse)

' Purpose: This function splits and reverses a name from LAST, FIRST to FIRST LAST
' Trigger: Called
' Updated: 12/8/2020

' Change Log:
'       12/8/2020: Initial creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    If InStr(strNametoReverse, ",") = False Then 'If they have a unique name then abort
        'fx_Reverse_Given_Name = Replace(strNametoReverse, ".", "")
        fx_Reverse_Given_Name = strNametoReverse
        Exit Function
    End If

Dim str_First_Name As String
    str_First_Name = Right(strNametoReverse, Len(strNametoReverse) - InStrRev(strNametoReverse, ",") - 1)

Dim str_Last_Name As String
    str_Last_Name = Left(strNametoReverse, InStrRev(strNametoReverse, ",") - 1)

Dim str_Full_Name As String
    str_Full_Name = str_First_Name & " " & str_Last_Name

' -----------
' Output the name
' -----------

fx_Reverse_Given_Name = Replace(str_Full_Name, ".", "") 'Output the new user name, removes any periods after a middle initial

End Function
Function fx_QuickSort(coll As Collection, intFirstRecordNum As Long, intLastRecordNum As Long) As Collection

' Purpose: To sort the bassed collection or array alphabetically.
' Trigger: Called
' Updated: 10/10/2022

' Change Log:
'       10/10/2022: Overhauled to clarify variables and allow for an array to be passed

' ***********************************************************************************************************************************

' Use Example: _
    Set coll_SortedPMs = fx_QuickSort(coll_UniquePMs, 1, coll_UniquePMs.count)

' LEGEND MANDATORY:
'   coll: The collection to be sorted
'   intFirstRecordNum:
'   intLastRecordNum:

' LEGEND OPTIONAL:
'

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim vCentreVal As Variant
    Dim vTemp As Variant
    
    Dim intTempLow As Long
        intTempLow = intFirstRecordNum
    
    Dim intTempHigh As Long
        intTempHigh = intLastRecordNum
    
' ---------------
' Sort the values
' ---------------
    
    vCentreVal = coll((intFirstRecordNum + intLastRecordNum) \ 2)
    Do While intTempLow <= intTempHigh
  
    Do While coll(intTempLow) < vCentreVal And intTempLow < intLastRecordNum
      intTempLow = intTempLow + 1
    Loop
    
    Do While vCentreVal < coll(intTempHigh) And intTempHigh > intFirstRecordNum
      intTempHigh = intTempHigh - 1
    Loop
    
    If intTempLow <= intTempHigh Then
    
      ' Swap values
      vTemp = coll(intTempLow)
      
      coll.Add coll(intTempHigh), After:=intTempLow
      coll.Remove intTempLow
      
      coll.Add vTemp, Before:=intTempHigh
      coll.Remove intTempHigh + 1
      
      ' Move to next positions
      intTempLow = intTempLow + 1
      intTempHigh = intTempHigh - 1
      
    End If
    
  Loop
  
  If intFirstRecordNum < intTempHigh Then fx_QuickSort coll, intFirstRecordNum, intTempHigh
  If intTempLow < intLastRecordNum Then fx_QuickSort coll, intTempLow, intLastRecordNum
  
  Set fx_QuickSort = coll
  
End Function

Function fx_Hide_Worksheets_For_Users()

' Purpose: To hide the extra worksheets that the "regular" users don't need to see.
' Trigger: Called
' Updated: 12/22/2020

' Change Log:
'       12/22/2020: Initial creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim ws As Worksheet

' -----------
' Loop through the worksheet names
' -----------

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard Review" And ws.Name <> "INSTRUCTIONS" Then
            ws.Visible = xlSheetHidden
        Else
            ws.Visible = xlSheetVisible
        End If
    Next ws

End Function
Function fx_Update_Data_SIC(strMatchField_Cur As String, strTargetField_Cur As String, strMatchField_Lookup As String, strTargetField_Lookup As String, arryCurData() As Variant, arryLookupData() As Variant)

' Purpose: To update the values for a given field using a second data source.
' Trigger: Called
' Updated: 2/17/2021
' Use Example: Application.Transpose(fx_Update_Data_SIC("Customer", "SIC Code", "Customer3", "SIC_Code3", arryCurData, arryTargetData))

' Change Log:
'       12/23/2020: Initial creation
'       12/23/2020: Made more robust by using the arryHeader and field lookups
'       12/28/2020: Added in the Unknown SIC Codes
'       2/17/2021: Updated to make the lookup case-insensitve

' ****************************************************************************

' -----------
' Declare your variables
' -----------
        
    ' Declare Current Data "Ranges"
    
    Dim arryHeader_Cur() As Variant
        arryHeader_Cur = Application.Transpose(Application.Index(arryCurData, 1, 0))
    
    Dim col_MatchField_Cur As Long
        col_MatchField_Cur = fx_Create_Headers(strMatchField_Cur, arryHeader_Cur)
    
    Dim col_TargetField_Cur As Long
        col_TargetField_Cur = fx_Create_Headers(strTargetField_Cur, arryHeader_Cur)
    
    ' Declare Lookup Data "Ranges"
    
    Dim arryHeader_Lookup() As Variant
        arryHeader_Lookup = Application.Transpose(Application.Index(arryLookupData, 1, 0))

    Dim col_MatchField_Lookup As Long
        col_MatchField_Lookup = fx_Create_Headers(strMatchField_Lookup, arryHeader_Lookup)
    
    Dim col_TargetField_Lookup As Long
        col_TargetField_Lookup = fx_Create_Headers(strTargetField_Lookup, arryHeader_Lookup)

    ' Declare Arrays
    
    Dim arryTargetFieldOnly_Cur() As Variant
        arryTargetFieldOnly_Cur = Application.Transpose(Application.Index(arryCurData, 0, col_TargetField_Cur))
        
    ' Declare Dictionaries
    
    Dim dict_LookupData As Scripting.Dictionary
        Set dict_LookupData = New Scripting.Dictionary
        dict_LookupData.CompareMode = TextCompare
        
    ' Declare Loop Variables
    
    Dim i As Long

' -----------
' Fill the Dictionary with the Lookup Data
' -----------
    
    On Error Resume Next
        For i = 1 To UBound(arryLookupData)
            dict_LookupData.Add Key:=arryLookupData(i, col_MatchField_Lookup), Item:=arryLookupData(i, col_TargetField_Lookup)
        Next i
    On Error GoTo 0
    
' -----------
' Loop through the Lookup Data to pull in the updated data for the target field
' -----------

    For i = 2 To UBound(arryCurData)
        If arryCurData(i, col_TargetField_Cur) = "9900" Or arryCurData(i, col_TargetField_Cur) = "Unknown" Then
            If dict_LookupData.Exists(arryCurData(i, col_MatchField_Cur)) Then
               arryTargetFieldOnly_Cur(i) = dict_LookupData.Item(arryCurData(i, col_MatchField_Cur))
            End If
        End If
    Next i

    'Output the values from the array
    fx_Update_Data_SIC = arryTargetFieldOnly_Cur

End Function
Function fx_Update_Single_Field(wsSource As Worksheet, wsDest As Worksheet, _
    str_Source_TargetField As String, str_Source_MatchField As String, str_Dest_TargetField As String, str_Dest_MatchField As String, _
    Optional int_SourceHeaderRow As Long, Optional bol_ConvertMatchSourcetoValues As Boolean, _
    Optional bol_CloseSourceWb As Boolean, Optional bol_SkipDuplicates As Boolean, Optional bol_BlanksOnly As Boolean, _
    Optional str_OnlyUseValue As String, Optional arry_OnlyUseMultipleValues As Variant, Optional bol_MultipleOnlyUseValues As Boolean, _
    Optional bol_MissingLookupData_MsgBox As Boolean, Optional bol_MissingLookupData_UseExistingData As Boolean, _
    Optional strMissingLookupData_ValuetoUse, Optional strWsNameLookup As String, _
    Optional str_FilterField_Dest As String, Optional str_FilterValue As String, Optional bol_FilterPassArray As Boolean)

' Purpose: To update the data in the Target Field in the Destination, based on data from the Target Field in the Source.
    ' 1) Add the data to be referenced into the dictionary
    ' 2) Look for the data to be updated for matches to the reference dictionary
    ' 3) Replace the existing data with the reference data

' Trigger: Called
' Updated: 6/13/2023

' Change Log:
'       2/16/2021:  Initial creation, based on fx_Update_Data_SIC
'       2/17/2021:  Updated to convert over to pulling in the applicable ranges.
'       2/26/2021:  Tweaked the names of the paramaters
'       2/26/2021:  Rewrote to include as much of the code as possible in the function.
'       6/22/2021:  Added the bol_CloseSourceWb variable and related code.
'       6/30/2021:  Added the code to ignore duplicates, just output the value once
'       7/14/2021:  Added the option to convert the Match_Source to values (for Acct #s w/ leading 0s)
'       7/14/2021:  Added the option to pass the int_SourceHeaderRow and the related code
'       10/5/2021:  Updated to use the passed Target & Match fields to determine the intLastRow
'       10/11/2021: Added the option for bol_BlanksOnly
'       10/12/2021: Added the option to ONLY update with a single value, if present (Ex. updating NPL flag for a borrower)
'       10/13/2021: Updated to convert the range from Text => General formatting, if bol_ConvertMatchSourcetoValues = True
'       4/18/2022:  Added the MsgBox for any missing data, and the bol to use it
'                   Updated the code to determine if there are missing fields to use a dictionary, and created a process to handle blanks
'       4/19/2022:  Added 'strMissingLookupData_ValuetoUse' to allow a user to pass a value that will be used for blanks
'                   Updated the names of some of the variables to help clarify
'       6/13/2022:  Added the 'strWsNameLookup' and related code when a value is missing.
'                   Added code to remove the leading line break in str_missingvalues
'       9/15/2022:  Added the 'Or InStr(1, str_FilterValue, arry_Dest_Filter(i)) > 0' to allow an 'array' to be passed as a criteria
'                   Simplified how the Arrays are determined
'       6/13/2023:  Added a simple example and some clarifications
'       6/29/2023:  Updated to allow multiple strings to be passed to str_OnlyUseValue and arryTemp_OnlyUseValue to determine if the value is in the data
'                   Cleaned up the code for multiple strings to be passed

' ********************************************************************************************************************************************************

' USE EXAMPLE 1 (BASIC): _
    Call fx_Update_Single_Field( _
        wsSource:=wsLists, wsDest:=wsData, _
        str_Source_TargetField:="3. Updated Portfolio", str_Source_MatchField:="3. Sub-Portfolio", _
        str_Dest_TargetField:="PE Updated Portfolio", str_Dest_MatchField:="Sub-Portfolio", _
        bol_MissingLookupData_MsgBox:=True)

' USE EXAMPLE 2: _
    Call fx_Update_Single_Field( _
        wsSource:=wsDetailDash, wsDest:=wsSageworks, _
        int_SourceHeaderRow:=4, _
        str_Source_TargetField:="14 Digit Acct#", _
        str_Source_MatchField:="Full Customer #", _
        str_Dest_TargetField:="Account Number", _
        str_Dest_MatchField:="Full Customer #", _
        bol_ConvertMatchSourcetoValues:=True, _
        strWsNameLookup:="County Lookup", _
        bol_SkipDuplicates:=True, _
        bol_CloseSourceWb:=True, _
        bol_BlanksOnly:= True, _
        str_OnlyUseValue:= "Y", _
        bol_MissingLookupData_MsgBox:=True, _
        bol_MissingLookupData_UseExistingData:=True)

' USE EXAMPLE 2: _
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="1 - Collateral (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="1 - Collateral (#)", _
        str_Dest_MatchField:="Helper", _
        str_FilterField_Dest:="Line of Business", _
        str_FilterValue:="Asset Based Lending, Commercial Real Estate, Middle Market Banking, Sponsor and Specialty Finance, Wealth", _
        bol_FilterPassArray:=True)

' LEGEND MANDATORY:
'   wsSource:
'   wsDest:
'   str_Source_TargetField:
'   str_Source_MatchField:
'   str_Dest_TargetField:
'   str_Dest_MatchField:

' LEGEND OPTIONAL:
'   int_SourceHeaderRow: The header row for the Source ws, if blank will default to 1
'   bol_ConvertMatchSourcetoValues: Converts to values only from the Source ws for the Match fields
'   bol_CloseSourceWb: Closes the Source wb when the code has finished
'   bol_SkipDuplicates: Removes duplicate values so the looked up data will only be used once
'   bol_BlanksOnly: Only updates the data if the field is currently blank

'   str_OnlyUseValue: Used to allow a SINGLE value to be used
'   arry_OnlyUseMultipleValues: Used to allow MULTIPLE values to be used

'   bol_MissingLookupData_MsgBox: Outputs a message box with a list of fields that are missing from the lookup
'   bol_MissingLookupData_UseExistingData: Will use the existing data instead of the lookup value
'   strMissingLookupData_ValuetoUse: If the value isn't in the lookup, and I didn't include a blank in the lookups, will use this value instead
'   strWsNameLookup: If the value isn't in the lookup it will say where it was looking to help with troubleshooting
'   str_FilterField_Dest: Used to filter down the values to be imported on
'   str_FilterValue: Used to filter down the values to be imported on
'   bol_FilterPassArray: Allows an array of values to be passed instead of a single value for the filter

' ********************************************************************************************************************************************************

' -------------------------------
' Declare your wsSource variables
' -------------------------------
        
With wsSource
        
    ' Dim wsSource Range References
    
    Dim intHeaderRow_wsSource As Long
    
        If int_SourceHeaderRow <> 0 Then
            intHeaderRow_wsSource = int_SourceHeaderRow
        Else
            intHeaderRow_wsSource = 1
        End If
    
    Dim intLastCol_wsSource As Long
        intLastCol_wsSource = .Cells(intHeaderRow_wsSource, Columns.count).End(xlToLeft).Column
        
    ' Dim wsSource Column References
        
    Dim arryHeader_wsSource() As Variant
        arryHeader_wsSource = Application.Transpose(.Range(.Cells(intHeaderRow_wsSource, 1), .Cells(intHeaderRow_wsSource, intLastCol_wsSource)))
        
    Dim col_Source_Target As Integer
        col_Source_Target = fx_Create_Headers(str_Source_TargetField, arryHeader_wsSource)

    Dim col_Source_Match As Integer
        col_Source_Match = fx_Create_Headers(str_Source_MatchField, arryHeader_wsSource)

    ' Dim wsSource Range References
        
    Dim intLastRow_wsSource As Long
        intLastRow_wsSource = WorksheetFunction.Max( _
        .Cells(Rows.count, col_Source_Target).End(xlUp).Row, _
        .Cells(Rows.count, col_Source_Match).End(xlUp).Row)
        
        If intLastRow_wsSource = 1 Then intLastRow_wsSource = 2
        
    ' Dim wsSource Ranges
        
    Dim rng_Source_Target As Range
    Set rng_Source_Target = .Range(.Cells(1, col_Source_Target), .Cells(intLastRow_wsSource, col_Source_Target))
        
    Dim rng_Source_Match As Range
    Set rng_Source_Match = .Range(.Cells(1, col_Source_Match), .Cells(intLastRow_wsSource, col_Source_Match))
        
    If bol_ConvertMatchSourcetoValues = True Then
        rng_Source_Match.NumberFormat = "General"
        rng_Source_Match.Value = rng_Source_Match.Value
        Set rng_Source_Match = .Range(.Cells(1, col_Source_Match), .Cells(intLastRow_wsSource, col_Source_Match))
    End If
        
    ' Dim wsSource Arrays
    
    Dim arry_Source_Target() As Variant
        arry_Source_Target = Application.Transpose(rng_Source_Target)
    
    Dim arry_Source_Match() As Variant
        arry_Source_Match = Application.Transpose(rng_Source_Match)
        
End With
        
' -----------------------------
' Declare your wsDest variables
' -----------------------------
        
With wsDest
        
    ' Dim wsDest Range References
    
    Dim intLastCol_wsDest As Long
        intLastCol_wsDest = wsDest.Cells(1, Columns.count).End(xlToLeft).Column
 
    ' Dim wsDest Column References
    
    Dim arryHeader_wsDest() As Variant
        arryHeader_wsDest = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, intLastCol_wsDest)))
        
    Dim col_Dest_Target As Integer
        col_Dest_Target = fx_Create_Headers(str_Dest_TargetField, arryHeader_wsDest)

    Dim col_Dest_Match As Integer
        col_Dest_Match = fx_Create_Headers(str_Dest_MatchField, arryHeader_wsDest)
        
    Dim col_Dest_FilterField As Integer
        col_Dest_FilterField = fx_Create_Headers(str_FilterField_Dest, arryHeader_wsDest)
        If col_Dest_FilterField = 0 Then col_Dest_FilterField = 999
 
    ' Dim wsDest Range References
 
    Dim intLastRow_wsDest As Long
        intLastRow_wsDest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.count, col_Dest_Target).End(xlUp).Row, _
        wsDest.Cells(Rows.count, col_Dest_Match).End(xlUp).Row)
        
        If intLastRow_wsDest = 1 Then intLastRow_wsDest = 2
 
    ' Dim Ranges
        
    Dim rng_Dest_Target As Range
    Set rng_Dest_Target = .Range(.Cells(1, col_Dest_Target), .Cells(intLastRow_wsDest, col_Dest_Target))
        
    ' Dim Arrays
    
    Dim arry_Dest_Target() As Variant
        arry_Dest_Target = Application.Transpose(.Range(.Cells(1, col_Dest_Target), .Cells(intLastRow_wsDest, col_Dest_Target)))
    
    Dim arry_Dest_Match() As Variant
        arry_Dest_Match = Application.Transpose(.Range(.Cells(1, col_Dest_Match), .Cells(intLastRow_wsDest, col_Dest_Match)))

    Dim arry_Dest_Filter() As Variant
        arry_Dest_Filter = Application.Transpose(.Range(.Cells(1, col_Dest_FilterField), .Cells(intLastRow_wsDest, col_Dest_FilterField)))
 
End With
        
' ----------------------------
' Declare your Other variables
' ----------------------------
        
    ' Dim Dictionaries
    
    Dim dict_LookupData As Scripting.Dictionary
        Set dict_LookupData = New Scripting.Dictionary
            dict_LookupData.CompareMode = TextCompare

    Dim dict_MissingFields As Scripting.Dictionary
        Set dict_MissingFields = New Scripting.Dictionary
            dict_MissingFields.CompareMode = TextCompare

    ' Declare Loop Variables
    
    Dim i As Long
    
    Dim cntr_MissingFields As Integer
    
    Dim str_MissingValues As String
        
    Dim val As Variant
    
    ' Declare Message Variables
    
    Dim strMissingDataMessage As String
    
    Dim str_OnlyUseMultipleValues As String
        If bol_MultipleOnlyUseValues = True Then
            str_OnlyUseMultipleValues = Join(arry_OnlyUseMultipleValues, ", ")
        End If
    
    If strWsNameLookup <> "" Then
        strMissingDataMessage = strWsNameLookup
    Else
        strMissingDataMessage = "(ex. 'Collateral Lookup')"
    End If

' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------

On Error Resume Next
        
    For i = 1 To UBound(arry_Source_Target)
        If arry_Source_Target(i) <> "" And arry_Source_Match(i) <> "" Then
            If str_OnlyUseValue <> "" Then
                If arry_Source_Target(i) = str_OnlyUseValue Then ' Only import if the target matches the passed value
                    dict_LookupData.Add Key:=arry_Source_Match(i), Item:=arry_Source_Target(i)
                End If

            ElseIf bol_MultipleOnlyUseValues = True Then
                If InStr(1, str_OnlyUseMultipleValues, arry_Source_Target(i)) > 0 Then ' Only import if the target matches the passed values
                    dict_LookupData.Add Key:=arry_Source_Match(i), Item:=arry_Source_Target(i)
                End If
                
            Else
                dict_LookupData.Add Key:=arry_Source_Match(i), Item:=arry_Source_Target(i)
            End If
        End If
    Next i

On Error GoTo 0
    
' -----------------------------------------------------------------------------
' Loop through the Lookup Data to pull in the updated data for the target field
' -----------------------------------------------------------------------------

    For i = 2 To UBound(arry_Dest_Match)
        
        If str_FilterField_Dest = "" Or arry_Dest_Filter(i) = str_FilterValue Or bol_FilterPassArray = True And InStr(1, str_FilterValue, arry_Dest_Filter(i)) > 0 Then
        
            If dict_LookupData.Exists(arry_Dest_Match(i)) Then
                
                If bol_BlanksOnly = True Then
                    If arry_Dest_Target(i) = "" Or arry_Dest_Target(i) = 0 Then
                        arry_Dest_Target(i) = dict_LookupData.Item(arry_Dest_Match(i))
                    End If
                Else
                    arry_Dest_Target(i) = dict_LookupData.Item(arry_Dest_Match(i))
                End If
                
                If bol_SkipDuplicates = True Then dict_LookupData.Remove (arry_Dest_Match(i)) ' Remove so it can only be imported once
            ElseIf arry_Dest_Match(i) = Empty Then
            
                ' If I have a record for a blank in the lookups use that, or use the strMissingLookupData_ValuetoUse if that was passed, otherwise abort
                If dict_LookupData.Exists(" ") = True Then
                    arry_Dest_Target(i) = dict_LookupData.Item(" ")
                
                ElseIf IsMissing(strMissingLookupData_ValuetoUse) = False Then
                    arry_Dest_Target(i) = strMissingLookupData_ValuetoUse
                    GoTo MissingDataMsgBox
                Else
                    GoTo MissingDataMsgBox
                End If
            
            Else

MissingDataMsgBox:

                If bol_MissingLookupData_MsgBox = True Then ' Let the user know that the data is missing
                    
                    ' Load the dictionary with each of the exceptions noted
                    On Error Resume Next
                        dict_MissingFields.Add Key:=arry_Dest_Match(i), Item:=arry_Dest_Match(i)
                        cntr_MissingFields = cntr_MissingFields + 1
                    On Error GoTo 0
                    
                End If
                
                If bol_MissingLookupData_UseExistingData = True Then ' Use the existing data to fill in the blank
                    arry_Dest_Target(i) = arry_Dest_Match(i)
                End If
                
                'wsDest.Cells(i, col_Dest_Target).Interior.Color = RGB(252, 213, 180) ' Highlight the missing data (disabled 6/13/2022)
            End If
        
        End If
        
    Next i

' -------------------------------------------------------
' Create the MsgBox if bol_MissingLookupData_MsgBox = True
' -------------------------------------------------------

        ' Create the list of fields
        For Each val In dict_MissingFields
            str_MissingValues = str_MissingValues & Chr(10) & "  > " & CStr(dict_MissingFields(val))
        Next val
        
        If Left(str_MissingValues, 1) = vbLf Then
            str_MissingValues = Right(str_MissingValues, Len(str_MissingValues) - 2)
        End If
        
        ' Output the Messagebox if there were any results
        If cntr_MissingFields > 0 Then
                
            MsgBox Title:="Missing Values in Lookup", _
                Buttons:=vbOKOnly + vbExclamation, _
                Prompt:="There is a missing record in the lookup table for:" & Chr(10) _
                & "'" & str_MissingValues & "'" & Chr(10) & Chr(10) _
                & "Please review the lookups in the applicable worksheet: " & Chr(10) _
                & strMissingDataMessage & Chr(10) & Chr(10) _
                & "Once the data has been reviewed re-run this process, or manually update the data."
        
        End If

    'Output the values from the array
    rng_Dest_Target.Value2 = Application.Transpose(arry_Dest_Target)

' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bol_CloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function

Function fx_Open_Workbook(strPromptTitle As String, Optional bolCloseIfOpen As Boolean) As Workbook
             
' Purpose: This function will prompt the user for the workbook to open and returns that workbook.
' Trigger: Called Function
' Updated: 11/19/2022

' Change Log:
'       2/12/2021: Initial Creation
'       2/12/2021: Added the code to abort if the user selects cancel.
'       2/12/2021: Added the code to determine if the Workbook is already open.
'       6/16/2021: Added the code to ChDrive and ChDir
'       11/19/2022: Added the bolCloseIfOpen value

' --------------------------------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Set wbTEST = fx_Open_Workbook(strPromptTitle:="Select the current Sageworks data dump")
    
' bolCloseIfOpen: If the workbook is already open then close it, and re-open

' ***********************************************************************************************************************************
             
On Error Resume Next
    ChDrive ThisWorkbook.path
    ChDir ThisWorkbook.path '"\(Source Data)"
On Error GoTo 0
             
' -----------------
' Declare Variables
' -----------------
             
    Dim str_wbPath As String
        str_wbPath = Application.GetOpenFilename( _
        Title:=strPromptTitle, FileFilter:="Excel Workbooks (*.xls*;*.csv),*.xls*;*.csv")
             
        If str_wbPath = "False" Then
            MsgBox "No Workbook was selected, the code cannont continue."
            myPrivateMacros.DisableForEfficiencyOff
            End
        End If
        
' -----------------------------------------
' Determine if the Workbook is already open
' -----------------------------------------
        
    Dim bolAlreadyOpen As Boolean
        
     Dim str_wbName As String
         str_wbName = Right(str_wbPath, Len(str_wbPath) - InStrRev(str_wbPath, "\"))
        
    On Error Resume Next
        Dim wb As Workbook
        Set wb = Workbooks(str_wbName)
        bolAlreadyOpen = Not wb Is Nothing
    On Error GoTo 0
        
' -------------------------
' Set the Workbook variable
' -------------------------
        
    If bolAlreadyOpen = True Then
        If bolCloseIfOpen = True Then
            Workbooks(str_wbName).Close savechanges:=False ' Added on 11/19/22
            Set fx_Open_Workbook = Workbooks.Open(str_wbPath, UpdateLinks:=False, ReadOnly:=True)
        Else
            Set fx_Open_Workbook = Workbooks(str_wbName)
        End If
        
    Else
        Set fx_Open_Workbook = Workbooks.Open(str_wbPath, UpdateLinks:=False, ReadOnly:=True)
    End If

End Function
Function fx_Update_Default_Directory()

' Purpose: This function will reset the Drive and Directory to wherever ThisWorkbook is located.
' Trigger: Called Function
' Updated: 8/12/2021

' Change Log:
'       8/12/2021: Initial Creation

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
' ----------------------------------
' Set the current directory and path
' ----------------------------------
    
On Error Resume Next
    ChDrive ThisWorkbook.path
        
        If objFSO.FolderExists(ThisWorkbook.path & "\(Source Data)") = True Then
            ChDir ThisWorkbook.path & "\(Source Data)"
        Else
            ChDir ThisWorkbook.path
        End If
        
On Error GoTo 0

    ' Release the Object
    Set objFSO = Nothing

End Function

Function fx_Find_LastRow(wsTarget As Worksheet, Optional intTargetColumn As Long, Optional bolIncludeSpecialCells As Boolean, Optional bolIncludeUsedRange As Boolean) As Long

' Purpose: To output the the Last Row for the passed ws using multiple options.
' Trigger: Called
' Updated: 3/6/2022

' Use Example: _
'   intLastRow = fx_Find_LastRow(wsData)

' Use Example 2: Using all of the optional variables _
'   intLastRow = fx_Find_LastRow(wsTarget:=wsTest, intTargetColumn:=2, bolIncludeSpecialCells:=True, bolIncludeUsedRange:=True)

'bolIncludeUsedRange: If this is True then the last row of the UsedRange will be included in the Max formula
'bolIncludeSpecialCells: If this is True then the SpecialCells(xlCellTypeLastCell) row will be included in the Max formula

' Change Log:
'       11/29/2021: Initial Creation
'       3/6/2022:   Overhauled to include error handling, and the if statements to breakout the determination of the Last Row
'                   Added the fx_Find_Row code as an alternative to handle filtered data

' ****************************************************************************
  
' -----------------
' Declare Variables
' -----------------

On Error Resume Next

    Dim intLastRow_1st As Long
    If intTargetColumn <> 0 Then
        intLastRow_1st = wsTarget.Cells(wsTarget.Rows.count, intTargetColumn).End(xlUp).Row
    Else
        intLastRow_1st = wsTarget.Cells(wsTarget.Rows.count, "A").End(xlUp).Row
    End If
        
    Dim intLastRow_2nd As Long
    If bolIncludeSpecialCells = True Then
        intLastRow_2nd = wsTarget.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row
    End If
    
    Dim intLastRow_3rd As Long
    If bolIncludeUsedRange = True Then
        intLastRow_3rd = wsTarget.UsedRange.Rows(wsTarget.UsedRange.Rows.count).Row
    End If
    
    Dim intLastRow_4th As Long
        intLastRow_4th = fx_Find_Row(ws:=wsTarget, strTarget:="") - 1

    Dim intLastRow_Max As Long

On Error Resume Next

' ---------------------------------
' Determine which intLastRow to use
' ---------------------------------

    intLastRow_Max = WorksheetFunction.Max(intLastRow_1st, intLastRow_2nd, intLastRow_3rd, intLastRow_4th)
        If intLastRow_Max = 0 Or intLastRow_Max = 1 Then intLastRow_Max = 2 ' Don't pass 0 or 1
        fx_Find_LastRow = intLastRow_Max
        
End Function
Function fx_Find_Row(ws As Worksheet, strTarget As String, Optional strTargetFieldName As String, Optional strTargetCol As String) As Long

' Purpose: To find the target value in the passed column for the passed worksheet.  Replaces the Find function, to account for hidden rows.
' Trigger: Called
' Updated: 3/6/2022

' Change Log:
'       12/26/2021: Initial Creation
'       12/27/2021: Made the intLastRow more dynamic, and added the 1 to capture a blank row
'       1/19/2022:  Added the code to allow strTargetCol to be passed
'       3/6/2022:   Added Error Handling around the Dictionary to allow duplicates
'                   Updated to handle situations where strTargetCol AND strTargetFieldName are not passed

' --------------------------------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Find_Row( _
        ws:=ThisWorkbook.Sheets("Projects"), _
        strTargetFieldName:="Project", _
        strTarget:="P.343 - Migrate to Win10")

' Use Example 2: Passing the Target Field Name _
    intRowCurProject = fx_Find_Row(ws:=wsProjects, strTarget:=strProjName, strTargetFieldName:="Project")

' Use Example 3: Passing the Target Column letter reference _
    intRowCurProject = fx_Find_Row(ws:=wsProjects, strTarget:=strProjName, strTargetCol:="B")

' ********************************************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    'Declare Header Variables
    
    Dim arryHeader_Data() As Variant
        arryHeader_Data = Application.Transpose(ws.Range(ws.Cells(1, 1), ws.Cells(1, 99)))
        
    Dim col_Target As Long
        If strTargetCol <> "" Then
            col_Target = ws.Range(strTargetCol & "1").Column
        ElseIf strTargetFieldName <> "" Then
            col_Target = fx_Create_Headers(strTargetFieldName, arryHeader_Data)
        Else
            col_Target = 1
        End If
    
    'Declare Other Variables

    Dim intLastRow As Long
        intLastRow = WorksheetFunction.Max( _
        ws.Cells(ws.Rows.count, col_Target).End(xlUp).Row, _
        ws.UsedRange.Rows(ws.UsedRange.Rows.count).Row) + 1
        
    Dim arryData() As Variant
        arryData = ws.Range(ws.Cells(1, col_Target), ws.Cells(intLastRow, col_Target))

    Dim dictData As New Scripting.Dictionary
        dictData.CompareMode = TextCompare
        
    Dim i As Long
        
' -------------------
' Fill the Dictionary
' -------------------
    
On Error Resume Next
    
    For i = 1 To UBound(arryData)
        dictData.Add Key:=arryData(i, 1), Item:=i
    Next i
    
On Error GoTo 0
    
' --------------------
' Find the Current Row
' --------------------
    
    fx_Find_Row = dictData(strTarget)

End Function

Function fx_Find_LastColumn(wsTarget As Worksheet, Optional bolIncludeSpecialCells As Boolean, Optional bolIncludeUsedRange As Boolean) As Long

' Purpose: To output the the Last Column for the passed ws using multiple options.
' Trigger: Called
' Updated: 3/6/2022

' Use Example: _
'   intLastCol = fx_Find_LastColumn(wsData)

'bolIncludeUsedRange: If this is True then the last Col of the UsedRange will be included in the Max formula
'bolIncludeSpecialCells: If this is True then the SpecialCells(xlCellTypeLastCell) col will be included in the Max formula

' Change Log:
'       3/6/2022:  Initial Creation, based on fx_Find_LastCol

' ****************************************************************************
  
' -----------------
' Declare Variables
' -----------------

On Error Resume Next

    Dim intLastCol_1st As Long
        intLastCol_1st = wsTarget.Cells(1, wsTarget.Columns.count).End(xlToLeft).Column
        
    Dim intLastCol_2nd As Long
    If bolIncludeSpecialCells = True Then
        intLastCol_2nd = wsTarget.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Column
    End If
    
    Dim intLastCol_3rd As Long
    If bolIncludeUsedRange = True Then
        intLastCol_3rd = wsTarget.UsedRange.Columns(wsTarget.UsedRange.Columns.count).Column
    End If

    Dim intLastCol_Max As Long

On Error Resume Next

' ---------------------------------
' Determine which intLastCol to use
' ---------------------------------

    intLastCol_Max = WorksheetFunction.Max(intLastCol_1st, intLastCol_2nd, intLastCol_3rd)
        If intLastCol_Max = 0 Or intLastCol_Max = 1 Then intLastCol_Max = 2 ' Don't pass 0 or 1
        fx_Find_LastColumn = intLastCol_Max
        
End Function

Function fx_Show_or_Hide_Fields(strWsTarget As String, strFilterToUse As String)

' Purpose: This function shows / hides fields, based on the passed Filter, and which fields have strikethrough formatting in wsLists.
' Trigger: Called Function
' Updated: 3/25/2022

' Use Example: _
    Call fx_Show_or_Hide_Fields(strWsTarget:="Dashboard Review", strFilterToUse:="Quarterly - CRE")

' Change Log:
'       3/25/2022:  Initial Creation

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    ' Declare Worksheets
    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(strWsTarget)
    
    Dim wsLists As Worksheet
    Set wsLists = ThisWorkbook.Sheets("LISTS")
    
    ' Declare Loop Variables
    Dim col_wsTarget As Long
    Dim row_FilterFields As Long
        
    ' Declare Strings
    Dim strCurTargetField As String
    Dim strCurFilterField As String

    ' Declare wsTarget Cell References
    Dim intLastCol_wsTarget As Long
        intLastCol_wsTarget = wsTarget.Cells(1, Columns.count).End(xlToLeft).Column
    
    ' Declare wsLists Cell References
    Dim intLastCol_wsLists As Long
        intLastCol_wsLists = wsLists.Cells(1, Columns.count).End(xlToLeft).Column
    
    Dim arryHeader_wsLists() As Variant
        arryHeader_wsLists = Application.Transpose(wsLists.Range(wsLists.Cells(1, 1), wsLists.Cells(1, intLastCol_wsLists)))
    
    Dim col_FilterFlags As Long
        col_FilterFlags = fx_Create_Headers(strFilterToUse, arryHeader_wsLists)

    Dim intLastRow_wsLists As Long
        intLastRow_wsLists = wsLists.Cells(Rows.count, col_FilterFlags).End(xlUp).Row

' -----------------------------------------------------------------------
' Filter the data based on which fields have the strikethrough formatting
' -----------------------------------------------------------------------

    For col_wsTarget = 1 To intLastCol_wsTarget
        strCurTargetField = wsTarget.Cells(1, col_wsTarget)
    
        For row_FilterFields = 2 To intLastRow_wsLists
            strCurFilterField = wsLists.Cells(row_FilterFields, col_FilterFlags)
            
            If strCurTargetField = strCurFilterField Then ' If we match
                If wsLists.Cells(row_FilterFields, col_FilterFlags).Font.Strikethrough = False Then
                    wsTarget.Columns(col_wsTarget).Hidden = False
                    '.Cells(row_FilterFields, col_FilterFlags).EntireColumn.Hidden = False
                Else
                    wsTarget.Columns(col_wsTarget).Hidden = True
                    'wsTarget.Cells(row_FilterFields, col_FilterFlags).EntireColumn.Hidden = True
                End If
                
            End If
    
        Next row_FilterFields
    
    Next col_wsTarget

End Function
Function fx_Sheet_Exists(strWbName As String, strWsName As String, Optional bolDeleteSheet As Boolean) As Boolean

' Purpose: To determine if a sheet exists, to be used in an IF statement.
' Trigger: Called
' Updated: 7/26/2022

' Change Log:
'       6/29/2021:  Intial Creation
'       6/29/2021:  Added the ErrorHandler for the 2015 #VALUE error when the ws doesn't exist
'       7/26/2022:  Added the optional 'bolDeleteSheet' and application.displayalert = false

' --------------------------------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
'    Dim bolSupportOpenAlready As Boolean
'    bolSupportOpenAlready = fx_Sheet_Exists( _
        strWbName:=strProjName & ".xlsx", _
        strWsName:="Next Actions")
'   If fx_Sheet_Exists(ThisWorkbook.Name, "VALIDATION") = False Then

' LEGEND:
'   bolDeleteSheet: If the sheet exists delete it and then pass False

' ********************************************************************************************************************************************************

On Error GoTo ErrorHandler

Application.DisplayAlerts = False
    If Evaluate("ISREF('[" & strWbName & "]" & strWsName & "'!A1)") = True Then
        If bolDeleteSheet = True Then
            Workbooks(strWbName).Sheets(strWsName).Delete
        Else
            fx_Sheet_Exists = True
        End If
    End If
Application.DisplayAlerts = True
        
    Exit Function

ErrorHandler:

fx_Sheet_Exists = False

End Function
Sub XXX_Test_fx_Sheet_Exists()

    Call fx_Sheet_Exists(ThisWorkbook.Name, "Complete Dashboard", True)
        ThisWorkbook.Sheets("Dashboard Review").Copy After:=ThisWorkbook.Sheets("Dashboard Review")
        ActiveSheet.Name = "Complete Dashboard"


End Sub
Function fx_Create_Dynamic_Lookup_List(wsDataSource As Worksheet, str_Dynamic_Lookup_Value As String, col_Dynamic_Lookup_Field As Long, Optional col_Criteria_Field As Long, Optional str_Criteria_Match_Value As Variant, Optional col_Target_Field As Long) As Variant

' Purpose: To create the dynamic list of values to be used in the ListBox, based on a change to the cmb_Dynamic_Borrower_Lookup.
' Trigger: Start typing in the Dynamic_Borrower_Lookup combo box (cmb_Dynamic_Borrower_Lookup_Change)
' Updated: 10/7/2022

' Change Log:
'       11/21/2021: Intial Creation for the PAR Agenda, taken from Sageworks Validation Dashboard code
'       1/18/2022:  Updated and converted to a function
'       10/6/2022:  Updated the naming of the fields, and handled the error if there was no LOB Lookup value
'       10/7/2022:  Updated to allow a seperate Dynamic Lookup and Target field

' ********************************************************************************************************************************************************

' Use Example: _
'    Dim arryBorrowersTemp As Variant
'       arryBorrowersTemp = myFunctions.fx_Create_Dynamic_Lookup_List( _
        wsDataSource:=wsData, _
        col_Dynamic_Lookup_Field:=col_Borrower, _
        str_Dynamic_Lookup_Value:=Me.cmb_Dynamic_Borrower.Value, _
        col_Criteria_Field:=col_LOBUpdated, _
        str_Criteria_Match_Value:=Me.lst_LOB.Value)

'    Me.lst_Borrowers.List = arryTargetValuesTemp

' ********************************************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Declare Strings
    Dim strLookupValue As String
    Dim strCriteriaValue As String
    Dim strTargetValue As String
    
    'Declare Cell References
    Dim intSourceRow As Long: intSourceRow = 2
    Dim intArryRow As Long: intArryRow = 0
    
    'Declare Arrays
    Dim arryTargetValues As Variant
        ReDim arryTargetValues(1 To 99999)

    If IsMissing(str_Criteria_Match_Value) Then str_Criteria_Match_Value = ""

' -------------------------------------
' Add the borrowers to the lookup Array
' -------------------------------------
       
    With wsDataSource
            
        Do While .Cells(intSourceRow, col_Dynamic_Lookup_Field).Value2 <> ""
            
            ' Set the loop variables
            strLookupValue = .Cells(intSourceRow, col_Dynamic_Lookup_Field).Value2
            If col_Criteria_Field <> 0 Then strCriteriaValue = .Cells(intSourceRow, col_Criteria_Field).Value2
            If col_Target_Field <> 0 Then strTargetValue = .Cells(intSourceRow, col_Target_Field).Value2
                
                'If the data matches add to the array
                If InStr(1, strLookupValue, str_Dynamic_Lookup_Value, vbTextCompare) Then
                    If col_Criteria_Field = 0 Or str_Criteria_Match_Value = strCriteriaValue Then
                        intArryRow = intArryRow + 1
                        If strTargetValue <> "" Then arryTargetValues(intArryRow) = strTargetValue Else arryTargetValues(intArryRow) = strLookupValue
                    End If
                End If
            
            intSourceRow = intSourceRow + 1
        Loop
    End With

    If intArryRow > 0 Then ' If nothing was passed, don't redim
        ReDim Preserve arryTargetValues(1 To intArryRow)
    End If
    
    'Output the results
    fx_Create_Dynamic_Lookup_List = arryTargetValues

End Function
Function fx_Copy_in_Data_for_Matching_Fields(wsSource As Worksheet, wsDest As Worksheet, Optional intSourceHeaderRow As Long, Optional intLastRowtoImport As Long, Optional strModuleName As String, Optional strControlTotalField As String, Optional bolCloseSourceWb As Boolean, Optional bolVisibleFieldsOnly As Boolean, Optional intCurRow_wsValidation As Long)

' Purpose: To copy the data from the source to the destination, wherever the fields match.
' Trigger: Called
' Updated: 3/1/2022
'
' Change Log:
'       9/18/2020:  Intial Creation based on CV Mod Agg. CV Tracker import code
'       11/3/2020:  Updated to include the strSourceDesc and strDestDesc to feed into the validation function
'       11/3/2020:  Updated to include the strModuleName to feed into the validation function
'       11/3/2020:  Removed the 'DisableforEfficiency' as it was disabiling it in my Main Procedure.
'       2/12/2021:  Updated to account for pulling in only visible data
'       2/12/2021:  Switched from the filtered boolean to intLastRowtoImport
'       2/12/2021:  Updated to use Arrays instead of Ranges for the import
'       2/12/2021:  Added the code related to bolCloseSourceWb
'       2/25/2021:  Updated the code for strControlTotalField and strModuleName to make it optional
'       2/25/2021:  Removed the old col_Bal_Dest and col_Bal_Source references
'       3/9/2021:   Added the code to use intLastUsedRow_Dest to delete any extraneous rows
'       3/15/2021:  Updated to include the Optional intHeaderRow field to handle ignoring headers
'       5/17/2021:  Updated the code related to the Data Validation to use the intFirstRowData_Source instead of defaulting to 1
'       5/17/2021:  Updated the intHeaderRow variable to use the code from intHeaderRow_Source
'       6/16/2021:  Made some minor improvements to the variables to make them more resilient.
'       6/16/2021:  Added the code to apply the formatting from the first row to the rest.
'       6/21/2021:  Added Option to only import the visible fields
'       6/22/2021:  Updated the code to assign intHeaderRow_Source, and related code in the data
'       3/1/2022:   Added the intCurRow_wsValidation variable to pass to fx_Create_Data_Validation_Control_Totals

' --------------------------------------------------------------------------------------------------------------------------------------------------------

'   Use Example: _
        Call fx_Copy_in_Data_for_Matching_Fields( _
            wsSource:=wsSource, _
            wsDest:=wsData, _
            intSourceHeaderRow:=1, _
            intLastRowtoImport:=0, _
            strModuleName:="o_11_Import_Data", _
            strControlTotalField:="New Direct Outstanding", _
            bolCloseSourceWb:=True, _
            intCurRow_wsValidation:=2, _
            bolVisibleFieldsOnly:=True)

'        Call fx_Copy_in_Data_for_Matching_Fields_v2( _
            wsSource:=wsSource, _
            wsDest:=wsData, _
            strModuleName:="o_11_Import_Sageworks_Customer_Data", _
            strControlTotalField:="Webster Outstanding (000's) - Book Balance", _
            bolCloseSourceWb:=False, _
            intCurRow_wsValidation:=2)

' ***********************************************************************************************************************************

' -----------------------------------------------
' Turn off any filtering from the source and dest
' -----------------------------------------------
        
    If wsSource.AutoFilterMode = True Then wsSource.AutoFilter.ShowAllData
        
    If wsDest.AutoFilterMode = True Then wsDest.AutoFilter.ShowAllData

' -----------------
' Declare Variables
' -----------------

    'Dim "Source" Integers
    
    Dim intLastRow_Source As Long
        If intLastRowtoImport > 0 Then 'If I passed the intLastRowtoImport variable use it
            intLastRow_Source = intLastRowtoImport
        Else
            intLastRow_Source = WorksheetFunction.Max( _
            wsSource.Cells(Rows.count, "A").End(xlUp).Row, _
            wsSource.Cells(Rows.count, "B").End(xlUp).Row, _
            wsSource.Cells(Rows.count, "C").End(xlUp).Row)
        End If

    Dim intHeaderRow_Source As Long
        
        If intSourceHeaderRow > 0 Then 'If I passed the intHeaderRow variable use it
            intHeaderRow_Source = intSourceHeaderRow
        Else
            intHeaderRow_Source = 1
        End If

    Dim intFirstRowData_Source As Long
        intFirstRowData_Source = intHeaderRow_Source + 1

    Dim intLastCol_Source As Long
        intLastCol_Source = WorksheetFunction.Max( _
        wsSource.Cells(intHeaderRow_Source, Columns.count).End(xlToLeft).Column, _
        wsSource.Rows(intHeaderRow_Source).Find("").Column - 1)
        
    'Dim "Dest" Integers

    Dim intLastRow_Dest As Long
        intLastRow_Dest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.count, "A").End(xlUp).Row, _
        wsDest.Cells(Rows.count, "B").End(xlUp).Row, _
        wsDest.Cells(Rows.count, "C").End(xlUp).Row)
        
        If intLastRow_Dest = 1 Then intLastRow_Dest = 2
    
    Dim intLastUsedRow_Dest As Long
        intLastUsedRow_Dest = wsDest.Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
    Dim intLastCol_Dest As Long
        intLastCol_Dest = wsDest.Cells(1, Columns.count).End(xlToLeft).Column
        intLastCol_Dest = WorksheetFunction.Max( _
        wsDest.Cells(1, Columns.count).End(xlToLeft).Column, _
        wsDest.Rows(1).Find("").Column - 1)
    
    'Dim Other Integers
        
    Dim intCurRowValidation As Long
        
    'Dim Ranges / "Ranges"
    
    Dim arryHeader_Source() As Variant
        arryHeader_Source = Application.Transpose( _
        wsSource.Range(wsSource.Cells(intHeaderRow_Source, 1), wsSource.Cells(intHeaderRow_Source, intLastCol_Source)))
        
    Dim arryHeader_Dest() As Variant
        arryHeader_Dest = Application.Transpose( _
        wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(1, intLastCol_Dest)))
        
    Dim arrySourceData() As Variant
    
    'Dim arryDestData() As Variant
    
    'Dim Values for Loops
        
    Dim strFieldName As String
    
    Dim intColNum_Source As Long
    
    Dim intColNum_Dest As Long
        
    Dim i As Long
    
    'Dim Strings
    
    Dim strSourceDesc As String
        strSourceDesc = wsSource.Parent.Name & " - " & wsSource.Name

    Dim strDestDesc As String
        strDestDesc = wsDest.Parent.Name & " - " & wsDest.Name

' ----------------------------------
' Copy over the data from the source
' ----------------------------------
        
    'Clear out the old data and cell fill
    wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(intLastRow_Dest, intLastCol_Dest)).ClearContents
    wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(intLastRow_Dest, intLastCol_Dest)).Interior.Color = xlNone
    wsDest.Range(wsDest.Cells(intLastRow_Dest + 1, 1), wsDest.Cells(intLastUsedRow_Dest, 1)).EntireRow.Delete

    'Loop through the fields in the destination
    For intColNum_Dest = 1 To intLastCol_Dest
        strFieldName = wsDest.Cells(1, intColNum_Dest).Value2
        intColNum_Source = fx_Create_Headers(strFieldName, arryHeader_Source)
        If intColNum_Source <> 0 Then
            arrySourceData = Application.Transpose(wsSource.Range(wsSource.Cells(intFirstRowData_Source, intColNum_Source), wsSource.Cells(intLastRow_Source, intColNum_Source)))
        End If
        
        If intColNum_Source > 0 Then
            If bolVisibleFieldsOnly = True And wsDest.Columns(intColNum_Dest).Hidden = False Then
                wsDest.Range(wsDest.Cells(2, intColNum_Dest), wsDest.Cells(intLastRow_Source - intHeaderRow_Source + 1, intColNum_Dest)).Value2 = _
                wsSource.Range(wsSource.Cells(intFirstRowData_Source, intColNum_Source), wsSource.Cells(intLastRow_Source, intColNum_Source)).Value2
            Else
                On Error Resume Next
                wsDest.Range(wsDest.Cells(2, intColNum_Dest), wsDest.Cells(intLastRow_Source - intHeaderRow_Source + 1, intColNum_Dest)).Value2 = _
                wsSource.Range(wsSource.Cells(intFirstRowData_Source, intColNum_Source), wsSource.Cells(intLastRow_Source, intColNum_Source)).Value2
                On Error GoTo 0
                

            End If
        End If
        
    Next intColNum_Dest

    'Reset the Last Row variable
    intLastRow_Dest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.count, "A").End(xlUp).Row, _
        wsDest.Cells(Rows.count, "B").End(xlUp).Row, _
        wsDest.Cells(Rows.count, "C").End(xlUp).Row)

' ------------------------------------------------------------
' Output the control totals to the Validation ws, if it exists
' ------------------------------------------------------------
    
If strControlTotalField <> "" Then
    
    Call fx_Create_Data_Validation_Control_Totals( _
        wsDataSource:=wsSource, _
        strModuleName:=strModuleName, _
        strSourceName:=strSourceDesc, _
        intHeaderRow:=intHeaderRow_Source, _
        intLastRowtoImport:=intLastRowtoImport, _
        strControlTotalField:=strControlTotalField, _
        intCurRow_wsValidation:=intCurRow_wsValidation)
    
    Call fx_Create_Data_Validation_Control_Totals( _
        wsDataSource:=wsDest, _
        strModuleName:=strModuleName, _
        strSourceName:=strDestDesc, _
        intHeaderRow:=1, _
        strControlTotalField:=strControlTotalField, _
        intCurRow_wsValidation:=intCurRow_wsValidation + 1)
    
End If
    
' -----------------------------------------------
' Apply the formatting to all rows from the first
' -----------------------------------------------
    
    Call fx_Steal_First_Row_Formating( _
        ws:=wsDest, _
        intFirstRow:=2, _
        intLastRow:=intLastRow_Dest, _
        intLastCol:=intLastCol_Dest)
    
' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bolCloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function


Function fx_Create_Data_Validation_Control_Totals(wsDataSource As Worksheet, strModuleName As String, strSourceName As String, intHeaderRow As Long, strControlTotalField As String, Optional intLastRowtoImport As Long, Optional intCurRow_wsValidation As Long, Optional dblTotalsFromSource As Double, Optional intRecordCountFromSource As Long)

' Purpose: To output the data validation control totals to the wsValidation, if it exists.
' Trigger: Called
' Updated: 1/5/2022


' Change Log:
'       9/26/2020: Initial Creation
'       11/3/2020: Updated to activate ThisWorkbook before checking for the Validation ws
'       12/19/2020: Made the intRecordCount more resiliant
'       12/19/2020: Added the ThisWorkbook.Name to the ISREF check
'       2/12/2021: Added the code for intLastRowtoImport
'       5/17/2021: Updated the calculation for strRng_Totals to go down to intRecordCount + intHeaderRow
'       6/16/2021: Added the code related to intCurRow_wsValidation
'       6/22/2021: Updated the code for intRecordCount
'       6/22/2021: Updated the Check for the Validation worksheet to reference ThisWorkbook, and avoid the .Activate
'       1/5/2022:  Added the code to bypass the sums and whatnot if dblTotalsFromSource or intRecordCountFromSource is passed
    
' --------------------------------------------------------------------------------------------------------------------------------------------------------
    
' Use Example: _
    Call fx_Create_Data_Validation_Control_Totals( _
        wsDataSource:=wsDest, _
        strModuleName:=strModuleName, _
        strSourceName:=strDestDesc, _
        intHeaderRow:=1, _
        strControlTotalField:=strControlTotalField, _
        intCurRow_wsValidation:=3)

' Use Example 2: _
    Call fx_Create_Data_Validation_Control_Totals( _
    wsDataSource:=ws3666_Source, _
    strModuleName:="o3_Import_BB_Data", _
    strSourceName:=ws3666_Source.Parent.Name & " - " & ws3666_Source.Name, _
    intHeaderRow:=intHeaderRow, _
    strControlTotalField:="Line Commitment", _
    intCurRow_wsValidation:=10, _
    intRecordCountFromSource:=intDestRowCounter - 2, _
    dblTotalsFromSource:=dblControlTotal)
    
' ***********************************************************************************************************************************

    'Only run of the VALIDATION ws exists
    If Evaluate("ISREF(" & "'[" & ThisWorkbook.Name & "]" & "Validation'" & "!A1)") = False Then
        Debug.Print "fx_Create_Data_Validation_Control_Totals failed becuase there is no ws called 'VALIDATION' in the Workbook"
        Exit Function
    End If

' ----------------------------
' Declare Validation Variables
' ----------------------------

    'Dim Worksheets

    Dim wsValidation As Worksheet
    Set wsValidation = ThisWorkbook.Sheets("VALIDATION")

    Dim wsSource As Worksheet
    Set wsSource = wsDataSource

    'Dim Cell References

    Dim intLastCol As Long
        intLastCol = wsSource.Cells(intHeaderRow, Columns.count).End(xlToLeft).Column
      
    Dim intCurRow As Long
        If intCurRow_wsValidation > 1 Then 'If I passed the intCurRow_wsValidation variable use it
            intCurRow = intCurRow_wsValidation
        Else
            intCurRow = wsValidation.Cells(Rows.count, "A").End(xlUp).Row + 1
        End If

    'Bypass the code if dblTotals was passed
    If dblTotalsFromSource > 0 Or intRecordCountFromSource > 0 Then GoTo Bypass

' ------------------------
' Declare Source Variables
' ------------------------

    'Dim "Ranges"
    
    Dim arryHeader() As Variant
        arryHeader = Application.Transpose(wsSource.Range(wsSource.Cells(intHeaderRow, 1), wsSource.Cells(intHeaderRow, intLastCol)))
        
    Dim intColTotals As Long
        intColTotals = fx_Create_Headers(strControlTotalField, arryHeader)
    
    'Dim Integers

    Dim intRecordCount As Long
        If intLastRowtoImport > 0 Then 'If I passed the intLastRowtoImport variable use it
            intRecordCount = intLastRowtoImport - intHeaderRow
        Else
            intRecordCount = WorksheetFunction.Max( _
            wsSource.Cells(Rows.count, "A").End(xlUp).Row, _
            wsSource.Cells(Rows.count, "B").End(xlUp).Row, _
            wsSource.Cells(Rows.count, "C").End(xlUp).Row) - intHeaderRow
        End If
    
    'Dim Other Variables

    Dim strCol_Totals As String
        strCol_Totals = Split(Cells(1, intColTotals).Address, "$")(1)
    
    Dim strRng_Totals As String
        strRng_Totals = strCol_Totals & "1:" & strCol_Totals & intRecordCount + intHeaderRow
        
    Dim dblTotals As Double
        dblTotals = Round(Application.WorksheetFunction.Sum(wsSource.Range(strRng_Totals)), 2)

Bypass:

    'Assign the Optional variables if they were passed
    If dblTotalsFromSource > 0 Then dblTotals = dblTotalsFromSource
    If intRecordCountFromSource > 0 Then intRecordCount = intRecordCountFromSource

' ------------------------------------------------------
' Output the validation totals from the passed variables
' ------------------------------------------------------

    With wsValidation
        .Range("A" & intCurRow) = Format(Now, "m/d/yyyy hh:mm")   'Date / Time
        .Range("B" & intCurRow) = strModuleName                   'Code Module
        .Range("C" & intCurRow) = strSourceName                   'Source
        .Range("D" & intCurRow) = Format(dblTotals, "$#,##0")     'Total
        .Range("E" & intCurRow) = Format(intRecordCount, "0,0")   'Count
    End With

End Function
Function fx_Steal_First_Row_Formating(ws As Worksheet, Optional intFirstRow As Long, Optional intLastCol As Long, Optional intLastRow As Long, Optional intSingleRow As Long)

' Purpose: To copy the formatting from the first row of data and apply to the rest of the data.
' Trigger: Called
' Updated: 1/19/2022

' Use Example: _
    Call fx_Steal_First_Row_Formating( _
        ws:=wsQCReview, _
        intFirstRow:=2, _
        intLastRow:=intLastRow, _
        intLastCol:=intLastCol)

' Use Example 2: _
    Call fx_Steal_First_Row_Formating( _
        ws:=wsData, _
        intSingleRow:=CurRow)

' Use Example 3: Call fx_Steal_First_Row_Formating(ws:=wsQCReview)

' Change Log:
'       5/17/2021:  Intial Creation
'       6/16/2021:  Added the 'Application.Goto' to reset the copy paste
'       12/6/2021:  Added the option to pass only a single row
'       12/8/2021:  Added the rngCur so the screen doesn't jump around
'       1/8/2022:   Updated some of the passed variables to be optional
'                   Defaulted intFirstRow to be 2 if not passed
'       1/19/2022:  Added the ws. qualifiers for LastRow and LastCol

' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    Dim rngCur As Range
    Set rngCur = ActiveCell

    'Declare Integers
    
    If intFirstRow = 0 Then intFirstRow = 2
    
    If intLastRow = 0 Then
       intLastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    End If
    
    If intLastCol = 0 Then
       intLastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    End If
    
    'Declare Ranges
    
With ws

    Dim rngFormat As Range
        Set rngFormat = .Range(.Cells(intFirstRow, 1), .Cells(intFirstRow, intLastCol))
    
    Dim rngTarget As Range
        If intSingleRow <> 0 Then
            Set rngTarget = .Range(.Cells(intSingleRow, 1), .Cells(intSingleRow, intLastCol))
        ElseIf intLastRow <> 0 Then
            Set rngTarget = .Range(.Cells(intFirstRow + 1, 1), .Cells(intLastRow, intLastCol))
        Else
            MsgBox "There was no row passed to the Steal First Row function."
        End If

End With

' ---------------------------------------------------------------------------------------------------
' Copy the formatting from the first row of data (intFirstRow) to the remaining rows (thru intLastRow)
' ---------------------------------------------------------------------------------------------------
    
    rngFormat.Copy: rngTarget.PasteSpecial xlPasteFormats
    
    'Go back to where you were before the code
    Application.CutCopyMode = False
    Application.GoTo Reference:=rngCur, Scroll:=False
    
End Function
Function fx_Delete_Unused_Data(ws As Worksheet, str_Target_Field As String, str_Value_To_Delete As String, _
Optional bol_DeleteDataOnly As Boolean, Optional bol_DeleteValues_PassArray As Boolean)

' Purpose: To delete data from the passed worksheet where the "Value To Delete" is in the Target Field.
' Trigger: Called
' Updated: 6/29/2023

' Use Example: _
    Call fx_Delete_Unused_Data( _
        ws:=wsSageworksRT_Dest, _
        str_Target_Field:="Line of Business", _
        str_Value_To_Delete:="Small Business")

' LEGEND OPTIONAL:
'   bol_DeleteDataOnly: Allows only the specific fields to be deleted, otherwise the entire row is deleted
'   bol_FilterPassArray: Allows an array of values to be passed instead of a single value for the filter

' Change Log:
'       9/15/2021:  Initial Creation
'       5/5/2023:   Added the 'Exit Function' if the intLastRow is 1, to handle situations where there is no data (ex. Policy Exceptions - Paid Off Loans)
'                   Updated the delete to keep the first row of data for the formatting
'       6/12/2023:  Added the option to pass an array of values to be deleted, and the code around arryValuesToDelete
'       6/29/2023:  Added the option to delete just the values, not the rows

' ****************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Cell References
       
    Dim intLastRow As Long
        intLastRow = WorksheetFunction.Max( _
            ws.Cells(Rows.count, "A").End(xlUp).Row, _
            ws.Cells(Rows.count, "B").End(xlUp).Row, _
            ws.Cells(Rows.count, "C").End(xlUp).Row)
       
       If intLastRow = 1 Then Exit Function
       
    Dim intLastCol As Integer
        intLastCol = ws.Cells(1, Columns.count).End(xlToLeft).Column

    ' Declare "Ranges"
    
    Dim arryHeader() As Variant
        arryHeader = Application.Transpose(ws.Range(ws.Cells(1, 1), ws.Cells(1, intLastCol)))

    Dim col_Target As Integer
        col_Target = fx_Create_Headers(str_Target_Field, arryHeader)
        
    ' Declare Arrays
    
    Dim arryValuesToDelete() As String
    
    If bol_DeleteValues_PassArray = True Then
        arryValuesToDelete = Split(str_Value_To_Delete, ", ")
    End If
    
' ----------------------
' Delete the Unused Data
' ----------------------
        
On Error Resume Next
        
With ws
    
    ' Sort the data to make deleting MUCH faster
    .Range(.Cells(1, 1), .Cells(intLastRow, intLastCol)).Sort _
        Key1:=.Cells(1, col_Target), Order1:=xlAscending, Header:=xlYes
    
    If bol_DeleteDataOnly = True Then
        
        If bol_DeleteValues_PassArray = False Then
        
            .Range("A1").AutoFilter Field:=col_Target, Criteria1:=str_Value_To_Delete, Operator:=xlFilterValues
                .Range(.Cells(2, col_Target), .Cells(intLastRow, col_Target)).SpecialCells(xlCellTypeVisible).ClearContents
                .Range("A1").AutoFilter Field:=col_Target
        Else
            .Range("A1").AutoFilter Field:=col_Target, Criteria1:=arryValuesToDelete, Operator:=xlFilterValues
                .Range(.Cells(2, col_Target), .Cells(intLastRow, col_Target)).SpecialCells(xlCellTypeVisible).ClearContents
                .Range("A1").AutoFilter Field:=col_Target
        
        End If
        
        Exit Function ' Skip the deleting the rows
    
    End If
    
    ' Filter based on the values and then delete the filtered rows
    If bol_DeleteValues_PassArray = False Then
    
        .Range("A1").AutoFilter Field:=col_Target, Criteria1:=str_Value_To_Delete, Operator:=xlFilterValues
            .Range("A2:A3").SpecialCells(xlCellTypeVisible).EntireRow.ClearContents
            .Range("A3:A" & intLastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .Range("A1").AutoFilter Field:=col_Target
    Else
        .Range("A1").AutoFilter Field:=col_Target, Criteria1:=arryValuesToDelete, Operator:=xlFilterValues
            .Range("A2:A3").SpecialCells(xlCellTypeVisible).EntireRow.ClearContents
            .Range("A3:A" & intLastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .Range("A1").AutoFilter Field:=col_Target
    
    End If
    
End With

On Error GoTo 0

End Function
Function fx_Convert_Values_In_Range(wsData As Worksheet, strTargetField As String, strNewValue As String, _
Optional strLookupField As String)

' Purpose: To convert values from their original value to a new value in the passed range.

' Trigger: Called
' Updated: 6/29/2023

' Change Log:
'       6/29/2023:  Initial creation

' ********************************************************************************************************************************************************

' USE EXAMPLE: _
    Call fx_Convert_Values_In_Range( _
        wsData:=wsData, _
        strLookupField:="Borrower", _
        strTargetField:="CRE Flag", _
        strNewValue:="Yes")

' LEGEND MANDATORY:
'   wsData: The worksheet being referenced

' ********************************************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
        
With wsData
        
    ' Declare Header Data
        
    Dim intLastCol As Long
        intLastCol = .Cells(1, Columns.count).End(xlToLeft).Column
        
    Dim arryHeader_wsData() As Variant
        arryHeader_wsData = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, intLastCol)))
        
    Dim col_TargetField As Integer
        col_TargetField = fx_Create_Headers(strTargetField, arryHeader_wsData)
        
    Dim intLastRow As Long
        intLastRow = .Cells(Rows.count, col_TargetField).End(xlUp).Row

    ' Declare Ranges / Arrays
    
    Dim rngTargetData As Range
    Set rngTargetData = .Range(.Cells(2, col_TargetField), .Cells(intLastRow, col_TargetField))
        
    Dim arryTargetData() As Variant
        arryTargetData = Application.Transpose(rngTargetData)
        
    ' Declare LookupField Variables
If strLookupField <> "" Then
        
    Dim col_LookupField As Integer
        col_LookupField = fx_Create_Headers(strLookupField, arryHeader_wsData)
        
    Dim rngLookupData As Range
    Set rngLookupData = .Range(.Cells(2, col_LookupField), .Cells(intLastRow, col_LookupField))
        
    Dim arryLookupData() As Variant
        arryLookupData = Application.Transpose(rngLookupData)
                
End If
        
End With
                        
' ----------------------------
' Declare your Other variables
' ----------------------------
        
    ' Declare Loop Variables
    
    Dim i As Long
    
' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------

On Error Resume Next
        
    For i = 1 To UBound(arryTargetData)

        If arryTargetData(i) <> "" Then
            arryTargetData(i) = strNewValue
        End If

    Next i

On Error GoTo 0

' ------------------------------------
' Output the results back to the field
' ------------------------------------

    rngTargetData.Value2 = Application.Transpose(arryTargetData)

End Function
Function fx_Update_Named_Range(strNamedRangeName As String)

' Purpose: To update the passed Named Range a change in the Change Log.
' Trigger: Called
' Updated: 3/6/2022

' Use Example: _
    Call fx_Update_Named_Range("ChangeLog_Data")

' Change Log:
'       12/8/2021:  Intial Creation
'       3/4/2022:   Added Error Handling for the intLastRow and intLastCol to handle if all of the empty rows/cols are hidden
'       3/6/2022:   Replaced the intLastRow and intLastCol w/ functions

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strWsName As String
        strWsName = ThisWorkbook.Names(strNamedRangeName).RefersToRange.Parent.Name
    
    Dim wsNamedRange As Worksheet
    Set wsNamedRange = ThisWorkbook.Sheets(strWsName)
    
    Dim intLastRow As Long
        intLastRow = fx_Find_LastRow(wsNamedRange)
        
    Dim intLastCol As Integer
        intLastCol = fx_Find_LastColumn(wsNamedRange)

' ----------------------
' Update the Named Range
' ----------------------

   ThisWorkbook.Names(strNamedRangeName).RefersToR1C1 = wsNamedRange.Range(wsNamedRange.Cells(1, 1), wsNamedRange.Cells(intLastRow, intLastCol))

End Function
