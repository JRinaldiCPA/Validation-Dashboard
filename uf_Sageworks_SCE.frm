VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Sageworks_SCE 
   Caption         =   "Customer Selector UserForm -  Privileged User Version"
   ClientHeight    =   5784
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   18708
   OleObjectBlob   =   "uf_Sageworks_SCE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Sageworks_SCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Declare Worksheets
    Dim wsData As Worksheet
    Dim wsDetailData As Worksheet
    
    Dim wsArrays As Worksheet
    Dim wsPivot As Worksheet

' Declare Strings
    Dim strCustomer As String
    Dim strLastCol_wsData As String

' Declare Integers
    Dim intLastRow As Long
    Dim intLastRow_wsArrays As Long
    
    Dim intLastCol As Long
    
' Declare Data "Ranges"
    Dim col_Borrower As Long
    Dim col_CustID As Long
    Dim col_LOB As Long
    Dim col_LOBUpdated As Long
    Dim col_Region As Long
    Dim col_Team As Long
    Dim col_PM_Name As Long
    Dim col_AM_Name As Long
    
    Dim col_BRG As Long
    Dim col_FRG As Long
    Dim col_CCRP As Long
    
    Dim col_Exposure As Long
    Dim col_Outstanding As Long
    
    Dim col_PM_Attest As Long
    Dim col_PM_Attest_Exp As Long
    Dim col_CovCompl As Long
    Dim col_CovCompl_Exp As Long
    Dim col_ChangeFlag As Long
    
' Declare Arrays / Other
    Dim bolPrivilegedUser As Boolean
    
    Dim arryHeader_Data() As Variant
    Dim arryHeader_Lists() As Variant

    Dim ary_Customers
    Dim ary_SelectedCustomers
    Dim ary_PM
    
    Dim ary_BorrowerLookupData

' Declare Dictionaries
    Dim dict_PMs As Scripting.Dictionary

' Declare "Booleans"
    Dim bol_AttestationStatus As String
    Dim bol_Edit_Filter As Boolean
    Dim bol_QC_Flags As Boolean
    Dim bol_CovCompliance As String
    Dim bol_wsDetailData_Exists As Boolean
    
Option Explicit
Private Sub UserForm_Initialize()
 
' Purpose:  To initialize the userform, including adding in the data from the arrays.
' Trigger:  Workbook Open
' Updated:  11/16/2020
' Author:   James Rinaldi

' Change Log:
'       3/23/2020: Intial Creation
'       8/19/2020: Added the logic to exclude the exempt customers
'       11/16/2020: I added the autofilter to hide the CRE and ABL LOBs
'       12/29/2020: Moved the CRE and ABL filtering to o_64_Update_Workbook_for_Admins

' ****************************************************************************

Call Me.o_02_Assign_Global_Variables

Call Me.o_03_Declare_Global_Arrays

    wsData.Unprotect

' -----------
' Initialize the initial values
' -----------
    Me.StartUpPosition = 0 'Allow you to set the position
        Me.Top = Application.Top + (Application.UsableHeight / 1.5) - (Me.Height / 2) 'Open near the bottom of the screen
        Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)
    
    'Add the values for the LOB ListBox
        Me.lst_LOB.List = Get_LOB_Array
        
' -----------
' Show the worksheets and objects that should be visible to a Privileged User
' -----------
        
    Call Me.o_65_Update_Workbook_for_SCEs
    Call Me.o_24_Populate_PM_Edit_Metrics
    Call Me.o_25_Populate_Portfolio_Metrics
    
End Sub
Sub o_02_Assign_Global_Variables()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called on Initialization
' Updated: 3/26/2023

' Change Log:
'       4/23/2020:  Intial Creation
'       2/25/2021:  Updated the intLastCol to make it more resiliant
'       1/4/2023:   Added the 'col_LOBUpdated' variable
'       3/26/2023:  Removed all references to wsLists

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Assign Sheets
    
        Set wsData = ThisWorkbook.Sheets("Dashboard Review")
        
        If Evaluate("ISREF(" & "'Detailed Dashboard'" & "!A1)") = True Then
            Set wsDetailData = ThisWorkbook.Sheets("Detailed Dashboard")
            bol_wsDetailData_Exists = True
        End If

        Set wsArrays = ThisWorkbook.Sheets("Array Values")
        Set wsPivot = ThisWorkbook.Sheets("PIVOT")
            
    ' Assign Integers / Strings
    
        intLastCol = WorksheetFunction.Max( _
            wsData.Cells(1, Columns.count).End(xlToLeft).Column, _
            wsData.Rows(1).Find("").Column - 1)
            
        intLastRow = wsData.Cells(Rows.count, "A").End(xlUp).Row
            
        strLastCol_wsData = Split(Cells(1, intLastCol).Address, "$")(1)
        
        intLastRow_wsArrays = wsArrays.Cells(Rows.count, "A").End(xlUp).Row
            If intLastRow_wsArrays = 1 Then intLastRow_wsArrays = 2
            
    ' Assign Arrays
        
        arryHeader_Data = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))

    ' Assign "Ranges"
        col_Borrower = fx_Create_Headers("Customer", arryHeader_Data)
        col_CustID = fx_Create_Headers("Unique Customer ID", arryHeader_Data)
        col_LOB = fx_Create_Headers("LOB", arryHeader_Data)
        col_LOBUpdated = fx_Create_Headers("Updated LOB", arryHeader_Data)
        col_PM_Name = fx_Create_Headers("PM", arryHeader_Data)
        col_AM_Name = fx_Create_Headers("Account Manager (PM & RM)", arryHeader_Data)
        col_Region = fx_Create_Headers("Region", arryHeader_Data)
        col_Team = fx_Create_Headers("Team", arryHeader_Data)
        
        col_BRG = fx_Create_Headers("BRG", arryHeader_Data)
        col_FRG = fx_Create_Headers("FRG", arryHeader_Data)
        col_CCRP = fx_Create_Headers("CCRP", arryHeader_Data)
        
        col_Exposure = fx_Create_Headers("Webster Commitment (000's) - Gross Exposure", arryHeader_Data)
        col_Outstanding = fx_Create_Headers("Webster Outstanding (000's) - Book Balance", arryHeader_Data)
        
        col_PM_Attest = fx_Create_Headers("PM Attestation", arryHeader_Data)
        col_PM_Attest_Exp = fx_Create_Headers("PM Attestation Explanation", arryHeader_Data)
        col_CovCompl = fx_Create_Headers("Covenant Compliance", arryHeader_Data)
        col_CovCompl_Exp = fx_Create_Headers("Covenant Compliance Explanation", arryHeader_Data)
        col_ChangeFlag = fx_Create_Headers("Change Flag", arryHeader_Data)
        
    ' Assign Booleans
    
        'Dim bolPrivilegedUser As Boolean
            bolPrivilegedUser = fx_Privileged_User

End Sub
Sub o_03_Declare_Global_Arrays()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called on Initialization
' Updated: 10/6/2022

' Change Log:
'       4/23/2020:  Intial Creation
'       12/14/2020: Removed the code related to the PM Helper
'       10/6/2022:  Updated arrays to point to the wsData, not wsLists

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Dim Dictionaries
    
    'Dim dict_PMs As Scripting.Dictionary
        Set dict_PMs = New Scripting.Dictionary

    'Dim Loop Values
    
    Dim i As Long
    
    Dim val As Variant
        
' -------------------
' Add values to array
' -------------------
        
    'Public ary_SelectedCustomers
        If wsArrays.Range("A2").Value2 = "" Then
            ReDim ary_SelectedCustomers(1) 'Only set if there aren't old customers in there
        Else
            ReDim ary_SelectedCustomers(1 To 999)
            
            For i = 2 To intLastRow_wsArrays
                ary_SelectedCustomers(i - 1) = wsArrays.Range("A" & i)
            Next i
            
        End If

    'Dim ary_BorrowerLookupData
        ary_BorrowerLookupData = wsData.Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, col_AM_Name))

End Sub
Private Sub lst_LOB_Click()
    
    Call Me.o_12_Create_Borrower_List_By_LOB
    
    Call Me.o_14_Create_AM_List_From_ListBox
    
    Me.cmb_Dynamic_AM.Value = Null
        Me.cmb_Dynamic_AM.SetFocus
            
End Sub
Private Sub lst_LOB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call Me.o_32_Filter_Customers_by_LOB
        
    ' Reset the AM Name lookup
    Me.cmb_Dynamic_AM.Value = Null
        Me.cmb_Dynamic_AM.SetFocus

    ' Update the PM Edit / Portfolio Metrics
    Call Me.o_24_Populate_PM_Edit_Metrics
    Call Me.o_25_Populate_Portfolio_Metrics

End Sub
Private Sub cmb_Dynamic_AM_Change()

    Call Me.o_15_Create_AM_List_From_DynamicLookup

End Sub
Private Sub lst_AM_Click()
        
    Call Me.o_13_Create_Borrower_List_By_PM

End Sub
Private Sub lst_AM_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    'Wipe out the customer filtering
    wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Borrower
    
    'Filter by the selected PM
    Call Me.o_33_Filter_Customers_by_PM
    
    ' Update the PM Edit / Portfolio Metrics
    Call Me.o_24_Populate_PM_Edit_Metrics
    Call Me.o_25_Populate_Portfolio_Metrics

End Sub
Private Sub lst_AM_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    'If I hit enter run the code
    
    If KeyCode = vbKeyReturn And Me.lst_AM.Value <> "" Then
        Call Me.o_33_Filter_Customers_by_PM
    End If

End Sub
Private Sub cmb_Dynamic_Borrower_Change()

    Call Me.o_11_Create_Borrower_List_Dynamic

End Sub
Private Sub lst_Borrowers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
      
    'If there is only one valid record select it instead of the blank, otherwise go w/ what I picked
      
    If lst_Borrowers.ListCount = 2 Then lst_Borrowers.Selected(0) = True
        If lst_Borrowers.Value = "" Then Exit Sub
        
    Call Me.o_21_Add_Customer_To_Selected_Customers_Array
        Call Me.o_31_Filter_Customers
        
    Call Me.o_25_Populate_Portfolio_Metrics 'Added on 12/10 to test for Jason

End Sub
Private Sub lst_Borrowers_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    'If I hit enter and have selected a customer filter the list to JUST that customer
    
    If lst_Borrowers.ListCount = 2 Then lst_Borrowers.Selected(0) = True
    
    If KeyCode = vbKeyReturn And Me.lst_Borrowers.Value <> "" Then
        Call Me.o_23_Filter_Single_Customer
            Me.cmb_Dynamic_Borrower.SetFocus
    End If

End Sub
Private Sub lst_Borrowers_Enter()

    'If there is only one valid record select it, otherwise abort

    If lst_Borrowers.ListCount = 2 Then lst_Borrowers.Selected(0) = True
        If lst_Borrowers.Value = "" Then Exit Sub
        If lst_Borrowers.ListCount > 2 Then Exit Sub
        
    'Call Me.o_21_Add_Customer_To_Selected_Customers_Array
    
    Call o_23_Filter_Single_Customer

End Sub
Private Sub cmd_Filter_Anomalies_Click()

    Call Me.o_37_Filter_Edits_For_PMs

End Sub
Private Sub cmd_Apply_QCFlags_Click()

    Call o_38_Apply_QC_Flags_For_SCEs

End Sub
Private Sub cmd_Filter_CovComp_Click()

    Call Me.o_39_Filter_Covenant_Compliance

End Sub
Private Sub cmd_Filter_Missing_Attestation_Click()

    Call Me.o_36_Filter_Attestation_Status

End Sub
Private Sub cmd_Clear_Filter_Click()

    Call Me.o_41_Clear_Filters
    
    Call Me.o_42_Clear_Saved_Array

    Me.lst_LOB.Value = ""
    Me.lst_AM.Clear
    Me.lst_Borrowers.Clear
    Me.cmb_Dynamic_Borrower.Value = Null
    Me.cmb_Dynamic_AM.Value = Null

    Me.txt_PM_Updates = ""
    Me.txt_Flags_Addressed = ""
    Me.txt_Attestations_Complete = ""
    
    Me.txt_Port_Outstanding = 0
    Me.txt_Port_Exposure = 0
    Me.txt_Port_BRG = 0
    Me.txt_Port_FRG = 0
    Me.txt_Port_CCRP = 0

End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub
Sub o_11_Create_Customer_List_Dynamic()

' Purpose: To create the list of customers to be used in the Customer ListBox.
' Trigger: Start typing in the DynamicSearch combo box
' Updated: 10/6/2022

' Change Log:
'       3/23/2020:  Intial Creation
'       11/24/2020: Updated to make the field references more dynamic
'       11/25/2020: Rewrote to use an Array instead of wsLists
'       10/6/2022:  Overhauled using the fx_Create_Dynamic_Lookup_List function

' ****************************************************************************

On Error GoTo ErrorHandler

' ---------------------------------------------------
' Copy the values from the collection to the list box
' ---------------------------------------------------
    
    Dim arryBorrowersTemp As Variant
        arryBorrowersTemp = myFunctions.fx_Create_Dynamic_Lookup_List( _
        wsDataSource:=wsData, _
        col_Target:=col_Borrower, _
        str_Dynamic_Lookup_Value:=Me.cmb_Dynamic_Borrower.Value, _
        col_Criteria_Field:=col_LOBUpdated, _
        str_Criteria_Match_Value:=Me.lst_LOB.Value)

    If Not IsEmpty(arryBorrowersTemp) Then
        Me.lst_Borrowers.List = arryBorrowersTemp
    End If

Exit Sub

' -------------------------------------------------------------------------------

ErrorHandler:

Global_Error_Handling SubName:="o_11_Create_Customer_List_Dynamic", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_12_Create_Borrower_List_By_LOB()

' Purpose: To create the list of customers to be used in the Customer ListBox, based on the selected LOB.
' Trigger: Select a customer from the LOB ListBox.
' Updated: 10/6/2022

' Change Log:
'       3/23/2020: Intial Creation
'       11/24/2020: Updated to make the field references more dynamic
'       11/25/2020: Rewrote to use an Array instead of wsLists
'       10/6/2022:  Overhauled using the fx_Create_Dynamic_Lookup_List function

' ****************************************************************************

On Error GoTo ErrorHandler

' ---------------------------------------------------
' Copy the values from the collection to the list box
' ---------------------------------------------------
    
    Dim arryBorrowersTemp As Variant
        arryBorrowersTemp = myFunctions.fx_Create_Dynamic_Lookup_List( _
        wsDataSource:=wsData, _
        col_Dynamic_Lookup_Field:=col_Borrower, _
        str_Dynamic_Lookup_Value:=Me.cmb_Dynamic_Borrower.Value, _
        col_Criteria_Field:=col_LOBUpdated, _
        str_Criteria_Match_Value:=Me.lst_LOB.Value)

    If Not IsEmpty(arryBorrowersTemp) Then
        Me.lst_Borrowers.List = arryBorrowersTemp
    End If

Exit Sub

' -------------------------------------------------------------------------------

ErrorHandler:

Global_Error_Handling SubName:="o_12_Create_Borrower_List_By_LOB", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
End Sub

Sub o_13_Create_Borrower_List_By_PM()

' Purpose: To create the list of customers to be used in the Customer ListBox, based on the selected PM.
' Trigger: Select a customer from the PM ListBox.
' Updated: 10/6/2022

' Change Log:
'       3/30/2020:  Intial Creation
'       11/24/2020: Updated to make the field references more dynamic
'       11/25/2020: Rewrote to use an Array instead of wsLists
'       12/8/2020:  Updated to include the RM values
'       10/6/2022:  Overhauled using the fx_Create_Dynamic_Lookup_List function

' ****************************************************************************

' ---------------------------------------------------
' Copy the values from the collection to the list box
' ---------------------------------------------------
    
    Dim arryBorrowersTemp As Variant
        arryBorrowersTemp = myFunctions.fx_Create_Dynamic_Lookup_List( _
        wsDataSource:=wsData, _
        col_Target:=col_AM_Name, _
        str_Dynamic_Lookup_Value:=Me.cmb_Dynamic_AM.Value, _
        col_Criteria_Field:=col_LOBUpdated, _
        str_Criteria_Match_Value:=Me.lst_LOB.Value)

    If Not IsEmpty(arryBorrowersTemp) Then
        Me.lst_Borrowers.List = arryBorrowersTemp
    End If

Exit Sub

' -------------------------------------------------------------------------------

' -----------
' Error Handler
' -----------

ErrorHandler:

Global_Error_Handling SubName:="o_13_Create_Borrower_List_By_PM", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_14_Create_AM_List_From_ListBox()
   
' Purpose: To create the list of PMs to be used in the PM ListBox, based on the selected LOB.
' Trigger: Select a LOB from the LOB ListBox.
' Updated: 11/25/2020

' Change Log:
'       3/30/2020: Intial Creation
'       11/24/2020: Updated to make the field references more dynamic
'       11/25/2020: Rewrote to use an Array instead of wsLists
'       12/7/2020: Rewrote using a dictionary to pull in the PMs and RMs to one list
'       12/8/2020: Added the functionality to sort the list of PMs, based on Axcel's code

' ****************************************************************************
        
' -----------------
' Declare Variables
' -----------------
    
    'Dim Strings
    
    Dim strLOB As String
    Dim strRM As String
    Dim strPM As String
    
    'Dim Loop Variables
    
    Dim x As Long
    
    'Dim Arrays / Dictionaries

    'Dim ary_PM As Variant
        ReDim ary_PM(1 To 999)
    
    Dim coll_UniquePMs As New Collection
    
    Dim coll_SortedPMs As New Collection
    
' -----------
' Run the loop
' -----------

    lst_AM.Clear

     For x = 2 To UBound(ary_BorrowerLookupData)
        strLOB = ary_BorrowerLookupData(x, 1)
        strRM = ary_BorrowerLookupData(x, 4)
        strPM = ary_BorrowerLookupData(x, 5)

            If Me.lst_LOB = strLOB Then
                On Error Resume Next
                    coll_UniquePMs.Add Key:=strPM, Item:=strPM
                    coll_UniquePMs.Add Key:=strRM, Item:=strRM
                On Error GoTo 0
            End If
        Next x
        
    'Output the values from the collection into the array
    
    If coll_UniquePMs.count = 0 Then Exit Sub
        
    Set coll_SortedPMs = fx_QuickSort(coll_UniquePMs, 1, coll_UniquePMs.count)
        
    For x = 1 To coll_UniquePMs.count
        ary_PM(x) = coll_UniquePMs(x)
    Next x
        
' -------------------------------------
' Copy the array into the Customer List
' -------------------------------------

    ReDim Preserve ary_PM(1 To coll_UniquePMs.count)

    Me.lst_AM.List = ary_PM
        
End Sub
Sub o_15_Create_AM_List_From_DynamicLookup()

' Purpose: To create the list of PMs to be used in the PM ListBox.
' Trigger: Start typing in the PM Dynamic Search combo box
' Updated: 12/8/2020

' Change Log:
'       9/24/2020: Intial Creation
'       9/24/2020: Updated to use a new collection instead of an array
'       9/25/2020: Auto filters to just the PMs customers if only their name is left
'       11/24/2020: Updated to make the field references more dynamic
'       11/25/2020: Rewrote to use an Array instead of wsLists
'       12/8/2020: Updated to use both RM and PM values

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    'Dim Strings
    
    Dim strLOB As String
    Dim strRM As String
    Dim strPM As String

    'Dim Loop Variables
    
    Dim x As Long
    
    'Dim Arrays / Dictionaries

    'Dim ary_PM As Variant
        ReDim ary_PM(1 To 9999)

    Dim coll_UniquePMs As New Collection

' -----------
' Copy the matching values into the collection
' -----------
       
    Me.lst_AM.Clear
       
     For x = 2 To UBound(ary_BorrowerLookupData)
        strLOB = ary_BorrowerLookupData(x, 1)
        strRM = ary_BorrowerLookupData(x, 4)
        strPM = ary_BorrowerLookupData(x, 5)
    
             If InStr(1, strPM, Me.cmb_Dynamic_AM.Value, vbTextCompare) Then
                 If IsNull(Me.lst_LOB) Or Me.lst_LOB = strLOB Then
                        On Error Resume Next
                            coll_UniquePMs.Add Key:=strPM, Item:=strPM
                        On Error GoTo 0
                 End If
             End If
    
             If InStr(1, strRM, Me.cmb_Dynamic_AM.Value, vbTextCompare) Then
                 If IsNull(Me.lst_LOB) Or Me.lst_LOB = strLOB Then
                        On Error Resume Next
                            coll_UniquePMs.Add Key:=strRM, Item:=strRM
                        On Error GoTo 0
                 End If
             End If
    
    Next x

    'Output the values from the collection into the array
    
    If coll_UniquePMs.count = 0 Then Exit Sub
    
    For x = 1 To coll_UniquePMs.count
        ary_PM(x) = coll_UniquePMs(x)
    Next x

' -----------
' Copy the array into the Customer List
' -----------

    ReDim Preserve ary_PM(1 To coll_UniquePMs.count)

    Me.lst_AM.List = ary_PM

' -----------
' If only one record is left, filter based on that
' -----------
    
    If lst_AM.ListCount = 1 Then
        lst_AM.Selected(0) = True: lst_AM.Selected(0) = True
        lst_AM.Value = lst_AM.List(lst_AM.ListIndex, 0)
        Call Me.o_33_Filter_Customers_by_PM
        cmb_Dynamic_AM.SetFocus
    End If

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

Global_Error_Handling SubName:="o_15_Create_AM_List_From_DynamicLookup", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_21_Add_Customer_To_Selected_Customers_Array()

' Purpose: To add to a growing array of customers that will then be used to filter the data.
' Trigger: Double Click on a Customer in the Customer ListBox
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

On Error GoTo ErrorHandler

ReDim Preserve ary_SelectedCustomers(1 To 9999)

' -----------------
' Declare Variables
' -----------------
    
    'Dim strCustomer As String
        strCustomer = lst_Borrowers.Value

    Dim i As Long
        i = 1

    Dim intArrayLast As Long

        Do Until ary_SelectedCustomers(i) = Empty
            intArrayLast = i
            i = i + 1
        Loop

        If ary_SelectedCustomers(1) = Empty Then intArrayLast = 0

' -----------
' Input the value into the array in the first empty slot
' -----------

    ary_SelectedCustomers(intArrayLast + 1) = strCustomer

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_21_Add_Customer_To_Selected_Customers_Array", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_23_Filter_Single_Customer()

' Purpose: To filter the list of customers in the Data ws based on only the currently swelected customer in the Customers List.
' Trigger: Called: uf_Sageworks_Regular
' Updated: 8/25/2020

' Change Log:
'          4/22/2020: Intial Creation
'          8/25/2020: Added error handling if a blank was selected instead of a customer

' ****************************************************************************

If Me.lst_Borrowers.Value = "" Then Exit Sub

    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Borrower, Criteria1:=Me.lst_Borrowers.Value, Operator:=xlFilterValues
    
    End With

End Sub
Sub o_24_Populate_PM_Edit_Metrics()

' Purpose: To add the Portfolio Metrics into the userform.
' Trigger: Called
' Updated: 7/8/2021

' Change Log:
'       10/2/2020: Intial Creation
'       1/9/2021: Added the abort if not in Quarterly mode
'       7/8/2021: Updated the calc for txt_Attestations_Complete to output the % complete

' ****************************************************************************

' Abort if it isn't the Quarterly Validation Dashboard
#If Quarterly = 0 Then
    Exit Sub
#End If

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    ' Dim Integers
    Dim int_PM_Updates As Long: int_PM_Updates = 0
    Dim int_Flags_Remaining As Long: int_Flags_Remaining = 0
    Dim int_Flags_Addressed As Long: int_Flags_Addressed = 0
    
    Dim int_PM_Attest_Complete As Long: int_PM_Attest_Complete = 0
    Dim int_PM_Attest_Blank As Long: int_PM_Attest_Blank = 0
    
    ' Dim Colors
    Dim intYellow As Long
        intYellow = RGB(254, 255, 102)

    Dim clrOrange As Long
        clrOrange = RGB(253, 223, 199)
        
    Dim intGreen As Long
        intGreen = RGB(236, 241, 222)

    ' Dim Ranges
    Dim rngData As Range
        Set rngData = wsData.Range(wsData.Cells(2, 1), wsData.Cells(intLastRow, intLastCol)) '.SpecialCells(xlCellTypeVisible)
        
    Dim rngAttestationCol As Range
        Set rngAttestationCol = wsData.Range(wsData.Cells(2, col_PM_Attest), wsData.Cells(intLastRow, col_PM_Attest)).SpecialCells(xlCellTypeVisible)

    ' Dim Loop Variables
    
    Dim cell As Variant
    
    'Dim i As Long

' -----------
' Set the values for Portfolio Metrics
' -----------

    ' Input the values for the PM Updates / Credit Risk Flags Addressed

    For Each cell In rngData.SpecialCells(xlCellTypeVisible)
        If cell.Interior.Color = intYellow Then
            int_PM_Updates = int_PM_Updates + 1
        ElseIf cell.Interior.Color = clrOrange Then
            int_Flags_Remaining = int_Flags_Remaining + 1
        ElseIf cell.Interior.Color = intGreen Then
            int_Flags_Addressed = int_Flags_Addressed + 1
        End If
    Next cell
    
    Me.txt_PM_Updates = int_PM_Updates
    
    Me.txt_Flags_Addressed = int_Flags_Addressed & " / " & int_Flags_Addressed + int_Flags_Remaining
    
    ' Input the values for the PM Attestations Completed
    
    For Each cell In rngAttestationCol
        If cell.Value = "" Then
            int_PM_Attest_Blank = int_PM_Attest_Blank + 1
        ElseIf cell.Value <> "" Then
            int_PM_Attest_Complete = int_PM_Attest_Complete + 1
        End If
    Next cell
    
    Me.txt_Attestations_Complete = int_PM_Attest_Complete & " / " & Format(int_PM_Attest_Complete + int_PM_Attest_Blank, "#,#") & _
        " (" & Format(int_PM_Attest_Complete / (int_PM_Attest_Complete + int_PM_Attest_Blank), "0%") & ")"

' -----------
' Error Handler
' -----------

ErrorHandler:
    Exit Sub

End Sub
Sub o_25_Populate_Portfolio_Metrics()

' Purpose: To add the Portfolio Metrics into the userform.
' Trigger: TBD
' Updated: 12/10/2020

' Change Log:
'       12/10/2020: Intial Creation
'       12/10/2020: Built out to include BRG, FRG, CCRP
'       12/14/2020: Updated to only show one digit if the BRG, FRG, CCRP rounds out to an integer

' ****************************************************************************

' -----------
' Set the values for the Portfolio Summary Metrics
' -----------

With wsData

On Error Resume Next

    ' Calculate Outstanding
    Dim dbl_Port_Outstanding As Double
        dbl_Port_Outstanding = Application.WorksheetFunction.Sum(.Range(.Cells(2, col_Outstanding), .Cells(intLastRow, col_Outstanding)).SpecialCells(xlCellTypeVisible))
        
        dbl_Port_Outstanding = dbl_Port_Outstanding / 10 ^ 3
    
    Me.txt_Port_Outstanding = Format(dbl_Port_Outstanding, "$#,## MM")

    ' Calculate Exposure
    Dim dbl_Port_Exposure As Double
        dbl_Port_Exposure = Application.WorksheetFunction.Sum(.Range(.Cells(2, col_Exposure), .Cells(intLastRow, col_Exposure)).SpecialCells(xlCellTypeVisible))
        
        dbl_Port_Exposure = dbl_Port_Exposure / 10 ^ 3
    
    Me.txt_Port_Exposure = Format(dbl_Port_Exposure, "$#,## MM")
    
    ' Calculate BRG
    Dim dbl_Port_BRG As Double
        dbl_Port_BRG = Application.WorksheetFunction.Average(.Range(.Cells(2, col_BRG), .Cells(intLastRow, col_BRG)).SpecialCells(xlCellTypeVisible))

    If dbl_Port_BRG <> Int(dbl_Port_BRG) Then
        Me.txt_Port_BRG = Format(dbl_Port_BRG, "#.##")
    Else
        Me.txt_Port_BRG = Format(dbl_Port_BRG, "#.00")
    End If
    
    ' Calculate FRG
    Dim dbl_Port_FRG As Double
        dbl_Port_FRG = Application.WorksheetFunction.Average(.Range(.Cells(2, col_FRG), .Cells(intLastRow, col_FRG)).SpecialCells(xlCellTypeVisible))

    If dbl_Port_FRG <> Int(dbl_Port_FRG) Then
        Me.txt_Port_FRG = Format(dbl_Port_FRG, "#.##")
    Else
        Me.txt_Port_FRG = Format(dbl_Port_FRG, "#.00")
    End If

    ' Calculate CCRP
    Dim dbl_Port_CCRP As Double
        dbl_Port_CCRP = Application.WorksheetFunction.Average(.Range(.Cells(2, col_CCRP), .Cells(intLastRow, col_CCRP)).SpecialCells(xlCellTypeVisible))

    If dbl_Port_CCRP <> Int(dbl_Port_CCRP) Then
        Me.txt_Port_CCRP = Format(dbl_Port_CCRP, "#.##")
    Else
        Me.txt_Port_CCRP = Format(dbl_Port_CCRP, "#.00")
    End If

On Error GoTo 0

End With

End Sub
Sub o_31_Filter_Customers()

' Purpose: To filter the list of customers in the Data ws based on the Customer Selected Array.
' Trigger: Called: uf_Sageworks_Regular.cmd_Filter_Customers
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

If ary_SelectedCustomers(1) = "" Then
    MsgBox "No Customers were selected, please double click to select customers or Filter by LOB"
    Exit Sub
End If

Call myPrivateMacros.DisableForEfficiency

On Error GoTo ErrorHandler

' -----------
' Filter the Dashboard Review worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Borrower, Criteria1:=ary_SelectedCustomers, Operator:=xlFilterValues
    
    End With

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_31_Filter_Customers", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_32_Filter_Customers_by_LOB()

' Purpose: To filter the list of customers in the Data ws based on the selected LOB.
' Trigger: Called: uf_Sageworks_Regular.cmd_Filter_Customers_by_LOB
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency

On Error GoTo ErrorHandler

Dim aryTEMP As Variant
    If Me.lst_LOB.Value = "Commercial Workout" Then aryTEMP = Application.Transpose(Me.lst_Borrowers.List)
        
' -----------
' Filter the Dashboard Review worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_LOB, Criteria1:=Me.lst_LOB.Value, Operator:=xlFilterValues
    End With

    If Me.lst_LOB.Value = "Commercial Workout" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_LOB
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Borrower, Criteria1:=aryTEMP, Operator:=xlFilterValues
    End If

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_32_Filter_Customers_by_LOB", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_33_Filter_Customers_by_PM()

' Purpose: To filter the list of customers in the Data ws based on the selected Market.
' Trigger: Called: uf_Sageworks_Regular.cmd_Filter_Customers_by_LOB
' Updated: 3/23/2020

' Change Log:
'       3/23/2020: Intial Creation
'       12/5/2020: Converted to use an AdvancedFilter to pull in PM and RM
'       12/8/2020: Added the If statment if there is only one value in the list
'       12/14/2020: Converted back to AutoFilter so flags can be applied to filtered accounts

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

Dim strSelectedPM As String
    If Me.lst_AM.Value <> "" Then
        strSelectedPM = Me.lst_AM.Value
    Else
        strSelectedPM = lst_AM.List(lst_AM.ListIndex, 0)
    End If

' -----------
' Filter the Dashboard Review worksheet based on the selected PM / RM
' -----------
    
    wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_AM_Name, Criteria1:="*" & strSelectedPM & "*", Operator:=xlFilterValues
    
    'wsData.Range("A:" & strLastCol_wsData).AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=[AryVal_RMPM_Filter], Unique:=False

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_33_Filter_Customers_by_PM", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_34_Save_Array_Values()

' Purpose: To save the selected customers so that they can be accessed later.
' Trigger: Called: uf_Sageworks_Regular.cmd_Filter_Customers_Click
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

' -----------
' Copy the values from the array into the worksheet
' -----------

On Error Resume Next

    wsArrays.Range("A2:A" & UBound(ary_SelectedCustomers) + 1) = WorksheetFunction.Transpose(ary_SelectedCustomers)

End Sub
Sub o_35_Filter_Only_Changes()

' Purpose: To filter the Data ws down to only records that had a Change before emailing Credit Risk.
' Trigger: Called: cmd_Email_Credit_Risk_Click
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency

On Error GoTo ErrorHandler
       
' -----------
' Filter the Dashboard Review worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
    
      .Cells.AutoFilter
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ChangeFlag, Criteria1:="CHANGE", Operator:=xlFilterValues
    
    End With

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_35_Filter_Only_Changes", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_36_Filter_Attestation_Status()

' Purpose: To filter the data down to just the customers that have or have not been attested to.
' Trigger: Called: cmd_Filter_Missing_Attestation
' Updated: 9/28/2020

' Change Log:
'          9/28/2020: Intial Creation

' ****************************************************************************

    Dim strDefaultCaption As String
        strDefaultCaption = "Check missing attestations"

' -----------
' Filter the data
' -----------
          
     If bol_AttestationStatus = "" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_PM_Attest, Criteria1:="*", Operator:=xlFilterValues
        Me.cmd_Filter_Missing_Attestation.BackColor = RGB(235, 241, 222)
        Me.cmd_Filter_Missing_Attestation.Caption = "Attestation Complete"
        bol_AttestationStatus = "Attestation Complete"
          
    ElseIf bol_AttestationStatus = "Attestation Complete" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_PM_Attest, Criteria1:="", Operator:=xlFilterValues
        Me.cmd_Filter_Missing_Attestation.BackColor = RGB(242, 220, 219)
        Me.cmd_Filter_Missing_Attestation.Caption = "Missing Attestation"
        bol_AttestationStatus = "No Attestation"
    Else
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_PM_Attest
        Me.cmd_Filter_Missing_Attestation.BackColor = RGB(240, 240, 240)
        Me.cmd_Filter_Missing_Attestation.Caption = strDefaultCaption
        bol_AttestationStatus = ""
        
    End If

End Sub
Sub o_37_Filter_Edits_For_PMs()

' Purpose: To filter the Data down to just the Edits that still need to be addressed by the PMs.
' Trigger: Called: cmd_Filter_Anomalies
' Updated: 11/23/2020

' Change Log:
'          11/23/2020: Intial Creation
    
' ****************************************************************************

On Error GoTo ErrorHandler

myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    'Dim "Ranges"
    
    Dim col_EditFlag As Long
        col_EditFlag = fx_Create_Headers("Edit Flag", arryHeader_Data)
        
    Dim rngData As Range
        Set rngData = wsData.Range(wsData.Cells(2, 1), wsData.Cells(intLastRow, intLastCol))

    Dim cell As Variant

    ' Dim Colors

    Dim clrOrange As Long
        clrOrange = RGB(253, 223, 199)
        
' -----------
' Refresh the Filter Flag data
' -----------
    
    'Clear out the old data
    wsData.Range(wsData.Cells(2, col_EditFlag), wsData.Cells(intLastRow, col_EditFlag)).ClearContents

    For Each cell In rngData.SpecialCells(xlCellTypeVisible) 'Visible only to account for filtered data
        If cell.Interior.Color = clrOrange Then
            wsData.Cells(cell.Row, col_EditFlag).Value2 = "Yes"
        End If
    Next cell
    
' -----------
' Filter the data
' -----------
          
    If bol_Edit_Filter = False Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_EditFlag, Criteria1:="Yes", Operator:=xlFilterValues
        Me.cmd_Filter_Anomalies.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_Anomalies.Caption = "Anomalies Only"
    ElseIf bol_Edit_Filter = True Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_EditFlag
        Me.cmd_Filter_Anomalies.BackColor = RGB(240, 240, 240)
        Me.cmd_Filter_Anomalies.Caption = "Filter Anomalies"
    End If
    
    bol_Edit_Filter = Not bol_Edit_Filter 'Switch the boolean

myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_37_Filter_Edits_For_PMs", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    myPrivateMacros.DisableForEfficiencyOff
End Sub
Sub o_38_Apply_QC_Flags_For_SCEs()

' Purpose: To apply the QC Flags on the Data to highlight in RED.
' Trigger: Called: cmd_Apply_QCFlags
' Updated: 6/15/2020

' Change Log:
'       12/1/2020: Intial Creation
'       12/21/2020: Updated so that only the specific field that caused the flag turns red (ex. % Chng. (YTD vs. PYTD))
'       6/15/2021: Removed the Spreads Ratings Outlook, as per Eric R.
    
' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------
    
    'Dim "Ranges"
        
    Dim col_BRG
        col_BRG = fx_Create_Headers("BRG", arryHeader_Data)
        
    Dim col_FRG
        col_FRG = fx_Create_Headers("FRG", arryHeader_Data)
        
    Dim col_Change_YTD_v_Prior
        col_Change_YTD_v_Prior = fx_Create_Headers("% Chng. (YTD vs. PYTD) (calculated)", arryHeader_Data)
        
'    Dim col_Spreads_Ratings_Outlook
'        col_Spreads_Ratings_Outlook = fx_Create_Headers("Spreads Ratings Outlook", arryHeader_Data)

    ' Dim Colors
    
    Dim intRed1 As Long
        intRed1 = RGB(242, 220, 219)
        
    Dim intWhite As Long
        intWhite = RGB(255, 255, 255)
    
    'Dim Loop Variables
    
    Dim cell As Variant
    
' -----------
' Apply the QC Flags to highlight in RED
' -----------

    If bol_QC_Flags = False Then
        Me.cmd_Apply_QCFlags.BackColor = RGB(240, 248, 224)
        Me.cmd_Apply_QCFlags.Caption = "SCE Flags On"
        GoTo QC_FLAGS_ON
    ElseIf bol_QC_Flags = True Then
        Me.cmd_Apply_QCFlags.BackColor = RGB(240, 240, 240)
        Me.cmd_Apply_QCFlags.Caption = "SCE  Flags Off"
        GoTo QC_FLAGS_OFF
    End If
    
    Exit Sub

' -----------
' Turn the Red QC Flags On
' -----------

QC_FLAGS_ON:
    With wsData

        'Flag the BRG AND Spreads Ratings Outlook if the '% Chng. (YTD vs. PYTD)' is negative
        For Each cell In .Range(.Cells(2, col_Change_YTD_v_Prior), .Cells(intLastRow, col_Change_YTD_v_Prior))
            If cell.Value < 0 Then
                cell.Interior.Color = intRed1
            End If
        Next cell
        
        'Flag the Bright Line Waiver <> "Y" if BRG >= 8 AND FRG >= 4 AND MARKET = NOT Remediation
        
        For Each cell In .Range(.Cells(2, col_Bright_Line_Waiver), .Cells(intLastRow, col_Bright_Line_Waiver))
            If .Cells(cell.Row, col_Bright_Line_Waiver) <> "Y" Then
                If .Cells(cell.Row, col_BRG) >= 8 And .Cells(cell.Row, col_BRG) <> "6W" And .Cells(cell.Row, col_FRG) >= 4 Then
                    If InStr(1, .Cells(cell.Row, col_Team), "Remediation", vbTextCompare) = 0 Then
                        .Cells(cell.Row, col_BRG).Interior.Color = intRed1
                        .Cells(cell.Row, col_FRG).Interior.Color = intRed1
                        .Cells(cell.Row, col_Team).Interior.Color = intRed1
                        .Cells(cell.Row, col_Bright_Line_Waiver).Interior.Color = intRed1
                    End If
                End If
            End If
        Next cell
    
    End With
    
    bol_QC_Flags = Not bol_QC_Flags 'Switch the boolean
    
    Exit Sub
    
' -----------
' Turn the Red QC Flags Off
' -----------
    
QC_FLAGS_OFF:
    With wsData

        ' Reset the BRG and Spreads Ratings Outlook interior colors
        For Each cell In .Range(.Cells(2, col_Change_YTD_v_Prior), .Cells(intLastRow, col_Change_YTD_v_Prior))
            If cell.Value < 0 Then
                cell.Interior.Color = .Cells(cell.Row, col_Change_YTD_v_Prior).Interior.Color
            End If
        Next cell
        
        ' Reset the Bright Line Waiver interior color
        
        For Each cell In .Range(.Cells(2, col_Bright_Line_Waiver), .Cells(intLastRow, col_Bright_Line_Waiver))
            If .Cells(cell.Row, col_BRG) >= 8 And .Cells(cell.Row, col_BRG) <> "6W" And .Cells(cell.Row, col_FRG) >= 4 Then
                If InStr(1, .Cells(cell.Row, col_Team), "Remediation", vbTextCompare) = 0 Then
                    .Cells(cell.Row, col_BRG).Interior.Color = intWhite
                    .Cells(cell.Row, col_FRG).Interior.Color = intWhite
                    .Cells(cell.Row, col_Team).Interior.Color = intWhite
                    .Cells(cell.Row, col_Bright_Line_Waiver).Interior.Color = intWhite
                End If
            End If
        Next cell
    
    End With
    
    bol_QC_Flags = Not bol_QC_Flags 'Switch the boolean
    
    Exit Sub
    
' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_38_Apply_QC_Flags_For_SCEs", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_39_Filter_Covenant_Compliance()

' Purpose: To filter the Data down to just the Edits that still need to be addressed by the PMs.
' Trigger: Called: cmd_Filter_Anomalies
' Updated: 11/23/2020

' Change Log:
'          11/23/2020: Intial Creation
    
' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Dim "Ranges"
    
    Dim col_EditFlag As Long
        col_EditFlag = fx_Create_Headers("Edit Flag", arryHeader_Data)
        
    Dim rngData As Range
        Set rngData = wsData.Range(wsData.Cells(2, 1), wsData.Cells(intLastRow, intLastCol))

    Dim cell As Variant

    ' Dim Colors
    
    Dim clrOrange As Long
        clrOrange = RGB(253, 223, 199)

    Dim strDefaultCaption As String
        strDefaultCaption = "Filter Cov Comp"

' -----------
' Filter the data
' -----------
          
     If bol_CovCompliance = "" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_CovCompl, Criteria1:="*In Compliance*", Operator:=xlFilterValues
        Me.cmd_Filter_CovComp.BackColor = RGB(235, 241, 222)
        Me.cmd_Filter_CovComp.Caption = "In Compliance"
        bol_CovCompliance = "In Compliance"
          
    ElseIf bol_CovCompliance = "In Compliance" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_CovCompl, Criteria1:="*Out of Compliance*", Operator:=xlFilterValues
        Me.cmd_Filter_CovComp.BackColor = RGB(242, 220, 219)
        Me.cmd_Filter_CovComp.Caption = "Out of Compliance"
        bol_CovCompliance = "Out of Compliance"
    
    ElseIf bol_CovCompliance = "Out of Compliance" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_CovCompl, Criteria1:="*Compliance Waived*", Operator:=xlFilterValues
        Me.cmd_Filter_CovComp.BackColor = RGB(242, 220, 219)
        Me.cmd_Filter_CovComp.Caption = "Compliance Waived"
        bol_CovCompliance = "Compliance Waived"
    
    Else
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_CovCompl
        Me.cmd_Filter_CovComp.BackColor = RGB(240, 240, 240)
        Me.cmd_Filter_CovComp.Caption = strDefaultCaption
        bol_CovCompliance = ""
        
    End If

End Sub
Sub o_41_Clear_Filters()

' Purpose: To reset all of the current filtering.
' Trigger: Called: uf_Sageworks_Regular.cmd_Clear_Filter
' Updated: 11/16/2020

' Change Log:
'       3/23/2020: Intial Creation
'       8/19/2020: Added the logic to exclude the exempt customers
'       11/16/2020: Added in the autofilter to hide CRE and ABL
'       12/29/2020: Removed the filtering for CRE and ABL and moved to o_64_Update_Workbook_for_Admins

' ****************************************************************************

' -----------
' Reset the AutoFilter for wsData
' -----------
    
    If wsData.AutoFilterMode = True Then wsData.AutoFilter.ShowAllData
    
    'Reset the Filter Edits button
    Me.cmd_Filter_Anomalies.BackColor = RGB(240, 240, 240)
    Me.cmd_Filter_Anomalies.Caption = "Filter Anomalies"
    
Exit Sub

End Sub
Sub o_42_Clear_Saved_Array()

' Purpose: To remove all of the values from the Selected Customers Array and the Array ws.
' Trigger: Called: uf_Sageworks_Regular.cmd_Clear_Filter
' Updated: 3/23/2020

' Change Log:
'       3/23/2020: Intial Creation

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Dim intLastRow_wsArrays As Long
        intLastRow_wsArrays = wsArrays.Cells(Rows.count, "A").End(xlUp).Row ' Reset intLastRow

' -----------
' Remove the old values and empty the array
' -----------

    If intLastRow <> 1 Then
        wsArrays.Range("A2:A" & intLastRow_wsArrays).Clear
    End If

    If IsEmpty(ary_SelectedCustomers) = False Then Erase ary_SelectedCustomers

End Sub
Sub o_65_Update_Workbook_for_SCEs()

' Purpose: To unhide the worksheets that the PMs don't need to see.
' Trigger: Called: uf_Sageworks_Regular
' Updated: 3/25/2022

' Change Log:
'       9/25/2020:  Initial Creation
'       11/20/2020: Updated to hide the Region column
'       12/28/2020: Updated to hide the Customer ID # column
'       12/29/2020: Updated to hide the AM column
'       2/13/2021:  Renamed and updated for SCEs
'       2/19/2021:  Added in the code to hide Customer Since, at Eric's behest
'       2/24/2021:  Updated to ONLY show the wsDetailData if it exists
'       3/25/2022:  Disabled the "Cust Since" related code since that field is no longer used

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

  '  Dim col_CustSince as Long
  '      col_CustSince = fx_Create_Headers("Customer Since", arryHeader_Data)

' -----------
' Unhide the worksheets, objects, etc.
' -----------

    ' Make the sheets visible for Admin Users

    wsPivot.Visible = xlSheetVisible
    If bol_wsDetailData_Exists = True Then wsDetailData.Visible = xlSheetVisible

    ' Make ALL fields visible for the Admin users except a select fiew

On Error Resume Next
    wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)).EntireColumn.Hidden = False
    wsData.Columns(col_Region).Hidden = True
    wsData.Columns(col_AM_Name).Hidden = True
    wsData.Columns(col_CustID).Hidden = True
On Error GoTo 0

' -----------
' Filter the data
' -----------

    ' If the AutoFilter isn't on already then turn it on
    If wsData.AutoFilterMode = False Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter
    Else
        wsData.AutoFilter.ShowAllData
    End If

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
