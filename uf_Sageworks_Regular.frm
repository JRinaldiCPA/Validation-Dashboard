VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Sageworks_Regular 
   Caption         =   "Customer Selector UserForm"
   ClientHeight    =   7976
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   14412
   OleObjectBlob   =   "uf_Sageworks_Regular.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Sageworks_Regular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim Worksheets
    Dim wsData As Worksheet
    Dim wsArrays As Worksheet
    Dim wsChangeLog As Worksheet
    Dim wsUpdates As Worksheet

'Dim Strings
    Dim strCustomer As String
    Dim strNewFileFullPath As String 'Used by o_51_Create_a_XLSX_Copy to create the XLSX to attach to the email
    Dim strLastCol_wsData As String
    Dim strUserID As String

'Dim Integers
    Dim intLastRow As Long
    Dim intLastRow_wsArrays As Long
    
    Dim intLastCol As Long
    
'Dim Data "Ranges"
    Dim col_Borrower As Long
    Dim col_LOB As Long
    Dim col_LOBUpdated As Long
    Dim col_Legacy_Bank As Long
    
    Dim col_Region As Long
    Dim col_Team As Long
    Dim col_PM_Name As Long
    Dim col_RM_Name As Long
    Dim col_AM_Name As Long
    Dim col_CRE_Flag As Long
    
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
        
'Dim Arrays / Other
    Dim arryHeader_Data()

    Dim ary_Customers
    Dim ary_SelectedCustomers
    Dim ary_PM
    
    Dim ary_BorrowerLookupData

'Dim Dictionaries
    Dim dict_PMs As Scripting.Dictionary

'Dim "Booleans"
    Dim bol_AttestationStatus As String
    Dim bol_Edit_Filter As Boolean
    Dim bol_QC_Flags As Boolean
    
Option Explicit

Private Sub cmd_Attestation_Click()

    uf_PM_Attestation.Show

End Sub

Private Sub UserForm_Initialize()

' Purpose:  To initialize the userform, including adding in the data from the arrays.
' Trigger:  Workbook Open
' Updated:  12/29/2020
' Author:   James Rinaldi

' Change Log:
'       3/23/2020:  Intial Creation
'       8/19/2020:  Added the logic to exclude the exempt customers
'       11/16/2020: I added the autofilter to hide the CRE and ABL LOBs
'       12/17/2020: Added the conditional compiler constant to determine if this is the Weekly or Quarterly file
'       12/29/2020: Moved the CRE and ABL filtering to o_64_Update_Workbook_for_Admins

' ****************************************************************************

Call Me.o_02_Assign_Global_Variables

Call Me.o_03_Declare_Global_Arrays

' -----------------------------
' Initialize the initial values
' -----------------------------

    Me.StartUpPosition = 0 'Allow you to set the position
        Me.Top = Application.Top + (Application.UsableHeight / 1.5) - (Me.Height / 2) 'Open near the bottom of the screen
        Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)
    
    'Add the values for the LOB ListBox
        Me.lst_LOB.List = Get_LOB_Array

    ' Show the Email Credit Risk if this is the Quarterly Validation Dashboard
    
    #If Quarterly = 1 Then
        Me.cmd_Email_Credit_Risk.Visible = True
    #End If

' -----------------------------------------------------------------------
' Hide the worksheets and objects that shouldn't be seen by regular users
' -----------------------------------------------------------------------

    Call Me.o_63_Update_Workbook_for_PMs
    
    Call Me.o_61_Protect_Ws
    
End Sub
Sub o_02_Assign_Global_Variables()

' Purpose: To assign all of the Public variables that were dimensioned "above the line".
' Trigger: Called on Initialization
' Updated: 3/26/2023

' Change Log:
'       4/23/2020:  Intial Creation
'       12/14/2020: Updated to reflect the changes to the Admin module
'       2/25/2021:  Updated the intLastCol to make it more resiliant
'       10/10/2022: Added the col_RM_Name variable
'       3/23/2023:  Added col_Legacy_Bank
'       3/26/2023:  Removed all references to wsLists

' ****************************************************************************

' ----------------
' Assign Variables
' ----------------
    
    ' Assign Sheets
    
        Set wsData = ThisWorkbook.Sheets("Dashboard Review")
        
        Set wsArrays = ThisWorkbook.Sheets("Array Values")
        Set wsChangeLog = ThisWorkbook.Sheets("Change Log")
        Set wsUpdates = ThisWorkbook.Sheets("Updates")
            
    ' Assign Integers / Strings
    
        intLastCol = WorksheetFunction.Max( _
            wsData.Cells(1, Columns.count).End(xlToLeft).Column, _
            wsData.Rows(1).Find("").Column - 1)
            
        intLastRow = wsData.Cells(Rows.count, "A").End(xlUp).Row
            
        strLastCol_wsData = Split(Cells(1, intLastCol).Address, "$")(1)
        
        intLastRow_wsArrays = wsArrays.Cells(Rows.count, "A").End(xlUp).Row
            If intLastRow_wsArrays = 1 Then intLastRow_wsArrays = 2
      
    ' Assign Strings
            
        strUserID = Application.UserName
            
    ' Assign Arrays
        
        arryHeader_Data = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))

    ' Assign "Ranges"
        col_Borrower = fx_Create_Headers("Customer", arryHeader_Data)
        col_LOB = fx_Create_Headers("LOB", arryHeader_Data)
        col_LOBUpdated = fx_Create_Headers("Updated LOB", arryHeader_Data)
        col_Legacy_Bank = fx_Create_Headers("Legacy Bank", arryHeader_Data)
        col_PM_Name = fx_Create_Headers("PM", arryHeader_Data)
        col_RM_Name = fx_Create_Headers("RM", arryHeader_Data)
        col_AM_Name = fx_Create_Headers("Account Manager (PM & RM)", arryHeader_Data)
        col_Region = fx_Create_Headers("Region", arryHeader_Data)
        col_Team = fx_Create_Headers("Team", arryHeader_Data)
        col_CRE_Flag = fx_Create_Headers("CRE Flag", arryHeader_Data)
        
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

End Sub
Private Sub cmb_Dynamic_AM_Change()

    Call Me.o_15_Create_AM_List_From_DynamicLookup

End Sub
Private Sub lst_AM_Click()
        
    Call Me.o_13_Create_Borrower_List_By_PM

End Sub
Private Sub lst_AM_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    'Wipe out the customer filtering
    Call Me.o_62_UnProtect_Ws
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Borrower
    Call Me.o_61_Protect_Ws
    
    'Filter by the selected PM
    Call Me.o_33_Filter_Customers_by_PM
        Unload Me
    
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
Private Sub cmd_Filter_Customers_by_LOB_Click()
    
    If Me.lst_AM <> "" Then
        Call Me.o_33_Filter_Customers_by_PM
    Else
        Call Me.o_32_Filter_Customers_by_LOB
    End If

End Sub
Private Sub cmd_Filter_Anomalies_Click()

    Call Me.o_37_Filter_Edits_For_PMs

End Sub
Private Sub cmd_Clear_Filter_Click()

    Call Me.o_41_Clear_Filters
    
    Call Me.o_42_Clear_Saved_Array

    Me.lst_LOB.Value = ""
    Me.lst_AM.Clear
    Me.lst_Borrowers.Clear
    Me.cmb_Dynamic_Borrower.Value = Null
    Me.cmb_Dynamic_AM.Value = Null

End Sub
Private Sub cmd_Email_Credit_Risk_Click()

Call myPrivateMacros.DisableForEfficiency

    Call Me.o_71_Create_Faux_Change_Log
    
    Call Me.o_35_Filter_Only_Changes
    
    Call Me.o_51_Create_a_XLSX_Copy

    Call Me.o_52_Email_Credit_Risk
       
Call myPrivateMacros.DisableForEfficiencyOff
    
    Unload Me

End Sub
Private Sub cmd_ZoomIn_Click()

    Call fxUserForm_ZoomIn(Me)
    
End Sub
Private Sub cmd_ZoomOut_Click()

    Call fxUserForm_ZoomOut(Me)

End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub
Sub o_11_Create_Borrower_List_Dynamic()

' Purpose: To create the list of customers to be used in the Customer ListBox.
' Trigger: Start typing in the DynamicSearch combo box
' Updated: 10/6/2022

' Change Log:
'       3/23/2020:  Intial Creation
'       11/24/2020: Updated to make the field references more dynamic
'       11/25/2020: Rewrote to use an Array instead of wsLists
'       10/6/2022:  Overhauled using the fx_Create_Dynamic_Lookup_List function

' ****************************************************************************

'On Error GoTo ErrorHandler

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

Global_Error_Handling SubName:="o_11_Create_Borrower_List_Dynamic", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

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
        col_Target_Field:=col_Borrower, _
        col_Dynamic_Lookup_Field:=col_AM_Name, _
        str_Dynamic_Lookup_Value:=Me.lst_AM.Value)
    
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
' Updated: 10/10/2022

' Change Log:
'       3/30/2020:  Intial Creation
'       11/24/2020: Updated to make the field references more dynamic
'       11/25/2020: Rewrote to use an Array instead of wsLists
'       12/7/2020:  Rewrote using a dictionary to pull in the PMs and RMs to one list
'       12/8/2020:  Added the functionality to sort the list of PMs, based on Axcel's code
'       10/10/2022: Updated ary_BorrowerLookupData to point to wsData, so needed to update cell references

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
        ReDim ary_PM(1 To 99999)
    
    Dim coll_UniquePMs As New Collection
    
    Dim coll_SortedPMs As New Collection
    
' ----------------------
' Create the list of AMs
' ----------------------

    lst_AM.Clear

     For x = 2 To UBound(ary_BorrowerLookupData)
        strLOB = ary_BorrowerLookupData(x, col_LOBUpdated)
        strRM = ary_BorrowerLookupData(x, col_RM_Name)
        strPM = ary_BorrowerLookupData(x, col_PM_Name)

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
        
' -----------
' Copy the array into the Customer List
' -----------

    ReDim Preserve ary_PM(1 To coll_UniquePMs.count)

    Me.lst_AM.List = ary_PM
        
End Sub
Sub o_15_Create_AM_List_From_DynamicLookup()

' Purpose: To create the list of PMs to be used in the PM ListBox.
' Trigger: Start typing in the PM Dynamic Search combo box
' Updated: 10/10/2022

' Change Log:
'       9/24/2020: Intial Creation
'       9/24/2020: Updated to use a new collection instead of an array
'       9/25/2020: Auto filters to just the PMs customers if only their name is left
'       11/24/2020: Updated to make the field references more dynamic
'       11/25/2020: Rewrote to use an Array instead of wsLists
'       12/8/2020: Updated to use both RM and PM values
'       10/10/2022: Updated ary_BorrowerLookupData to point to wsData, so needed to update cell references

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
        ReDim ary_PM(1 To 999)

    Dim coll_UniquePMs As New Collection

' --------------------------------------------
' Copy the matching values into the collection
' --------------------------------------------
       
    Me.lst_AM.Clear
       
     For x = 2 To UBound(ary_BorrowerLookupData)
        strLOB = ary_BorrowerLookupData(x, col_LOBUpdated)
        strRM = ary_BorrowerLookupData(x, col_RM_Name)
        strPM = ary_BorrowerLookupData(x, col_PM_Name)
    
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

' -------------------------------------
' Copy the array into the Customer List
' -------------------------------------

    ReDim Preserve ary_PM(1 To coll_UniquePMs.count)

    Me.lst_AM.List = ary_PM

' ------------------------------------------------
' If only one record is left, filter based on that
' ------------------------------------------------
    
    If lst_AM.ListCount = 1 Then
        lst_AM.Selected(0) = True: lst_AM.Selected(0) = True
        lst_AM.Value = lst_AM.List(lst_AM.ListIndex, 0)
        Call Me.o_33_Filter_Customers_by_PM
        cmb_Dynamic_AM.SetFocus
    End If

Exit Sub

' -------------
' Error Handler
' -------------

ErrorHandler:

Global_Error_Handling SubName:="o_15_Create_AM_List_From_DynamicLookup", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_21_Add_Customer_To_Selected_Customers_Array()

' Purpose: To add to a growing array of customers that will then be used to filter the data.
' Trigger: Double Click Customer
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

Call Me.o_62_UnProtect_Ws

    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Borrower, Criteria1:=Me.lst_Borrowers.Value, Operator:=xlFilterValues
    
    End With

Call Me.o_61_Protect_Ws

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

Call Me.o_62_UnProtect_Ws

' -----------
' Filter the Dashboard Review worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Borrower, Criteria1:=ary_SelectedCustomers, Operator:=xlFilterValues
    
    End With

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call Me.o_61_Protect_Ws

Call myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_31_Filter_Customers", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call Me.o_61_Protect_Ws
    
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

Call Me.o_62_UnProtect_Ws
        
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
    
Call Me.o_61_Protect_Ws

Call myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_32_Filter_Customers_by_LOB", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call Me.o_61_Protect_Ws
    
    Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_33_Filter_Customers_by_PM()

' Purpose: To filter the list of customers in the Data ws based on the selected Market.
' Trigger: Called: uf_Sageworks_Regular.cmd_Filter_Customers_by_LOB
' Updated: 3/23/2020

' Change Log:
'       3/23/2020: Intial Creation
'       12/5/2020: Converted to use an AdvancedFilter to pull in PM and RM
'       12/14/2020: Converted back to AutoFilter so flags can be applied to filtered accounts

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency

On Error GoTo ErrorHandler

Call Me.o_62_UnProtect_Ws

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
' Filter the Dashboard Review worksheet based on the customers selected previously
' -----------
    
    wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_AM_Name, Criteria1:="*" & strSelectedPM & "*", Operator:=xlFilterValues
    
    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call Me.o_61_Protect_Ws

Call myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_33_Filter_Customers_by_PM", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call Me.o_61_Protect_Ws
    
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

Call Me.o_62_UnProtect_Ws
       
' -----------
' Filter the Dashboard Review worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
    
      .Cells.AutoFilter
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ChangeFlag, Criteria1:="CHANGE", Operator:=xlFilterValues
    
    End With

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call Me.o_61_Protect_Ws

Call myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_35_Filter_Only_Changes", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call Me.o_61_Protect_Ws
    
    Call myPrivateMacros.DisableForEfficiencyOff

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

Call Me.o_62_UnProtect_Ws

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

    Call Me.o_61_Protect_Ws

myPrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_37_Filter_Edits_For_PMs", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    Call Me.o_61_Protect_Ws
    myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_41_Clear_Filters()

' Purpose: To reset all of the current filtering.
' Trigger: Called: uf_Sageworks_Regular.cmd_Clear_Filter
' Updated: 11/16/2020

' Change Log:
'   3/23/2020: Intial Creation
'   8/19/2020: Added the logic to exclude the exempt customers
'   11/16/2020: Added in the autofilter to hide CRE and ABL

' ****************************************************************************

On Error GoTo ErrorHandler

Call Me.o_62_UnProtect_Ws

' -----------
' If the AutoFilter is on turn it off and then reapply
' -----------

    If wsData.AutoFilterMode = True Then wsData.AutoFilter.ShowAllData
    
    'Reset the Filter Edits button
    Me.cmd_Filter_Anomalies.BackColor = RGB(240, 240, 240)
    Me.cmd_Filter_Anomalies.Caption = "Filter Anomalies"
    
Call Me.o_61_Protect_Ws

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_41_Clear_Filters", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    Call Me.o_61_Protect_Ws

End Sub
Sub o_42_Clear_Saved_Array()

On Error GoTo ErrorHandler

' Purpose: To remove all of the values from the Selected Customers Array and the Array ws.
' Trigger: Called: uf_Sageworks_Regular.cmd_Clear_Filter
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Dim intLastRow_wsArrays As Long
        intLastRow_wsArrays = wsArrays.Cells(Rows.count, "A").End(xlUp).Row

' -----------
' Remove the old values and empty the array
' -----------

    If intLastRow <> 1 Then
        wsArrays.Range("A2:A" & intLastRow_wsArrays).Clear
    End If

    If IsEmpty(ary_SelectedCustomers) = False Then Erase ary_SelectedCustomers

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_42_Clear_Saved_Array", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_51_Create_a_XLSX_Copy()

On Error GoTo ErrorHandler

' Purpose: To create a copy of the workbook in XLSX to aid in providing it to Lizzy.
' Trigger: Called: uf_Sageworks_Regular
' Updated: 3/31/2020

' Change Log:
'   3/31/2020: Updated the strOldFileName to include TEMP and look simpler

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strFullName As String
        strFullName = myFunctions.fx_Name_Reverse()

    Dim objFSO As Object
        Set objFSO = VBA.CreateObject("Scripting.FileSystemObject")

    'Public strNewFileFullPath As String
        strNewFileFullPath = ThisWorkbook.path & "\" & Format(Now, "mm-dd") & " " & Format(Now, "HH-MM-SS") & " Sageworks Validation Dashboard Update by " & strFullName
        strNewFileFullPath = Replace(strNewFileFullPath, "/", "\")

    Dim strOldFileName As String
        strOldFileName = "SAGEWORKS DASHBOARD - TEMP" & "(" & Format(Now, "mm-dd") & " " & Format(Now, "HH-MM-SS") & ")" & ".xlsm"
    
    Dim strOldFileFullPath As String
        strOldFileFullPath = ThisWorkbook.path & "\" & strOldFileName
        strOldFileFullPath = Replace(strOldFileFullPath, "/", "\")

' -----------
' Create a copy of the workbook using the current date, time, and the individuals name
' -----------

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
        
        ThisWorkbook.Save
        ThisWorkbook.SaveCopyAs strOldFileFullPath
        
        Workbooks.Open strOldFileFullPath
        
        Workbooks(strOldFileName).SaveAs _
            Filename:=strNewFileFullPath, FileFormat:=xlOpenXMLWorkbook
        
        ActiveWorkbook.Close
                
    On Error Resume Next
        If Right(strOldFileFullPath, 5) = ".xlsm" Then Kill strOldFileFullPath
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

Debug.Print ThisWorkbook.FullName 'Take the full path incase something goes awry

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

Global_Error_Handling SubName:="o_51_Create_a_XLSX_Copy", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_52_Email_Credit_Risk()

On Error GoTo ErrorHandler

' Purpose: To attach the template to an email to send to Credit Reporting & Analysis.
' Trigger: Called: uf_Sageworks_Regular
' Updated: 3/31/2021

' Change Log:
'       3/23/2020: Intial Creation
'       12/30/2020: Added in the strTempVersionv2
'       3/31/2021: Added extra code to avoid spell corrector
'       4/30/2021: Removed the code related to strTempVersion, which pulled the version from the title

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

On Error Resume Next

    Dim strTempVersionv2 As String
        strTempVersionv2 = [V_Dashboard_Version]

' -----------
' Send the Save Me email
' -----------
    
        strbody = "James," & vbNewLine & vbNewLine & _
            "This is an automated notification of an update to the Sageworks Dashboard." & vbNewLine & vbNewLine & _
            "Thanks," & vbNewLine & _
            "The Sageworks Validation Dashboard" & vbNewLine & vbNewLine & _
            "Diagnostic Info:" & vbNewLine & vbNewLine & _
            "Dashboard Version: " & strTempVersionv2 & vbNewLine & _
            "User: " & strUserID & vbNewLine & _
            "File Path: " & ThisWorkbook.FullName
        
        With OutMail
            .CC = "JRinaldi@WebsterBank.com"
            .Subject = "Sageworks Validation Dashboard Update - Auto Email"
            .Body = strbody
            .Display
            .Attachments.Add strNewFileFullPath & ".xlsx"
                ''Application.SendKeys "%s"
            '   3/31/2021: Added extra code to avoid spell corrector
            'Send Key: Send
            Application.SendKeys "%{s}", True
            'Send Key: Cancel/Escape Spell Checking
            Application.SendKeys "{ESC}", True
            'Send Key: Yes
            Application.SendKeys "%{y}", True
            
        End With
        
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing

    MsgBox _
    Title:="It worked!", _
    Buttons:=vbInformation, _
    Prompt:="Your email has been sent, and the Sageworks Validation Dashboard was saved, you may now exit. " & Chr(10) & Chr(10) _
    & "If you want to make additional changes please open the Excel workbook called 'Sageworks Validation Dashboard (v XX.X)' not the TEMP file with your name in the title."

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_52_Email_Credit_Risk", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_61_Protect_Ws()

' Purpose: To protect the Dashboard Review worksheet from manipulation.
' Trigger: Called: uf_Sageworks_Regular
' Updated: 12/17/2020

' Change Log:
'       12/17/2020: Added the conditional compiler constant to abort if it isn't the Quarterly Validation file.

' ****************************************************************************

#If Quarterly = 0 Then
    Exit Sub
#End If

On Error GoTo ErrorHandler

' -----------
' Turn on data protection
' -----------

    With wsData
        .Protect AllowFiltering:=True, AllowSorting:=True, AllowFormattingCells:=True
        .EnableAutoFilter = True
        .EnableSelection = xlUnlockedCells
    End With

Exit Sub

ErrorHandler:

    Debug.Print Global_Error_Handling("o_61_Protect_Ws", Err.Source, Err.Number, Err.Description)
    Global_Error_Handling SubName:="o_21_Add_Customer_To_Selected_Customers_Array", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_62_UnProtect_Ws()

    wsData.Unprotect

End Sub
Sub o_63_Update_Workbook_for_PMs()

' Purpose: To hide the worksheets that the PMs don't need to see.
' Trigger: Called: uf_Sageworks_Regular
' Updated: 3/23/2023

' Change Log:
'       9/25/2020: Initial Creation
'       11/20/2020: Updataed to hide the Region column
'       12/17/2020: Added the conditional compiler constant to hide certain fields if it isn't the Quarterly Validation file.
'       12/22/2020: Made it so the fields UNHIDE if Quarterly Mode is on
'       12/22/2020: Switched the WS Hiding with fx_Hide_Worksheets_For_Users
'       2/19/2021:  Added in the code to hide Customer Since, as per Eric R.
'       2/19/2021:  Added in the code to hide Collateral Description, as per Eric R.
'       12/28/2021: Hide the Instructions when not attesting
'       3/25/2022:  Disabled the "Cust Since" related code since that field is no longer used
'       3/23/2023:  Added 'Legacy Bank' and 'Segment' to be hidden
'       6/29/2023:  Added 'Team' and 'CRE Flag' to be hidden

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    'Dim Worksheets
    
        Dim wsInstructions As Worksheet
        Set wsInstructions = ThisWorkbook.Sheets("Instructions")
    
    'Dim "Ranges"
        
        Dim col_RiskExposure As Long
            col_RiskExposure = fx_Create_Headers("Webster Commitment (000's) - Risk Exposure", arryHeader_Data)
        
        Dim col_CustID As Long
            col_CustID = fx_Create_Headers("Unique Customer ID", arryHeader_Data)

        Dim col_FilterFinal As Long
            col_FilterFinal = fx_Create_Headers("Filter Flag (Final)", arryHeader_Data)

        Dim col_Review As Long
            col_Review = fx_Create_Headers("Review", arryHeader_Data)
            
        Dim col_PaidOff As Long
            col_PaidOff = fx_Create_Headers("Paid Off", arryHeader_Data)
            
        Dim col_EditFlag As Long
            col_EditFlag = fx_Create_Headers("Edit Flag", arryHeader_Data)
                        
        Dim col_CollatDesc As Long
            col_CollatDesc = fx_Create_Headers("Collateral Description", arryHeader_Data)
            
' -----------
' Hide the worksheets, objects, etc.
' -----------

On Error Resume Next
    
    ' Make most of the worksheets hidden for the PMs
    fx_Hide_Worksheets_For_Users

    ' Hide the Instructions when not in the Quarterly Attestation Mode
    
    #If Quarterly = 0 Then
        wsInstructions.Visible = xlSheetHidden
    #End If

    ' Make certain fields hidden for the PMs

    wsData.Columns(col_Region).Hidden = True
    wsData.Columns(col_Team).Hidden = True
    wsData.Columns(col_CRE_Flag).Hidden = True
    wsData.Columns(col_CustID).Hidden = True
    wsData.Columns(col_RiskExposure).Hidden = True
    wsData.Columns(col_AM_Name).Hidden = True
    wsData.Columns(col_FilterFinal).Hidden = True
    wsData.Columns(col_Review).Hidden = True
    wsData.Columns(col_PaidOff).Hidden = True
    wsData.Columns(col_EditFlag).Hidden = True
    wsData.Columns(col_CovCompl_Exp).Hidden = True
    wsData.Columns(col_PM_Attest_Exp).Hidden = True
    wsData.Columns(col_Legacy_Bank).Hidden = True
    wsData.Columns(col_LOB).Hidden = True

' Make certain fields hidden for the PMs for the Weekly file
    #If Quarterly = 0 Then
        wsData.Columns(col_PM_Attest).Hidden = True
        wsData.Columns(col_CovCompl).Hidden = True
        wsData.Columns(col_ChangeFlag).Hidden = True
    #Else
        wsData.Columns(col_PM_Attest).Hidden = False
        wsData.Columns(col_CovCompl).Hidden = False
        wsData.Columns(col_ChangeFlag).Hidden = False
    #End If
On Error GoTo 0

' -----------
' Filter the Dashboard data
' -----------

    'If the AutoFilter isn't on already then turn it on
    Call Me.o_62_UnProtect_Ws
        
        If wsData.AutoFilterMode = False Then
            wsData.Range("A:" & strLastCol_wsData).AutoFilter
        Else
            wsData.AutoFilter.ShowAllData
        End If

    Call Me.o_61_Protect_Ws

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_71_Create_Faux_Change_Log()

' Purpose: To create a change log based on actual changes made, but not recorded in the "real" Change Log.
' Trigger: Manual
' Updated: 3/31/2021
' Change Log:
'          9/1/2020: Updated for the Sageworks Dashboard project
'       1/9/2021: Added the strChangeType to replace the N/A
'       3/31/2021: Declared and assigned a value to wsPrior

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    ' Dim Integers
    
    Dim CurRowLog As Long
        CurRowLog = [MATCH(TRUE,INDEX(ISBLANK('Change Log'!A:A),0),0)]

    Dim rowID As Long
    
    Dim colID As Long
    
    ' Dim Strings / Ranges
    
    Dim strOldValue As String
    
    Dim strNewValue As String
 
    Dim strFullName As String
        strFullName = myFunctions.fx_Name_Reverse()
 
    Dim strDateTime As String
        strDateTime = Format(Now, "m/d/yyyy hh:mm:ss")
        
    Dim strChangeType As String
    
    'Dim Colors
    
    Dim intYellow As Long
        intYellow = RGB(254, 255, 102)

    Dim clrOrange As Long
        clrOrange = RGB(253, 223, 199)
        
    Dim intGreen As Long
        intGreen = RGB(236, 241, 222)
        
    '   3/31/2021: Declared and assigned a value to wsPrior
    Dim wsPrior As Worksheet
      Set wsPrior = ThisWorkbook.Sheets("Prior Dashboard")
 
' -----------
' Capture the data
' -----------
           
    For rowID = 2 To intLastRow
        With wsData
            If .Cells(rowID, col_ChangeFlag).Value2 = "CHANGE" Then 'If the record was flagged as having changes
                
                For colID = 1 To col_ChangeFlag - 1 'Everything before the Change column
                    If .Cells(rowID, colID).Value2 <> wsPrior.Cells(rowID, colID).Value2 Then
                    
                        ' Set the Change Type
                    
                        If .Cells(rowID, colID).Interior.Color = intYellow Then
                            strChangeType = "PM Change (Yellow)"
                        ElseIf .Cells(rowID, colID).Interior.Color = clrOrange Then
                            strChangeType = "Credit Risk Change (Orange)"
                        ElseIf .Cells(rowID, colID).Interior.Color = intGreen Then
                            strChangeType = "PM Resolved Credit Risk Change (Green)"
                        ElseIf .Cells(1, colID) = "Covenant Compliance" Then
                            strChangeType = "Covenant Compliance Attestation"
                        ElseIf .Cells(1, colID) = "PM Attestation" Then
                            strChangeType = "User Attestation"
                        Else
                            strChangeType = "N/A"
                        End If
                        
                        ' Copy the data into the change log
                        
                        With wsChangeLog
                            .Range("A" & CurRowLog).Value2 = strDateTime                                ' Change Made Data
                            .Range("B" & CurRowLog).Value2 = strFullName                                ' By Who
                            .Range("C" & CurRowLog).Value2 = wsData.Cells(rowID, col_LOB)            ' LOB
                            .Range("D" & CurRowLog).Value2 = wsData.Cells(rowID, col_Borrower)       ' Customer
                            .Range("E" & CurRowLog).Value2 = wsData.Cells(1, colID)                     ' Field Changed
                            .Range("F" & CurRowLog).Value2 = wsPrior.Cells(rowID, colID).Value          ' Old Value
                            .Range("G" & CurRowLog).Value2 = wsData.Cells(rowID, colID).Value           ' New Value
                            .Range("H" & CurRowLog).Value2 = strChangeType                              ' Change Type
                            .Range("I" & CurRowLog).Value2 = "Faux Log"                                 ' Source
                            CurRowLog = CurRowLog + 1
                       End With
                    
                    End If
                    
                Next colID
            
            End If
        End With
    
    Next rowID

End Sub

