VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_PM_Attestation 
   Caption         =   "Customer Attestation"
   ClientHeight    =   4536
   ClientLeft      =   72
   ClientTop       =   480
   ClientWidth     =   9384.001
   OleObjectBlob   =   "uf_PM_Attestation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_PM_Attestation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare Integers
Dim intCustRow As Long

' Declare Ranges / "Ranges"
Dim col_LOB As Long
Dim col_Cust As Long

Dim col_Attestation As Long
Dim col_Attestation_Exp As Long
Dim col_CovCompliance As Long

' Declare Strings
Dim strFullName As String
Dim strTimeStamp As String

Dim strOverrideExp_OrigTextCaption As String
Dim strOverrideExp_PlaceholderText As String
Dim strOverrideExp_NewTextCaption As String

' Declare Dictionary
Dim dict_Anomalies As Scripting.Dictionary

Option Explicit

Private Sub frm_Override_Click()

End Sub

Private Sub UserForm_Initialize()

' Purpose:  To initialize the userform and properly put it on the screen.
' Trigger:  Workbook Open
' Updated:  3/31/2021
' Author:   James Rinaldi

' Change Log:
'       12/3/2020: Intial Creation
'       12/5/2020: Added additional code to add the Customer Name details in.
'       3/31/2021: Changed the column index to look to borrower name

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency

Call PublicVariables.Assign_Public_Variables
Call Me.o_02_Assign_Global_Variables

' -----------
' Declare variables
' -----------

    Dim strCustomer As String
        ''strCustomer = wsData.Cells(ActiveCell.Row, 4)
        ' 3/31/2021: Changed the column index to look to borrower name
        strCustomer = wsData.Cells(ActiveCell.Row, 1)

' -----------
' Initialize the initial values
' -----------
    
    Me.StartUpPosition = 0 'Allow you to set the position
        Me.Top = Application.Top + (Application.UsableHeight / 2) - (Me.Height / 2)
        Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)
        Me.Height = 142

    Me.cmd_Cancel.SetFocus

' -----------
' Update the UserForm with details about the customer
' -----------

    uf_PM_Attestation.Caption = uf_PM_Attestation.Caption & " - " & strCustomer
    Me.lbl_Attest_Text.Caption = Replace(Me.lbl_Attest_Text.Caption, "the selected customer", strCustomer)

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Global_Variables()

' Purpose: To Assign all of the Public variables that were declared "above the line".
' Trigger: Called on Initialization
' Updated: 12/27/2020

' Change Log:
'       12/9/2020: Intial Creation
'       12/27/2020: Updated to add in the Anomalies Dictionary

' ****************************************************************************

' -----------
' Declare variables
' -----------

    ' Assign Integers
    intCustRow = ActiveCell.Row
               
    ' Assign "Ranges"
    col_LOB = fx_Create_Headers("LOB", arryHeader)
    col_Cust = fx_Create_Headers("Customer", arryHeader)

    col_Attestation = fx_Create_Headers("PM Attestation", arryHeader)
    col_Attestation_Exp = fx_Create_Headers("PM Attestation Explanation", arryHeader)
    col_CovCompliance = fx_Create_Headers("Covenant Compliance", arryHeader)
    
    ' Assign Dictionary
    Set dict_Anomalies = fx_Remaining_Anomalies_v2(intCustRow) 'Customer Row
    
    ' Dim Strings
    Dim strAnomalyText As String
        If dict_Anomalies("Unique Anomalies Found Count") > 1 Then
            strAnomalyText = "Anomalies are"
        Else
            strAnomalyText = "Anomaly is"
        End If
    
    ' Assign Strings
    strFullName = myFunctions.fx_Name_Reverse()
    strTimeStamp = Format(Now, "m/d/yyyy hh:mm")
    
    strOverrideExp_OrigTextCaption = "Please provide one sentence to explain why the remaining [ANOMALY] being overridden:"
    strOverrideExp_PlaceholderText = "[ANOMALY]"
    strOverrideExp_NewTextCaption = Replace(strOverrideExp_OrigTextCaption, strOverrideExp_PlaceholderText, strAnomalyText)

End Sub
Sub o_03_Adjust_UserForm()

' Purpose: To unhide the cmd_OverrideAttest button and explanation on the PM_Attestation UserForm if all that remains is unique QC Flags.
' Trigger: Called by cmd_Attestation_Click
' Updated: 12/27/2020

' Change Log:
'          12/27/2020: Intial Creation

' ****************************************************************************

' -----------
' Unhide the Override Explanation to prompt the user for more detail before recording the attestation.
' -----------

    Me.Height = 250
    Me.lbl_Override_ExplanationPrompt.Caption = strOverrideExp_NewTextCaption
    Me.txt_Override_Explanation.SetFocus
               
End Sub
Private Sub cmd_Attestation_Click()
    
' -----------
' Declare variables
' -----------
    
    Dim strAnomalyText As String
        If dict_Anomalies("Unique Anomalies Found Count") > 1 Then
            strAnomalyText = "Anomalies remain: "
        Else
            strAnomalyText = "Anomaly remains: "
        End If
    
    ' Update the form if only Unique Anomalies remain
    
    If dict_Anomalies("Anomalies Found Count") = dict_Anomalies("Unique Anomalies Found Count") And _
    dict_Anomalies("Unique Anomalies Found Count") > 0 Then
        MsgBox Title:="Caution", Buttons:=vbInformation, _
        Prompt:="The below unaddressed " & strAnomalyText & Chr(10) & _
        fx_Remaining_Anomalies_List(Selection.Row) & Chr(10) & Chr(10) & _
        "If you wish to continue you can override."
        Call Me.o_03_Adjust_UserForm
    Else
        Call Me.o_1_Record_Attestation
    End If
    
End Sub
Private Sub cmd_Override_SubmitExplanation_Click()
    
    If Len(Me.txt_Override_Explanation.Value) >= 5 Then
        Call Me.o_2_Record_Override_Explanation
        Call Me.o_1_Record_Attestation
        Unload Me
    Else
        MsgBox Title:="Ah Ah Ah", Buttons:=vbExclamation, Prompt:="You need to include the explanation for the override."
    End If

End Sub
Private Sub txt_Override_Explanation_Change()

'When the user starts to type the Override Explanation default to the Submit button instead of Cancel

    Me.cmd_Override_SubmitExplanation.Default = True

End Sub
Private Sub cmd_Cancel_Click()

    Unload Me

End Sub
Sub o_1_Record_Attestation()

' Purpose: To record the user attestation for each customer.
' Trigger: cmd_Attestation Click
' Updated: 12/28/2020

' Change Log:
'       12/3/2020: Initial Creation, based on the initial attestation code.
'       12/9/2020: Streamlined by moving all of the variables to their own procedure.
'       12/28/2020: Added more fields, including Old Value, Field Changed, and updated strFullAttestation

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim strFullAttestation As String
        strFullAttestation = strFullName & " (" & strTimeStamp & ")"
    
    Dim strOldValue As String
        strOldValue = wsData.Cells(intCustRow, col_Attestation).Value2

' -----------
' Add in the attestation
' -----------
    
    'Make the change, but don't trigger the Worksheet Change Event
    Application.EnableEvents = False
        wsData.Cells(intCustRow, col_Attestation) = strFullAttestation
    Application.EnableEvents = True
    
    'Track the change in the change log
    With wsChangeLog
        .Range("A" & intCurRow_ChangeLog).Value2 = strTimeStamp                                 ' Change Made Data
        .Range("B" & intCurRow_ChangeLog).Value2 = strFullName                                  ' By Who
        .Range("C" & intCurRow_ChangeLog).Value2 = wsData.Cells(intCustRow, col_LOB)         ' LOB
        .Range("D" & intCurRow_ChangeLog).Value2 = wsData.Cells(intCustRow, col_Cust)        ' Customer
        .Range("E" & intCurRow_ChangeLog).Value2 = "PM Attestation"                             ' Field Changed
        .Range("F" & intCurRow_ChangeLog).Value2 = strOldValue                                  ' Old Value
        .Range("G" & intCurRow_ChangeLog).Value2 = strFullAttestation                           ' Attestation
        .Range("H" & intCurRow_ChangeLog).Value2 = "User Attestation"                           ' Change Type
        .Range("I" & intCurRow_ChangeLog).Value2 = "Change Log"                                 ' Source
    End With

Unload Me
    
End Sub
Sub o_2_Record_Override_Explanation()

' Purpose: To record the explanation given for the Override of the remaining Unique Flags.
' Trigger: cmd_Override_SubmitExplanation_Click
' Updated: 12/10/2020

' Change Log:
'       12/10/2020: Initial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim strCovCompExplanation
        strCovCompExplanation = Me.txt_Override_Explanation
    
' -----------
' Add in the attestation
' -----------
    
    'Make the change, but don't trigger the Worksheet Change Event
    Application.EnableEvents = False
        wsData.Cells(intCustRow, col_Attestation_Exp) = strCovCompExplanation
    Application.EnableEvents = True
    
End Sub
