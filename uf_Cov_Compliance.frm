VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Cov_Compliance 
   Caption         =   "Covenant Compliance Attestation"
   ClientHeight    =   7680
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   10824
   OleObjectBlob   =   "uf_Cov_Compliance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Cov_Compliance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare Integers
Dim intCustRow As Long

' Declare Ranges / "Ranges"
Dim arryHeader() As Variant

Dim col_LOB As Long
Dim col_Cust As Long

Dim col_CovCompl As Long
Dim col_CovCompl_Exp As Long

' Declare Strings
Dim strFullName As String
Dim strTimeStamp As String
Dim strCovCompSelection As String

Dim strCovCompExp_OrigTextCaption As String
Dim strCovCompExp_PlaceholderText As String
Dim strCovCompExp_NewTextCaption As String

Option Explicit
Private Sub UserForm_Initialize()

' Purpose:  To initialize the userform and properly put it on the screen.
' Trigger:  Workbook Open
' Updated:  3/31/2021
' Author:   James Rinaldi

' Change Log:
'       12/9/2020: Intial Creation
'       3/31/2021: Changed the column index to look to borrower name

' ****************************************************************************

Call PublicVariables.Assign_Public_Variables
Call uf_Cov_Compliance.o_02_Assign_Global_Variables

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

    Me.Height = frm_CovComp_Explanation.Top + 21
    
    Me.cmd_Cancel.SetFocus

' -----------
' Update the UserForm with details about the customer
' -----------

    uf_Cov_Compliance.Caption = uf_Cov_Compliance.Caption & " - " & strCustomer
    Me.lbl_Attest_Text.Caption = Replace(Me.lbl_Attest_Text.Caption, "the selected customer", strCustomer)

End Sub
Sub o_02_Assign_Global_Variables()

' Purpose: To Assign all of the Public variables that were declared "above the line".
' Trigger: Called on Initialization
' Updated: 12/9/2020

' Change Log:
'       12/9/2020: Intial Creation
'       3/10/2021: Updated to convert to arryHeader

' ****************************************************************************

' -----------
' Assign your variables
' -----------

    ' Assign Integers

    intCustRow = ActiveCell.Row
               
    ' Assign Ranges / "Ranges"

    'Set rngHeader = wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol))
    arryHeader = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))
        
    col_LOB = fx_Create_Headers("LOB", arryHeader)
    col_Cust = fx_Create_Headers("Customer", arryHeader)
        
    col_CovCompl = fx_Create_Headers("Covenant Compliance", arryHeader)
    col_CovCompl_Exp = fx_Create_Headers("Covenant Compliance Explanation", arryHeader)
    
    ' Assign Strings
    
    strFullName = myFunctions.fx_Name_Reverse()
    strTimeStamp = Format(Now, "m/d/yyyy hh:mm")
    
    strCovCompExp_OrigTextCaption = "Please provide one sentence to explain why the Borrower [COV COMPL STATUS]:"
    strCovCompExp_PlaceholderText = "[COV COMPL STATUS]"

End Sub
Sub o_03_Adjust_UserForm()

' Purpose: To adjust the size and location of buttons on the user form when the Covenant Compliance Explanation is needed.
' Trigger: Called on selecting an Out Of Compliance option.
' Updated: 12/21/2020

' Change Log:
'       12/21/2020: Intial Creation

' ****************************************************************************

' -----------
' Unhide the Cov Compliance Explanation to prompt the user for more detail before recording the attestation.
' -----------

    Me.Height = 405
    Me.lbl_CovComp_ExplanationPrompt.Caption = strCovCompExp_NewTextCaption
    Me.txt_CovComp_Explanation.SetFocus

End Sub

Private Sub cmd_CovComp_InComp_Click()

    strCovCompSelection = "In Compliance"
    
    Call Me.o_1_Record_Attestation
        Unload Me

End Sub
Private Sub cmd_CovComp_OutOfComp_Click()

' -----------
' Update the text
' -----------

    strCovCompSelection = "Out of Compliance"
    strCovCompExp_NewTextCaption = Replace(strCovCompExp_OrigTextCaption, strCovCompExp_PlaceholderText, "is " & strCovCompSelection)
        
' -----------
' Unhide the Cov Compliance Explanation to prompt the user for more detail before recording the attestation.
' -----------

    Call Me.o_03_Adjust_UserForm

End Sub
Private Sub cmd_CovComp_Waived_Click()

' -----------
' Update the text
' -----------

    strCovCompSelection = "Compliance Waived"
    strCovCompExp_NewTextCaption = Replace(strCovCompExp_OrigTextCaption, strCovCompExp_PlaceholderText, "has " & strCovCompSelection)

' -----------
' Unhide the Cov Compliance Explanation to prompt the user for more detail before recording the attestation.
' -----------

    Call Me.o_03_Adjust_UserForm

End Sub
Private Sub cmd_CovComp_UnderForb_Click()

' -----------
' Update the text
' -----------

    strCovCompSelection = "Under Forbearance"
    strCovCompExp_NewTextCaption = Replace(strCovCompExp_OrigTextCaption, strCovCompExp_PlaceholderText, "is " & strCovCompSelection)

' -----------
' Unhide the Cov Compliance Explanation to prompt the user for more detail before recording the attestation.
' -----------

    Call Me.o_03_Adjust_UserForm

End Sub
Private Sub cmd_CovComp_SubmitExplanation_Click()

    If Len(Me.txt_CovComp_Explanation.Value) >= 5 Then
        Call Me.o_2_Record_CovComp_Explanation
        Call Me.o_1_Record_Attestation
        Unload Me
    Else
        MsgBox Title:="Ah Ah Ah", Buttons:=vbExclamation, Prompt:="You need to include the rationale for why you selected " & strCovCompSelection
    End If

End Sub
Private Sub txt_CovComp_Explanation_Change()

'When the user starts to type the Cov Compliance Explanation default to the Submit button instead of Cancel

    Me.cmd_CovComp_SubmitExplanation.Default = True

End Sub
Private Sub cmd_Cancel_Click()

    Unload Me

End Sub
Sub o_1_Record_Attestation()

' Purpose: To record the attestation for Covenant Compliance for each customer.
' Trigger: Various Attestation cmd Buttons
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
        strFullAttestation = strFullName & " (" & strTimeStamp & ")" & " - " & strCovCompSelection
    
    Dim strOldValue As String
        strOldValue = wsData.Cells(intCustRow, col_CovCompl).Value2
    
' -----------
' Add in the attestation
' -----------
    
    'Make the change, but don't trigger the Worksheet Change Event
    Application.EnableEvents = False
        wsData.Cells(intCustRow, col_CovCompl) = strFullAttestation
    Application.EnableEvents = True
    
    'Track the change in the change log
    With wsChangeLog
        .Range("A" & intCurRow_ChangeLog).Value2 = strTimeStamp                                 ' Change Made Data
        .Range("B" & intCurRow_ChangeLog).Value2 = strFullName                                  ' By Who
        .Range("C" & intCurRow_ChangeLog).Value2 = wsData.Cells(intCustRow, col_LOB)         ' LOB
        .Range("D" & intCurRow_ChangeLog).Value2 = wsData.Cells(intCustRow, col_Cust)        ' Customer
        .Range("E" & intCurRow_ChangeLog).Value2 = "Covenant Compliance"                        ' Field Changed
        .Range("F" & intCurRow_ChangeLog).Value2 = strOldValue                                  ' Old Value
        .Range("G" & intCurRow_ChangeLog).Value2 = strFullAttestation                           ' Attestation
        .Range("H" & intCurRow_ChangeLog).Value2 = "Covenant Compliance Attestation"            ' Change Type
        .Range("I" & intCurRow_ChangeLog).Value2 = "Change Log"                                 ' Source
    End With
    
End Sub
Sub o_2_Record_CovComp_Explanation()

' Purpose: To record the explanation given for a "Out of Compliance" or "Compliance Waived" selection.
' Trigger: cmd_CovComp_SubmitExplanation_Click
' Updated: 12/10/2020

' Change Log:
'       12/10/2020: Initial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim strCovCompExplanation
        'strCovCompExplanation = strFullName & " (" & strTimeStamp & ")" & Chr(10) & Chr(10) & Me.txt_CovComp_Explanation
        strCovCompExplanation = Me.txt_CovComp_Explanation
    
' -----------
' Add in the attestation
' -----------
    
    'Make the change, but don't trigger the Worksheet Change Event
    Application.EnableEvents = False
        wsData.Cells(intCustRow, col_CovCompl_Exp) = strCovCompExplanation
    Application.EnableEvents = True
    
End Sub



