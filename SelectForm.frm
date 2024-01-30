VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectForm 
   Caption         =   "UserForm Selector"
   ClientHeight    =   2784
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6912
   OleObjectBlob   =   "SelectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub UserForm_Initialize()

' Purpose:  To initialize the userform and properly put it on the screen.
' Trigger:  Workbook Open
' Updated:  12/3/2020
' Author:   James Rinaldi

' Change Log:
'   12/3/2020: Intial Creation

' ****************************************************************************

' -----------
' Initialize the initial values
' -----------
    
    Me.StartUpPosition = 0 'Allow you to set the position
        Me.Top = Application.Top + (Application.UsableHeight / 2) - (Me.Height / 2)
        Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)

Me.cmd_SCE.SetFocus

End Sub
Private Sub cmd_Cancel_Click()

    Unload Me

End Sub
Private Sub cmd_Admin_Click()

    Unload Me
    strUserFormSelection = "Admin"

End Sub
Private Sub cmd_SCE_Click()

    Unload Me
    strUserFormSelection = "SCE"

End Sub
Private Sub cmd_Regular_Click()

    Unload Me
    strUserFormSelection = "Regular"
    
End Sub
