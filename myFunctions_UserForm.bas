Attribute VB_Name = "myFunctions_UserForm"
Option Explicit
Function fxUserForm_ZoomIn(objTargetUserForm As Object)

' Purpose: To allow the user to zoom in on a UserForm.
' Trigger: Called by UserForm
' Updated: 10/7/2022

' Change Log:
'       10/7/2022:  Initial Creation

' ********************************************************************************************************************************************************

    ' Adjust the size of the UserForm
    objTargetUserForm.Height = objTargetUserForm.Height * 1.1
    objTargetUserForm.Width = objTargetUserForm.Width * 1.1
    
    ' Adjust the zoom
    objTargetUserForm.Zoom = objTargetUserForm.Zoom * 1.1

    ' Recenter the form
    objTargetUserForm.Top = Application.Top + (Application.UsableHeight / 1.5) - (objTargetUserForm.Height / 2)
    objTargetUserForm.Left = Application.Left + (Application.UsableWidth / 2) - (objTargetUserForm.Width / 2)

End Function
Function fxUserForm_ZoomOut(objTargetUserForm As Object)

' Purpose: To allow the user to zoom out on a UserForm.
' Trigger: Called by UserForm
' Updated: 10/7/2022

' Change Log:
'       10/7/2022:  Initial Creation

' Note:
'       It's actually .909, not .90 to get from 110% back down to 100% (or close enough)

' ********************************************************************************************************************************************************

    ' Adjust the size of the UserForm
    objTargetUserForm.Height = objTargetUserForm.Height * 0.909
    objTargetUserForm.Width = objTargetUserForm.Width * 0.909
    
    ' Adjust the zoom
    objTargetUserForm.Zoom = objTargetUserForm.Zoom * 0.909

    ' Recenter the form
    objTargetUserForm.Top = Application.Top + (Application.UsableHeight / 1.5) - (objTargetUserForm.Height / 2)
    objTargetUserForm.Left = Application.Left + (Application.UsableWidth / 2) - (objTargetUserForm.Width / 2)

End Function
