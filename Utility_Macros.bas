Attribute VB_Name = "Utility_Macros"
Option Explicit
Sub u_Delete_Hidden_Sheets()

' Purpose: To delete all of the hidden Worksheets in the Sageworks Dashboard Workbook.
' Trigger: Manual
' Updated: 2/1/2021

' ****************************************************************************

    Dim ws As Worksheet
    
    Debug.Print ActiveWorkbook.Name
    
    Application.DisplayAlerts = False
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetHidden Then ws.Delete
    Next ws
    
    Application.DisplayAlerts = True

End Sub
Sub u_Remove_Our_Cell_Fill()

' Purpose: To remove all of the cell fill except the greys.
' Trigger: Manual
' Updated: 2/1/2021

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    ' Dim Loop

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("Dashboard Review")

    Dim cell As Variant

    ' Dim Colors
    Dim clrLightGray As Long
        clrLightGray = RGB(217, 217, 217)

    Dim clrDarkGray As Long
        clrDarkGray = RGB(64, 64, 64)

    ' Dim Integers

    Dim intLastRow As Long
        intLastRow = 669
        
    Dim intLastCol As Long
        intLastCol = 47

' -----------
' Remove the colors
' -----------

    With wsData
        For Each cell In .Range(.Cells(2, 1), .Cells(intLastRow, intLastCol))
        
            If cell.Interior.Color <> clrLightGray And cell.Interior.Color <> clrDarkGray Then
                cell.Interior.Color = xlNone
            End If
    
        Next cell
    End With


End Sub
Sub u_Wipe_Checklist()

' Purpose: To remove the X marks in the Checklist ws.
' Trigger: Button
' Updated: 8/10/2021

' Change Log:
'       8/10/2021: Initial Creation

' ****************************************************************************

myPrivateMacros.DisableForEfficiency

    ' Assign Worksheets
    
    Dim wsChecklist As Worksheet
    Set wsChecklist = ThisWorkbook.Sheets("CHECKLIST")

    ' Declare Integers
       
    Dim intLastRow As Long
        intLastRow = wsChecklist.Cells(Rows.count, "C").End(xlUp).Row

    Dim i As Long
    
' ------------
' Wipe the X's
' ------------

    With wsChecklist
                
        For i = 2 To intLastRow
            If .Range("C" & i).Value = "X" Then .Range("C" & i).Value = ""
        Next i
        
    End With

myPrivateMacros.DisableForEfficiencyOff

End Sub
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Long
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    directory = ActiveWorkbook.path & "\VisualBasic"
    count = 0
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
End Sub


