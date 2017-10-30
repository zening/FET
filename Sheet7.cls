VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CB_Add_Res_Click()

If ActiveCell.Row >= Sheets("Frame Services").Range("B1").Value Then MsgBox ("Please select Task Item Line - Exiting Routine"): Exit Sub

Application.ScreenUpdating = False
With Sheets("Frame Services")
    .Unprotect
    .Range("A" & ActiveCell.Offset(1).Row).EntireRow.Insert
    .Range("A" & ActiveCell.Row).EntireRow.Copy
    .Range("A" & ActiveCell.Offset(1).Row).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & ActiveCell.Row) = ""
    .Range("B" & ActiveCell.Row & ":C" & ActiveCell.Row) = 0
    .Range("F" & ActiveCell.Row) = "Project Staff"
    .Range("G" & ActiveCell.Row) = ""
    .Range("J" & ActiveCell.Row) = 0
    .Range("N" & ActiveCell.Row & ":O" & ActiveCell.Row) = 0
    .Range("Q" & ActiveCell.Row) = 0.15
    .Range("S" & ActiveCell.Row & ":T" & ActiveCell.Row) = 0
    .Range("V" & ActiveCell.Row) = 0.15
    .Range("X" & ActiveCell.Row & ":Y" & ActiveCell.Row) = 0
    .Range("AA" & ActiveCell.Row) = 0.15
    .Range("AC" & ActiveCell.Row & ":AD" & ActiveCell.Row) = 0
    .Range("AF" & ActiveCell.Row) = 0.15
    .Range("AH" & ActiveCell.Row & ":AI" & ActiveCell.Row) = 0
    .Range("AK" & ActiveCell.Row) = 0.15
    .Range("AM" & ActiveCell.Row & ":AN" & ActiveCell.Row) = 0
    .Range("AR" & ActiveCell.Row) = ""
    .Range("AS" & ActiveCell.Row) = 0
    .Range("AT" & ActiveCell.Row) = 0.15
    .Protect
    .Range("A" & ActiveCell.Row).Select
End With

Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub

Public Sub CB_Remove_Task_Click()

If Sheets("Frame Services").Range("A" & ActiveCell.Row).Value <> "" Then MsgBox "This is a Task - Remove Task from Scope - Exiting Routine": Exit Sub

Application.ScreenUpdating = False
ActiveSheet.Unprotect

ActiveCell.EntireRow.Select
Selection.Delete

ActiveSheet.Protect
Application.ScreenUpdating = True
End Sub

Private Sub PopulatePRButton_Click() 'Resource List Autopopulate Button
Application.ScreenUpdating = False
ActiveSheet.Unprotect
Dim TStRow As Long
Dim TEndRow As Long
Dim RStRow As Long
Dim REndRow As Long
Dim PSSRow As Long
Dim StrRes As String
Dim r As Integer
Dim StrResList As String
Dim X As Integer
Dim t As Integer
Dim PSR As Long

'Defines Ranges for Task List vs. Resource List
TStRow = Sheets("Frame Services").Range("A1").Value 'Task Start Row
TEndRow = Sheets("Frame Services").Range("B1").Value 'Task End Row
RStRow = Sheets("Frame Services").Range("C1").Value 'Resource Start Row
REndRow = Sheets("Frame Services").Range("D1").Value 'Resource End Row
PSSRow = Sheets("Frame Services").Range("E1").Value 'Project Setup Start Row

'Show the formula row
Sheets("Frame Services").Range("A" & RStRow).EntireRow.Hidden = False
     
'Loop 1 = Add New Resources, Loop 2 = Remove Unnecessary Resources
For X = 4 To TEndRow 'Task List Scanner
    StrRes = Sheets("Frame Services").Range("F" & X).Value
    t = Sheets("Frame Services").Range("E" & X).Value
        For r = RStRow To REndRow 'Resource List Scanner
         StrResList = Sheets("Frame Services").Range("B" & r).Value
            If StrRes = "" Then Exit For
            If t = 0 Then Exit For
            If StrRes = "Project Staff" Then Exit For
            If StrRes = "Technical Staff" Then Exit For
            If StrRes = "Consulting Staff" Then Exit For
            If StrRes = "Applications Staff" Then Exit For
            If StrResList = "" Then Exit For
            If StrRes = StrResList Then Exit For
            If StrResList = "Projects" Then Range("B" & r) = StrRes: Exit For
               If StrRes <> StrResList Then r = r
               If r = REndRow And StrRes <> StrResList Then
               Range("A" & RStRow).EntireRow.Select
               ActiveCell.Offset(1).EntireRow.Insert
               Selection.Copy
               ActiveCell.Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
               Application.CutCopyMode = False
               Range("A" & ActiveCell.Row) = "[Personnel]"
               Range("B" & ActiveCell.Row) = StrRes
               Range("L" & ActiveCell.Row).Formula = "=L" & ActiveCell.Offset(1).Row & "-1"
               REndRow = Sheets("Frame Services").Range("D1").Value
               PSR = r + 1
               With Sheets("Project Setup Request")
                   .Range("A" & PSSRow).Offset(1).EntireRow.Insert
                   .Range("A" & PSSRow).EntireRow.Copy
                   .Range("A" & PSSRow).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
                   .Range("A" & PSSRow).Offset(1).EntireRow.Hidden = False
               End With
               Exit For
               End If
        Next r
Next X

For r = RStRow To REndRow
    PSR = Sheets("Frame Services").Range("L" & r).Value
    If r > REndRow Then Sheets("Frame Services").Range("L" & RStRow).Formula = Sheets("Frame Services").Range("E1").Value _
                        : Sheets("Frame Services").Range("B" & RStRow) = "Project Staff" _
                        : Exit For
    If RStRow = REndRow Then Sheets("Frame Services").Range("A" & RStRow).EntireRow.Hidden = True _
                        : Sheets("Project Setup Request").Range("A" & PSR).EntireRow.Hidden = True _
                        : Exit For
    Sheets("Frame Services").Range("B" & r).Select
    StrResList = Sheets("Frame Services").Range("B" & r).Value
    t = Sheets("Frame Services").Range("F" & r).Value
    
    'Hide the formula row and show the row if t > 0
    If t = 0 Then Sheets("Frame Services").Range("A" & r).EntireRow.Hidden = True: Sheets("Project Setup Request").Range("A" & PSR).EntireRow.Hidden = True
    If t > 0 Then Sheets("Frame Services").Range("A" & r).EntireRow.Hidden = False: Sheets("Project Setup Request").Range("A" & PSR).EntireRow.Hidden = False
    REndRow = Sheets("Frame Services").Range("D1").Value
Next r

ActiveSheet.Protect
Application.ScreenUpdating = True
End Sub
    
Private Sub CB_TAHide_Click()
Application.ScreenUpdating = False
Worksheets("Frame Services").Unprotect
Worksheets("Frame Services").Columns("N:AW").EntireColumn.Hidden = True
Worksheets("Frame Services").Protect
CB_TAHide.Visible = False
CB_TAShow.Caption = "Show Travel"
Application.ScreenUpdating = True
End Sub

Private Sub CB_TAShow_Click()
Application.ScreenUpdating = False
Worksheets("Frame Services").Unprotect
'If CB_Expend_BOE.Caption = "+" Then
If (CB_TAShow.Caption = "Show Travel") Then
    Worksheets("Frame Services").Columns("N:AW").EntireColumn.Hidden = False
    CB_TAShow.Caption = "Hide Travel"
    CB_TAHide.Visible = True
Else
    Worksheets("Frame Services").Columns("N:AW").EntireColumn.Hidden = True
    CB_TAHide.Visible = False
    CB_TAShow.Caption = "Show Travel"
End If

Worksheets("Frame Services").Protect
Application.ScreenUpdating = True
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
