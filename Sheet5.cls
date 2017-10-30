VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CB_Invoicing_Click()
Application.ScreenUpdating = False
ActiveSheet.Unprotect

Dim InputData As Long
Dim vVal As Long

InputData = Application.InputBox(Prompt:="Number of Months:", Title:="Invoicing Period", Type:=1)
Run InputData
vVal = 6 + InputData

If InputData <= 0 Then Application.ScreenUpdating = True: Exit Sub
If InputData > 24 Then MsgBox ("Contracts beyond 24 months will be supported in future versions!"): Exit Sub
If InputData > 0 Then Worksheets("Amortised Summary").Range(Columns(7), Columns(30)).Hidden = True _
                     : Worksheets("Amortised Summary").Range(Columns(7), Columns(vVal)).Hidden = False

ActiveSheet.Protect
Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_Activate()

Me.Unprotect
Application.ScreenUpdating = False
    Dim Col As Range
    Dim minWidth As Long
    For Each Col In Columns("A:AG")
    If Col.Hidden = False Then
    minWidth = Col.ColumnWidth

    Col.EntireColumn.AutoFit
    newWidth = Col.ColumnWidth
    If minWidth > newWidth Then
        Col.ColumnWidth = minWidth
    End If
    End If
    Next Col

Application.ScreenUpdating = True
'Me.Protect

End Sub
