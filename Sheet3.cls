VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()

Me.Unprotect
Application.ScreenUpdating = False
    Dim Col As Range
    Dim minWidth As Long
    For Each Col In Columns("A:Z")
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
