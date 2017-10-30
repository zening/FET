VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Dim r As Integer

Application.ScreenUpdating = False

For f = 20 To 400 'Loop till the cell found
    
    If Range("A" & f).Value = "Frame Services" Then Exit For

Next f
r = f + 2

Dim i As Integer
i = 0
For X = 19 To 37 Step 2 'scan Frame Summary page for sub-contractors

    If Sheets("Frame Summary").Range("A" & X).EntireRow.Hidden = False Then
        Sheets("Quote").Range("A" & r + i).EntireRow.Hidden = False
    Else
        Sheets("Quote").Range("A" & r + i).EntireRow.Hidden = True
    End If
    i = i + 1

Next X
Application.ScreenUpdating = True

End Sub


