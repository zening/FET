VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectSetupRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    'Check if the Target Address is same as you have given
    'In the above example i have taken A4 Cell, so I am
    'Comparing this with $A$4
    Range(Target.Range.Address).Cells.Select
    'MsgBox (Range(Target.Range.Address).Row)
    If Range(Target.Range.Address).Value > 1 Then
        SplitRow (Range(Target.Range.Address).Row)
    End If
End Sub

Private Sub SplitRow(r)
    Dim msg As String
    Dim splitValue As String
    Dim strArray() As String
    Dim intCount As Integer
    Dim BudgetHours As Integer
    BudgetHours = 0

    splitValue = InputBox("The Budget Hours is " & Range("H" & r).Value & ", You may split it with comma separated values: (ie. x,x )")
    If splitValue = vbNullString Then
        MsgBox ("User canceled!")
        Exit Sub
    End If
    strArray = Split(splitValue, ",")
    
    For intCount = LBound(strArray) To UBound(strArray)
      BudgetHours = BudgetHours + (Trim(strArray(intCount)))
    Next
    If BudgetHours <> Range("H" & r).Value Then
        MsgBox ("Total hours must be the same as the Budget Hours! ")
        Exit Sub
    Else
        'msg = Range("H" & r).Value & " " & splitValue
        'MsgBox (BudgetHours)
        For intCount = LBound(strArray) To UBound(strArray)
         '   BudgetHours = BudgetHours + (Trim(strArray(intCount)))
            Range("H" & r).Offset(intCount + 1).EntireRow.Insert
            Range("H" & r).Offset(intCount + 1).Value = Trim(strArray(intCount))
            Range("I" & r).Offset(intCount + 1).Formula = "=H" & r + intCount + 1 & "*F" & r
            Range("J" & r).Offset(intCount + 1).Formula = "=H" & r + intCount + 1 & "*G" & r
        Next
        'remove the bugdet hours
        Range("K" & r).Value = Range("H" & r).Value
        Range("H" & r).Value = ""
        Range("I" & r).Value = ""
        Range("J" & r).Value = ""
    End If
    
End Sub



