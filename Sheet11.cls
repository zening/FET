VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub FormulaFiller() 'Adjust Financial Calcs to Contract Type -> Fills down the list.
For r = 5 To Range("A2").Value
    If r = Range("A2").Value Then Exit For
    
    If CB_TnM = True Then
        Range("AK" & r).Formula = "=((B" & r & "*C" & r & ")+D" & r & ")*AJ" & r
        Range("AN" & r).Formula = "=(((B" & r & "*C" & r & ")+AK" & r & "+D" & r & ")/(100%-AL" & r & "))+AI" & r
    End If
    
    If CB_Fixed = True Then
            Range("AK" & r).Formula = "=((B" & r & "+D" & r & ")*AJ" & r & ")"
            Range("AN" & r).Formula = "=((B" & r & "+AK" & r & "+D" & r & ")/(100%-AL" & r & "))+AI" & r
    End If
    
Next r
End Sub

Private Sub CB_TnM_Click()
Application.ScreenUpdating = False

If CB_TnM = True Then Sheets(11).Unprotect: CB_Fixed = False _
                      : Sheets(11).Columns("C:D").EntireColumn.Hidden = False _
                      : Range("B3") = "Rate": Call FormulaFiller

If CB_TnM = False Then CB_Fixed = True

Application.ScreenUpdating = True
'ActiveSheet.Protect ' -> For some reason results in inexplicable error of Range Class hidden settings.
End Sub

Private Sub CB_Fixed_Click()
Application.ScreenUpdating = False

If CB_Fixed = True Then Sheets(11).Unprotect: CB_TnM = False _
                      : Sheets(11).Columns("C:D").EntireColumn.Hidden = True _
                      : Range("B3") = "Buy": Call FormulaFiller

If CB_Fixed = False Then CB_TnM = True

Application.ScreenUpdating = True
'ActiveSheet.Protect ' -> For some reason results in inexplicable error of Range Class hidden settings.
End Sub

Private Sub CB_Subc_TA_Click()
Application.ScreenUpdating = False

If CB_Subc_TA = True Then Sheets(11).Columns("E:AI").EntireColumn.Hidden = False
If CB_Subc_TA = False Then Sheets(11).Columns("E:AI").EntireColumn.Hidden = True

Application.ScreenUpdating = True
End Sub

Private Sub CB_EnableS_Click()
Application.ScreenUpdating = False
Dim FSR As Long 'Frame Sum Row
Dim FSRE As Long 'Frame Sum Row End
Dim CSR As Long 'Client Sum Row
Dim CSRE As Long 'Client Sum Row End

FSR = Sheets(11).Range("B2").Value
FSRE = Sheets(11).Range("D2").Value
CSR = FSR - 1
CSRE = FSRE - 1

If CB_EnableS = True And Sheets("Document Control").CB_Itemised_Tasks = False Then _
                           : Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                           : Sheets("Frame Summary").Rows(FSR).EntireRow.Hidden = False _
                           : Sheets("Client Summary").Rows(CSR).EntireRow.Hidden = False _
                           : Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                           
If CB_EnableS = True And Sheets("Document Control").CB_Itemised_Tasks = True Then _
                           : Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                           : Sheets("Frame Summary").Rows(FSR & ":" & FSRE).EntireRow.Hidden = False _
                           : Sheets("Client Summary").Rows(CSR & ":" & CSRE).EntireRow.Hidden = False _
                           : Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                           
If CB_EnableS = False Then Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                           : Sheets("Frame Summary").Rows(FSR & ":" & FSRE).EntireRow.Hidden = True _
                           : Sheets("Client Summary").Rows(CSR & ":" & CSRE).EntireRow.Hidden = True _
                           : Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
 
Application.ScreenUpdating = True
End Sub
