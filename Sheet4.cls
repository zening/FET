VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CB_Expend_BOE_Click()
Me.Unprotect
Application.ScreenUpdating = False
Dim DES As Integer
Dim DErows As Integer
Dim BES As Integer
Dim BEE As Integer

Dim i As Integer
DES = Range("G3").Value + 1
DErows = Range("N" & DES + 1).Value
BES = Range("G3").Value + 1 + DErows + 3 'BOE block start here
BEE = BES + Range("N" & BES + 1).Value + 1

Rows(BEE & ":" & BEE + 4).EntireRow.Insert
Range("A" & BEE - 1).EntireRow.Copy
For i = 0 To 4
Range("A" & BEE + i).EntireRow.PasteSpecial xlPasteAllUsingSourceTheme
Range("B" & BEE + i).Value = ""
Next i

Range("N" & BES + 1).Value = Range("N" & BES + 1).Value + 5
Range("B" & BEE).Select
CB_Shrink_BOE.Enabled = True
Application.ScreenUpdating = True
Me.Protect
End Sub

Private Sub CB_Expend_Del_Click()
Me.Unprotect
Application.ScreenUpdating = False
Dim DES As Integer
Dim DEE As Integer

Dim i As Integer
DES = Range("G3").Value + 1
DEE = DES + Range("N" & DES + 1).Value + 1

Rows(DEE & ":" & DEE + 4).EntireRow.Insert
Range("A" & DEE - 1).EntireRow.Copy
For i = 0 To 4
Range("A" & DEE + i).EntireRow.PasteSpecial xlPasteAllUsingSourceTheme
Range("B" & DEE + i).Value = ""
Next i

Range("N" & DES + 1).Value = Range("N" & DES + 1).Value + 5
Range("B" & DEE).Select
CB_Shrink_Del.Enabled = True
Application.ScreenUpdating = True
Me.Protect

End Sub

Private Sub CB_Expend_GB_Click()
Application.ScreenUpdating = False
If CB_Expend_GB.Caption = "+" Then
With Sheets("Scope")
    .Unprotect
    .Rows((Sheets("Scope").Range("A3").Value + 6) & ":" & Sheets("Scope").Range("A3").Value + 13).Hidden = False
    .Protect
End With
CB_Expend_GB.Caption = "-"
Else
With Sheets("Scope")
    .Unprotect
    .Rows((Sheets("Scope").Range("A3").Value + 6) & ":" & Sheets("Scope").Range("A3").Value + 13).Hidden = True
    .Protect
End With
CB_Expend_GB.Caption = "+"
End If

Application.ScreenUpdating = True
End Sub

Private Sub CB_Shrink_BOE_Click()
Me.Unprotect

Dim DES As Integer
Dim DErows As Integer
Dim BES As Integer
Dim BEE As Integer

Dim i As Integer
DES = Range("G3").Value + 1
DErows = Range("N" & DES + 1).Value
BES = Range("G3").Value + 1 + DErows + 3 'BOE block start here
BEE = Range("N" & BES + 1).Value

Rows(BES + BEE).EntireRow.Delete
Range("N" & BES + 1).Value = Range("N" & BES + 1).Value - 1
Range("B" & BES + BEE - 1).Select

If Range("N" & BES + 1).Value < 6 Then
    CB_Shrink_BOE.Enabled = False
End If
Me.Protect
End Sub

Private Sub CB_Shrink_Del_Click()
Me.Unprotect

Dim DES As Integer
Dim DEE As Integer

Dim i As Integer
DES = Range("G3").Value + 1
DEE = Range("N" & DES + 1).Value
Rows(DES + DEE).EntireRow.Delete
Range("N" & DES + 1).Value = Range("N" & DES + 1).Value - 1
Range("B" & DES + DEE - 1).Select

If Range("N" & DES + 1).Value < 6 Then
    CB_Shrink_Del.Enabled = False
End If
Me.Protect

End Sub

Private Sub CB_Shrink_GB_Click()
Dim GBE As Integer
Dim i As Integer
GBE = Range("N6").Value
Rows(GBE - 1).EntireRow.Delete
Range("N6") = GBE - 1
Range("B" & GBE - 2).Select

If Range("N6").Value < 13 Then
    CB_Shrink_GB.Enabled = False
End If
End Sub

'Subcontractor Change Tracker
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("A" & Range("A4").Value, "C" & Range("A4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(9).Name = Range("A" & Range("A4").Value).Value
        If Range("C" & Range("A4").Value).Value = "Fixed Price" Then Sheets(9).CB_Fixed = True
        If Range("C" & Range("A4").Value).Value = "Time & Materials" Then Sheets(9).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
    If Not Intersect(Target, Range("A" & Range("B4").Value, "C" & Range("B4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(10).Name = Range("A" & Range("B4").Value).Value
        If Range("C" & Range("B4").Value).Value = "Fixed Price" Then Sheets(10).CB_Fixed = True
        If Range("C" & Range("B4").Value).Value = "Time & Materials" Then Sheets(10).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
    If Not Intersect(Target, Range("A" & Range("C4").Value, "C" & Range("C4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(11).Name = Range("A" & Range("C4").Value).Value
        If Range("C" & Range("C4").Value).Value = "Fixed Price" Then Sheets(11).CB_Fixed = True
        If Range("C" & Range("C4").Value).Value = "Time & Materials" Then Sheets(11).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
    If Not Intersect(Target, Range("A" & Range("D4").Value, "C" & Range("D4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(12).Name = Range("A" & Range("D4").Value).Value
        If Range("C" & Range("D4").Value).Value = "Fixed Price" Then Sheets(12).CB_Fixed = True
        If Range("C" & Range("D4").Value).Value = "Time & Materials" Then Sheets(12).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
    If Not Intersect(Target, Range("A" & Range("E4").Value, "C" & Range("E4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(13).Name = Range("A" & Range("E4").Value).Value
        If Range("C" & Range("E4").Value).Value = "Fixed Price" Then Sheets(13).CB_Fixed = True
        If Range("C" & Range("E4").Value).Value = "Time & Materials" Then Sheets(13).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
    If Not Intersect(Target, Range("A" & Range("F4").Value, "C" & Range("F4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(14).Name = Range("A" & Range("F4").Value).Value
        If Range("C" & Range("F4").Value).Value = "Fixed Price" Then Sheets(14).CB_Fixed = True
        If Range("C" & Range("F4").Value).Value = "Time & Materials" Then Sheets(14).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
    If Not Intersect(Target, Range("A" & Range("G4").Value, "C" & Range("G4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(15).Name = Range("A" & Range("G4").Value).Value
        If Range("C" & Range("G4").Value).Value = "Fixed Price" Then Sheets(15).CB_Fixed = True
        If Range("C" & Range("G4").Value).Value = "Time & Materials" Then Sheets(15).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
    If Not Intersect(Target, Range("A" & Range("H4").Value, "C" & Range("H4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(16).Name = Range("A" & Range("H4").Value).Value
        If Range("C" & Range("H4").Value).Value = "Fixed Price" Then Sheets(16).CB_Fixed = True
        If Range("C" & Range("H4").Value).Value = "Time & Materials" Then Sheets(16).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
    If Not Intersect(Target, Range("A" & Range("I4").Value, "C" & Range("I4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(17).Name = Range("A" & Range("I4").Value).Value
        If Range("C" & Range("I4").Value).Value = "Fixed Price" Then Sheets(17).CB_Fixed = True
        If Range("C" & Range("I4").Value).Value = "Time & Materials" Then Sheets(17).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
    If Not Intersect(Target, Range("A" & Range("J4").Value, "C" & Range("J4").Value)) Is Nothing Then
    Application.ScreenUpdating = False
        Sheets(18).Name = Range("A" & Range("J4").Value).Value
        If Range("C" & Range("J4").Value).Value = "Fixed Price" Then Sheets(18).CB_Fixed = True
        If Range("C" & Range("J4").Value).Value = "Time & Materials" Then Sheets(18).CB_TnM = True
    Application.ScreenUpdating = True
    End If
    
End Sub


Public Sub CB_Add_Scope_Click() 'Scope Controls

'Row Checks: Adding lines between Frame & Subc Scopes / Not using Segment Title
If ActiveCell.Row < Range("C3").Value Then MsgBox ("Please Select Task Row - Exiting Routine."): Exit Sub
If ActiveCell.Row > Range("F3").Value Then MsgBox ("Please Select Task Row - Exiting Routine."): Exit Sub
If Range("G" & ActiveCell.Row).Value <> "" Then MsgBox ("Please Select Task Row - Exiting Routine."): Exit Sub

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.Unprotect

Dim FSI As Long 'Frame Services Insert
Dim FSV As Long 'Frame Services
Dim FSR As Long 'Frame Summary Row
Dim Subc As String
Dim ItemRow As Long

'Insert row below current line, copy formula's to mimic the 'Fill Down' effect
With ActiveCell
    .Offset(1).EntireRow.Insert
    .EntireRow.Copy
    .Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
End With

ItemRow = ActiveCell.Row + 1

Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect: Sheets("Amortised Summary").Unprotect

'If in Frame Scope - Add Scope line to Frame Services
If ActiveCell.Row > Range("B3").Value And ActiveCell.Row < Range("D3").Value Then
    Range("A" & ActiveCell.Row) = "[Task]"
    Range("J" & ActiveCell.Row) = ""
    ItemRow = ActiveCell.Row
    FSI = Range("O" & ActiveCell.Row).Offset(1).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    With Sheets("Frame Services")
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSV).Formula = "=Scope!A" & ItemRow
        .Range("B" & FSV & ":C" & FSV).Value = 0
        .Range("F" & FSV) = "Project Staff"
        .Range("G" & FSV) = ""
        .Range("J" & FSV) = 0
        .Range("N" & FSV) = 0
        .Range("O" & FSV) = 0
        .Range("Q" & FSV) = 0.15
        .Range("S" & FSV) = 0
        .Range("T" & FSV) = 0
        .Range("V" & FSV) = 0.15
        .Range("X" & FSV) = 0
        .Range("Y" & FSV) = 0
        .Range("AA" & FSV) = 0.15
        .Range("AC" & FSV) = 0
        .Range("AD" & FSV) = 0
        .Range("AF" & FSV) = 0.15
        .Range("AH" & FSV) = 0
        .Range("AI" & FSV) = 0
        .Range("AK" & FSV) = 0.15
        .Range("AM" & FSV) = 0
        .Range("AN" & FSV) = 0
        .Range("AP" & FSV) = 0.15
        .Range("AR" & FSV) = ""
        .Range("AS" & FSV) = 0
        .Range("AT" & FSV) = 0.15
        .Protect
    End With
    
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='Frame Services'!A" & FSV
        .Range("C" & FSR).Formula = "=SUM('Frame Services'!K" & FSV _
         & ":INDIRECT(""'Frame Services'!K""&(ROW('Frame Services'!K" & (FSV + 1) _
         & ")-1)))"
        .Range("D" & FSR).Formula = "=SUM('Frame Services'!B" & FSV _
         & ":INDIRECT(""'Frame Services'!B""&(ROW('Frame Services'!B" & (FSV + 1) _
         & ")-1)))"
        .Range("E" & FSR).Formula = "=SUM('Frame Services'!AV" & FSV _
         & ":INDIRECT(""'Frame Services'!AV""&(ROW('Frame Services'!AV" & (FSV + 1) _
         & ")-1)))"
        .Range("F" & FSR).Formula = "=SUM('Frame Services'!AX" & FSV _
         & ":INDIRECT(""'Frame Services'!AX""&(ROW('Frame Services'!AX" & (FSV + 1) _
         & ")-1)))"
        .Range("C" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!K" & Sheets("Scope").Range("O" & ItemRow).Offset(-1).Value _
        & ":INDIRECT(""'Frame Services'!K""&(ROW('Frame Services'!K" & FSV & ")-1)))"
        .Range("D" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!B" & Sheets("Scope").Range("O" & ItemRow).Offset(-1).Value _
        & ":INDIRECT(""'Frame Services'!B""&(ROW('Frame Services'!B" & FSV & ")-1)))"
        .Range("E" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!AV" & Sheets("Scope").Range("O" & ItemRow).Offset(-1).Value _
        & ":INDIRECT(""'Frame Services'!AV""&(ROW('Frame Services'!AV" & FSV & ")-1)))"
        .Range("F" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!AX" & Sheets("Scope").Range("O" & ItemRow).Offset(-1).Value _
        & ":INDIRECT(""'Frame Services'!AX""&(ROW('Frame Services'!AX" & FSV & ")-1)))"
    End With
    
    'If this is the second task in a segment, the first's formulae will be skewed to cover both. Routine fixes first lines formulae
    If Sheets("Frame Services").Range("C" & FSV).Offset(-3).Value = "%" Or Sheets("Frame Services").Range("C" & FSV).Offset(-3).Value = "" Then
        With Sheets("Frame Summary")
        .Range("C" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!K" & Sheets("Frame Services").Range("C" & FSV).Offset(-2).Value + 1 _
         & ":INDIRECT(""'Frame Services'!K""&(ROW('Frame Services'!K" & FSV _
         & ")-1)))"
        .Range("D" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!B" & Sheets("Frame Services").Range("C" & FSV).Offset(-2).Value + 1 _
         & ":INDIRECT(""'Frame Services'!B""&(ROW('Frame Services'!B" & FSV _
         & ")-1)))"
        .Range("E" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!AV" & Sheets("Frame Services").Range("C" & FSV).Offset(-2).Value + 1 _
         & ":INDIRECT(""'Frame Services'!AV""&(ROW('Frame Services'!AV" & FSV _
         & ")-1)))"
        .Range("F" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!AX" & Sheets("Frame Services").Range("C" & FSV).Offset(-2).Value + 1 _
         & ":INDIRECT(""'Frame Services'!AX""&(ROW('Frame Services'!AX" & FSV _
         & ")-1)))"
         End With
    End If
    Sheets("Frame Summary").Protect
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ItemRow).Formula = "=Row('Frame Services'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ItemRow).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

'If in Subc Scope - Add Scope line to Relevant Subc
'Scope: Subc 1 - Sheet(9)
If ActiveCell.Row > Range("A4").Value And ActiveCell.Row < Range("B4").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("A4").Value).Value
    With Sheets(9)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

'Scope: Subc 2 - Sheet(10)
If ActiveCell.Row > Range("B4").Value And ActiveCell.Row < Range("C4").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("B4").Value).Value
    With Sheets(10)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

'Scope: Subc 3 - Sheet(11)
If ActiveCell.Row > Range("C4").Value And ActiveCell.Row < Range("D4").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("C4").Value).Value
    With Sheets(11)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

'Scope: Subc 4 - Sheet(12)
If ActiveCell.Row > Range("D4").Value And ActiveCell.Row < Range("E4").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("D4").Value).Value
    With Sheets(12)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

'Scope: Subc 5 - Sheet(13)
If ActiveCell.Row > Range("E4").Value And ActiveCell.Row < Range("F4").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("E4").Value).Value
    With Sheets(13)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

'Scope: Subc 6 - Sheet(14)
If ActiveCell.Row > Range("F4").Value And ActiveCell.Row < Range("G4").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("F4").Value).Value
    With Sheets(14)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If
    
'Scope: Subc 7 - Sheet(15)
If ActiveCell.Row > Range("G4").Value And ActiveCell.Row < Range("H4").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("G4").Value).Value
    With Sheets(15)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

'Scope: Subc 8 - Sheet(16)
If ActiveCell.Row > Range("H4").Value And ActiveCell.Row < Range("I4").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("H4").Value).Value
    With Sheets(16)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

'Scope: Subc 9 - Sheet(17)
If ActiveCell.Row > Range("I4").Value And ActiveCell.Row < Range("J4").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("I4").Value).Value
    With Sheets(17)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

'Scope: Subc 10 - Sheet(18)
If ActiveCell.Row > Range("J4").Value And ActiveCell.Row < Range("F3").Value Then
    Range("J" & ActiveCell.Row) = ""
    Range("A" & ActiveCell.Row) = "[Task]"
    FSI = Range("O" & ActiveCell.Row).Value - 1
    FSV = FSI + 1
    FSR = Range("P" & ActiveCell.Row).Value
    Subc = Range("A" & Range("J4").Value).Value
    With Sheets(18)
        .Unprotect
        .Rows(FSI).Offset(1).EntireRow.Insert
        .Rows(FSI).Copy
        .Rows(FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("B" & FSV & ":" & "F" & FSV) = 0
        .Range("H" & FSV) = 0.15
        .Range("J" & FSV & ":" & "K" & FSV) = 0
        .Range("M" & FSV) = 0.15
        .Range("O" & FSV & ":" & "P" & FSV) = 0
        .Range("R" & FSV) = 0.15
        .Range("T" & FSV & ":" & "U" & FSV) = 0
        .Range("W" & FSV) = 0.15
        .Range("Y" & FSV & ":" & "Z" & FSV) = 0
        .Range("AB" & FSV) = 0.15
        .Range("AD" & FSV & ":" & "AE" & FSV) = 0
        .Range("AG" & FSV) = 0.15
        .Range("AJ" & FSV) = 0
        .Range("AL" & FSV) = 0
        '.Protect
    End With
        
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Insert
        .Rows(FSR).Offset(-1).Copy
        .Rows(FSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Range("A" & FSR).Formula = "='" & Subc & "'!A" & FSV
        .Protect
    End With
    
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Protect
    End With
    
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Insert
        .Rows(FSR).Offset(-2).Copy
        .Rows(FSR).Offset(-1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    End With
    
    Sheets("Scope").Range("O" & ActiveCell.Row).Formula = "=Row('" & Subc & "'!A" & FSV & ")"
    Sheets("Scope").Range("P" & ActiveCell.Row).Formula = "=Row('Frame Summary'!A" & FSR & ")"
End If

Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect: Sheets("Amortised Summary").Protect

Sheets("Scope").Protect
Application.CutCopyMode = False
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub

Private Sub CB_AddMultiple_Click()
Dim InputData As Variant

InputData = Application.InputBox(Prompt:="Number of Scope Lines:", Title:="Scope List", Type:=1)
If InputData <= 0 Then Exit Sub

For r = 1 To InputData
Call Sheets("Scope").CB_Add_Scope_Click
Next r

End Sub

Public Sub ServiceSummaryUnhide() 'Unhides Segment Titles on Frame / Client Summaries

For r = Sheets("Scope").Range("C3").Value To Sheets("Scope").Range("F3").Value 'Frame Scope Start to End
    If Sheets("Scope").Range("G" & r).Value = "" Then r = r
    If Sheets("Scope").Range("G" & r).Value <> "" And Sheets("Scope").Rows(r).Hidden = False Then Sheets("Frame Summary").Rows(Sheets("Scope").Range("G" & r).Value).Hidden = False _
                                                     : Sheets("Client Summary").Rows(Sheets("Scope").Range("G" & r).Value - 1).Hidden = False
Next r

End Sub

Public Sub UnusedHide()

For r = Sheets("Scope").Range("C3").Value To Sheets("Scope").Range("F3").Value 'Frame Scope Start to End
    If Sheets("Scope").Range("P" & r).Value <> "" And Sheets("Scope").Rows(r).Hidden = True Then Sheets("Frame Summary").Rows(Sheets("Scope").Range("P" & r).Value).Hidden = True _
                                    : Sheets("Client Summary").Rows(Sheets("Scope").Range("P" & r).Value - 1).Hidden = True
Next r

End Sub

Private Sub CB_Expand_Click()
Application.ScreenUpdating = False
With Sheets("Scope")
    .Unprotect
    .Rows((Sheets("Scope").Range("A3").Value + 6) & ":" & Sheets("Scope").Range("A3").Value + 13).Hidden = False
    .Rows((Sheets("Scope").Range("G3").Value + 7) & ":" & Sheets("Scope").Range("G3").Value + 14).Hidden = False
    .Rows((Sheets("Scope").Range("H3").Value - 7) & ":" & Sheets("Scope").Range("H3").Value - 1).Hidden = False
    .Protect
End With
Application.ScreenUpdating = True
End Sub

Public Sub CB_Cut_Scope_Click()

Dim RVT As Long 'Row Value Tester
Dim ItemRow As Long
Dim FSS As Long 'Frame Services Task Start Row
Dim FSE As Long 'Frame Services Task End Row
Dim FSR As Long 'Frame Summary Row
Dim FSV As Long

'Row Checks: Cutting lines between Frame & Subc Scopes / Not using Segment Title / Not only Task
If ActiveCell.Row < Range("C3").Value Then MsgBox ("Please Select Task Row - Exiting Routine."): Exit Sub
If ActiveCell.Row > Range("F3").Value Then MsgBox ("Please Select Task Row - Exiting Routine."): Exit Sub
If Range("G" & ActiveCell.Row).Value <> "" Then MsgBox ("Please Select Task Row - Exiting Routine."): Exit Sub
If Sheets("Scope").Range("G" & ActiveCell.Offset(-1).Row).Value <> "" And Range("A" & ActiveCell.Offset(1).Row).Value = "" Then MsgBox "This is the only Task - Remove Segment or Disable Subcontractor - Exiting Routine": Exit Sub

Application.EnableEvents = False
Application.ScreenUpdating = False
ItemRow = ActiveCell.Row
FSS = Range("O" & ActiveCell.Row).Value 'Determines Task Start Row
FSE = Range("O" & ActiveCell.Row).Offset(1).Value - 1 'Determines Task End Row
FSR = Range("P" & ActiveCell.Row).Value
FSV = Range("O" & ActiveCell.Row).Offset(-1).Value

Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect: Sheets("Amortised Summary").Unprotect

'If in Frame Scope:
If ActiveCell.Row > Range("B3").Value And ActiveCell.Row < Range("D3").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets("Frame Services")
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        .Protect
    End With
    If IsError(Sheets("Frame Summary").Range("C" & FSR).Offset(-1)) Then
        With Sheets("Frame Summary")
            .Unprotect
            .Range("C" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!K" & Sheets("Scope").Range("O" & ItemRow).Offset(-1).Value _
            & ":INDIRECT(""'Frame Services'!K""&(ROW('Frame Services'!K" & FSS & ")-1)))"
            .Range("D" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!B" & Sheets("Scope").Range("O" & ItemRow).Offset(-1).Value _
            & ":INDIRECT(""'Frame Services'!B""&(ROW('Frame Services'!B" & FSS & ")-1)))"
            .Range("E" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!AV" & Sheets("Scope").Range("O" & ItemRow).Offset(-1).Value _
            & ":INDIRECT(""'Frame Services'!AV""&(ROW('Frame Services'!AV" & FSS & ")-1)))"
            .Range("F" & FSR).Offset(-1).Formula = "=SUM('Frame Services'!AX" & Sheets("Scope").Range("O" & ItemRow).Offset(-1).Value _
            & ":INDIRECT(""'Frame Services'!AX""&(ROW('Frame Services'!AX" & FSS & ")-1)))"
            .Protect
        End With
    End If
End If

'If in Subc 1 Scope:
If ActiveCell.Row > Range("A4").Value And ActiveCell.Row < Range("B4").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(9)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

'If in Subc 2 Scope:
If ActiveCell.Row > Range("B4").Value And ActiveCell.Row < Range("C4").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(10)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

'If in Subc 3 Scope:
If ActiveCell.Row > Range("C4").Value And ActiveCell.Row < Range("D4").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(11)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

'If in Subc 4 Scope:
If ActiveCell.Row > Range("D4").Value And ActiveCell.Row < Range("E4").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(12)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

'If in Subc 5 Scope:
If ActiveCell.Row > Range("E4").Value And ActiveCell.Row < Range("F4").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(13)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

'If in Subc 6 Scope:
If ActiveCell.Row > Range("F4").Value And ActiveCell.Row < Range("G4").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(14)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

'If in Subc 7 Scope:
If ActiveCell.Row > Range("G4").Value And ActiveCell.Row < Range("H4").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(15)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

'If in Subc 8 Scope:
If ActiveCell.Row > Range("H4").Value And ActiveCell.Row < Range("I4").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(16)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

'If in Subc 9 Scope:
If ActiveCell.Row > Range("I4").Value And ActiveCell.Row < Range("J4").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(17)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

'If in Subc 10 Scope:
If ActiveCell.Row > Range("J4").Value And ActiveCell.Row < Range("F3").Value Then
    With Sheets("Amortised Summary")
        .Rows(FSR).Offset(-1).Delete
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(FSR).Offset(-1).Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(FSR).Delete
        .Protect
    End With
    With Sheets(18)
        .Unprotect
        .Rows(FSS & ":" & FSE).Delete 'Captures any additional resources
        '.Protect
    End With
End If

Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect: Sheets("Amortised Summary").Protect

With Sheets("Scope")
    .Unprotect
    .Rows(ItemRow).Delete
    .Protect
    .Range("A" & ItemRow).Offset(-1).Select
End With

Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub

Private Sub PhaseButton_Click() 'Creates new phase in Client Sum / Frame Sum / Scope / Relevant Services
If MsgBox("This will insert a new segment below the current selection: Continue?", vbYesNo, "Change Log") = vbYes Then
Application.EnableEvents = False

Dim FSRow As Long
Dim FSI As Long 'Frame Summary Insert Row
Dim CSI As Long 'Client Summary Insert Row
Dim SRSt As Long 'Segment Row Starting
Dim IRSt As Long 'Insertion Row Starting
Dim SREn As Long 'Segment Row Ending
Dim TaskRow1 As Long
Dim TaskRow2 As Long
Dim RVT As Long 'Row Value Tester

If ActiveCell.Row < Range("C3").Value Then MsgBox ("Please Select Frame Scope Rows - Exiting Routine."): Exit Sub
If ActiveCell.Row > Range("D3").Value Then MsgBox ("Please Select Frame Scope Rows - Exiting Routine."): Exit Sub
If Range("H" & ActiveCell.Row).Value <> "" Then MsgBox ("Please Select Task Row - Exiting Routine."): Exit Sub
    
Sheets("Scope").Unprotect
Application.ScreenUpdating = False

For X = ActiveCell.Row To 400 'Loop down till populated line -> Store Location Values -> Generate Line Variables
    Sheets("Scope").Range("A" & X).EntireRow.Select
    SRSt = Sheets("Scope").Range("H" & X).Value
    If ActiveCell.Row > Range("D3").Value Then MsgBox ("Add Segments Before Project Close - Exiting Routine."): Exit Sub
    If SRSt = "0" Then X = X Else FSRow = Sheets("Scope").Range("G" & X).Value _
                                 : IRSt = Sheets("Scope").Range("O" & X).Value: Exit For
    
Next X

SREn = SRSt + 2
FSI = FSRow - 1
CSI = FSI
TaskRow1 = IRSt + 1
TaskRow2 = TaskRow1 + 1

With Sheets("Scope")
    .Rows(SRSt).Offset(-1).Select 'Bottom Row of Current Segment
    .Range("A" & SRSt).EntireRow.Insert '[Blank Row]
    .Range("A" & SRSt).EntireRow.Insert '[Title Row]
    .Range("A" & SRSt).EntireRow.Insert '[Task Row]
    '[Title Row]
    .Range("A" & ActiveCell.Offset(4).Row & ":AZ" & ActiveCell.Offset(4).Row).Copy
    .Range("A" & ActiveCell.Offset(1).Row).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & ActiveCell.Row) = "[Segment]"
    .Range("G" & ActiveCell.Row).Formula = "=ROW('Frame Summary'!A" & FSRow & ")"
    .Range("H" & ActiveCell.Row).Formula = "=Row(A" & SRSt & ")"
    .Range("I" & ActiveCell.Row).Formula = "=Row(A" & SREn & ")"
    '[Task Row]
    .Range("A" & ActiveCell.Offset(4).Row & ":AZ" & ActiveCell.Offset(4).Row).Copy
    .Range("A" & ActiveCell.Offset(1).Row).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & ActiveCell.Row) = "[Task]"
    .Range("J" & ActiveCell.Row) = ""
    .Range("A" & SRSt).Offset(-1).EntireRow.Copy '[Blank Row]
    .Range("A" & SRSt).Offset(2).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
End With

'Insert to Services
Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect: Sheets("Amortised Summary").Unprotect

With Sheets("Frame Services")
    .Activate
    .Unprotect
    .Rows(IRSt).Offset(-1).Select
    .Rows(IRSt).EntireRow.Insert
    .Rows(IRSt).EntireRow.Insert
    .Rows(IRSt).EntireRow.Insert
    .Range("A" & ActiveCell.Offset(4).Row & ":AZ" & ActiveCell.Offset(4).Row).Copy
    .Range("A" & ActiveCell.Offset(1).Row).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & ActiveCell.Offset(4).Row & ":AZ" & ActiveCell.Offset(4).Row).Copy
    .Range("A" & ActiveCell.Offset(1).Row).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & IRSt).Formula = "='Scope'!A" & SRSt
    .Range("D" & IRSt).Formula = "=ROW(D" & TaskRow2 & ")"
    .Range("AZ" & IRSt).Formula = "=(SUM(K" & TaskRow1 & ":K" & TaskRow2 & ")+Sum(AV" & TaskRow1 & ":AV" & TaskRow2 & "))"
    .Range("BA" & IRSt).Formula = "=SUM(AY" & TaskRow1 & ":AY" & TaskRow2 & ")"
    .Range("BB" & IRSt).Formula = "=SUM(M" & TaskRow1 & ":M" & TaskRow2 & ")+Sum(Sum(AW" & TaskRow1 & ":AW" & TaskRow2 & ")-Sum(AV" & TaskRow1 & ":AV" & TaskRow2 & "))"
    .Range("BC" & IRSt).Formula = "=SUM(AX" & TaskRow1 & ":AX" & TaskRow2 & ")"
    .Range("B" & TaskRow1 & ":C" & TaskRow1) = 0
    .Range("F" & TaskRow1) = "Project Staff"
    .Range("G" & TaskRow1) = ""
    .Range("J" & TaskRow1) = "0"
    .Range("N" & TaskRow1) = 0
    .Range("O" & TaskRow1) = 0
    .Range("Q" & TaskRow1) = 0.15
    .Range("S" & TaskRow1) = 0
    .Range("T" & TaskRow1) = 0
    .Range("V" & TaskRow1) = 0.15
    .Range("X" & TaskRow1) = 0
    .Range("Y" & TaskRow1) = 0
    .Range("AA" & TaskRow1) = 0.15
    .Range("AC" & TaskRow1) = 0
    .Range("AD" & TaskRow1) = 0
    .Range("AF" & TaskRow1) = 0.15
    .Range("AH" & TaskRow1) = 0
    .Range("AI" & TaskRow1) = 0
    .Range("AK" & TaskRow1) = 0.15
    .Range("AM" & TaskRow1) = 0
    .Range("AN" & TaskRow1) = 0
    .Range("AP" & TaskRow1) = 0.15
    .Range("AR" & TaskRow1) = ""
    .Range("AS" & TaskRow1) = 0
    .Range("AT" & TaskRow1) = 0.15
    .Protect
    .Range("A5").Select
End With

With Sheets("Frame Summary")
    .Activate 'Insert to Frame Summary
    .Unprotect
    .Range("A" & FSI).Offset(1).EntireRow.Insert
    .Range("A" & FSI).Offset(1).EntireRow.Insert
    .Range("A" & FSI).Offset(3).EntireRow.Copy 'Segment Title
    .Range("A" & FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & FSI).Offset(1).Formula = "='Frame Services'!A" & IRSt
    .Range("B" & FSI).Offset(1).Formula = "='Frame Services'!AZ" & IRSt
    .Range("G" & FSI).Offset(1).Formula = "='Frame Services'!BA" & IRSt
    .Range("H" & FSI).Offset(1).Formula = "='Frame Services'!BB" & IRSt
    .Range("A" & FSI).EntireRow.Copy 'Task Item
    .Range("A" & FSI).Offset(2).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & FSI).Offset(2).Formula = "='Frame Services'!A" & TaskRow1
    .Range("C" & FSI).Offset(2).Formula = "=SUM('Frame Services'!K" & TaskRow1 _
     & ":INDIRECT(""'Frame Services'!K""&(ROW('Frame Services'!K" & (TaskRow1 + 1) _
     & ")-1)))"
    .Range("D" & FSI).Offset(2).Formula = "=SUM('Frame Services'!B" & TaskRow1 _
     & ":INDIRECT(""'Frame Services'!B""&(ROW('Frame Services'!B" & (TaskRow1 + 1) _
     & ")-1)))"
    .Range("E" & FSI).Offset(2).Formula = "=SUM('Frame Services'!AV" & TaskRow1 _
     & ":INDIRECT(""'Frame Services'!AV""&(ROW('Frame Services'!AV" & (TaskRow1 + 1) _
     & ")-1)))"
    .Range("F" & FSI).Offset(2).Formula = "=SUM('Frame Services'!AX" & TaskRow1 _
     & ":INDIRECT(""'Frame Services'!AX""&(ROW('Frame Services'!AX" & (TaskRow1 + 1) _
     & ")-1)))"
    .Protect
    .Range("B3").Select
End With

With Sheets("Client Summary")
    .Activate 'Insert to Client Summary
    .Unprotect
    .Range("A" & CSI).EntireRow.Insert
    .Range("A" & CSI).EntireRow.Insert
    '.Range("A" & CSI).Offset(-2).EntireRow.Copy
    .Range("A" & CSI).Offset(2).EntireRow.Copy   'copy Project Close formula
    .Range("A" & CSI).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & CSI).Offset(-1).EntireRow.Copy
    .Range("A" & CSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Protect
    .Range("G3").Select
End With

With Sheets("Amortised Summary")
    .Range("A" & CSI).EntireRow.Insert
    .Range("A" & CSI).EntireRow.Insert
    '.Range("A" & CSI).Offset(-2).EntireRow.Copy
    .Range("A" & CSI).Offset(2).EntireRow.Copy   'copy Project Close row
    .Range("A" & CSI).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("G" & CSI).Formula = "=C" & CSI
    .Range("H" & CSI, "AD" & CSI).Value = 0
    .Range("A" & CSI).Offset(-1).EntireRow.Copy
    .Range("A" & CSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
End With

Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect: Sheets("Amortised Summary").Protect

With Sheets("Scope")
    .Activate
    .Range("G" & SRSt).Formula = "=ROW('Frame Summary'!A" & FSRow & ")"
    .Range("O" & SRSt).Formula = "=ROW('Frame Services'!A" & IRSt & ")"
    .Range("P" & SRSt).Formula = "=ROW('Frame Summary'!A" & FSRow & ")"
    .Range("O" & SRSt).Offset(1).Formula = "=ROW('Frame Services'!A" & TaskRow1 & ")"
    .Range("P" & SRSt).Offset(1).Formula = "=ROW('Frame Summary'!A" & (FSRow + 1) & ")"
    .Range("O" & SRSt).Offset(2).Formula = "=ROW('Frame Services'!A" & TaskRow2 & ")"
    .Protect
    .Range("A" & SRSt).Offset(1).Select
End With

End If

Application.CutCopyMode = False
Application.EnableEvents = True
Application.ScreenUpdating = True

Call Sheets("Document Control").CB_Itemised_Tasks_Click 'Aligns new segment to current sheet setting

End Sub

Private Sub CutPhaseButton_Click()

If ActiveCell.Row < Range("C3").Value Then MsgBox ("Please Select Frame Scope Segments - Exiting Routine."): Exit Sub
If ActiveCell.Row > Range("D3").Value Then MsgBox ("Please Select Frame Scope Segments - Exiting Routine."): Exit Sub
If Range("H" & ActiveCell.Row).Value <> "" Then MsgBox ("Please Select Segement Task Row - Exiting Routine."): Exit Sub

If MsgBox("This will DELETE the selected segment and all tasks: Continue?", vbYesNo, "Change Log") = vbYes Then
Application.ScreenUpdating = False

Dim ARow As Long
Dim DR As Long
Dim FSR As Long 'Frame Summary Insert Row
Dim CSR As Long 'Client Summary Insert Row
Dim SRSt As Long 'Segment Row Starting
Dim SREn As Long 'Segment Row Ending
Dim PSR As Long 'Frame Services Start - Captures additional resources
Dim PFR As Long 'Frame Services End
Dim RVT As Long 'Row Value Tester

ARow = ActiveCell.Row

For X = 1 To 100 'Loop up till Segment Title line -> Store Location Values -> Generate Line Variables
    DR = ARow - X
    Sheets("Scope").Activate
    Sheets("Scope").Unprotect
    Sheets("Scope").Range("A" & DR).EntireRow.Select
    SRSt = Sheets("Scope").Range("H" & DR).Value  '- Hidden Value
    SREn = Sheets("Scope").Range("I" & DR).Value  '- Hidden Value
    PSR = Sheets("Scope").Range("O" & DR).Value '- Hidden Value
    FSR = Sheets("Scope").Range("G" & DR).Value '- Hidden Value
    If SRSt = "0" Then X = X
    If SRSt <> 0 Then Exit For
Next X

CSR = FSR - 1
PFR = Sheets("Frame Services").Range("D" & PSR).Value

Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect: Sheets("Amortised Summary").Unprotect

With Sheets("Amortised Summary")
    .Range("A" & CSR & ":A" & (CSR + SREn - SRSt - 1)).EntireRow.Delete 'Segment Start -> Last Task (Next Segment -1)
End With

With Sheets("Client Summary")
    .Activate
    .Unprotect
    .Range("A" & CSR).EntireRow.Select
    .Range("A" & CSR & ":A" & (CSR + SREn - SRSt - 1)).EntireRow.Delete 'Segment Start -> Last Task (Next Segment -1)
    .Protect
    .Range("G3").Select
End With

With Sheets("Frame Summary")
    .Activate
    .Unprotect
    .Range("A" & FSR).EntireRow.Select
    .Range("A" & FSR & ":A" & (FSR + SREn - SRSt - 1)).EntireRow.Delete 'Segment Start -> Last Task (Next Segment -1)
    .Protect
    .Range("B3").Select
End With

With Sheets("Frame Services")
    .Activate
    .Unprotect
    .Range("A" & PSR & ":AK" & PFR).EntireRow.Select
    .Range("A" & PSR & ":AK" & PFR).EntireRow.Delete
    .Protect
    .Range("A5").Select
End With

Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect: Sheets("Amortised Summary").Protect

With Sheets("Scope")
    .Activate
    .Unprotect
    .Range("A" & SRSt & ":AK" & SREn).EntireRow.Select
    .Range("A" & SRSt & ":AK" & SREn).EntireRow.Delete
    .Protect
    .Range("A" & SRSt).Select
End With

End If

Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub


