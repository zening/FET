VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub CB_AddEquipment_Click()

Dim ItemRow As Long
Dim SummRow As Long
Dim Summ1R As Long
Dim ClieRow As Long
Dim AdHocRow As Long
Dim FEQS As Long 'Equip Start
Dim RVT As Long 'Row Value Tester

ActiveSheet.Protect
Range("A" & ActiveCell.Row).Activate

FEQS = Sheets("Frame Summary").Range("F2").Value
ItemRow = ActiveCell.Row
SummRow = Range("S" & ActiveCell.Row).Value
Summ1R = SummRow + 1
ClieRow = SummRow - 1
AdHocRow = ItemRow + 11

If ActiveCell.Row < 7 Then MsgBox ("Please select Equipment Item Line - Exiting Routine"): Exit Sub
If ActiveCell.Row > Range("J1").Value Then MsgBox ("Please select Equipment Item Line - Exiting Routine"): Exit Sub

RVT = Sheets("Equipment").Range("N" & ActiveCell.Row).Value 'Determines if Supplier title, if so then select next line.
If RVT <> "0" Then ActiveCell.Offset(1).Select

Application.ScreenUpdating = False
ActiveSheet.Unprotect
Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect: Sheets("Amortised Summary").Unprotect

With Sheets("Equipment")
    .Range("A" & ItemRow).Offset(1).EntireRow.Insert
    .Range("A" & ItemRow).EntireRow.Copy
    .Range("A" & ItemRow).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & ItemRow).Offset(1) = "[Item Part Number]"
    .Range("B" & ItemRow).Offset(1) = "[Item Name & Description]"
    .Range("C" & ItemRow & ":E" & ItemRow).Offset(1) = 0
    .Range("G" & ItemRow).Offset(1) = 0
    .Range("I" & ItemRow).Offset(1) = 0
    .Range("O" & ItemRow).Offset(1) = ""
End With

With Sheets("Frame Summary")
    '.Unprotect
    .Rows(SummRow).Offset(1).EntireRow.Insert
    .Rows(SummRow).Copy
    .Rows(SummRow).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    '.Protect
End With

With Sheets("Client Summary")
    '.Unprotect
    .Rows(ClieRow).Offset(1).EntireRow.Insert
    .Rows(ClieRow).Copy
    .Rows(ClieRow).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    '.Protect
End With

With Sheets("Amortised Summary")
    '.Unprotect
    .Rows(ClieRow).Offset(1).EntireRow.Insert
    .Rows(ClieRow).Copy
    .Rows(ClieRow).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    '.Protect
End With

'use Frame Summary data for Quote sheet
AdHocRow = SummRow - 24

With Sheets("Quote")
    .Unprotect
    .Rows(AdHocRow).Offset(1).EntireRow.Insert
    .Rows(AdHocRow).Copy
    .Rows(AdHocRow).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
End With

With Sheets("Equipment")
    .Range("S" & ItemRow).Offset(1).Formula = "=ROW('Frame Summary'!A" & Summ1R & ")"
    .Range("A" & ItemRow).Offset(1).Select
    .Range("H" & Range("J1").Value).Formula = "=SUM(H5:H" & Range("J1").Value - 1 & ")"
    .Range("K" & Range("J1").Value).Formula = "=SUM(K5:K" & Range("J1").Value - 1 & ")"
    .Range("M" & Range("J1").Value).Formula = "=SUM(M5:M" & Range("J1").Value - 1 & ")"
    .Range("N" & Range("J1").Value).Formula = "=SUM(N5:N" & Range("J1").Value - 1 & ")"
    .Protect
End With
Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect: Sheets("Amortised Summary").Protect
Application.ScreenUpdating = True
End Sub

Public Sub CB_Remove_Item_Click()

Dim ItemRow As Long
Dim RVS As String
Dim RVT As String
Dim FEQS As Long
Dim SummRow As Long
Dim ClieRow As Long
Dim AdHocRow As Long
Dim RVX As Long 'Row Value Tester

Range("A" & ActiveCell.Row).Activate

FEQS = Sheets("Frame Summary").Range("F2").Value
ItemRow = ActiveCell.Row
SummRow = Range("S" & ActiveCell.Row).Value
ClieRow = SummRow - 1
AdHocRow = ItemRow + 11 'Fix later - Static + 11 reference will cause issues.

Application.ScreenUpdating = False

If ActiveCell.Row < 6 Then MsgBox ("This is Not An Equipment Item - Exiting Routine"): Exit Sub
If ActiveCell.Row > Range("J1").Value Then MsgBox ("This is Not An Equipment Item - Exiting Routine"): Exit Sub
If Range("C" & ActiveCell.Row).Value = "" Then MsgBox ("Select Equipment Item - Use: Remove Supplier - Exiting Routine"): Exit Sub
If Range("P" & ActiveCell.Row).Offset(-1).Value <> "" And Range("P" & ActiveCell.Row).Offset(2).Value <> "" Then MsgBox ("This is the only Equipment Item - Use: Remove Supplier - Exiting Routine"): Exit Sub
If Range("J1").Value = 9 Then Range("A7") = "[Item/Config Number]" _
                            : Range("B7") = "[Item Name & Description] // [BOM Title]" _
                            : Range("C7:E7") = 0 _
                            : Range("G" & ItemRow).Offset(1) = 0 _
                            : Range("I" & ItemRow).Offset(1) = 0 _
                            : Range("O" & ItemRow).Offset(1) = "" _
                            : MsgBox ("This is the only Item - Exiting Routine") _
                            : Exit Sub




With Sheets("Equipment")
    .Unprotect
    .Rows(ItemRow).EntireRow.Delete
    .Protect
End With
With Sheets("Frame Summary")
    .Unprotect
    .Rows(SummRow).EntireRow.Delete
    .Protect
End With
With Sheets("Client Summary")
    .Unprotect
    .Rows(ClieRow).EntireRow.Delete
    .Protect
End With
With Sheets("Amortised Summary")
    .Unprotect
    .Rows(ClieRow).EntireRow.Delete
    .Protect
End With
With Sheets("Quote")
    .Rows(AdHocRow).EntireRow.Delete
End With
With Sheets("Equipment")
    .Activate
    .Range("A" & ItemRow).Offset(-1).Select
    .Protect
End With

Application.ScreenUpdating = True
End Sub

Private Sub CB_AddMultiple_Click()
Dim InputData As Variant

InputData = Application.InputBox(Prompt:="Number of Equipment Lines:", Title:="Equipment List", Type:=1)
If InputData <= 0 Then Exit Sub

For r = 1 To InputData
    Call Sheets("Equipment").CB_AddEquipment_Click
Next r

End Sub

Private Sub CB_CutMultiple_Click()

Dim InputData As Variant
Dim DR As Long
Dim ARow As Long
Dim FEQS As Long
Dim SummRow As Long
Dim ClieRow As Long
Dim AdHocRow As Long
Dim RVS As Long
Dim RVT As Long
Dim RVX As Long 'Row Value Tester

Range("A" & ActiveCell.Row).Activate

FEQS = Sheets("Frame Summary").Range("F2").Value
RVX = Sheets("Equipment").Range("P" & ActiveCell.Row).Value 'Row Tester Value

If RVX <> "0" Then MsgBox "This is the Supplier - Use: Remove Supplier - Exiting Routine": Exit Sub 'Exit if Supplier Row
If ActiveCell.Row > Range("J1").Value Then MsgBox "This is Not An Equipment Item - Exiting Routine": Exit Sub 'Exit if beyond Eq List
If Range("J1").Value = 9 Then Range("A7") = "[Item Part Number]" _
                            : Range("B7") = "[Item Name & Description]" _
                            : Range("C7:E7") = 0 _
                            : Range("G" & ItemRow).Offset(1) = 0 _
                            : Range("I" & ItemRow).Offset(1) = 0 _
                            : Range("O" & ItemRow).Offset(1) = "" _
                            : MsgBox ("This is the only Item - Exiting Routine") _
                            : Exit Sub
                            
Application.ScreenUpdating = False
InputData = Application.InputBox(Prompt:="BEWARE: Remove X Number of Equipment Lines Below:", Title:="Equipment List", Type:=1)
If InputData <= 0 Then Exit Sub

RVX = ActiveCell.Row 'Reuse Test Variable -> Start marker
For r = 1 To InputData 'Loops down to end of selecter Supplier's Item List
    Range("A" & RVX).Offset(r).Select
    If Range("A" & RVX).Offset(r) = "" Then Range("A" & RVX).Offset(r - 1).Select: Exit For
Next r

ARow = ActiveCell.Row
For r = 0 To InputData
    DR = ARow - r
    If Range("Q" & DR).Offset(-1) <> "" Then MsgBox "No More Equipment To Remove For This Supplier - Exiting Routine": Exit For
    If Range("J1").Value = 9 Then MsgBox "No More Equipment To Remove - Exiting Routine": Exit For
    SummRow = Sheets("Equipment").Range("S" & DR).Value
    ClieRow = SummRow - 1
    AdHocRow = DR + 11 'Likely to generate issue - Fix later

    With Sheets("Amortised Summary")
        .Unprotect
        .Rows(ClieRow).EntireRow.Delete
        .Protect
    End With
    With Sheets("Client Summary")
        .Unprotect
        .Rows(ClieRow).EntireRow.Delete
        .Protect
    End With
    With Sheets("Frame Summary")
        .Unprotect
        .Rows(SummRow).EntireRow.Delete
        .Protect
    End With
    With Sheets("Quote")
        .Rows(AdHocRow).EntireRow.Delete
    End With
    With Sheets("Equipment")
        .Unprotect
        .Rows(DR).EntireRow.Delete
        .Range("S" & Sheets("Equipment").Range("J1").Value).Formula = "=Row('Frame Summary'!A" & Sheets("Frame Summary").Range("F3").Value - 1 & ")"
        .Protect
    End With
Next r

Range("A6").Select
Application.ScreenUpdating = True
End Sub

Private Sub CB_AddSupplier_Click()
Application.ScreenUpdating = False

If ActiveCell.Row < 7 Then
'MsgBox ("Please select Equipment Item Line - Exiting Routine"): Exit Sub
    Range("A7").Select
    
End If

If ActiveCell.Row > Range("J1").Value Then MsgBox ("Please select Equipment Item Line - Exiting Routine"): Exit Sub

If MsgBox("This will insert a new Equipment Supplier below the present one: Continue?", vbYesNo, "Change Log") = vbYes Then
Dim FSRow As Long
Dim FSI As Long 'Frame Summary Insert Row
Dim CSI As Long 'Client Summary Insert Row
Dim SRSt As Long 'Segment Row Starting
Dim IRSt As Long 'Insertion Row Starting
Dim SREn As Long 'Segment Row Ending
Dim SRSup As Long
Dim FSSup As Long
Dim QI As Long
Dim QN As Long
Dim FSN As Long
Dim QSR As Long
Dim RVT As Long 'Row Value Tester

Range("A" & ActiveCell.Row).Activate

Sheets("Equipment").Unprotect
If ActiveCell.Row > Range("J1").Value Then Range("A" & Range("O6").Value).Offset(-1).Select 'If below list, select bottom supplier

For X = ActiveCell.Row To 400 'Loop till populated line -> Store Location Values -> Generate Line Variables
    Sheets("Equipment").Range("A" & X).EntireRow.Select
    SRSt = Sheets("Equipment").Range("P" & X).Value  '- Hidden Value
    QSR = Sheets("Equipment").Range("R" & X).Value  '- Hidden Value
    FSRow = Sheets("Equipment").Range("S" & X).Value '- Hidden Value
    If SRSt = "0" Then X = X
    If SRSt <> 0 Then Exit For
Next X

SRSup = SRSt + 1
SREn = SRSt + 2
IRSt = SRSt - 1
FSI = FSRow - 1
FSN = FSRow + 1
QI = QSR - 1
QN = QSR + 1
CSI = FSI - 2
FSSup = Sheets("Frame Summary").Range("F2").Value

With Sheets("Equipment")
    .Rows(IRSt).Offset(1).EntireRow.Insert
    .Rows(IRSt).Offset(1).EntireRow.Insert
    .Rows(IRSt).Offset(1).EntireRow.Insert
    .Range("A6").EntireRow.Copy
    .Rows(IRSt).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & ActiveCell.Row) = "[Supplier Name]"
    .Range("B" & ActiveCell.Row) = "[Quote Number] / [Configurator File ID]"
    .Range("R" & ActiveCell.Row).Formula = "=ROW('Frame Summary'!A" & FSRow & ")"
    .Range("A7").EntireRow.Copy
    .Rows(SRSt).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & ActiveCell.Row) = "[Item Part Number]"
    .Range("B" & ActiveCell.Row) = "[Item Name & Description]"
    .Range("C" & ActiveCell.Row & ":E" & ActiveCell.Row) = 0
    .Range("G" & ActiveCell.Row) = 0
    .Range("I" & ActiveCell.Row) = 0
    .Range("O" & ActiveCell.Row) = ""
    .Range("Q" & SRSt).Formula = "=Row(A" & SREn & ")"
End With

With Sheets("Frame Summary")
    .Activate 'Insert to Frame Summary
    .Unprotect
    .Range("A" & FSI).EntireRow.Select   'Supplier Row
    .Range("A" & FSI).Offset(1).EntireRow.Insert
    .Range("A" & FSI).Offset(1).EntireRow.Insert
    .Range("A" & FSSup).Offset(1).EntireRow.Copy
    .Range("A" & FSI).Offset(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & ActiveCell.Row).Formula = "='Equipment'!A" & SRSt
    .Range("A" & FSSup).Offset(2).EntireRow.Copy    'Item Row
    .Range("A" & FSI).Offset(2).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & ActiveCell.Row).Formula = "='Equipment'!A" & SRSup
    .Range("B" & ActiveCell.Row).Formula = "='Equipment'!K" & SRSup
    .Range("G" & ActiveCell.Row).Formula = "=Equipment!N" & SRSup
    .Range("H" & ActiveCell.Row).Formula = "='Equipment'!M" & SRSup
    .Range("B3").Select
    .Protect
End With

With Sheets("Client Summary")
    .Activate 'Insert to Client Summary
    .Unprotect
    .Range("A" & CSI).EntireRow.Select
    .Range("A" & CSI).Offset(2).EntireRow.Insert
    .Range("A" & CSI).Offset(2).EntireRow.Insert
    .Range("A" & CSI).EntireRow.Copy
    .Range("A" & CSI).Offset(2).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & CSI).Offset(1).EntireRow.Copy
    .Range("A" & CSI).Offset(3).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("F3").Select
    .Protect
End With

With Sheets("Amortised Summary")
    .Unprotect
    .Range("A" & CSI).Offset(2).EntireRow.Insert
    .Range("A" & CSI).Offset(2).EntireRow.Insert
    .Range("A" & CSI).Copy
    .Range("A" & CSI).Offset(2).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("F" & CSI).Copy
    .Range("F" & CSI).Offset(2).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("A" & CSI).Offset(2).IndentLevel = 2
    .Range("A" & CSI).Offset(2).EntireRow.Interior.color = RGB(255, 255, 255)
    .Range("F" & CSI).Offset(2).IndentLevel = 1
    .Range("A" & CSI).Offset(1).EntireRow.Copy
    .Range("A" & CSI).Offset(3).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Protect
End With

'use Frame Summary data for Quote
QI = FSI - 24
QSR = QI + 1
With Sheets("Quote")
    .Unprotect
    .Rows(QI).Offset(1).EntireRow.Insert
    .Rows(QI).Offset(1).EntireRow.Insert
    .Rows(17).Copy
    .Rows(QSR).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("B" & QSR).Formula = "='Equipment'!B" & SRSt
    .Rows(18).Copy
    .Rows(QI).Offset(2).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    .Range("B" & QSR + 1).Formula = "='Equipment'!B" & SRSup
End With

With Sheets("Equipment")
    .Activate
    .Unprotect
    .Range("S" & SRSt).Formula = "=ROW('Frame Summary'!A" & FSRow & ")"
    .Range("R" & SRSt).Formula = "=ROW('Quote'!A" & QSR & ")"
    .Range("S" & SRSup).Formula = "=ROW('Frame Summary'!A" & FSN & ")"
    .Range("A" & SRSup).Select
    .Range("H" & Range("J1").Value).Formula = "=SUM(H5:H" & Range("J1").Value - 1 & ")"
    .Range("K" & Range("J1").Value).Formula = "=SUM(K5:K" & Range("J1").Value - 1 & ")"
    .Range("M" & Range("J1").Value).Formula = "=SUM(M5:M" & Range("J1").Value - 1 & ")"
    .Range("N" & Range("J1").Value).Formula = "=SUM(N5:N" & Range("J1").Value - 1 & ")"
    .Protect
End With

End If
Application.ScreenUpdating = True
End Sub

Private Sub CB_CutSupplier_Click()

If MsgBox("This will DELETE the selected Supplier and all associated equipment: Continue?", vbYesNo, "Change Log") = vbYes Then

'Routine Checker
If ActiveCell.Row < 6 Then MsgBox "This is Not A Removable Item - Exiting Routine": Exit Sub
If ActiveCell.Row > Range("J1").Value Then MsgBox "This is Not A Removable Item - Exiting Routine": Exit Sub
If Range("J1").Value = 9 Then MsgBox "No More Suppliers or Equipment To Remove - Exiting Routine": Exit Sub
If Range("P" & ActiveCell.Row).Value <> "0" Then ActiveCell.Offset(1).Select 'Determines if Supplier title, if so then selects next line.

Dim ARow As Long
Dim DR As Long
Dim FSRE As Long
Dim FSRS As Long 'Frame Summary Insert Row
Dim CSR As Long 'Client Summary Insert Row
Dim SRSt As Long 'Segment Row Starting
Dim SREn As Long 'Segment Row Ending
Dim QSR As Long
Dim QFR As Long
Dim RDif As Long 'Row Difference
Dim RVT As Long 'Row Value Tester

Range("A" & ActiveCell.Row).Activate

Application.ScreenUpdating = False
ARow = ActiveCell.Row
Sheets("Equipment").Unprotect

For X = 1 To 200 'Loop up till populated line -> Store Location Values -> Generate Line Variables
    DR = ARow - X
    Sheets("Equipment").Range("A" & DR).EntireRow.Select
    SRSt = Sheets("Equipment").Range("P" & DR).Value  '- Hidden Value
    SREn = Sheets("Equipment").Range("Q" & DR).Value  '- Hidden Value
    QSR = Sheets("Equipment").Range("R" & DR).Value '- Hidden Value
    FSRS = Sheets("Equipment").Range("S" & DR).Value '- Hidden Value
    If SRSt = "0" Then X = X
    If SRSt <> 0 Then Exit For
Next X

RDif = SREn - SRSt - 1
FSRE = FSRS + RDif
CSRS = FSRS - 1
CSRE = FSRE - 1
QFR = QSR + SREn - SRSt

With Sheets("Amortised Summary")
    .Unprotect
    .Range("A" & CSRS & ":AK" & CSRE).EntireRow.Delete
    .Protect
End With

With Sheets("Client Summary")
    .Unprotect
    .Range("A" & CSRS & ":AK" & CSRE).EntireRow.Delete
    .Protect
End With

With Sheets("Frame Summary")
    .Unprotect
    .Range("A" & FSRS & ":AK" & FSRE).EntireRow.Delete
    .Protect
End With

With Sheets("Quote")
    .Unprotect
    .Range("A" & QSR & ":AK" & QFR).EntireRow.Delete
End With

With Sheets("Equipment")
    .Range("A" & SRSt & ":AK" & SREn).EntireRow.Delete
    .Range("S" & Sheets("Equipment").Range("J1").Value).Formula = "=Row('Frame Summary'!A" & Sheets("Frame Summary").Range("F3").Value - 1 & ")"
    .Range("A6").Select
    .Protect
End With

End If
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

