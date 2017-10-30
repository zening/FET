VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''''''''''''''''''''Control Panel Buttons'''''''''''''''''''''''''''''''


Private Sub Scheduler_Click() 'Creates MS Project file from Services
Application.ScreenUpdating = False

Dim objProj As Object
Dim strTask As String
Dim strResource As String
Dim strWork As String
Dim SchedName As String
Dim NewScheduleNameS As String
Dim LongNameS As String
Dim LastRow As Variant
Dim NewSchedule

Set objProj = CreateObject("MSProject.Application")
    objProj.Visible = True
Set NewSchedule = objProj.Projects.Add()
    NewSchedule.Title = Sheets("Document Control").Range("A6").Value
Set ActiveProject = NewSchedule

'Projects
With Sheets("Frame Services")
    .Activate
    .Unprotect
End With

LastRow = Sheets("Frame Services").Range("B1").Value

For i = 4 To LastRow
    Worksheets("Frame Services").Range("A" & i).Select
    strTask = Worksheets("Frame Services").Range("A" & i)
    NewSchedule.Tasks.Add (strTask)
    strResource = Worksheets("Frame Services").Range("F" & i) 'See if you can add: ( & ":G" i) to combine resource / skillset
    NewSchedule.Tasks(i - 3).ResourceNames = strResource
    strWork = Worksheets("Frame Services").Range("E" & i)
    NewSchedule.Tasks(i - 3).Work = strWork
    'Sets indent layer to group tasks by segment
    If NewSchedule.Tasks(i - 3).ResourceNames <> "" Then NewSchedule.Tasks(i - 3).OutlineLevel = 2 Else NewSchedule.Tasks(i - 3).OutlineLevel = 1
        If i <> 4 Then
            If (NewSchedule.Tasks(i - 3).ResourceNames <> "" And NewSchedule.Tasks(i - 3 - 1).OutlineLevel = 2) Then NewSchedule.Tasks(i - 3).Predecessors = NewSchedule.Tasks(i - 3 - 1)
        Else
        End If

Next i

Sheets("Frame Services").Protect
Application.ScreenUpdating = True
End Sub

Public Sub CB_Export_Click()
Application.ScreenUpdating = False
Dim InputData As Variant

InputData = Application.InputBox(Prompt:="Enter 1 for a Scope List .doc or 2 for Workbook Output .pdf", Title:="Export Data", Type:=1)
If InputData < 1 Then MsgBox ("Please use valid input - Exiting Routine"): Exit Sub
If InputData > 2 Then MsgBox ("Please use valid input - Exiting Routine"): Exit Sub
If InputData = 1 Then Call Sheets("Document Control").CB_Output_Click
If InputData = 2 Then Call Sheets("Document Control").CB_PDF_Click

Application.ScreenUpdating = True
End Sub

Public Sub CB_Output_Click()
Application.ScreenUpdating = False

Dim strTask As String
Dim strComm As String
Dim NewDoc
Dim objDoc As Object

Set objDoc = CreateObject("Word.Application")
    objDoc.Visible = True
Set NewDoc = objDoc.Documents.Add()
Set ActiveDocument = NewDoc

objDoc.Selection.TypeText Text:="Cost Estimate Output"
NewDoc.Content.InsertParagraphAfter

For i = Sheets("Scope").Range("A3").Value To Sheets("Scope").Range("H3").Value
    strTask = Sheets("Scope").Range("A" & i)
    strComm = Sheets("Scope").Range("J" & i)
    If strTask = "" Then strTask = Sheets("Scope").Range("B" & i)
    If strTask = "" And Sheets("Scope").Rows(i).Hidden = False Then NewDoc.Content.InsertParagraphAfter
    If strTask <> "" And Sheets("Scope").Rows(i).Hidden = False Then NewDoc.Content.InsertAfter (strTask & " - " & strComm): NewDoc.Content.InsertParagraphAfter
    
Next i

Application.ScreenUpdating = True
End Sub

Public Sub CB_PDF_Click()
Dim f As Boolean
Dim wsh As Worksheet
'To do: Figure out how to 'print' document in A3 / Landscape
'Loop selects all visible sheets then defaults
f = True
    For Each wsh In Worksheets
        If wsh.Name = "User Guide" Then Exit For
        If wsh.Visible = xlSheetVisible Then
            wsh.Select Replace:=f
            f = False
        End If
    Next wsh
        
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=Sheets("Document Control").Range("A6").Value, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
Sheets(1).Select
End Sub

Private Sub CB_Instructions_Click() 'Brings up Instructions
Application.ScreenUpdating = False
Worksheets("User Guide").Visible = xlSheetVisible
Worksheets("User Guide").Activate
Application.ScreenUpdating = True
End Sub

'Version Up
Private Sub CB_Version_Click()
Dim LastRow As Long
Dim ActBook As Workbook
Dim NewFile As String
Dim LongName As String

Application.ScreenUpdating = False
ActiveSheet.Unprotect

If MsgBox("Were changes made to this document?", vbYesNo, "Change Log") = vbYes Then
LastRow = Sheets("Document Control").Cells(Rows.count, 9).End(xlUp).Row + 1
    'Maps File Path to Name - saving in the origin folder
    NewFile = (Worksheets("Document Control").Range("B1") & " - " & Worksheets("Document Control").Range("F2") & " - " & Format(Now(), "yyyymmdd") & " " & Format(Now(), "hhmmss") & ".xlsm")
    LongName = ThisWorkbook.path & "\" & NewFile

        'Adds change log
    With Sheets("Document Control")
       .Cells(LastRow, 9) = NewFile
       .Cells(LastRow, 17) = Application.UserName
       .Cells(LastRow, 19) = InputBox("Please list changes below:", "Comments")
    End With
    'Maps Save to SaveAs by checking if file of current name exists
    If NewFile <> "" And NewFile <> "False" Then
        ActiveWorkbook.SaveAs Filename:=LongName, _
        FileFormat:=52
        Set ActBook = ActiveWorkbook
        Workbooks.Open LongName
    End If
      
    Else
    Exit Sub
    
End If
Application.ScreenUpdating = True
End Sub

''''''''''''''''''''''''''''''''Work Type''''''''''''''''''''''''''''''''
Private Sub CB_EQONLY_Click()
If CB_Project.Value = True Or CB_Consulting.Value = True Then MsgBox ("Work type may only be changed on fresh workbooks - Exiting Routine"): CB_EQONLY.Value = False: Exit Sub
Application.ScreenUpdating = False
If CB_EQONLY.Value = True Then CB_Consulting.Value = False: CB_Project.Value = False _
                             : CB_Equipment.Value = True: CB_Ad_Hoc_Quote.Value = True
If CB_EQONLY.Value = False Then CB_Equipment.Value = False: CB_Ad_Hoc_Quote.Value = False
Application.ScreenUpdating = True
End Sub

Private Sub CB_Consulting_Click()
If CB_Project.Value = True Or CB_EQONLY.Value = True Then MsgBox ("Work type may only be changed on fresh workbooks - Exiting Routine"): CB_Consulting.Value = False: Exit Sub
If CB_Consulting.Value = True Then Application.ScreenUpdating = False _
                                 : Sheets("Scope").Visible = xlSheetVisible _
                                 : Sheets("Client Summary").Visible = xlSheetVisible _
                                 : Sheets("Frame Summary").Visible = xlSheetVisible _
                                 : Sheets("Rates").Visible = xlSheetVisible _
                                 : Sheets("Frame Services").Visible = xlSheetVisible _
                                 : Sheets("Scope").Unprotect _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("I3").Value).Value = "Contract Prelims:" _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("J3").Value).Value = "Gap Analysis:" _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("K3").Value).Value = "Solution Delivery:" _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("L3").Value).Value = "Documentation Development:" _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("M3").Value).Value = "Handover Activities:" _
                                 : Sheets("Scope").Protect

If CB_Consulting.Value = False Then Sheets("Scope").Visible = xlSheetHidden _
                                 : Sheets("Client Summary").Visible = xlSheetHidden _
                                 : Sheets("Frame Summary").Visible = xlSheetHidden _
                                 : Sheets("Scope").Visible = xlSheetHidden _
                                 : Sheets("Frame Services").Visible = xlSheetHidden _
                                 : Sheets("Rates").Visible = xlSheetHidden
Application.ScreenUpdating = True
End Sub

Private Sub CB_Project_Click()
If CB_Consulting.Value = True Or CB_EQONLY.Value = True Then MsgBox ("Work type may only be changed on fresh workbooks - Exiting Routine"): CB_Project.Value = False: Exit Sub

If CB_Project.Value = True Then Application.ScreenUpdating = False _
                                 : Sheets("Scope").Visible = xlSheetVisible _
                                 : Sheets("Client Summary").Visible = xlSheetVisible _
                                 : Sheets("Frame Summary").Visible = xlSheetVisible _
                                 : Sheets("Rates").Visible = xlSheetVisible _
                                 : Sheets("Frame Services").Visible = xlSheetVisible _
                                 : Sheets("Scope").Unprotect _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("I3").Value).Value = "Project Prelims:" _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("J3").Value).Value = "Project Management:" _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("K3").Value).Value = "Solution Design:" _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("L3").Value).Value = "Implementation:" _
                                 : Sheets("Scope").Range("A" & Sheets("Scope").Range("M3").Value).Value = "Project Close:" _
                                 : Sheets("Scope").Protect

If CB_Project.Value = False Then Sheets("Scope").Visible = xlSheetHidden _
                                 : Sheets("Client Summary").Visible = xlSheetHidden _
                                 : Sheets("Frame Summary").Visible = xlSheetHidden _
                                 : Sheets("Scope").Visible = xlSheetHidden _
                                 : Sheets("Frame Services").Visible = xlSheetHidden _
                                 : Sheets("Rates").Visible = xlSheetHidden
                                 
Application.ScreenUpdating = True
End Sub


''''''''''''''''''''''''Client Rates''''''''''''''''''''''''''''''''''''
'R1: Standard // R2: Perth Airport // R3: MelAir (Additional Scope Sheet) // R4: Lend Lease
'R5: Lend Lease // R6: DIIPSIM (Unused) // R7: Mod B: On-Base (Unused) // R8: Mod B: Off-Base (Unused)

Private Sub CB_R1_Click()
Application.ScreenUpdating = False
If CB_R1.Value = True Then Sheets("Rates").Unprotect: Sheets("Rates").Columns("C:G").EntireColumn.Hidden = False _
                                : CB_R2.Value = False: CB_R3.Value = False: CB_R4.Value = False: CB_R5.Value = False _
                                : CB_R6.Value = False: CB_R7.Value = False: CB_R8.Value = False _
                                : Sheets("Frame Services").Unprotect _
                                : Sheets("Frame Services").Range("H1") = "R1" _
                                : Sheets("Frame Services").Protect _
                                : Sheets("Rates").Protect
If CB_R1.Value = False Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("C:G").EntireColumn.Hidden = True _
                                : Sheets("Rates").Protect
Application.ScreenUpdating = True
End Sub

Private Sub CB_R2_Click()
Application.ScreenUpdating = False

If CB_R2.Value = True Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("H:M").EntireColumn.Hidden = False _
                                : CB_R1.Value = False: CB_R3.Value = False: CB_R4.Value = False: CB_R5.Value = False _
                                : CB_R6.Value = False: CB_R7.Value = False: CB_R8.Value = False _
                                : Sheets("Frame Services").Unprotect _
                                : Sheets("Frame Services").Range("H1") = "R2" _
                                : Sheets("Frame Services").Protect _
                                : Sheets("Rates").Protect
If CB_R2.Value = False Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("H:M").EntireColumn.Hidden = True _
                                : Sheets("Rates").Protect
Application.ScreenUpdating = True
End Sub

Private Sub CB_R3_Click()
Application.ScreenUpdating = False

If CB_R3.Value = True Then Sheets("Rates").Unprotect _
                                  : Sheets("Rates").Columns("N:S").EntireColumn.Hidden = False _
                                  : CB_R1.Value = False: CB_R2.Value = False: CB_R4.Value = False: CB_R5.Value = False _
                                  : CB_R6.Value = False: CB_R7.Value = False: CB_R8.Value = False _
                                  : Sheets("Frame Services").Unprotect _
                                  : Sheets("Frame Services").Range("H1") = "R3" _
                                  : Sheets("Frame Services").Protect _
                                  : Sheets("MelAir SOW").Visible = xlSheetVisible _
                                  : CB_SiteMisc.Value = True _
                                  : Sheets("Rates").Protect
                                  
If CB_R3.Value = False Then Sheets("Rates").Unprotect _
                                  : Sheets("MelAir SOW").Visible = xlSheetHidden _
                                  : CB_SiteMisc.Value = False _
                                  : Sheets("Rates").Protect
                                  
Application.ScreenUpdating = True
End Sub

Private Sub CB_R4_Click()
Application.ScreenUpdating = False
If CB_R4.Value = True Then Sheets("Rates").Unprotect _
                            : Sheets("Rates").Columns("T:Y").EntireColumn.Hidden = False _
                            : CB_R1.Value = False: CB_R2.Value = False: CB_R3.Value = False: CB_R5.Value = False _
                            : CB_R6.Value = False: CB_R7.Value = False: CB_R8.Value = False _
                            : Sheets("Frame Services").Unprotect _
                            : Sheets("Frame Services").Range("H1") = "R4" _
                            : Sheets("Frame Services").Protect _
                            : Sheets("Rates").Protect
If CB_R4.Value = False Then Sheets("Rates").Unprotect _
                            : Sheets("Rates").Columns("T:Y").EntireColumn.Hidden = True _
                            : Sheets("Rates").Protect
Application.ScreenUpdating = True
End Sub

Private Sub CB_R5_Click()
Application.ScreenUpdating = False
If CB_R5.Value = True Then Sheets("Rates").Unprotect: Sheets("Rates").Columns("Z:AE").EntireColumn.Hidden = False _
                                : CB_R1.Value = False: CB_R2.Value = False: CB_R3.Value = False: CB_R4.Value = False _
                                : CB_R6.Value = False: CB_R7.Value = False: CB_R8.Value = False _
                                : Sheets("Frame Services").Unprotect _
                                : Sheets("Frame Services").Range("H1") = "R5" _
                                : Sheets("Frame Services").Protect _
                                : Sheets("Rates").Protect
If CB_R5.Value = False Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("Z:AE").EntireColumn.Hidden = True _
                                : Sheets("Rates").Protect
Application.ScreenUpdating = True
End Sub

Private Sub CB_R6_Click()
Application.ScreenUpdating = False

If CB_R6.Value = True Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("AF:AK").EntireColumn.Hidden = False _
                                : CB_R1.Value = False: CB_R2.Value = False: CB_R3.Value = False: CB_R4.Value = False _
                                : CB_R5.Value = False: CB_R7.Value = False: CB_R8.Value = False _
                                : Sheets("Frame Services").Unprotect _
                                : Sheets("Frame Services").Range("H1") = "R6" _
                                : Sheets("Frame Services").Protect _
                                : Sheets("Rates").Protect
If CB_R6.Value = False Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("AF:AK").EntireColumn.Hidden = True _
                                : Sheets("Rates").Protect
Application.ScreenUpdating = True
End Sub

Private Sub CB_R7_Click()
Application.ScreenUpdating = False

If CB_R7.Value = True Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("AL:AQ").EntireColumn.Hidden = False _
                                : CB_R1.Value = False: CB_R2.Value = False: CB_R3.Value = False: CB_R4.Value = False _
                                : CB_R5.Value = False: CB_R6.Value = False: CB_R8.Value = False _
                                : Sheets("Frame Services").Unprotect _
                                : Sheets("Frame Services").Range("H1") = "R7" _
                                : Sheets("Frame Services").Protect _
                                : Sheets("Rates").Protect
If CB_R7.Value = False Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("AL:AQ").EntireColumn.Hidden = True _
                                : Sheets("Rates").Protect
Application.ScreenUpdating = True
End Sub

Private Sub CB_R8_Click()
Application.ScreenUpdating = False
If CB_R8.Value = True Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("AR:AW").EntireColumn.Hidden = False _
                                : CB_R1.Value = False: CB_R2.Value = False: CB_R3.Value = False: CB_R4.Value = False _
                                : CB_R5.Value = False: CB_R6.Value = False: CB_R7.Value = False _
                                : Sheets("Frame Services").Unprotect _
                                : Sheets("Frame Services").Range("H1") = "R8" _
                                : Sheets("Frame Services").Protect _
                                : Sheets("Rates").Protect
If CB_R8.Value = False Then Sheets("Rates").Unprotect _
                                : Sheets("Rates").Columns("AR:AW").EntireColumn.Hidden = True _
                                : Sheets("Rates").Protect
Application.ScreenUpdating = True
End Sub

''''''''''''''''''''''''WORKSHEET FORMS''''''''''''''''''''''''''''

Private Sub CB_Ad_Hoc_Quote_Click()
Application.ScreenUpdating = False
If CB_Ad_Hoc_Quote.Value = True Then Sheets("Quote").Visible = xlSheetVisible
If CB_Ad_Hoc_Quote.Value = False Then Sheets("Quote").Visible = xlSheetHidden
Application.ScreenUpdating = True
End Sub

Private Sub CB_CS_Click()
Application.ScreenUpdating = False
If CB_CS.Value = True Then Sheets("Client Summary").Visible = xlSheetVisible
If CB_CS.Value = False Then Sheets("Client Summary").Visible = xlSheetHidden
Application.ScreenUpdating = True
End Sub

Private Sub CB_SetupForm_Click()  'Changes sheet to visible.
Application.ScreenUpdating = False
If CB_SetupForm = True Then Worksheets("Project Setup Request").Visible = xlSheetVisible
If CB_SetupForm = False Then Worksheets("Project Setup Request").Visible = xlSheetHidden
Application.ScreenUpdating = True
End Sub

Private Sub CB_Tech_Click()
Application.ScreenUpdating = False
If CB_Tech.Value = True Then Sheets("Technology Request").Visible = xlSheetVisible
If CB_Tech.Value = False Then Sheets("Technology Request").Visible = xlSheetHidden
Application.ScreenUpdating = True
End Sub

Private Sub CB_Variation_Click()
Application.ScreenUpdating = False
If CB_Variation.Value = True Then Sheets("Variation Request").Visible = xlSheetVisible
If CB_Variation.Value = False Then Sheets("Variation Request").Visible = xlSheetHidden
Application.ScreenUpdating = True
End Sub

''''''''''''''''''CONTRACT OPTIONS''''''''''''''''''''''''''''''''''

Private Sub CB_Amortised_Click()
Application.ScreenUpdating = False

Dim InputData As Variant
InputData = MsgBox("This is a Sales feature - Requires BDM Approval: Continue?", vbYesNo, "Change Log")

If InputData = vbYes Then Sheets("Amortised Summary").Visible = xlSheetVisible: CB_Amortised = True
If InputData = vbNo Then Sheets("Amortised Summary").Visible = xlSheetHidden: CB_Amortised = False

Application.ScreenUpdating = True
End Sub

Private Sub CB_BG_Click() 'Bank Guarantee
Application.ScreenUpdating = False
Dim FBGS As Long
Dim FBGE As Long
Dim CBGS As Long
Dim CBGE As Long
FBGS = Sheets("Frame Summary").Range("H2").Value
FBGE = FBGS + 3
CBGS = FBGS - 2
CBGE = FBGS - 1
If CB_BG.Value = True Then Sheets("Client Summary").Unprotect _
                          : Sheets("Client Summary").Rows(CBGS & ":" & CBGE).EntireRow.Hidden = False _
                          : Sheets("Client Summary").Protect _
                          : Sheets("Frame Summary").Unprotect _
                          : Sheets("Frame Summary").Rows(FBGS & ":" & FBGE).Hidden = False _
                          : Sheets("Frame Summary").Range("E" & FBGS).Value = "Yes" _
                          : Sheets("Frame Summary").Protect
If CB_BG.Value = False Then Sheets("Client Summary").Unprotect _
                          : Sheets("Client Summary").Rows(CBGS & ":" & CBGE).EntireRow.Hidden = True _
                          : Sheets("Client Summary").Protect _
                          : Sheets("Frame Summary").Unprotect _
                          : Sheets("Frame Summary").Range("E" & FBGS).Value = "No" _
                          : Sheets("Frame Summary").Rows(FBGS & ":" & FBGE).Hidden = True _
                          : Sheets("Frame Summary").Protect
Application.ScreenUpdating = True
End Sub

Private Sub CB_Cli_Con_Click()
Application.ScreenUpdating = False
If CB_Cli_Con = True Then Sheets("Client Summary").Unprotect _
                        : Sheets("Client Summary").Columns(2).Hidden = False _
                        : Sheets("Client Summary").Protect
If CB_Cli_Con = False Then Sheets("Client Summary").Unprotect _
                         : Sheets("Client Summary").Columns(2).Hidden = True _
                         : Sheets("Client Summary").Protect
Application.ScreenUpdating = True
End Sub

Private Sub CB_Equipment_Click()
Application.ScreenUpdating = False
Dim CSEFRow As Long
Dim CSELRow As Long
Dim FSEFRow As Long
Dim FSELRow As Long
FSEFRow = Sheets("Frame Summary").Range("F2").Value
FSELRow = Sheets("Frame Summary").Range("F3").Value
CSEFRow = FSEFRow - 2
CSELRow = FSELRow - 2
If CB_Equipment.Value = True Then Sheets("Equipment").Visible = xlSheetVisible _
                                : Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                                : Sheets("Frame Summary").Rows(FSEFRow & ":" & FSELRow).EntireRow.Hidden = False _
                                : Sheets("Client Summary").Rows(CSEFRow & ":" & CSELRow).EntireRow.Hidden = False _
                                : Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect _
                                : Sheets("Quote").Range("A17").Value = "=Equipment!A6" _
                                : Sheets("Quote").Range("B17").Value = "=Equipment!B6" _
                                : Sheets("Quote").Range("H17").Value = "" _
                                : Sheets("Quote").Range("I17").Value = "" _
                                : Sheets("Quote").Range("A18").Value = "=Equipment!A7" _
                                : Sheets("Quote").Range("B18").Value = "=Equipment!B7" _
                                : Sheets("Quote").Range("J18").Value = "=Equipment!C7" _
                                : Sheets("Quote").Range("K18").Value = "=Equipment!L7"
If CB_Equipment.Value = False Then Sheets("Equipment").Visible = xlSheetHidden _
                                 : Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                                 : Sheets("Frame Summary").Rows(FSEFRow & ":" & FSELRow).EntireRow.Hidden = True _
                                 : Sheets("Client Summary").Rows(CSEFRow & ":" & CSELRow).EntireRow.Hidden = True _
                                 : Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect _
                                 : CB_Itemise_BOM.Value = False
Call CB_Itemise_BOM_Click
Application.ScreenUpdating = True
End Sub

Public Sub CB_Itemise_BOM_Click()

Dim FEQS As Long 'Equip Start
Dim FEQE As Long 'Equip End
Dim FBoMS As Long 'BoM Start
Dim FBoME As Long 'BoM End
Dim FESubT As Long 'Equipment Subtotal Row
Dim CEQS As Long 'Equip Start
Dim CEQE As Long 'Equip End
Dim CBoMS As Long 'BoM Start
Dim CBoME As Long 'BoM End
Dim CESubT As Long 'Equipment Subtotal Row

If CB_Equipment.Value = False Then Exit Sub

FEQS = Sheets("Frame Summary").Range("F2").Value
FEQE = Sheets("Frame Summary").Range("F3").Value
FBoMS = FEQS + 1
FBoME = FEQE - 2
FESubT = FEQE - 1
CEQS = FEQS - 2
CEQE = FEQE - 2
CBoMS = CEQS + 2
CBoME = CEQE - 1
CESubT = CEQE

Application.ScreenUpdating = False

If CB_Itemise_BOM = False Then Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                                : Sheets("Frame Summary").Rows(FESubT).EntireRow.Hidden = False _
                                : Sheets("Frame Summary").Rows(FBoMS & ":" & FBoME).EntireRow.Hidden = True _
                                : Sheets("Client Summary").Rows(CESubT).EntireRow.Hidden = False _
                                : Sheets("Client Summary").Rows(CBoMS & ":" & CBoME).EntireRow.Hidden = True _
                                : Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
If CB_Itemise_BOM = True Then Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                                : Sheets("Frame Summary").Rows(FESubT).EntireRow.Hidden = True _
                                : Sheets("Frame Summary").Rows(FBoMS & ":" & FBoME).EntireRow.Hidden = False _
                                : Sheets("Client Summary").Rows(CESubT).EntireRow.Hidden = True _
                                : Sheets("Client Summary").Rows(CBoMS & ":" & CBoME).EntireRow.Hidden = False _
                                : Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect

Application.ScreenUpdating = True
End Sub

'Itemise Tasks -> If True: Unhide entire section // If False: Hide Entire Section + Unhide Segment Titles
Public Sub CB_Itemised_Tasks_Click()
Application.ScreenUpdating = False

Dim FSS As Long
Dim FSE As Long
Dim CSS As Long
Dim CSE As Long
Dim FSbcS As Long
Dim FSbcE As Long
Dim CSbcS As Long
Dim CSbcE As Long

FSS = Sheets("Frame Summary").Range("D2").Value + 1 'Frame Services Start -> Leaves title Row
FSE = Sheets("Frame Summary").Range("D3").Value - 1 'Frame Services End -> Leaves blank divider
CSS = FSS - 1 'Client Summary Start
CSE = FSE - 1 'Client Summary End
FSbcS = Sheets("Frame Summary").Range("E2").Value + 1 'Subcontractor Start -> Leaves title Row
FSbcE = Sheets("Frame Summary").Range("E3").Value - 1 'Subcontractor End -> Leaves blank divider
CSbcS = FSbcS - 1 'Client Summary Subcontractor Start
CSbcE = FSbcE - 1 'Client Summary Subcontractor End

Sheets("Scope").Unprotect
Sheets("Frame Summary").Unprotect
Sheets("Client Summary").Unprotect

'Two Loop Structure -> Frame Loop // Subc Loop (If Subc = Visible Then ...)
If CB_Itemised_Tasks = False Then Sheets("Frame Summary").Rows(FSS & ":" & FSbcE).EntireRow.Hidden = True _
                                : Sheets("Client Summary").Rows(CSS & ":" & CSbcE).EntireRow.Hidden = True _
                                : Call Sheets("Scope").ServiceSummaryUnhide
                                
If CB_Itemised_Tasks = True Then Sheets("Frame Summary").Rows(FSS & ":" & FSbcE).EntireRow.Hidden = False _
                                : Sheets("Client Summary").Rows(CSS & ":" & CSbcE).EntireRow.Hidden = False _
                                : Call Sheets("Scope").UnusedHide
                                
If Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value).Hidden = False Then Sheets("Frame Summary").Rows(Sheets("Frame Summary").Range("D3") & ":" & Sheets("Frame Summary").Range("E2")).Hidden = False
If Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value).Hidden = False Then Sheets("Client Summary").Rows(Sheets("Frame Summary").Range("D3") & ":" & Sheets("Frame Summary").Range("E2")).Offset(-1).Hidden = False
If Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value).Hidden = True Then Sheets("Frame Summary").Rows(Sheets("Frame Summary").Range("E2").Value).Hidden = True
If Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value).Hidden = True Then Sheets("Client Summary").Rows(Sheets("Frame Summary").Range("D3") & ":" & Sheets("Frame Summary").Range("E2")).Offset(-1).Hidden = True

Sheets("Frame Summary").Rows(Sheets("Frame Summary").Range("D3").Value).Hidden = False 'Reveals Row Separator
Sheets("Scope").Protect
Sheets("Frame Summary").Protect
Sheets("Client Summary").Protect
                                                                
Application.ScreenUpdating = True
End Sub

Private Sub CB_SiteMisc_Click()
Application.ScreenUpdating = False
Dim FMiscSt As Long
Dim FMiscEn As Long
Dim CMiscSt As Long
Dim CMiscEn As Long

FMiscSt = Sheets("Frame Summary").Range("G2").Value
FMiscEn = Sheets("Frame Summary").Range("G3").Value
CMiscSt = FMiscSt - 2
CMiscEn = FMiscEn - 1

If CB_SiteMisc.Value = True Then Sheets("Miscellaneous").Visible = xlSheetVisible _
                            : Sheets("Permits").Visible = xlSheetVisible _
                            : Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                            : Sheets("Frame Summary").Rows(FMiscSt & ":" & FMiscEn).EntireRow.Hidden = False _
                            : Sheets("Client Summary").Rows(CMiscSt & ":" & CMiscEn).EntireRow.Hidden = False _
                            : Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                                                                                
If CB_SiteMisc.Value = False Then Sheets("Miscellaneous").Visible = xlSheetHidden _
                            : Sheets("Permits").Visible = xlSheetHidden _
                            : Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                            : Sheets("Frame Summary").Rows(FMiscSt & ":" & FMiscEn).EntireRow.Hidden = True _
                            : Sheets("Client Summary").Rows(CMiscSt & ":" & CMiscEn).EntireRow.Hidden = True _
                            : Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                            
Application.ScreenUpdating = True
End Sub

Private Sub CB_Third_Parties_Click()

Application.ScreenUpdating = False
Dim InputData As Variant

Dim FRow As Long
Dim lRow As Long
Dim CSFRow As Long
Dim CSLRow As Long
Dim FS1Row As Long
Dim FS2Row As Long
Dim FS3Row As Long
Dim FS4Row As Long
Dim FS5Row As Long
Dim FS6Row As Long
Dim FS7Row As Long
Dim FS8Row As Long
Dim FS9Row As Long

FRow = Sheets("Frame Summary").Range("E2").Value 'Subc Start
lRow = Sheets("Frame Summary").Range("E3").Value 'Subc Start
FS1Row = Sheets("Scope").Range("G" & Sheets("Scope").Range("B4").Value) - 1 'Subc1 ends at Subc2 Summary Start Row - 1
FS2Row = Sheets("Scope").Range("G" & Sheets("Scope").Range("C4").Value) - 1 'Subc2 ends at Subc3 Summary Start Row - 1
FS3Row = Sheets("Scope").Range("G" & Sheets("Scope").Range("D4").Value) - 1 'Subc3 ends at Subc4 Summary Start Row - 1
FS4Row = Sheets("Scope").Range("G" & Sheets("Scope").Range("E4").Value) - 1 'Subc4 ends at Subc5 Summary Start Row - 1
FS5Row = Sheets("Scope").Range("G" & Sheets("Scope").Range("F4").Value) - 1 'Subc5 ends at Subc6 Summary Start Row - 1
FS6Row = Sheets("Scope").Range("G" & Sheets("Scope").Range("G4").Value) - 1 'Subc6 ends at Subc7 Summary Start Row - 1
FS7Row = Sheets("Scope").Range("G" & Sheets("Scope").Range("H4").Value) - 1 'Subc7 ends at Subc8 Summary Start Row - 1
FS8Row = Sheets("Scope").Range("G" & Sheets("Scope").Range("I4").Value) - 1 'Subc8 ends at Subc9 Summary Start Row - 1
FS9Row = Sheets("Scope").Range("G" & Sheets("Scope").Range("J4").Value) - 1 'Subc9 ends at Subc10 Summary Start Row - 1
CSFRow = FRow - 2
CSLRow = lRow - 2

If CB_Third_Parties.Value = True Then Run InputData
If CB_Third_Parties.Value = False Then Sheets(9).Visible = xlSheetHidden: Sheets(10).Visible = xlSheetHidden _
                        : Sheets(11).Visible = xlSheetHidden: Sheets(12).Visible = xlSheetHidden _
                        : Sheets(13).Visible = xlSheetHidden: Sheets(14).Visible = xlSheetHidden _
                        : Sheets(15).Visible = xlSheetHidden: Sheets(16).Visible = xlSheetHidden _
                        : Sheets(17).Visible = xlSheetHidden: Sheets(18).Visible = xlSheetHidden _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("F3").Value)).EntireRow.Hidden = True _
                        : Sheets("Frame Summary").Rows(FRow & ":" & lRow).EntireRow.Hidden = True _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & CSLRow).EntireRow.Hidden = True _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect: Exit Sub

InputData = Application.InputBox(Prompt:="Number of Third Parties:", Title:="Third Parties", Type:=1)
    If InputData = 1 Then Sheets(9).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("B4").Value - 1)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & FS1Row).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & FRow).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                    
    If InputData = 2 Then Sheets(9).Visible = xlSheetVisible: Sheets(10).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("C4").Value - 1)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & FS2Row).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & FS1Row).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                        
    If InputData = 3 Then Sheets(9).Visible = xlSheetVisible: Sheets(10).Visible = xlSheetVisible _
                        : Sheets(11).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("D4").Value - 1)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & FS3Row).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & FS2Row).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                        
    If InputData = 4 Then Sheets(9).Visible = xlSheetVisible: Sheets(10).Visible = xlSheetVisible _
                        : Sheets(11).Visible = xlSheetVisible: Sheets(12).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("E4").Value - 1)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & FS4Row).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & FS3Row).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                        
    If InputData = 5 Then Sheets(9).Visible = xlSheetVisible: Sheets(10).Visible = xlSheetVisible _
                        : Sheets(11).Visible = xlSheetVisible: Sheets(12).Visible = xlSheetVisible _
                        : Sheets(13).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("F4").Value - 1)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & FS5Row).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & FS4Row).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                        
    If InputData = 6 Then Sheets(9).Visible = xlSheetVisible: Sheets(10).Visible = xlSheetVisible _
                        : Sheets(11).Visible = xlSheetVisible: Sheets(12).Visible = xlSheetVisible _
                        : Sheets(13).Visible = xlSheetVisible: Sheets(14).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("G4").Value - 1)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & FS6Row).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & FS5Row).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                        
    If InputData = 7 Then Sheets(9).Visible = xlSheetVisible: Sheets(10).Visible = xlSheetVisible _
                        : Sheets(11).Visible = xlSheetVisible: Sheets(12).Visible = xlSheetVisible _
                        : Sheets(13).Visible = xlSheetVisible: Sheets(14).Visible = xlSheetVisible _
                        : Sheets(15).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("H4").Value - 1)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & FS7Row).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & FS6Row).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                        
    If InputData = 8 Then Sheets(9).Visible = xlSheetVisible: Sheets(10).Visible = xlSheetVisible _
                        : Sheets(11).Visible = xlSheetVisible: Sheets(12).Visible = xlSheetVisible _
                        : Sheets(13).Visible = xlSheetVisible: Sheets(14).Visible = xlSheetVisible _
                        : Sheets(15).Visible = xlSheetVisible: Sheets(16).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("I4").Value - 1)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & FS8Row).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & FS7Row).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                        
    If InputData = 9 Then Sheets(9).Visible = xlSheetVisible: Sheets(10).Visible = xlSheetVisible _
                        : Sheets(11).Visible = xlSheetVisible: Sheets(12).Visible = xlSheetVisible _
                        : Sheets(13).Visible = xlSheetVisible: Sheets(14).Visible = xlSheetVisible _
                        : Sheets(15).Visible = xlSheetVisible: Sheets(16).Visible = xlSheetVisible _
                        : Sheets(17).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("J4").Value - 1)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & FS9Row).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & FS8Row).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                        
   If InputData = 10 Then Sheets(9).Visible = xlSheetVisible: Sheets(10).Visible = xlSheetVisible _
                        : Sheets(11).Visible = xlSheetVisible: Sheets(12).Visible = xlSheetVisible _
                        : Sheets(13).Visible = xlSheetVisible: Sheets(14).Visible = xlSheetVisible _
                        : Sheets(15).Visible = xlSheetVisible: Sheets(16).Visible = xlSheetVisible _
                        : Sheets(17).Visible = xlSheetVisible: Sheets(18).Visible = xlSheetVisible _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("F3").Value)).EntireRow.Hidden = False _
                        : Sheets("Frame Summary").Rows(FRow & ":" & lRow).EntireRow.Hidden = False _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & CSLRow).EntireRow.Hidden = False _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect
                        
    If InputData <= 0 Then Sheets(9).Visible = xlSheetHidden
                        
   If InputData = "" Then Sheets(9).Visible = xlSheetHidden
    
                       If Sheets(9).Visible = xlSheetHidden Then Sheets(10).Visible = xlSheetHidden _
                        : Sheets(11).Visible = xlSheetHidden: Sheets(12).Visible = xlSheetHidden _
                        : Sheets(13).Visible = xlSheetHidden: Sheets(14).Visible = xlSheetHidden _
                        : Sheets(15).Visible = xlSheetHidden: Sheets(16).Visible = xlSheetHidden _
                        : Sheets(17).Visible = xlSheetHidden: Sheets(18).Visible = xlSheetHidden _
                        : Sheets("Scope").Unprotect: Sheets("Frame Summary").Unprotect: Sheets("Client Summary").Unprotect _
                        : Sheets("Scope").Rows(Sheets("Scope").Range("E3").Value & ":" & (Sheets("Scope").Range("F3").Value)).EntireRow.Hidden = True _
                        : Sheets("Frame Summary").Rows(FRow & ":" & lRow).EntireRow.Hidden = True _
                        : Sheets("Client Summary").Rows(CSFRow & ":" & CSLRow).EntireRow.Hidden = True _
                        : Sheets("Scope").Protect: Sheets("Frame Summary").Protect: Sheets("Client Summary").Protect

If Sheets("Document Control").CB_Itemised_Tasks = False Then Call Sheets("Document Control").CB_Itemised_Tasks_Click
If Sheets("Document Control").CB_Itemised_Tasks = True Then Call Sheets("Document Control").CB_Itemised_Tasks_Click

Application.ScreenUpdating = True
End Sub



