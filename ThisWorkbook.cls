VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const sFilePath = "C:\Users\Leo\Documents\GitHub\FET\Sheet"
Const strDelim = ","

Sub Main()
    CreateTxt_Output
    ExportVisualBasicCode
End Sub

Public Sub CreateTxt_Output()
Dim ws As Worksheet
Dim rng1 As Range
Dim X
Dim lRow As Long
Dim lCol As Long
Dim strTmp As String
Dim lFnum As Long
Dim lngCnt As Long
Dim strOut As String

lFnum = FreeFile
For Each ws In ActiveWorkbook.Worksheets
    lngCnt = lngCnt + 1
    Open (sFilePath & lngCnt & ".csv") For Output As lFnum
    'test that sheet has been used
    Set rng1 = ws.UsedRange
    If Not rng1 Is Nothing Then
        'only multi-cell ranges can be written to a 2D array
        If rng1.Cells.count > 1 Then
            X = ws.UsedRange.Formula
            For lRow = 1 To UBound(X, 1)
                strOut = IIf(InStr(X(lRow, 1), strDelim) > 0, """" & X(lRow, 1) & """", X(lRow, 1))
                For lCol = 2 To UBound(X, 2)
                    'write each line to CSV
                    strOut = strOut & (strDelim & IIf(InStr(X(lRow, lCol), strDelim) > 0, """" & X(lRow, lCol) & """", X(lRow, lCol)))
                Next lCol
                Print #lFnum, strOut
            Next lRow
        Else
            Print #lFnum, IIf(InStr(rng1.Formula, strDelim) > 0, """" & rng1.Formula & """", rng1.Formula)
            End If
    End If
    Close lFnum
    Debug.Print "Exported " & sFilePath & lngCnt & ".csv"
Next ws
'MsgBox "Done!", vbOKOnly
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
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    'Dim fso As New FileSystemObject
    
    directory = "C:\Users\Leo\Documents\GitHub\FET"
    count = 0
    
    'If Not fso.FolderExists(directory) Then
        'Call fso.CreateFolder(directory)
    'End If
    'Set fso = Nothing
    
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
    
    'Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    'Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub



