'PDF rename
' Search for Duplicate PDF assign numbers to duplicate found
' Rename PDF in folder with new pdf_1
'file location in below path
'"C:\2_VBA_Selenium\VISA_2025\PDF Files Renames_with_Nos.xlsm"



'Run Step - 1
Sub NumberDuplicates_Step1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim dict As Object
    
    Set ws = ThisWorkbook.Sheets("Sheet2") ' Change Sheet1 to your sheet name
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each cell In ws.Range("A1:A" & lastRow)
        If Not dict.exists(cell.Value) Then
            dict.Add cell.Value, 1
            cell.Offset(0, 1).Value = 1 ' Start with number 1 for the first occurrence
        Else
            dict(cell.Value) = dict(cell.Value) + 1
            cell.Offset(0, 1).Value = dict(cell.Value) ' Increment the number for duplicates
        End If
    Next cell
End Sub
'Run Step - 2
Sub rename_pdf_Step2()
Dim LR As Long, sht As String, FPath1 As String, FName1 As String
sht = Sheets("RawData").Name
LR = Sheets(sht).Range("A" & Rows.Count).End(xlUp).Row
FPath = Range("J1").Text
    For i = 6 To LR
        FName1 = FPath & "\" & Sheets(sht).Range("A" & i).Value & ".pdf"
        fname2 = FPath & "\" & Sheets(sht).Range("B" & i).Value & ".pdf"
        Name FName1 As fname2
    Next
MsgBox "Done Rename File"
End Sub


''''''''''Additional steps

Sub get_pdf_name()
Dim FR As Long, sh As String, FPath As String, FName As String
sh = Sheets("RawData").Name
FR = Sheets(sh).Range("A" & Rows.Count).End(xlUp).Row + 1
FPath = Range("J1").Text
FName = Dir(FPath & "\" & "*.pdf*")
Do While Len(FName)
    Sheets(sh).Range("A" & FR).Value = Left(FName, Len(FName) - 4)
    FR = FR + 1
FName = Dir
Loop
MsgBox "Done Extract PDF Names"
End Sub
Sub rename_pdf()
Dim LR As Long, sht As String, FPath1 As String, FName1 As String
sht = Sheets("RawData").Name
LR = Sheets(sht).Range("A" & Rows.Count).End(xlUp).Row
FPath = Range("J1").Text
    For i = 6 To LR
        FName1 = FPath & "\" & Sheets(sht).Range("A" & i).Value & ".pdf"
        fname2 = FPath & "\" & Sheets(sht).Range("B" & i).Value & ".pdf"
        Name FName1 As fname2
    Next
MsgBox "Done Rename File"
End Sub
