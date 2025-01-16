Sub GetValuesFromDynamicRange()
    Dim cell As Range
    Dim cellValue As Variant
    Dim lastRow As Long
    
    ' Find the last row with data in column A
    lastRow = Sheets("Sheet1").Cells(Rows.count, 1).End(xlUp).row
    
    ' Loop through each cell in the range A1 to the last row in column A
    For Each cell In Sheets("Sheet1").Range("A1:A" & lastRow)
        cellValue = cell.Value
        'MsgBox "The value in cell " & cell.Address & "---- Case No is ---- " & cellValue
        MsgBox "Current Case ---- " & cellValue, vbInformation, "Case Number"
    Next cell
End Sub
