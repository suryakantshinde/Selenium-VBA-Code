Private Sub Main()
On Error GoTo EH
    
    Dim i As Long
    For i = 1 To 2
        ' Generate type mismatch error
         Error 13
continue:
    Next i

done:
    Exit Sub
EH:
    Debug.Print i, Err.Description
    On Error GoTo -1 ' clear the error
    GoTo continue ' return to the code
End Sub
