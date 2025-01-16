                ' Find the table
                Dim table As WebElement
                Set table = driver.FindElementById("tranList")
                
                ' Get all rows
                Dim rows_REP As WebElements
                Set rows_REP = table.FindElementsByTag("tr")
                
                ' Check for duplicates in Column 2
                Dim data As Collection
                Set data = New Collection
                
                Dim Z As Integer
                Dim cellText As String
                Dim duplicateFound As Boolean
                
                On Error Resume Next ' Use error handling outside the loop
                
                duplicateFound = False ' Initialize the duplicate found flag
                
                For Z = 2 To rows_REP.count ' Assuming first row is header
                    cellText = rows_REP(Z).FindElementsByTag("td")(2).Text ' Column 2
                    
                    ' Attempt to add the cellText to the collection
                    If cellText = " REP" Then ' Only check for duplicates if the text is "REP"
                        data.Add cellText, CStr(cellText)
                        
                        If Err.Number <> 0 Then ' Duplicate found
                            duplicateFound = True
                            Debug.Print "Duplicate found: " & cellText
                            
                            ' Handle the specific case for "REP"
                            Debug.Print "Transfer to Merch Not Done"
                            ' Add your specific action here for "REP"
                            
                            Err.Clear ' Clear the error
                        End If
                    End If
                Next Z
                
                On Error GoTo 0 ' Resume normal error handling
                
                ' Check if no duplicate was found
                If Not duplicateFound Then
                    Debug.Print "No duplicate found for REP."
                    ' Add your specific action here if no duplicate is found for "REP"
                    GoTo Process_ThisCase
                Else
                    GoTo TransferToMerch_NotDone
                End If
                
                ' Label for processing this case
GoTo Process_ThisCase:
                    ' Add your specific action for processing the case here
                    Debug.Print "Processing this case..."
                    Exit Sub ' Ensure to exit to avoid falling through
                
                ' Label for transferring to merch not done
GoTo TransferToMerch_NotDone:
                    ' Add your specific action for transferring to merch not done here
                    Debug.Print "Transfer to Merch Not Done"
                    Exit Sub ' Ensure to exit to avoid falling through
