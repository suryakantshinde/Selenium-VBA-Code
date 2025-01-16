'-----------------------------------------------------------------------------------
'Search for Refund
'-----------------------------------------------------------------------------------
Dim tbl1 As WebElement
Dim row1 As WebElement
Dim icon1 As WebElement
Dim refftype1 As WebElement
Dim reffvalue1 As WebElement
Dim txtvalue1 As String
row_count = 0
Set tbl1 = driver.FindElementByClass("zebratable")
For Each row1 In tbl1.FindElementsByTag("tr")
        If row1.FindElementsByTag("td").count > 5 Then
            Set icon1 = row1.FindElementsByTag("td")(1)
            'Set refftype1 = row.FindElementsByTag("td")(2)
            Set reffvalue1 = row1.FindElementsByTag("td")(5)
            If (Trim(reffvalue1.Text)) = "Refund" Then
             driver.Window.Activate.Close
             driver.SwitchToPreviousWindow
             Range("F" & i).Value = "Record not process as Refund Done"
                GoTo Surya
                Dim tm As Selenium.SelectElement
                'Set tm = driver.FindElementById("actTList_" & row_count).AsSelect
                'If InStr(1, icon1.Text, "Transfer to Merch") > 0 Then
               ' tm.SelectByText "Transfer to Merch"
                Else
               
                'GoTo Surya
                End If
                driver.Wait 1000
Exit For
            'End If
                row_count = row_count + 1
        End If
Next
