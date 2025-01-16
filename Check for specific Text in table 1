Dim tbl As WebElement
Dim row As WebElement
Dim icon As WebElement
Dim refftype As WebElement
Dim reffvalue As WebElement
Dim txtvalue As String
row_count = 0
Set tbl = driver.FindElementByXPath("/html/body/div[4]/div[2]/form/div/div[6]/div[1]")
For Each row In tbl.FindElementsByTag("tr")
        '1
        If row.FindElementsByTag("td").count > 1 Then
            Set icon = row.FindElementsByTag("td")(1)
            Set refftype = row.FindElementsByTag("td")(1)
            Set reffvalue = row.FindElementsByTag("td")(2)
           '2
            If (Trim(reffvalue.Text)) = "Key Merchant" Then
                        icon.Click
                        GoTo Found_Key_Merchant
Exit For
            Else
            '2
            End If
        '1
        End If
row_count = row_count + 1
Next
'================================================================================================
' Add Key Merchant if not present in Table
  GoTo Not_Found_Key_Merchant
'================================================================================================
continue:
Next i
Found_Key_Merchant:
          'icon.Click
          driver.Wait 100
          driver.FindElementByLinkText("Update").Click
          driver.Wait 100
          'Clear Property Value
          driver.FindElementByXPath("//*[@id='ID205aby']").Clear
          'Insert Property Value
           driver.FindElementByXPath("//*[@id='ID205aby']").SendKeys Range("B" & i)
           driver.Wait 1000
          driver.FindElementByClass("Update").Click
          driver.Wait 1000
          Range("E" & i).Value = "Record updated successfully"
GoTo continue
'================================================================================================
Not_Found_Key_Merchant:
      'Click on Add and Update Properties
      driver.FindElementByXPath("//*[@id='btnAdd']").Click
      'click on check box
      driver.FindElementByName("SEL_042").Click
      'enter value in text box
      driver.FindElementByXPath("//*[@id='ID205aby_042']").SendKeys Range("B" & i)
      'Click on Add Button
      'driver.FindElementByClass("Add").Click
      driver.FindElementByXPath("/html/body/div[5]/div[2]/div/div[1]/div[3]/span/span[1]/span/button").Click
      driver.Wait 1000
      Range("E" & i).Value = "Record updated successfully"
GoTo continue
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
 'Exit Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
ActiveWorkbook.Save
MsgBox "Completed Successfully", vbInformation
driver.Quit
'------------------------------------------------------------------------------------------------------
End Sub
