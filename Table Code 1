Dim tbl1 As WebElement
Dim row1 As WebElement
Dim icon1 As WebElement
Dim refftype1 As WebElement
Dim reffvalue1 As WebElement
Dim reffvalue2 As WebElement
Dim txtvalue1 As String
row_count = 0
Set tbl1 = driver.FindElementById("accountListTable")
For Each row1 In tbl1.FindElementsByTag("tr")
        If row1.FindElementsByTag("td").count > 1 Then   '...................1
            Set icon1 = row1.FindElementsByTag("td")(1)
            Set refftype1 = row1.FindElementsByTag("td")(2)
            Set reffvalue1 = row1.FindElementsByTag("td")(2)
            Set reffvalue2 = row1.FindElementsByTag("td")(3)
             
                    If (Trim(reffvalue1.Text)) = "Local Payments Acct" Then '...................2
                           
                             If (Trim(reffvalue2.Text)) = "INR" Then '...................3
                                icon1.Click
                            Else
                            End If '...................3
                            Exit For
                    End If '...................2
                    
        End If '...................1
              row_count = row_count + 1
Next
'-----------------------------------------------------------------------------------
