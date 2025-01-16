Private Keys As New Selenium.Keys
Private Assert As New Selenium.Assert
Sub Download_Onus_Chargeback_Automation_Run()
On Error Resume Next
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim driver As New Selenium.ChromeDriver
Dim count As Long
Dim fs As String
Dim button As WebElement
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sheets("SignOn").Activate
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Set driver = CreateObject("Selenium.ChromeDriver")
        driver.Get "https://in-ssoca.fiservapp.com/idp/startSSO.ping?PartnerSpId=APM0002648_OMNIPAY-INDIA_IDP"
        driver.Window.Maximize
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Enter User Name
        driver.FindElementByXPath("//*[@id='username']").Clear
        driver.FindElementByXPath("//*[@id='username']").SendKeys Range("B8")
'Enter Password
        driver.FindElementByXPath("//*[@id='password']").Clear
        driver.FindElementByXPath("//*[@id='password']").SendKeys Range("B12")
'Click on Login Button
        driver.FindElementByXPath("//*[@id='signOnButton']").Click
        ' allow to load the results page
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  MsgBox "Please do..." & vbCrLf & vbCrLf & "1 - PingID multi-factor Authentication" & vbCrLf & vbCrLf & "2 - Update Privileges if needed" & vbCrLf & vbCrLf & "Then Click OK button to continue running automation", vbInformation, "PingID-Authentication"
  driver.wait 5000
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'7 - Click on Merchant Activity
driver.FindElementByXPath("//li/a/span[contains(text(),'Transaction Activity')]", 10000).Click
'8 - Click on Authorisation History
driver.FindElementByXPath("//li/a[contains(text(),'Processed Transactions')]", 10000).Click
driver.wait 1000
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sheets("RawData").Activate
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'postingDateStart=id:id_55B
driver.FindElementById("id_55B", 10000).Clear
driver.FindElementById("id_55B").SendKeys Sheets("Rawdata", 10000).Range("B1")
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'postingDateEnd=id:id_28B
driver.FindElementById("id_28B").Clear
driver.FindElementById("id_28B").SendKeys Sheets("Rawdata", 10000).Range("D1")
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim i, lr As Long
lr = ThisWorkbook.ActiveSheet.Cells(Rows.count, 1).End(xlUp).row
'1
For i = 3 To lr
        Sheets("RawData").Activate
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        driver.FindElementByXPath("//*[@id='ex2h']/fieldset/table/tbody/tr[3]/td[4]/input", 10000).Clear
        driver.FindElementByXPath("//*[@id='ex2h']/fieldset/table/tbody/tr[3]/td[4]/input", 10000).SendKeys Range("A" & i).Text
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Click on Search Button
        driver.FindElementByXPath("//button[@id='search']", 10000).Click
        driver.Timeouts.ImplicitWait = 5000
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '2
        If driver.FindElementById("idtopmessage").IsDisplayed = True Then
                Sheets("RawData").Range("B" & i) = "No Record Found"
                Range("B" & i).Font.Color = vbRed
                Else
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Dim tbl As WebElement
                Dim row As WebElement
                Dim icon As WebElement
                Dim refftype As WebElement
                Dim reffvalue As WebElement
                Dim txtvalue As String
                Set tbl = driver.FindElementByClass("zebratable")
                '3
                For Each row In tbl.FindElementsByTag("tr")
                    '4
                    If row.FindElementsByTag("td").count > 0 Then
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                        Set MerchantName = row.FindElementsByTag("td")(21)
                        Set Mids = row.FindElementsByTag("td")(1)
                        Set TID = row.FindElementsByTag("td")(29)
                        Set CardNo = row.FindElementsByTag("td")(6)
                        Set Merc_Tran_Ref = row.FindElementsByTag("td")(18)
                        Set Batch_No = row.FindElementsByTag("td")(2)
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            Sheets("RawData").Range("B" & i) = MerchantName.Text
                            Sheets("RawData").Range("C" & i) = "'" & Mids.Text
                            Sheets("RawData").Range("D" & i) = "'" & TID.Text
                            Sheets("RawData").Range("E" & i) = "'" & CardNo.Text
                            Sheets("RawData").Range("F" & i) = "'" & Merc_Tran_Ref.Text
                            Sheets("RawData").Range("G" & i) = "'" & Batch_No.Text
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    '4
                    End If
                            driver.wait 100
                '3
                Next
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                driver.wait 5000
                ActiveWorkbook.Save
        '2
        End If
'1
Next i
driver.wait 1000
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
driver.Quit
MsgBox "Completed Successfully", vbInformation
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
End Sub


