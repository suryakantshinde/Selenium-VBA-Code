Private Keys As New Selenium.Keys
Private Assert As New Selenium.Assert
Sub Merchant_Debit_FDMSA_VROL_Acceptance_Start()
'On Error GoTo SuryaError
Dim driver As New Selenium.ChromeDriver
Dim count As Long
Dim fs As String
Dim button As WebElement
Dim i, lr As Long
Sheets("Instructions").Activate
        Set driver = CreateObject("Selenium.ChromeDriver")
        driver.Get "https://www2.omnipaygroup.com/ramtool"
'01 - Max Crome windows
        driver.Window.Maximize
'02 - Enter User Name
        driver.FindElementById("69").SendKeys Range("C3")
'03 - Enter Password
        driver.FindElementById("76").SendKeys Range("C4")
'04 - Click on Login Button
        driver.FindElementByXPath("//input[@value='Login']").Click
'-----------------------------------------------------------'-----------------------------------------------------------
     ' Allow to load the results page
        driver.wait 2000
'-----------------------------------------------------------'-----------------------------------------------------------
        
Sheets("RawData").Activate
lr = ThisWorkbook.ActiveSheet.Cells(Rows.count, 1).End(xlUp).row
For i = 2 To lr
'01 - Click on Chargebacks
       driver.FindElementByXPath("//a[@href = '#Chargebacks']").Click

'02 - Click on Case List
       driver.FindElementByLinkText("Case List").Click
       driver.wait 2000
       
'03 - Insert - Case ID/VROL/MCM ID
       'Sheets("MC").Activate
       driver.FindElementById("ID18e").Clear
       driver.FindElementById("ID18e").SendKeys Range("A" & i)

'04 - Click on Search Button
       driver.FindElementByXPath("//button[@id='search']").Click
       driver.wait 2000

'05 - Click on record found
        driver.FindElementByXPath("/html/body/div[5]/div[2]/form/div/div[2]/div[2]/div[4]/table/tbody[2]/tr/td[2]").Click
        driver.wait 4000

'06 - Click on Transaction tab
'     Choose Transfer to Merchant from drop down list (of CBK1)
        driver.FindElementById("actTList_0").Click

'07 - This Combo-box we need to select from drop down "VCR Accept Decision"
Dim tbl1 As WebElement
Dim row1 As WebElement
Dim icon1 As WebElement
Dim refftype1 As WebElement
Dim reffvalue1 As WebElement
Dim txtvalue1 As String
row_count = 0
Set tbl1 = driver.FindElementById("tranList")
For Each row1 In tbl1.FindElementsByTag("tr")
        If row1.FindElementsByTag("td").count > 0 Then
            Set icon1 = row1.FindElementsByTag("td")(1)
            'Set refftype1 = row.FindElementsByTag("td")(2)
            Set reffvalue1 = row1.FindElementsByTag("td")(2)
            If (Trim(reffvalue1.Text)) = "CBK1" Then
                icon1.Click
                Dim tm As Selenium.SelectElement
                Set tm = driver.FindElementById("actTList_" & row_count).AsSelect
                If InStr(1, icon1.Text, "Transfer to Merch") > 0 Then
                tm.SelectByText "Transfer to Merch"
                Else
                GoTo Surya
                End If
                driver.wait 1000
Exit For
            End If
                row_count = row_count + 1
        End If
Next

'08 - Click on Go Button
    driver.FindElementByXPath("//*[@id='tranList']/tbody[2]/tr[" & row_count + 1 & "]/td[1]/span/input").Click

'09 - Switch to Popup Window
        driver.wait 5000
        driver.SwitchToNextWindow
        driver.wait 5000
        
'10 - Click on Save Button
        driver.FindElementById("saveBtn").Click
        driver.wait 5000
        
'11 - Switch to default Window
        driver.SwitchToPreviousWindow
'-----------------------------------------------------------------------------------------------------------------------------------
Surya:
'12 - Click on Tab {Letters/Attachments} id - tab3
        driver.FindElementById("tab3").Click

''Criteria - 1
Dim tbl2 As WebElement
Dim row2 As WebElement
Dim icon2 As WebElement
Dim refftype2 As WebElement
Dim reffvalue2 As WebElement
Dim txtvalue2 As String
row_count = 0
Set tbl2 = driver.FindElementById("letterList")
For Each row2 In tbl2.FindElementsByTag("tr")
        If row2.FindElementsByTag("td").count > 0 Then
            Set icon2 = row2.FindElementsByTag("td")(1)
            'Set refftype = row.FindElementsByTag("td")(2)
            Set reffvalue2 = row2.FindElementsByTag("td")(2)
            If (Trim(reffvalue2.Text)) = "VCRAcceptDecision.xml" Then
                GoTo Criteria2
            End If
                 driver.wait 1000
Exit For
            End If
                row_count = row_count + 1
Next

''Criteria - 2
'Table
' Find Tran.Type - Purchase  Txn Kind - PRE
'--------------------------------------------------------
Dim tbl As WebElement
Dim row As WebElement
Dim icon As WebElement
Dim refftype As WebElement
Dim reffvalue As WebElement
Dim txtvalue As String
row_count = 0
Set tbl = driver.FindElementById("letterList")
For Each row In tbl.FindElementsByTag("tr")
        If row.FindElementsByTag("td").count > 0 Then
            Set icon = row.FindElementsByTag("td")(1)
            'Set refftype = row.FindElementsByTag("td")(2)
            Set reffvalue = row.FindElementsByTag("td")(2)
            If (Trim(reffvalue.Text) = "VCRDisputeAllocation.xml" Or Trim(reffvalue.Text) = "VCRDisputeCollaboration.xml") Then
                icon.Click
                Dim lfs As Selenium.SelectElement
                Set lfs = driver.FindElementById("actLList_" & row_count).AsSelect
                If InStr(1, icon.Text, "VCR Accept Decision") > 0 Then
                lfs.SelectByText "VCR Accept Decision"
                Else
                GoTo Surya
                End If
                     driver.wait 1000
Exit For
            End If
                row_count = row_count + 1
        End If
Next


'13 - Click on Go Button
 driver.FindElementByXPath("//*[@id='letterList']/tbody[2]/tr[" & row_count + 1 & "]/td[1]/span/input").Click
'-----------------------------------------------------------------------------------------------------------------------------------
'14 - Switch to Popup Window
        driver.wait 5000
        driver.SwitchToNextWindow
        driver.wait 5000
        
'15 - Click on Save Button
        driver.FindElementById("saveBtn").Click
        driver.wait 5000
        
'16 - Switch to default Window
        driver.SwitchToPreviousWindow
Criteria2:
Range("D" & i).Value = "Error in Record"
'17 - Click on Tab - User Notes
        driver.FindElementById("tab5").Click

'18 - Add notes
        driver.FindElementByName("44c").Click
        driver.FindElementByName("44c").SendKeys Range("B" & i)
'19 - Click on Add Button
       driver.FindElementByXPath("/html/body/div[3]/div[2]/form/div/div[9]/table/tbody/tr[1]/td[2]/input").Click
driver.wait 1000
onum = onum + 1
driver.wait 100
'20 - Add Result/status in column C
Range("C" & i).Value = "Record updated"
Next i
'-----------------------------------------------------------'-----------------------------------------------------------
ActiveWorkbook.Save
MsgBox "Completed Successfully", vbInformation
driver.Quit
'------------------------------------------------------------------------------------------------------
'Exit Sub
'SuryaError:
'MsgBox "Contact Surya Error in code", vbCritical
End Sub



