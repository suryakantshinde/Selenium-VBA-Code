Private Keys As New Selenium.Keys
Private Assert As New Selenium.Assert
Sub Mastercard_PreArb_VBA_Selenium_Start()
'On Error Resume Next
Dim driver As New Selenium.ChromeDriver
'Dim driver As New Selenium.EdgeDriver
Dim count As Long
Dim fs As String
Dim button As WebElement
Dim i, lr As Long
Sheets("Instructions").Activate
        Set driver = CreateObject("Selenium.ChromeDriver")
        'Set driver = CreateObject("Selenium.EdgeDriver")
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
  '-----------------------------------------------------------'-----------------------------------------------------------
' ********** PIN ENTER AUTOMATIC FROM EMAIL **************

'
Const DontWaitUntilFinished = False, ShowWindow = 1, DontShowWindow = 0, WaitUntilFinished = True
Set oShell = CreateObject("WScript.Shell")
            'change path if needed -----------------------------------------------------------------------------------& Change Path if needed "C:\Selenium_VBA_Projects\Visa_Representment_VBASelenium\PIN\outputpin.txt"
Cmd1 = "cmd /c C:\windows\system32\wscript.exe C:\Pre_Arb_Accept_Selenium\PIN\CreatePin.vbs " & "C:\Pre_Arb_Accept_Selenium\PIN\outputpin.txt"
oShell.Run Cmd1, DontShowWindow, WaitUntilFinished
'-----------------------------------------------------------'-----------------------------------------------------------
'---- Read the text File ----
FSOPasteTextFileContent
'-----------------------------------------------------------'-----------------------------------------------------------
'To Read PIN from Sheets("Instructions").Range("Q1")
'-----------------------------------------------------------'-----------------------------------------------------------
'TextString
Sheets("Instructions").Activate
driver.FindElementByXPath("//*[@id='46aAeiL']").Clear
driver.FindElementByXPath("//*[@id='46aAeiL']").SendKeys Range("Q1")
driver.wait 5000
'click on update pin
'driver.FindElementByXPath("/html/body/div/div[2]/form/div[2]/div/input[1]").Click
'-----------------------------------------------------------'-----------------------------------------------------------
driver.wait 1000
'-----------------------------------------------------------'-----------------------------------------------------------
'==================================================================================================================================
'Click on Data Access Privileges
driver.FindElementByXPath("//*[@id='twofactor']").Click
'check Enhanced Data Access Privileges

'View Card Number
driver.FindElementByXPath("//*[@id='field-view-card-number']").Click

'Download Card Number
driver.FindElementByXPath("//*[@id='field-download-card-number']").Click

'View Bank Account
'driver.FindElementByXPath("//*[@id='field-view-bank-account']").Click

'View Merchant PII Data
'driver.FindElementByXPath("//*[@id='field-view-merchant-pii']").Click

'Update Merchant PII Data
'driver.FindElementByXPath("//*[@id='field-update-merchant-pii']").Click

'View Sensitive Document/Report - PCI
driver.FindElementByXPath("//*[@id='field-view-sens-doc-pci']").Click

'View Sensitive Document/Report - PII
driver.FindElementByXPath("//*[@id='field-view-sens-doc-pii']").Click

'-----------------------------------------------------------'-----------------------------------------------------------
'Click button Update Privileges
driver.FindElementByXPath("/html/body/div[2]/div[1]/div[3]/span/span[1]/span/button").Click
'-----------------------------------------------------------'-----------------------------------------------------------
driver.wait 3000
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
        driver.wait 2000

'04 - Click on Search Button
       driver.FindElementByXPath("//button[@id='search']").Click
       driver.wait 2000
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
'Condition - 1
'***********************************************************************************************
'If Case found and in Status its "Open" or "Closed" then move to next new case = Case Not Processed
'***********************************************************************************************
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
''Search in Table Status - "Open" or "Closed"
Dim tbl1 As WebElement
Dim row1 As WebElement
Dim icon1 As WebElement
Dim refftype1 As WebElement
Dim reffvalue1 As WebElement
Dim txtvalue1 As String
row_count = 0
Set tbl1 = driver.FindElementById("omcbkCaseListTable")
For Each row1 In tbl1.FindElementsByTag("tr")
        If row1.FindElementsByTag("td").count > 3 Then
            Set icon1 = row1.FindElementsByTag("td")(1)
            'Set refftype1 = row.FindElementsByTag("td")(2)
            Set reffvalue1 = row1.FindElementsByTag("td")(3)
            If (Trim(reffvalue1.Text)) = "Closed" Or (Trim(reffvalue1.Text)) = "Open" Then
                'MsgBox "The Case Status is " & reffvalue1.Text, vbInformation
                Range("G" & i).Value = "The Case Status is " & reffvalue1.Text
                Range("G" & i).Interior.Color = vbRed
'-----------------------------------------------------------------------------------
'Condition - 2
'If Case found and in Status its Pre-ArbitrationPre-Arbitration then Process = Case Processed

Else

''01 - Click on record found
      driver.FindElementByXPath("/html/body/div[5]/div[2]/form/div/div[2]/div[2]/div[4]/table/tbody[2]/tr/td[2]").Click
      driver.wait 4000

''02 - Go To Transaction Activity
driver.FindElementByXPath("//a[@href = '#Transaction Activity']").Click

''03 - Click on  Processed Transactions
driver.FindElementByLinkText("Processed Transactions").Click Keys.Control
driver.wait 5000

driver.SwitchToNextWindow
'-----------------------------------------------------------------------------------
'Full Card No. or Last 4 digits: = id - ID_17A
       driver.FindElementById("ID_17A").Clear
       driver.FindElementById("ID_17A").SendKeys Range("B" & i)

'Transaction Type search for -->  Refund (Credit)
       driver.FindElementByXPath("//*[@id='ex2h']/fieldset/table/tbody/tr[2]/td[4]/select").AsSelect.SelectByText (Trim("Refund (Credit)"))
       
'Posting Date - Start: we need to change to Previous Year = id - id_55B
       driver.FindElementById("id_55B").Clear
       driver.FindElementById("id_55B").SendKeys Range("C" & i)
       
'End:Date - id - id_28B
       driver.FindElementById("id_28B").Clear
       driver.FindElementById("id_28B").SendKeys Range("D" & i)

'Click on Search Button
       driver.FindElementByXPath("//button[@id='search']").Click
       driver.wait 5000
'1
If driver.FindElementById("idtopmessage").IsDisplayed = True Then
'GoTo NoResultsFound
            driver.Window.Activate.Close
            driver.SwitchToPreviousWindow
'21 - Click on Tab {Letters/Attachments} id - tab3
driver.FindElementById("tab3").Click

''12 - This Combo-box we need to select from drop down "Accept-MC Pre-Arb"
Dim tbl2 As WebElement
Dim row2 As WebElement
Dim icon2 As WebElement
Dim refftype2 As WebElement
Dim reffvalue2 As WebElement
Dim txtvalue2 As String
row_count = 0
Set tbl2 = driver.FindElementById("letterList")
For Each row2 In tbl2.FindElementsByTag("tr")
        '2
        If row2.FindElementsByTag("td").count > 1 Then
            Set icon2 = row2.FindElementsByTag("td")(1)
            Set refftype2 = row2.FindElementsByTag("td")(2)
            Set reffvalue2 = row2.FindElementsByTag("td")(2)
            'If (Trim(reffvalue.Text) = "PRE-ARBITRATION_19827495.xml") Then
            'If InStr(1, Trim(reffvalue.Text), "PRE-ARBITRATION_*" & ".xlsm") > 0 Then
          '3
            If InStr(1, Trim(reffvalue2.Text), "PRE-ARBITRATION_") > 0 And Right(UCase(Trim(reffvalue2.Text)), 3) = "XML" Then
                icon2.Click
                Dim lfs As Selenium.SelectElement
                Set lfs = driver.FindElementById("actLList_" & row_count).AsSelect
             '4
                If InStr(1, icon2.Text, "Accept-MC Pre-Arb") > 0 Then
                lfs.SelectByText "Accept-MC Pre-Arb"
                Else
                End If
             '4
             driver.wait 1000
Exit For
            End If
            '3
                row_count = row_count + 1
        End If
       '2
Next
'13 - Click on Go Button
    driver.FindElementByXPath("//*[@id='letterList']/tbody[2]/tr[" & row_count + 1 & "]/td[1]/span/input").Click
'14 - Switch to Popup Window
        driver.wait 1000
        driver.SwitchToNextWindow
        driver.wait 1000

        driver.FindElementByXPath("//*[@id='41A']").SendKeys Range("E" & i)
        driver.wait 100

'15 - Click on Save Button
        driver.FindElementByXPath("/html/body/div/div[2]/form/table/tbody/tr[11]/td/input[1]").Click
        driver.wait 5000

'16 - Switch to default Window
        driver.SwitchToPreviousWindow
        driver.wait 1000
''-----------------------------------------------------------------------------------------------------------------------------------
 '21 - Click on Tab {User Notes} id - tab3
driver.FindElementById("tab5").Click
      driver.wait 1000
driver.FindElementByName("44c").SendKeys Range("F" & i)
driver.wait 2000
'click on Add button
driver.FindElementByXPath("/html/body/div[3]/div[2]/form/div/div[9]/table/tbody/tr[1]/td[2]/input").Click
driver.wait 1000
Range("G" & i).Value = "Pre-arbitration CB Accepted"
Range("G" & i).Interior.Color = vbGreen
Else
            driver.Window.Activate.Close
            driver.SwitchToPreviousWindow
            Range("G" & i).Value = "Refund (Credit)"
            Range("G" & i).Interior.Color = vbYellow
'-----------------------------------------------------------------------------------
End If
'1
driver.wait 1000
On Error Resume Next
Exit For
              row_count = row_count + 1
        End If
        End If
Next
driver.wait 1000
'==================================================================================================================================
Next i
'-----------------------------------------------------------'-----------------------------------------------------------
ActiveWorkbook.Save
MsgBox "Completed Successfully", vbInformation
driver.Quit
'------------------------------------------------------------------------------------------------------
Exit Sub
SuryaError:
MsgBox "Contact Surya Error in code", vbCritical
End Sub


'''Read Text File - That is read PIN
Sub FSOPasteTextFileContent()
    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
                    'change path if needed --------------------
    Set FileToRead = FSO.OpenTextFile("C:\Pre_Arb_Accept_Selenium\PIN\outputpin.txt", ForReading) 'add here the path of your text file
    TextString = FileToRead.ReadAll

    FileToRead.Close
    'change Sheet name and column
    ThisWorkbook.Sheets("Instructions").Range("Q1").Value = TextString 'you can specify the worksheet and cell where to paste the text file’s content
End Sub
