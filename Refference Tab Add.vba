Private Keys As New Selenium.Keys
Private Assert As New Selenium.Assert
Sub Add_Reference_Tab_Start()
'On Error GoTo Run_Next_Record
'==================================================================================================================
Dim driver As New Selenium.ChromeDriver
'Dim driver As New Selenium.EdgeDriver
'==================================================================================================================
Dim count As Long
Dim fs As String
Dim button As WebElement
Dim i, lr As Long
'==================================================================================================================
Sheets("SignOn").Activate
        Set driver = CreateObject("Selenium.ChromeDriver")
        'Set driver = CreateObject("Selenium.EdgeDriver")
'Change Website if Needed
'==================================================================================================================
        'New Omnipay Link
        '-------------------------
        'driver.Get "https://in-ssoca.fiservapp.com/idp/startSSO.ping?PartnerSpId=APM0002648_OMNIPAY-INDIA_IDP"
        
        'Old Omnipay Link
        '-------------------------
        driver.Get "https://www2.omnipaygroup.com/ramtool?"
'==================================================================================================================
'Max Windows
        driver.Window.Maximize
'==================================================================================================================
     'Login for Old Omnipay
     '-------------------------------
            'Enter User Name
                driver.FindElementByXPath("//*[@id='69']").Clear
                driver.FindElementByXPath("//*[@id='69']").SendKeys Range("B8")
            'Enter Password
                driver.FindElementByXPath("//*[@id='76']").Clear
                driver.FindElementByXPath("//*[@id='76']").SendKeys Range("B12")
            'Click on Login Button
                  driver.FindElementByXPath("//input[@value='Login']").Click
'==================================================================================================================
     'Login for New Omnipay
     '-------------------------------
            'Enter User Name
                'driver.FindElementByXPath("//*[@id='username']").Clear
                'driver.FindElementByXPath("//*[@id='username']").SendKeys Range("B8")
            'Enter Password
                'driver.FindElementByXPath("//*[@id='password']").Clear
                'driver.FindElementByXPath("//*[@id='password']").SendKeys Range("B12")
            'Click on Login Button
                 'driver.FindElementByXPath("//*[@id='signOnButton']").Click
'==================================================================================================================
    MsgBox "Please do PingID multi-factor authentication then click on OK button to continue running automation", vbInformation, "PingID-Authentication"
    driver.wait 5000
'==================================================================================================================
'Below select Institution if needed
'==================================================================================================================
    driver.FindElementById("selectinst", "10000").Click
    driver.FindElementByLinkText("00000047 - ICICI MS", 10000).Click
    driver.wait 500
'==================================================================================================================
'check and click TwoFactor
    driver.FindElementByXPath("//*[@id='twofactor']").Click
'==================================================================================================================
    'View Card Number
        ClickPrivilegeIfAvailable driver, "field-view-card-number"
    'Download Card Number
        ClickPrivilegeIfAvailable driver, "field-download-card-number"
    'View Bank Account
        ClickPrivilegeIfAvailable driver, "field-view-bank-account"
    'Update Bank Account
        ClickPrivilegeIfAvailable driver, "field-update-bank-account"
    'View Merchant PII Data
        ClickPrivilegeIfAvailable driver, "field-view-merchant-pii"
    'Update Merchant PII Data
        ClickPrivilegeIfAvailable driver, "field-update-merchant-pii"
    'View Sensitive Document/Report - PCI
        ClickPrivilegeIfAvailable driver, "field-view-sens-doc-pci"
    'View Sensitive Document/Report - PII
        ClickPrivilegeIfAvailable driver, "field-view-sens-doc-pii"
'==================================================================================================================
'Click button Update Privileges
    driver.wait 500
    driver.FindElementByXPath("/html/body/div[2]/div[1]/div[3]/span/span[1]/span/button", 10000).Click
'==================================================================================================================
    Sheets("RawData").Activate
    lr = ThisWorkbook.ActiveSheet.Cells(Rows.count, 1).End(xlUp).row
    For i = 2 To lr
'==================================================================================================================
            'clickonMerchant Administration  link    Merchant Administration
                driver.FindElementByLinkText("Merchant Administration").Click
                driver.wait 100
            'clickonMerchant Application Setup   link    Merchant Maintenance
                driver.FindElementByLinkText("Merchant Maintenance").Click
                driver.wait 100
            'selectMerchant Application List link    Maintain Mantenanace Detail
               driver.FindElementByLinkText("Maintain Merchant Details").Click
'==================================================================================================================
            '01 - Click on Select Merchant Number
                   driver.FindElementByXPath("//button[@id='merchbutton-button']", 10000).Click
                   
            '02 - Enter Merchant Number - //*[@id="id_40A"]
                  driver.FindElementByXPath("//input[@id='id_40A']", 10000).Clear
                  driver.FindElementByXPath("//input[@id='id_40A']", 10000).SendKeys Range("A" & i)
            
            '03 - Click on Change Button - //*[@id="changeMerchBtn"]
                   driver.FindElementByXPath("//button[@id='changeMerchBtn']", 10000).Click
'==================================================================================================================
'Start Writing you code from here |  Start Writing you code from here |   Start Writing you code from here |   Start Writing you code from here |
'==================================================================================================================
'>>>>> Click on References Tab >>>>>
driver.FindElementByLinkText("References", 10000).Click
driver.wait 500
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Tick Check Box - MVV Value
'Search for "MVV Value"
'IF MVV Value is present in below table then
'exit
'else check the box by ticking it
'==================================================================================================================
        Dim tbl1 As WebElement
        Dim row1 As WebElement
        Dim icon1 As WebElement
        Dim refftype1 As WebElement
        Dim reffvalue1 As WebElement
        Dim txtvalue1 As String
        row_count = 0
        Set tbl1 = driver.FindElementById("referenceListTable")
        '<1>
        For Each row1 In tbl1.FindElementsByTag("tr")
                '<2>
                If row1.FindElementsByTag("td").count > 2 Then
                   Set icon1 = row1.FindElementsByTag("td")(1)
                    Set reffvalue1 = row1.FindElementsByTag("td")(2)
                    '<3>
                    If (Trim(reffvalue1.Text)) = "MVV Value" Then '----------------------(Change the Reference Type as per your need)
                        Range("C" & i).Value = "MVV Value-Allready Added"
                        GoTo continue:
'==================================================================================================================
        Exit For
                    '<3>
                    End If
                        row_count = row_count + 1
                '<2>
                End If
        '<1>
        Next
        driver.wait 1000
'==================================================================================================================
''>>>>> Click on Add New References Button >>>>>
driver.Timeouts.ImplicitWait = 2000
driver.FindElementById("addReference", 10000).Click
driver.wait 2000
'==================================================================================================================
'''Click on Check Box to add reference (here i am adding MVV Value)
driver.FindElementByXPath("//*[@id='FB_48FRT_453']/label/input").Click
'==================================================================================================================
'Insert reference value in text box from excel (in front of MVV Value above)
driver.FindElementByXPath("//*[@id='ID48FRT_453']").SendKeys Range("B" & i)   'take value from Column "B"
Range("C" & i).Value = "Record updated - Added References with Value"
'==================================================================================================================
'click on Add
driver.FindElementByXPath("/html/body/div[7]/div[2]/form/div/div[7]/div[1]/div[3]/span/span[1]/span/button").Click
driver.wait 1000
continue:
driver.wait 1000
'==================================================================================================================
'End Writing you code | End Writing you code | End Writing you code | End Writing you code | End Writing you code | End Writing you code |
'==================================================================================================================
Next i
'==================================================================================================================
done:
    Exit Sub
'==================================================================================================================
Run_Next_Record:
    Range("H" & i).Value = "Record not updated - Account Locked"
    driver.wait 2000
'==================================================================================================================
        ActiveWorkbook.Save
'==================================================================================================================
        MsgBox "Completed Successfully", vbInformation
driver.Quit
'==================================================================================================================
End Sub

Private Sub ClickPrivilegeIfAvailable(driver As Selenium.EdgeDriver, id As String)
    If driver.FindElementById(id).Value > 0 Then
        driver.FindElementById(id).Click
    End If
End Sub


