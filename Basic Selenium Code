Private Keys As New Selenium.Keys
Private Assert As New Selenium.Assert
Sub Update_Property_Tab_Start()
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


'Add your Code here



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
Private Function GetPropValue(driver As Selenium.ChromeDriver) As String
    GetPropValue = ""
    Set tbl = driver.FindElementByXPath("//form/div/div[6]/div[1]/div[3]/table/tbody[2]")
    For Each row In tbl.FindElementsByXPath("./tr")
            Set icon = row.FindElementsByTag("td")(1)
            Set refftype = row.FindElementsByTag("td")(2)
            Set reffvalue = row.FindElementsByTag("td")(3)
            If (refftype.Text = "CPV GROUP ID") Then
                GetPropValue = reffvalue.Text
                Exit For
            End If
    Next
End Function


