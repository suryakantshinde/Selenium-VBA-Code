Private Keys As New Selenium.Keys
Private Assert As New Selenium.Assert
Sub Omnipay_Account_Number_Change_Start()
'On Error GoTo SuryaError
Dim driver As New Selenium.EdgeDriver
Dim count As Long
Dim fs As String
Dim button As WebElement
Dim i, lr As Long
Dim onum
onum = 2
Sheets("Instructions").Activate
        Set driver = CreateObject("Selenium.EdgeDriver")
        driver.Get "https://www2.omnipaygroup.com/ramtool"
'01 - Max Crome windows
        driver.Window.Maximize
        
'02 - Enter User Name
        driver.FindElementByXPath("//input[@id='69']").Clear
        driver.FindElementByXPath("//input[@id='69']").SendKeys Range("D4")
        
'03 - Enter Password
        driver.FindElementByXPath("//input[@id='76']").Clear
        driver.FindElementByXPath("//input[@id='76']").SendKeys Range("D5")
        
'04 - Click on Login Button
        driver.FindElementByXPath("//input[@value='Login']").Click
'-----------------------------------------------------------'-----------------------------------------------------------
     ' Allow to load the results page
     MsgBox "Please enter Pin from Email then Click ok", vbInformation
        driver.wait 5000
'-----------------------------------------------------------'-----------------------------------------------------------
 driver.WaitUntilNotPresent

  
 'Set elemSearch = driver.FindElement(By.Name, "btnK")
  Set elemSearch = driver.WaitUntilNotPresent(By.Name, "")
  
  'Click on Data Access Privileges
driver.FindElementByXPath("//*[@id='twofactor']").Click
'check Enhanced Data Access Privileges

'View Card Number
driver.FindElementByXPath("//*[@id='field-view-card-number']").Click

'View Bank Account
driver.FindElementByXPath("//*[@id='field-view-bank-account']").Click

'Update Bank Account
driver.FindElementByXPath("//*[@id='field-update-bank-account']").Click

'View Merchant PII Data
driver.FindElementByXPath("//*[@id='field-view-merchant-pii']").Click

'Update Merchant PII Data
driver.FindElementByXPath("//*[@id='field-update-merchant-pii']").Click

'Download Card Number
'driver.FindElementByXPath("//*[@id='field-download-card-number']").Click

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
           
'01 -  Click on
            driver.FindElementByLinkText("Merchant Administration").Click
            'driver.FindElementByXPath("//span[normalize-space()='Merchant Administration']").Click
            driver.wait 100
            
            driver.FindElementByLinkText("Merchant Maintenance").Click
            'driver.FindElementByXPath("//a[@class='yuimenuitemlabel yuimenuitemlabel-hassubmenu yuimenuitemlabel-selected yuimenuitemlabel-hassubmenu-selected']").Click
            
            driver.FindElementByLinkText("Maintain Merchant Details").Click
            'driver.FindElementByXPath("//a[@class='yuimenuitemlabel yuimenuitemlabel-selected']").Click
            
            driver.wait 2000
            
'02 -   Click on Select Merchant Number
            'driver.FindElementByXPath("//*[@id='merchbutton-button']").Click
            driver.FindElementByXPath("//button[@id='merchbutton-button']").Click

'03 -   Enter Merchant Number - //*[@id="id_40A"]
            'driver.FindElementByXPath("//*[@id='id_40A']").Clear
            driver.FindElementByXPath("//input[@id='id_40A']").Clear
            'driver.FindElementByXPath("//*[@id='id_40A']").SendKeys Range("A" & i)
            driver.FindElementByXPath("//input[@id='id_40A']").SendKeys Range("A" & i)

'04 -   Click on Change Button - //*[@id="changeMerchBtn"]
            'driver.FindElementByXPath("//*[@id='changeMerchBtn']").Click
            driver.FindElementByXPath("//button[@id='changeMerchBtn']").Click

driver.wait 2000

'05 -   Click on Accounts Tab - /html/body/div[6]/div[2]/form/div/div[2]/ul/li[3]/a/span
            driver.FindElementByXPath("//span[normalize-space()='Accounts']").Click
            driver.wait 9000
'===============================================================================================
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Local Payments Acct | Local Payments Acct | Local Payments Acct | Local Payments Acct | Local Payments Acct |
'===============================================================================================

Dim tbl_lpa As WebElement
Dim row_lpa As WebElement
Dim icon_lpa As WebElement
Dim refftype_lpa As WebElement
Dim reffvalue_lpa1 As WebElement
Dim reffvalue_lpa2 As WebElement
Dim txtvalue_lpa As String
row_count = 0
Set tbl_lpa = driver.FindElementById("accountListTable")
For Each row_lpa In tbl_lpa.FindElementsByTag("tr")
        If row_lpa.FindElementsByTag("td").count > 1 Then   '...................1
            Set icon_lpa = row_lpa.FindElementsByTag("td")(1)
            Set refftype_lpa = row_lpa.FindElementsByTag("td")(2)
            Set reffvalue_lpa1 = row_lpa.FindElementsByTag("td")(3)
            Set reffvalue_lpa2 = row_lpa.FindElementsByTag("td")(4)
             
                    If (Trim(reffvalue_lpa1.Text)) = "Local Payments Acct" Then '...................2
                           
                             If (Trim(reffvalue_lpa2.Text)) = "INR" Then '...................3
                                'Click on Icon
                                icon_lpa.Click
 '------------------------------------------------------------------------------------------------------------------------------------
                                '1 - Click on Edit Button
                                driver.FindElementByLinkText("Edit").Click
                                driver.wait 4000
                                
                                 '2 - Bank Account - //*[@id="ID21aaa"]
                                driver.FindElementByXPath("//*[@id='ID21aaa']").Clear
                                driver.FindElementByXPath("//*[@id='ID21aaa']").SendKeys Range("B" & i)
                                
                                '3 - MICR - //*[@id="ID21AAACA"]
                                driver.FindElementByXPath("//*[@id='ID21AAACA']").Clear
                                driver.FindElementByXPath("//*[@id='ID21AAACA']").SendKeys Range("C" & i)
                                
                                '4 - Bank Name - //*[@id="ID53AA"]
                                driver.FindElementByXPath("//*[@id='ID53AA']").Clear
                                driver.FindElementByXPath("//*[@id='ID53AA']").SendKeys Range("D" & i)
                                
                                '5 - Bank City - //*[@id="ID53AB"]
                                driver.FindElementByXPath("//*[@id='ID53AB']").Clear
                                driver.FindElementByXPath("//*[@id='ID53AB']").SendKeys Range("E" & i)
                                
                                '6 - Account Name - //*[@id="ID21aaaA"]
                                driver.FindElementByXPath("//*[@id='ID21aaaA']").Clear
                                driver.FindElementByXPath("//*[@id='ID21aaaA']").SendKeys Range("F" & i)
                                
                                '7 - Bank Sort Code - //*[@id="ID12AAAcl"]
                                driver.FindElementByXPath("//*[@id='ID12AAAcl']").Clear
                                driver.FindElementByXPath("//*[@id='ID12AAAcl']").SendKeys Range("G" & i)
                                
                                '8 - Copy Payable Entries - //*[@id="ID18abbcc1"]
                                driver.FindElementByXPath("//*[@id='ID18abbcc1']").Click
                                
                                driver.wait 1000
                                '9 - Click on Update Button - //button[text()='Update']
                                driver.FindElementByXPath("//button[text()='Update']").Click
                                driver.wait 1000
                                
                                '10 - Click on Close Button - //button[text()='Close']
                                driver.FindElementByXPath("//button[text()='Close']").Click
                                driver.wait 1000
 '------------------------------------------------------------------------------------------------------------------------------------
                            Else
                            End If '...................3
                            Exit For
                    End If '...................2
                    
        End If '...................1
              row_count = row_count + 1
Next
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
driver.wait 10000
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'05 -   Click on Accounts Tab - /html/body/div[6]/div[2]/form/div/div[2]/ul/li[3]/a/span
            driver.FindElementByXPath("//span[normalize-space()='Accounts']").Click
            driver.wait 9000
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Paymnt Acct Retail | Paymnt Acct Retail | Paymnt Acct Retail | Paymnt Acct Retail | Paymnt Acct Retail |
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim tbl_par As WebElement
Dim row_par As WebElement
Dim icon_par As WebElement
Dim refftype_par As WebElement
Dim reffvalue_par1 As WebElement
Dim reffvalue_par2 As WebElement
Dim txtvalue_par As String
row_count = 0
Set tbl_par = driver.FindElementById("accountListTable")
For Each row_par In tbl_par.FindElementsByTag("tr")
        If row_par.FindElementsByTag("td").count > 1 Then   '...................1
            Set icon_par = row_par.FindElementsByTag("td")(1)
            Set refftype_par = row_par.FindElementsByTag("td")(2)
            Set reffvalue_par1 = row_par.FindElementsByTag("td")(3)
            Set reffvalue_par2 = row_par.FindElementsByTag("td")(4)
             
                    If (Trim(reffvalue_par1.Text)) = "Paymnt Acct Retail" Then '...................2
                           
                             If (Trim(reffvalue_par2.Text)) = "INR" Then '...................3
                                'Click on Icon
                                icon_par.Click
 '------------------------------------------------------------------------------------------------------------------------------------
                                '1 - Click on Edit Button
                                driver.FindElementByLinkText("Edit").Click
                                driver.wait 4000
                                
                                 '2 - Bank Account - //*[@id="ID21aaa"]
                                driver.FindElementByXPath("//*[@id='ID21aaa']").Clear
                                driver.FindElementByXPath("//*[@id='ID21aaa']").SendKeys Range("B" & i)
                                
                                '3 - MICR - //*[@id="ID21AAACA"]
                                driver.FindElementByXPath("//*[@id='ID21AAACA']").Clear
                                driver.FindElementByXPath("//*[@id='ID21AAACA']").SendKeys Range("C" & i)
                                
                                '4 - Bank Name - //*[@id="ID53AA"]
                                driver.FindElementByXPath("//*[@id='ID53AA']").Clear
                                driver.FindElementByXPath("//*[@id='ID53AA']").SendKeys Range("D" & i)
                                
                                '5 - Bank City - //*[@id="ID53AB"]
                                driver.FindElementByXPath("//*[@id='ID53AB']").Clear
                                driver.FindElementByXPath("//*[@id='ID53AB']").SendKeys Range("E" & i)
                                
                                '6 - Account Name - //*[@id="ID21aaaA"]
                                driver.FindElementByXPath("//*[@id='ID21aaaA']").Clear
                                driver.FindElementByXPath("//*[@id='ID21aaaA']").SendKeys Range("F" & i)
                                
                                '7 - Bank Sort Code - //*[@id="ID12AAAcl"]
                                driver.FindElementByXPath("//*[@id='ID12AAAcl']").Clear
                                driver.FindElementByXPath("//*[@id='ID12AAAcl']").SendKeys Range("G" & i)
                                
                                '8 - Copy Payable Entries - //*[@id="ID18abbcc1"]
                                driver.FindElementByXPath("//*[@id='ID18abbcc1']").Click
                                
                                driver.wait 1000
                                '9 - Click on Update Button - //button[text()='Update']
                                driver.FindElementByXPath("//button[text()='Update']").Click
                                driver.wait 1000
                                
                                '10 - Click on Close Button - //button[text()='Close']
                                driver.FindElementByXPath("//button[text()='Close']").Click
                                driver.wait 1000
 '------------------------------------------------------------------------------------------------------------------------------------
                            Else
                            End If '...................3
                            Exit For
                    End If '...................2
                    
        End If '...................1
              row_count = row_count + 1
Next
'===============================================================================================
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
driver.wait 1000
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
driver.FindElementByLinkText("References").Click
driver.wait 1000
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Search for NEFT | Search for NEFT | Search for NEFT | Search for NEFT | Search for NEFT | Search for NEFT |
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'referenceListTable
Dim tbl_neft As WebElement
Dim row_neft As WebElement
Dim icon_neft As WebElement
Dim refftype_neft As WebElement
Dim reffvalue_neft As WebElement
Dim txtvalue_neft As String
row_count = 0
Set tbl_neft = driver.FindElementById("referenceListTable")
For Each row_neft In tbl_neft.FindElementsByTag("tr")
        If row_neft.FindElementsByTag("td").count > 1 Then   '...................1
            Set icon_neft = row_neft.FindElementsByTag("td")(1)
            Set refftype_neft = row_neft.FindElementsByTag("td")(2)
            Set reffvalue_neft = row_neft.FindElementsByTag("td")(2)
            
                    If (Trim(reffvalue_neft.Text)) = "NEFT" Then '...................2
                                'Click on Icon
                                icon_neft.Click
                                
                                '1 - Click on Delete Button
                                driver.FindElementByLinkText("Delete").Click
                                driver.wait 500
                                
                                '10 - Click on Delete Button - //button[text()='Delete']
                                driver.FindElementByXPath("//button[text()='Delete']").Click
                                driver.wait 1000
 '------------------------------------------------------------------------------------------------------------------------------------
 '------------------------------------------------------------------------------------------------------------------------------------
                            Exit For
                    End If '...................2
                    
        End If '...................1
              row_count = row_count + 1
Next
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
driver.wait 1000

'Click on Add New Reference Button
driver.FindElementById("addReference").Click
driver.wait 3000

'click on check box NEFT
If Sheets("RawData").Range("H" & i).Text = "TTUMP" Then  ''' ICIC Bank
    ''Other Bank will select NEFT - Yes
       driver.FindElementByName("SEL_364").Click
driver.wait 100
'Select Yes from dropdown
driver.FindElementByName("48FRT_364").AsSelect.SelectByText ("Yes")
driver.wait 100
'10 - Click on Add Button - //button[text()='Add']
driver.FindElementByXPath("//button[text()='Add']").Click
driver.wait 5000
ElseIf Sheets("RawData").Range("H" & i).Text = "NEFT" Then  ''' Other Banks
    ''Other Bank will select NEFT - Yes
       driver.FindElementByName("SEL_366").Click
driver.wait 100
'Select Yes from dropdown
driver.FindElementByName("48FRT_366").AsSelect.SelectByText ("Yes")
driver.wait 100

'10 - Click on Add Button - //button[text()='Add']
driver.FindElementByXPath("//button[text()='Add']").Click
driver.wait 5000
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Details Tab| Details Tab | Details Tab | Details Tab | Details Tab | Details Tab | Details Tab | Details Tab | Details Tab |
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Click on Detail Tab
driver.FindElementByLinkText("Detail").Click
driver.wait 1000

'select Client Status as "Active" - 78ccc
driver.FindElementByName("78ccc").AsSelect.SelectByText ("Active")
driver.wait 1000

'10 - Click on Save Button - //button[text()='Save']
driver.FindElementByXPath("//button[text()='Save']").Click
driver.wait 1000
'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------

Range("I" & i).Value = "Record updated"
Next i
'-----------------------------------------------------------'-----------------------------------------------------------
ActiveWorkbook.Save
MsgBox "Completed Successfully", vbInformation
driver.Quit
'------------------------------------------------------------------------------------------------------
Exit Sub
SuryaError:
MsgBox "Error in code - Contact Surya", vbCritical
End Sub



