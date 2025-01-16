Sub VISA_Allocation_Pre_Arbitration() ' visaonline.com
'On Error GoTo Surya

Dim driver As New Selenium.ChromeDriver
Dim count As Long
Dim fs As String
Dim FSO As Scripting.FileSystemObject
Dim logFile As Scripting.TextStream

Set FSO = New Scripting.FileSystemObject
Set logFile = FSO.CreateTextFile(ThisWorkbook.Path & "\LogFile.txt")
logFile.WriteLine (Now & "----Macro Started----")
    
Sheets("Instructions").Select
        
        Set driver = CreateObject("Selenium.ChromeDriver")
        logFile.WriteLine Now() & "----Chrome Drivers Started"
        'Visa resolve Online India  (New)– Domestic  url
        driver.Get "https://www.visaonline.com/login/LoginMain.aspx/"
                    'https://www.visaonline.com/login/loginmain.aspx
        driver.wait 1000
        logFile.WriteLine Now() & "----Log into visaonline Website Successfully----"
        driver.Window.Maximize
        logFile.WriteLine Now() & "----Window Maximize"
        
'Enter User Name
        driver.FindElementById("txtUsername").SendKeys Range("C3")
        logFile.WriteLine Now() & "----Entered User Name as----" & Range("C3")
         driver.wait 1000
'Enter Password
        driver.FindElementByXPath("/html/body/form/div[3]/div/section[1]/section/section[2]/div/section/ul/li[2]/input").SendKeys Range("C4")
        logFile.WriteLine Now() & "----Entered Password as----" & Range("C4")
        driver.wait 1000
'Click on Login Button
         'driver.FindElementByXPath("//input[@value='Log In']").Click
         driver.FindElementById("btnLogin").Click
        ' allow to load the results page
        logFile.WriteLine Now() & "----Clicked on Login Button Successfully----"
        driver.wait 12000

    'driver.Wait 50
    
driver.FindElementByXPath("/html/body/vol-root/vol-timeout-monitor/div/vol-notifications/div[1]/vol-base-layout/div/main/vol-subheader-trim/vol-home-page/section/div/div/div[1]/div/div[1]/div[2]/div/ul/li[2]/a").Click
driver.wait 9000

'click on "Visa Resolve Online"  id - ROL
'    driver.FindElementById("ROL").Click
    logFile.WriteLine Now() & "----Clicked on Visa Resolve Online Successfully----"
    driver.Window.SwitchToNextWindow
    driver.Window.Maximize
    
    
    
'-----------------------------------------------------------'-----------------------------------------------------------
Sheets("RawData").Activate
Dim i, lr As Long
lr = ThisWorkbook.ActiveSheet.Cells(Rows.count, 1).End(xlUp).row
For i = 2 To lr
'================================================================
'Click on Quick Search id - searchField
driver.wait 3000
driver.FindElementById("searchField").Clear
driver.FindElementById("searchField").SendKeys Range("A" & i)
driver.wait 100
'===================================================================================================================================
  
'Click on Go Button - name - gobutton
driver.FindElementByName("gobutton").Click
driver.wait 5000
'===================================================================================================================================
'click on "Disputes"   link    Disputes
driver.SwitchToFrame ("bodyframe")
driver.FindElementByXPath("//a[contains(text(),'Dispute Details')]").Click
logFile.WriteLine Now() & "----Clicked on Disputes Successfully----"
driver.wait 5000
'driver.Window.SwitchToPreviousWindow.Close
'driver.Wait 5000
'===================================================================================================================================
'Click on Respond Button
driver.Window.SwitchToNextWindow
driver.Window.Maximize
driver.wait 3000
driver.FindElementByXPath("//*[@id='container']/div/div[4]/div[2]/div/div[2]/div/div[3]/div/div[2]/div[2]/button[9]").Click
driver.wait 5000
'===================================================================================================================================
'Why are you initiating Pre-Arbitration? select - Compelling Evidence
driver.FindElementByXPath("/html/body/div/div/div[4]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div/div[2]/div[3]/div/div/div/div/div/div[1]/div/select/option[3]").Click
driver.wait 100
'===================================================================================================================================
'Compelling Evidence Type: - select Documentation to prove the cardholder is in possession of and /or using the merchandise
driver.FindElementByXPath("/html/body/div/div/div[4]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div/div[2]/div[3]/div/div/div/div/div/div[31]/div/div[1]/div/select/option[2]").Click
driver.wait 100
'===================================================================================================================================
'Click on Attach Documents
driver.FindElementByXPath("//span[contains(text(),'Attach Documents')]").Click
'=====================================================================================================================
'Switch to Attach Document Window
driver.wait 1000
driver.SwitchToNextWindow
driver.wait 3000

'Attach file
Dim CF As WebElement
file_path = "C:\Visa_Resolve_Online_India\Attachments\"
'file_path = "C:\new\IMS_CB_Visa_Resolve_Online_India_New\New folder\Attachments\"
CaseNo = Range("B" & i)

sk_str = ""
For cc = 0 To 10 Step 1
    If cc = 0 Then
       suffix = ""
    Else
       suffix = "_" & cc
    End If
If Dir(file_path & CaseNo & suffix & ".pdf") <> "" Then
    sk_str = file_path & CaseNo & suffix & ".pdf"
    Set CF = driver.FindElementByXPath("//input[@type='file']")
    'Select File Type: --> Adobe Acrobat Document (*.pdf) - name - 44e
    CF.SendKeys (sk_str)
'=====================================================================================================================
    'Select document Type = Other
    driver.FindElementById("documentType").AsSelect.SelectByText ("Other")
    documentTypeValue = OTHER
    driver.wait 500
'=====================================================================================================================
   '24 - Click on Attach Button
    driver.FindElementByXPath("//*[@id='attachRemoteButton']").Click
End If
driver.wait 4000
Next
'=====================================================================================================================
'Click on Close Button
driver.FindElementByXPath("//*[@id='closeButton']").Click
driver.wait 1000
'=====================================================================================================================
'Enter Comments
driver.SwitchToPreviousWindow
driver.FindElementByXPath("/html/body/div/div/div[4]/div[2]/div[2]/div[2]/div[1]/div[2]/div[2]/div[1]/div/div/div[2]/div[1]/div/textarea").SendKeys Range("E" & i)
driver.wait 1000
'=====================================================================================================================
'click on Submit Button
driver.FindElementByXPath("//*[@id='container']/div/div[4]/div[2]/div/div[2]/div/div[3]/div/div[3]/div[3]/button[4]").Click
driver.wait 9000
Dim w As Selenium.Window
For Each w In driver.Windows
    Debug.Print w.Title
 w.Activate
 Next w
 
'Click on Home
driver.FindElementByXPath("/html/body/div[1]/div[1]/div/form[1]/table/tbody/tr[2]/td[4]/table/tbody/tr/td[3]/a").Click
'=====================================================================================================================
Range("F" & i).Value = "Done"

Next i
driver.Quit
MsgBox "Completed Sucessfully", vbInformation
Surya:
Exit Sub
logFile.WriteLine Now() & "----Error----"
driver.Quit
End Sub
