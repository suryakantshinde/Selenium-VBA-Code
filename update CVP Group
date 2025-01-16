Private Keys As New Selenium.Keys
Private Assert As New Selenium.Assert
Sub Update_CPV_Group_Automation()
On Error GoTo Surya

Dim driver As New Selenium.ChromeDriver
Dim count As Long
Dim fs As String
Dim FSO As Scripting.FileSystemObject
Dim logFile As Scripting.TextStream

Set FSO = New Scripting.FileSystemObject
Set logFile = FSO.CreateTextFile(ThisWorkbook.Path & "\LogFile.txt")
logFile.WriteLine (Now & "----Macro Started----")
    
Sheets("Login").Select
        
        Set driver = CreateObject("Selenium.ChromeDriver")
        logFile.WriteLine Now() & "----Chrome Drivers Started"
        driver.Get "https://www2.omnipaygroup.com/ramtool"
        logFile.WriteLine Now() & "----Log into Onnipay Website Successfully----"
        'driver.Window.Maximize
        logFile.WriteLine Now() & "----Window Maximize"
        
'Enter User Name
        driver.FindElementById("69").SendKeys Range("B2")
        logFile.WriteLine Now() & "----Entered User Name as----" & Range("B2")
'Enter Password
        driver.FindElementById("76").SendKeys Range("B3")
        logFile.WriteLine Now() & "----Entered Password as----" & Range("B3")
'Click on Login Button
        driver.FindElementByXPath("//input[@value='Login']").Click
        ' allow to load the results page
        logFile.WriteLine Now() & "----Clicked on Login Button Successfully----"
        driver.wait 300
'================================================================
'clickonMerchant Administration  link    Merchant Administration
    driver.FindElementByLinkText("Merchant Administration").Click
    logFile.WriteLine Now() & "----Clicked on Merchant Administration Successfully----"
    driver.wait 200
    
'clickonMerchant Application Setup   link    Merchant Maintenance
    driver.FindElementByLinkText("Merchant Maintenance").Click
    logFile.WriteLine Now() & "----Clicked on Merchant Maintenance Setup Successfully----"
    driver.wait 300
'selectMerchant Application List link    Maintain Mantenanace Detail
    'driver.FindElementByXPath("/html/body/div[1]/div[2]/div/div/div/div/div/ul/li[5]/div/div[1]/ul[2]/li/div/div[1]/ul[1]/li/a").Click
    driver.FindElementByCss("a[href*='MERCH_MAINTAIN_DETAILS']").Click
    logFile.WriteLine Now() & "----Clicked on Maintain Mantenanace Detail Successfully----"
'================================================================
driver.wait 100
'-----------------------------------------------------------'-----------------------------------------------------------
Sheets("Update_CPV").Activate
Dim i, lr As Long
lr = ThisWorkbook.ActiveSheet.Cells(Rows.count, 1).End(xlUp).row
For i = 2 To lr

''Steps:

''1]  Click on Enter Merchant Number Button
driver.FindElementById("merchbutton-button").Click
logFile.WriteLine Now() & "----Click on Enter Merchant Number Button----"
driver.wait 300

''2] Enter Merchant Number in Text Box
driver.FindElementById("id_40A").Clear
driver.FindElementById("id_40A").SendKeys Range("A" & i)
logFile.WriteLine Now() & "----Entered Merchant No----" & Range("A" & i)
driver.wait 300

''3] Click on Change Button
driver.FindElementById("changeMerchBtn").Click
'logFile.WriteLine Now() & "----Click on Change Button----"
driver.wait 500

''4] Click on Property Tab
driver.FindElementByLinkText("Properties").Click
'logFile.WriteLine Now() & "----Click on Property Tab----"
driver.wait 300

''5] Click on Icon and select update from drop down
'---------------------------------------------------------------------------------------
Dim tbl As WebElement
Dim row As WebElement
Dim icon As WebElement
Dim refftype As WebElement
Dim reffvalue As WebElement
Dim txtvalue As String
Set tbl = driver.FindElementByXPath("//form/div/div[6]/div[1]/div[3]/table/tbody[2]")
txtvalue = Range("B" & i).Value
For Each row In tbl.FindElementsByXPath("./tr")
        Set icon = row.FindElementsByTag("td")(1)
        Set refftype = row.FindElementsByTag("td")(2)
        Set reffvalue = row.FindElementsByTag("td")(3)
        If (refftype.Text = "CPV GROUP ID") Then
            icon.Click
            driver.FindElementByLinkText("Update").Click              ' click on update dropdown
    driver.wait 300
  'Logic
'---------------------------------------------------------------------------------------------------
                Dim strTemp As WebElement
                Dim strFind As String
                Dim CS As WebElement
                Dim PropValu As String
              
                Set strTemp = driver.FindElementByXPath("/html/body/div[5]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[3]/div/fieldset/p[2]/span[2]/input")
                 PropValu = strTemp.Value
                 
                If InStr(1, strTemp.Value, txtvalue) > 0 Then
                Range("D" & i).Value = strTemp.Value
                    strTemp.Clear
                    
                    strTemp.SendKeys Replace(PropValu, txtvalue, "")
'---------------------------------------------------------------------------------------------------
                driver.FindElementByXPath("//button[@class ='update']").Click
                driver.wait 300
                 Range("C" & i).Value = "Record Updated"       ' click on update button
                logFile.WriteLine Now() & "----Record Updated----"
                Range("E" & i).Value = GetPropValue(driver)
                driver.wait 300
            Else
                Range("C" & i).Value = "Record Not Updated"
                logFile.WriteLine Now() & "----Record Not Updated----"
                driver.FindElementByXPath("//button[@class ='cancel']").Click
                driver.wait 300
            End If
ActiveCell.Offset(1, 0).Select
Exit For
        End If
        Next
 driver.wait 300
 
'------------------------------------------------------------------------------------
Next i
driver.Quit
MsgBox "Completed Sucessfully", vbInformation
Exit Sub
Surya: Range("D" & i).Value = "Error"
logFile.WriteLine Now() & "----Error----"
driver.Quit


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

