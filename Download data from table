Private Keys As New Selenium.Keys
Private Assert As New Selenium.Assert
Sub Download_Auth_History_Automation()
On Error GoTo Surya
Dim driver As New Selenium.ChromeDriver
Dim count As Long
Dim fs As String
Dim button As WebElement
Sheets("Output").Activate ''Clear Previous Data
Range("A2:G104876").ClearContents
Sheets("Input").Activate
Set driver = CreateObject("Selenium.ChromeDriver")
driver.Get "https://www2.omnipaygroup.com/ramtool"
driver.Window.Maximize
'1 - Enter User Name
driver.FindElementById("69").SendKeys Range("R2")
'2 - Enter Password
driver.FindElementById("76").SendKeys Range("R3")
'3 - Click on Login Button
driver.FindElementByXPath("//input[@value='Login']").Click
'-----------------------------------------------------------'
' Allow to load the results page
driver.wait 2000
'-----------------------------------------------------------'
'4 - Click on Select Institution
driver.FindElementById("selectinst").Click
'5 - Select 00000066 - Robinsons Bank
driver.FindElementByXPath("/html/body/div[1]/div[3]/div[1]/div[1]/ul/li[2]/a").Click
'6 - Click on two factor Authentication
driver.FindElementById("twofactor").Click
driver.wait 50
driver.FindElementByXPath("//button[text()='Generate PIN']").Click
driver.wait 60000
'7 - Click on Merchant Activity
driver.FindElementByXPath("//a[@href = '#Merchant Activity']").Click
'8 - Click on Authorisation History
driver.FindElementByLinkText("Authorisation History").Click
driver.wait 7000
'-----------------------------------------------------------'
onum = 2
Sheets("Input").Activate
Dim i, lr As Long
lr = ThisWorkbook.ActiveSheet.Cells(Rows.count, 1).End(xlUp).row
For i = 2 To lr
Sheets("Input").Activate
''Steps:
'=============================================================
'7 - Click on Merchant Activity
driver.FindElementByXPath("//a[@href = '#Merchant Activity']").Click
'8 - Click on Authorisation History
driver.FindElementByLinkText("Authorisation History").Click
'9 - Enter Start Date
' driver.SwitchToPreviousWindow
'driver.FindElementById("ex2h").Click
driver.FindElementByName("55B").Clear
driver.FindElementByName("55B").SendKeys Range("A" & i).Text
driver.wait 200
'10 - Enter End Date
driver.FindElementByName("28B").Clear
driver.FindElementByName("28B").SendKeys Range("B" & i).Text
'11 - Enter Merchant No
driver.FindElementByName("40").Clear
driver.FindElementByName("40").SendKeys Range("C" & i).Text
driver.wait 100
'12 - Enter Auth Code
driver.FindElementByName("12k").Clear
driver.FindElementByName("12k").SendKeys Range("D" & i).Text
driver.wait 100
'13 - EnterClick on Search Button
driver.FindElementByXPath("//button[@id='search']").Click
driver.Timeouts.ImplicitWait = 10


''''=============================================================


Dim tbl As WebElement
Dim row As WebElement
Dim icon As WebElement
Dim refftype As WebElement
Dim reffvalue As WebElement
Dim txtvalue As String
Set tbl = driver.FindElementByClass("zebratable")
For Each row In tbl.FindElementsByTag("tr")
If row.FindElementsByTag("td").count > 0 Then
Set merch_num = row.FindElementsByTag("td")(1)
Set card_num = row.FindElementsByTag("td")(2)
Set Expiry = row.FindElementsByTag("td")(4)
Set amount = row.FindElementsByTag("td")(5)
Set auth_code = row.FindElementsByTag("td")(6)
Set auth_date = row.FindElementsByTag("td")(7)
Set file_date = row.FindElementsByTag("td")(8)
Sheets("Output").Range("A" & onum) = merch_num.Text
Sheets("Output").Range("B" & onum) = card_num.Text
Sheets("Output").Range("C" & onum) = Expiry.Text
Sheets("Output").Range("D" & onum) = amount.Text
Sheets("Output").Range("E" & onum) = auth_code.Text
Sheets("Output").Range("F" & onum) = auth_date.Text
Sheets("Output").Range("G" & onum) = file_date.Text
onum = onum + 1
End If
driver.wait 100
Next

'------------------------------------------------------------------------------------------------------
Next i
'------------------------------------------------------------------------------------------------------
Sheets("Output").Activate
Columns("B:B").Select
Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=False, Space:=False, OTHER:=False, FieldInfo _
:=Array(1, 2), TrailingMinusNumbers:=True
Range("A1").Select
'------------------------------------------------------------------------------------------------------
driver.Quit
MsgBox "Completed Successfully", vbInformation
'------------------------------------------------------------------------------------------------------
Exit Sub
Surya:
MsgBox "Contact Surya Error in code", vbCritical
End Sub


