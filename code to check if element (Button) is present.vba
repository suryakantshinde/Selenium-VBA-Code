===================================================================
Perfect code to check if element (Button) is present
===================================================================
Dim driver As New WebDriver
Dim element As WebElement
Dim isVisible As Boolean

' Initialize the driver and navigate to the website
driver.Start "chrome", "http://example.com"

' Try to find the button element
On Error Resume Next
Set element = driver.FindElementByCssSelector(".button-class")
On Error GoTo 0

' Check if the element is found
If Not element Is Nothing Then
    ' Check if the element is visible
    isVisible = element.Displayed
    If isVisible Then
        MsgBox "Button is present and visible!"
    Else
        MsgBox "Button is present but not visible!"
    End If
Else
    MsgBox "Button is not present!"
End If

' Close the browser
driver.Quit

===================================================================
===================================================================

Usage with FindElement:

Dim driver As New FirefoxDriver
driver.Get "https://en.wikipedia.org/wiki/Main_Page"
Set ele = driver.FindElementById("searchInput", Raise:=False, Timeout:=0)
If Not ele Is Nothing Then
  ele.SendKeys "xyz"
End If
===================================================================
Usage with FindElements:

Dim driver As New FirefoxDriver
driver.Get "https://en.wikipedia.org/wiki/Main_Page"
Set elts = driver.FindElementsById("searchInput")
If elts.Count > 0 Then
  elts(0).SendKeys "xyz"
End If
===================================================================
Usage with IsElementPresent:

Dim By As New By
Dim driver As New FirefoxDriver
driver.Get "https://en.wikipedia.org/wiki/Main_Page"
If driver.IsElementPresent(By.Id("searchInput")) Then
  Debug.Print "Element is present"
End If
===================================================================
===================================================================

' Wait for 3 seconds

Dim By As Selenium.By
Set By = New Selenium.By
If driver.IsElementPresent(By.ID("saveBtn"), 250000) Then
 driver.FindElementById("saveBtn").Click
Else
 driver.Window.Activate.Close
 driver.SwitchToPreviousWindow
End If
===================================================================
===================================================================
On Error Resume Next
Do While .FindElementById("theBttnbobjid_1545642213647_dialog_submitBtn") Is Nothing
    DoEvents
Loop
On Error Goto 0
===================================================================

===================================================================

===================================================================
/html/body/ngb-modal-window[2]/div/div/app-charge-slip-modal/div[1]/fa


Private Sub UserForm_Terminate()
Dim form As New UserForm_IDFC
ActiveWorkbook.Save
    ActiveWorkbook.Saved = True
    'ThisWorkbook.Save
    If MsgBox("Do you want to close Excel?", vbOKCancel + vbQuestion, "Exit Excel?") <> vbOK Then
        Load form
        ActiveWorkbook.Save
      form.Show
    Else
    ActiveWorkbook.Save
   ActiveWorkbook.Close SaveChanges:=True
   ActiveWorkbook.Save
      ActiveWorkbook.Close False
    End If
End Sub



===================================================================
VBA Internet Explorer wait for web page to load
===================================================================
Dim t As Date, ele As Object
Const MAX_WAIT_SEC As Long = 10 '<==Adjust wait time

While ie.Busy Or ie.readyState < 4: DoEvents: Wend
t = timer
Do 
    DoEvents
    On Error Resume Next
    Set ele = IE.document.getElementByID("firstname")
    If Timer - t > MAX_WAIT_SEC Then Exit Do
    On Error GoTo 0
Loop While ele Is Nothing

If Not ele Is Nothing Then
    'do something 
End If

===================================================================
        ' Click on Respond Button
Set element_respond_button = driver.FindElementByXPath("xpath")
        
        On Error GoTo 0

' Check if the element is found
If Not element_respond_button Is Nothing Then
    ' Check if the element is visible
    isVisible = element_respond_button.IsDisplayed
    If isVisible Then
        'MsgBox "Button is present and visible!"
                driver.FindElementByXPath("xpath").Click
                driver.Wait 5000
    Else
                driver.Window.Close
                driver.SwitchToPreviousWindow
                Range("D" & i).Value = "Case allready Processed"
                GoTo continue
    End If
Else
End If
===================================================================
/html[1]/body[1]/div[12]/ul[1]
csspath - html>body>div:nth-of-type(12)>ul>li:nth-of-type(2)>label>input
