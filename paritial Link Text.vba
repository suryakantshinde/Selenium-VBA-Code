Option Explicit
Dim BRW As Selenium.ChromeDriver

Sub LinkText()
Sheets(1).Select
Set BRW = New Selenium.ChromeDriver
BRW.Get ("https://en.wikipedia.org/wiki/Main_Page")
Dim ATAG As WebElement
Set ATAG = BRW.FindElementByLinkText("Raymond Pace Alexander")
ATAG.Click
'Dim atags As WebElements
'Set atags = brw.FindElementsByTag("a")
'Raymond Pace Alexander
'For Each atag In atags
       ' if atag.Attribute("href") = "" then
'Next
End Sub
'-------------------------------------------------------------------------------------------------
Sub paritial_LinkText()
  Sheets(1).Select
Set BRW = New Selenium.ChromeDriver
BRW.Get ("https://en.wikipedia.org/wiki/Main_Page")
'wikimedia
Dim a As WebElement
Dim atags As WebElements
Set atags = BRW.FindElementsByPartialLinkText("Wikimedia")
Set a = atags(2)
a.Click
End Sub
'-------------------------------------------------------------------------------------------------

