Option Explicit
Dim BRW As Selenium.ChromeDriver

Sub ExtractListtags1()
Sheets(1).Select
Set BRW = New Selenium.ChromeDriver
BRW.Get ("https://en.wikipedia.org/wiki/Main_Page")

Dim mylist As Selenium.WebElement

Dim mylists As Selenium.WebElements
Set mylists = BRW.FindElementsByTag("li")

Dim r As Long
r = 1
For Each mylist In mylists

            Cells(r, 1) = mylist.Text
            r = r + 1

Next
End Sub
'---------------------------------------------------------------------------------------
Sub ExtractListtags2()
' using id attribute
Sheets(1).Select
Set BRW = New Selenium.ChromeDriver
BRW.Get ("https://en.wikipedia.org/wiki/Main_Page")
Dim mylist As Selenium.WebElement
Dim mylists As Selenium.WebElements
Set mylists = BRW.FindElementById("mp-otd").FindElementsByTag("li")
Dim r As Long
r = 1
For Each mylist In mylists
            Cells(r, 1) = mylist.Text
            r = r + 1
Next
'mp-otd
End Sub
'---------------------------------------------------------------------------------------
Option Explicit
Dim BRW As Selenium.ChromeDriver
Sub LinkText()
'project -how to land on current date page
'September 15
Set BRW = New Selenium.ChromeDriver
BRW.Get "https://en.wikipedia.org/wiki/Main_Page"
Dim dt As String
dt = Format(Date, "mmmm dd")
Dim ATAG As WebElement
Set ATAG = BRW.FindElementByLinkText(dt)
ATAG.Click
End Sub









