Option Explicit

Dim brw As Selenium.ChromeDriver

Sub workwithDropDown()
''using the findelementbyID method
Set brw = New Selenium.ChromeDriver
brw.Get ("https://www.amazon.com")
'searchDropdownBox
'brw.FindElementById("searchDropdownBox").Click
Dim ele As SelectElement
Set ele = brw.FindElementById("searchDropdownBox").AsSelect
ele.SelectByIndex (1)
ele.SelectByText "Baby"
ele.SelectByValue "search-alias=beauty-intl-ship"
End Sub
'-----------------------------------------------------------------------------------------

Sub exportoptionvalues()
'collection loop
Set brw = New Selenium.ChromeDriver
brw.Get ("https://www.amazon.com")
Dim ele As WebElement
Dim eles As WebElements
Set eles = brw.FindElementsByTag("option")
Dim r As Long
Sheets(1).Select
r = 1
For Each ele In eles
        Cells(r, 1) = ele.Text
        r = r + 1
Next ele
End Sub
'-----------------------------------------------------------------------------------------

Sub selectValue_xpath()
Set brw = New Selenium.ChromeDriver
brw.Get ("https://www.amazon.com")
Dim ele As SelectElement
Set ele = brw.FindElementByXPath("//*[@id='searchDropdownBox']").AsSelect
'Set ele = brw.FindElementByXPath("//*[@class='nav-search-dropdown searchSelect nav-progressive-attrubute nav-progressive-search-dropdown']").AsSelect']").AsSelect
ele.SelectByValue "search-alias=beauty-intl-ship"
End Sub
'-----------------------------------------------------------------------------------------
Option Explicit
Dim brw As Selenium.ChromeDriver

Sub workwithDropDown()
''we want to click on sub levels
'Dim ele As WebElement
'Set ele = brw.FindElementByXPath("//*[@id='nav-hamburger-menu']")
'ele.Click
Set brw = New Selenium.ChromeDriver
brw.Get ("https://www.amazon.com")

'brw.FindElementById("nav-hamburger-menu").Click

brw.FindElementByXPath("//*[@id='nav-hamburger-menu']").Click
'brw.FindElementByXPath("/html/body/div[3]/div[2]/div/ul[1]/li[8]/a]").Click
brw.FindElementByXPath("//*[@id='hmenu-content']/ul[1]/li[8]/a").Click
'//*[@id="hmenu-content"]/ul[6]/li[8]/a
brw.FindElementByXPath("//*[@id='hmenu-content']/ul[6]/li[8]/a").Click
End Sub
'-----------------------------------------------------------------------------------------

Sub workwithDropDown_litag()
''we want to click on sub levels
'Dim ele As WebElement
'Set ele = brw.FindElementByXPath("//*[@id='nav-hamburger-menu']")
'ele.Click
Set brw = New Selenium.ChromeDriver
brw.Get ("https://www.amazon.com")
brw.FindElementById("nav-hamburger-menu").Click
Dim li As WebElement
Dim list_tags As WebElements
Set list_tags = brw.FindElementById("hmenu-content").FindElementsByTag("li")
Set li = list_tags(8)
li.Click
' finding parenting element and going inside its child element and sub child element
brw.FindElementById("hmenu-content").FindElementsByTag("ul")(8).FindElementsByTag("li")(8).Click
End Sub
'-----------------------------------------------------------------------------------------





