sub check
Dim CheckBox As WebElement
Set CheckBox = driver.FindElementByXPath("(//span[text()='caseTcn']/following::input)[1]")
If CheckBox.IsSelected Then
MsgBox "Its Already Tick"
Else
MsgBox "No Tick"
End If
end sub
