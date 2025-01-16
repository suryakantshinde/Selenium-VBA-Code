
Private Assert as New Selenium.Assert
Sub AlertDemo
Set dlg = driver.SwitchToAlert(Raise:=False)

'assert make sure exists
Assert.False dlg is nothing, "No alert!"
Assert.Equals "Text in Msgbox", dlg.Text

'close alert
dlg.Accept 'to click ok button
dlg.Dismiss ' to click on cancel button

--------------------------------------------------------------
--------------------------------------------------------------

'test for dialog box
'debug.present IsDialogPresent(driver)
if is IsDialogPresent(driver) then
	Set dlg = driver.SwitchToAlert(Raise:=False)
	'close alert by clicking OK button
	dlg.Dismiss
end if 
----------------------------------------------------------------
'Return true is an alert is present, false otherwise
Private Finction IsDialogPresent(driver as WebDriver) as boolean
	on Error Resume Next
	T = driver.Title
	isDialogPresent = (26 = Err.Number)
End Function

--------------------------------------------------------------
--------------------------------------------------------------
dim dlg as selenium Alert
Set dlg = driver.SwitchToAlert(Raise:=False)
if not driver is nothing then
	else
	'close alert
	driver.Accept 'to click ok button
	end if
end sub
