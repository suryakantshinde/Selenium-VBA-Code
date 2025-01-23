'driver.FindElementById("Button").waitText "Hello",20000
'driver.FindElementById("mytext").waitEnabled False, 20000

'================================================================================================================
'Perfect code
'Below code will wait untill "Visa Resolve Online - India" is available then it will click on it
'================================================================================================================
On Error Resume Next
Do While IsError(driver.FindElementByLinkText("Visa Resolve Online - India")) = True
Loop
driver.FindElementByLinkText("Visa Resolve Online - India", 10000).Click
'================================================================================================================
'================================================================================================================


Sub wait()
'----------------------------------------------------------------------------------------------------------------------
'Wait for Download Button to be visible
'----------------------------------------------------------------------------------------------------------------------
   
   Dim element As WebElement
    Dim timeout As Integer
    timeout = 60 ' Set your timeout in seconds
    Dim startTime As Double
    startTime = Timer

    Do
        On Error Resume Next
        Set element = driver.FindElementByXPath("//*[@id='download']") ' Change to your desired element identifier

        If element.IsDisplayed = True Then
            Exit Do
        End If

        If Timer - startTime > timeout Then
         '   MsgBox "Timeout while waiting for element"
           ' Exit Do
        End If
        Application.wait (Now + TimeValue("0:00:1")) ' Wait for 1 second before checking again
    Loop
' Continue with actions on the found element or other operations

End Sub

'----------------------------------------------------------------------------------------------------------------------
Sub wait2()

Do Until .FindElementById("ContentPlaceHolder1_Label2").Text = "?? ??? ??????? ?????"
    Application.wait Now() + TimeValue("00:00:01")
Loop
'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
'Multi-line msgbox
'Multi-line msgbox
'Multi-line msgbox
'Multi-line msgbox
Private Sub about_button_Click()
    MsgBox "Name: gemUI" & vbCrLf & "Version: 1.0" & vbCrLf & "Build: 0001" & _
        vbCrLf & "(C) 2018 Josh Face", , "About gemUI"
End Sub
'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------'----------------------------------------------------------------------------------------------------------------------
  MsgBox "Please do PingID multi-factor Authentication" & vbCrLf & "UPDATE PRIVILIGES" & vbCrLf & "Then Click OK button to continue running automation", vbInformation, "PingID-Authentication"
  driver.wait 5000

'----------------------------------------------------------------------------------------------------------------------'----------------------------------------------------------------------------------------------------------------------
  MsgBox "Please do..." & vbCrLf & vbCrLf & "1 - PingID multi-factor Authentication" & vbCrLf & vbCrLf & "2 - Update Priviliges if needed" & vbCrLf & vbCrLf & "Then Click OK button to continue running automation", vbInformation, "PingID-Authentication"
 driver.wait 5000

'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
'Show/Get Current Case No in Msgbox

Sub GetValuesFromDynamicRange()
    Dim cell As Range
    Dim cellValue As Variant
    Dim lastRow As Long
    
    ' Find the last row with data in column A
    lastRow = Sheets("Sheet1").Cells(Rows.count, 1).End(xlUp).row
    
    ' Loop through each cell in the range A1 to the last row in column A
    For Each cell In Sheets("Sheet1").Range("A1:A" & lastRow)
        cellValue = cell.Value
        'MsgBox "The value in cell " & cell.Address & "---- Case No is ---- " & cellValue
        MsgBox "Current Case ---- " & cellValue, vbInformation, "Case Number"
    Next cell
End Sub
