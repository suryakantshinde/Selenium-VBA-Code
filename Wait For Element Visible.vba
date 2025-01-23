Sub WaitForElementVisible()
    Dim driver As New WebDriver
    Dim element As Object
    Dim timeout As Double
    Dim startTime As Double
    
    ' Start the WebDriver (e.g., ChromeDriver)
    driver.Start "chrome", "https://yourwebsite.com"
    driver.Get "/"

    ' Set the timeout (in seconds)
    timeout = 10 ' Adjust this value as needed
    startTime = Timer

    ' Loop until the element is found or timeout is reached
    Do
        On Error Resume Next
        Set element = driver.FindElementByCss(".your-element-class")
        On Error GoTo 0
        If Not element Is Nothing Then
            If element.IsDisplayed Then Exit Do
        End If
        DoEvents ' Allow other events to process
    Loop While Timer - startTime < timeout
    
    ' Check if element is found and visible
    If Not element Is Nothing And element.IsDisplayed Then
        MsgBox "Element is visible!"
    Else
        MsgBox "Element not found or not visible within timeout."
    End If
    
    ' Clean up
    driver.Quit
End Sub
