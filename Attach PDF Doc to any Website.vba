Dim CF As WebElement
file_path = "C:\Selenium_VBA_Projects\Rupay_Dispute_Management\Attachments\"
CaseNo = Range("F" & i)

sk_str = ""
For CC = 0 To 10 Step 1
    If CC = 0 Then
       suffix = ""
    Else
       suffix = "_" & CC
    End If
If Dir(file_path & CaseNo & suffix & ".pdf") <> "" Then
    sk_str = file_path & CaseNo & suffix & ".pdf"
    Set CF = driver.FindElementByXPath("//input[@type='file']")
    'Select File Type: --> Adobe Acrobat Document (*.pdf) - name - 44e
    CF.SendKeys (sk_str)
    '24 - Click on Upload Button
    driver.FindElementByXPath("//button[@class='btn btn-success btn-xs']").Click
End If
driver.Wait 1000
Next
