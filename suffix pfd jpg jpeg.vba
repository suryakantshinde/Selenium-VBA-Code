driver.Wait 50
Dim CF As WebElement
'change path
'file_path = "C:\MasterCard_Connect_Docs\Attachments\"
file_path = "C:\2024-Automations\3-MasterCard_Connect_Docs\PDF_Attachments\"
CaseNo = Range("B" & i)
sk_str = ""

For CC = Asc("a") To Asc("z")
    If CC = 0 Then
        ' suffix = ""
    Else
        suffix = Chr(CC)
    End If
    
    If Dir(file_path & CaseNo & suffix & ".pdf") <> "" Then
        sk_str = file_path & CaseNo & suffix & ".pdf"
        Set CF = driver.FindElementByXPath("//input[@type='file']")
        CF.SendKeys (sk_str)
        driver.Wait 1000
    End If
    
    If Dir(file_path & CaseNo & suffix & ".jpg") <> "" Then
        sk_str = file_path & CaseNo & suffix & ".jpg"
        Set CF = driver.FindElementByXPath("//input[@type='file']")
        CF.SendKeys (sk_str)
        driver.Wait 1000
    End If
    
    If Dir(file_path & CaseNo & suffix & ".jpeg") <> "" Then
        sk_str = file_path & CaseNo & suffix & ".jpeg"
        Set CF = driver.FindElementByXPath("//input[@type='file']")
        CF.SendKeys (sk_str)
        driver.Wait 1000
    End If
Next CC
    driver.Wait 1000
