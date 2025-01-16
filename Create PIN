'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------
' ********** PIN ENTER AUTOMATIC FROM EMAIL **************

''Const DontWaitUntilFinished = False, ShowWindow = 1, DontShowWindow = 0, WaitUntilFinished = True
''Set oShell = CreateObject("WScript.Shell")
''            'change path if needed -----------------------------------------------------------------------------------& Change Path if needed "C:\Selenium_VBA_Projects\Visa_Representment_VBASelenium\PIN\outputpin.txt"
''Cmd1 = "cmd /c C:\windows\system32\wscript.exe C:\Fee_Collection_Selenium\PIN\CreatePin.vbs " & "C:\Fee_Collection_Selenium\PIN\outputpin.txt"
''
''oShell.Run Cmd1, DontShowWindow, WaitUntilFinished
'''-----------------------------------------------------------'-----------------------------------------------------------
'''---- Read the text File ----
''FSOPasteTextFileContent
'''-----------------------------------------------------------'-----------------------------------------------------------
'''To Read PIN from Sheets("Instructions").Range("Q1")
'''-----------------------------------------------------------'-----------------------------------------------------------
'''TextString
''Sheets("Instructions").Activate
''driver.FindElementByXPath("//*[@id='46aAeiL']").SendKeys Range("Q1")


'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------



'''Read Text File - That is read PIN

Sub FSOPasteTextFileContent()
    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
                    'change path if needed --------------------
    Set FileToRead = FSO.OpenTextFile("C:\Fee_Collection_Selenium\PIN\outputpin.txt", ForReading) 'add here the path of your text file
    TextString = FileToRead.ReadAll
    
    FileToRead.Close
    'change Sheet name and column
    ThisWorkbook.Sheets("Instructions").Range("Q1").Value = TextString 'you can specify the worksheet and cell where to paste the text file’s content

End Sub


'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------
'-----------------------------------------------------------'-----------------------------------------------------------

'Create PIN
'Name - CreatePin.vbs

'write code in above vbs script as


'pinfile = WScript.Arguments.Item(0)
'Set objOutlook = CreateObject("Outlook.Application")
'Set OutlookSetNameSpace = objOutlook.GetNamespace("MAPI")
'Set OutlookSetFolder = OutlookSetNameSpace.GetDefaultFolder(6) ' 6 for Inbox folder
'Set objAllMails = OutlookSetFolder.Items
'MailProperty = "From"
'MailDateTime = FormatDateTime(Date)
'MailPropertyValue = "sysadmin@omnipay.ie"
'
'    Set ObjFilteredMails = objAllMails.Restrict("[" & MailProperty & "] = '" & MailPropertyValue & "' And [LastModificationTime] >= '" & MailDateTime & "'")
'    OrigUreadCount = ObjFilteredMails.count
'    counter = 0
'    Do
'        WScript.Sleep (5000)
'        Set ObjFilteredMails = objAllMails.Restrict("[" & MailProperty & "] = '" & MailPropertyValue & "' And [LastModificationTime] >= '" & MailDateTime & "'") ' needs quoted
'
'        UreadCount = ObjFilteredMails.count
'
'        If CInt(UreadCount) > CInt(OrigUreadCount) Or counter > 3 Then
'            MailString = ObjFilteredMails(CInt(UreadCount)).Body
'
'            StrArray = Split(MailString, vbNewLine)
'            For i = LBound(StrArray) To UBound(StrArray) - 1
'
'                If InStr(1, StrArray(i), "PIN") > 0 And InStr(1, StrArray(i), "Request") <= 0 And IsNumeric(Right(Trim(StrArray(i)), 5)) Then
'                    PinCode = Replace(StrArray(i), "PIN", "")
'                    PinCode = Replace(PinCode, ":", "")
'                    PinCode = Trim(Right(Trim(PinCode), 10))
'                End If
'            Next
'            Set FSO = CreateObject("Scripting.FileSystemObject")
'            Set ts = FSO.CreateTextFile(pinfile)
'            ts.WriteLine PinCode
'            ts.Close
'            Set FSO = Nothing
'            Set ts = Nothing
'            Exit Do
'        End If
'        counter = counter + 1
'    Loop
'Set objOutlook = Nothing



'create notepad as "outputpin.txt"


