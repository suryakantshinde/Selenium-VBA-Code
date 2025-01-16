Sub MoveFiles()

  'Declare Variables
Dim FSO
Dim sFile As String
Dim sSfolder As String
Dim sDFolder As String

'This is Your File Name which you want to Copy
sFile = "e-charge-slip.pdf"

'Change to match the source folder path
sSfolder = Sheets("Data").Range("L4")

'Change to match the destination folder path
sDFolder = Sheets("Data").Range("L5")


'Create Object
Set FSO = CreateObject("Scripting.FileSystemObject")
'-------

'Checking If File Is Located in the Source Folder
If Not FSO.FileExists(sSfolder & sFile) Then
    'MsgBox "Specified File Not Found", vbInformation, "Not Found"
    
'Copying If the Same File is Not Located in the Destination Folder
ElseIf Not FSO.FileExists(sDFolder & sFile) Then
    FSO.CopyFile (sSfolder & sFile), sDFolder, True
    '==================================================================================================
    'MsgBox "Specified File Copied Successfully", vbInformation, "Done!"
Else
    'MsgBox "Specified File Already Exists In The Destination Folder", vbExclamation, "File Already Exists"
End If
End Sub
Sub DeleteFile()
   
Dim MyFolderPath As String
'Path where the folder is located, change the path as per your requirement
MyFolderPath = Sheets("Data").Range("L4").Value
'Check if the folder already exists or not
If Dir(MyFolderPath, vbDirectory) <> "" Then
'Delete all xlsx files in the folder and subfolders ,   to delete any type of files use [*.*] istead of *.xlsx
        Call Shell("cmd.exe /S /C" & "Del " & MyFolderPath & "\e-charge-slip.pdf" & "  /S /Q")
'MsgBox "Files deleted"
Else
'MsgBox "Folder does not Exist"
End If
End Sub

Sub Rename_the_File()

Dim OldFile As String
Dim NewFile As String
Dim MyPath As String
Dim MyFile As String
Dim Num As Integer

Num = 1
MyPath = Sheets("Data").Range("L6")
MyFile = Dir(MyPath & "*.pdf")

On Error GoTo Er_R
Do Until IsEmpty(MyFile)
    OldFile = MyFile
    NewFile = Range("B" & lr).Value & ".pdf"
    'Name OldFile As NewFile ''''''''''''''''''''''''''''Error here
    Name MyPath & OldFile As MyPath & NewFile
Num = Num + 1
MyFile = Dir()
Loop
Er_R:
End Sub
'ActiveWorkbook.SaveAs Filename:="C:\Users\f4i65kw\Desktop\PRE AUTH CANCELLATION\Download_Charge_Slip\" & _
Range("B" & i).Value

