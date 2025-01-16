Dim FilePath, FileName, ModuleName, ProcName
'#################################################################################################################

FilePath = "<folder path>"

FileName = "<macro name with extension>"

ModuleName = "<module name of the macro>"

ProcName = "<procedure name>"

'#################################################################################################################
Dim oXL 
Dim wb

Set oXL = CreateObject("Excel.Application")
oXL.AutomationSecurity = 1
oXL.Visible = True 
oXL.EnableEvents = False
Set wb = oXL.Workbooks.Open(FilePath &  FileName)
oXL.EnableEvents = True
oXL.Application.OnTime Now + TimeValue("0:00:05"), FileName & "!" & ModuleName & "." & ProcName

'wb.Activate
Set oXL = Nothing

'#################################################################################################################
