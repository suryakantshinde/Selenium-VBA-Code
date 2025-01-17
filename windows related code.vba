Option Explicit
Dim brw As Selenium.ChromeDriver

Sub collectionloop()

Set brw = New Selenium.ChromeDriver
brw.Get "https://en.wikipedia.org/wiki/Main_Page"

Dim mykeys As Selenium.keys
Set mykeys = New Selenium.keys


brw.FindElementByLinkText("Current events").Click mykeys.Control

brw.FindElementByLinkText("Contact us").Click mykeys.Control

Dim mywindow As Selenium.Window

For Each mywindow In brw.Windows

                   '
                'Debug.Print mywindow.Title

        If mywindow.Title Like "*" & "Portal" & "*" Then
                    mywindow.Activate
                    mywindow.Close
        Else
        End If

Next mywindow

brw.Windows(1).Activate

End Sub

'---------------------------------------------------------------------------------------------------
Sub define_windows()

Set brw = New Selenium.ChromeDriver
brw.Get "https://en.wikipedia.org/wiki/Main_Page"

Dim mykeys As Selenium.keys
Set mykeys = New Selenium.keys

Dim main_w As Selenium.Window
Set main_w = brw.Windows(1)


brw.FindElementByLinkText("Current events").Click mykeys.Control
Dim currentevent_w As Selenium.Window
Set currentevent_w = brw.Windows(2)


brw.FindElementByLinkText("Contact us").Click mykeys.Control

Dim contact_w As Selenium.Window
Set contact_w = brw.Windows(3)


main_w.Activate
contact_w.Activate
currentevent_w.Activate
currentevent_w.Close
contact_w.Close
main_w.Close

End Sub
'-------------------------------------------------------------------------------------------
 Option Explicit

Dim brw As Selenium.ChromeDriver


Sub newwindow_open()

Set brw = New Selenium.ChromeDriver
brw.Get "https://en.wikipedia.org/wiki/Main_Page"

Dim mykeys As Selenium.keys
Set mykeys = New Selenium.keys


brw.FindElementByLinkText("Current events").Click mykeys.Control
'debug.Print brw.Title  ' url



brw.Window.SwitchToNextWindow
Debug.Print brw.Title
brw.FindElementByLinkText("Sport").Click

brw.Window.SwitchToPreviousWindow

brw.FindElementByLinkText("Contact us").Click mykeys.Control

brw.Window.SwitchToNextWindow
brw.Window.SwitchToNextWindow


End Sub
'-------------------------------------------------------------------------------------------

Sub newwindow_open2()

Set brw = New Selenium.ChromeDriver
brw.Get "https://en.wikipedia.org/wiki/Main_Page"

Dim mykeys As Selenium.keys
Set mykeys = New Selenium.keys


brw.FindElementByLinkText("Current events").Click mykeys.Control

brw.FindElementByLinkText("Contact us").Click mykeys.Control


brw.Windows(3).Activate
brw.Windows(2).Activate
brw.Windows(1).Activate


End Sub               

