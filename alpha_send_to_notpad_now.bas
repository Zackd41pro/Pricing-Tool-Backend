Attribute VB_Name = "alpha_send_to_notpad_now"

'how to send to notpad as a post
Sub test()
    Dim myApp As String
    myApp = Shell("Notepad", vbNormalFocus)
    SendKeys "test", True
    SendKeys "hello", True
End Sub
