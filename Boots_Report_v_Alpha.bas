Attribute VB_Name = "Boots_Report_v_Alpha"
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                                                                            'Author: Zachary Daugherty
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                                '©: 2020-2021
                    'If you want to make edits or additions to this module please contact me to make sure it is included with the live production group.
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Nessasary Librarys
                                                        'made for :Report_VX
                                                                 ':boots_main
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Purpose Case
                    'This Module is built to handle exporting of information
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function status()
    Call MsgBox("Boots_Report_Vx is in alpha!", , "Warning!")
    Call MsgBox("Boots_Report_Vx Status:" & Chr(10) & _
    "------------------------------------------------------------" & Chr(10) & _
    "Public functions: " & Chr(10) & _
    "" & Chr(10) & _
    Chr(10) & "Private functions:" & Chr(10) & _
    "" & Chr(10) & _
    "", , "showing status for Boots_Report_Vx")
End Function


'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'from alpha generate log

'https://www.exceltip.com/files-workbook-and-worksheets-in-vba/log-files-using-vba-in-microsoft-excel.html

Sub ALPHA_LogInformation(LogMessage As String)
    Const LogFileName As String = "C:\test log\TEXTFILE.LOG"
    Dim FileNum As Integer
    
    FileNum = FreeFile ' next file number
    Open LogFileName For Append As #FileNum ' creates the file if it doesn't exist
        Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file

End Sub


Public Sub ALPHA_DisplayLastLogInformation()
'from origen
'    Const LogFileName As String = "C:\test log\TEXTFILE.LOG"
'    Dim FileNum As Integer, tLine As String
'
'        FileNum = FreeFile ' next file number
'        Open LogFileName For Input Access Read Shared As #f ' open the file for reading
'            Do While Not EOF(FileNum)
'            Line Input #FileNum, tLine ' read a line from the text file
'            Loop ' until the last line is read
'        Close #FileNum ' close the file
'
'    MsgBox tLine, vbInformation, "Last log information:"



'from 'http://codevba.com/office/read_text_file_line_by_line.htm#.X1JUDPZFwdU'
    Dim strFilename As String: strFilename = "C:\test log\TEXTFILE.LOG"
    Dim strTextLine As String
    Dim iFile As Integer: iFile = FreeFile
    
    Open strFilename For Input As #iFile
    Do Until EOF(1)
        Line Input #1, strTextLine
        
    Loop
    Close #iFile
    Stop
End Sub


Sub ALPHA_DeleteLogFile(FullFileName As String)

    On Error Resume Next ' ignore possible errors
        Kill FullFileName ' delete the file if it exists and it is possible
    On Error GoTo 0 ' break on errors

End Sub


Private Sub ALPHA_Workbook_Open()

ALPHA_LogInformation (ThisWorkbook.Name & " opened by " & Application.username & " " & Format(Now, "yyyy-mm-dd hh:mm"))

End Sub

Sub ALPHA_OpenInNotepad()
'https://www.mrexcel.com/board/threads/excel-vba-open-txt-file-as-notepad-not-excel.410578/
Dim MyTxtFile
    MyTxtFile = Shell("C:\WINDOWS\notepad.exe C:\test log\TEXTFILE.LOG", 1)
End Sub


'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'from alpha send to notpad now

'how to send to notpad as a post
Sub test_a()
    Dim myApp As String
    myApp = Shell("Notepad", vbNormalFocus)
    SendKeys "test", True
    SendKeys "hello", True
End Sub
