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
                                                                 ':NA
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
    Call MsgBox("Report_Vx Status:" & Chr(10) & _
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

ALPHA_LogInformation (ThisWorkbook.Name & " opened by " & Application.UserName & " " & Format(Now, "yyyy-mm-dd hh:mm"))

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


'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'from alpha_make_dir

'https://stackoverflow.com/questions/43658276/create-folder-path-if-does-not-exist-saving-issue


'-----------------------------------------------------------------------------------------------------------
'from add ref page
'-----------------------------------------------------------------------------------------------------------

Sub alpha_AddReference()
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    Set vbProj = ActiveWorkbook.VBProject

    '~~> Check if "Microsoft VBScript Regular Expressions 5.5" is already added
    For Each chkRef In vbProj.References
        If chkRef.Name = "VBScript_RegExp_55" Then
            BoolExists = True
            GoTo CleanUp
        End If
    Next

    vbProj.References.AddFromFile "C:\Windows\SysWOW64\scrrun.dll"

CleanUp:
    If BoolExists = True Then
        MsgBox "Reference already exists"
    Else
        MsgBox "Reference Added Successfully"
    End If

    Set vbProj = Nothing
    Set VBAEditor = Nothing
End Sub

'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------

'requires reference to Microsoft Scripting Runtime
Function alpha_MkDir(strDir As String, strPath As String)

Dim fso As New FileSystemObject
Dim path As String

'examples for what are the input arguments
'strDir = "Folder"
'strPath = "C:\"

path = strPath & strDir

If Not fso.FolderExists(path) Then

' doesn't exist, so create the folder
          fso.CreateFolder path

End If

End Function


Sub test_b()
    Call MkDir("test log", "C:\")
End Sub
