Attribute VB_Name = "Boots_Report_v_Alpha"
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                                                                            'Author: Zachary Daugherty
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                                '�: 2020-2021
                    'If you want to make edits or additions to this module please contact me to make sure it is included with the live production group.
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Nessasary Librarys
                                                        'made for :Report_VX
                                                                 ':boots_main
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Purpose Case
                    'This Module is built to handle the exporting & importing of information
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Enum boots_report_pos
    'make date
        i_date_made_row = 1
            i_date_made_col = 10
    'log len
        log_row_length_row = 1
            log_row_lenght_col = 11
    'data
        p_indent_row = 1
            p_indent_col = 1
                p_time_row = boots_report_pos.p_indent_row + 0
                    p_time_col = boots_report_pos.p_indent_col + 1
                        p_text_row = boots_report_pos.p_time_row + 0
                            p_text_col = boots_report_pos.p_time_col + 1
End Enum

Enum Flush_selection
    Save
    Save_Exit
    Delete
    CleanUp
End Enum

Enum Push_selection
    text
    table_open
    table_close
    Variable
    Trigger_S
    Trigger_e
    Flag
    Error_
    Display_now
End Enum

Const Log_indent_spaces As String = "    "

Public Function status()
    MsgBox ("old to delete")
'    Call MsgBox("Boots_Report_Vx is in alpha!", , "Warning!")
'    Call MsgBox("Boots_Report_Vx Status:" & Chr(10) & _
'    "------------------------------------------------------------" & Chr(10) & _
'    "Public functions: " & Chr(10) & _
'    "" & Chr(10) & _
'    Chr(10) & "Private functions:" & Chr(10) & _
'    "" & Chr(10) & _
'    "", , "showing status for Boots_Report_Vx")
End Function

Public Function Log_get_indent_value_V0() As Long
    Dim wb As Workbook
    Dim sht As Worksheet
    
    Set wb = ActiveWorkbook
    On Error Resume Next
        Set sht = wb.Sheets("LOG_" & Boots_Main_V_alpha.get_username)
    On Error Resume Next
        Log_get_indent_value_V0 = sht.Cells(Boots_Report_v_Alpha.Log_get_length_of_log_list_V1, 1).value
End Function

Public Function Log_get_length_of_log_list_V1() As Long
    'define variables
        'addresses
            Dim wb As Workbook
            Dim sht As Worksheet
            Dim home_sht As Worksheet
        'containers
            Dim s As String
    'set variables
        'addresses
            Set wb = ActiveWorkbook
            Set home_sht = ActiveSheet
            'get namespace of the sheet
                Set sht = wb.Sheets("LOG_" & Boots_Main_V_alpha.get_username)
        'containers
            s = "Empty"
            Log_get_length_of_log_list_V1 = 0
        'alpha check container
            If (sht.Cells(boots_report_pos.log_row_length_row, boots_report_pos.log_row_lenght_col).value = "") Then
                sht.Cells(boots_report_pos.log_row_length_row, boots_report_pos.log_row_lenght_col).Formula = "=COUNT(A:A)"
            End If
    'fetch length of the log
        Log_get_length_of_log_list_V1 = sht.Cells(boots_report_pos.log_row_length_row, boots_report_pos.log_row_lenght_col).value
    'cleanup
        s = "Empty"
End Function

Public Function Log_Initalize(Optional Further_definitions As String) As Boolean
    'define variables
        'addresses
            Dim wb As Workbook
            Dim sht As Worksheet
            Dim home_sht As Worksheet
        'containers
            Dim s As String
            Dim i As Long
            Dim j As Long
            Dim count As Long
    'set variables
        'addresses
            Set wb = ActiveWorkbook
            Set home_sht = ActiveSheet
            'create log session or update log session
                s = "LOG_" & Boots_Main_V_alpha.get_username
                Application.Wait (Now + TimeValue("0:00:05"))
                If (Boots_Main_V_alpha.sheet_exist(wb, s) = False) Then 'if the log page dont exist make it
                'make sheet
                    Call Boots_Main_V_alpha.make_sheet_V1(wb, s, -1, "d_report")
                Else
                'sheet exist
                    Call Boots_Report_v_Alpha.Log_Flush(Save)
                    Call Boots_Report_v_Alpha.Log_Flush(Delete)
                    Call Boots_Main_V_alpha.make_sheet_V1(wb, s, -1, "d_report")
                End If
            'set sht
                Set sht = wb.Sheets(s)
        'containers
            s = "Empty"
            i = -1
            j = -1
            count = -1
    'format the log and add name space
        sht.Activate
        Boots_Report_v_Alpha.Log_format_page
        sht.visible = 2
    'make note of the modules that are currently installed
        'update table
            Boots_Main_V_alpha.get_project_files
        'push record
            Call Boots_Report_v_Alpha.Log_Push(text, " Table Displaying Currently Installed Project Object Files:...")
            Boots_Report_v_Alpha.Log_Push (table_open)
            'make push of project objects
                i = Boots_Main_V_alpha.get_project_files(na)
                For count = 1 To i
                    If (count <> i) Then
                    'if entry is not the last one do this
                        Call Boots_Report_v_Alpha.Log_Push(text, Boots_Main_V_alpha.get_project_files(get_index, count))
                        'get the plugins required for this module
                            'get length of list
                                j = Boots_Report_v_Alpha.Log_get_length_of_log_list_V1
                            'determine if needed to look for function list in the specified project file E.G. meaning the version reported back was not 'NA'
                                s = sht.Cells(boots_report_pos.p_text_row + j - 1, boots_report_pos.p_text_col).value
                                'parse
                                    s = String_V1.Disassociate_by_Char_V2(">", s, Right_C, "d_report")
                                        s = String_V1.Disassociate_by_Char_V2(">", s, Right_C, "d_report")
                                            s = String_V1.Disassociate_by_Char_V2("<", s, Right_C, "d_report")
                                                s = String_V1.Disassociate_by_Char_V2(">", s, Left_C, "d_report")
                                'check the entry for not NA entry
                                    If (UCase(s) <> "NA") Then
                                        'Get module dependancy
                                            'setup
                                                'get the namespace of the ENV for modules
                                                    s = sht.Cells(boots_report_pos.p_text_row + j - 1, boots_report_pos.p_text_col).value
                                                    'parse for namespace
                                                        s = String_V1.Disassociate_by_Char_V2(">", s, Right_C, "d_report")
                                                            s = String_V1.Disassociate_by_Char_V2("<", s, Right_C, "d_report")
                                                                s = String_V1.Disassociate_by_Char_V2(">", s, Left_C, "d_report")
                                                'fetch Module dependables
                                                    s = s + ".LOG_push_project_file_requirements"
                                                    s = Run(s)
                                            'paste log lines for module dependancy
                                                'indent line
                                                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                                'list
                                                    Call Boots_Report_v_Alpha.Log_Push(text, "Showing Object File Dependants:...")
                                                'indent line
                                                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                                'pull out variables
                                                    Call Boots_Report_v_Alpha.Log_Push(text, s)
                                                'de-indent line
                                                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                            'cleanup
                                                s = "Empty"
                                        'get registered function stability status
                                            'setup
                                                'prep log list
                                                    'list
                                                        Call Boots_Report_v_Alpha.Log_Push(text, "Reporting Project Function Stability Status:...")
                                                    'indent line
                                                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                                'get the namespace of the ENV for modules
                                                    s = sht.Cells(boots_report_pos.p_text_row + j - 1, boots_report_pos.p_text_col).value
                                                    'parse for namespace
                                                        s = String_V1.Disassociate_by_Char_V2(">", s, Right_C, "d_report")
                                                            s = String_V1.Disassociate_by_Char_V2("<", s, Right_C, "d_report")
                                                                s = String_V1.Disassociate_by_Char_V2(">", s, Left_C, "d_report")
                                            'paste log lines for function stability status
                                                'push functions log list
                                                    s = s + ".LOG_Push_Functions_v1"
                                                    s = Run(s)
                                                'de-indent line
                                                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                        'de-indent line for project file end
                                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                    End If
                    'cleanup
                        s = "Empty"
                        j = -1
                        
                    Else
                    'if last field entry do this (last field will give total list length
                        Call Boots_Report_v_Alpha.Log_Push(text, "total entrys:" & (count - 1)) ' gives this lengh of all the listed projectfiles
                    End If
                Next count
                'cleanup
                    i = -1
                    count = -1
                    s = "empty"
            'table close
                Boots_Report_v_Alpha.Log_Push (table_close)
            'programs start
                Call Boots_Report_v_Alpha.Log_Push(text, ".")
                Call Boots_Report_v_Alpha.Log_Push(text, ".")
                Call Boots_Report_v_Alpha.Log_Push(text, ".")
                Call Boots_Report_v_Alpha.Log_Push(text, ".")
                Call Boots_Report_v_Alpha.Log_Push(text, ".")
                Call Boots_Report_v_Alpha.Log_Push(text, "__________Finished Object Name Declare!...")
                Call Boots_Report_v_Alpha.Log_Push(text, "__________Program Start!...")
End Function

Public Function Log_Push(ByVal Action As Push_selection, Optional text As String, Optional more_instructions As String) As Variant
    'this function is made to push all log entrys to a sheet stored in the project so that if there are errors it is easy to report infomration on what went wrong, or ect.
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            Log_Push = "Log_Push - Public - Stable 10/30/2020 - help file:N"
            Exit Function
        End If
    
    'code start
        'define variables
            'addresses
                Dim wb As Workbook
                Dim sht As Worksheet
                Dim home_sht As Worksheet
                Dim Logfile_Env
            'container
                Dim i As Long
                Dim s As String
        'set variables
            Set wb = ActiveWorkbook
            Set home_sht = ActiveSheet
            On Error GoTo Log_Push_Error_Unable_To_Find_Log
                Set sht = wb.Sheets("LOG_" & Boots_Main_V_alpha.get_username)
            On Error GoTo 0
        'find open position on the table
            i = Boots_Report_v_Alpha.Log_get_length_of_log_list_V1
        'check if there is an indent mark in the empty pos
            If (sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value <> "") Then
                Stop
                'check if it is plus
                    If (sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = "+") Then
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value + 1
                        GoTo Log_Push_exit_indent
                    End If
                'check if it is minus
                    If (sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = "-") Then
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value - 1
                        'check is indent is now negative if so make 0
                            If (s < 0) Then
                                sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = 0
                                MsgBox ("Log_Push Error: indent '-' made the indent value less than zero now made zero")
                            End If
                            GoTo Log_Push_exit_indent
                    End If
            Else
            'get indent value from line above
                'check for a valid grab position
                    'MsgBox ("need to add a proper check for row overflow")
                    If ((i <= 0) Or (i >= 1048577)) Then 'meaning the row value is 0 or greater than the max on a sheet
                        If (i <= 0) Then
                            i = 1
                        End If
                        If (i >= 1048577) Then
                            Stop
                        End If
                    End If
                'get indent value
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                'check if indent is x<0 then make zero
                    If (sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value < 0) Then
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = 0
                    End If
            End If
Log_Push_exit_indent:
            'set now
                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
        'run action
            Select Case Action
            Case Push_selection.Display_now
                'final text
                    Boots_Report_v_Alpha.Log_Before_Close_or_Error
                'compress log removes blank lines
                    Boots_Report_v_Alpha.Log_compress_blank_space
                'export
                    Boots_Report_v_Alpha.Log_Flush (Save)
                'display now
                    s = "C:\WINDOWS\notepad.exe " & root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\" & "Users\" & Boots_Main_V_alpha.get_username & "\" & "Log-" & Month(Date) & "-" & Day(Date) & "-" & Year(Date) & ".log"
                    Logfile_Env = Shell(s, 1)
                    Application.DisplayAlerts = False
                    Application.ScreenUpdating = False
                        sht.visible = xlSheetVisible
                        sht.Delete
                    Application.DisplayAlerts = True
                    Application.ScreenUpdating = True
                    ActiveWorkbook.Close
            Case Push_selection.Error_
                'get indent
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value + 1
                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                            i = i + 1
                                sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                                i = i + 1
                'open a new error report FLAG
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "________________________________ERROR TRIGGERED_______________________________"
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                i = i + 1
            Case Push_selection.Flag
                'open new flag
                    'get indent
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value + 1
                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "????????????????????????????????????????????????????????????????????????????????"
                    i = i + 1
                    'new flag line
                        'get indent
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                        'set now
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                        'title
                            sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "????????????????????????????????????????????????????????????????????????????????"
                'Flag Text
                    'get indent
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "_________________________________FLAG TRIGGERED_______________________________"
                        i = i + 1
            Case Push_selection.text
                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = text
            Case Push_selection.Trigger_e
                sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value - 1
                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = ""
                'check if indent is x<0 then make zero
                    If (sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value < 0) Then
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = 0
                    End If
            Case Push_selection.Trigger_S
                sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value + 1
                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = ""
            Case Push_selection.Variable
                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "____Displaying Variable: " & text & "____"
            Case Push_selection.table_close
                'set text
                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "________________________________________________________________________________"
                    i = i + 1
                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/"
                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                    i = i + 1
                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value - 1
            Case Push_selection.table_open
                'set text
                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/"
                    i = i + 1
                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = "___________________________________NEW TABLE____________________________________"
                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value + 1
                    sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value + 1
            Case Else
                Call MsgBox("Fatal Error: Module 'Boots_Report' Called 'Push Log' and was unable to determine a action selection", , "Fatal Error: Module 'Boots_Report' Called 'Push Log'")
                Stop
            End Select
    'code end
        'cleanup
            Log_Push = True
            Exit Function
    'error reporting
Log_Push_Error_Unable_To_Find_Log:
        'Log_Push_Error_Unable_To_Find_Log:
            Call Boots_Report_v_Alpha.Push_notification_message("Boots_report_v_alpha.log_push: ERROR! log_push was unable to find the log sheet. the log could of been flushed or removed...")
            Stop
            End
End Function

Public Function Log_Flush(ByVal Action As Flush_selection, Optional Further_definitions As String)
    Dim wb As Workbook
    Dim sht As Worksheet
    
    Dim i As Long
    Dim j As Long
    Dim line As Long
    Dim s As String
    
    Set wb = ActiveWorkbook
    Set sht = wb.Sheets("LOG_" & Boots_Main_V_alpha.get_username)
    
    'get log len
        i = Log_get_length_of_log_list_V1
        line = 0
    'check for delete action
        If (Action = Delete) Then
            Application.DisplayAlerts = False
            Application.ScreenUpdating = False
                On Error Resume Next
                    sht.visible = xlSheetVisible
                On Error Resume Next
                    sht.Delete
                On Error GoTo 0
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
            Exit Function
        End If
    'check for cleanup to exit
        Call Log_compress_blank_space
        If (Action = CleanUp) Then
            Exit Function
        End If
    'post to log and delete lines that are posted
        If (Action = Save) Then
            'get line for posting then delete that line from log page
                For line = 0 To i - 1
                    'get date
                        s = sht.Cells(boots_report_pos.p_time_row + line, boots_report_pos.p_time_col).value & " == "
                    'find indent value and add in front of text
                    For j = 1 To sht.Cells(boots_report_pos.p_indent_row + line, boots_report_pos.p_indent_col).value
                        s = s + Log_indent_spaces
                    Next j
                    'install line to post
                        s = s & sht.Cells(boots_report_pos.p_text_row + line, boots_report_pos.p_text_col).value
                        Boots_Report_v_Alpha.Log_Flush_Line_pvt_v0 (s)
                    'delete line
                        sht.Cells(boots_report_pos.p_indent_row + line, boots_report_pos.p_indent_col).value = ""
                        sht.Cells(boots_report_pos.p_time_row + line, boots_report_pos.p_time_col).value = ""
                        sht.Cells(boots_report_pos.p_text_row + line, boots_report_pos.p_text_col).value = ""
                    'cleanup for line get
                        s = ""
                Next line
        End If
    'delete log page as saving happens in the action group above
        If (Action = Save_Exit) Then
            'close or error text
                Boots_Report_v_Alpha.Log_Before_Close_or_Error
            'get line for posting then delete that line from log page
                For line = 0 To i - 1
                    'get date
                        s = sht.Cells(boots_report_pos.p_time_row + line, boots_report_pos.p_time_col).value & " == "
                    'find indent value and add in front of text
                    For j = 1 To sht.Cells(boots_report_pos.p_indent_row + line, boots_report_pos.p_indent_col).value
                        s = s + Log_indent_spaces
                    Next j
                    'install line to post
                        s = s & sht.Cells(boots_report_pos.p_text_row + line, boots_report_pos.p_text_col).value
                        Boots_Report_v_Alpha.Log_Flush_Line_pvt_v0 (s)
                    'delete line
                        sht.Cells(boots_report_pos.p_indent_row + line, boots_report_pos.p_indent_col).value = ""
                        sht.Cells(boots_report_pos.p_time_row + line, boots_report_pos.p_time_col).value = ""
                        sht.Cells(boots_report_pos.p_text_row + line, boots_report_pos.p_text_col).value = ""
                    'cleanup for line get
                        s = ""
                Next line
            'sht delete
                sht.Delete
        End If
    
End Function

Private Sub Log_format_page()
    Cells.Select
    With Selection.Font
        .Color = -16711936
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With ActiveWorkbook.Sheets("LOG_" & Boots_Main_V_alpha.get_username).Tab
        .Color = 65280
        .TintAndShade = 0
    End With
    x = 2
    Range("A1").Select
    Cells(1, 10).value = Date
    Cells(1, 1).value = 0
    Cells(1, 2).value = Now()
    Cells(1, 3).value = "LOG Session Created from " & ActiveWorkbook.Name
    
    Cells(x, 3).value = "LOG_" & Boots_Main_V_alpha.get_username & " was created on " & Now()
    Cells(x, 2).value = Now()
    Cells(x, 1).value = 0
End Sub

Private Sub Log_compress_blank_space()
    Dim wb As Workbook
    Dim sht As Worksheet
    Dim i As Long
    Dim j As Long
    Dim z As Long
    Dim s As String
    
    Set wb = ActiveWorkbook
    Set sht = wb.Sheets("LOG_" & Boots_Main_V_alpha.get_username)
    i = 0
    
    For j = 0 To 50
        s = sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value
        If (s = "") Then
            i = i + 1
            s = i & ":" & i
            sht.Rows(s).Delete Shift:=xlUp
            j = 0
            i = 0
            z = z + 1
            If (z > 1000) Then
                j = 50
            End If
        Else
            j = 0
            i = i + 1
        End If
    Next j
End Sub

Private Function Log_Flush_Line_pvt_v0(ByVal LogMessage As String) As Boolean
'https://www.exceltip.com/files-workbook-and-worksheets-in-vba/log-files-using-vba-in-microsoft-excel.html

    Dim s As String
    Dim s_2 As String
    Dim LogFileName As String
    
    'check for locations existance if not make
        Call Boots_Report_v_Alpha.DIR_Make(root.get_save_location, root.get_drive_location)
        Call Boots_Report_v_Alpha.DIR_Make(root.get_project_name, root.get_drive_location & root.get_save_location)
        Call Boots_Report_v_Alpha.DIR_Make(root.get_version & "\", root.get_drive_location & root.get_save_location & root.get_project_name)
        Call Boots_Report_v_Alpha.DIR_Make("Users\", root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\")
        Call Boots_Report_v_Alpha.DIR_Make(Boots_Main_V_alpha.get_username & "\", root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\" & "Users\")
        
    
    s_2 = root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\" & "Users\" & Boots_Main_V_alpha.get_username & "\" & "Log-" & Month(Date) & "-" & Day(Date) & "-" & Year(Date) & ".log"
    LogFileName = s_2

    'Boots_report_v_alpha.DIR_Make(
    Dim FileNum As Integer
        
    FileNum = FreeFile ' next file number
    Open LogFileName For Append As #FileNum ' creates the file if it doesn't exist
        Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file

End Function
 
Public Function Log_Before_Close_or_Error() As Variant
    'log
        For i = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
        Next i
        Call Boots_Report_v_Alpha.Log_Push(text, ".")
        Call Boots_Report_v_Alpha.Log_Push(text, ".")
        Call Boots_Report_v_Alpha.Log_Push(text, ".")
        Call Boots_Report_v_Alpha.Log_Push(text, ".")
        Call Boots_Report_v_Alpha.Log_Push(text, ".")
        Call Boots_Report_v_Alpha.Log_Push(text, "--------------------------------------------------------------------------------------------------------------------")
        Call Boots_Report_v_Alpha.Log_Push(text, "--------------------------------------------------------------------------------------------------------------------")
        Call Boots_Report_v_Alpha.Log_Push(text, "End of workbook session...")
        Call Boots_Report_v_Alpha.Log_Push(text, "Preping Save to file for log...")
        Call Boots_Report_v_Alpha.Log_Push(text, "Finished...")
        Call Boots_Report_v_Alpha.Log_Push(text, "--------------------------------------------------------------------------------------------------------------------")
        Call Boots_Report_v_Alpha.Log_Push(text, "--------------------------------------------------------------------------------------------------------------------")
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

                                                                        'DIR LIB
                                                                        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Public Function DIR_Make(strDir As String, strPath As String)
'https://stackoverflow.com/questions/43658276/create-folder-path-if-does-not-exist-saving-issue
'requires reference to Microsoft Scripting Runtime
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

Public Function DIR_Flush() As Variant
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
        On Error Resume Next
            ActiveWorkbook.Sheets(Boots_Main_V_alpha.get_username & "_DIR_Search").visible = -1
        On Error GoTo 0
        On Error Resume Next
            ActiveWorkbook.Sheets(Boots_Main_V_alpha.get_username & "_DIR_Search").Delete
        On Error GoTo 0
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    DIR_Flush = True
End Function

Public Function DIR_GET_vA(ByVal DirLocation As String) As Variant

    'VIA (http://www.xl-central.com/list-the-files-in-a-folder-and-subfolders.html)

    'Set a reference to Microsoft Scripting Runtime by using
    'Tools > References in the Visual Basic Editor (Alt+F11)
    
    'Declare the variables
    Dim objFSO As Scripting.FileSystemObject
    Dim objTopFolder As Scripting.Folder
    Dim strTopFolderName As String
    Dim sht As Worksheet
    
    'check for old version
        For Each sht In ThisWorkbook.Worksheets
                If Application.Proper(sht.Name) = Application.Proper(Boots_Main_V_alpha.get_username & "_DIR_Search") Then
                    Application.DisplayAlerts = False
                    Application.ScreenUpdating = False
                        sht.visible = xlSheetVisible
                        sht.Delete
                    Application.DisplayAlerts = True
                    Application.ScreenUpdating = True
                    Exit For
                End If
        Next sht
    'create new sht
        Call ActiveWorkbook.Sheets.Add
        ActiveSheet.Name = Boots_Main_V_alpha.get_username & "_DIR_Search"
        Set sht = ActiveSheet
        Boots_Report_v_Alpha.DIR_format_page_v0
        sht.visible = xlSheetVeryHidden
    'Insert the headers for Columns A through F
    sht.Range("A1").value = "File Name"
    sht.Range("B1").value = "Path"
    sht.Range("C1").value = "File Size"
    sht.Range("D1").value = "File Type"
    sht.Range("E1").value = "Date Created"
    sht.Range("F1").value = "Date Last Accessed"
    sht.Range("G1").value = "Date Last Modified"
    
    'Assign the top folder to a variable
    strTopFolderName = DirLocation
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Get the top folder
    Set objTopFolder = objFSO.GetFolder(strTopFolderName)
    
    'Call the DIR_RecursiveFolder_VA routine
    Call Boots_Report_v_Alpha.DIR_RecursiveFolder_VA(objTopFolder, True, sht)
    
    'Change the width of the columns to achieve the best fit
    Columns.AutoFit
    
    DIR_GET_vA = True
End Function


Private Function DIR_RecursiveFolder_VA(objFolder As Scripting.Folder, _
    IncludeSubFolders As Boolean, sht As Worksheet) As Variant

    'VIA (http://www.xl-central.com/list-the-files-in-a-folder-and-subfolders.html)

    'Declare the variables
    Dim objFile As Scripting.File
    Dim objSubFolder As Scripting.Folder
    Dim NextRow As Long
    
    'Find the next available row
    NextRow = sht.Cells(Rows.count, "A").End(xlUp).row + 1
    
    'Loop through each file in the folder
    For Each objFile In objFolder.Files
        sht.Cells(NextRow, "A").value = objFile.Name
        sht.Cells(NextRow, "B").value = objFile.path
        sht.Cells(NextRow, "C").value = objFile.size
        sht.Cells(NextRow, "D").value = objFile.Type
        sht.Cells(NextRow, "E").value = objFile.DateCreated
        sht.Cells(NextRow, "F").value = objFile.DateLastAccessed
        sht.Cells(NextRow, "G").value = objFile.DateLastModified
        NextRow = NextRow + 1
    Next objFile
    
    'Loop through files in the subfolders
    If IncludeSubFolders Then
        For Each objSubFolder In objFolder.SubFolders
            Call DIR_RecursiveFolder_VA(objSubFolder, True, sht)
        Next objSubFolder
    End If
    DIR_RecursiveFolder_VA = True
End Function

Private Function DIR_format_page_v0() As Variant
    Cells.Select
    With Selection.Font
        .Color = -16711936
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With ActiveWorkbook.Sheets(Boots_Main_V_alpha.get_username & "_DIR_Search").Tab
        .Color = 65280
        .TintAndShade = 0
    End With
    DIR_format_page_v0 = True
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

                                                                        'Alpha Code
                                                                        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


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


'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'from alpha send to notpad now

'how to send to notpad as a post
Sub Push_notification_message(ByVal text_to_display As String)
    Dim myApp As String
    Dim s As String
    myApp = Shell("Notepad", vbNormalFocus)
    SendKeys "___________________________________This is an Automated message from the terminal___________________________________" & Chr(10), True
SendKeys text_to_display, True
End Sub


'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'depreciated and old code
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/



Public Function Log_get_length_of_log_list_VA() As Long
    'define variables
        'addresses
            Dim wb As Workbook
            Dim sht As Worksheet
            Dim home_sht As Worksheet
        'containers
            Dim s As String
    'set variables
        'addresses
            Set wb = ActiveWorkbook
            Set home_sht = ActiveSheet
            'get namespace of the sheet
                Set sht = wb.Sheets("LOG_" & Boots_Main_V_alpha.get_username)
        'containers
            s = "Empty"
            Log_get_length_of_log_list_VA = 0
    'fetch length of the log
restart_Log_get_length_of_log_list_VA:
        s = sht.Cells(boots_report_pos.p_indent_row + Log_get_length_of_log_list_VA, boots_report_pos.p_indent_col).value
        If (s <> "") Then
            Log_get_length_of_log_list_VA = Log_get_length_of_log_list_VA + 1
            GoTo restart_Log_get_length_of_log_list_VA
        End If
    'cleanup
        s = "Empty"
End Function
