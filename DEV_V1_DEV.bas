Attribute VB_Name = "DEV_V1_DEV"
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
                                                        'made for :DEV_V1
                                                                 ':NA
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Purpose Case
                    'This Module is built to handle all memory and backround operations for the price tool program
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'GLOBALS
    'DEFINE PRIVATE GLOBALS
        Private Enum POS_V0_01A
        'currently NOT functional as of (8/7/2020) checked by: (Zachary Daugherty)
            'Created By (Zachary Daugherty)(8/7/2020)
            'Purpose Case & notes:
                'dev index location
            'Library Refrences required
                'na
            'Modules Required
                'na
            'Inputs
                'Internal:
                    'workbook.object
                        'dev page
                'required:
                    'na
                'optional:
                    'na
            'returned outputs
                'returns index
        'code start
            'login table
                DEV_login_USER_ONLINE_row = 3 'set as global pos
                    DEV_login_USER_ONLINE_col = 2
                DEV_login_Signin_time_row = POS_V0_01A.DEV_login_USER_ONLINE_row
                    DEV_login_Signin_time_col = 3
                DEV_login_Marked_for_signout_row = POS_V0_01A.DEV_login_USER_ONLINE_row
                    DEV_login_Marked_for_signout_col = 4
                dev_login_bottom_row = 100
                    dev_login_bottom_col = POS_V0_01A.DEV_login_Marked_for_signout_col
                
            'log table
                DEV_log_log_row = 103
                    DEV_log_log_col = 2
                DEV_log_Timestamp_row = POS_V0_01A.DEV_log_log_row
                    DEV_log_Timestamp_col = 3
        'code end
        End Enum

        Public Function status()
            Call MsgBox("DEV_Vx Status:" & Chr(10) & _
            "------------------------------------------------------------" & Chr(10) & _
            "Public functions: " & Chr(10) & _
            "          welcome: Stable" & Chr(10) & _
            " get_username: Stable" & Chr(10) & _
            "     On_Startup: update" & Chr(10) & _
            "ON_Shutdown: update" & Chr(10) & _
            "                   Log: update" & Chr(10) & _
            "  check_user_in: depreciated" & Chr(10) & _
            Chr(10) & "Private functions:" & Chr(10) & _
            "DEV_page_Exist: stable" & Chr(10) & _
            "check_user_out: depreciated" & Chr(10) & _
            "", , "showing status for Dev_v1_dev")
        End Function
        
        Public Sub welcome()
            If (ActiveWorkbook.ReadOnly = False) Then
                MsgBox ("--------------------------------------------------------------------------------------------" & Chr(10) & _
                    "________Welcome to the Product Sales Pricing Tool- Data Editor________" & Chr(10) & _
                    "                          Version: Alpha 1.1.7 status update" & Chr(10) & _
                    "--------------------------------------------------------------------------------------------")
                    
                MsgBox ("--------------------------------------------------------------------------------------------" & Chr(10) & _
                    "DEV NOTES:" & Chr(10) & _
                    "      Makes sure read only opens are addressed as normal user" & Chr(10) & _
                    "            functionality will be a read only open" & Chr(10) & _
                    "      Add in dev a marker for keeping track of the sheets in the wb," & Chr(10) & _
                    "            this will allow marking for changes without having there be" & Chr(10) & _
                    "            issues with revisions" & Chr(10) & _
                    "      need to add admin user interface in the future." & Chr(10) & _
                    "      need to do an audit on readonly opens." & Chr(10) & _
                    "      need todo an audit of if a module is missing and the behavior that follows." & Chr(10) & _
                    "      need to update error logging tool to allow the exporting of the log right away.")
                    
            End If
        End Sub

        Public Sub On_Startup_V0_01(Optional dont_report As Boolean)
        'currently functional as of (8/7/2020) checked by: (Zachary daugherty)
            'Created By (Zachary Daugherty)(8/7/2020)
            'Purpose Case & notes:
                'runs and sets up any starting operations for the program
                'including
                    'setup of the dev page and if it does not exist makes a new one
            'Library Refrences required
                '(information not filled out)
            'Modules Required
                '(information not filled out)
            'Inputs
                'Internal:
                    'na
                'required:
                    'Na
                'optional:
                    'Na
            'returned outputs
                'Na
        'code start
            'define varables
                'system protection
                    Dim anti_loop As Long
                'memory
                    Dim Home_sheet As Worksheet     'on startup sheet treated as the home position to return to post check
                    Dim condition As Boolean        'store T/F
                    Dim s As String                 'stores string info
                    Dim i As Long                'long storage
                'cursor information
                    Dim proj_workbook As Workbook   'active workbook selection on open
                    Dim cell_cursor As Range        'cursor selection of current position of the cell
                    Dim cell_row As Long
                    Dim cell_col As Long
                    Dim sheet_cursor As Worksheet  'cursor selection of current position of the sheet
            'log
                If dont_report = False Then
                    Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "On startup started")
                End If
            'return
restart_on_startup:
            'set global varables
                Set proj_workbook = ActiveWorkbook
                Set Home_sheet = proj_workbook.ActiveSheet
            'check to see if DEV page exists
                'setup variables
                    Set sheet_cursor = ActiveSheet
                    Set cell_cursor = ActiveCell
                    s = "empty"
                If dont_report = False Then
                    condition = DEV_V1_DEV.DEV_page_Exist
                Else
                    condition = True
                End If
                If dont_report = False Then
                    Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "condition check")
                End If
                If (condition = False) Then
                    If dont_report = False Then
                        Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "condition check passed as false")
                    End If
                    'make devpage
                            'make screen updating off
                                Application.ScreenUpdating = False
                        proj_workbook.Sheets.Add    'add new sheet
                        Set sheet_cursor = ActiveSheet  'set cursor to active
                        sheet_cursor.Name = "DEV"   'rename active sheet
                        sheet_cursor.Visible = 2    'make dev hidden code lvl permission
                        Home_sheet.Activate 'return to home position
                            'make screen updating on
                                Application.ScreenUpdating = True
                        anti_loop = anti_loop + 1   'anti loop iteration
                        If (anti_loop < 6) Then
                            If dont_report = False Then
                                Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "anti loop triggered")
                            End If
                            GoTo restart_on_startup
                        Else
                            If dont_report = False Then
                                Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "ERROR CODE STUCK IN LOOP PROTECTION TRIGGERED")
                            End If
                            MsgBox ("ERROR CODE STUCK IN LOOP PROTECTION TRIGGERED")
                            Stop
                        End If

                Else
                    If dont_report = False Then
                        Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "condition check passed as true")
                    End If
                End If
                'cleanup
                    condition = False
                    anti_loop = 0
                    Set sheet_cursor = Nothing
                    Set cell_cursor = Nothing
                    s = "empty"
            'check for table positions
                    'check to make DEV is hidden
                        proj_workbook.Sheets("DEV").Visible = 1
                Set sheet_cursor = proj_workbook.Sheets("Dev")  'set cursor
                'check for table names and setup if not valid and if not throw error
                    If dont_report = False Then
                        Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "dev page position checks")
                    End If
                    Set cell_cursor = sheet_cursor.Cells(POS_V0_01A.DEV_login_USER_ONLINE_row, POS_V0_01A.DEV_login_USER_ONLINE_col)
                        If cell_cursor.value <> "Users Online" Then
                            If cell_cursor.value = "" Then
                                cell_cursor.value = "Users Online"
                            Else
                                GoTo Error_Startup_Tracked_Cell_filled
                            End If
                        End If
                    Set cell_cursor = sheet_cursor.Cells(POS_V0_01A.DEV_login_Signin_time_row, POS_V0_01A.DEV_login_Signin_time_col)
                        If cell_cursor.value <> "Sign in time" Then
                            If cell_cursor.value = "" Then
                                cell_cursor.value = "Sign in time"
                            Else
                                GoTo Error_Startup_Tracked_Cell_filled
                            End If
                        End If
                    Set cell_cursor = sheet_cursor.Cells(POS_V0_01A.DEV_login_Marked_for_signout_row, POS_V0_01A.DEV_login_Marked_for_signout_col)
                        If cell_cursor.value <> "Marked for Signout" Then
                            If cell_cursor.value = "" Then
                                cell_cursor.value = "Marked for Signout"
                            Else
                                GoTo Error_Startup_Tracked_Cell_filled
                            End If
                        End If
                    Set cell_cursor = sheet_cursor.Cells(POS_V0_01A.DEV_log_log_row, POS_V0_01A.DEV_log_log_col)
                        If cell_cursor.value <> "Action Log" Then
                            If cell_cursor.value = "" Then
                                cell_cursor.value = "Action Log"
                            Else
                                GoTo Error_Startup_Tracked_Cell_filled
                            End If
                        End If
                    'format login: set row pos
                        cell_row = POS_V0_01A.DEV_login_USER_ONLINE_row
                        cell_col = POS_V0_01A.DEV_login_USER_ONLINE_col
                    'format login
                        For i = 0 To POS_V0_01A.dev_login_bottom_row - 3
                            sheet_cursor.Cells(cell_row + i, cell_col).Interior.Color = 500
                            sheet_cursor.Cells(cell_row + i, cell_col + 1).Interior.Color = 500
                            sheet_cursor.Cells(cell_row + i, cell_col + 2).Interior.Color = 500
                        Next i
                        'cleanup
                            i = -1
                    'format log: set row pos
                        cell_row = POS_V0_01A.DEV_log_log_row
                        cell_col = POS_V0_01A.DEV_log_log_col
                    'format log

                        'if not formated format
                            If (sheet_cursor.Cells(cell_row, cell_col).Interior.Color <> 500) Then
                                'not formatted correctly so setup
                                    MsgBox ("Please wait Fixing Errors: Module: DEV_VX: Function: On_Startup is repairing the dev page this function will take a couple mins please wait...")
                                    For i = 0 To 1048575 - POS_V0_01A.DEV_log_log_row
                                        sheet_cursor.Cells(cell_row + i, cell_col).Interior.Color = 500
                                        sheet_cursor.Cells(cell_row + i, cell_col + 1).Interior.Color = 500
                                        sheet_cursor.Cells(cell_row + i, cell_col + 2).Interior.Color = 500
                                    Next i
                            End If
                'cleanup
                    Set sheet_cursor = Nothing
                    Set cell_cursor = Nothing
                    cell_row = -1
                    cell_col = -1
            'add creation to log
                If dont_report = False Then
                    Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "On startup finished")
                End If
            'return user to sheet they where on
                Home_sheet.Activate
            'exit
                Exit Sub
            'error handling
Error_Startup_Tracked_Cell_filled:
                If dont_report = False Then
                    Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "Error: On Startup unable to set table header as there is unexpected information in the field")
                End If
                Call MsgBox("Error: On Startup unable to set table header as there is unexpected information in the field", , "ERROR Startup unable to store table header")
                Stop
        End Sub
        
        Public Sub ON_Shutdown_V0_01()
            Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "Start on shutdown")
            MsgBox ("On shutdown not setup")
            'check_user_out_v0_01 (get_username)
            Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "Finish on shutdown")
        End Sub
        
    'DEFINE PRIVATE GLOBAL VARIABLES
        'NA
        
    'DEFINE PUBLIC GLOBAL VARIABLES
        'NA

'ROUTINES

    'CHECKER OPERATIONS
        Private Function DEV_page_Exist(Optional dont_report As Boolean) As Boolean
        'currently functional as of (8/7/2020) checked by: (Zachary Daugherty)
            'Created By (Zachary Daugherty)(8/7/2020)
            'Purpose Case & notes:
                'returns if the page exists
            'Library Refrences required
                'Workbooks.object
            'Modules Required
                'na
            'Inputs
                'Internal:
                    'workbooks.object
                'required:
                    'na
                'optional:
                    'Na
            'returned outputs
                'true: if the sheet exist
                'false: if not
        'code start
            'define varables
                'system protection
                    Dim anti_loop As Long
                'memory
                    Dim Home_sheet As Worksheet     'on startup sheet treated as the home position to return to post check
                    Dim arr() As String             'storage array for memory operations
                    Dim i As Long                'iterator and int storage
                    Dim s As String                 'string storage
                    Dim total_sheets_num As Long 'stores number in existance
                    
                'cursor information
                    Dim proj_workbook As Workbook   'active workbook selection on open
                    Dim sheet_cursor As Worksheet   'cursor selection of current position of the sheet
                    Dim cell_cursor As Range        'cursor selection of current position of the cell
            'log
                If dont_report = False Then
                    Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "DEV_Page_Exist started")
                End If
            'goto return
Restart_if_exist_check:
            'set varables
                Set proj_workbook = ActiveWorkbook
                Set sheet_cursor = ActiveSheet
                Set Home_sheet = proj_workbook.ActiveSheet
                Set cell_cursor = ActiveCell
            'get ammount of sheets that exist
                total_sheets_num = proj_workbook.Sheets.count
            'iterate through the sheets to see if dev exists
                For i = 1 To total_sheets_num
                    'get name of sheet
                        s = proj_workbook.Sheets(i).Name
                    'check if the sheet name is DEV
                        If (s = "DEV") Then
                            'if not in index position 1 move to 1
                                If (i > 1) Then
                                    proj_workbook.Sheets(i).Move _
                                        Before:=ActiveWorkbook.Sheets(1)
                                    anti_loop = antiloop + 1
                                        If dont_report = False Then
                                            Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "anti loop triggered")
                                        End If
                                    If (anti_loop < 6) Then
                                        GoTo Restart_if_exist_check
                                    Else
                                        If dont_report = False Then
                                            Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "ERROR CODE STUCK IN LOOP PROTECTION TRIGGERED")
                                        End If
                                        MsgBox ("ERROR CODE STUCK IN LOOP PROTECTION TRIGGERED")
                                        Stop
                                    End If
                                End If
                            'enviorment exists exit
                                DEV_page_Exist = True
                                GoTo page_exist_exit
                        Else
                            'fall through
                        End If
                Next i
            'sheet not found make sheet
                proj_workbook.Sheets.Add
                Set sheet_cursor = ActiveSheet
                sheet_cursor.Name = "DEV"
                DEV_page_Exist = True
                DEV_V1_DEV.On_Startup_V0_01 (True)
                Exit Function
                'log
page_exist_exit:
                If dont_report = False Then
                    Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "DEV_Page_Exist finished")
                End If

        End Function
    
    'GET/SET Operations
        'GET
            Public Function get_username(Optional dont_report As Boolean) As String
            'currently NOT functional as of (8/7/2020) checked by: (Zachary Daugherty)
                'Created By (Zachary Daugherty)(8/7/2020)
                'Purpose Case & notes:
                    'made to use functions to return the user name of the account in this session
                'Library Refrences required
                    'workbook.object
                'Modules Required
                    'Na
                'Inputs
                    'Internal:
                        'Na
                    'required:
                        'Na
                    'optional:
                        'Na
                'returned outputs
                    'username as string
            'code start
                get_username = (Environ$("Username"))
            'code end
            End Function
        'SET
            Public Sub log(ByVal user As String, ByVal message As String)
            'currently functional as of (8/19/20) checked by: (Zachary Daugherty)
                'Created By (Zachary Daugherty)(8/19/2020)
                'Purpose Case & notes:
                    'reports actions to the log for debugging reasons
                'Library Refrences required
                    'workbook.object
                'Modules Required
                    'na
                'Inputs
                    'Internal:
                        'time now
                    'required:
                        'username
                        'message to post to log
                    'optional:
                        'naa
                'returned outputs
                    'na
                'code start
                    'define varaibles
                        'cursor
                            Dim wb_cursor As Workbook
                            Dim sheet_cursor As Worksheet
                            Dim cell_cursor As Range
                            Dim home_pos As Worksheet
                        'storage
                            Dim s As String
                            Dim i As Long
                            Dim i_2 As Long
                            Dim anti_loop As Long
restart_log:
                    'setup var
                        Application.ScreenUpdating = False
                        Application.DisplayAlerts = False
                        Set wb_cursor = ActiveWorkbook
                        Set home_pos = ActiveSheet
                        On Error GoTo Fatal_Error_Log_cant_find_dev_page
                            Set sheet_cursor = wb_cursor.Sheets("DEV")
                                On Error GoTo 0
                        Set cell_cursor = sheet_cursor.Cells(POS_V0_01A.DEV_log_log_row, POS_V0_01A.DEV_log_log_col)
                    'report
                        wb_cursor.Sheets("DEV").Visible = 1
                            wb_cursor.Sheets("DEV").Activate
                            Cells(1048576, 2).Select
                            Set cell_cursor = ActiveCell
                            cell_cursor.End(xlUp).Offset(1, 0).Select
                        Set cell_cursor = ActiveCell
                        If cell_cursor.row = 1048570 Then   'clear log file
                            ActiveCell.Offset(-1, 0).Activate
                            Set cell_cursor = ActiveCell
                            i = cell_cursor.row
                            i_2 = cell_cursor.Column
                            cell_cursor.End(xlUp).Offset(1, 0).Activate
                            Set cell_cursor = Range(Cells(ActiveCell.row, ActiveCell.Column), Cells(i, i_2 + 1))
                            Cells(1, 1).Activate
                            cell_cursor.Select
                            cell_cursor.value = ""
                            anti_loop = anti_loop + 1
                            If anti_loop < 5 Then
                                GoTo restart_log
                            Else
                                MsgBox ("ERROR DEV Function LOG Failed to Escape log reset loop please check code.")
                                Stop
                            End If
                        End If
                        cell_cursor.value = "UserName: " & user & ": " & message
                        cell_cursor.Offset(0, 1).value = Now()
                        wb_cursor.Sheets("DEV").Visible = 2
                        Application.ScreenUpdating = True
                        Application.DisplayAlerts = True
                'code end
                    Exit Sub
                'error handler
Fatal_Error_Log_cant_find_dev_page:
                    DEV_V1_DEV.DEV_page_Exist (True)
                    Exit Sub
            End Sub
            
            Public Sub check_user_in_v0_01(ByVal user As String)
            'currently NOT functional as of (8/7/2020) checked by: (Zachary Daugherty)
                'Created By (Zachary Daugherty)(8/7/2020)
                'Purpose Case & notes:
                    'mark a user as in the doc & at what time
                'Library Refrences required
                    'na
                'Modules Required
                    'na
                'Inputs
                    'Internal:
                        'na
                    'required:
                        'user: as it finds who to mark what
                    'optional:
                        'na
                'returned outputs
                    'na
            'code start
                'define varables
                    'memory
                        Dim arr() As String 'made to store RAM
                        Dim s As String 'store string info
                        Dim i As Long    'store int and iterator
                        Dim count As Long 'counter
                    'cursor
                        Dim current_pos As Worksheet    'accounted as the starting position before this operation to return to it after done
                        Dim proj_workbook As Workbook   'selected as the local master
                        Dim cursor_sheet As Worksheet   'cursor location on what sheet specified
                        Dim cursor_row As Long       'self explained
                        Dim cursor_col As Long       'self explained
                'set varables
                    'log
                        Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "Start check_user_in")
                    'breakout
                    Set proj_workbook = ActiveWorkbook
                    Set current_pos = proj_workbook.ActiveSheet
                    Set cursor_sheet = current_pos
                    cursor_row = ActiveCell.row
                    cursor_col = ActiveCell.Column
                'get list
                        'make page updates false
                            Application.ScreenUpdating = False
                            Application.DisplayAlerts = False
                    'set DEV as cursor
                        Set cursor_sheet = proj_workbook.Sheets("DEV")
                            cursor_row = POS_V0_01A.DEV_login_USER_ONLINE_row + 1
                            cursor_col = POS_V0_01A.DEV_login_USER_ONLINE_col
                    'setup of arr
                        s = cursor_sheet.Cells(cursor_row, cursor_col).value
                        If (s <> "") Then
                            'stop multi login
                            count = 1
                            For i = 1 To 3
                                cursor_row = cursor_row + 1
                                s = cursor_sheet.Cells(cursor_row, cursor_col).value
                                If s = "" Then
                                    Exit For
                                Else
                                    i = 1
                                    count = count + 1
                                End If
                            Next i
                            ReDim arr(count + 1, 2)
                        Else
                            MsgBox ("devnote: ignore, need to setup multi login perms")
                            ReDim arr(1, 2)
                        End If
                        arr(0, 0) = "size of arr"
                        arr(0, 1) = count + 1
                    'cleanup
                        i = -1
                        count = -1
                'fill arr list
                    'set cell cursor to top of the list
                        cursor_row = POS_V0_01A.DEV_login_USER_ONLINE_row + 1
                        cursor_col = POS_V0_01A.DEV_login_USER_ONLINE_col
                    'fill
                        count = 0
                        For i = 1 To (CInt(arr(0, 1)) - 1)
                            arr(i, 0) = cursor_sheet.Cells(cursor_row, cursor_col).value
                            arr(i, 1) = cursor_sheet.Cells(cursor_row, cursor_col + 1).value
                            arr(i, 2) = cursor_sheet.Cells(cursor_row, cursor_col + 2).value
                            cursor_row = cursor_row + 1
                            count = count + 1
                        Next i
                    'check for dup
                        For i = 1 To (CInt(arr(0, 1)) - 1)
                            If (user = arr(i, 0)) Then
                                'dupe found
                                    'set pos for update
                                        cursor_row = POS_V0_01A.DEV_login_USER_ONLINE_row + i
                                        cursor_col = POS_V0_01A.DEV_login_USER_ONLINE_col
                                        cursor_sheet.Cells(cursor_row, cursor_col).value = user
                                        cursor_sheet.Cells(cursor_row, cursor_col + 1).value = Now()
                                        cursor_sheet.Cells(cursor_row, cursor_col + 2).value = DateAdd("d", 1, Now())
                                            GoTo check_user_in_cleanup:
                                'breakout
                            End If
                        Next i
                        'add new user to arr
                            arr(arr(0, 1), 0) = user
                            arr(arr(0, 1), 1) = Now()
                            arr(arr(0, 1), 2) = DateAdd("d", 1, CDate(arr(arr(0, 1), 1)))
                'reset cell position
                    cursor_row = POS_V0_01A.DEV_login_USER_ONLINE_row
                    cursor_col = POS_V0_01A.DEV_login_USER_ONLINE_col
                'add to list
                    cursor_row = cursor_row + CInt(arr(0, 1))
                    cursor_sheet.Cells(cursor_row, cursor_col).value = arr(CInt(arr(0, 1)), 0)
                    cursor_sheet.Cells(cursor_row, cursor_col + 1).value = arr(CInt(arr(0, 1)), 1)
                    cursor_sheet.Cells(cursor_row, cursor_col + 2).value = arr(CInt(arr(0, 1)), 2)
                'checkpoint
check_user_in_cleanup:
                'cleanup
                    cursor_row = 1
                    cursor_col = 1
                    current_pos.Activate
                    Application.ScreenUpdating = True
                    Application.DisplayAlerts = True
                    Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "Finished check_user_in")
            'code end
                Exit Sub
            'error handler
                'na
            End Sub
            
            Private Sub check_user_out_v0_01(ByVal user As String)
                Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "Start check_user_out")
                MsgBox ("dev note checkout not setup: Russ ignore this and just hit END")
                Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "Finished check_user_out")
            End Sub


Sub test()
    check_user_in_v0_01 (get_username)
End Sub



















