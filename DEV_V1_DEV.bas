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
    'run if readonly
        If (ActiveWorkbook.ReadOnly = False) Then
            MsgBox ("--------------------------------------------------------------------------------------------" & Chr(10) & _
                "________Welcome to the Product Sales Pricing Tool- Data Editor________" & Chr(10) & _
                "                          Version: Alpha 1.1.8-2 front loader update" & Chr(10) & _
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
    'run regardless of read only
        MsgBox ("Need to remove 'unit list' marked range from the DTS refactor code as it has been removed from the data set")
        Boots_Report_v_Alpha.status
        dev_v1_dev.status
        DTS_V1_DEV.status
        matrix_V2.status
        SP_V1_DEV.status
        String_V1.status
End Sub

Public Sub On_Startup_V0_01(ByVal username As String, dont_report As Boolean)
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
    'make or check local locations made
        Call MsgBox("using dir create for testing", , "Using dev code from boots_main_v_alpha")
            Call alpha_MkDir("Pricetool-Alpha-omega", "C:\")
            Call alpha_MkDir("version-0", "C:\Pricetool-Alpha-omega\")
            Call alpha_MkDir("Users", "C:\Pricetool-Alpha-omega\version-0\")
    'create user log
        Call MsgBox("dev note add create right here", , "on startup vo_01 add creation of new user log")
        dev_v1_dev.check_user_in_v0_01 (dev_v1_dev.get_username)
    'log if checking for dev page
        If dont_report = False Then
            Call MsgBox("to remove and use new system", , "On_Startup_V0_01 using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "On startup started")
        End If
    'display startup messages
        dev_v1_dev.welcome
    'return flag
restart_on_startup:
    'set global varables
        Set proj_workbook = ActiveWorkbook
        Set Home_sheet = proj_workbook.ActiveSheet
    'check to see if DEV page exists
        'setup variables
            Set sheet_cursor = ActiveSheet
            Set cell_cursor = ActiveCell
            s = "empty"
            'set if reporting or not
                If dont_report = False Then
                    condition = dev_v1_dev.DEV_page_Exist
                Else
                    condition = True
                End If
        'make log post
            If dont_report = False Then
                Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
                'Call dev_v1_dev.log(dev_v1_dev.get_username, "condition check") 'check started and seeing values
            End If
        'logic to check for dev page exist
            If (condition = False) Then
            'dev page dont exist make new one
                'log
                If dont_report = False Then
                    Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "condition check passed as false")
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
                    'anti loop check to prevent infinite loop
                        anti_loop = anti_loop + 1   'anti loop iteration
                        'if you have not looped through this more than 6 times do if else do else
                        If (anti_loop < 6) Then
                        'log
                            If dont_report = False Then
                                Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
                                'Call dev_v1_dev.log(dev_v1_dev.get_username, "anti loop triggered")
                            End If
                        'goto flag
                            GoTo restart_on_startup
                        Else
                        'stuck in a loop
                            'log
                                If dont_report = False Then
                                    Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
                                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "ERROR CODE STUCK IN LOOP PROTECTION TRIGGERED")
                                End If
                            'report to user
                                MsgBox ("ERROR CODE STUCK IN LOOP PROTECTION TRIGGERED")
                                Stop
                        End If
            Else
            'dev page does exist
                'log
                    If dont_report = False Then
                        Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
                        'Call dev_v1_dev.log(dev_v1_dev.get_username, "condition check passed as true")
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
                Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
                'Call dev_v1_dev.log(dev_v1_dev.get_username, "dev page position checks")
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
            Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "On startup finished")
        End If
    'return user to sheet they where on
        Home_sheet.Activate
    'exit
        Exit Sub
    'error handling
Error_Startup_Tracked_Cell_filled:
        If dont_report = False Then
            Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "Error: On Startup unable to set table header as there is unexpected information in the field")
        End If
        Call MsgBox("Error: On Startup unable to set table header as there is unexpected information in the field", , "ERROR Startup unable to store table header")
        Stop
End Sub

Public Sub ON_Shutdown_V0_01()
    Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
    'Call dev_v1_dev.log(dev_v1_dev.get_username, "Start on shutdown")
    MsgBox ("On shutdown not setup")
    'check_user_out_v0_01 (get_username)
    Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
    'Call dev_v1_dev.log(dev_v1_dev.get_username, "Finish on shutdown")
End Sub

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
            Call MsgBox("Dev_page_exist using log replace with new report", , "Dev_page_exist using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "DEV_Page_Exist started")
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
                                    Call MsgBox("Dev_page_exist using log replace with new report", , "Dev_page_exist using log")
                                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "anti loop triggered")
                                End If
                            If (anti_loop < 6) Then
                                GoTo Restart_if_exist_check
                            Else
                                If dont_report = False Then
                                    Call MsgBox("Dev_page_exist using log replace with new report", , "Dev_page_exist using log")
                                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "ERROR CODE STUCK IN LOOP PROTECTION TRIGGERED")
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
        dev_v1_dev.On_Startup_V0_01 (True)
        Exit Function
        'log
page_exist_exit:
        If dont_report = False Then
            Call MsgBox("Dev_page_exist using log replace with new report", , "Dev_page_exist using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "DEV_Page_Exist finished")
        End If

End Function

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

Public Sub check_user_in_v0_01(ByVal user As String)

End Sub

Private Sub check_user_out_v0_01(ByVal user As String)
    
End Sub



















