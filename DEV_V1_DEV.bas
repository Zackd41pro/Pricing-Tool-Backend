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

Private Enum dev_POS
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
    'safe to pull requests
        dev_p_safety_to_get_values_header_row = 3
            dev_p_safety_to_get_values_header_col = 2
                dev_p_safety_to_get_values_value_row = dev_POS.dev_p_safety_to_get_values_header_row + 1
                    dev_p_safety_to_get_values_value_col = dev_POS.dev_p_safety_to_get_values_header_col + 0
    'login table
        dev_p_login_user_online_time_header_row = 3
            dev_p_login_user_online_time_header_col = 4
                DEV_p_login_USER_ONLINE_header_row = dev_POS.dev_p_login_user_online_time_header_row + 0
                    DEV_p_login_USER_ONLINE_header_col = dev_POS.dev_p_login_user_online_time_header_col + 1
        
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
                "                          Version: Alpha 1.1.8-5 cleanup update" & Chr(10) & _
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
        'Boots_Report_v_Alpha.status
        DEV_V1_DEV.status
        matrix_V2.status
        SP_V1_DEV.status
        String_V1.status
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
    'create user log
        Call MsgBox("dev_v1.onstartup dev note add log right here")
    'log if checking for dev page
        If dont_report = False Then
            Call MsgBox("to remove and use new system", , "On_Startup_V0_01 using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "On startup started")
        End If
    'display startup messages
        DEV_V1_DEV.welcome
    'return flag
restart_on_startup:
    'set global varables
        Set proj_workbook = ActiveWorkbook
        Set Home_sheet = ActiveSheet
    'check to see if DEV page exists
        'setup variables
            Set sheet_cursor = ActiveSheet
            Set cell_cursor = ActiveCell
            s = "empty"
            'set if reporting or not
                MsgBox ("change the statement below to 'boots' not 'dev'")
                If dont_report = False Then
                    condition = DEV_V1_DEV.DEV_page_Exist
                Else
                    condition = True
                End If
        'make log post
            If dont_report = False Then
                Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
                'Call dev_v1_dev.log(dev_v1_dev.get_username, "condition check") 'check started and seeing values
            End If
        'if dev page dont exist then make page now
            If (condition = False) Then
            'dev page dont exist make new one
                'log
                If dont_report = False Then
                    Call MsgBox("to remove logging feature moving to more stable process", , "On_Startup_V0_01 using log")
                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "condition check passed as false")
                End If
                'make devpage
                    Call Boots_Main_V_alpha.make_sheet(proj_workbook, "DEV", 2, True)
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
            End If
        'cleanup
            condition = False
            anti_loop = 0
            Set sheet_cursor = Nothing
            Set cell_cursor = Nothing
            s = "empty"
    'format dev
        Call DEV_V1_DEV.format_Dev(proj_workbook)
        'cleanup
            'na
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

Public Sub check_user_in_v0_01(ByVal user As String)
    
End Sub

Public Sub check_user_out_v0_01(ByVal user As String)
    
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
        Set Home_sheet = ActiveSheet
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
                        MsgBox ("need to remove the reorder as causing crash")
                        Stop
'                        If (i > 1) Then
'                            proj_workbook.Sheets(i).Move _
'                                Before:=ActiveWorkbook.Sheets(1)
'                            anti_loop = anti_loop + 1
'                                If dont_report = False Then
'                                    Call MsgBox("Dev_page_exist using log replace with new report", , "Dev_page_exist using log")
'                                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "anti loop triggered")
'                                End If
'                            If (anti_loop < 6) Then
'                                GoTo Restart_if_exist_check
'                            Else
'                                If dont_report = False Then
'                                    Call MsgBox("Dev_page_exist using log replace with new report", , "Dev_page_exist using log")
'                                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "ERROR CODE STUCK IN LOOP PROTECTION TRIGGERED")
'                                End If
'                                MsgBox ("ERROR CODE STUCK IN LOOP PROTECTION TRIGGERED")
'                                Stop
'                            End If
'                        End If
                    'enviorment exists exit
                        DEV_page_Exist = True
                        GoTo page_exist_exit
                Else
                    'fall through
                End If
        Next i
    'sheet not found
        DEV_page_Exist = False
        Exit Function
        'log
page_exist_exit:
        If dont_report = False Then
            Call MsgBox("Dev_page_exist using log replace with new report", , "Dev_page_exist using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "DEV_Page_Exist finished")
        End If

End Function

Private Function format_Dev(ByVal wb As Workbook) As Boolean
    Call MsgBox("<private>dev_vx.format_dev needs green text added", , "<private>dev_vx.format_dev")
    'define variables
        Dim sht As Worksheet
        Dim home As Worksheet
    'setup variables
        Set home = ActiveSheet
        Set sht = wb.Sheets("DEV")
    'format
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sht.visible = xlSheetVisible

        'format safty check
            'header
                sht.Cells(dev_POS.dev_p_safety_to_get_values_header_row, dev_POS.dev_p_safety_to_get_values_header_col).value = "safety check"
        'format login
            'header time
                sht.Cells(dev_POS.dev_p_login_user_online_time_header_row, dev_POS.dev_p_login_user_online_time_header_col).value = "Sign-in Time"
            'header user
                sht.Cells(dev_POS.DEV_p_login_USER_ONLINE_header_row, dev_POS.DEV_p_login_USER_ONLINE_header_col).value = "Name of Accnt"
        'full sheet
            sht.Activate
            Cells.Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = -0.499984740745262
                    .PatternTintAndShade = 0
                End With
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
                Cells.EntireColumn.AutoFit
                Cells.Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Range("A1").Select
        'cleanup
            sht.visible = xlSheetVeryHidden
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
            home.Activate
End Function

















