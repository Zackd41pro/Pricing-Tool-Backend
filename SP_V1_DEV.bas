Attribute VB_Name = "SP_V1_DEV"
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
                                                        'made for :SP_V1
                                                                 ':DEV_V1_DEV
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Purpose Case
                    'This Module is built to handle all referances to the Price Tool STEEL PRESET database For proper Referenceing and Updating
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Enum SP_POS
        'currently functional as of (8/7/2020) checked by: (zdaugherty)
            'Created By (Zachary Daugherty)(8/6/2020)
            'Purpose Case & notes:
                'POS Enum is to be called to act as a check condition to verify that the code and the sheet agrees on
                    'the locational position of where things are on the sheet.
            'Library Refrences required
                'Na
            'Modules Required
                '(information not filled out)
            'Inputs
                'Internal:
                    'Na
                'required:
                    'Na
                'optional:
                    'Na
            'returned outputs
                'returns indexed number
        'code start
            'other table information
                SP_I_Const_Plate_row = 2
                    SP_I_Const_Plate_col = 3
                SP_I_Const_Structural_row = 2
                    SP_I_Const_Structural_col = 4
            'number of entry fields watching
                SP_Q_A_Number_tracked_cells = 4
                SP_Q_B_Number_tracked_cells = 4
                SP_Q_G_Number_tracked_cells = 2
                SP_Q_Number_total_A_Tracked_Cells = SP_POS.SP_Q_A_Number_tracked_cells + SP_POS.SP_Q_G_Number_tracked_cells
                SP_Q_Number_total_B_Tracked_Cells = SP_POS.SP_Q_B_Number_tracked_cells + SP_POS.SP_Q_G_Number_tracked_cells
                SP_Q_Number_total_tracked_cells = SP_POS.SP_Q_A_Number_tracked_cells + SP_POS.SP_Q_B_Number_tracked_cells + SP_POS.SP_Q_G_Number_tracked_cells
            'array of positions
                'A
                    SP_I_A_Prefix_row = 4
                        SP_I_A_Prefix_col = 1
                    SP_I_A_Description_row = SP_POS.SP_I_A_Prefix_row
                        SP_I_A_Description_col = 2
                    SP_I_A_Cost_per_lb_row = SP_POS.SP_I_A_Prefix_row
                        SP_I_A_Cost_per_lb_col = 3
                    SP_I_A_Cost_per_lb_Wdrop_row = SP_POS.SP_I_A_Prefix_row
                        SP_I_A_Cost_per_lb_Wdrop_col = 4
                'B
                    SP_I_B_Prefix_row = 4
                        SP_I_B_Prefix_col = 7
                    SP_I_B_Description_row = SP_POS.SP_I_B_Prefix_row
                        SP_I_B_Description_col = 8
                    SP_I_B_Cost_per_lb_row = SP_POS.SP_I_B_Prefix_row
                        SP_I_B_Cost_per_lb_col = 9
                    SP_I_B_Cost_per_lb_Wdrop_row = SP_POS.SP_I_B_Prefix_row
                        SP_I_B_Cost_per_lb_Wdrop_col = 10
        End Enum
        
        Public Function status()
            Call MsgBox("SP_Vx Status:" & Chr(10) & _
            "------------------------------------------------------------" & Chr(10) & _
            "Public functions: " & Chr(10) & _
            " Check_SP_A_Table: Stable" & Chr(10) & _
            " Check_SP_B_Table: Stable" & Chr(10) & _
            "    get_sheet_name: Stable" & Chr(10) & _
            "              get_size_A: Stable" & Chr(10) & _
            "              get_size_B_V1: Stable" & Chr(10) & _
            Chr(10) & "Private functions:" & Chr(10) & _
            "na" & Chr(10) & _
            "", , "showing status for SP_v1_dev")
        End Function
        
    Public Function get_sheet_name(Optional more_instructions As Variant) As Variant
        'currently functional as of (9/3/2020) checked by: (zachary daugherty)
        'Created By (Zachary Daugherty)(9/3/2020)
        'Purpose Case & notes:
            'returns the sheet name associated with this module
        'Library Refrences required
            'workbook.object
        'Modules Required
            'string_v1
        'inputs
            'Internal:
                'na
            'required:
                'na
            'optional:
                'na
        'returned outputs
            'returns:
                'name of sheet as string
'            'check for instructions
'                If (dont_show_instructions = False) Then
'                    MsgBox ("Showning instructions for: SP_Vx:get_expected_sheet_name.__ this function is designed to return the expected name of the sheet associated with this module")
'                    Stop
'                    Exit Function
'                End If
        'check for log reporting
            If (more_instructions = "Log_Report") Then
                get_size_B_V1 = "get_size_B_V1 - Public - Stable with logs 11/12/2020 - Missing help file"
                Exit Function
            End If
        'code start
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "sp_v1_dev.get_sheet_name starting...")
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                get_sheet_name = "Steel Presets"
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "sp_v1_dev.get_sheet_name Finish...")
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            Exit Function
        End Function
        
        Public Function Check_SP_A_Table_V1(Optional more_instructions As String) As Variant
        'currently functional as of (8/11/2020) checked by: (zachary daugherty)
            'Created By (Zachary Daugherty)(8/11/2020)
            'Purpose Case & notes:
                'Check_SP_A_TABLE is a function that SHOULD be called at the start of any GET OR SET operation on the SP_A_table.
                    'As when any operation is done to the data table it should first be verifyed that the current version of the
                    'program and the addressed cell locations agree on their position.
                'This is done by Calling POS ENUM and compairing the positional data indexed per what is expected.
            'Library Refrences required
                'Na
            'Modules Required
                'DEV_V1_DEV
            'inputs
                'Internal:
                    'SP_VX.POS
                'required:
                    'na
                'optional:
                    'na
            'returned outputs
                'returns:
                    'true: if all positions match up
                    'false: if any positions do not match up
        'check for log reporting
            If (more_instructions = "Log_Report") Then
                Check_SP_A_Table_V1 = "Check_SP_A_Table_V1 - Public - Stable - no log reporting & no help file"
                Exit Function
            End If
        'code start
            'devnote
                Boots_Report_v_Alpha.Push_notification_message ("SP_V1_dev.Check_SP_A_Table_V1 has old error codes and must be refactored...")
                Boots_Report_v_Alpha.Push_notification_message ("SP_V1_dev.Check_SP_A_Table_V1 has no log reporting & no help file...")
                Boots_Report_v_Alpha.Push_notification_message ("DTS_V2A.CHECK_DTS_TABLE_V0_01A: Needs to have an error code added for the matrix array size as it can fail if the sizes are not set right see 'HP_V3_stable.DO_Check_HP_A_Table_V1' for an example....")
            'define varables
                'memory
                    Dim arr() As String             'designed as ram storage
                    Dim condition As Boolean        'store T/F
                    Dim i As Long                'iterator and int storage
                    Dim s As String                 'string storage
                'cursor
                    Dim proj_wb As Workbook         'set local workbook
                    Dim cursor_sheet As Worksheet   'sheet the cursor is on
                    Dim cursor_row As Long       'self explains
                    Dim cursor_col As Long       'self explains
                'ref
                    Dim ref_rng As Range            'reference range in question
            'setup variables
                'log
                    Call MsgBox("check sp table a using log replace", , "check sp table a using log")
                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "Check_SP_A_Table Started")
                'breakout
                Set proj_wb = ActiveWorkbook
                On Error GoTo FATAL_ERROR_CHECK_SP_A_SET_SP_ENV_For_A 'set error handler
                    Set cursor_sheet = proj_wb.Sheets("STEEL PRESETS")
                On Error GoTo 0 'set error handler back to norm
                cursor_row = 1
                cursor_col = 1
                
            'setup arr
                'redefine size of the arr
                    ReDim arr(1 To SP_POS.SP_Q_Number_total_A_Tracked_Cells, 1 To 5) 'see line below for definitions
                        'arr memory assignments
                            '(<specific index>,<1 to 5>)
                            '(<specific index>,<1:row of enum>)
                            '(<specific index>,<2:col of enum>)
                            '(<specific index>,<3:row of range>)
                            '(<specific index>,<4:col of range>)
                            '(<specific index>,<5: conditional if match>)
                'fill arr
                    'Collect information
                        i = 0
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        'NOTICE CODE IN THIS BLOCK IS STD AND THE OPERATIONS ARE THE SAME SO DEV NOTES ON THE FIRST FOLLOW THRU
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        
                        'compair <SP_GENERAL_PREFIX> expected location
                            s = "SP_GENERAL_PREFIX"   'expected range name for search
                            i = i + 1               'iterate arr position from x to x + 1 in the array
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_A 'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                            Set ref_rng = Range(s)  'set range
                            On Error GoTo 0 'reset error handler
                                arr(i, 1) = CStr(ref_rng.row)       'get range row pos
                                arr(i, 2) = CStr(ref_rng.Column)    'get range col pos
                                arr(i, 3) = SP_POS.SP_I_A_Prefix_row   'get enum row pos
                                arr(i, 4) = SP_POS.SP_I_A_Prefix_col   'get enum col pos
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                                    arr(i, 5) = s & ": " & True 'if true report text
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                                    condition = True    'if true at the end of the block throw error as there is a miss match
                                End If
                        'compair <SP_GENERAL_DESCRIPTION> expected location
                            s = "SP_GENERAL_DESCRIPTION"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_A
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_A_Description_row
                                arr(i, 4) = SP_POS.SP_I_A_Description_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <SP_GENERAL_COST_PER_LB> expected location
                            s = "SP_GENERAL_COST_PER_LB"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_A
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_A_Cost_per_lb_row
                                arr(i, 4) = SP_POS.SP_I_A_Cost_per_lb_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <SP_GENERAL_COST_PER_LB_W_DROP> expected location
                            s = "SP_GENERAL_COST_PER_LB_W_DROP"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_A
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_A_Cost_per_lb_Wdrop_row
                                arr(i, 4) = SP_POS.SP_I_A_Cost_per_lb_Wdrop_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <SP_GLOBAL_PLATE> expected location
                            s = "SP_GLOBAL_PLATE"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_A
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_Const_Plate_row
                                arr(i, 4) = SP_POS.SP_I_Const_Plate_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <SP_GLOBAL_STRUCTURAL> expected location
                            s = "SP_GLOBAL_STRUCTURAL"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_A
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_Const_Structural_row
                                arr(i, 4) = SP_POS.SP_I_Const_Structural_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        '__________________________________________END of CODE BLOCK___________________________________________
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        'cleanup
                            i = 0
                            s = "Empty"
            'compile report
                'check to see if failure condition is met
                    
                    If (condition = True) Then
                        GoTo ERROR_CHECK_sp_FAILED_POS_CHECK_For_A
                    End If
                'return true
                    Call MsgBox("check sp a table using log replace", , "check sp a table using log")
                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "Check_SP_A_Table Finished")
                    Check_SP_A_Table_V1 = True   'passed all checks
                    Exit Function
        'code end
        'error handle
ERROR_FATAL_check_sp_range_error_For_A:
            Call MsgBox("check sp a table using log replace", , "check sp a table using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "FATAL ERROR: MODULE:(SP_VX)FUNCTION:(CHECK_SP_TABLE_A) UNABLE TO LOCATE THE SPECIFIED RANGE:<" & s & "> please check the name mannager for errors. fix and then re-run")
            Call MsgBox("FATAL ERROR: MODULE:(SP_VX)FUNCTION:(CHECK_SP_TABLE_A) UNABLE TO LOCATE THE SPECIFIED RANGE:<" & s & "> please check the name mannager for errors. fix and then re-run", , "Fatal error")
            Stop
FATAL_ERROR_CHECK_SP_A_SET_SP_ENV_For_A:
            Call MsgBox("check sp a table using log replace", , "check sp a table using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "FATAL_ERROR: MODULE: (SP_VX)FUNCTION: (CHECK_SP_TABLE_A) UNABLE TO FIND OR SET SHEET STEEL PRESETS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.")
            Call MsgBox("FATAL_ERROR: MODULE: (SP_VX)FUNCTION: (CHECK_SP_TABLE_A) UNABLE TO FIND OR SET SHEET STEEL PRESETS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.", , "FATAL ERROR: SET SP SHEET ENV")
            Stop
ERROR_CHECK_sp_FAILED_POS_CHECK_For_A:
            Call MsgBox("check sp a table using log replace", , "check sp a table using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "ERROR: MODULE: (SP_VX)FUNCTION: (CHECK_SP_TABLE_A) FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: " & vbCrLf & vbCrLf & arr(1, 5) & vbCrLf & arr(2, 5) & vbCrLf & arr(3, 5) & vbCrLf & arr(4, 5) & vbCrLf & arr(5, 5) & vbCrLf & arr(6, 5) & vbCrLf & arr(7, 5) & vbCrLf & arr(8, 5) & vbCrLf & arr(9, 5) & vbCrLf & arr(10, 5) & vbCrLf & arr(11, 5) & vbCrLf & arr(12, 5) & vbCrLf & arr(13, 5) & vbCrLf & arr(14, 5) & vbCrLf & arr(15, 5))
            Call MsgBox("ERROR: MODULE: (SP_VX)FUNCTION: (CHECK_SP_TABLE_A) FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: " & vbCrLf & vbCrLf & arr(1, 5) & vbCrLf & arr(2, 5) & vbCrLf & arr(3, 5) & vbCrLf & arr(4, 5) & vbCrLf & arr(5, 5) & vbCrLf & arr(6, 5) & vbCrLf & arr(7, 5) & vbCrLf & arr(8, 5) & vbCrLf & arr(9, 5) & vbCrLf & arr(10, 5) & vbCrLf & arr(11, 5) & vbCrLf & arr(12, 5) & vbCrLf & arr(13, 5) & vbCrLf & arr(14, 5) & vbCrLf & arr(15, 5))
            Stop
        End Function
    
Public Function Check_SP_B_Table_V0_01A() As Boolean
        'currently functional as of (8/18/2020) checked by: (zachary daugherty)
            'Created By (Zachary Daugherty)(8/18/2020)
            'Purpose Case & notes:
                'Check_SP_B_TABLE is a function that SHOULD be called at the start of any GET OR SET operation on the SP_B_table.
                    'As when any operation is done to the data table it should first be verifyed that the current version of the
                    'program and the addressed cell locations agree on their position.
                'This is done by Calling POS ENUM and compairing the positional data indexed per what is expected.
            'Library Refrences required
                'Na
            'Modules Required
                'DEV_V1_DEV
            'inputs
                'Internal:
                    'SP_VX.POS
                'required:
                    'na
                'optional:
                    'na
            'returned outputs
                'returns:
                    'true: if all positions match up
                    'false: if any positions do not match up
        'devnote
            Boots_Report_v_Alpha.Push_notification_message ("SP_V1_dev.Check_SP_B_Table_V0_01A has old error codes and must be refactored...")
            Boots_Report_v_Alpha.Push_notification_message ("SP_V1_dev.Check_SP_B_Table_V0_01A has no log reporting & no help file...")
            Boots_Report_v_Alpha.Push_notification_message ("DTS_V2A.CHECK_DTS_TABLE_V0_01A: Needs to have an error code added for the matrix array size as it can fail if the sizes are not set right see 'HP_V3_stable.DO_Check_HP_A_Table_V1' for an example....")
        'code start
            'define varables
                'memory
                    Dim arr() As String             'designed as ram storage
                    Dim condition As Boolean        'store T/F
                    Dim i As Long                'iterator and int storage
                    Dim s As String                 'string storage
                'cursor
                    Dim proj_wb As Workbook         'set local workbook
                    Dim cursor_sheet As Worksheet   'sheet the cursor is on
                    Dim cursor_row As Long       'self explains
                    Dim cursor_col As Long       'self explains
                'ref
                    Dim ref_rng As Range            'reference range in question
            'setup variables
                'log
                    Call MsgBox("check sp b table using log replace", , "check sp b table using log")
                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "Check_sp_B_table Start")
                'breakout
                Set proj_wb = ActiveWorkbook
                On Error GoTo FATAL_ERROR_CHECK_SP_A_SET_SP_ENV_For_B 'set error handler
                    Set cursor_sheet = proj_wb.Sheets("STEEL PRESETS")
                On Error GoTo 0 'set error handler back to norm
                cursor_row = 1
                cursor_col = 1
            'setup arr
                'redefine size of the arr
                    ReDim arr(1 To SP_POS.SP_Q_Number_total_B_Tracked_Cells, 1 To 5) 'see line below for definitions
                        'arr memory assignments
                            '(<specific index>,<1 to 5>)
                            '(<specific index>,<1:row of enum>)
                            '(<specific index>,<2:col of enum>)
                            '(<specific index>,<3:row of range>)
                            '(<specific index>,<4:col of range>)
                            '(<specific index>,<5: conditional if match>)
                'fill arr
                    'Collect information
                        i = 0
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        'NOTICE CODE IN THIS BLOCK IS STD AND THE OPERATIONS ARE THE SAME SO DEV NOTES ON THE FIRST FOLLOW THRU
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        
                        'compair <SP_PROPRIETARY_PREFIX> expected location
                            s = "SP_PROPRIETARY_PREFIX"   'expected range name for search
                            i = i + 1               'iterate arr position from x to x + 1 in the array
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_B 'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                            Set ref_rng = Range(s)  'set range
                            On Error GoTo 0 'reset error handler
                                arr(i, 1) = CStr(ref_rng.row)       'get range row pos
                                arr(i, 2) = CStr(ref_rng.Column)    'get range col pos
                                arr(i, 3) = SP_POS.SP_I_B_Prefix_row   'get enum row pos
                                arr(i, 4) = SP_POS.SP_I_B_Prefix_col   'get enum col pos
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                                    arr(i, 5) = s & ": " & True 'if true report text
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                                    condition = True    'if true at the end of the block throw error as there is a miss match
                                End If
                        'compair <SP_PROPRIETARY_DESCRIPTION> expected location
                            s = "SP_PROPRIETARY_DESCRIPTION"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_B
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_B_Description_row
                                arr(i, 4) = SP_POS.SP_I_B_Description_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <SP_PROPRIETARY_COST_PER_LB> expected location
                            s = "SP_PROPRIETARY_COST_PER_LB"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_B
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_B_Cost_per_lb_row
                                arr(i, 4) = SP_POS.SP_I_B_Cost_per_lb_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <SP_PROPRIETARY_COST_PER_LB_W_DROP> expected location
                            s = "SP_PROPRIETARY_COST_PER_LB_W_DROP"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_B
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_B_Cost_per_lb_Wdrop_row
                                arr(i, 4) = SP_POS.SP_I_B_Cost_per_lb_Wdrop_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <SP_GLOBAL_PLATE> expected location
                            s = "SP_GLOBAL_PLATE"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_B
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_Const_Plate_row
                                arr(i, 4) = SP_POS.SP_I_Const_Plate_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <SP_GLOBAL_STRUCTURAL> expected location
                            s = "SP_GLOBAL_STRUCTURAL"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_sp_range_error_For_B
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = SP_POS.SP_I_Const_Structural_row
                                arr(i, 4) = SP_POS.SP_I_Const_Structural_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        '__________________________________________END of CODE BLOCK___________________________________________
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        'cleanup
                            i = 0
                            s = "Empty"
            'compile report
                'check to see if failure condition is met
                    If (condition = True) Then
                        GoTo ERROR_CHECK_sp_FAILED_POS_CHECK_For_B
                    End If
                'return true
                    Call MsgBox("check sp b table using log replace", , "check sp b table using log")
                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "check_sp_B_table finished")
                    Check_SP_B_Table_V0_01A = True   'passed all checks
                    Exit Function
        'code end
        'error handle
ERROR_FATAL_check_sp_range_error_For_B:
            Call MsgBox("check sp b table using log replace", , "check sp b table using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "FATAL ERROR: MODULE:(SP_VX)FUNCTION:(CHECK_SP_TABLE_A) UNABLE TO LOCATE THE SPECIFIED RANGE:<" & s & ">")
            Call MsgBox("FATAL ERROR: MODULE:(SP_VX)FUNCTION:(CHECK_SP_TABLE_A) UNABLE TO LOCATE THE SPECIFIED RANGE:<" & s & "> please check the name mannager for errors. fix and then re-run", , "Fatal error")
            Stop
FATAL_ERROR_CHECK_SP_A_SET_SP_ENV_For_B:
            Call MsgBox("check sp b table using log replace", , "check sp b table using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "FATAL_ERROR: MODULE: (SP_VX)FUNCTION: (CHECK_SP_TABLE_A) UNABLE TO FIND OR SET SHEET STEEL PRESETS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.")
            Call MsgBox("FATAL_ERROR: MODULE: (SP_VX)FUNCTION: (CHECK_SP_TABLE_A) UNABLE TO FIND OR SET SHEET STEEL PRESETS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.", , "FATAL ERROR: SET SP SHEET ENV")
            Stop
ERROR_CHECK_sp_FAILED_POS_CHECK_For_B:
            Call MsgBox("check sp b table using log replace", , "check sp b table using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "ERROR: MODULE: (SP_VX)FUNCTION: (CHECK_SP_TABLE_A) FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: " & vbCrLf & vbCrLf & arr(1, 5) & vbCrLf & arr(2, 5) & vbCrLf & arr(3, 5) & vbCrLf & arr(4, 5) & vbCrLf & arr(5, 5) & vbCrLf & arr(6, 5) & vbCrLf & arr(7, 5) & vbCrLf & arr(8, 5) & vbCrLf & arr(9, 5) & vbCrLf & arr(10, 5) & vbCrLf & arr(11, 5) & vbCrLf & arr(12, 5) & vbCrLf & arr(13, 5) & vbCrLf & arr(14, 5) & vbCrLf & arr(15, 5))
            Call MsgBox("ERROR: MODULE: (SP_VX)FUNCTION: (CHECK_SP_TABLE_A) FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: " & vbCrLf & vbCrLf & arr(1, 5) & vbCrLf & arr(2, 5) & vbCrLf & arr(3, 5) & vbCrLf & arr(4, 5) & vbCrLf & arr(5, 5) & vbCrLf & arr(6, 5) & vbCrLf & arr(7, 5) & vbCrLf & arr(8, 5) & vbCrLf & arr(9, 5) & vbCrLf & arr(10, 5) & vbCrLf & arr(11, 5) & vbCrLf & arr(12, 5) & vbCrLf & arr(13, 5) & vbCrLf & arr(14, 5) & vbCrLf & arr(15, 5))
            Stop
        End Function
    

        Public Function get_size_A(Optional more_instructions As String) As Variant
        'currently functional as of (9/2/2020) checked by: (Zachary Daugherty)
            'Created By (Zachary daugherty)(8/28/2020)
            'Purpose Case & notes:
                'returns the total number of rows in Table A ('Steel Presets General')
            'Library Refrences required
                'workbook.object
            'Modules Required
                'na
            'Inputs
                'Internal:
                    'na
                'required:
                    'na
                'optional:
                    'na
            'returned outputs
                'returns the size as a length
            'check for log reporting
                If (more_instructions = "Log_Report") Then
                    get_size_A = "get_size_A - Public - Stable"
                    Exit Function
                End If
            'code start
                Call Boots_Report_v_Alpha.Log_Push(text, "SP_v1_dev.get_size_A Start...")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'define variables
                    Call Boots_Report_v_Alpha.Log_Push(text, "Setup Variables...")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    'positional
                        Dim wb As Workbook
                        Dim home_pos As Worksheet
                        Dim current_sht As Worksheet
                        Dim row As Long
                        Dim col As Long
                    'memory
                        'na
                    'const
                        Dim dist_to_goalpost As Long
                    'containers
                        Dim s As String
                    'setup variables
                        Set wb = ActiveWorkbook
                        Set home_pos = ActiveSheet
                        On Error GoTo SP_get_size_A_cant_find_SP_SHEET
                            Set current_sht = wb.Sheets("STEEL PRESETS")
                        On Error GoTo 0
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
            'Seting start location
                Call Boots_Report_v_Alpha.Log_Push(text, "Setting Start location...")
                row = SP_POS.SP_I_A_Prefix_row
                col = SP_POS.SP_I_A_Prefix_col
                s = current_sht.Cells(row, col).value
            'get lenght to bottom
            Call Boots_Report_v_Alpha.Log_Push(text, "Browsing for the Goalpost position...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                On Error GoTo sp_A_cant_find_goalpost
                    dist_to_goalpost = Range("SP_GENERAL_GOALPOST").row - row
                On Error GoTo 0
                'get size
                    Call Boots_Report_v_Alpha.Log_Push(text, "Returning result...")
                    get_size_A = dist_to_goalpost
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'code end
                'cleanup
                            Call Boots_Report_v_Alpha.Log_Push(text, "SP_V1_dev.get_size_A Finished...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                'return
                    Exit Function
            'error handling
SP_get_size_A_cant_find_SP_SHEET:
                'SP_get_size_A_cant_find_SP_SHEET:
                    'set error report
                        Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                        Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: SP_V1_DEV.get_size_A was unable to find the sheet named 'Steel Presets'")
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'generate and store required information
                        Call Boots_Report_v_Alpha.Log_Push(text, "generating sheet list...")
                            Call Boots_Main_V_alpha.get_sheet_list
                        'push generated information
                            Call Boots_Report_v_Alpha.Log_Push(table_open)
                                Call Boots_Report_v_Alpha.Log_Push(text, "Displaying All existing sheets...")
                                For z = 1 To (wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).End(xlDown).row - wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).row)
                                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Sheet List: " & wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row + z, boots_pos.p_sheet_name_col).value & _
                                        " ='visible stat': " & wb.Sheets("boots").Cells(boots_pos.p_sheet_visible_status_row + z, boots_pos.p_sheet_visible_status_col).value)
                                Next z
                            Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'showing important stored vars
                        Call Boots_Report_v_Alpha.Log_Push(text, "Showing important Variables...")
                            Call Boots_Report_v_Alpha.Log_Push(table_open)
                                'wb
                                    If wb Is Nothing Then
                                        Call Boots_Report_v_Alpha.Log_Push(Variable, "wb: 'NOTHING' as workbook")
                                    Else
                                        Call Boots_Report_v_Alpha.Log_Push(Variable, "wb: '" & wb.path & "\" & wb.Name & "' as workbook")
                                    End If
                                'home_pos
                                    If home_pos Is Nothing Then
                                        Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: 'NOTHING' as worksheet")
                                    Else
                                        Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: '" & home_pos.Parent.path & "\" & home_pos.Parent.Name & " == " & home_pos.index & ": " & home_pos.Name & "' as worksheet")
                                    End If
                                'current_sht
                                    If current_sht Is Nothing Then
                                        Call Boots_Report_v_Alpha.Log_Push(Variable, "current_sht: 'NOTHING' as worksheet")
                                    Else
                                        Call Boots_Report_v_Alpha.Log_Push(Variable, "current_sht: '" & current_sht.Parent.path & "\" & current_sht.Parent.Name & " == " & current_sht.index & ": " & current_sht.Name & "' as worksheet")
                                    End If
                                'activeworkbook
                                    If ActiveWorkbook Is Nothing Then
                                        Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                                    Else
                                        Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                                    End If
                                'more_instructions
                                    Call Boots_Report_v_Alpha.Log_Push(Variable, "more_instructions: '" & more_instructions & "' as string")
                                
                            Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'close error code
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'exit
                        'indent out
                            For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                            Next z
                        'call end statement
                            Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                            End
sp_A_cant_find_goalpost:
                'sp_A_cant_find_goalpost
                    'set error report
                        Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                        Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: SP_VX: FUNCTION  GET_SIZE_A: was unable to find the range named 'SP_GENERAL_GOALPOST'.")
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'generate and store required information
                        Call Boots_Report_v_Alpha.Log_Push(text, "generating sheet list...")
                            Call Boots_Main_V_alpha.get_sheet_list
                        'push generated information
                            Call Boots_Report_v_Alpha.Log_Push(table_open)
                                Call Boots_Report_v_Alpha.Log_Push(text, "Displaying All existing sheets...")
                                For z = 1 To (wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).End(xlDown).row - wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).row)
                                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Sheet List: " & wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row + z, boots_pos.p_sheet_name_col).value & _
                                        " ='visible stat': " & wb.Sheets("boots").Cells(boots_pos.p_sheet_visible_status_row + z, boots_pos.p_sheet_visible_status_col).value)
                                Next z
                            Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'give addresses
                        Call Boots_Report_v_Alpha.Log_Push(text, "generating Address list...")
                        Call Boots_Report_v_Alpha.Log_Push(table_open)
                        'home_pos
                            If home_pos Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: 'NOTHING' as worksheet")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: '" & home_pos.Parent.path & "\" & home_pos.Parent.Name & " == " & home_pos.index & ": " & home_pos.Name & "' as worksheet")
                            End If
                        'current_sht
                            If current_sht Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "current_sht: 'NOTHING' as worksheet")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "current_sht: '" & current_sht.Parent.path & "\" & current_sht.Parent.Name & " == " & current_sht.index & ": " & current_sht.Name & "' as worksheet")
                            End If
                        'wb
                            If wb Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "wb: 'NOTHING' as workbook")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "wb: '" & wb.path & "\" & wb.Name & "' as workbook")
                            End If
                        'activeworkbook
                            If ActiveWorkbook Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                            End If
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'TABLE CLOSE
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                        'end procedure
                            For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                            Next z
                        'call end statement
                            Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                            End
        End Function
        
        Public Function get_size_B_V1(Optional more_instructions As String) As Variant
        'currently functional as of (9/2/2020) checked by: (Zachary Daugherty)
            'Created By (Zachary daugherty)(8/28/2020)
            'Purpose Case & notes:
                'returns the total number of rows in Table B ('STEEL PRESETS PROPRIETARY')
            'Library Refrences required
                'workbook.object
            'Modules Required
                'na
            'Inputs
                'Internal:
                    'na
                'required:
                    'na
                'optional:
                    'na
            'returned outputs
                'returns the size as a length
            'check for log reporting
                If (more_instructions = "Log_Report") Then
                    get_size_B_V1 = "get_size_B_V1 - Public - Stable"
                    Exit Function
                End If
            'code start
                Call Boots_Report_v_Alpha.Log_Push(text, "Sp_v1_dev.get_size_B_V1 Start...")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'define variables
                    Call Boots_Report_v_Alpha.Log_Push(text, "Setup variables...")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    'positional
                        Dim wb As Workbook
                        Dim home_pos As Worksheet
                        Dim current_sht As Worksheet
                        Dim row As Long
                        Dim col As Long
                    'memory
                        'na
                    'const
                        Dim dist_to_goalpost As Long
                    'containers
                        Dim s As String
                    'setup variables
                        Set wb = ActiveWorkbook
                        Set home_pos = ActiveSheet
                        On Error GoTo SP_get_size_B_V1_cant_find_SP_SHEET
                            Set current_sht = wb.Sheets("STEEL PRESETS")
                        On Error GoTo 0
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                'move to start location
                    Call Boots_Report_v_Alpha.Log_Push(text, "moving to start location...")
                    row = SP_POS.SP_I_B_Prefix_row
                    col = SP_POS.SP_I_B_Prefix_col
                    s = current_sht.Cells(row, col).value
                'get lenght to bottom
                    Call Boots_Report_v_Alpha.Log_Push(text, "getting length to the bottom...")
                    On Error GoTo sp_B_cant_find_goalpost
                        dist_to_goalpost = Range("SP_PROPRIETARY_GOALPOST").row - row
                    On Error GoTo 0
                'get size
                    get_size_B_V1 = dist_to_goalpost
                    Call Boots_Report_v_Alpha.Log_Push(text, "Sp_v1_dev.get_size_B_V1 finish...")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'code end
                Exit Function
            'error handling
SP_get_size_B_V1_cant_find_SP_SHEET:
                'SP_get_size_B_V1_cant_find_SP_SHEET
                    'set error report
                        Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                        Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: SP_V1_dev.get_size_B_V1: was unable to find the sheet named 'Steel Presets'")
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'generate and store required information
                        Call Boots_Report_v_Alpha.Log_Push(text, "generating sheet list...")
                        Call Boots_Main_V_alpha.get_sheet_list
                    'push generated information
                        Call Boots_Report_v_Alpha.Log_Push(table_open)
                            Call Boots_Report_v_Alpha.Log_Push(text, "Displaying All existing sheets...")
                            For z = 1 To (wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).End(xlDown).row - wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).row)
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Sheet List: " & wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row + z, boots_pos.p_sheet_name_col).value & _
                                    " ='visible stat': " & wb.Sheets("boots").Cells(boots_pos.p_sheet_visible_status_row + z, boots_pos.p_sheet_visible_status_col).value)
                            Next z
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'give addresses
                        Call Boots_Report_v_Alpha.Log_Push(text, "generating Address list...")
                        Call Boots_Report_v_Alpha.Log_Push(table_open)
                        'home_pos
                            If home_pos Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: 'NOTHING' as worksheet")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: '" & home_pos.Parent.path & "\" & home_pos.Parent.Name & " == " & home_pos.index & ": " & home_pos.Name & "' as worksheet")
                            End If
                        'current_sht
                            If current_sht Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "current_sht: 'NOTHING' as worksheet")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "current_sht: '" & current_sht.Parent.path & "\" & current_sht.Parent.Name & " == " & current_sht.index & ": " & current_sht.Name & "' as worksheet")
                            End If
                        'wb
                            If wb Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "wb: 'NOTHING' as workbook")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "wb: '" & wb.path & "\" & wb.Name & "' as workbook")
                            End If
                        'activeworkbook
                            If ActiveWorkbook Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                            End If
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'TABLE CLOSE
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                        'end procedure
                            For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                            Next z
                        'call end statement
                            Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                            End
sp_B_cant_find_goalpost:
                'SP_get_size_B_V1_cant_find_SP_SHEET
                    'set error report
                        Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                        Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: SP_V1_dev.get_size_B_V1: was unable to find the range named 'SP_GENERAL_GOALPOST'. please check your code.")
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'generate and store required information
                        Call Boots_Report_v_Alpha.Log_Push(text, "generating sheet list...")
                        Call Boots_Main_V_alpha.get_sheet_list
                    'push generated information
                        Call Boots_Report_v_Alpha.Log_Push(table_open)
                            Call Boots_Report_v_Alpha.Log_Push(text, "Displaying All existing sheets...")
                            For z = 1 To (wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).End(xlDown).row - wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).row)
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Sheet List: " & wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row + z, boots_pos.p_sheet_name_col).value & _
                                    " ='visible stat': " & wb.Sheets("boots").Cells(boots_pos.p_sheet_visible_status_row + z, boots_pos.p_sheet_visible_status_col).value)
                            Next z
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'give addresses
                        Call Boots_Report_v_Alpha.Log_Push(text, "generating Address list...")
                        Call Boots_Report_v_Alpha.Log_Push(table_open)
                        'home_pos
                            If home_pos Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: 'NOTHING' as worksheet")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: '" & home_pos.Parent.path & "\" & home_pos.Parent.Name & " == " & home_pos.index & ": " & home_pos.Name & "' as worksheet")
                            End If
                        'current_sht
                            If current_sht Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "current_sht: 'NOTHING' as worksheet")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "current_sht: '" & current_sht.Parent.path & "\" & current_sht.Parent.Name & " == " & current_sht.index & ": " & current_sht.Name & "' as worksheet")
                            End If
                        'wb
                            If wb Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "wb: 'NOTHING' as workbook")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "wb: '" & wb.path & "\" & wb.Name & "' as workbook")
                            End If
                        'activeworkbook
                            If ActiveWorkbook Is Nothing Then
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                            Else
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                            End If
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'TABLE CLOSE
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                        'end procedure
                            For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                            Next z
                        'call end statement
                            Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                            End
        End Function

















