Attribute VB_Name = "HP_V3_stable"
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
                                                                        'Purpose Case
                    'This Module is built to handle all referances to the Price Tool HARDWARE PRESETS (HP_v3_stable) database For proper Referenceing and Updating
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Enum HP_POS_1
    'Purpose Case & notes:
        'POS Enum is to be called to act as a check condition to verify that the code and the sheet agrees on
            'the locational position of where things are on the sheet.
        'list
            'other table information
                
            'number of tracked fields
                Q_HP_TABLE_A_NUMBER_OF_TRACKED_POSITIONs = 7
                Q_HP_TABLE_B_NUMBER_OF_TRACKED_POSITIONS = 3
                Q_HP_OTHER_NUMBER_OF_TRACKED_POSITIONS = 2
                Q_HP_total_number_of_tracked_positions = HP_POS_1.Q_HP_OTHER_NUMBER_OF_TRACKED_POSITIONS + HP_POS_1.Q_HP_TABLE_A_NUMBER_OF_TRACKED_POSITIONs + HP_POS_1.Q_HP_TABLE_B_NUMBER_OF_TRACKED_POSITIONS
            'array positions
                ' table A
                    A_HP_GENERAL_PREFIX_row = 4
                        A_HP_GENERAL_PREFIX_col = 1
                    A_HP_GENERAL_DESCRIPTION_row = HP_POS_1.A_HP_GENERAL_PREFIX_row
                        A_HP_GENERAL_DESCRIPTION_col = 2
                    A_HP_GENERAL_XSMALL_row = HP_POS_1.A_HP_GENERAL_PREFIX_row
                        A_HP_GENERAL_XSMALL_col = 3
                    A_HP_GENERAL_SMALL_row = HP_POS_1.A_HP_GENERAL_PREFIX_row
                        A_HP_GENERAL_SMALL_col = 4
                    A_HP_GENERAL_MEDIUM_row = HP_POS_1.A_HP_GENERAL_PREFIX_row
                        A_HP_GENERAL_MEDIUM_col = 5
                    A_HP_GENERAL_LARGE_row = HP_POS_1.A_HP_GENERAL_PREFIX_row
                        A_HP_GENERAL_LARGE_col = 6
                    A_HP_GENERAL_XLARGE_row = HP_POS_1.A_HP_GENERAL_PREFIX_row
                        A_HP_GENERAL_XLARGE_col = 7
                'table B
                    B_HP_PROPRIETARY_PART_NUMBER_row = 4
                        B_HP_PROPRIETARY_part_number_col = 10
                    B_HP_PROPRIETARY_DESCRIPTION_row = HP_POS_1.B_HP_PROPRIETARY_PART_NUMBER_row
                        B_HP_PROPRIETARY_DESCRIPTION_col = 11
                    B_HP_PROPRIETARY_UNIT_COST_row = HP_POS_1.B_HP_PROPRIETARY_PART_NUMBER_row
                        B_HP_PROPRIETARY_UNIT_COST_col = 12
End Enum

Public Enum HP_get_table_A_sizing
    'Purpose Case & notes:
        'POS Enum is to be called to give position and value to check against
    'list
        'sizing
            x_small '= "0.25" '1/4
            Small '= "0.375"  '3/8
            medium '= "0.625" '5/8
            Large '= "0.875"  '7/8
            x_large '= "1.5"  '1-1/2
End Enum
        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                                                
                                                                        'GET GLOBAL FUNCTIONS

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
        
Public Function get_table_A_sizing_V0(ByVal size As HP_get_table_A_sizing, Optional more_instructions As String) As Variant
'Created By (Zachary Daugherty)(12/18/2020)
'Purpose Case & notes:
    If (more_instructions = "help") Then
        'Boots_Report_v_Alpha.Push_notification_message ("")
        Exit Function
    End If
'check for log reporting
    If (more_instructions = "Log_Report") Then
        get_table_A_sizing_V0 = "get_table_A_sizing_V0 - Public - under development 12-18-2020 - help file:N"
        Exit Function
    End If
'code start
    If (more_instructions <> "d_report") Then Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.size As HP_get_table_A_sizing Starting...")
    If (more_instructions <> "d_report") Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
    Select Case size
        Case HP_get_table_A_sizing.Large
            get_table_A_sizing_V0 = 0.875
        Case HP_get_table_A_sizing.medium
            get_table_A_sizing_V0 = 0.625
        Case HP_get_table_A_sizing.Small
            get_table_A_sizing_V0 = 0.375
        Case HP_get_table_A_sizing.x_large
            get_table_A_sizing_V0 = 1.5
        Case HP_get_table_A_sizing.x_small
            get_table_A_sizing_V0 = 0.25
        Case Else
            Boots_Report_v_Alpha.Push_notification_message ("HP_v3_stable.get_table_A_sizing_V0: selected a HP_get_table_A_sizing option that is not programed.")
            Stop
    End Select
    If (more_instructions <> "d_report") Then Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.size As HP_get_table_A_sizing Finish...")
    If (more_instructions <> "d_report") Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    If (more_instructions <> "d_report") Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
End Function
        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                                                
                                                'LOG VERSION REPORTING FUNCTIONS ONLY DO NOT REMOVE

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
        
Private Function LOG_push_version(ByVal pos As Long, Optional more_instructions As String) As Variant 'for version reporting
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            LOG_push_version = "LOG_push_version - Private - Stable"
            Exit Function
        End If
    
'code start
    Dim version As String
    Dim sht As Worksheet
    
    version = "3 stable |12/04/2020|"
    
    If (Boots_Main_V_alpha.sheet_exist(ActiveWorkbook, "Boots") = True) Then
        Set sht = ActiveWorkbook.Sheets("Boots")
    Else
        Exit Function
    End If
    sht.Cells(boots_pos.p_track_module_version_row + pos, boots_pos.p_track_module_version_col).value = version
    
End Function

Private Function LOG_push_project_file_requirements(Optional more_instructions As String) As Variant
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            LOG_push_project_file_requirements = "LOG_push_project_file_requirements - Private - Stable"
            Exit Function
        End If
        
'code start
    LOG_push_project_file_requirements = _
    "<boots_main_v_alpha><boots_report_v_alpha><string_v1>" & Chr(149) & _
    ""
End Function

Private Function LOG_Push_Functions_v1(Optional more_instructions As String) As Variant
'this function exists to report the status of code to the boots report log manager to allow the reading of information on all of the projects
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            LOG_Push_Functions_v1 = "LOG_Push_Functions_v1 - Private - Stable"
            Exit Function
        End If
'code start
    'define variables
        'addresses
            Dim sht As Worksheet
            Dim home As Worksheet
            Dim wb As Workbook
        'containers
            Dim i As Long
            Dim s As String
    'setup variables
        'addresses
            Set wb = ActiveWorkbook
            Set home = ActiveSheet
            Set sht = wb.Sheets("LOG_" & Boots_Main_V_alpha.get_username)
    'find open position on the log table
        i = Boots_Report_v_Alpha.Log_get_length_of_log_list_V1
    'get each status
        'log functions
            'LOG header
                s = "__________Project Object LOG Functions__________"
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning LOG_push_version
                    s = HP_V3_stable.LOG_push_version(0, "Log_Report")
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning LOG_push_project_file_requirements
                    s = HP_V3_stable.LOG_push_project_file_requirements("Log_Report")
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning LOG_Push_Functions_v1
                    s = HP_V3_stable.LOG_Push_Functions_v1("Log_Report")
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
            'ENUM
                s = "__________Project Object ENUM Functions__________"
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning HP_POS_1
                    s = "HP_POS_1 - Public - Stable 12-16-2020"
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'HP_get_table_A_sizing
                    s = "HP_get_table_A_sizing - Public - Stable 12-18-2020"
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
            'DO
                s = "__________Project Object DO Functions__________"
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning DO_Check_HP_A_Table_V1
                    s = HP_V3_stable.DO_Check_HP_A_Table_V1("Log_Report")
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning Do_Check_HP_B_Table_V1
                    s = HP_V3_stable.Do_Check_HP_B_Table_V1("Log_Report")
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning get_table_A_sizing_V0
                    s = HP_V3_stable.get_table_A_sizing_V0(1, "Log_Report")
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
            'get
                s = "__________Project Object Get Functions__________"
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning get_HP_sheet_name_v1
                    s = HP_V3_stable.get_HP_sheet_name_v1("Log_Report")
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning Get_size_HP_A_V1
                    s = HP_V3_stable.Get_size_HP_A_V1("Log_Report")
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                'Returning Get_size_HP_B_V1
                    s = HP_V3_stable.Get_size_HP_B_V1("Log_Report")
                        sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
            'close table
                s = "__________________________________________________________________________________________________"
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
                s = "__________________________________________________________________________________________________"
                    sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                            sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                    i = i + 1
    'cleanup
        On Error Resume Next
            home.Activate
        On Error GoTo 0
        LOG_Push_Functions = True
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

                                                                        'Do Statements
                                                                        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Public Function DO_Check_HP_A_Table_V1(Optional more_instructions As String) As Variant
    'Created By (Zachary Daugherty)(11/17/2020)
    'Purpose Case & notes:
        If (more_instructions = "help") Then
            Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "HP_V3_stable.DO_Check_HP_A_Table_V1: Help File Triggered..." & Chr(13) & _
                "What is this function for? (updated:12/03/2020):" & Chr(13) & _
                "    DO_Check_HP_A_Table_V1 is a function that SHOULD be called at the start of any GET OR SET operation on the HP_B_table." & Chr(13) & _
                "    As when any operation is done to the data table it should first be verifyed that the current version of the program" & Chr(13) & _
                "    and the addressed cell locations agree on their position. This is done by Calling ENUM HP_POS_1 and compairing the" & Chr(13) & _
                "    positional data indexed per what is expected." & Chr(13) & _
                "Should i call this function directly?(updated:12/03/2020):" & Chr(13) & _
                "    you can to check the operation status of the table but" & Chr(13) & _
                "    IF YOU PLAN ON DOING OPERATIONS TO THE TABLE REGARDING READING OR WRITING MAKE SURE THAT OPERATION IS CALLED IN 'DTH_VA.RUN' or" & Chr(13) & _
                "    other appropriate run functions this is done to protect the integrity of the file as if any run operation interacts" & Chr(13) & _
                "    with the HP sheet this will/should be called anyway before anything is done" & Chr(13) & _
                "What is returned from this function?(updated:12/03/2020):" & Chr(13) & _
                "    the function will return 'true' if all positions listed match the enumeration positions and there is no special calls done in more_instructions" & Chr(13) & _
                "    the function will return 'false' if all positions listed don't the enumeration positions and there is no special calls done in more_instructions" & Chr(13) & _
                "Listing Off dependants of Function (updated:12/03/2020):..." & Chr(13) & _
                "    HP_V3_stable." & Chr(13) & _
                "        HP_V3_stable.|parent module|" & Chr(13) & _
                "    Boots_Main_V_alpha." & Chr(13) & _
                "        Boots_Main_V_alpha.get_username" & Chr(13) & _
                "    Boots_Report_v_Alpha.:" & Chr(13) & _
                "        Boots_Report_v_Alpha.Log_Push" & Chr(13) & _
                "        Boots_Report_v_Alpha.Push_notification_message" & Chr(13) & _
                "        Boots_Report_v_Alpha.Log_get_indent_value_V0" & Chr(13) & _
                "    String_V1.:" & Chr(13) & _
                "        String_V1.get_Special_Char_V1")
            Exit Function
        End If
'check for log reporting
    If (more_instructions = "Log_Report") Then
        DO_Check_HP_A_Table_V1 = "DO_Check_HP_A_Table_V1 - Public - Stable 12-03-2020 - help file:Y"
        Exit Function
    End If
'code start
    Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.DO_Check_HP_A_Table_V1 Starting...")
    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
    'define varables
         Call Boots_Report_v_Alpha.Log_Push(text, "Define Varables...")
          Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
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
            Set proj_wb = ActiveWorkbook
            On Error GoTo FATAL_ERROR_CHECK_HP_A_SET_HP_ENV_For_A 'set error handler
                Set cursor_sheet = proj_wb.Sheets("HARDWARE PRESETS")
            On Error GoTo 0 'set error handler back to norm
            cursor_row = 1
            cursor_col = 1
         Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    'setup arr
         Call Boots_Report_v_Alpha.Log_Push(text, "Setup of Array...")
          Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
        'redefine size of the arr
            Call Boots_Report_v_Alpha.Log_Push(text, "Redefine Array")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            ReDim arr(1 To HP_POS_1.Q_HP_TABLE_A_NUMBER_OF_TRACKED_POSITIONs, 1 To 5) 'see line below for definitions
                'arr memory assignments
                    '(<specific index>,<1 to 5>)
                    '(<specific index>,<1:row of enum>)
                    '(<specific index>,<2:col of enum>)
                    '(<specific index>,<3:row of range>)
                    '(<specific index>,<4:col of range>)
                    '(<specific index>,<5: conditional if match>)
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'fill arr
            Call Boots_Report_v_Alpha.Log_Push(text, "Filling arr...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'Collect information
                i = 0
                '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                'NOTICE CODE IN THIS BLOCK IS STD AND THE OPERATIONS ARE THE SAME SO DEV NOTES ON THE FIRST FOLLOW THRU
                '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                'compair <HP_GENERAL_PREFIX> expected location
                    s = "HP_GENERAL_PREFIX"                                 'expected range name for search
                    i = i + 1                                               'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_For_A    'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                              'set range
                    On Error GoTo 0                                         'reset error handler
                    On Error GoTo ERROR_FATAL_DO_Check_HP_A_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                       'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                    'get range col pos
                        arr(i, 3) = HP_POS_1.A_HP_GENERAL_PREFIX_row        'get enum row pos
                        arr(i, 4) = HP_POS_1.A_HP_GENERAL_PREFIX_col        'get enum col pos
                    On Error GoTo 0
                    If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                        arr(i, 5) = s & ": " & True 'if true report text
                    Else
                        arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                        condition = True    'if true at the end of the block throw error as there is a miss match
                    End If
                'compair <HP_GENERAL_DESCRIPTION> expected location
                    s = "HP_GENERAL_DESCRIPTION"                            'expected range name for search
                    i = i + 1                                               'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_For_A    'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                              'set range
                    On Error GoTo 0                                         'reset error handler
                    On Error GoTo ERROR_FATAL_DO_Check_HP_A_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                       'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                    'get range col pos
                        arr(i, 3) = HP_POS_1.A_HP_GENERAL_DESCRIPTION_row   'get enum row pos
                        arr(i, 4) = HP_POS_1.A_HP_GENERAL_DESCRIPTION_col   'get enum col pos
                    On Error GoTo 0
                    If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                        arr(i, 5) = s & ": " & True 'if true report text
                    Else
                        arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                        condition = True    'if true at the end of the block throw error as there is a miss match
                    End If
                'compair <HP_GENERAL_XSMALL> expected location
                    s = "HP_GENERAL_XSMALL"                                 'expected range name for search
                    i = i + 1                                               'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_For_A    'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                              'set range
                    On Error GoTo 0                                         'reset error handler
                    On Error GoTo ERROR_FATAL_DO_Check_HP_A_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                       'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                    'get range col pos
                        arr(i, 3) = HP_POS_1.A_HP_GENERAL_XSMALL_row
                        arr(i, 4) = HP_POS_1.A_HP_GENERAL_XSMALL_col        'get enum col pos
                    On Error GoTo 0
                    If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                        arr(i, 5) = s & ": " & True 'if true report text
                    Else
                        arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                        condition = True    'if true at the end of the block throw error as there is a miss match
                    End If
                'compair <HP_GENERAL_SMALL> expected location
                    s = "HP_GENERAL_SMALL"                                  'expected range name for search
                    i = i + 1                                               'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_For_A    'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                              'set range
                    On Error GoTo 0                                         'reset error handler
                    On Error GoTo ERROR_FATAL_DO_Check_HP_A_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                       'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                    'get range col pos
                        arr(i, 3) = HP_POS_1.A_HP_GENERAL_SMALL_row         'get enum row pos
                        arr(i, 4) = HP_POS_1.A_HP_GENERAL_SMALL_col         'get enum col pos
                    On Error GoTo 0
                    If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                        arr(i, 5) = s & ": " & True 'if true report text
                    Else
                        arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                        condition = True    'if true at the end of the block throw error as there is a miss match
                    End If
                'compair <HP_GENERAL_MEDIUM> expected location
                    s = "HP_GENERAL_MEDIUM"                                 'expected range name for search
                    i = i + 1                                               'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_For_A    'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                              'set range
                    On Error GoTo 0                                         'reset error handler
                    On Error GoTo ERROR_FATAL_DO_Check_HP_A_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                       'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                    'get range col pos
                        arr(i, 3) = HP_POS_1.A_HP_GENERAL_MEDIUM_row        'get enum row pos
                        arr(i, 4) = HP_POS_1.A_HP_GENERAL_MEDIUM_col        'get enum col pos
                    On Error GoTo 0
                    If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                        arr(i, 5) = s & ": " & True 'if true report text
                    Else
                        arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                        condition = True    'if true at the end of the block throw error as there is a miss match
                    End If
                'compair <HP_GENERAL_LARGE> expected location
                    s = "HP_GENERAL_LARGE"                                          'expected range name for search
                    i = i + 1                                                       'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_For_A            'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                                      'set range
                    On Error GoTo 0                                                 'reset error handler
                    On Error GoTo ERROR_FATAL_DO_Check_HP_A_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                               'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                            'get range col pos
                        arr(i, 3) = HP_POS_1.A_HP_GENERAL_LARGE_row                 'get enum row pos
                        arr(i, 4) = HP_POS_1.A_HP_GENERAL_LARGE_col                 'get enum col pos
                    On Error GoTo 0
                    If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                        arr(i, 5) = s & ": " & True                                 'if true report text
                    Else
                        arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                        condition = True    'if true at the end of the block throw error as there is a miss match
                    End If
                'compair <HP_GENERAL_XLARGE> expected location
                    s = "HP_GENERAL_XLARGE"                                 'expected range name for search
                    i = i + 1                                               'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_For_A    'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                              'set range
                    On Error GoTo 0                                         'reset error handler
                    On Error GoTo ERROR_FATAL_DO_Check_HP_A_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                       'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                    'get range col pos
                        arr(i, 3) = HP_POS_1.A_HP_GENERAL_XLARGE_row        'get enum row pos
                        arr(i, 4) = HP_POS_1.A_HP_GENERAL_XLARGE_col        'get enum col pos
                    On Error GoTo 0
                        If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                            arr(i, 5) = s & ": " & True 'if true report text
                        Else
                            arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                            condition = True    'if true at the end of the block throw error as there is a miss match
                        End If
                '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                '__________________________________________END of CODE BLOCK___________________________________________
                '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                'cleanup
                    i = 0
                    s = "Empty"
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'cleanup
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    'compile report
        'check to see if failure condition is met
            If (condition = True) Then
                GoTo ERROR_CHECK_HP_FAILED_POS_CHECK_For_A
            End If
        'return true
            DO_Check_HP_A_Table_V1 = True   'passed all checks
            Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.DO_Check_HP_A_Table_V1 Finishing...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            Exit Function
'code end
'error handle
ERROR_FATAL_DO_Check_HP_A_Table_V1_matrix_sz:
    Call Boots_Report_v_Alpha.Log_Push(Error_)
        Call Boots_Report_v_Alpha.Log_Push(Flag)
            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: HP_V3_stable.DO_Check_HP_A_Table_V1 MATRIX TABLE WAS NOT ABLE TO FILL AS THE SIZE WAS NOT LARGE ENOUGH...")
        Call Boots_Report_v_Alpha.Log_Push(table_close)
        Call Boots_Report_v_Alpha.Log_Push(text, "listing matrix information...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            Call Boots_Report_v_Alpha.Log_Push(text, "arr position:" & i & " was unable to add the data position '" & s & "' please check that the redim has enough space for all needed information...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Call Boots_Report_v_Alpha.Log_Push(text, "Listing addresses...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'proj_wb
                If proj_wb Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: 'NOTHING' as Workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: '" & proj_wb.path & "/" & proj_wb.Name & "' as Workbook")
                End If
            'cursor_sht
                If cursor_sheet Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: 'NOTHING' as worksheet")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: '" & cursor_sheet.Parent.path & "\" & cursor_sheet.Parent.Name & " == " & cursor_sheet.index & ": " & cursor_sheet.Name & "' as worksheet")
                End If
            'activeworkbook
                If ActiveWorkbook Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                End If
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Call Boots_Report_v_Alpha.Log_Push(table_close)
    'exit procedure
    Call Boots_Report_v_Alpha.Log_Push(text, "error code last updated on: 12/3/2020")
    For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Next z
    Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
    End
ERROR_FATAL_check_HP_range_error_For_A:
    Call Boots_Report_v_Alpha.Log_Push(Error_)
        Call Boots_Report_v_Alpha.Log_Push(Flag)
            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: HP_V3_stable.DO_Check_HP_A_Table_V1 UNABLE TO LOCATE THE SPECIFIED RANGE:")
            Call Boots_Report_v_Alpha.Log_Push(text, "'<" & s & ">")
            Call Boots_Report_v_Alpha.Log_Push(text, "please check the name mannager for errors. fix and then re-run")
        Call Boots_Report_v_Alpha.Log_Push(table_close)
        Call Boots_Report_v_Alpha.Log_Push(text, "Listing addresses...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'proj_wb
                If proj_wb Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: 'NOTHING' as Workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: '" & proj_wb.path & "/" & proj_wb.Name & "' as Workbook")
                End If
            'cursor_sht
                If cursor_sheet Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: 'NOTHING' as worksheet")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: '" & cursor_sheet.Parent.path & "\" & cursor_sheet.Parent.Name & " == " & cursor_sheet.index & ": " & cursor_sheet.Name & "' as worksheet")
                End If
            'activeworkbook
                If ActiveWorkbook Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                End If
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Call Boots_Report_v_Alpha.Log_Push(table_close)
    'exit procedure
    Call Boots_Report_v_Alpha.Log_Push(text, "error code last updated on: 12/3/2020")
    For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Next z
    Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
    End
FATAL_ERROR_CHECK_HP_A_SET_HP_ENV_For_A:
    Call Boots_Report_v_Alpha.Log_Push(Error_)
        Call Boots_Report_v_Alpha.Log_Push(Flag)
            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: HP_V3_stable.DO_Check_HP_A_Table_V1 UNABLE TO FIND OR SET SHEET Hardware Presets IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.")
        Call Boots_Report_v_Alpha.Log_Push(table_close)
                Call Boots_Report_v_Alpha.Log_Push(text, "Listing addresses...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'proj_wb
                If proj_wb Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: 'NOTHING' as Workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: '" & proj_wb.path & "/" & proj_wb.Name & "' as Workbook")
                End If
            'cursor_sht
                If cursor_sheet Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: 'NOTHING' as worksheet")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: '" & cursor_sheet.Parent.path & "\" & cursor_sheet.Parent.Name & " == " & cursor_sheet.index & ": " & cursor_sheet.Name & "' as worksheet")
                End If
            'activeworkbook
                If ActiveWorkbook Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                End If
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Call Boots_Report_v_Alpha.Log_Push(table_close)
    'exit procedure
        Call Boots_Report_v_Alpha.Log_Push(text, "error code last updated on: 12/3/2020")
        For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Next z
        Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
        End
ERROR_CHECK_HP_FAILED_POS_CHECK_For_A:
    Call Boots_Report_v_Alpha.Log_Push(Error_)
        Call Boots_Report_v_Alpha.Log_Push(Flag)
            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: HP_V3_stable.DO_Check_HP_A_Table_V1 FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: ")
            Call Boots_Report_v_Alpha.Log_Push(text, arr(1, 5))
            Call Boots_Report_v_Alpha.Log_Push(text, arr(2, 5))
            Call Boots_Report_v_Alpha.Log_Push(text, arr(3, 5))
            Call Boots_Report_v_Alpha.Log_Push(text, arr(4, 5))
            Call Boots_Report_v_Alpha.Log_Push(text, arr(5, 5))
            Call Boots_Report_v_Alpha.Log_Push(text, arr(6, 5))
            Call Boots_Report_v_Alpha.Log_Push(text, arr(7, 5))
        Call Boots_Report_v_Alpha.Log_Push(table_close)
                Call Boots_Report_v_Alpha.Log_Push(text, "Listing addresses...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'proj_wb
                If proj_wb Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: 'NOTHING' as Workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: '" & proj_wb.path & "/" & proj_wb.Name & "' as Workbook")
                End If
            'cursor_sht
                If cursor_sheet Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: 'NOTHING' as worksheet")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: '" & cursor_sheet.Parent.path & "\" & cursor_sheet.Parent.Name & " == " & cursor_sheet.index & ": " & cursor_sheet.Name & "' as worksheet")
                End If
            'activeworkbook
                If ActiveWorkbook Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                End If
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Call Boots_Report_v_Alpha.Log_Push(table_close)
    'exit procedure
        Call Boots_Report_v_Alpha.Log_Push(text, "error code last updated on: 12/3/2020")
        For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Next z
        Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
        End
End Function

Public Function Do_Check_HP_B_Table_V1(Optional more_instructions As String) As Variant
    'Created By (Zachary Daugherty)(8/11/2020)
    'Purpose Case & notes:
        If (more_instructions = "help") Then
            Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "HP_V3_stable.Do_Check_HP_B_Table_V1: Help File Triggered..." & Chr(13) & _
                "What is this function for? (updated:12/02/2020):" & Chr(13) & _
                "    Do_Check_HP_B_Table_V1 is a function that SHOULD be called at the start of any GET OR SET operation on the HP_B_table." & Chr(13) & _
                "    As when any operation is done to the data table it should first be verifyed that the current version of the program" & Chr(13) & _
                "    and the addressed cell locations agree on their position. This is done by Calling ENUM HP_POS_1 and compairing the" & Chr(13) & _
                "    positional data indexed per what is expected." & Chr(13) & _
                "Should i call this function directly?(updated:12/02/2020):" & Chr(13) & _
                "    you can to check the operation status of the table but" & Chr(13) & _
                "    IF YOU PLAN ON DOING OPERATIONS TO THE TABLE REGARDING READING OR WRITING MAKE SURE THAT OPERATION IS CALLED IN 'DTH_VA.RUN' or" & Chr(13) & _
                "    other appropriate run functions this is done to protect the integrity of the file as if any run operation interacts" & Chr(13) & _
                "    with the HP sheet this will/should be called anyway before anything is done" & Chr(13) & _
                "What is returned from this function?(updated:12/02/2020):" & Chr(13) & _
                "    the function will return 'true' if all positions listed match the enumeration positions and there is no special calls done in more_instructions" & Chr(13) & _
                "    the function will return 'false' if all positions listed don't the enumeration positions and there is no special calls done in more_instructions" & Chr(13) & _
                "Listing Off dependants of Function (updated:12/02/2020):..." & Chr(13) & _
                "    HP_V3_stable." & Chr(13) & _
                "        HP_V3_stable.|parent module|" & Chr(13) & _
                "    Boots_Main_V_alpha." & Chr(13) & _
                "        Boots_Main_V_alpha.get_username" & Chr(13) & _
                "    Boots_Report_v_Alpha.:" & Chr(13) & _
                "        Boots_Report_v_Alpha.Log_Push" & Chr(13) & _
                "        Boots_Report_v_Alpha.Push_notification_message" & Chr(13) & _
                "        Boots_Report_v_Alpha.Log_get_indent_value_V0" & Chr(13) & _
                "    String_V1.:" & Chr(13) & _
                "        String_V1.get_Special_Char_V1")
            Exit Function
        End If
    'returned outputs
        'returns:
            'true: if all positions match up
            'false: if any positions do not match up
'check for log reporting
    If (more_instructions = "Log_Report") Then
        Do_Check_HP_B_Table_V1 = "Do_Check_HP_B_Table_V1 - Public - Stable 12/03/2020 - help file:Y"
        Exit Function
    End If
'code start
    Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.Do_Check_HP_B_Table_V1 Starting...")
    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
    'define varables
         Call Boots_Report_v_Alpha.Log_Push(text, "Define Varables...")
          Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
        'memory
            Dim arr() As String             'designed as ram storage
            Dim condition As Boolean        'store T/F
            Dim i As Long                   'iterator and int storage
            Dim s As String                 'string storage
        'cursor
            Dim proj_wb As Workbook         'set local workbook
            Dim cursor_sheet As Worksheet   'sheet the cursor is on
            Dim cursor_row As Long          'self explains
            Dim cursor_col As Long          'self explains
        'ref
            Dim ref_rng As Range            'reference range in question
        'setup variables
            Set proj_wb = ActiveWorkbook
            On Error GoTo FATAL_ERROR_CHECK_HP_B_SET_HP_ENV_For_B 'set error handler
                Set cursor_sheet = proj_wb.Sheets("HARDWARE PRESETS")
            On Error GoTo 0 'set error handler back to norm
            cursor_row = 1
            cursor_col = 1
         Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    'setup arr
         Call Boots_Report_v_Alpha.Log_Push(text, "Setup of Array...")
          Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
        'redefine size of the arr
            Call Boots_Report_v_Alpha.Log_Push(text, "Redefine Array")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            ReDim arr(1 To HP_POS_1.Q_HP_TABLE_B_NUMBER_OF_TRACKED_POSITIONS, 1 To 5) 'see line below for definitions
                'arr memory assignments
                    '(<specific index>,<1 to 5>)
                    '(<specific index>,<1:row of enum>)
                    '(<specific index>,<2:col of enum>)
                    '(<specific index>,<3:row of range>)
                    '(<specific index>,<4:col of range>)
                    '(<specific index>,<5: conditional if match>)
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'fill arr
            Call Boots_Report_v_Alpha.Log_Push(text, "Filling arr...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'Collect information
                i = 0
                '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                'NOTICE CODE IN THIS BLOCK IS STD AND THE OPERATIONS ARE THE SAME SO DEV NOTES ON THE FIRST FOLLOW THRU
                '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                'compair <HP_PROPRIETARY_PART_NUMBER> expected location
                    s = "HP_PROPRIETARY_PART_NUMBER"                            'expected range name for search
                    i = i + 1                                                   'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_for_b        'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                                  'set range
                    On Error GoTo 0                                             'reset error handler
                    On Error GoTo ERROR_FATAL_Do_Check_HP_B_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                           'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                        'get range col pos
                        arr(i, 3) = HP_POS_1.B_HP_PROPRIETARY_PART_NUMBER_row   'get enum row pos
                        arr(i, 4) = HP_POS_1.B_HP_PROPRIETARY_part_number_col   'get enum col pos
                    On Error GoTo 0
                    If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                        arr(i, 5) = s & ": " & True 'if true report text
                    Else
                        arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                        condition = True    'if true at the end of the block throw error as there is a miss match
                    End If
                'compair <HP_PROPRIETARY_DESCRIPTION> expected location
                    s = "HP_PROPRIETARY_DESCRIPTION"                            'expected range name for search
                    i = i + 1                                                   'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_for_b        'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                                  'set range
                    On Error GoTo 0                                             'reset error handler
                    On Error GoTo ERROR_FATAL_Do_Check_HP_B_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                           'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                        'get range col pos
                        arr(i, 3) = HP_POS_1.B_HP_PROPRIETARY_DESCRIPTION_row   'get enum row pos
                        arr(i, 4) = HP_POS_1.B_HP_PROPRIETARY_DESCRIPTION_col   'get enum col pos
                    On Error GoTo 0
                    If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                        arr(i, 5) = s & ": " & True 'if true report text
                    Else
                        arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                        condition = True    'if true at the end of the block throw error as there is a miss match
                    End If
                'compair <HP_PROPRIETARY_UNIT_COST> expected location
                    s = "HP_PROPRIETARY_UNIT_COST"                                 'expected range name for search
                    i = i + 1                                               'iterate arr position from x to x + 1 in the array
                    On Error GoTo ERROR_FATAL_check_HP_range_error_for_b    'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                        Set ref_rng = Range(s)                              'set range
                    On Error GoTo 0                                         'reset error handler
                    On Error GoTo ERROR_FATAL_Do_Check_HP_B_Table_V1_matrix_sz
                        arr(i, 1) = CStr(ref_rng.row)                       'get range row pos
                        arr(i, 2) = CStr(ref_rng.Column)                    'get range col pos
                        arr(i, 3) = HP_POS_1.B_HP_PROPRIETARY_UNIT_COST_row 'get enum row pos
                        arr(i, 4) = HP_POS_1.B_HP_PROPRIETARY_UNIT_COST_col 'get enum col pos
                    On Error GoTo 0
                    If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                        arr(i, 5) = s & ": " & True 'if true report text
                    Else
                        arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                        condition = True    'if true at the end of the block throw error as there is a miss match
                    End If
                '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                '__________________________________________END of CODE BLOCK___________________________________________
                '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                'cleanup
                    i = 0
                    s = "Empty"
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'cleanup
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    'compile report
        'check to see if failure condition is met
            If (condition = True) Then
                GoTo ERROR_CHECK_HP_FAILED_POS_CHECK_For_B
            End If
        'return true
            Do_Check_HP_B_Table_V1 = True   'passed all checks
            Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.Do_Check_HP_B_Table_V1 Finishing...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            Exit Function
'code end
'error handle
ERROR_FATAL_Do_Check_HP_B_Table_V1_matrix_sz:
    Call Boots_Report_v_Alpha.Log_Push(Error_)
        Call Boots_Report_v_Alpha.Log_Push(Flag)
            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: HP_V3_stable.Do_Check_HP_B_Table_V1 MATRIX TABLE WAS NOT ABLE TO FILL AS THE SIZE WAS NOT LARGE ENOUGH...")
        Call Boots_Report_v_Alpha.Log_Push(table_close)
        Call Boots_Report_v_Alpha.Log_Push(text, "listing matrix information...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            Call Boots_Report_v_Alpha.Log_Push(text, "arr position:" & i & " was unable to add the data position '" & s & "' please check that the redim has enough space for all needed information...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Call Boots_Report_v_Alpha.Log_Push(text, "Listing addresses...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'proj_wb
                If proj_wb Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: 'NOTHING' as Workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: '" & proj_wb.path & "/" & proj_wb.Name & "' as Workbook")
                End If
            'cursor_sht
                If cursor_sheet Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: 'NOTHING' as worksheet")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: '" & cursor_sheet.Parent.path & "\" & cursor_sheet.Parent.Name & " == " & cursor_sheet.index & ": " & cursor_sheet.Name & "' as worksheet")
                End If
            'activeworkbook
                If ActiveWorkbook Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                End If
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Call Boots_Report_v_Alpha.Log_Push(table_close)
    'exit procedure
    Call Boots_Report_v_Alpha.Log_Push(text, "error code last updated on: 12/3/2020")
    For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Next z
    Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
    End
ERROR_FATAL_check_HP_range_error_for_b:
    Call Boots_Report_v_Alpha.Log_Push(Error_)
        Call Boots_Report_v_Alpha.Log_Push(Flag)
            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: HP_V3_stable.Do_Check_HP_B_Table_V1 UNABLE TO LOCATE THE SPECIFIED RANGE:")
            Call Boots_Report_v_Alpha.Log_Push(text, "'<" & s & ">")
            Call Boots_Report_v_Alpha.Log_Push(text, "please check the name mannager for errors. fix and then re-run")
        Call Boots_Report_v_Alpha.Log_Push(table_close)
        Call Boots_Report_v_Alpha.Log_Push(text, "Listing addresses...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'proj_wb
                If proj_wb Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: 'NOTHING' as Workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: '" & proj_wb.path & "/" & proj_wb.Name & "' as Workbook")
                End If
            'cursor_sht
                If cursor_sheet Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: 'NOTHING' as worksheet")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: '" & cursor_sheet.Parent.path & "\" & cursor_sheet.Parent.Name & " == " & cursor_sheet.index & ": " & cursor_sheet.Name & "' as worksheet")
                End If
            'activeworkbook
                If ActiveWorkbook Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                End If
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Call Boots_Report_v_Alpha.Log_Push(table_close)
    'exit procedure
    Call Boots_Report_v_Alpha.Log_Push(text, "error code last updated on: 12/3/2020")
    For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Next z
    Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
    End
FATAL_ERROR_CHECK_HP_B_SET_HP_ENV_For_B:
    Call Boots_Report_v_Alpha.Log_Push(Error_)
        Call Boots_Report_v_Alpha.Log_Push(Flag)
            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: HP_V3_stable.Do_Check_HP_B_Table_V1 UNABLE TO FIND OR SET SHEET Hardware Presets IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.")
        Call Boots_Report_v_Alpha.Log_Push(table_close)
                Call Boots_Report_v_Alpha.Log_Push(text, "Listing addresses...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'proj_wb
                If proj_wb Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: 'NOTHING' as Workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: '" & proj_wb.path & "/" & proj_wb.Name & "' as Workbook")
                End If
            'cursor_sht
                If cursor_sheet Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: 'NOTHING' as worksheet")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: '" & cursor_sheet.Parent.path & "\" & cursor_sheet.Parent.Name & " == " & cursor_sheet.index & ": " & cursor_sheet.Name & "' as worksheet")
                End If
            'activeworkbook
                If ActiveWorkbook Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                End If
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Call Boots_Report_v_Alpha.Log_Push(table_close)
    'exit procedure
        Call Boots_Report_v_Alpha.Log_Push(text, "error code last updated on: 12/3/2020")
        For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Next z
        Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
        End
ERROR_CHECK_HP_FAILED_POS_CHECK_For_B:
    Call Boots_Report_v_Alpha.Log_Push(Error_)
        Call Boots_Report_v_Alpha.Log_Push(Flag)
            Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: HP_V3_stable.Do_Check_HP_B_Table_V1 FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: ")
            Call Boots_Report_v_Alpha.Log_Push(text, arr(1, 5))
            Call Boots_Report_v_Alpha.Log_Push(text, arr(2, 5))
            Call Boots_Report_v_Alpha.Log_Push(text, arr(3, 5))
        Call Boots_Report_v_Alpha.Log_Push(table_close)
                Call Boots_Report_v_Alpha.Log_Push(text, "Listing addresses...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'proj_wb
                If proj_wb Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: 'NOTHING' as Workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: '" & proj_wb.path & "/" & proj_wb.Name & "' as Workbook")
                End If
            'cursor_sht
                If cursor_sheet Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: 'NOTHING' as worksheet")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: '" & cursor_sheet.Parent.path & "\" & cursor_sheet.Parent.Name & " == " & cursor_sheet.index & ": " & cursor_sheet.Name & "' as worksheet")
                End If
            'activeworkbook
                If ActiveWorkbook Is Nothing Then
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: 'NOTHING' as workbook")
                Else
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Activeworkbook: '" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "' as workbook")
                End If
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Call Boots_Report_v_Alpha.Log_Push(table_close)
    'exit procedure
        Call Boots_Report_v_Alpha.Log_Push(text, "error code last updated on: 12/3/2020")
        For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Next z
        Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
        End
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

                                                                        'Get Statements
                                                                        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Public Function Get_size_HP_A_V1(Optional more_instructions As String) As Variant
        'Created By (Zachary daugherty)(11/10/20)
        'Purpose Case & notes:
            If (more_instructions = "help") Then
                        Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "HP_V3_stable.Get_size_HP_A_V1: Help File Triggered..." & Chr(13) & _
                            "What is this function for? (updated:12/04/2020):" & Chr(13) & _
                            "    Get_size_HP_A_V1 is used to return the size of the Hardware presets sheet size" & Chr(13) & _
                            "Should i call this function directly? (updated:12/04/2020):" & Chr(13) & _
                            "    calling this function directly will work but all this will do is give the table size." & Chr(13) & _
                            "What is returned from this function? (updated:12/04/2020):" & Chr(13) & _
                            "    the function will return a number equal to the size of the sheet" & Chr(13) & _
                            "Listing Off dependants of Function (updated:12/04/2020):..." & Chr(13) & _
                            "    HP_V3_stable." & Chr(13) & _
                            "        HP_V3_stable.|parent module|" & Chr(13) & _
                            "    Boots_Main_V_alpha." & Chr(13) & _
                            "        Boots_Main_V_alpha.get_sheet_list" & Chr(13) & _
                            "    Boots_Report_v_Alpha.:" & Chr(13) & _
                            "        Boots_Report_v_Alpha.Log_Push" & Chr(13) & _
                            "        Boots_Report_v_Alpha.Push_notification_message" & Chr(13) & _
                            "        Boots_Report_v_Alpha.Log_get_indent_value_V0")
                        Exit Function
                    End If
        'check for log reporting
            If (more_instructions = "Log_Report") Then
                Get_size_HP_A_V1 = "Get_size_HP_A_V1 - Public - Stable with logs 12/4/2020 - help file:Y"
                Exit Function
            End If
        'code start
            Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.Get_size_HP_A_V1 Start...")
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
                    On Error GoTo HP_get_size_A_cant_find_HP_SHEET
                        Set current_sht = wb.Sheets("HARDWARE PRESETS")
                    On Error GoTo 0
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
            'Seting start location
            Call Boots_Report_v_Alpha.Log_Push(text, "Setting Start location...")
                row = HP_POS_1.A_HP_GENERAL_PREFIX_row
                col = HP_POS_1.A_HP_GENERAL_PREFIX_col
                s = current_sht.Cells(row, col).value
            'Browsing for the Goalpost position
            Call Boots_Report_v_Alpha.Log_Push(text, "Browsing for the Goalpost position...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                On Error GoTo HP_A_cant_find_goalpost
                    dist_to_goalpost = Range("HP_GENERAL_GOALPOST").row - row
                On Error GoTo 0
                'Returning result...
                    Call Boots_Report_v_Alpha.Log_Push(text, "Returning result...")
                    Get_size_HP_A_V1 = dist_to_goalpost
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'code end
                'cleanup
                            Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.Get_size_HP_A_V1 Finished...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                'return
                    Exit Function
        'error handling
HP_get_size_A_cant_find_HP_SHEET:
            'SP_get_size_A_cant_find_SP_SHEET:
                'set error report
                    Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                    Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                        Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: HP_V3_stable.Get_size_HP_A_V1 was unable to find the sheet named 'HARDWARE PRESETS'")
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
HP_A_cant_find_goalpost:
            'sp_A_cant_find_goalpost
                'set error report
                    Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                    Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                        Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: HP_V3_stable: FUNCTION  GET_SIZE_HP_A_V1: was unable to find the range named 'HP_GENERAL_GOALPOST'.")
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

Public Function Get_size_HP_B_V1(Optional more_instructions As String) As Variant
        'Created By (Zachary daugherty)(11/12/20)
        'Purpose Case & notes:
            If (more_instructions = "help") Then
                Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "HP_V3_stable.Get_size_HP_B_V1: Help File Triggered..." & Chr(13) & _
                    "What is this function for? (updated:12/04/2020):" & Chr(13) & _
                    "    Get_size_HP_B_V1 returns the size of the hardware presets page" & Chr(13) & _
                    "Should i call this function directly?(updated:12/04/2020):" & Chr(13) & _
                    "    you can but since this is a get function its just returning the size" & Chr(13) & _
                    "What is returned from this function?(updated:12/04/2020):" & Chr(13) & _
                    "    just the size of the hardware presets." & Chr(13) & _
                    "Listing Off dependants of Function (updated:12/04/2020):..." & Chr(13) & _
                    "    HP_V3_stable." & Chr(13) & _
                    "        HP_V3_stable.|parent module|" & Chr(13) & _
                    "    Boots_Main_V_alpha." & Chr(13) & _
                    "        Boots_Main_V_alpha.get_sheet_list" & Chr(13) & _
                    "    Boots_Report_v_Alpha.:" & Chr(13) & _
                    "        Boots_Report_v_Alpha.Log_Push" & Chr(13) & _
                    "        Boots_Report_v_Alpha.Push_notification_message")
                Exit Function
            End If
        'check for log reporting
            If (more_instructions = "Log_Report") Then
                Get_size_HP_B_V1 = "Get_size_HP_B_V1 - Public - Stable with logs 11/11/2020 - help file:Y"
                Exit Function
            End If
        'code start
            Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.Get_size_HP_B_V1 Start...")
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
                    On Error GoTo HP_get_size_B_cant_find_HP_SHEET
                        Set current_sht = wb.Sheets("HARDWARE PRESETS")
                    On Error GoTo 0
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
            'Seting start location
            Call Boots_Report_v_Alpha.Log_Push(text, "Setting Start location...")
                row = HP_POS_1.B_HP_PROPRIETARY_PART_NUMBER_row
                col = HP_POS_1.B_HP_PROPRIETARY_part_number_col
                s = current_sht.Cells(row, col).value
            'Browsing for the Goalpost position
            Call Boots_Report_v_Alpha.Log_Push(text, "Browsing for the Goalpost position...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                On Error GoTo HP_B_cant_find_goalpost
                    dist_to_goalpost = Range("HP_PROPRIETARY_GOALPOST").row - row
                On Error GoTo 0
                'Returning result...
                    Call Boots_Report_v_Alpha.Log_Push(text, "Returning result...")
                    Get_size_HP_B_V1 = dist_to_goalpost
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'code end
                'cleanup
                            Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.Get_size_HP_B_V1 Finished...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                'return
                    Exit Function
        'error handling
HP_get_size_B_cant_find_HP_SHEET:
            'SP_get_size_A_cant_find_SP_SHEET:
                'set error report
                    Call Boots_Report_v_Alpha.Log_Push(Error_)
                    Call Boots_Report_v_Alpha.Log_Push(Flag)
                        Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: HP_V3_stable.Get_size_HP_B_V1 was unable to find the sheet named 'HARDWARE PRESETS'")
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
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Next z
                    'call end statement
                        Call Boots_Report_v_Alpha.Log_Push(Display_now)
                        End
HP_B_cant_find_goalpost:
            'sp_A_cant_find_goalpost
                'set error report
                    Call Boots_Report_v_Alpha.Log_Push(Error_)
                    Call Boots_Report_v_Alpha.Log_Push(Flag)
                        Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: HP_V3_stable: FUNCTION  GET_SIZE_HP_B_V1: was unable to find the range named 'HP_PROPRIETARY_GOALPOST'.")
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
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Next z
                    'call end statement
                        Call Boots_Report_v_Alpha.Log_Push(Display_now)
                        End
End Function


Public Function get_HP_sheet_name_v1(Optional more_instructions As Variant) As Variant
    'Created By (Zachary Daugherty)(12/01/2020)
    'Purpose Case & notes:
        If (more_instructions = "help") Then
            Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "HP_V3_stable.get_HP_sheet_name_v1: Help File Triggered..." & Chr(13) & _
                "What is this function for? (updated:12/04/2020):" & Chr(13) & _
                "    get_HP_sheet_name_v1 returns the name of the Hardware presets sheet name" & Chr(13) & _
                "Should i call this function directly?(updated:12/04/2020):" & Chr(13) & _
                "    you can but since this is a get function its just returning the name" & Chr(13) & _
                "What is returned from this function?(updated:12/04/2020):" & Chr(13) & _
                "    just the name of the hardware presets page name. why? this is to make it easy to change the name of the sheet and not need to make a change to each" & Chr(13) & _
                "    function individually" & Chr(13) & _
                "Listing Off dependants of Function (updated:12/02/2020):..." & Chr(13) & _
                "    HP_V3_stable." & Chr(13) & _
                "        HP_V3_stable.|parent module|" & Chr(13) & _
                "    Boots_Report_v_Alpha.:" & Chr(13) & _
                "        Boots_Report_v_Alpha.Log_Push" & Chr(13) & _
                "        Boots_Report_v_Alpha.Push_notification_message")
            Exit Function
        End If
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            get_HP_sheet_name_v1 = "HP_V3_stable.get_HP_sheet_name_v1 - Public - Stable with logs 11/12/2020 - help file:Y"
            Exit Function
        End If
    'code start
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.get_HP_sheet_name_v1 starting...")
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            get_HP_sheet_name_v1 = "HARDWARE PRESETS"
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "HP_V3_stable.get_HP_sheet_name_v1 Finish...")
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    Exit Function
End Function
