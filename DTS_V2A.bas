Attribute VB_Name = "DTS_V2A"
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
                                                        'please see function : .LOG_push_project_file_requirements
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Purpose Case
                    'This Module is built to handle all referances to the Price Tool DTS database For proper Referenceing and Updating
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Enum DTS_run_choices_V1
    'Purpose Case & notes:
        'gives the enumeration of choices that are setup for options
    'list
        DTS_update_unit_cost
End Enum
Public Enum get_choices_v1
    'Purpose Case & notes:
        'gives the enumeration of choices that are setup for options
    'list
        get_size
End Enum
Public Enum DTS_POS_2A
    'Purpose Case & notes:
        'POS Enum is to be called to act as a check condition to verify that the code and the sheet agrees on
            'the locational position of where things are on the sheet.
    'list
        'other table information
            DTS_I_Inflation_Const_row = 1
                DTS_I_Inflation_Const_col = 5
        'number of entry fields watching
            DTS_Q_number_of_tracked_locations = 14
        'array of positions
            DTS_I_part_number_row = 3 'used as global table header row position
                DTS_I_part_number_col = 1
            DTS_I_AKA_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_AKA_col = 2
            DTS_I_Description_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Description_col = 3
            DTS_I_Unit_cost_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Unit_cost_col = 4
            DTS_I_Adjusted_unit_cost_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Adjusted_unit_cost_col = 5
            DTS_I_Unit_weight_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Unit_weight_col = 6
            DTS_I_Shop_origin_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Shop_origin_col = 7
            DTS_I_Status_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Status_col = 8
            DTS_I_Job_other_info_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Job_other_info_col = 9
            DTS_I_Vendor_info_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Vendor_info_col = 10
            DTS_I_Vendor_phone_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Vendor_phone_col = 11
            DTS_I_Vendor_fax_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Vendor_fax_col = 12
            DTS_I_Vendor_part_number_row = DTS_POS_2A.DTS_I_part_number_row
                DTS_I_Vendor_part_number_col = 13
End Enum
        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                                                
                                                                        'GET GLOBAL FUNCTIONS

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
        
        Public Const get_global_unit_cost_refresh_ignore_trigger As String = "<skip>"
        
        Public Const get_global_decoder_symbol = "-"
        
        
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
            
            version = "2.0 Alpha with some logs"
            
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
            "<Boots_main_Valpha>" & Chr(149) & _
            "<Boots_Report_Valpha>" & Chr(149) & _
            "<SP_V1>" & Chr(149) & _
            "<Matrix_V2>" & Chr(149) & _
            "<String_V1>" & Chr(149)
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
                    'returning log push version
                        s = DTS_V2A.LOG_push_version(0, "Log_Report")
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                        i = i + 1
                    'returning LOG_push_project_file_requirements
                        s = DTS_V2A.LOG_push_project_file_requirements("Log_Report")
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                        i = i + 1
                    'returning LOG_Push_Functions_v1
                        s = DTS_V2A.LOG_Push_Functions_v1("Log_Report")
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                        i = i + 1
                'returning listed enums
                    'Enum header
                        s = "__________Project Object ENUMS__________"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                    'DTS_run_choices_V1
                        s = "ENUM: DTS_run_choices_V1 - Public"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                        i = i + 1
                    'get_choices_v1
                        s = "ENUM: get_choices_v1 - Pubilc"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                        i = i + 1
                    'DTS_POS_2A
                        s = "ENUM: DTS_POS_2A - Public"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                        i = i + 1
                'returning globals
                    'Global variable header
                        s = "__________Project Object global variable set__________"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                    'get_global_unit_cost_refresh_ignore_trigger
                        s = "get_global_unit_cost_refresh_ignore_trigger - Public - Stable"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                        i = i + 1
                    'get_global_decoder_symbol
                        s = "get_global_decoder_symbol - Public - Stable"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                        i = i + 1
                'returning DO Statements
                    'DO header
                        s = "__________Project Object DO Statements__________"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                    'Run
                        s = DTS_V2A.run_V0(0, "Log_Report")
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                    'Run_unit_cost_refresh
                        s = DTS_V2A.Run_DTS_unit_cost_refresh_v0("Log_Report")
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                'returning Get Statements
                    'get header
                        s = "__________Project Object Get Statements__________"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                    'get
                        s = DTS_V2A.Get_V0(0, "Log_Report")
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                    'get_size_V0
                        s = DTS_V2A.get_size_V0("Log_Report")
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                'returning Utility Statements
                    'get header
                        s = "__________Project Object Utility Statements__________"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                    'Check_DTS_Table_V0_01A
                        s = Check_DTS_Table_V0_01A("Log_Report")
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                    'check
                        s = "Check - Public - Unstable: needs updates has not specific code"
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
        
        Public Function run_V0(ByVal choice As DTS_run_choices_V1, Optional more_instructions As String) As Variant
            'check for log reporting
                If (more_instructions = "Log_Report") Then
                    run_V0 = "run_v0 - Public - Need Log report 10/28/20"
                    Exit Function
                End If
            'code start
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_vX.run_V0... Starting...")
                'define variables
                    Dim condition As Boolean
                    Dim i As Long
                'setup variables
                    'na
                'start check
                    MsgBox ("'dts_vx_dev.run' need to add boots check insted of the one used on dts as it can then use a standard check for a page exist.")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    condition = DTS_V2A.check(True)
                'check for check pass and if so run command else throw error
                    Call Boots_Report_v_Alpha.Log_Push(text, "Checking if conditions are met to run any actions...")
                    If (condition = True) Then
                        Call Boots_Report_v_Alpha.Log_Push(text, "Passed!...")
                        Select Case choice
                            Case DTS_update_unit_cost
                                DTS_V2A.Run_DTS_unit_cost_refresh_v0 ("lS2bjzvsk4BmFl5vpN3W")
                        End Select
                    Else
                        Call Boots_Report_v_Alpha.Log_Push(text, "Failed!...")
                        GoTo Dts_error_run_check_not_passed
                    End If
            'code end
                run_V0 = True
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_vX.run_V0... Finished...")
                Exit Function
            'error handling
Dts_error_run_check_not_passed:
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                    Call Boots_Report_v_Alpha.Log_Push(Flag)
                        Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR:...")
                        Call Boots_Report_v_Alpha.Log_Push(text, "Subroutine procedures set to check if operations could be complete failed required checks please see the log...")
                    Call Boots_Report_v_Alpha.Log_Push(table_close, "")
                    Call Boots_Report_v_Alpha.Log_Push(text, "__________CRASH!____________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "__________CRASH!____________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "__________CRASH!____________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "__________CRASH!____________________")
                Call Boots_Report_v_Alpha.Log_Push(table_close, "")
                Call Boots_Report_v_Alpha.Log_Push(text, ".")
                For i = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                Next i
                Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                End

        End Function
        
Private Function Run_DTS_unit_cost_refresh_v0(Optional more_instructions As String) As Variant
    'Created By (Zachary Daugherty)(8/25/20)
    'Purpose Case & notes:
        If (more_instructions = "help") Then
            Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "DTS_v2A.Run_DTS_unit_cost_refresh_v0: Help File Triggered..." & Chr(13) & _
                "What is this function for? |updated:12-09-2020|:" & Chr(13) & _
                "    Run_DTS_unit_cost_refresh_v0 is a private function called when all other checks are complete called from check functions WARNING!" & Chr(13) & _
                "        this function should not be called directly always use the run function that is designated to the module. this is because" & Chr(13) & _
                "        this function has no logic for proper procedure checks this is just the update function for the unit cost function. meaning" & Chr(13) & _
                "        if called improperly stuff can get overwritten inproperly." & Chr(13) & _
                "Should i call this function directly? |updated:12-09-2020|:" & Chr(13) & _
                "    see question above." & Chr(13) & _
                "What is returned from this function? |updated:12-09-2020|:" & Chr(13) & _
                "    if function completes properly, true will be returned to prove that it was run correctly" & Chr(13) & _
                "Listing off dependants of function |updated:12-09-2020|:" & Chr(13) & _
                "    DTS_V2A." & Chr(13) & _
                "        DTS_V2A.|parent module|" & Chr(13) & _
                "    Boots_Main_V_alpha." & Chr(13) & _
                "        Boots_Main_V_alpha.get_sheet_list" & Chr(13) & "    Boots_Report_v_Alpha." & Chr(13) & "        Boots_Report_v_Alpha.Log_Push" & Chr(13) & "        Boots_Report_v_Alpha.Push_notification_message" & Chr(13) & _
                "    matrix_V2." & Chr(13) & _
                "        matrix_V2.matrix_dimensions_v1" & Chr(13) & _
                "    SP_V1_DEV." & Chr(13) & _
                "        SP_V1_DEV.get_size_A" & Chr(13) & "        SP_V1_DEV.get_size_B_V1" & Chr(13) & "        SP_V1_DEV.get_sheet_name" & Chr(13) & _
                "    String_V1." & Chr(13) & _
                "        String_V1.Disassociate_by_Char_V2" & Chr(13) & "        String_V1.is_same_V1")
            Exit Function
        End If
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            Run_DTS_unit_cost_refresh_v0 = "Run_DTS_unit_cost_refresh_v0 - Private - stable 12-09-2020 - Help file:Y"
            Exit Function
        End If
    'code start
            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_vX.Run_DTS_unit_cost_refresh_v0... Starting...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
        'call protection
            If (more_instructions <> "lS2bjzvsk4BmFl5vpN3W") Then 'random string
                Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "DTS_v2A.Run_DTS_unit_cost_refresh_v0: Call Protection triggered..." & Chr(13) & _
                "DTS_v2A.Run_DTS_unit_cost_refresh_v0 did not have the correct key sent to run this command this is in place to prevent this function being called by accident" & Chr(13) & _
                "    if you did mean to call this function please enter the random string specified in the function in the more instructions field to bypass this protection...")
                End
            End If
        'define variables
            'positional
                Dim wb As Workbook              'cursor position
                Dim home_pos As Worksheet       'returns to this position post subroutine
                Dim current_sht As Worksheet    'cursor position
                Dim row As Long                 'cursor position
                Dim col As Long                 'cursor position
                Dim line As Long
            'memory management
                Dim Memory_Main() As String
                Dim SP_decoder_A() As String
                Dim SP_decoder_B() As String
                Dim Lookup() As String
                Dim size_of_dts As Long
                Dim size_of_sp_A As Long
                Dim size_of_sp_B As Long
            'globals
                Dim DTS_Inflation_value As Double
                Dim SP_GLOBAL_STRUCTURAL_value As Double
                Dim SP_GLOBAL_plate_value As Double
            'containers
                Dim L As Long
                Dim L_2 As Long
                Dim s As String
                Dim condition As Boolean
                Dim error As String
                Dim anti_loop As Long
        'setup variables
            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Setting up variables... Start...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'setup pos
                Set wb = ActiveWorkbook
                Set home_pos = ActiveSheet
                On Error GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_find_SHEET  'goto error handler
                    error = "DTS"
                    Set current_sht = wb.Sheets(error)                      'setting name
                    error = ""
                On Error GoTo 0                                             'returns error handler to default
                row = -1
                col = -1
                s = "empty"
                L = -1
                L_2 = -1
                SP_GLOBAL_STRUCTURAL_value = -1
                SP_GLOBAL_plate_value = -1
                DTS_Inflation_value = -1
                size_of_dts = -1
                size_of_sp_A = -1
                size_of_sp_B = -1
            'setup array tables
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Array table setup... start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'check for valid sizes and get sizes
                    'get size of DTS
                        Call Boots_Report_v_Alpha.Log_Push(text, "Fetching size of DTS TABLE...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        size_of_dts = DTS_V2A.get_size_V0()
                    'get size of steel presets A
                        Call Boots_Report_v_Alpha.Log_Push(text, "Fetching size of Steel Presets Table A...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        size_of_sp_A = SP_V1_DEV.get_size_A
                    'get size of steel presets B
                        Call Boots_Report_v_Alpha.Log_Push(text, "Fetching size of Steel Presets Table B...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        size_of_sp_B = SP_V1_DEV.get_size_B_V1
                'initialize arrays
                    Call Boots_Report_v_Alpha.Log_Push(text, "Initalizing Storage for Tables: Memory Main, Sp_decoder_A , Sp_decoder_B, Lookup...")
                    ReDim Memory_Main(0 To size_of_dts, 0 To DTS_POS_2A.DTS_Q_number_of_tracked_locations)
                    ReDim SP_decoder_A(0 To size_of_sp_A, 1 To SP_POS.SP_Q_Number_total_A_Tracked_Cells)
                    ReDim SP_decoder_B(0 To size_of_sp_B, 1 To SP_POS.SP_Q_Number_total_B_Tracked_Cells)
                    'lookup setup
                        ReDim Lookup(0 To size_of_sp_A + size_of_sp_B, 0 To 3)
                        'address assignment
                            Lookup(0, 0) = "Lookup Code"
                            Lookup(0, 1) = "Sheet its on"
                            Lookup(0, 2) = "What table its on"
                            Lookup(0, 3) = "address in array"
                'end of array tables
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Array table setup... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'setup globals
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Setting Up Global values... start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'dts
                    Call Boots_Report_v_Alpha.Log_Push(text, "Get DTS Sheet... Inflation Values")
                    On Error GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_find_SHEET
                        error = "DTS"
                        Set current_sht = wb.Sheets(error)
                        error = ""
                    On Error GoTo 0
                    DTS_Inflation_value = current_sht.Cells(DTS_POS_2A.DTS_I_Inflation_Const_row, DTS_POS_2A.DTS_I_Inflation_Const_col).value
                'SP
                    Call Boots_Report_v_Alpha.Log_Push(text, "Get SP Sheet... Structural value & Plate value")
                    On Error GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_find_SHEET
                        error = "STEEL PRESETS"
                        Set current_sht = wb.Sheets(error)
                        error = ""
                    On Error GoTo 0
                    SP_GLOBAL_STRUCTURAL_value = current_sht.Cells(SP_POS.SP_I_Const_Structural_row, SP_POS.SP_I_Const_Structural_col).value
                    SP_GLOBAL_plate_value = current_sht.Cells(SP_POS.SP_I_Const_Plate_row, SP_POS.SP_I_Const_Plate_col).value
                'return cursor to home
                    home_pos.Activate
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Setting Up Global values... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'setup complete
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Setting Up Variables.. Finished")
        'load tables
            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Load table Info... start")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'dts main
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Loading DTS Information to array... Start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'set focus
                    On Error GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_find_SHEET
                        error = "DTS"
                            Set current_sht = wb.Sheets(error)
                        error = ""
                        'set row col
                            row = DTS_POS_2A.DTS_I_part_number_row
                            col = DTS_POS_2A.DTS_I_part_number_col

                    On Error GoTo 0
                'get
                    For L = 0 To size_of_dts
                        For L_2 = 1 To DTS_POS_2A.DTS_Q_number_of_tracked_locations
                            Memory_Main(L, L_2) = current_sht.Cells(row, col).value
                            col = DTS_POS_2A.DTS_I_part_number_col + L_2
                        Next L_2
                        row = DTS_POS_2A.DTS_I_part_number_row + L + 1
                        col = DTS_POS_2A.DTS_I_part_number_col
                    Next L
                'cleanup
                    row = -1
                    col = -1
                    L = -1
                    L_2 = -1
                    Set current_sht = Nothing
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Loading DTS Information to array... Finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'SP_DECODER_A
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Loading Steel Presets A Information to array... Start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'set focus
                    On Error GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_find_SHEET
                        error = "Steel Presets"
                            Set current_sht = wb.Sheets(error)
                        error = ""
                        'set row col
                            row = SP_POS.SP_I_A_Prefix_row
                            col = SP_POS.SP_I_A_Prefix_col
                    On Error GoTo 0
                'get
                    For L = 0 To size_of_sp_A
                        For L_2 = 1 To SP_POS.SP_Q_Number_total_A_Tracked_Cells
                            SP_decoder_A(L, L_2) = current_sht.Cells(row, col).value
                            col = SP_POS.SP_I_A_Prefix_col + L_2
                        Next L_2
                    row = SP_POS.SP_I_A_Prefix_row + L + 1
                    col = SP_POS.SP_I_A_Prefix_col
                Next L
                'cleanup
                    row = -1
                    col = -1
                    L = -1
                    L_2 = -1
                    Set current_sht = Nothing
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Loading Steel Presets A Information to array... Finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'SP_DECODER_B
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Loading Steel Presets B Information to array... Start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'set focus
                    On Error GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_find_SHEET
                        error = "Steel Presets"
                            Set current_sht = wb.Sheets(error)
                        error = ""
                        'set row col
                            row = SP_POS.SP_I_B_Prefix_row
                            col = SP_POS.SP_I_B_Prefix_col
                    On Error GoTo 0
                'get
                    For L = 0 To size_of_sp_B
                        For L_2 = 1 To SP_POS.SP_Q_Number_total_B_Tracked_Cells
                            SP_decoder_B(L, L_2) = current_sht.Cells(row, col).value
                            col = SP_POS.SP_I_B_Prefix_col + L_2
                        Next L_2
                    row = SP_POS.SP_I_B_Prefix_row + L + 1
                    col = SP_POS.SP_I_B_Prefix_col
                    Next L
                'cleanup
                    row = -1
                    col = -1
                    L = -1
                    L_2 = -1
                    Set current_sht = Nothing
                    home_pos.Activate
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Loading Steel Presets B Information to array... Finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'assemble lookup table
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: assemble lookup table Information to array... Start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'initialize variables
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: assemble lookup table initialize variables... Start")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    L = 0
                    L_2 = 1
                    line = 0
                    s = ""
                    'Fetching lookup array dimension information
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Fetching lookup array dimension information... matrix_V2.matrix_dimensions_v1... start")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = matrix_V2.matrix_dimensions_v1(Lookup(), "d_report")
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Fetching lookup array dimension information... matrix_V2.matrix_dimensions_v1... finish")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    'parse matrix dim
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: resolve dimension information... start")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        Call Boots_Report_v_Alpha.Log_Push(text, "resolve dimension information (step1-3).... String_V1.Disassociate_by_Char_V2 running...")
                            s = String_V1.Disassociate_by_Char_V2("<", s, Right_C, "d_report")
                        Call Boots_Report_v_Alpha.Log_Push(text, "resolve dimension information (step2-3).... String_V1.Disassociate_by_Char_V2 running...")
                            s = String_V1.Disassociate_by_Char_V2("<", s, Right_C, "d_report")
                        Call Boots_Report_v_Alpha.Log_Push(text, "resolve dimension information (step3-3).... String_V1.Disassociate_by_Char_V2 running...")
                            line = CLng(String_V1.Disassociate_by_Char_V2(">", s, Left_C, "d_report"))
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: resolve dimension information... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    'cleanup
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: assemble lookup table initialize variables... Finish")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                'Filling lookup table with all possible values and table positional data...
DTS_incoding_of_table_names:
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Filling lookup table with all possible values and table positional data... start")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    For L = 0 To line
                        'check to see where table data should be grabed from, note will fill in table A first then B
                            'section for table a
                                If (L > 0) Then
                                    If (L <= size_of_sp_A - 1) Then '-1 is included on the end to skip the goalpost of the table
                                        Lookup(L, 0) = SP_decoder_A(L, 1)
                                                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: fetching sheet name: SP_V1_DEV.get_sheet_name start")
                                                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                                    Lookup(L, 1) = SP_V1_DEV.get_sheet_name("d_report")
                                                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: fetching sheet name: SP_V1_DEV.get_sheet_name finish")
                                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                        Lookup(L, 2) = "A"
                                        Lookup(L, 3) = L
                                    Else
                                        'fall through marker
                                            Lookup(L, 0) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger
                                            Lookup(L, 1) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger
                                            Lookup(L, 2) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger
                                            Lookup(L, 3) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger
                                    End If
                                End If
                            'section for table b
                                If (L > 0) Then
                                    If ((L > size_of_sp_A) And (L < size_of_sp_A + size_of_sp_B)) Then
                                        L_2 = L - size_of_sp_A
                                        Lookup(L, 0) = SP_decoder_B(L_2, 1)
                                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: fetching sheet name: SP_V1_DEV.get_sheet_name start")
                                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                            Lookup(L, 1) = SP_V1_DEV.get_sheet_name("d_report")
                                            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: fetching sheet name: SP_V1_DEV.get_sheet_name finish")
                                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                        Lookup(L, 2) = "B"
                                        Lookup(L, 3) = L_2
                                    Else
                                        'fall through marker
                                            If (L > size_of_sp_A) Then
                                                Lookup(L, 0) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger
                                                Lookup(L, 1) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger
                                                Lookup(L, 2) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger
                                                Lookup(L, 3) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger
                                            End If
                                    End If
                                End If
                    Next L
                'cleanup
                    L = -1
                    L_2 = -1
                    s = "empty"
                    line = -1
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Filling lookup table with all possible values and table positional data... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: assemble lookup table Information to array... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Load table Info... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'check lookup table for duplicate entrys
            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: check lookup table for duplicate entrys... start") '2-3
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'initialize variable
                L = 0
                L_2 = 0
                s = ""
                condition = False
            'start
                'Get size of the matrix
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix... start") '3-4
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: matrix_V2.matrix_dimensions_v1... |1-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = matrix_V2.matrix_dimensions_v1(Lookup(), "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: matrix_V2.matrix_dimensions_v1... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... |2-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = String_V1.Disassociate_by_Char_V2(">", s, Right_C, "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... |3-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = String_V1.Disassociate_by_Char_V2(">", s, Right_C, "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... |4-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = String_V1.Disassociate_by_Char_V2("<", s, Right_C, "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                            
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... |5-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            line = String_V1.Disassociate_by_Char_V2(">", s, Left_C, "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    'cleanup
                        s = "Empty"
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: Get size of the matrix... finish") '4-3
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        
                'loop through lookup table
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: loop through lookup table for duplicate entrys... start") '3-4
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    
                    For L = 1 To (line - 1)
                        'value to search by
                            'if set to empty skip
                                If (Lookup(L, 1) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger) Then
                                    GoTo Run_DTS_unit_cost_refresh_v0_ignore_entry
                                End If
                        'compair against all other entrys
                            For L_2 = 1 To (line - 1)
                                'check to see if index is the same if so skip check
                                    If (L = L_2) Then
                                        GoTo Run_DTS_unit_cost_refresh_v0_skip_check
                                    End If
                                'set value to check thru
                                    s = Lookup(L, 0)
                                'do check
                                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: loop through lookup table for duplicate entrys match possible: String_V1.is_same_V1... start") '4-5
                                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                        condition = String_V1.is_same_V1(s, Lookup(L_2, 0), "d_report")
                                        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: loop through lookup table for duplicate entrys match possible: String_V1.is_same_V1... finish") '5-4
                                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                'if condition is true then throw error
                                    If (condition = True) Then
                                        error = "in array: lookup:(" & L_2 & ",0) value:'" & Lookup(L_2, 0) & "'. is the same as the value in: lookupL(" & L & ",0)"
                                        GoTo Run_DTS_unit_cost_refresh_v0_duplicate_lookups
                                    Else
                                        
                                    End If
                                'goto
Run_DTS_unit_cost_refresh_v0_skip_check:
                            Next L_2
                        'goto
Run_DTS_unit_cost_refresh_v0_ignore_entry:
                    Next L
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: loop through lookup table for duplicate entrys... finish") '4-3
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'cleanup
                L = -1
                L_2 = -1
                s = "empty"
                line = -1
                condition = False
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.Run_DTS_unit_cost_refresh_v0: check lookup table for duplicate entrys... finish") '3-2
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'do update
            Call Boots_Report_v_Alpha.Log_Push(text, "DTS_vX.Run_DTS_unit_cost_refresh_v0: run update Starting...") '2-3
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'initalize variables
                On Error GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_find_SHEET
                    error = "DTS"
                    Set current_sht = wb.Sheets(error)
                    error = ""
                On Error GoTo 0
                row = DTS_POS_2A.DTS_I_part_number_row
                col = DTS_POS_2A.DTS_I_part_number_col
                L = 0
                L_2 = 0
            'iterate thru memory main
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_vX.Run_DTS_unit_cost_refresh_v0: run update: iterating thru memory main to find matches and return values...") '3-na
                For L = 1 To size_of_dts
                    'change pos
                        row = DTS_POS_2A.DTS_I_part_number_row + L
                    'set smart code
                        s = Memory_Main(L, 2)
                    'check for empty or ignore trigger
                        If ((s <> DTS_V2A.get_global_unit_cost_refresh_ignore_trigger) And (s <> "")) Then
                            'decode smart code
Run_DTS_unit_cost_refresh_v0_part_numb_check:
                                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_vX.Run_DTS_unit_cost_refresh_v0: run update: String_V1.Disassociate_by_Char_V2 start...") '3-4
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                    s = String_V1.Disassociate_by_Char_V2(DTS_V2A.get_global_decoder_symbol, s, Left_C, "d_report")
                                    Call Boots_Report_v_Alpha.Log_Push(text, "DTS_vX.Run_DTS_unit_cost_refresh_v0: run update: String_V1.Disassociate_by_Char_V2 finish...") '4-3
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                            'search for key in lookup array
                                For L_2 = 1 To (size_of_sp_A + size_of_sp_B)
                                    If (s = Lookup(L_2, 0)) Then
                                        'match found return value to sheet
                                            'locate which chart
                                                If (Lookup(L_2, 2) = "A") Then
                                                    'match found in decode table 'A'
                                                        'return value to dts
                                                            On Error GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_find_range
                                                                error = "DTS_Unit_cost"
                                                                    current_sht.Range(error).Offset(L, 0).value = SP_decoder_A(CLng(Lookup(L_2, 3)), 4) 'paste to sheet name <current_sht> then move cursor to range <error> offset down to pos <L>: to get the value find in array <lookup> and return address of the match. convert to <long> variable and then user that long to look in array <sp_decoder_a> at the value of <long> then return that value to sheet
                                                                error = ""
                                                            On Error GoTo 0
                                                Else
                                                    If (Lookup(L_2, 2) = "B") Then
                                                        'match found in decode table 'B'
                                                            'return value to dts
                                                                On Error GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_find_range
                                                                    error = "DTS_Unit_cost"
                                                                    current_sht.Range(error).Offset(L, 0).value = SP_decoder_B(CLng(Lookup(L_2, 3)), 4) 'paste to sheet name <current_sht> then move cursor to range <error> offset down to pos <L>: to get the value find in array <lookup> and return address of the match. convert to <long> variable and then user that long to look in array <sp_decoder_a> at the value of <long> then return that value to sheet
                                                                    error = ""
                                                                On Error GoTo 0
                                                    Else
                                                        error = CStr(Lookup(L_2, 2))
                                                        GoTo dts_Run_DTS_unit_cost_refresh_v0_cant_locate_table
                                                    End If
                                                End If
                                    End If
                                Next L_2
                            'fall through statement Smart code not found
                        Else
                            'check for non aka code
                                If (condition = False) Then
                                    condition = True
                                    s = Memory_Main(L, 1)
                                    anti_loop = anti_loop + 1
                                    If (anti_loop < 6) Then
                                        GoTo Run_DTS_unit_cost_refresh_v0_part_numb_check
                                    Else
                                        MsgBox ("anti loop triggered please check code")
                                        Stop
                                    End If
                                End If
                        End If
                        'reset check
                            condition = False
                            anti_loop = 0
                Next L
            'cleanup
                Set current_sht = Nothing
                row = -1
                col = -1
                L = -1
                s = "empty"
                L_2 = -1
                home_pos.Activate
                Call Boots_Report_v_Alpha.Log_Push(text, "DTS_vX.Run_DTS_unit_cost_refresh_v0: run update finish...") '3-2
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    'code end
        Run_DTS_unit_cost_refresh_v0 = True
        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_vX.Run_DTS_unit_cost_refresh_v0... finish...") '2-1
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Exit Function
    'error handling
dts_Run_DTS_unit_cost_refresh_v0_cant_find_SHEET:
        'dts_Run_DTS_unit_cost_refresh_v0_cant_find_SHEET
            Run_DTS_unit_cost_refresh_v0 = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL Error: DTS_Vx: sub: Run_DTS_unit_cost_refresh_v0: was unable to find the sheet named '" & error & "', please check your code.")
                Call Boots_Report_v_Alpha.Log_Push(text, "Displaying Snapshot of Values:...")
                'table
                    Call Boots_Report_v_Alpha.Log_Push(table_open, "")
                    'push generated sheet list
                        Boots_Main_V_alpha.get_sheet_list
                        Call Boots_Report_v_Alpha.Log_Push(text, "Posting sheet manager list...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            For z = 1 To (wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).End(xlDown).row - wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).row)
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Sheet List: " & wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row + z, boots_pos.p_sheet_name_col).value & _
                                    " ='visible stat': " & wb.Sheets("boots").Cells(boots_pos.p_sheet_visible_status_row + z, boots_pos.p_sheet_visible_status_col).value)
                            Next z
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                    'adresses
                        Call Boots_Report_v_Alpha.Log_Push(text, "Posting address mannager...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            'home_pos
                                If home_pos Is Nothing Then
                                    Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: 'NOTHING' as worksheet")
                                Else
                                    Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: '" & home_pos.Parent.path & "\" & home_pos.Parent.Name & " == " & home_pos.index & ": " & home_pos.Name & "' as worksheet")
                                End If
                            'wb
                                If wb Is Nothing Then
                                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Wb: 'NOTHING' as workbook")
                                Else
                                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Wb: '" & wb.path & "\" & wb.Name & "' as workbook")
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
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    'TABLE CLOSE
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'end procedure
                        For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                        Next z
                    'call end statement
                        Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                        End
Run_DTS_unit_cost_refresh_v0_duplicate_lookups:
        'Run_DTS_unit_cost_refresh_v0_duplicate_lookups
            Run_DTS_unit_cost_refresh_v0 = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL Error: DTS_Vx: Function: Run_DTS_unit_cost_refresh_v0:During the assembly '" & error & "' please make the nessasary changes to the tables to not have duplicate values")
                'table
                     Call Boots_Report_v_Alpha.Log_Push(text, "Showing all values of lookup...")
                    Call Boots_Report_v_Alpha.Log_Push(table_open, "")
                        For z = 0 To (UBound(Lookup(), 1) - LBound(Lookup(), 1))
                             Call Boots_Report_v_Alpha.Log_Push(Variable, "<" & Lookup(z, 0) & "><" & Lookup(z, 1) & "><" & Lookup(z, 2) & "><" & Lookup(z, 3) & ">")
                        Next z
                    'table close
                        Call Boots_Report_v_Alpha.Log_Push(table_close)
                    Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'end procedure
                        For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                        Next z
                    'call end statement
                        Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                        End
dts_Run_DTS_unit_cost_refresh_v0_cant_find_range:
        'dts_Run_DTS_unit_cost_refresh_v0_cant_find_range:
            Run_DTS_unit_cost_refresh_v0 = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL Error: Dts_vx: Function: Run_DTS_unit_cost_refresh_v0: Range(" & error & ") was unable to be located")
                'table
                    Call Boots_Report_v_Alpha.Log_Push(table_open, "")
                        Boots_Main_V_alpha.get_sheet_list
                    'sheets
                        Call Boots_Report_v_Alpha.Log_Push(text, "Displaying Snapshot of Sheets:...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            For z = 1 To (wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).End(xlDown).row - wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).row)
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "Sheet List: " & wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row + z, boots_pos.p_sheet_name_col).value & _
                                    " ='visible stat': " & wb.Sheets("boots").Cells(boots_pos.p_sheet_visible_status_row + z, boots_pos.p_sheet_visible_status_col).value)
                            Next z
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Call Boots_Report_v_Alpha.Log_Push(text, "...")
                    'addresses
                        Call Boots_Report_v_Alpha.Log_Push(text, "Displaying Snapshot of Addresses:...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
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
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Call Boots_Report_v_Alpha.Log_Push(text, "...")
                'close table
                    Call Boots_Report_v_Alpha.Log_Push(table_close)
                    Call Boots_Report_v_Alpha.Log_Push(table_close)
                     'end procedure
                        For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                        Next z
                    'call end statement
                        Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                        End
dts_Run_DTS_unit_cost_refresh_v0_cant_locate_table:
        'dts_Run_DTS_unit_cost_refresh_v0_cant_locate_table:
            Run_DTS_unit_cost_refresh_v0 = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL Error: Dts_vx: Function: Run_DTS_unit_cost_refresh_v0:(see next line)")
                Call Boots_Report_v_Alpha.Log_Push(text, "Function was unable to locate the table named:'" & error & "'(see next line)")
                Call Boots_Report_v_Alpha.Log_Push(text, "Please see the goto 'DTS_incoding_of_table_names' as this is where the table chars are assigned")
            'table
                Call Boots_Report_v_Alpha.Log_Push(text, "Displaying Snapshot of addresses:...")
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
                'table close
                    Call Boots_Report_v_Alpha.Log_Push(table_close)
                    Call Boots_Report_v_Alpha.Log_Push(table_close)
                    'end procedure
                        For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                        Next z
                    'call end statement
                        Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                        End
End Function



'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

                                                                        'Get Statements
                                                                        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Public Function Get_V0(ByVal operation As get_choices_v1, Optional more_instructions As String) As Variant
'currently functional as of (10/14/20) checked by: (Zachary Daugherty)
    'Created By (Zachary Daugherty)(10/14/20)
    'Purpose Case & notes:
        'select get options from menu
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
        'varies
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            Get_V0 = "Get_V0 - Public - Need Log report 10/28/20"
            Exit Function
        End If
    'code start
        'find selected option
            MsgBox ("need to add notation for fallthrough statement")
            Stop
            Select Case operation
                Case get_choices_v1.get_size
                    Get_V0 = DTS_V2A.get_size_V0
                    GoTo get_v0_exit
                Case Else
                    Stop
            End Select
    'cleanup
get_v0_exit:
End Function

Private Function get_size_V0(Optional more_instructions As String) As Variant
'currently functional as of (9/2/2020) checked by: (Zachary Daugherty)
    'Created By (Zachary Daugherty)(9/2/2020)
    'Purpose Case & notes:
        'returns the size of the DTS Table rows
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
        'gives size of the table as long
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            get_size_V0 = "get_size_V0 - Private - Stable 10/28/20"
            Exit Function
        End If
        
        
        
        '<debug note>
            Call Boots_Report_v_Alpha.Log_Push(text, "-------------------------------------------------------------------------")
                Call Boots_Report_v_Alpha.Log_Push(text, "The Calling of 'DTS_V2A.get_size_V0' is not properly setup for dev notes yet please fix: missing error reporting")
                Call Boots_Report_v_Alpha.Log_Push(text, "The Calling of 'DTS_V2A.get_size_V0' is not properly setup for dev notes yet please fix: missing error reporting")
                Call Boots_Report_v_Alpha.Log_Push(text, "The Calling of 'DTS_V2A.get_size_V0' is not properly setup for dev notes yet please fix: missing error reporting")
                Call Boots_Report_v_Alpha.Log_Push(text, "The Calling of 'DTS_V2A.get_size_V0' is not properly setup for dev notes yet please fix: missing error reporting")
            Call Boots_Report_v_Alpha.Log_Push(text, "-------------------------------------------------------------------------")
        '<end of debug note>
        
        
        
    'code start
        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.get_size_v0 Start...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
        'define variables
            'positional
                Dim wb As Workbook              'cursor position
                Dim home_pos As Worksheet       'returns to this position post subroutine
                Dim current_sht As Worksheet    'cursor position
                Dim row As Long                 'cursor position
                Dim col As Long                 'cursor position
            'memory
                Dim arr() As String                         'storage of data
                Dim delete_empty_rows_condition As Boolean  'conditional check
                Dim anti_loop As Long                       'protections for looping operations: protects against infinite loops
            'const
                Dim dist_to_goalpost As Long                'dist to goal from top of the table
            'containers
                Dim i As Long       'int storage 1
                Dim i_2 As Long     'int sotrage 2
                Dim s As String     'string storage
        'restart trigger
get_size_V0_restart:               'goto flag
        'setup variables
            Call Boots_Report_v_Alpha.Log_Push(text, "Setting Up Variables... ")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            Set wb = ActiveWorkbook
            Set home_pos = ActiveSheet
            On Error GoTo dts_get_cant_find_DTS_SHEET   'goto error handler
                Set current_sht = wb.Sheets("DTS")      'setting name
            On Error GoTo 0                             'returns error handler to default
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
        'move to start location
            Call Boots_Report_v_Alpha.Log_Push(text, "Moving to Starting position...")
            row = DTS_POS_2A.DTS_I_part_number_row     'fetching indexed information from enumeration
            col = DTS_POS_2A.DTS_I_part_number_col     'fetching indexed information from enumeration
            s = current_sht.Cells(row, col).value   'fetching indexed information from enumeration
        'get lenght to bottom
            Call Boots_Report_v_Alpha.Log_Push(text, "Fetching the length to the botom of the table...")
            On Error GoTo dts_cant_find_goalpost                    'goto error handler
                dist_to_goalpost = Range("DTS_GOALPOST").row - row  'setting definition
            On Error GoTo 0                                         'returns error handler to default
        'setup arr
            Call Boots_Report_v_Alpha.Log_Push(text, "Defining arr arry size...")
            ReDim arr(1 To dist_to_goalpost)        'defining dimensions for the array
                'arr memory guide
                    'arr(<X:true if cell empty, false if not>)

        'all rows with data if there is no data in cell location mark for removal
            Call Boots_Report_v_Alpha.Log_Push(text, "Removal of blank space... Start...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            Call Boots_Report_v_Alpha.Log_Push(text, "____Discription: program to check the part number field first then check other cols if empty and will only display in the log if empty____")
            For i = 1 To dist_to_goalpost
                row = row + 1
                s = current_sht.Cells(row, col)
                If (s = "") Then    'check the other cols in the row to see if any data is stored.
                    For i_2 = 1 To DTS_POS_2A.DTS_Q_number_of_tracked_locations - 1
                        s = current_sht.Cells(row, col + 1)
                        If (s <> "") Then
                            arr(i) = False  'row does not need to be deleted as it has values in fields other than part number
                            Exit For
                        End If
                        delete_empty_rows_condition = True
                        arr(i) = True   'row is entirly empty so mark for delete
                    Next i_2
                    If (arr(i) = True) Then
                        Call Boots_Report_v_Alpha.Log_Push(text, "Row:'" & row & "' part number field was empty checking looking in other cols...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        Call Boots_Report_v_Alpha.Log_Push(text, "Row:'" & row & "' has NO information stored inside mark for removal...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                    End If
                Else
                    arr(i) = False  'row does not need to be deleted
                End If
            Next i
            'cleanup
                i = -1
                i_2 = -1
                s = "empty"
                Call Boots_Report_v_Alpha.Log_Push(text, "Removal of blank space... Finished...")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
        'check for delete condition to be true\
            If (delete_empty_rows_condition = True) Then
                Call Boots_Report_v_Alpha.Log_Push(text, "Removing Blank space from the DTS page... Start...")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'hide updating
                    Application.ScreenUpdating = False
                    Application.DisplayAlerts = False
                'move to start location
                    Call Boots_Report_v_Alpha.Log_Push(text, "Moving to starting location...")
                    row = DTS_POS_2A.DTS_I_part_number_row
                    col = DTS_POS_2A.DTS_I_part_number_col
                    s = current_sht.Cells(row, col).value
                'iterate through to find empty then delete by moving everything up eliminating the blank space
                    Call Boots_Report_v_Alpha.Log_Push(text, "Loop through the rows and delete the marked locations... Start...")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    For i = 1 To dist_to_goalpost
                        row = row + 1
                        s = current_sht.Cells(row, col).value
                        If (arr(i) = "True") Then
                            Call Boots_Report_v_Alpha.Log_Push(text, "Deleting Row:'" & row & "' empty...")
                            'setup
                                On Error GoTo dts_get_cant_find_DTS_SHEET
                                    Set current_sht = wb.Sheets("DTS")
                                On Error GoTo 0
                                If ((current_sht.visible = xlSheetVeryHidden) Or (current_sht.visible = xlSheetHidden)) Then
                                    i_2 = current_sht.visible
                                    current_sht.visible = xlSheetVisible
                                Else
                                    i_2 = -1
                                End If
                                s = CStr(row) & ":" & CStr(row)
                                current_sht.Activate
                                Cells(row, col).Select
                            'delete row
                                Rows(s).Select
                                Selection.Delete Shift:=xlUp
                                i = i + 1
                        End If
                    Next i
                    Call Boots_Report_v_Alpha.Log_Push(text, "Loop through the rows and delete the marked locations... Finish...")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                'restart if things were deleted, reset some variables then goto 'get_size_V0_restart'
                    
                    Call Boots_Report_v_Alpha.Push_notification_message("DEVNOTE:'DTS_V2.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    
                    Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                    Call Boots_Report_v_Alpha.Log_Push(text, "DEVNOTE:'DTS_V2.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    Call Boots_Report_v_Alpha.Log_Push(text, "DEVNOTE:'DTS_V2.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    Call Boots_Report_v_Alpha.Log_Push(text, "DEVNOTE:'DTS_V2.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    Call Boots_Report_v_Alpha.Log_Push(text, "DEVNOTE:'DTS_V2.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    Call Boots_Report_v_Alpha.Log_Push(table_close, "")
                    
ActiveWorkbook.Sheets("LOG_" & Boots_Main_V_alpha.get_username).visible = -1
Stop 'error code test for indent see inside the if statement
                    
                    If (delete_empty_rows_condition = True) Then
                        current_sht.visible = i_2
                        If (anti_loop < 30) Then
                            Call Boots_Report_v_Alpha.Log_Push(text, "restarting Some checks becasuse rows were deleted...")
                            anti_loop = anti_loop + 1
                            'reset variables
                                home_pos.Activate
                                Set home_pos = Nothing
                                Set wb = Nothing
                                Set current_sht = Nothing
                                row = -1
                                col = -1
                                dist_to_goalpost = -1
                                i = -1
                                i_2 = -1
                                s = "empty"
                                ReDim arr(0)
                                delete_empty_rows_condition = False
                            'do goto
                                Call Boots_Report_v_Alpha.Log_Push(text, "Restart Procedure selected...")
                                Call Boots_Report_v_Alpha.Log_Push(text, "Removing Blank space from the DTS page... Abandoned...")
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "") 'THIS IS HERE TO FIX INDENT FROM THE GOTO JUMP
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "") 'THIS IS HERE TO FIX INDENT FROM THE GOTO JUMP
                            
                            ActiveWorkbook.Sheets("LOG_" & Boots_Main_V_alpha.get_username).visible = -1
                            Stop 'error code test for indent
                            
                                GoTo get_size_V0_restart
                        Else
                            MsgBox ("FATAL ERROR: ANTI_LOOP Triggered check code")
                            Stop
                        End If
                    End If
                'unhide updating
                    Application.ScreenUpdating = True
                    Application.DisplayAlerts = True
                    Call Boots_Report_v_Alpha.Log_Push(text, "Removing Blank space from the DTS page... Finish...")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
            'cleanup
                Call Boots_Report_v_Alpha.Log_Push(text, "Removing Blank space from the DTS page... Finish...")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
            End If
            'get final size
                get_size_V0 = dist_to_goalpost
    'cleanup
        Call Boots_Report_v_Alpha.Log_Push(text, "DTS_V2A.get_size_v0 Finish...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
    'code end
        Exit Function
    'error handling
dts_get_cant_find_DTS_SHEET:
        'dts_get_cant_find_DTS_SHEET:
            'set error report
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                    Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: dts_vx.Get_size_v0 was unable to locate the DTS sheet please check the enviorment & Log...")
                Call Boots_Report_v_Alpha.Log_Push(table_close)
            'generate required information
                Call Boots_Report_v_Alpha.Log_Push(text, "generating sheet list...")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    Call Boots_Main_V_alpha.get_sheet_list
            'push generated information
                Call Boots_Report_v_Alpha.Log_Push(table_open, "")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    For z = 1 To (wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).End(xlDown).row - wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).row)
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "Sheet List: " & wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row + z, boots_pos.p_sheet_name_col).value & _
                            " ='visible stat': " & wb.Sheets("boots").Cells(boots_pos.p_sheet_visible_status_row + z, boots_pos.p_sheet_visible_status_col).value)
                    Next z
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
            'page break
                Call Boots_Report_v_Alpha.Log_Push(text, "-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_")
                Call Boots_Report_v_Alpha.Log_Push(text, "-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_")
            'list local variables
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
                'other
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "more_instructions: '" & more_instructions & "' as string")
            'prep for post
                'close error table
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                    Call Boots_Report_v_Alpha.Log_Push(table_close, "")
                'indent out
                    For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                    Next z
                'call end statement
                    Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                    End
dts_cant_find_goalpost:
        'dts_cant_find_goalpost
            'set error report
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                    Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: dts_vx.Get_size_v0 was unable to return / find the DTS page Goalpost...")
                    Call Boots_Report_v_Alpha.Log_Push(text, "Please check the Range mannager in the Workbook for its existance...")
                Call Boots_Report_v_Alpha.Log_Push(table_close, "")
            'listing local variables
                Call Boots_Report_v_Alpha.Log_Push(text, "Listing Snapshot of some variables...")
                Call Boots_Report_v_Alpha.Log_Push(table_open, "")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Wb: '" & wb.path & "\" & wb.Name & "' as workbook")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: '" & home_pos.Parent.path & "\" & home_pos.Parent.Name & " == " & home_pos.index & ": " & home_pos.Name & "' as worksheet")
                Call Boots_Report_v_Alpha.Log_Push(table_close, "")
            'prep for post
                'close error table
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                    Call Boots_Report_v_Alpha.Log_Push(table_close, "")
                'indent out
                    For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                    Next z
                'call end statement
                    Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                    End
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

                                                                        'Utility Statements
                                                                        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


Private Function Check_DTS_Table_V0_01A(Optional more_instructions As String) As Variant
        'currently functional as of (8/7/2020) checked by: (zdaugherty)
            'Created By (Zachary Daugherty)(8/6/2020)
            'Purpose Case & notes:
                'Check_DTS_TABLE is a function that SHOULD be called at the start of any GET OR SET operation on the DTS page.
                    'As when any operation is done to the data table it should first be verifyed that the current version of the
                    'program and the addressed cell locations agree on their position.
                'This is done by Calling POS ENUM and compairing the positional data indexed per what is expected.
            'Library Refrences required
                'Na
            'Modules Required
                'na
            'inputs
                'Internal:
                    'DTS_VX_PROD.POS
                'required:
                    'na
                'optional:
                    'if visually is set to true:
                        'will walk through each position with user able to visually check positions.
            'returned outputs
                'returns:
                    'true: if all positions match up
                    'false: if any positions do not match up
        'check for log reporting
                If (more_instructions = "Log_Report") Then
                    Check_DTS_Table_V0_01A = "Check_DTS_Table_V0_01A - Private - Need Log report 10/28/20"
                    Exit Function
                End If
        'code start
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
                    
                'breakout
                Set proj_wb = ActiveWorkbook
                
                On Error GoTo FATAL_ERROR_CHECK_DTS_SET_DTS_ENV 'set error handler
                    Set cursor_sheet = proj_wb.Sheets("DTS")
                On Error GoTo 0 'set error handler back to norm
                cursor_row = 1
                cursor_col = 1
            'setup arr
                'redefine size of the arr
                    ReDim arr(1 To DTS_POS_2A.DTS_Q_number_of_tracked_locations, 1 To 5) 'see line below for definitions
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
                        
                        'compair <part number> expected location
                            s = "DTS_Part_number"   'expected range name for search
                            i = i + 1               'iterate arr position from x to x + 1 in the array
                            On Error GoTo ERROR_FATAL_check_dts_range_error 'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                                Set ref_rng = Range(s)  'set range
                            On Error GoTo 0 'reset error handler
                                arr(i, 1) = CStr(ref_rng.row)       'get range row pos
                                arr(i, 2) = CStr(ref_rng.Column)    'get range col pos
                                arr(i, 3) = DTS_POS_2A.DTS_I_part_number_row    'get enum row pos
                                arr(i, 4) = DTS_POS_2A.DTS_I_part_number_col    'get enum col pos
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then   'compair rows to rows and cols to cols
                                    arr(i, 5) = s & ": " & True 'if true report text
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                                    condition = True    'if true at the end of the block throw error as there is a miss match
                                End If
                        'compair <DTS_AKA> expected location
                            s = "DTS_AKA"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_AKA_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_AKA_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Description> expected location
                            s = "DTS_Description"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Description_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Description_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Unit_cost> expected location
                            s = "DTS_Unit_cost"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Unit_cost_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Unit_cost_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Adjusted_unit_cost> expected location
                            s = "DTS_Adjusted_unit_cost"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Adjusted_unit_cost_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Adjusted_unit_cost_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Unit_weight> expected location
                            s = "DTS_Unit_weight"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Unit_weight_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Unit_weight_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Shop_origin> expected location
                            s = "DTS_Shop_origin"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Shop_origin_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Shop_origin_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Status> expected location
                            s = "DTS_Status"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Status_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Status_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Job_other_info> expected location
                            s = "DTS_Job_other_info"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Job_other_info_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Job_other_info_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Vendor_info> expected location
                            s = "DTS_Vendor_info"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Vendor_info_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Vendor_info_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Vendor_phone> expected location
                            s = "DTS_Vendor_phone"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Vendor_phone_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Vendor_phone_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Vendor_fax> expected location
                            s = "DTS_Vendor_fax"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Vendor_fax_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Vendor_fax_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Vendor_part_number> expected location
                            s = "DTS_Vendor_part_number"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Vendor_part_number_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Vendor_part_number_col
                                If ((arr(i, 1) = arr(i, 3)) And (arr(i, 2) = arr(i, 4))) Then
                                    arr(i, 5) = s & ": " & True
                                Else
                                    arr(i, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Inflation_Const> expected location
                            s = "DTS_Inflation_Const"
                            i = i + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(i, 1) = CStr(ref_rng.row)
                                arr(i, 2) = CStr(ref_rng.Column)
                                arr(i, 3) = DTS_POS_2A.DTS_I_Inflation_Const_row
                                arr(i, 4) = DTS_POS_2A.DTS_I_Inflation_Const_col
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
                        GoTo ERROR_CHECK_DTS_FAILED_POS_CHECK
                    End If
                'return true
                    Check_DTS_Table_V0_01A = True   'passed all checks
                    Exit Function
        'code end
        'error handle
ERROR_FATAL_check_dts_range_error:
            'ERROR_FATAL_check_dts_range_error:
            Check_DTS_Table_V0_01A = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: Displaying Snapshot of Values:...")
                Call Boots_Report_v_Alpha.Log_Push(table_open, "")
                    'other
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "Check_DTS_Table_V0_01A: '" & Check_DTS_Table_V0_01A & "' as variant")
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "More_instructions: '" & more_instructions & "' as string")
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "condition: '" & condition & "' as Boolean")
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "i: '" & i & "' as Long")
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "s: '" & s & "' as String")
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
                    'other
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_row: '" & cursor_row & "' as long")
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_col: '" & cursor_col & "' as long")
                    'ref rng
                        If ref_rng Is Nothing Then
                            Call Boots_Report_v_Alpha.Log_Push(Variable, "ref_rng: 'NOTHING' as range")
                        Else
                            Call Boots_Report_v_Alpha.Log_Push(Variable, "ref_rng: '" & ref_rng.Name & " value=" & ref_rng.value & "' as range")
                        End If
                    'other
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "arr: '<please check the array>' as string")
                Call Boots_Report_v_Alpha.Log_Push(table_close, "")
                Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                    Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: MODULE:(DTS_VX)FUNCTION:(CHECK_DTS_TABLE) UNABLE TO LOCATE THE SPECIFIED RANGE:<" & s & "> please check the name mannager for errors. fix and then re-run")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                Call Boots_Report_v_Alpha.Log_Push(table_close)
                For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                Next z
                Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                End
FATAL_ERROR_CHECK_DTS_SET_DTS_ENV:
            Call MsgBox("check dts_table using log replace", , "check dts_table using log")
            'Call __________.log(__________.get_username, "FATAL_ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) UNABLE TO FIND OR SET SHEET DTS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.")
            Call MsgBox("FATAL_ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) UNABLE TO FIND OR SET SHEET DTS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.", , "FATAL ERROR: SET DTS SHEET ENV")
            Stop
            Exit Function
ERROR_CHECK_DTS_FAILED_POS_CHECK:
            'ERROR_CHECK_DTS_FAILED_POS_CHECK
            Check_DTS_Table_V0_01A = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE...")
                Call Boots_Report_v_Alpha.Log_Push(text, "Displaying Snapshot of Values:...")
                Call Boots_Report_v_Alpha.Log_Push(table_open, "")
                'array values
                    Call Boots_Report_v_Alpha.Log_Push(text, "Showing Array table values:...")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        For z = 1 To (UBound(arr, 1) - LBound(arr, 1))
                            Call Boots_Report_v_Alpha.Log_Push(text, arr(z, 5))
                        Next z
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                'address information
                    Call Boots_Report_v_Alpha.Log_Push(text, "Showing Address information:...")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    'proj_wb
                        If proj_wb Is Nothing Then
                            Call Boots_Report_v_Alpha.Log_Push(Variable, "proj_wb: 'NOTHING' as workbook")
                        Else
                            Call Boots_Report_v_Alpha.Log_Push(Variable, "proj_wb: '" & proj_wb.path & "\" & proj_wb.Name & "' as worksheet")
                        End If
                    'cursor_sheet
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
                    'other
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_row: '" & cursor_row & "' as long")
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_col: '" & cursor_col & "' as long")
                        'range
                            Call Boots_Report_v_Alpha.Log_Push(text, "REF_RNG:...")
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "ref_rng: '" & ref_rng.Parent.Name & "' as Parent name")
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "ref_rng: '" & ref_rng.row & "' as row")
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "ref_rng: '" & ref_rng.Column & "' as column")
                                Call Boots_Report_v_Alpha.Log_Push(Variable, "ref_rng: '" & ref_rng.value & "' as value")
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'close table
                Call Boots_Report_v_Alpha.Log_Push(table_close)
                Call Boots_Report_v_Alpha.Log_Push(table_close)
            'end
                For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                Next z
                Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                End
                
                
                
                
                
            Call MsgBox("check dts_table using log replace", , "check dts_table using log")
            
            Call MsgBox("ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: " & vbCrLf & vbCrLf & arr(1, 5) & vbCrLf & arr(2, 5) & vbCrLf & arr(3, 5) & vbCrLf & arr(4, 5) & vbCrLf & arr(5, 5) & vbCrLf & arr(6, 5) & vbCrLf & arr(7, 5) & vbCrLf & arr(8, 5) & vbCrLf & arr(9, 5) & vbCrLf & arr(10, 5) & vbCrLf & arr(11, 5) & vbCrLf & arr(12, 5) & vbCrLf & arr(13, 5) & vbCrLf & arr(14, 5))
            Stop
            Exit Function
        'end error handle code
        End Function
        
Private Function check(Optional dont_show_information As Boolean, Optional more_instructions As String) As Variant
    'currently functional as of (9/3/2020) checked by: (Zachary Daugherty)
    'Created By (Zachary Daugherty)(9/3/2020)
    'Purpose Case & notes:
        'this function is a check function on if data is stored in the proper place before updates or data is manipulated for the dts page
    'Library Refrences required
        'workbook.object
    'Modules Required
        'string_v1
    'Inputs
        'Internal:
            'na
        'required:
            'na
        'optional:
            'na
    'returned outputs
        'true if check is passed
        'false if stuff should not be updated
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            check = "check - Need Log report 10/28/20"
            Exit Function
        End If
    'code start
        'check for dont_show_information
            If (dont_show_information = False) Then
                MsgBox ("_________________String_Vx.check instructions_________________" & String_V1.get_Special_Char_V1(carriage_return, True) & _
                "function is called to make sure all the keystone locations of the data set are anchored to the right positions. data sets that should be checked are the " & _
                "following: DTS Tables and SP Tables. if either of these do not pass checks there will be errors generated to fix the positional data of the file")
                Stop
                Exit Function
            End If
        'define variables
            Dim condition As Boolean
        'run
            condition = DTS_V2A.Check_DTS_Table_V0_01A
            If (condition = True) Then
                condition = False
                condition = SP_V1_DEV.Check_SP_A_Table_V1
            End If
            If (condition = True) Then
                condition = False
                condition = SP_V1_DEV.Check_SP_B_Table_V0_01A
            End If
            
            check = condition
    'code end
        Exit Function
    'error handle
        'na
    'end error handle
End Function
        




























