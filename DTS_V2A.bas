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

Public Enum run_choices_V1
    'Purpose Case & notes:
        'gives the enumeration of choices that are setup for options
    'list
        update_unit_cost
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
        
        Private Function LOG_push_version(ByVal pos As Long) 'for version reporting
            Dim version As String
            Dim sht As Worksheet
            
            version = "2.0 Alpha without function logs"
            
            If (Boots_Main_V_alpha.sheet_exist(ActiveWorkbook, "Boots") = True) Then
                Set sht = ActiveWorkbook.Sheets("Boots")
            Else
                Exit Function
            End If
            sht.Cells(boots_pos.p_track_module_version_row + pos, boots_pos.p_track_module_version_col).value = version
            
        End Function
        
        Private Function LOG_push_project_file_requirements() As Variant
            LOG_push_project_file_requirements = _
            "<Boots_main_Valpha>" & Chr(149) & _
            "<Boots_Report_Valpha>" & Chr(149) & _
            "<SP_V1>" & Chr(149) & _
            "<Matrix_V2>" & Chr(149) & _
            "<String_V1>" & Chr(149)
        End Function
        
        Private Function LOG_Push_Functions_v1() As Variant
            'add discription
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
Log_Push_restart_size_check:
                i = i + 1
                s = sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value
                If (s <> "") Then
                    GoTo Log_Push_restart_size_check
                End If
            'get each status
                'returning listed enums
                    'Enum header
                        s = "__________Project Object ENUMS__________"
                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
                                            i = i + 1
                    'run_choices_V1
                        s = "ENUM: run_choices_V1 - Public"
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
                        s = DTS_V2A.Run_unit_cost_refresh_v0(True, "Log_Report")
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
        
        Public Function run_V0(ByVal choice As run_choices_V1, Optional more_instructions As String) As Variant
            'check for log reporting
                If (more_instructions = "Log_Report") Then
                    run_V0 = "run_v0 - Public - Out of Date: Using old code rather than the standardized code for checking page existance"
                    Exit Function
                End If
            'code start
                MsgBox ("need to add green text: DTS_V2A")
                'define variables
                    Dim condition As Boolean
                'setup variables
                    'na
                'start check
                    MsgBox ("'dts_vx_dev.run' need to add boots check insted of the one used on dts as it can then use a standard check for a page exist.")
                    condition = DTS_V2A.check(True)
                'check for check pass and if so run command else throw error
                    If (condition = True) Then
                        Select Case choice
                            Case update_unit_cost
                                DTS_V2A.Run_unit_cost_refresh_v0 (True)
                            Case Else
                                Stop
                        End Select
                    Else
                        GoTo Dts_error_run_check_not_passed
                        Stop
                    End If
            'code end
                run_V0 = 1
                Exit Function
            'error handling
Dts_error_run_check_not_passed:
                MsgBox ("Error DTS_VX: function check did not pass the nessasary checks to run commands for DTS please check your code and postions.")
                Stop
                Exit Function
        End Function
        
Private Function Run_unit_cost_refresh_v0(Optional dont_show_information As Boolean, Optional more_instructions As String) As Variant
'currently functional as of (9/2/2020) checked by: (Zachary Daugherty)
    'Created By (Zachary Daugherty)(8/25/20)
    'Purpose Case & notes:
        'this function will address updating the dts page unit cost
    'Library Refrences required
        'workbook.object
    'Modules Required
        'SP_V1
        'String_V1
        'Matrix_V2
    'Inputs
        'Internal:
            'na
        'required:
            'na
        'optional:
            'Na
    'returned outputs
        'na
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            Run_unit_cost_refresh_v0 = "Run_unit_cost_refresh_v0 - Private - Out Dated: Using old code this function is missing the proper refs to run"
            Exit Function
        End If
    'code start
        'check for show instructions
            If (dont_show_information = False) Then
                MsgBox ("this Function is designed to:" & String_V1.get_Special_Char_V1(carriage_return, True) & _
                "--------------------------------------------------------------------" & String_V1.get_Special_Char_V1(carriage_return, True) & _
                "Take the indexed information from the Steel preset data tables and apply the new unit cost to the DTS table. " & _
                "This function is private and should normally be called thru the RUN function to prevent mistakes" & String_V1.get_Special_Char_V1(carriage_return, True) & _
                "")
                Stop
                Exit Function
            End If
        'define variables
            'positional
                Dim wb As Workbook              'cursor position
                Dim home_pos As Worksheet       'returns to this position post subroutine
                Dim current_sht As Worksheet    'cursor position
                Dim row As Long                 'cursor position
                Dim col As Long                 'cursor position
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
            'setup pos
                Set wb = ActiveWorkbook
                Set home_pos = ActiveSheet
                On Error GoTo dts_Run_unit_cost_refresh_v0_cant_find_SHEET         'goto error handler
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
                Stop
                SP_GLOBAL_plate_value = -1
                DTS_Inflation_value = -1
                size_of_dts = -1
                size_of_sp_A = -1
                size_of_sp_B = -1
            'setup array tables
                'check for valid sizes and get sizes
                    'get size of DTS
                        size_of_dts = DTS_V2A.get_size_V0()
                    'get size of steel presets A
                        size_of_sp_A = SP_V1_DEV.get_size_A
                    'get size of steel presets B
                        size_of_sp_B = SP_V1_DEV.get_size_B
                'initialize arrays
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
            'setup globals
                'dts
                    On Error GoTo dts_Run_unit_cost_refresh_v0_cant_find_SHEET
                        error = "DTS"
                        Set current_sht = wb.Sheets(error)
                        error = ""
                    On Error GoTo 0
                    DTS_Inflation_value = current_sht.Cells(DTS_POS_2A.DTS_I_Inflation_Const_row, DTS_POS_2A.DTS_I_Inflation_Const_col).value * 100
                'SP
                    On Error GoTo dts_Run_unit_cost_refresh_v0_cant_find_SHEET
                        error = "STEEL PRESETS"
                        Set current_sht = wb.Sheets(error)
                        error = ""
                    On Error GoTo 0
                    SP_GLOBAL_STRUCTURAL_value = current_sht.Cells(SP_POS.SP_I_Const_Structural_row, SP_POS.SP_I_Const_Structural_col).value * 100
                    SP_GLOBAL_plate_value = current_sht.Cells(SP_POS.SP_I_Const_Plate_row, SP_POS.SP_I_Const_Plate_col).value * 100
                'return cursor to home
                    home_pos.Activate
        'load tables
            'dts main
                'set focus
                    On Error GoTo dts_Run_unit_cost_refresh_v0_cant_find_SHEET
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
            'SP_DECODER_A
                'set focus
                    On Error GoTo dts_Run_unit_cost_refresh_v0_cant_find_SHEET
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
            'SP_DECODER_B
                'set focus
                    On Error GoTo dts_Run_unit_cost_refresh_v0_cant_find_SHEET
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
        'assemble lookup table
            'initialize variables
                L = 0
                L_2 = 1
            'start
                Stop 'DEBUG NEED TO setup
                matrix_V2.matrix_dimensions (Lookup())
                Stop
incoding_of_table_names:
                For L = 0 To (UBound(Lookup(), 1) - LBound(Lookup(), 1))
                    'check to see where table data should be grabed from note will fill in table A first then B
                        'section for table a
                            If (L > 0) Then
                                If (L <= size_of_sp_A - 1) Then '-1 is included on the end to skip the goalpost of the table
                                    Lookup(L, 0) = SP_decoder_A(L, 1)
                                    Lookup(L, 1) = SP_V1_DEV.get_sheet_name(True)
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
                                    Lookup(L, 1) = SP_V1_DEV.get_sheet_name(True)
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
        'check lookup table for duplicate entrys
            'initialize variable
                L = 0
                L_2 = 0
                s = ""
                condition = False
            'start
                Stop 'NEED TO setup
                matrix_V2.matrix_dimensions (Lookup())
                Stop
                'loop through lookup tbale
                    For L = 1 To (UBound(Lookup(), 1) - LBound(Lookup(), 1))
                        'value to search by
                            'if set to empty skip
                                If (Lookup(L, 1) = DTS_V2A.get_global_unit_cost_refresh_ignore_trigger) Then
                                    GoTo Run_unit_cost_refresh_v0_ignore_entry
                                End If
                        'compair against all other entrys
                            For L_2 = 1 To (UBound(Lookup(), 1) - LBound(Lookup(), 1))
                                'check to see if index is the same if so skip check
                                    If (L = L_2) Then
                                        GoTo Run_unit_cost_refresh_v0_skip_check
                                    End If
                                'set value to check thru
                                    s = Lookup(L, 0)
                                'do check
                                    condition = String_V1.is_same_V1(s, Lookup(L_2, 0), True)
                                'if condition is true then throw error
                                    If (condition = True) Then
                                        error = "in array: lookup:(" & L_2 & ",0) value:'" & Lookup(L_2, 0) & "'. is the same as the value in: lookupL(" & L & ",0)"
                                        GoTo Run_unit_cost_refresh_v0_duplicate_lookups
                                    End If
                                'goto
Run_unit_cost_refresh_v0_skip_check:
                            Next L_2
                        'goto
Run_unit_cost_refresh_v0_ignore_entry:
                    Next L
            'cleanup
                L = -1
                L_2 = -1
                s = "empty"
                condition = False
        'do update
            'initalize variables
                On Error GoTo dts_Run_unit_cost_refresh_v0_cant_find_SHEET
                    error = "DTS"
                    Set current_sht = wb.Sheets(error)
                    error = ""
                On Error GoTo 0
                row = DTS_POS_2A.DTS_I_part_number_row
                col = DTS_POS_2A.DTS_I_part_number_col
                L = 0
                L_2 = 0
            'run
                'iterate thru memory main
                    For L = 1 To size_of_dts
                        'change pos
                            row = DTS_POS_2A.DTS_I_part_number_row + L
                        'set smart code
                            s = Memory_Main(L, 2)
                        'check for empty or ignore trigger
                            If ((s <> DTS_V2A.get_global_unit_cost_refresh_ignore_trigger) And (s <> "")) Then
                                'decode smart code
Run_unit_cost_refresh_v0_part_numb_check:
                                    s = String_V1.Disassociate_by_Char_V1(DTS_V2A.get_global_decoder_symbol, s, Left_C, True)
                                'search for key in lookup array
                                    For L_2 = 1 To (size_of_sp_A + size_of_sp_B)
                                        If (s = Lookup(L_2, 0)) Then
                                            'match found return value to sheet
                                                'locate which chart
                                                    If (Lookup(L_2, 2) = "A") Then
                                                        'match found in decode table 'A'
                                                            'return value to dts
                                                                On Error GoTo dts_Run_unit_cost_refresh_v0_cant_find_range
                                                                    error = "DTS_Unit_cost"
                                                                        current_sht.Range(error).Offset(L, 0).value = SP_decoder_A(CLng(Lookup(L_2, 3)), 4) 'paste to sheet name <current_sht> then move cursor to range <error> offset down to pos <L>: to get the value find in array <lookup> and return address of the match. convert to <long> variable and then user that long to look in array <sp_decoder_a> at the value of <long> then return that value to sheet
                                                                    error = ""
                                                                On Error GoTo 0
                                                    Else
                                                        If (Lookup(L_2, 2) = "B") Then
                                                            'match found in decode table 'B'
                                                                'return value to dts
                                                                    On Error GoTo dts_Run_unit_cost_refresh_v0_cant_find_range
                                                                        error = "DTS_Unit_cost"
                                                                        current_sht.Range(error).Offset(L, 0).value = SP_decoder_B(CLng(Lookup(L_2, 3)), 4)
                                                                        error = ""
                                                                    On Error GoTo 0
                                                        Else
                                                            Stop 'throw error
                                                            error = CStr(Lookup(L_2, 2))
                                                            GoTo dts_Run_unit_cost_refresh_v0_cant_locate_table
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
                                            GoTo Run_unit_cost_refresh_v0_part_numb_check
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
    'code end
        Exit Function
    'error handling
dts_Run_unit_cost_refresh_v0_cant_find_SHEET:
        Call MsgBox("FATAL Error: DTS_Vx: sub: Run_unit_cost_refresh_v0: was unaable to find the sheet named '" & error & "', please check your code.", , "FATAL Error: DTS_Vx: sub: Run_unit_cost_refresh_v0:: #1")
        Stop
Run_unit_cost_refresh_v0_duplicate_lookups:
        Call MsgBox("FATAL Error: DTS_Vx: Function: Run_unit_cost_refresh_v0:" & Chr(10) & "During the assembly " & error & Chr(10) & " please make the nessasary changes to the tables to not have duplicate values", , "FATAL Error: DTS_Vx: Function: Run_unit_cost_refresh_v0: #2")
        Stop
dts_Run_unit_cost_refresh_v0_cant_find_range:
        Call MsgBox("FATAL Error: Dts_vx: Function: Run_unit_cost_refresh_v0:" & Chr(10) & "Range(" & error & ") was unable to be located", , "FATAL Error: DTS_Vx: Function: Run_unit_cost_refresh_v0: #3")
        Stop
dts_Run_unit_cost_refresh_v0_cant_locate_table:
        Call MsgBox("FATAL Error: Dts_vx: Function: Run_unit_cost_refresh_v0:" & Chr(10) & "Function was unable to locate the table named:'" & error & "'" & Chr(10) & "Please see the goto 'incoding_of_table_names' as this is where the table chars are assigned", , "FATAL Error: Dts_vx: Function: Run_unit_cost_refresh_v0: #4")
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
            Get_V0 = "Get_V0 - Public - Stable"
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
            get_size_V0 = "get_size_V0 - Private - Stable"
            Exit Function
        End If
    'code start
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
            Set wb = ActiveWorkbook
            Set home_pos = ActiveSheet
            On Error GoTo dts_get_cant_find_DTS_SHEET   'goto error handler
                Set current_sht = wb.Sheets("DTS")      'setting name
            On Error GoTo 0                             'returns error handler to default
        'move to start location
            row = DTS_POS_2A.DTS_I_part_number_row     'fetching indexed information from enumeration
            col = DTS_POS_2A.DTS_I_part_number_col     'fetching indexed information from enumeration
            s = current_sht.Cells(row, col).value   'fetching indexed information from enumeration
        'get lenght to bottom
            On Error GoTo dts_cant_find_goalpost                    'goto error handler
                dist_to_goalpost = Range("DTS_GOALPOST").row - row  'setting definition
            On Error GoTo 0                                         'returns error handler to default
        'setup arr
            ReDim arr(1 To dist_to_goalpost)        'defining dimensions for the array
                'arr memory guide
                    'arr(<X:true if cell empty, false if not>)

        'all rows with data if there is no data in cell location mark for removal
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
                Else
                    arr(i) = False  'row does not need to be deleted
                End If
            Next i
            'cleanup
                i = -1
                i_2 = -1
                s = "empty"
        'check for delete condition to be true\
            If (delete_empty_rows_condition = True) Then
                'hide updating
                    Application.ScreenUpdating = False
                    Application.DisplayAlerts = False
                'move to start location
                    row = DTS_POS_2A.DTS_I_part_number_row
                    col = DTS_POS_2A.DTS_I_part_number_col
                    s = current_sht.Cells(row, col).value
                'iterate through to find empty then delete by moving everything up eliminating the blank space
                    For i = 1 To dist_to_goalpost
                        row = row + 1
                        s = current_sht.Cells(row, col).value
                        If (arr(i) = "True") Then
                            Stop
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
                                Stop
                            'delete row
                                Rows(s).Select
                                Selection.Delete Shift:=xlUp
                                i = i + 1
                        End If
                    Next i
                'start updating
                    'restart if things were deleted, reset some variables then goto 'get_size_V0_restart'
                        If (delete_empty_rows_condition = True) Then
                            current_sht.visible = i_2
                            If (anti_loop < 30) Then
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
                                    GoTo get_size_V0_restart
                            Else
                                MsgBox ("FATAL ERROR: ANTI_LOOP Triggered check code")
                                Stop
                            End If
                        End If
                'unhide updating
                    Application.ScreenUpdating = True
                    Application.DisplayAlerts = True
            End If
            'get final size
                get_size_V0 = dist_to_goalpost
    'code end
        Exit Function
    'error handling
dts_get_cant_find_DTS_SHEET:
        MsgBox ("Error: Dts_vx: FUNCTION  get_size_V0: was unable to find the sheet named dts, please check your code.")
        Stop
dts_cant_find_goalpost:
        MsgBox ("Error: Dts_vx: FUNCTION  get_size_V0: was unable to findd the range named DTS_GOALPOST. please check your code.")
        Stop
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
                'DEV_V1_DEV
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
                    Check_DTS_Table_V0_01A = "Check_DTS_Table_V0_01A - Out of date: Software issues have not been resolved as a field has been removed - Private"
                    Exit Function
                End If
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
                'report to log
                    Call MsgBox("check dts table using log replace", , "check dts table using log")
                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "CHECK DTS_TABLE_V0_01A Started")
                'breakout
                Set proj_wb = ActiveWorkbook
                On Error GoTo FATAL_ERROR_CHECK_DTS_SET_DTS_ENV 'set error handler
                    Set cursor_sheet = proj_wb.Sheets("DTS")
                On Error GoTo 0 'set error handler back to norm
                cursor_row = 1
                cursor_col = 1
            'debug
                proj_wb.Sheets("Boots").visible = -1
                Stop 'debug
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
                                Stop
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
                    Call MsgBox("check dts_table using log replace", , "check dts_table using log")
                    'Call dev_v1_dev.log(dev_v1_dev.get_username, "CHECK DTS_TABLE_V0_01A Finished")
                    Exit Function
        'code end
        'error handle
ERROR_FATAL_check_dts_range_error:
            Check_DTS_Table_V0_01A = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: Displaying Snapshot of Values:...")
                Call Boots_Report_v_Alpha.Log_Push(table_open, "")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Check_DTS_Table_V0_01A: '" & Check_DTS_Table_V0_01A & "' as variant")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "More_instructions: '" & more_instructions & "' as string")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "condition: '" & condition & "' as Boolean")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "i: '" & i & "' as Long")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "s: '" & s & "' as String")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Proj_wb: '" & proj_wb.path & "/" & proj_wb.Name & "' as Workbook")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_sheet: '" & cursor_sheet.Name & " :visible = " & cursor_sheet.visible & "' as worksheet")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_row: '" & cursor_row & "' as long")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "cursor_col: '" & cursor_col & "' as long")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "ref_rng: '" & ref_rng.Name & " value=" & ref_rng.value & "' as range")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "arr: '<please check the array>' as string")
                Call Boots_Report_v_Alpha.Log_Push(table_close, "")
                Call Boots_Report_v_Alpha.Log_Push(Flag, "")
                    Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR: MODULE:(DTS_VX)FUNCTION:(CHECK_DTS_TABLE) UNABLE TO LOCATE THE SPECIFIED RANGE:<" & s & "> please check the name mannager for errors. fix and then re-run")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                    Call Boots_Report_v_Alpha.Log_Push(text, "CRASH!__________________________________________________________")
                    For i = 1 To 10
                        Call Boots_Report_v_Alpha.Log_Push(text, "________________________________________________________________")
                    Next i
                    
                Call Boots_Report_v_Alpha.Log_Push(Display_now, "")
                Stop
            Exit Function
FATAL_ERROR_CHECK_DTS_SET_DTS_ENV:
            Call MsgBox("check dts_table using log replace", , "check dts_table using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "FATAL_ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) UNABLE TO FIND OR SET SHEET DTS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.")
            Call MsgBox("FATAL_ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) UNABLE TO FIND OR SET SHEET DTS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.", , "FATAL ERROR: SET DTS SHEET ENV")
            Stop
            Exit Function
ERROR_CHECK_DTS_FAILED_POS_CHECK:
            Call MsgBox("check dts_table using log replace", , "check dts_table using log")
            'Call dev_v1_dev.log(dev_v1_dev.get_username, "ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: " & vbCrLf & vbCrLf & arr(1, 5) & vbCrLf & arr(2, 5) & vbCrLf & arr(3, 5) & vbCrLf & arr(4, 5) & vbCrLf & arr(5, 5) & vbCrLf & arr(6, 5) & vbCrLf & arr(7, 5) & vbCrLf & arr(8, 5) & vbCrLf & arr(9, 5) & vbCrLf & arr(10, 5) & vbCrLf & arr(11, 5) & vbCrLf & arr(12, 5) & vbCrLf & arr(13, 5) & vbCrLf & arr(14, 5) & vbCrLf & arr(15, 5))
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
            check = "check - Stable"
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
                condition = SP_V1_DEV.Check_SP_A_Table_V0_01A
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
        




























