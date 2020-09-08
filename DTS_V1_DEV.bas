Attribute VB_Name = "DTS_V1_DEV"
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
                                                        'made for :DTS_V1
                                                                 ':DEV_v1_DEV
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Purpose Case
                    'This Module is built to handle all referances to the Price Tool DTS database For proper Referenceing and Updating
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


        Enum run_choices_V0
        'currently functional as of (8/21/2020) checked by: (zachary daugherty)
            'Created By (zachary Daugherty)(8/21/20)
            'Purpose Case & notes:
                'gives the enumeration of choices that are setup for options
            'Library Refrences required
                'na
            'Modules Required
                'workbook.object
            'Inputs
                'Internal:
                    'na
                'required:
                    'na
                'optional:
                    'na
            'returned outputs
                'index
            'code start
                update_unit_cost
            'code end
                'na
            'error handle
                'na
            'end error handle
        End Enum
        
        Public Enum DTS_POS
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
                DTS_I_Inflation_Const_row = 1
                    DTS_I_Inflation_Const_col = 6
            'number of entry fields watching
                DTS_Q_number_of_tracked_locations = 15
            'array of positions
                DTS_I_part_number_row = 3 'used as global table header row position
                    DTS_I_part_number_col = 1
                DTS_I_AKA_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_AKA_col = 2
                DTS_I_Description_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Description_col = 3
                DTS_I_Unit_list_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Unit_list_col = 4
                DTS_I_Unit_cost_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Unit_cost_col = 5
                DTS_I_Adjusted_unit_cost_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Adjusted_unit_cost_col = 6
                DTS_I_Unit_weight_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Unit_weight_col = 7
                DTS_I_Shop_origin_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Shop_origin_col = 8
                DTS_I_Status_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Status_col = 9
                DTS_I_Job_other_info_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Job_other_info_col = 10
                DTS_I_Vendor_info_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Vendor_info_col = 11
                DTS_I_Vendor_phone_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Vendor_phone_col = 12
                DTS_I_Vendor_fax_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Vendor_fax_col = 13
                DTS_I_Vendor_part_number_row = DTS_POS.DTS_I_part_number_row
                    DTS_I_Vendor_part_number_col = 14

        End Enum
        
        Private Function get_global_unit_cost_refresh_ignore_trigger() As String
            get_global_unit_cost_refresh_ignore_trigger = "<skip>"
        End Function
        
        Private Function get_global_decoder_symbol() As String
            get_global_decoder_symbol = "-"
        End Function
        
        Private Function get_unit_cost_refresh_Steel_presets_grab_row() As Long
            
        End Function
        
        Private Function Check_DTS_Table_V0_01A() As Boolean
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
        'code start
            'define varables
                'memory
                    Dim arr() As String             'designed as ram storage
                    Dim condition As Boolean        'store T/F
                    Dim I As Long                'iterator and int storage
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
                    Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "CHECK DTS_TABLE_V0_01A Started")
                'breakout
                Set proj_wb = ActiveWorkbook
                On Error GoTo FATAL_ERROR_CHECK_DTS_SET_DTS_ENV 'set error handler
                    Set cursor_sheet = proj_wb.Sheets("DTS")
                On Error GoTo 0 'set error handler back to norm
                cursor_row = 1
                cursor_col = 1
            'check for visual
                If (Visualy = True) Then
                    MsgBox ("not setup")
                    Exit Function
                End If
            'setup arr
                'redefine size of the arr
                    ReDim arr(1 To DTS_POS.DTS_Q_number_of_tracked_locations, 1 To 5) 'see line below for definitions
                        'arr memory assignments
                            '(<specific index>,<1 to 5>)
                            '(<specific index>,<1:row of enum>)
                            '(<specific index>,<2:col of enum>)
                            '(<specific index>,<3:row of range>)
                            '(<specific index>,<4:col of range>)
                            '(<specific index>,<5: conditional if match>)
                'fill arr
                    'Collect information
                        I = 0
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        'NOTICE CODE IN THIS BLOCK IS STD AND THE OPERATIONS ARE THE SAME SO DEV NOTES ON THE FIRST FOLLOW THRU
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        
                        'compair <part number> expected location
                            s = "DTS_Part_number"   'expected range name for search
                            I = I + 1               'iterate arr position from x to x + 1 in the array
                            On Error GoTo ERROR_FATAL_check_dts_range_error 'if specified range 'S' is unable to be found or set goto Error handler at bottom of this function
                                Set ref_rng = Range(s)  'set range
                            On Error GoTo 0 'reset error handler
                                arr(I, 1) = CStr(ref_rng.row)       'get range row pos
                                arr(I, 2) = CStr(ref_rng.Column)    'get range col pos
                                arr(I, 3) = DTS_POS.DTS_I_part_number_row    'get enum row pos
                                arr(I, 4) = DTS_POS.DTS_I_part_number_col    'get enum col pos
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then   'compair rows to rows and cols to cols
                                    arr(I, 5) = s & ": " & True 'if true report text
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair" 'if False report text
                                    condition = True    'if true at the end of the block throw error as there is a miss match
                                End If
                        'compair <DTS_AKA> expected location
                            s = "DTS_AKA"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_AKA_row
                                arr(I, 4) = DTS_POS.DTS_I_AKA_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Description> expected location
                            s = "DTS_Description"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Description_row
                                arr(I, 4) = DTS_POS.DTS_I_Description_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Unit_list> expected location
                            s = "DTS_Unit_list"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Unit_list_row
                                arr(I, 4) = DTS_POS.DTS_I_Unit_list_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Unit_cost> expected location
                            s = "DTS_Unit_cost"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Unit_cost_row
                                arr(I, 4) = DTS_POS.DTS_I_Unit_cost_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Adjusted_unit_cost> expected location
                            s = "DTS_Adjusted_unit_cost"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Adjusted_unit_cost_row
                                arr(I, 4) = DTS_POS.DTS_I_Adjusted_unit_cost_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Unit_weight> expected location
                            s = "DTS_Unit_weight"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Unit_weight_row
                                arr(I, 4) = DTS_POS.DTS_I_Unit_weight_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Shop_origin> expected location
                            s = "DTS_Shop_origin"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Shop_origin_row
                                arr(I, 4) = DTS_POS.DTS_I_Shop_origin_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Status> expected location
                            s = "DTS_Status"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Status_row
                                arr(I, 4) = DTS_POS.DTS_I_Status_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Job_other_info> expected location
                            s = "DTS_Job_other_info"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Job_other_info_row
                                arr(I, 4) = DTS_POS.DTS_I_Job_other_info_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Vendor_info> expected location
                            s = "DTS_Vendor_info"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Vendor_info_row
                                arr(I, 4) = DTS_POS.DTS_I_Vendor_info_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Vendor_phone> expected location
                            s = "DTS_Vendor_phone"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Vendor_phone_row
                                arr(I, 4) = DTS_POS.DTS_I_Vendor_phone_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Vendor_fax> expected location
                            s = "DTS_Vendor_fax"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Vendor_fax_row
                                arr(I, 4) = DTS_POS.DTS_I_Vendor_fax_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Vendor_part_number> expected location
                            s = "DTS_Vendor_part_number"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Vendor_part_number_row
                                arr(I, 4) = DTS_POS.DTS_I_Vendor_part_number_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        'compair <DTS_Inflation_Const> expected location
                            s = "DTS_Inflation_Const"
                            I = I + 1
                            On Error GoTo ERROR_FATAL_check_dts_range_error
                            Set ref_rng = Range(s)
                            On Error GoTo 0
                                arr(I, 1) = CStr(ref_rng.row)
                                arr(I, 2) = CStr(ref_rng.Column)
                                arr(I, 3) = DTS_POS.DTS_I_Inflation_Const_row
                                arr(I, 4) = DTS_POS.DTS_I_Inflation_Const_col
                                If ((arr(I, 1) = arr(I, 3)) And (arr(I, 2) = arr(I, 4))) Then
                                    arr(I, 5) = s & ": " & True
                                Else
                                    arr(I, 5) = s & ": " & False & vbCrLf & "____please check and compair"
                                    condition = True
                                End If
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        '__________________________________________END of CODE BLOCK___________________________________________
                        '-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                        'cleanup
                            I = 0
                            s = "Empty"
            'compile report
                'check to see if failure condition is met
                    If (condition = True) Then
                        GoTo ERROR_CHECK_DTS_FAILED_POS_CHECK
                    End If
                'return true
                    Check_DTS_Table_V0_01A = True   'passed all checks
                    Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "CHECK DTS_TABLE_V0_01A Finished")
                    Exit Function
        'code end
        'error handle
ERROR_FATAL_check_dts_range_error:
            Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "FATAL ERROR: MODULE:(DTS_VX)FUNCTION:(CHECK_DTS_TABLE) UNABLE TO LOCATE THE SPECIFIED RANGE:<" & s & "> please check the name mannager for errors. fix and then re-run")
            Call MsgBox("FATAL ERROR: MODULE:(DTS_VX)FUNCTION:(CHECK_DTS_TABLE) UNABLE TO LOCATE THE SPECIFIED RANGE:<" & s & "> please check the name mannager for errors. fix and then re-run", , "Fatal error")
            Stop
FATAL_ERROR_CHECK_DTS_SET_DTS_ENV:
            Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "FATAL_ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) UNABLE TO FIND OR SET SHEET DTS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.")
            Call MsgBox("FATAL_ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) UNABLE TO FIND OR SET SHEET DTS IN THE PROJECT WORKBOOK PLEASE CHECK FOR RIGHT CALL OR POS OR WORKBOOK.", , "FATAL ERROR: SET DTS SHEET ENV")
            Stop
ERROR_CHECK_DTS_FAILED_POS_CHECK:
            Call DEV_V1_DEV.log(DEV_V1_DEV.get_username, "ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: " & vbCrLf & vbCrLf & arr(1, 5) & vbCrLf & arr(2, 5) & vbCrLf & arr(3, 5) & vbCrLf & arr(4, 5) & vbCrLf & arr(5, 5) & vbCrLf & arr(6, 5) & vbCrLf & arr(7, 5) & vbCrLf & arr(8, 5) & vbCrLf & arr(9, 5) & vbCrLf & arr(10, 5) & vbCrLf & arr(11, 5) & vbCrLf & arr(12, 5) & vbCrLf & arr(13, 5) & vbCrLf & arr(14, 5) & vbCrLf & arr(15, 5))
            Call MsgBox("ERROR: MODULE: (DTS_VX)FUNCTION: (CHECK_DTS_TABLE) FAILED POSITIONAL CHECK REPORT LISTED BELOW A FAIL IS LISTED AS FALSE: " & vbCrLf & vbCrLf & arr(1, 5) & vbCrLf & arr(2, 5) & vbCrLf & arr(3, 5) & vbCrLf & arr(4, 5) & vbCrLf & arr(5, 5) & vbCrLf & arr(6, 5) & vbCrLf & arr(7, 5) & vbCrLf & arr(8, 5) & vbCrLf & arr(9, 5) & vbCrLf & arr(10, 5) & vbCrLf & arr(11, 5) & vbCrLf & arr(12, 5) & vbCrLf & arr(13, 5) & vbCrLf & arr(14, 5) & vbCrLf & arr(15, 5))
            Stop
        'end error handle code
        End Function
        
        Public Function get_size() As Long
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
                        Dim I As Long       'int storage 1
                        Dim i_2 As Long     'int sotrage 2
                        Dim s As String     'string storage
                'restart trigger
get_size_restart:               'goto flag
                'setup variables
                    Set wb = ActiveWorkbook
                    Set home_pos = wb.ActiveSheet
                    On Error GoTo dts_get_cant_find_DTS_SHEET   'goto error handler
                        Set current_sht = wb.Sheets("DTS")      'setting name
                    On Error GoTo 0                             'returns error handler to default
                'move to start location
                    row = DTS_POS.DTS_I_part_number_row     'fetching indexed information from enumeration
                    col = DTS_POS.DTS_I_part_number_col     'fetching indexed information from enumeration
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
                    For I = 1 To dist_to_goalpost
                        row = row + 1
                        s = current_sht.Cells(row, col)
                        If (s = "") Then    'check the other cols in the row to see if any data is stored.
                            For i_2 = 1 To DTS_POS.DTS_Q_number_of_tracked_locations - 1
                                s = current_sht.Cells(row, col + 1)
                                If (s <> "") Then
                                    arr(I) = False  'row does not need to be deleted as it has values in fields other than part number
                                    Exit For
                                End If
                                delete_empty_rows_condition = True
                                arr(I) = True   'row is entirly empty so mark for delete
                            Next i_2
                        Else
                            arr(I) = False  'row does not need to be deleted
                        End If
                    Next I
                    'cleanup
                        I = -1
                        i_2 = -1
                        s = "empty"
                'check for delete condition to be true\
                    If (delete_empty_rows_condition = True) Then
                        'hide updating
                            Application.ScreenUpdating = False
                            Application.DisplayAlerts = False
                        'move to start location
                            row = DTS_POS.DTS_I_part_number_row
                            col = DTS_POS.DTS_I_part_number_col
                            s = current_sht.Cells(row, col).value
                        'iterate through to find empty then delete by moving everything up eliminating the blank space
                            For I = 1 To dist_to_goalpost
                                row = row + 1
                                s = current_sht.Cells(row, col).value
                                If (arr(I) = "True") Then
                                    Stop
                                    'setup
                                        On Error GoTo dts_get_cant_find_DTS_SHEET
                                            Set current_sht = wb.Sheets("DTS")
                                        On Error GoTo 0
                                        If ((current_sht.Visible = xlSheetVeryHidden) Or (current_sht.Visible = xlSheetHidden)) Then
                                            i_2 = current_sht.Visible
                                            current_sht.Visible = xlSheetVisible
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
                                        I = I + 1
                                End If
                            Next I
                        'start updating
                            'restart if things were deleted, reset some variables then goto 'get_size_restart'
                                If (delete_empty_rows_condition = True) Then
                                    current_sht.Visible = i_2
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
                                            I = -1
                                            i_2 = -1
                                            s = "empty"
                                            ReDim arr(0)
                                            delete_empty_rows_condition = False
                                        'do goto
                                            GoTo get_size_restart
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
                        get_size = dist_to_goalpost
            'code end
                Exit Function
            'error handling
dts_get_cant_find_DTS_SHEET:
                MsgBox ("Error: Dts_vx: FUNCCTION  GET_SIZE: was unaable to find the sheet named dts, please check your code.")
                Stop
dts_cant_find_goalpost:
                MsgBox ("Error: Dts_vx: FUNCCTION  GET_SIZE: was unable to findd the range named DTS_GOALPOST. please check your code.")
                Stop
        End Function
        
        Public Sub run(ByVal choice As run_choices_V0)
            MsgBox ("need to add green text: Dts_v1_dev")
            
            'code start
                'define variables
                    Dim condition As Boolean
                'setup variables
                    'na
                'start check
                    condition = DTS_V1_DEV.check(True)
                'check for check pass and if so run command else throw error
                    If (condition = True) Then
                        Select Case choice
                            Case update_unit_cost
                                DTS_V1_DEV.unit_cost_refresh (True)
                            Case Else
                                Stop
                        End Select
                    Else
                        GoTo Dts_error_run_check_not_passed
                        Stop
                    End If
            'code end
                Exit Sub
            'error handling
Dts_error_run_check_not_passed:
                MsgBox ("Error DTS_VX: function check did not pass the nessasary checks to run commands for DTS please check your code and postions.")
                Stop
        End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'START OF private routines belonging to Run()
'-----------------------------------------------------------------------------------------------------------------------------------------------
            Private Function check(Optional dont_show_information As Boolean) As Boolean
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
                        condition = DTS_V1_DEV.Check_DTS_Table_V0_01A
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
                    
                    
                    Private Function unit_cost_refresh(Optional dont_show_information As Boolean)
                    'currently functional as of (9/2/2020) checked by: (Zachary Daugherty)
                        'Created By (Zachary Daugherty)(8/25/20)
                        'Purpose Case & notes:
                            'this function will address updating the dts page unit cost
                        'Library Refrences required
                            'workbook.object
                        'Modules Required
                            'SP_V1
                            'String_v1
                            'array_V1
                        'Inputs
                            'Internal:
                                'na
                            'required:
                                'na
                            'optional:
                                'Na
                        'returned outputs
                            'na
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
                                    Set home_pos = wb.ActiveSheet
                                    On Error GoTo dts_unit_cost_refresh_cant_find_SHEET         'goto error handler
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
                                            size_of_dts = DTS_V1_DEV.get_size()
                                        'get size of steel presets A
                                            size_of_sp_A = SP_V1_DEV.get_size_A
                                        'get size of steel presets B
                                            size_of_sp_B = SP_V1_DEV.get_size_B
                                    'initialize arrays
                                        ReDim Memory_Main(0 To size_of_dts, 0 To DTS_POS.DTS_Q_number_of_tracked_locations)
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
                                        On Error GoTo dts_unit_cost_refresh_cant_find_SHEET
                                            error = "DTS"
                                            Set current_sht = wb.Sheets(error)
                                            error = ""
                                        On Error GoTo 0
                                        DTS_Inflation_value = current_sht.Cells(DTS_POS.DTS_I_Inflation_Const_row, DTS_POS.DTS_I_Inflation_Const_col).value * 100
                                    'SP
                                        On Error GoTo dts_unit_cost_refresh_cant_find_SHEET
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
                                        On Error GoTo dts_unit_cost_refresh_cant_find_SHEET
                                            error = "DTS"
                                                Set current_sht = wb.Sheets(error)
                                            error = ""
                                            'set row col
                                                row = DTS_POS.DTS_I_part_number_row
                                                col = DTS_POS.DTS_I_part_number_col

                                        On Error GoTo 0
                                    'get
                                        For L = 0 To size_of_dts
                                            For L_2 = 1 To DTS_POS.DTS_Q_number_of_tracked_locations
                                                Memory_Main(L, L_2) = current_sht.Cells(row, col).value
                                                col = DTS_POS.DTS_I_part_number_col + L_2
                                            Next L_2
                                            row = DTS_POS.DTS_I_part_number_row + L + 1
                                            col = DTS_POS.DTS_I_part_number_col
                                        Next L
                                    'cleanup
                                        row = -1
                                        col = -1
                                        L = -1
                                        L_2 = -1
                                        Set current_sht = Nothing
                                'SP_DECODER_A
                                    'set focus
                                        On Error GoTo dts_unit_cost_refresh_cant_find_SHEET
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
                                        On Error GoTo dts_unit_cost_refresh_cant_find_SHEET
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
                                    Array_V1.ArrayDimensions_Alpha 'use this function for array bounds
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
                                                            Lookup(L, 0) = DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger
                                                            Lookup(L, 1) = DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger
                                                            Lookup(L, 2) = DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger
                                                            Lookup(L, 3) = DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger
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
                                                                Lookup(L, 0) = DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger
                                                                Lookup(L, 1) = DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger
                                                                Lookup(L, 2) = DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger
                                                                Lookup(L, 3) = DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger
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
                                    Array_V1.ArrayDimensions_Alpha 'use this function for array bounds
                                    Stop
                                    'loop through lookup tbale
                                        For L = 1 To (UBound(Lookup(), 1) - LBound(Lookup(), 1))
                                            'value to search by
                                                'if set to empty skip
                                                    If (Lookup(L, 1) = DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger) Then
                                                        GoTo unit_cost_refresh_ignore_entry
                                                    End If
                                            'compair against all other entrys
                                                For L_2 = 1 To (UBound(Lookup(), 1) - LBound(Lookup(), 1))
                                                    'check to see if index is the same if so skip check
                                                        If (L = L_2) Then
                                                            GoTo unit_cost_refresh_skip_check
                                                        End If
                                                    'set value to check thru
                                                        s = Lookup(L, 0)
                                                    'do check
                                                        condition = String_V1.is_same_V1(s, Lookup(L_2, 0), True)
                                                    'if condition is true then throw error
                                                        If (condition = True) Then
                                                            error = "in array: lookup:(" & L_2 & ",0) value:'" & Lookup(L_2, 0) & "'. is the same as the value in: lookupL(" & L & ",0)"
                                                            GoTo unit_cost_refresh_duplicate_lookups
                                                        End If
                                                    'goto
unit_cost_refresh_skip_check:
                                                Next L_2
                                            'goto
unit_cost_refresh_ignore_entry:
                                        Next L
                                'cleanup
                                    L = -1
                                    L_2 = -1
                                    s = "empty"
                                    condition = False
                            'do update
                                'initalize variables
                                    On Error GoTo dts_unit_cost_refresh_cant_find_SHEET
                                        error = "DTS"
                                        Set current_sht = wb.Sheets(error)
                                        error = ""
                                    On Error GoTo 0
                                    row = DTS_POS.DTS_I_part_number_row
                                    col = DTS_POS.DTS_I_part_number_col
                                    L = 0
                                    L_2 = 0
                                'run
                                    'iterate thru memory main
                                        For L = 1 To size_of_dts
                                            'change pos
                                                row = DTS_POS.DTS_I_part_number_row + L
                                            'set smart code
                                                s = Memory_Main(L, 2)
                                            'check for empty or ignore trigger
                                                If ((s <> DTS_V1_DEV.get_global_unit_cost_refresh_ignore_trigger) And (s <> "")) Then
                                                    'decode smart code
unit_cost_refresh_part_numb_check:
                                                        s = String_V1.Disassociate_by_Char_V1(get_global_decoder_symbol, s, Left, True)
                                                    'search for key in lookup array
                                                        For L_2 = 1 To (size_of_sp_A + size_of_sp_B)
                                                            If (s = Lookup(L_2, 0)) Then
                                                                'match found return value to sheet
                                                                    'locate which chart
                                                                        If (Lookup(L_2, 2) = "A") Then
                                                                            'match found in decode table 'A'
                                                                                'return value to dts
                                                                                    On Error GoTo dts_unit_cost_refresh_cant_find_range
                                                                                        error = "DTS_Unit_cost"
                                                                                            current_sht.Range(error).Offset(L, 0).value = SP_decoder_A(CLng(Lookup(L_2, 3)), 4) 'paste to sheet name <current_sht> then move cursor to range <error> offset down to pos <L>: to get the value find in array <lookup> and return address of the match. convert to <long> variable and then user that long to look in array <sp_decoder_a> at the value of <long> then return that value to sheet
                                                                                        error = ""
                                                                                    On Error GoTo 0
                                                                        Else
                                                                            If (Lookup(L_2, 2) = "B") Then
                                                                                'match found in decode table 'B'
                                                                                    'return value to dts
                                                                                        On Error GoTo dts_unit_cost_refresh_cant_find_range
                                                                                            error = "DTS_Unit_cost"
                                                                                            current_sht.Range(error).Offset(L, 0).value = SP_decoder_B(CLng(Lookup(L_2, 3)), 4)
                                                                                            error = ""
                                                                                        On Error GoTo 0
                                                                            Else
                                                                                Stop 'throw error
                                                                                error = CStr(Lookup(L_2, 2))
                                                                                GoTo dts_unit_cost_refresh_cant_locate_table
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
                                                                GoTo unit_cost_refresh_part_numb_check
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
dts_unit_cost_refresh_cant_find_SHEET:
                            Call MsgBox("FATAL Error: DTS_Vx: sub: unit_cost_refresh: was unaable to find the sheet named '" & error & "', please check your code.", , "FATAL Error: DTS_Vx: sub: unit_cost_refresh:: #1")
                            Stop
unit_cost_refresh_duplicate_lookups:
                            Call MsgBox("FATAL Error: DTS_Vx: Function: Unit_cost_refresh:" & Chr(10) & "During the assembly " & error & Chr(10) & " please make the nessasary changes to the tables to not have duplicate values", , "FATAL Error: DTS_Vx: Function: Unit_cost_refresh: #2")
                            Stop
dts_unit_cost_refresh_cant_find_range:
                            Call MsgBox("FATAL Error: Dts_vx: Function: Unit_cost_refresh:" & Chr(10) & "Range(" & error & ") was unable to be located", , "FATAL Error: DTS_Vx: Function: Unit_cost_refresh: #3")
                            Stop
dts_unit_cost_refresh_cant_locate_table:
                            Call MsgBox("FATAL Error: Dts_vx: Function: Unit_cost_refresh:" & Chr(10) & "Function was unable to locate the table named:'" & error & "'" & Chr(10) & "Please see the goto 'incoding_of_table_names' as this is where the table chars are assigned", , "FATAL Error: Dts_vx: Function: Unit_cost_refresh: #4")
                    End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------
'END OF private routines belonging to Run()
'-----------------------------------------------------------------------------------------------------------------------------------------------




























