Attribute VB_Name = "DTH_VA"
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
                                            'This Module is built to handle all referances to the Price Tool
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                                                
                                                                        'GET GLOBAL FUNCTIONS

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Public Enum DTH_run_choices_V1
    'Purpose Case & notes:
        'gives the enumeration of choices that are setup for options
    'list
        DTH_update_unit_cost
End Enum

Public Enum DTH_POS_1A
    'Purpose Case & notes:
        'POS Enum is to be called to act as a check condition to verify that the code and the sheet agrees on
            'the locational position of where things are on the sheet.
    'list
        'other table information
            DTH_Inflation_Const_ROW = 1
                DTH_Inflation_Const_COL = 5
        'number of entry fields watching
            DTH_Q_number_pos = 15
            DTH_Q_number_other = 1
            DTH_Q_total_number_of_tracked_locations = DTH_POS_1A.DTH_Q_number_other + DTH_POS_1A.DTH_Q_number_pos
        'array of positions
            DTH_Part_Number_ROW = 3         'preset holder
                DTH_part_number_col = 1     'preset holder
            DTH_AKA_Number_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_AKA_Number_COL = DTH_POS_1A.DTH_part_number_col + 1
            DTH_Description_row = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_Description_COL = DTH_POS_1A.DTH_AKA_Number_COL + 1
            DTH_UNIT_COST_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_UNIT_COST_COL = DTH_POS_1A.DTH_Description_COL + 1
            DTH_ADJUSTED_UNIT_COST_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_ADJUSTED_UNIT_COST_COL = DTH_POS_1A.DTH_UNIT_COST_COL + 1
            DTH_Unit_Weight_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_Unit_Weight_COL = DTH_POS_1A.DTH_ADJUSTED_UNIT_COST_COL + 1
            DTH_Shop_Origin_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_Shop_Origin_COL = DTH_POS_1A.DTH_Unit_Weight_COL + 1
            DTH_Status_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_Status_COL = DTH_POS_1A.DTH_Shop_Origin_COL + 1
            DTH_CURRENT_SAP_COST_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_CURRENT_SAP_COST_COL = DTH_POS_1A.DTH_Status_COL + 1
            DTH_LAST_DATE_OF_PURCHASE_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_LAST_DATE_OF_PURCHASE_COL = DTH_POS_1A.DTH_CURRENT_SAP_COST_COL + 1
            DTH_JOB_OTHER_INFO_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_JOB_OTHER_INFO_COL = DTH_POS_1A.DTH_LAST_DATE_OF_PURCHASE_COL + 1
            DTH_Vendor_Information_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_Vendor_Information_COL = DTH_POS_1A.DTH_JOB_OTHER_INFO_COL + 1
            DTH_Vendor_Phone_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_Vendor_Phone_COL = DTH_POS_1A.DTH_Vendor_Information_COL + 1
            DTH_Vendor_Fax_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_Vendor_Fax_COL = DTH_POS_1A.DTH_Vendor_Phone_COL + 1
            DTH_Vendor_Part_Number_ROW = DTH_POS_1A.DTH_Part_Number_ROW
                DTH_Vendor_Part_Number_COL = DTH_POS_1A.DTH_Vendor_Fax_COL + 1
End Enum

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
            "" & Chr(149) & _
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
'                        s = "__________Project Object LOG Functions__________"
'                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
'                                    sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
'                                        sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
'                                            i = i + 1
                    'returning log push version
'                        s = DTH_V2A.LOG_push_version(0, "Log_Report")
'                            sht.Cells(boots_report_pos.p_indent_row + i, boots_report_pos.p_indent_col).value = sht.Cells(boots_report_pos.p_indent_row + i - 1, boots_report_pos.p_indent_col).value
'                                sht.Cells(boots_report_pos.p_time_row + i, boots_report_pos.p_time_col).value = Now()
'                                    sht.Cells(boots_report_pos.p_text_row + i, boots_report_pos.p_text_col).value = s
'                                        i = i + 1
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

Public Function run_V0(ByVal choice As DTH_run_choices_V1, Optional more_instructions As String) As Variant
            'check for log reporting
                If (more_instructions = "Log_Report") Then
                    run_V0 = "run_v0 - Public - still being built 12/2/2020"
                    Exit Function
                End If
                MsgBox ("DTH_VA.RUN_V0 change out msg when complete")
                Stop
            'code start
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.run_V0... Starting...")
                'define variables
                    Dim condition As Boolean
                    Dim i As Long
                'setup variables
                    'na
                'start check
                    MsgBox ("'DTH_VA.run' need to add boots check insted of the one used on dts as it can then use a standard check for a page exist.")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    condition = DTH_VA.check_table_operation_V0
                'check for check pass and if so run command else throw error
                    Call Boots_Report_v_Alpha.Log_Push(text, "Checking if conditions are met to run any actions...")
                    If (condition = True) Then
                        Call Boots_Report_v_Alpha.Log_Push(text, "Passed!...")
                        Select Case choice
                            Case DTH_update_unit_cost
                                DTH_VA.Run_DTH_unit_cost_refresh_v0
                        End Select
                    Else
                        Call Boots_Report_v_Alpha.Log_Push(text, "Failed!...")
                        GoTo Dth_error_run_check_not_passed
                    End If
            'code end
                run_V0 = True
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.run_V0... Finished...")
                Exit Function
            'error handling
Dth_error_run_check_not_passed:
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

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

                                                                        'Get Statements
                                                                        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Public Function get_size_V0(Optional more_instructions As String) As Variant
    'Created By (Zachary Daugherty)(11/24/20)
    'Purpose Case & notes:
        If (more_instructions = "help") Then
            Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "DTH_VA.get_size_V0: Help File Triggered..." & Chr(13) & _
                "What is this function for? (updated:12/04/2020):" & Chr(13) & _
                "    get_size_V0 is used to return the size of the DTH table length as well as audit out any blank space through out the" & Chr(13) & _
                "    table." & Chr(13) & _
                "Should i call this function directly? (updated:12/04/2020):" & Chr(13) & _
                "    calling this function directly will work but all this will do is give the table size and clean it of blankspace." & Chr(13) & _
                "What is returned from this function? (updated:12/04/2020):" & Chr(13) & _
                "    the function will return a number equal to the post audited length of the page" & Chr(13) & _
                "Listing Off dependants of Function (updated:12/04/2020):..." & Chr(13) & _
                "    DTH_VA." & Chr(13) & _
                "        DTH_VA.|parent module|" & Chr(13) & _
                "    Boots_Main_V_alpha." & Chr(13) & _
                "        Boots_Main_V_alpha.get_username" & Chr(13) & _
                "        Boots_Main_V_alpha.get_sheet_list" & Chr(13) & _
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
            get_size_V0 = "DTH_VA.get_size_V0 - Public - Stable 12/04/2020 - help file:Y"
            Exit Function
        End If
        
        '<debug note>
            Call Boots_Report_v_Alpha.Log_Push(text, "-------------------------------------------------------------------------")
                Call Boots_Report_v_Alpha.Log_Push(text, "The Calling of 'DTH_VA.get_size_V0' is not properly setup for dev notes yet please fix: missing error reporting")
                Call Boots_Report_v_Alpha.Log_Push(text, "The Calling of 'DTH_VA.get_size_V0' is not properly setup for dev notes yet please fix: missing error reporting")
                Call Boots_Report_v_Alpha.Log_Push(text, "The Calling of 'DTH_VA.get_size_V0' is not properly setup for dev notes yet please fix: missing error reporting")
                Call Boots_Report_v_Alpha.Log_Push(text, "The Calling of 'DTH_VA.get_size_V0' is not properly setup for dev notes yet please fix: missing error reporting")
            Call Boots_Report_v_Alpha.Log_Push(text, "-------------------------------------------------------------------------")
        '<end of debug note>
          
    'code start
        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.get_size_V0 Start...")
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
get_size_V0_restart:                'goto flag
        'setup variables
            Call Boots_Report_v_Alpha.Log_Push(text, "Setting Up Variables... ")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            Set wb = ActiveWorkbook
            Set home_pos = ActiveSheet
            On Error GoTo DTH_get_cant_find_DTH_SHEET   'goto error handler
                Set current_sht = wb.Sheets("DTH")      'setting name
            On Error GoTo 0                             'returns error handler to default
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'move to start location
            Call Boots_Report_v_Alpha.Log_Push(text, "Moving to Starting position...")
            row = DTH_POS_1A.DTH_Part_Number_ROW      'fetching indexed information from enumeration
            col = DTH_POS_1A.DTH_part_number_col      'fetching indexed information from enumeration
            s = current_sht.Cells(row, col).value     'fetching indexed information from enumeration
        'get lenght to bottom
            Call Boots_Report_v_Alpha.Log_Push(text, "Fetching the length to the botom of the table...")
            On Error GoTo DTH_cant_find_goalpost                    'goto error handler
                dist_to_goalpost = Range("DTH_GOALPOST").row - row  'setting definition
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
                    For i_2 = 1 To DTH_POS_1A.DTH_Q_number_pos
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
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
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
                Call Boots_Report_v_Alpha.Log_Push(text, "Removing Blank space from the DTH page... Start...")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'hide updating
                    Application.ScreenUpdating = False
                    Application.DisplayAlerts = False
                'move to start location
                    Call Boots_Report_v_Alpha.Log_Push(text, "Moving to starting location...")
                    row = DTH_POS_1A.DTH_Part_Number_ROW
                    col = DTH_POS_1A.DTH_part_number_col
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
                                On Error GoTo DTH_get_cant_find_DTH_SHEET
                                    Set current_sht = wb.Sheets("DTH")
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
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                'restart if things were deleted, reset some variables then goto 'get_size_V0_restart'
                    
                    Call Boots_Report_v_Alpha.Push_notification_message("DEVNOTE:'DTH_VA.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    
                    Call Boots_Report_v_Alpha.Log_Push(Flag)
                    Call Boots_Report_v_Alpha.Log_Push(text, "DEVNOTE:'DTH_VA.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    Call Boots_Report_v_Alpha.Log_Push(text, "DEVNOTE:'DTH_VA.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    Call Boots_Report_v_Alpha.Log_Push(text, "DEVNOTE:'DTH_VA.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    Call Boots_Report_v_Alpha.Log_Push(text, "DEVNOTE:'DTH_VA.GET_SIZE_V0' NEED TO UPDATE ERROR ANTILOOP TRIGGERED TO MODERN CALLING PROCEDURE")
                    Call Boots_Report_v_Alpha.Log_Push(table_close)
                    
ActiveWorkbook.Sheets("LOG_" & Boots_Main_V_alpha.get_username).visible = -1
Stop 'error code test for indent see inside if statement
                    
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
                                Call Boots_Report_v_Alpha.Log_Push(text, "Removing Blank space from the DTH page... Abandoned...")
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
                    Call Boots_Report_v_Alpha.Log_Push(text, "Removing Blank space from the DTH page... Finish...")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'cleanup
                Call Boots_Report_v_Alpha.Log_Push(text, "Removing Blank space from the DTH page... Finish...")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            End If
            'get final size
                get_size_V0 = dist_to_goalpost
    'cleanup
        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.get_size_V0 Finish...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    'code end
        Exit Function
    'error handling
DTH_get_cant_find_DTH_SHEET:
        'DTH_get_cant_find_DTH_SHEET:
            'set error report
                Call Boots_Report_v_Alpha.Log_Push(Error_)
                Call Boots_Report_v_Alpha.Log_Push(Flag)
                    Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: DTH_vx.Get_size_v0 was unable to locate the DTH sheet please check the enviorment & Log...")
                Call Boots_Report_v_Alpha.Log_Push(table_close)
            'generate required information
                Call Boots_Report_v_Alpha.Log_Push(text, "generating sheet list...")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    Call Boots_Main_V_alpha.get_sheet_list
            'push generated information
                Call Boots_Report_v_Alpha.Log_Push(table_open)
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    For z = 1 To (wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).End(xlDown).row - wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).row)
                        Call Boots_Report_v_Alpha.Log_Push(Variable, "Sheet List: " & wb.Sheets("boots").Cells(boots_pos.p_sheet_name_row + z, boots_pos.p_sheet_name_col).value & _
                            " ='visible stat': " & wb.Sheets("boots").Cells(boots_pos.p_sheet_visible_status_row + z, boots_pos.p_sheet_visible_status_col).value)
                    Next z
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
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
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Call Boots_Report_v_Alpha.Log_Push(table_close)
                'indent out
                    For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Next z
                'call end statement
                    Call Boots_Report_v_Alpha.Log_Push(Display_now)
                    End
DTH_cant_find_goalpost:
        'DTH_cant_find_goalpost
            'set error report
                Call Boots_Report_v_Alpha.Log_Push(Error_)
                Call Boots_Report_v_Alpha.Log_Push(Flag)
                    Call Boots_Report_v_Alpha.Log_Push(text, "FATAL ERROR!: DTH_vx.Get_size_v0 was unable to return / find the DTH page Goalpost...")
                    Call Boots_Report_v_Alpha.Log_Push(text, "Please check the Range mannager in the Workbook for its existance...")
                Call Boots_Report_v_Alpha.Log_Push(table_close)
            'listing local variables
                Call Boots_Report_v_Alpha.Log_Push(text, "Listing Snapshot of some variables...")
                Call Boots_Report_v_Alpha.Log_Push(table_open)
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "Wb: '" & wb.path & "\" & wb.Name & "' as workbook")
                    Call Boots_Report_v_Alpha.Log_Push(Variable, "home_pos: '" & home_pos.Parent.path & "\" & home_pos.Parent.Name & " == " & home_pos.index & ": " & home_pos.Name & "' as worksheet")
                Call Boots_Report_v_Alpha.Log_Push(table_close)
            'prep for post
                'close error table
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Call Boots_Report_v_Alpha.Log_Push(table_close)
                'indent out
                    For z = 1 To Boots_Report_v_Alpha.Log_get_indent_value_V0
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Next z
                'call end statement
                    Call Boots_Report_v_Alpha.Log_Push(Display_now)
                    End
End Function

Private Function Run_DTH_unit_cost_refresh_v0(Optional more_instructions As String) As Variant
    'Created By (Zachary Daugherty)(8/25/20)
    'Purpose Case & notes:
    Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "DTH_VA.Run_DTH_unit_cost_refresh_v0: devnote need to add in help file still coding...")
    
        If (more_instructions = "help") Then
            Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "DTH_VA.Run_DTH_unit_cost_refresh_v0: Help File Triggered...")
            Exit Function
        End If
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            Run_DTH_unit_cost_refresh_v0 = "DTH_VA.Run_DTH_unit_cost_refresh_v0 - Private - in development 12-09-2020 - Help file:N"
            Exit Function
        End If
    'code start
            Call Boots_Report_v_Alpha.Log_Push(text, "DTh_va.Run_DTH_unit_cost_refresh_v0... Starting...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
        'call protection
            If (more_instructions <> "EbjtnI9SqGwmO8miHIjS") Then 'random string
                Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "DTH_VA.Run_DTH_unit_cost_refresh_v0: Call Protection triggered..." & Chr(13) & _
                "DTH_VA.Run_DTH_unit_cost_refresh_v0 did not have the correct key sent to run this command this is in place to prevent this function being called by accident" & Chr(13) & _
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
                Dim HP_decoder_A() As String
                Dim HP_decoder_B() As String
                Dim Lookup() As String
                Dim size_of_dth As Long
                Dim size_of_HP_A As Long
                Dim size_of_HP_B As Long
            'globals
                Dim DTH_inflation_Value As Double
                Dim 'sp_global_structural_value to remove As Double
                Dim 'sp_global_plate_value to remove As Double
            'containers
                Dim L As Long
                Dim L_2 As Long
                Dim s As String
                Dim condition As Boolean
                Dim error As String
                Dim anti_loop As Long
        'setup variables
            Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Setting up variables... Start...")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'setup pos
                Set wb = ActiveWorkbook
                Set home_pos = ActiveSheet
                On Error GoTo DTH_Run_DTH_unit_cost_refresh_v0_cant_find_SHEET  'goto error handler
                    error = "DTS"
                    Set current_sht = wb.Sheets(error)                      'setting name
                    error = ""
                On Error GoTo 0                                             'returns error handler to default
                row = -1
                col = -1
                s = "empty"
                L = -1
                L_2 = -1
                'sp_global_structural_value to remove = -1
                'sp_global_plate_value to remove = -1
                DTH_inflation_Value = -1
                size_of_dth = -1
                size_of_HP_A = -1
                size_of_HP_B = -1
            'setup array tables
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Array table setup... start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'check for valid sizes and get sizes
                    'get size of DTS
                        Call Boots_Report_v_Alpha.Log_Push(text, "Fetching size of DTH TABLE...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        Stop ' debug need to test that the value is returned correctly
                        size_of_dth = DTH_VA.get_size_V0()
                        Stop ' end of debug
                    'get size of hardware presets A
                        Call Boots_Report_v_Alpha.Log_Push(text, "Fetching size of Hardware Presets Table A...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        Stop ' debug need to test that the value is returned correctly
                        size_of_HP_A = HP_V3_stable.Get_size_HP_A_V1 'old spot: SP_V1_DEV.get_size_A
                        Stop ' end of debug
                    'get size of hardware presets B
                        Call Boots_Report_v_Alpha.Log_Push(text, "Fetching size of Hardware Presets Table B...")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        Stop ' debug need to test that the value is returned correctly
                        size_of_HP_B = HP_V3_stable.Get_size_HP_B_V1 'old spot: SP_V1_DEV.get_size_B_V1
                        Stop ' end of debug
                'initialize arrays
                    Call Boots_Report_v_Alpha.Log_Push(text, "Initalizing Storage for Tables: Memory Main, HP_decoder_A , HP_decoder_B, Lookup...")
                    Stop ' debug need to test that the size is correctly as dth_pos return
                    ReDim Memory_Main(0 To size_of_dth, 0 To DTH_POS_1A.DTH_Q_total_number_of_tracked_locations) 'old: DTS_POS_2A.DTS_Q_number_of_tracked_locations)
                    Stop ' end of debug
                    Stop ' debug need to test that the size is correctly as dth_pos return
                    ReDim HP_decoder_A(0 To size_of_HP_A, 1 To HP_POS_1.Q_HP_TABLE_A_NUMBER_OF_TRACKED_POSITIONS) 'old spot: SP_POS.SP_Q_Number_total_A_Tracked_Cells)
                    Stop ' end of debug
                    Stop ' debug need to test that the size is correctly as dth_pos return
                    ReDim HP_decoder_B(0 To size_of_HP_B, 1 To HP_POS_1.Q_HP_TABLE_B_NUMBER_OF_TRACKED_POSITIONS) 'old spot: SP_POS.SP_Q_Number_total_B_Tracked_Cells)
                    Stop ' end of debug
                    'lookup setup
                        Stop ' debug need to test that the size is correctly as dth_pos return
                        ReDim Lookup(0 To size_of_HP_A + size_of_HP_B, 0 To 3)
                        'address assignment
                            Lookup(0, 0) = "Lookup Code"
                            Lookup(0, 1) = "Sheet its on"
                            Lookup(0, 2) = "What table its on"
                            Lookup(0, 3) = "address in array"
                        Stop ' end of debug
                'end of array tables
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Array table setup... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'setup globals
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Setting Up Global values... start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'dth
                    Call Boots_Report_v_Alpha.Log_Push(text, "Get DTH Sheet... Inflation Values")
                    On Error GoTo DTH_Run_DTH_unit_cost_refresh_v0_cant_find_SHEET
                        error = "DTH"
                        Set current_sht = wb.Sheets(error)
                        error = ""
                    On Error GoTo 0
                    Stop ' debug need to test that the value is returned correctly
                    DTH_inflation_Value = current_sht.Cells(DTH_POS_1A.DTH_Inflation_Const_ROW, DTH_POS_1A.DTH_Inflation_Const_ROW).value ' old value: DTS_POS_2A.DTS_I_Inflation_Const_row, DTS_POS_2A.DTS_I_Inflation_Const_col).value
                    Stop ' end of debug
'                'SP
'                    Call Boots_Report_v_Alpha.Log_Push(text, "Get SP Sheet... Structural value & Plate value")
'                    On Error GoTo DTH_Run_DTH_unit_cost_refresh_v0_cant_find_SHEET
'                        error = "STEEL PRESETS"
'                        Set current_sht = wb.Sheets(error)
'                        error = ""
'                    On Error GoTo 0
'                    'sp_global_structural_value to remove = current_sht.Cells(SP_POS.SP_I_Const_Structural_row, SP_POS.SP_I_Const_Structural_col).value
'                    'sp_global_plate_value to remove = current_sht.Cells(SP_POS.SP_I_Const_Plate_row, SP_POS.SP_I_Const_Plate_col).value
                'return cursor to home
                    home_pos.Activate
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Setting Up Global values... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'setup complete
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Setting Up Variables.. Finished")
        'load tables
            Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Load table Info... start")
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'dts main
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Loading DTH Information to array... Start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'set focus
                    On Error GoTo DTH_Run_DTH_unit_cost_refresh_v0_cant_find_SHEET
                        error = "DTH"
                            Set current_sht = wb.Sheets(error)
                        error = ""
                        'set row col
                        Stop ' debug need to test that the value is returned correctly
                            row = DTH_POS_1A.DTH_Part_Number_ROW 'old spot: DTS_POS_2A.DTS_I_part_number_row
                            col = DTH_POS_1A.DTH_part_number_col 'old spot: DTS_POS_2A.DTS_I_part_number_col
                        Stop ' end of debug
                    On Error GoTo 0
                'get
                Stop ' debug need to test that the value is returned correctly
                    For L = 0 To size_of_dth
                        For L_2 = 1 To DTH_POS_1A.DTH_Q_total_number_of_tracked_locations 'old spot: DTS_POS_2A.DTS_Q_number_of_tracked_locations
                            Memory_Main(L, L_2) = current_sht.Cells(row, col).value
                            col = DTH_POS_1A.DTH_part_number_col + L_2 'old spot: DTS_POS_2A.DTS_I_part_number_col + L_2
                        Next L_2
                        Stop ' debug need to test that the value is returned correctly
                        row = DTH_POS_1A.DTH_part_number_col + L + 1 'old post: DTS_POS_2A.DTS_I_part_number_row + L + 1
                        col = DTH_POS_1A.DTH_part_number_col         'old post: DTS_POS_2A.DTS_I_part_number_col
                        Stop ' end of debug
                    Next L
                Stop ' end of debug
                'cleanup
                    row = -1
                    col = -1
                    L = -1
                    L_2 = -1
                    Set current_sht = Nothing
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Loading DTH Information to array... Finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'HP_decoder_A
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Loading Steel Presets A Information to array... Start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'set focus
                    On Error GoTo DTH_Run_DTH_unit_cost_refresh_v0_cant_find_SHEET
                        error = "Hardware Presets"
                            Set current_sht = wb.Sheets(error)
                        error = ""
                        'set row col
                            Stop ' debug need to test that the value is returned correctly
                            row = HP_POS_1.A_HP_GENERAL_PREFIX_row 'old post: SP_POS.SP_I_A_Prefix_row
                            col = HP_POS_1.A_HP_GENERAL_PREFIX_col 'old post: SP_POS.SP_I_A_Prefix_col
                            Stop ' end of debug
                    On Error GoTo 0
                'get
                    Stop ' debug need to test that the value is returned correctly
                    For L = 0 To size_of_HP_A
                        For L_2 = 1 To HP_POS_1.Q_HP_TABLE_A_NUMBER_OF_TRACKED_POSITIONS 'old post: SP_POS.SP_Q_Number_total_A_Tracked_Cells
                            HP_decoder_A(L, L_2) = current_sht.Cells(row, col).value
                            Stop ' debug need to test that the value is returned correctly
                            col = HP_POS_1.A_HP_GENERAL_PREFIX_col + L_2 'old post: SP_POS.SP_I_A_Prefix_col + L_2
                            Stop ' end of debug
                        Next L_2
                    Stop ' debug need to test that the value is returned correctly
                    row = HP_POS_1.A_HP_GENERAL_PREFIX_row + L + 1 'old post: SP_POS.SP_I_A_Prefix_row + L + 1
                    col = HP_POS_1.A_HP_GENERAL_PREFIX_col         'old post: SP_POS.SP_I_A_Prefix_col
                    Next L
                    Stop ' end of debug
                'cleanup
                    row = -1
                    col = -1
                    L = -1
                    L_2 = -1
                    Set current_sht = Nothing
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Loading Steel Presets A Information to array... Finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'HP_decoder_B
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Loading Steel Presets B Information to array... Start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'set focus
                    
                    On Error GoTo DTH_Run_DTH_unit_cost_refresh_v0_cant_find_SHEET
                        Stop ' debug need to test that the value is returned correctly
                        error = "Hardware Presets"
                            Set current_sht = wb.Sheets(error)
                        error = ""
                        Stop ' end of debug
                        'set row col
                            Stop ' debug need to test that the value is returned correctly
                            row = HP_POS_1.B_HP_PROPRIETARY_PART_NUMBER_row 'old post: SP_POS.SP_I_B_Prefix_row
                            col = HP_POS_1.B_HP_PROPRIETARY_part_number_col 'old post: SP_POS.SP_I_B_Prefix_col
                            Stop ' end of debug
                    On Error GoTo 0
                'get
                    Stop ' debug need to test that the value is returned correctly
                    For L = 0 To size_of_HP_B
                        For L_2 = 1 To SP_POS.SP_Q_Number_total_B_Tracked_Cells
                            HP_decoder_B(L, L_2) = current_sht.Cells(row, col).value
                            col = SP_POS.SP_I_B_Prefix_col + L_2
                        Next L_2
                    row = SP_POS.SP_I_B_Prefix_row + L + 1
                    col = SP_POS.SP_I_B_Prefix_col
                    Next L
                    Stop ' end of debug
                'cleanup
                    row = -1
                    col = -1
                    L = -1
                    L_2 = -1
                    Set current_sht = Nothing
                    home_pos.Activate
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Loading Steel Presets B Information to array... Finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'assemble lookup table
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: assemble lookup table Information to array... Start")
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'initialize variables
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: assemble lookup table initialize variables... Start")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    L = 0
                    L_2 = 1
                    line = 0
                    s = ""
                    'Fetching lookup array dimension information
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Fetching lookup array dimension information... matrix_V2.matrix_dimensions_v1... start")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = matrix_V2.matrix_dimensions_v1(Lookup(), "d_report")
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Fetching lookup array dimension information... matrix_V2.matrix_dimensions_v1... finish")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    'parse matrix dim
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: resolve dimension information... start")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        Call Boots_Report_v_Alpha.Log_Push(text, "resolve dimension information (step1-3).... String_V1.Disassociate_by_Char_V2 running...")
                            s = String_V1.Disassociate_by_Char_V2("<", s, Right_C, "d_report")
                        Call Boots_Report_v_Alpha.Log_Push(text, "resolve dimension information (step2-3).... String_V1.Disassociate_by_Char_V2 running...")
                            s = String_V1.Disassociate_by_Char_V2("<", s, Right_C, "d_report")
                        Call Boots_Report_v_Alpha.Log_Push(text, "resolve dimension information (step3-3).... String_V1.Disassociate_by_Char_V2 running...")
                            line = CLng(String_V1.Disassociate_by_Char_V2(">", s, Left_C, "d_report"))
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: resolve dimension information... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    'cleanup
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: assemble lookup table initialize variables... Finish")
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                'Filling lookup table with all possible values and table positional data...
DTH_incoding_of_table_names:
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Filling lookup table with all possible values and table positional data... start")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    For L = 0 To line
                        'check to see where table data should be grabed from, note will fill in table A first then B
                            'section for table a
                                If (L > 0) Then
                                    If (L <= size_of_HP_A - 1) Then '-1 is included on the end to skip the goalpost of the table
                                        Lookup(L, 0) = HP_decoder_A(L, 1)
                                                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: fetching sheet name: HP_V3_stable.get_HP_sheet_name_v1 start")
                                                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                                    Stop ' debug need to test that the value is returned correctly
                                                    Lookup(L, 1) = HP_V3_stable.get_HP_sheet_name_v1("d_report") 'old post: SP_V1_DEV.get_sheet_name("d_report")
                                                    Stop ' end of debug
                                                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: fetching sheet name: HP_V3_stable.get_HP_sheet_name_v1 finish")
                                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                        Lookup(L, 2) = "A"
                                        Lookup(L, 3) = L
                                    Else
                                        'fall through marker
                                            Lookup(L, 0) = DTH_VA.get_global_unit_cost_refresh_ignore_trigger
                                            Lookup(L, 1) = DTH_VA.get_global_unit_cost_refresh_ignore_trigger
                                            Lookup(L, 2) = DTH_VA.get_global_unit_cost_refresh_ignore_trigger
                                            Lookup(L, 3) = DTH_VA.get_global_unit_cost_refresh_ignore_trigger
                                    End If
                                End If
                            'section for table b
                                If (L > 0) Then
                                    If ((L > size_of_HP_A) And (L < size_of_HP_A + size_of_HP_B)) Then
                                        L_2 = L - size_of_HP_A
                                        Lookup(L, 0) = HP_decoder_B(L_2, 1)
                                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: fetching sheet name: HP_V3_stable.get_HP_sheet_name_v1 start")
                                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                            Stop ' debug need to test that the value is returned correctly
                                            Lookup(L, 1) = HP_V3_stable.get_HP_sheet_name_v1("d_report") 'old post: SP_V1_DEV.get_sheet_name("d_report")
                                            Stop ' end of debug
                                            Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: fetching sheet name: HP_V3_stable.get_HP_sheet_name_v1 finish")
                                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                        Lookup(L, 2) = "B"
                                        Lookup(L, 3) = L_2
                                    Else
                                        'fall through marker
                                            If (L > size_of_HP_A) Then
                                                Lookup(L, 0) = DTH_VA.get_global_unit_cost_refresh_ignore_trigger
                                                Lookup(L, 1) = DTH_VA.get_global_unit_cost_refresh_ignore_trigger
                                                Lookup(L, 2) = DTH_VA.get_global_unit_cost_refresh_ignore_trigger
                                                Lookup(L, 3) = DTH_VA.get_global_unit_cost_refresh_ignore_trigger
                                            End If
                                    End If
                                End If
                    Next L
                'cleanup
                    L = -1
                    L_2 = -1
                    s = "empty"
                    line = -1
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Filling lookup table with all possible values and table positional data... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: assemble lookup table Information to array... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Load table Info... finish")
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'check lookup table for duplicate entrys
            Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: check lookup table for duplicate entrys... start") '2-3
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'initialize variable
                L = 0
                L_2 = 0
                s = ""
                condition = False
            'start
                'Get size of the matrix
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix... start") '3-4
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: matrix_V2.matrix_dimensions_v1... |1-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = matrix_V2.matrix_dimensions_v1(Lookup(), "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: matrix_V2.matrix_dimensions_v1... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... |2-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = String_V1.Disassociate_by_Char_V2(">", s, Right_C, "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... |3-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = String_V1.Disassociate_by_Char_V2(">", s, Right_C, "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... |4-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            s = String_V1.Disassociate_by_Char_V2("<", s, Right_C, "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                            
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... |5-5| start") '4-5
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            line = String_V1.Disassociate_by_Char_V2(">", s, Left_C, "d_report")
                            Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix: String_V1.Disassociate_by_Char_V2... finish") '5-4
                            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                    'cleanup
                        s = "Empty"
                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: Get size of the matrix... finish") '4-3
                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        
                'loop through lookup table
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: loop through lookup table for duplicate entrys... start") '3-4
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    
                    For L = 1 To (line - 1)
                        'value to search by
                            'if set to empty skip
                                If (Lookup(L, 1) = DTH_VA.get_global_unit_cost_refresh_ignore_trigger) Then
                                    GoTo Run_DTH_unit_cost_refresh_v0_ignore_entry
                                End If
                        'compair against all other entrys
                            For L_2 = 1 To (line - 1)
                                'check to see if index is the same if so skip check
                                    If (L = L_2) Then
                                        GoTo Run_DTH_unit_cost_refresh_v0_skip_check
                                    End If
                                'set value to check thru
                                    s = Lookup(L, 0)
                                'do check
                                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: loop through lookup table for duplicate entrys match possible: String_V1.is_same_V1... start") '4-5
                                        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                        condition = String_V1.is_same_V1(s, Lookup(L_2, 0), "d_report")
                                        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: loop through lookup table for duplicate entrys match possible: String_V1.is_same_V1... finish") '5-4
                                        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                'if condition is true then throw error
                                    If (condition = True) Then
                                        error = "in array: lookup:(" & L_2 & ",0) value:'" & Lookup(L_2, 0) & "'. is the same as the value in: lookupL(" & L & ",0)"
                                        GoTo Run_DTH_unit_cost_refresh_v0_duplicate_lookups
                                    Else
                                        
                                    End If
                                'goto
Run_DTH_unit_cost_refresh_v0_skip_check:
                            Next L_2
                        'goto
Run_DTH_unit_cost_refresh_v0_ignore_entry:
                    Next L
                    Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: loop through lookup table for duplicate entrys... finish") '4-3
                    Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'cleanup
                L = -1
                L_2 = -1
                s = "empty"
                line = -1
                condition = False
                Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.Run_DTH_unit_cost_refresh_v0: check lookup table for duplicate entrys... finish") '3-2
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'do update
            Call Boots_Report_v_Alpha.Log_Push(text, "DTh_va.Run_DTH_unit_cost_refresh_v0: run update Starting...") '2-3
            Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            'initalize variables
                On Error GoTo DTH_Run_DTH_unit_cost_refresh_v0_cant_find_SHEET
                    Stop ' debug need to test that the value is returned correctly
                    error = "DTH"
                    Set current_sht = wb.Sheets(error)
                    Stop ' end of debug
                    error = ""
                On Error GoTo 0
                Stop ' debug need to test that the value is returned correctly
                row = DTH_POS_1A.DTH_Part_Number_ROW 'old post: DTS_POS_2A.DTS_I_part_number_row
                col = DTH_POS_1A.DTH_part_number_col 'old post: DTS_POS_2A.DTS_I_part_number_col
                Stop ' end of debug
                L = 0
                L_2 = 0
            'iterate thru memory main
                Call Boots_Report_v_Alpha.Log_Push(text, "DTh_va.Run_DTH_unit_cost_refresh_v0: run update: iterating thru memory main to find matches and return values...") '3-na
                For L = 1 To size_of_dth
                    'change pos
                        Stop ' debug need to test that the value is returned correctly
                        row = DTH_POS_1A.DTH_Part_Number_ROW + L 'old post: DTS_POS_2A.DTS_I_part_number_row + L
                        Stop ' end of debug
                    'set smart code
                        s = Memory_Main(L, 2)
                    'check for empty or ignore trigger
                        If ((s <> DTH_VA.get_global_unit_cost_refresh_ignore_trigger) And (s <> "")) Then
                            'decode smart code
Run_DTH_unit_cost_refresh_v0_part_numb_check:
                                Call Boots_Report_v_Alpha.Log_Push(text, "DTh_va.Run_DTH_unit_cost_refresh_v0: run update: String_V1.Disassociate_by_Char_V2 start...") '3-4
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                    s = String_V1.Disassociate_by_Char_V2(DTH_VA.get_global_decoder_symbol, s, Left_C, "d_report")
                                    Call Boots_Report_v_Alpha.Log_Push(text, "DTh_va.Run_DTH_unit_cost_refresh_v0: run update: String_V1.Disassociate_by_Char_V2 finish...") '4-3
                                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                            'search for key in lookup array
                                For L_2 = 1 To (size_of_HP_A + size_of_HP_B)
                                    If (s = Lookup(L_2, 0)) Then
                                        'match found return value to sheet
                                            'locate which chart
                                                If (Lookup(L_2, 2) = "A") Then
                                                    'match found in decode table 'A'
                                                        'return value to dts
                                                            On Error GoTo DTH_Run_DTH_unit_cost_refresh_v0_cant_find_range
                                                                Stop ' debug need to test that the value is returned correctly
                                                                error = "DTH_UNIT_COST" 'this is the col pos that the value is returned to the row is decided by 'L'
                                                                    current_sht.Range(error).Offset(L, 0).value = HP_decoder_A(CLng(Lookup(L_2, 3)), 4) 'paste to sheet name <current_sht> then move cursor to range <error> offset down to pos <L>: to get the value find in array <lookup> and return address of the match. convert to <long> variable and then user that long to look in array <HP_decoder_a> at the value of <long> then return that value to sheet
                                                                error = ""
                                                                Stop ' end of debug
                                                            On Error GoTo 0
                                                Else
                                                    If (Lookup(L_2, 2) = "B") Then
                                                        'match found in decode table 'B'
                                                            'return value to dts
                                                                On Error GoTo DTH_Run_DTH_unit_cost_refresh_v0_cant_find_range
                                                                    Stop ' debug need to test that the value is returned correctly
                                                                    error = "DTH_UNIT_COST" 'this is the col pos that the value is returned to the row is decided by 'L'
                                                                    current_sht.Range(error).Offset(L, 0).value = HP_decoder_B(CLng(Lookup(L_2, 3)), 4) 'paste to sheet name <current_sht> then move cursor to range <error> offset down to pos <L>: to get the value find in array <lookup> and return address of the match. convert to <long> variable and then user that long to look in array <HP_decoder_a> at the value of <long> then return that value to sheet
                                                                    error = ""
                                                                    Stop ' end of debug
                                                                On Error GoTo 0
                                                    Else
                                                        error = CStr(Lookup(L_2, 2))
                                                        GoTo dth_Run_DTH_unit_cost_refresh_v0_cant_locate_table
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
                                        GoTo Run_DTH_unit_cost_refresh_v0_part_numb_check
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
                Call Boots_Report_v_Alpha.Log_Push(text, "DTh_va.Run_DTH_unit_cost_refresh_v0: run update finish...") '3-2
                Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
    'code end
        Run_DTH_unit_cost_refresh_v0 = True
        Call Boots_Report_v_Alpha.Log_Push(text, "DTh_va.Run_DTH_unit_cost_refresh_v0... finish...") '2-1
            Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Exit Function
    'error handling
DTH_Run_DTH_unit_cost_refresh_v0_cant_find_SHEET:
        'DTH_Run_DTH_unit_cost_refresh_v0_cant_find_SHEET
            Run_DTH_unit_cost_refresh_v0 = False
                Call Boots_Report_v_Alpha.Log_Push(Error_)
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL Error: DTH_VA: sub: Run_DTH_unit_cost_refresh_v0: was unable to find the sheet named '" & error & "', please check your code.")
                Call Boots_Report_v_Alpha.Log_Push(text, "Displaying Snapshot of Values:...")
                'table
                    Call Boots_Report_v_Alpha.Log_Push(table_open)
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
Run_DTH_unit_cost_refresh_v0_duplicate_lookups:
        'Run_DTH_unit_cost_refresh_v0_duplicate_lookups
            Run_DTH_unit_cost_refresh_v0 = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL Error: DTh_va: Function: Run_DTH_unit_cost_refresh_v0:During the assembly '" & error & "' please make the nessasary changes to the tables to not have duplicate values")
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
DTH_Run_DTH_unit_cost_refresh_v0_cant_find_range:
        'DTH_Run_DTH_unit_cost_refresh_v0_cant_find_range:
            Run_DTH_unit_cost_refresh_v0 = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL Error: DTh_va: Function: Run_DTH_unit_cost_refresh_v0: Range(" & error & ") was unable to be located")
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
dth_Run_DTH_unit_cost_refresh_v0_cant_locate_table:
        'dth_Run_DTH_unit_cost_refresh_v0_cant_locate_table:
            Run_DTH_unit_cost_refresh_v0 = False
                Call Boots_Report_v_Alpha.Log_Push(Error_, "")
                Call Boots_Report_v_Alpha.Log_Push(text, "FATAL Error: DTh_va: Function: Run_DTH_unit_cost_refresh_v0:(see next line)")
                Call Boots_Report_v_Alpha.Log_Push(text, "Function was unable to locate the table named:'" & error & "'(see next line)")
                Call Boots_Report_v_Alpha.Log_Push(text, "Please see the goto 'DTH_incoding_of_table_names' as this is where the table chars are assigned")
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

                                                                        'Utility Statements
                                                                        
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Private Function check_table_operation_V0(Optional more_instructions As String) As Variant
    'Created By (Zachary Daugherty)(12/2/2020)
    'Purpose Case & notes:
        'this function is a check function on if data is stored in the proper place before updates or data is manipulated for the DTH page
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
            check_table_operation_V0 = "DTH_VA.check_table_operation_V0 - stable but missing table checks 12/02/2020 - no help file"
            Exit Function
        End If
        Boots_Report_v_Alpha.Push_notification_message ("DTH_VA.check_table_operation_V0 DEVNOTE: need to add proper devnotes and proper log reporting" & String_V1.get_Special_Char_V1(carriage_return, True) & _
            "currently this function only calls and returns log information about starting and stoping...")
        Stop
    'check for dont_show_information
'        If (dont_show_information = False) Then
'            MsgBox ("_________________String_Vx.check instructions_________________" & String_V1.get_Special_Char_V1(carriage_return, True) & _
'            "function is called to make sure all the keystone locations of the data set are anchored to the right positions. data sets that should be checked are the " & _
'            "following: DTS Tables and SP Tables. if either of these do not pass checks there will be errors generated to fix the positional data of the file")
'            Stop
'            Exit Function
'        End If
    'code start
        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.check_table_operation_V0 Starting...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
        'define variables
            Dim condition As Boolean
        'run
        
            MsgBox ("need to add DTH table CHECKS")
            Stop
            'Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'condition = DTS_V2A.Check_DTS_Table_V0_01A
            If (condition = True) Then
                condition = False
                Stop 'need to verify function below
                'Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    'condition = HP_V3_stable.DO_Check_HP_A_Table_V1
            End If
            If (condition = True) Then
                condition = False
                Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                    condition = HP_V3_stable.Do_Check_HP_B_Table_V1 ' code was verifyed as of 12/02/2020
            End If
            
            check_table_operation_V0 = condition
    'code end
        Call Boots_Report_v_Alpha.Log_Push(text, "DTH_VA.check_table_operation_V0 Finishing...")
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)   'to escape local master
        Call Boots_Report_v_Alpha.Log_Push(Trigger_e)   'to fix indent
        Exit Function
    'error handle
        'na
    'end error handle
End Function
