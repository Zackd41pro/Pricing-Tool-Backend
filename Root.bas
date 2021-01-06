Attribute VB_Name = "root"
Public Function get_drive_location() As String
    get_drive_location = "C:\"
End Function

Public Function get_save_location() As String
    get_save_location = "ZEDVBA\"
End Function

Public Function get_project_name() As String
    get_project_name = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)) & "\" 'ActiveWorkbook.Name
End Function

Public Function get_version() As Double
    get_version = 1.301
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                                                
                                                                        'Startup Code

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Public Function On_startup() As Variant
    'dims
        'addresses
        Dim wb As Workbook
        Dim sht As Worksheet
        'containers
            Dim i As Long
            Dim j As Long
            Dim s As String
            Dim s_2 As String
            Dim arr() As Variant
        'dir
            Dim DIR_row As Long
            Dim DIR_col As Long
    'sets
        Set wb = ActiveWorkbook
    'start boots
        Boots_Main_V_alpha.run_on_start
    'create Master DIR
        Call Boots_Report_v_Alpha.DIR_Make(root.get_save_location, root.get_drive_location)
        Call Boots_Report_v_Alpha.DIR_Make(root.get_project_name, root.get_drive_location & root.get_save_location)
        'make version DIR
            Call Boots_Report_v_Alpha.DIR_Make(root.get_version & "\", root.get_drive_location & root.get_save_location & root.get_project_name)
            'make user DIR
                Call Boots_Report_v_Alpha.DIR_Make("Users\", root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\")
                    'make 'this' specific user DIR
                        Call Boots_Report_v_Alpha.DIR_Make(Boots_Main_V_alpha.get_username & "\", root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\" & "Users\")
    'devnote todo list
        
        Boots_Report_v_Alpha.Push_notification_message ("need to check out this code for geting functions and subs of a specific file 'https://stackoverflow.com/questions/2630872/how-to-get-the-list-of-function-and-sub-of-a-given-module-name-in-excel-vba'" & Chr(13) & _
        "https://www.vitoshacademy.com/vba-listing-all-procedures-in-all-modules/")
    
        Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "location of terminal push:'Root.on_startup' todos" & Chr(13) & _
            "DEVELOPER NOTES:" & Chr(13) & _
            "    |completed tasks are in another terminal message|" & Chr(13) & _
            "    List of todos:" & Chr(13) & _
            "        boots_report_v_alpha." & Chr(13) & _
            "            boots_report_v_alpha.Log_Push" & Chr(13) & _
            "                take steps to add in a level of details integer to change how detailed the logs are" & Chr(13) & _
            "        DTH_VA." & Chr(13) & _
            "            need to run tests for DTH_VA.run_V0 and the connected functions tied to " & Chr(13) & _
            "            need to run tests for DTH_VA.RUN_DTH_unit_cost_refresh_v0 (private function)" & Chr(13) & _
            "        THISWORKBOOK." & Chr(13) & _
            "            ADD PLEASE WAIT PAGE SO IT DOES NOT LOOK LIKE A ERROR IS HAPPENING DURING BOOT AND SHUTDOWN")
        Boots_Report_v_Alpha.Push_notification_message (Chr(13) & Chr(13) & "location of terminal push:'Root.on_startup' completed todos" & Chr(13) & _
            "        HP_V3_stable." & Chr(13) & _
            "            |completed 12/02/2020| HP_V3_stable.Do_Check_HP_B_Table_V1" & Chr(13) & _
            "                |completed 12/02/2020| change out the Error codes that are incorrect/old" & Chr(13) & _
            "            |completed 12/03/2020| HP_V3_stable.Check_HP_A_Table_VA" & Chr(13) & _
            "                |completed 12/03/2020| have Check_HP_A_Table_VA reflect the same changes as HP_V3_stable.Do_Check_HP_B_Table_V1 once B is completly updated" & Chr(13) & _
            "        boots_report_v_alpha." & Chr(13) & _
            "            |completed 12/04/2020| boots_report_v_alpha.Log_get_length_of_log_list_V1" & Chr(13) & _
            "                |completed 12/04/2020| need to change out how the way this is calculated as it takes a really long time" & Chr(13) & _
            "        DTH_VA." & Chr(13) & _
            "            |completed 12-16-2020| need to add DTH_VA. FUNCTION TO CHECK THE TABLE AS IT DOES NOT EXIST")
    'check for debug mode
        'create DIR search
            Boots_Report_v_Alpha.DIR_GET_vA (root.get_drive_location & root.get_save_location)
            Set sht = wb.Sheets(Boots_Main_V_alpha.get_username & "_DIR_Search")
            'get dir lengths
                DIR_row = sht.Range("A1").End(xlDown).row
                DIR_col = sht.Range("A1").End(xlToRight).Column
                If (sht.Cells(2, 1).value = "") Then
                    DIR_row = 1
                End If
            'set dir
                ReDim arr(1 To DIR_row, 1 To DIR_col)
            'load arr
                For i = 1 To DIR_row
                    For j = 1 To DIR_col
                        arr(i, j) = sht.Cells(i, j).value
                    Next j
                Next i
                'cleanup
                    i = 0
                    j = 0
        'search through
            'setup search value
                s = root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\" & "users\" & Boots_Main_V_alpha.get_username & "\debug.txt"
            'do search
                For i = 1 To DIR_row
                    s_2 = arr(i, 2)
                    If (UCase(s_2) = UCase(s)) Then
                        Boots_Report_v_Alpha.DIR_Flush
                        MsgBox ("Debug Mode Enabled: TO Disable remove:" & Chr(10) & s & Chr(10) & "From file")
                        End
                    End If
                Next i
                Boots_Report_v_Alpha.DIR_Flush
End Function




















