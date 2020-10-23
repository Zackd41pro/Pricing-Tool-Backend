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
    get_version = 1.2
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                                                
                                                                        'Startup Code

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Public Function On_startup() As Variant
    'dims
        Dim wb As Workbook
        Dim sht As Worksheet
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

    'create Master DIR
        Call Boots_Report_v_Alpha.DIR_Make(root.get_save_location, root.get_drive_location)
        Call Boots_Report_v_Alpha.DIR_Make(root.get_project_name, root.get_drive_location & root.get_save_location)
        'make version DIR
            Call Boots_Report_v_Alpha.DIR_Make(root.get_version & "\", root.get_drive_location & root.get_save_location & root.get_project_name)
            'make user DIR
                Call Boots_Report_v_Alpha.DIR_Make("Users\", root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\")
                    'make 'this' specific user DIR
                        Call Boots_Report_v_Alpha.DIR_Make(Boots_Main_V_alpha.get_username & "\", root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\" & "Users\")
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




















