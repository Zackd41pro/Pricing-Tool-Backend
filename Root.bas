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
    get_version = 1.1912
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
        Dim arr() As Variant
    'sets
        Set wb = ActiveWorkbook

    'create Master DIR
        Call Boots_Main_V_alpha.Make_Dir(root.get_save_location, root.get_drive_location)
        Call Boots_Main_V_alpha.Make_Dir(root.get_project_name, root.get_drive_location & root.get_save_location)
        'make version DIR
            Call Boots_Main_V_alpha.Make_Dir(root.get_version & "\", root.get_drive_location & root.get_save_location & root.get_project_name)
            'make user DIR
                Call Boots_Main_V_alpha.Make_Dir("Users\", root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\")
                    'make 'this' specific user DIR
                        Call Boots_Main_V_alpha.Make_Dir(Boots_Main_V_alpha.get_username & "\", root.get_drive_location & root.get_save_location & root.get_project_name & root.get_version & "\" & "Users\")
    'check for debug mode
        Boots_Report_v_Alpha.DIR_GET_vA (root.get_drive_location & root.get_save_location)
        Set sht = wb.Sheets(Boots_Main_V_alpha.get_username & "_DIR_Search")
        ReDim arr(1 To sht.Range("A1").End(xlDown).row, 1 To sht.Range("A1").End(xlToRight).Column)
        MsgBox ("change the above function into a DIR function")
        Stop
End Function

