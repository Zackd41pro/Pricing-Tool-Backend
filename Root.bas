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
    get_version = 1.1911
End Function
