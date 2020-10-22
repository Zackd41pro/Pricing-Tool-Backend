Attribute VB_Name = "Alpha_code"
Option Explicit

'https://www.excelhowto.com/macros/how-to-list-all-files-in-folder-and-sub-folders-use-excel-vba/
Sub ListAllFilesInAllFolders(ByVal path As String)
 
    Dim MyPath As String, MyFolderName As String, MyFileName As String
    Dim i As Integer, F As Boolean
    Dim objShell As Object, objFolder As Object, AllFolders As Object, AllFiles As Object
    Dim MySheet As Worksheet
     
    On Error Resume Next
     
    '************************
    'Select folder
    'Set objShell = CreateObject("Shell.Application")
    'Set objFolder = objShell.browseforfolder(0, "", 0, 0)
    'If Not objFolder Is Nothing Then
        MyPath = path & "\" 'objFolder.self.path & "\"
    'Else
        'Exit Sub
       'MyPath = "G:\BackUp\"
    'End If
    Set objFolder = Nothing
    Set objShell = Nothing
     
    '************************
    'List all folders
     
    Set AllFolders = CreateObject("Scripting.Dictionary")
    Set AllFiles = CreateObject("Scripting.Dictionary")
    AllFolders.Add (MyPath), ""
    i = 0
    Do While i < AllFolders.count
        Key = AllFolders.Keys
        MyFolderName = Dir(Key(i), vbDirectory)
        Do While MyFolderName <> ""
            If MyFolderName <> "." And MyFolderName <> ".." Then
                If (GetAttr(Key(i) & MyFolderName) And vbDirectory) = vbDirectory Then
                    AllFolders.Add (Key(i) & MyFolderName & "\"), ""
                End If
            End If
            MyFolderName = Dir
        Loop
        i = i + 1
    Loop
     
    'List all files
    For Each Key In AllFolders.Keys
        MyFileName = Dir(Key & "*.*")
        'MyFileName = Dir(Key & "*.PDF")    'only PDF files
        Do While MyFileName <> ""
            DateStamp = FileDateTime(Key & MyFileName)
            AllFiles.Add (Key & MyFileName & ":" & DateStamp), DateStamp
            MyFileName = Dir
        Loop
    Next
     
    '************************
    'List all files in Files sheet
     
    For Each MySheet In ThisWorkbook.Worksheets
        If MySheet.Name = "Files" Then
            Sheets("Files").Cells.Delete
            F = True
            Exit For
        Else
            F = False
        End If
    Next
    If Not F Then Sheets.Add.Name = Boots_Main_V_alpha.get_username & "_location_Files"
 
    'Sheets("Files").[A1].Resize(AllFolders.Count, 1) = WorksheetFunction.Transpose(AllFolders.keys)
    Sheets(Boots_Main_V_alpha.get_username & "_location_Files").[A1].Resize(AllFiles.count, 1) = WorksheetFunction.Transpose(AllFiles.Keys)
    Set AllFolders = Nothing
    Set AllFiles = Nothing
End Sub











































