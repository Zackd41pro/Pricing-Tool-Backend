Attribute VB_Name = "alpha_make_dir"
'https://stackoverflow.com/questions/43658276/create-folder-path-if-does-not-exist-saving-issue


'-----------------------------------------------------------------------------------------------------------
'from add ref page
'-----------------------------------------------------------------------------------------------------------

Sub AddReference()
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    Set vbProj = ActiveWorkbook.VBProject

    '~~> Check if "Microsoft VBScript Regular Expressions 5.5" is already added
    For Each chkRef In vbProj.References
        If chkRef.Name = "VBScript_RegExp_55" Then
            BoolExists = True
            GoTo CleanUp
        End If
    Next

    vbProj.References.AddFromFile "C:\Windows\SysWOW64\scrrun.dll"

CleanUp:
    If BoolExists = True Then
        MsgBox "Reference already exists"
    Else
        MsgBox "Reference Added Successfully"
    End If

    Set vbProj = Nothing
    Set VBAEditor = Nothing
End Sub

'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------

'requires reference to Microsoft Scripting Runtime
Function MkDir(strDir As String, strPath As String)

Dim fso As New FileSystemObject
Dim path As String

'examples for what are the input arguments
'strDir = "Folder"
'strPath = "C:\"

path = strPath & strDir

If Not fso.FolderExists(path) Then

' doesn't exist, so create the folder
          fso.CreateFolder path

End If

End Function


Sub test()
    Call MkDir("test log", "C:\")
End Sub






























