Attribute VB_Name = "Boots_Main_V_alpha"
Public Sub First_time_Run_only()
    'adds needed refs
        AddReference_part1_vba_app_extensibility_5_3
        AddReference_part2_vbscript
    'setup specific locations
        Call alpha_MkDir("Pricetool-Alpha-omega", "C:\")
        Call alpha_MkDir("version-0", "C:\Pricetool-Alpha-omega\")
        Call alpha_MkDir("Users", "C:\Pricetool-Alpha-omega\version-0\")
    'calls in code from specified location
        
    'setup thisworkbook runtime
End Sub
















'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'from alpha_make_dir

'https://stackoverflow.com/questions/43658276/create-folder-path-if-does-not-exist-saving-issue


'-----------------------------------------------------------------------------------------------------------
'from add ref page
'-----------------------------------------------------------------------------------------------------------

Private Sub AddReference_part2_vbscript()
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

Private Sub AddReference_part1_vba_app_extensibility_5_3()
     'Macro purpose:  To add a reference to the project using the GUID for the
     'reference library

    Dim strGUID As String, theRef As Variant, i As Long

     'Update the GUID you need below.
    strGUID = "{0002E157-0000-0000-C000-000000000046}"      'https://social.msdn.microsoft.com/Forums/en-US/57813453-9a21-4080-9d4a-e548e715d7ca/add-visual-basic-extensibility-library-through-code?forum=isvvba

     'Set to continue in case of error
    On Error Resume Next

     'Remove any missing references
    For i = ThisWorkbook.VBProject.References.count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.isbroken = True Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i

     'Clear any errors so that error trapping for GUID additions can be evaluated
    Err.Clear

     'Add the reference
    ThisWorkbook.VBProject.References.AddFromGuid _
    GUID:=strGUID, Major:=1, Minor:=0

     'If an error was encountered, inform the user
    Select Case Err.Number
    Case Is = 32813
         'Reference already in use.  No action necessary
    Case Is = vbNullString
         'Reference added without issue
    Case Else
         'An unknown error was encountered, so alert the user
        MsgBox "A problem was encountered trying to" & vbNewLine _
        & "add or remove a reference in this file" & vbNewLine & "Please check the " _
        & "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
    End Select
    On Error GoTo 0
End Sub


'requires reference to Microsoft Scripting Runtime
Function alpha_MkDir(strDir As String, strPath As String)

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


Sub test_b()
    Call alpha_MkDir("Pricetool-Alpha-omega", "C:\")
    Call alpha_MkDir("version-0", "C:\Pricetool-Alpha-omega\")
    Call alpha_MkDir("Users", "C:\Pricetool-Alpha-omega\version-0\")
End Sub

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

'https://stackoverflow.com/questions/55345116/excel-vba-determine-if-a-module-is-included-in-a-project

'http://www.vbaexpress.com/kb/getarticle.php?kb_id=250

Public Sub AddCode(newMacro As String, RandUniqStr As String, _
    Optional wb As Workbook)
    Dim VBC, modCode As String
    If wb Is Nothing Then Set wb = ThisWorkbook
    For Each VBC In wb.VBProject.VBComponents
        If VBC.CodeModule.CountOfLines > 0 Then
            modCode = VBC.CodeModule.Lines(1, VBC.CodeModule.CountOfLines)
            If modCode Like "*" & RandUniqStr & "*" And Not modCode Like "*" & newMacro & "*" Then
                VBC.CodeModule.InsertLines VBC.CodeModule.CountOfLines + 1, newMacro
                Exit Sub
            End If
        End If
    Next VBC
End Sub
 
Public Sub delCode(MacroNm As String, RandUniqStr As String, _
    Optional wb As Workbook)
    Dim VBC, i As Integer, procName As String, VBCM, j As Integer
    If wb Is Nothing Then Set wb = ThisWorkbook
    For Each VBC In wb.VBProject.VBComponents
        Set VBCM = VBC.CodeModule
        If VBCM.CountOfLines > 0 Then
            If VBCM.Lines(1, VBCM.CountOfLines) Like "*" & RandUniqStr & "*" Then
                i = VBCM.CountOfDeclarationLines + 1
                Do Until i >= VBCM.CountOfLines
                    procName = VBCM.ProcOfLine(i, 0)
                    If UCase(procName) = UCase(MacroNm) Then
                        j = VBCM.ProcCountLines(procName, 0)
                        VBCM.DeleteLines i, j
                        Exit Sub
                    End If
                    i = i + VBCM.ProcCountLines(procName, 0)
                Loop
            End If
        End If
    Next VBC
End Sub
 
Sub TestingIt()
    Dim prmtrs As String, toAdd As String
    prmtrs = "Key1:=Range(""C1""), Order1:=xlAscending, Header:=xlNo"
    toAdd = "Sub CreatedMacro()" & vbCrLf & " Cells.Sort " & prmtrs & vbCrLf & "End Sub"
    delCode "CreatedMacro", "a1b2c3d4e5f6g7h8i9", ThisWorkbook
    AddCode toAdd, "a1b2c3d4e5f6g7h8i9", ThisWorkbook
End Sub

Sub CreatedMacro()
 Cells.Sort Key1:=Range("C1"), Order1:=xlAscending, Header:=xlNo
End Sub




'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

'https://www.exceltip.com/modules-class-modules-in-vba/add-content-to-a-module-from-a-file-using-vba-in-microsoft-excel.html


Sub ImportModuleCode(ByVal wb As Workbook, ByVal ModuleName As String, ByVal ImportFromFile As String)
' imports code to ModuleName in wb from a textfile named ImportFromFile
Dim VBCM As CodeModule
    If Dir(ImportFromFile) = "" Then Exit Sub
    On Error Resume Next
    Set VBCM = wb.VBProject.VBComponents(ModuleName).CodeModule
    If Not VBCM Is Nothing Then
        VBCM.AddFromFile ImportFromFile
        Set VBCM = Nothing
    End If
    On Error GoTo 0
End Sub


Sub test()
    Call Boots_Main_V_alpha.ImportModuleCode(ActiveWorkbook, "thisworkbook", "C:\tp\test.txt")
End Sub



