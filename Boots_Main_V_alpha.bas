Attribute VB_Name = "Boots_Main_V_alpha"
Public Enum boots_pos
    'sheet list
        i_sheet_count = 3
        p_sheet_name_row = 3
            p_sheet_name_col = 2
                p_sheet_visible_status_row = boots_pos.p_sheet_name_row + 0
                    p_sheet_visible_status_col = boots_pos.p_sheet_name_col + 1
                        p_sheet_color_row = boots_pos.p_sheet_visible_status_row + 0
                            p_sheet_color_col = boots_pos.p_sheet_visible_status_col + 1
                    
End Enum

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'first time boot code

Public Sub First_time_Run_only()
'adds needed refs
    AddReference_part1_vba_app_extensibility_5_3
    AddReference_part2_vbscript
'setup specific locations
    Call Make_Dir("Pricetool-Alpha-omega", "C:\")
    Call Make_Dir("version-0", "C:\Pricetool-Alpha-omega\")
    Call Make_Dir("Users", "C:\Pricetool-Alpha-omega\version-0\")
        
'calls in code from specified location
    
'setup thisworkbook runtime
    
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

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'boots env

Public Function run_on_start()
    MsgBox ("boots_main_startup add green text & instructions")
    'define varaiables
        'address
            Dim wb As Workbook
            Dim sht As Worksheets
        'container
            Dim bool As Boolean
    'setup varaibles
        Set wb = ActiveWorkbook
    'does boots env exist
        bool = boots_main_v_alpha.sheet_exist(wb, "Boots")
        If (bool = False) Then
            Call boots_main_v_alpha.make_sheet(wb, "Boots", -1, True)
        End If
        'format boots
            Call boots_main_v_alpha.boots_format
        'cleanup
            bool = False
    'fetch sheet list
        
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'boots functions public

'https://stackoverflow.com/questions/43658276/create-folder-path-if-does-not-exist-saving-issue
'requires reference to Microsoft Scripting Runtime
Public Function Make_Dir(strDir As String, strPath As String)
    MsgBox ("boots.make_dir needs green text & instructions")
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

Public Function make_sheet(ByVal wb As Workbook, ByVal sheet_name As String, Optional visible As Long, Optional dont_show_instructions As Boolean) As Boolean
    MsgBox ("boots_main.make_sheet needs to have green text added & instructions")
    If (dont_show_instructions = False) Then
        'show instructions
            Stop
    End If
    'define variables
        Dim home As Worksheet
        Dim sht As Worksheet
    'setup variables
        'check visible
            If ((visible <> -1) And (visible <> 2)) Then
                If (visible <> 0) Then
                    Call MsgBox("legal calls for 'make_sheet' are 0, -1, or 2")
                End If
                visible = 0
            End If
        'setup other variables
            Set home = ActiveSheet
    'run
        'make screen updating off
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            wb.Sheets.Add    'add new sheet
            Set sht = ActiveSheet  'set cursor to active
            sht.Name = sheet_name   'rename active sheet
            sht.visible = visible    'make dev hidden code lvl permission
            home.Activate 'return to home position
        'make screen updating on
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
End Function

Public Function sheet_exist(ByVal wb As Workbook, ByVal sheet As String) As Boolean
    MsgBox ("boots_main.sheet_exist need green text & instructions")
    'define variables
        Dim i As Long
        Dim s As String
        sheet = UCase(sheet)
    For i = 1 To wb.Sheets.count
        s = UCase(wb.Sheets(i).Name)
        If (s = sheet) Then
            sheet_exist = True
            Exit Function
        End If
    Next i
End Function

Public Function get_sheet_list() As Boolean
    MsgBox ("boots_main.get_sheet_list need green text & instructions")
    'define variables
        'address
            Dim wb As Workbook
            Dim sht As Worksheet
            Dim home As Worksheet
        'containers
            Dim arr() As String
            Dim i As Long
            Dim i_2 As Long
            Dim s As String
    'setup variables
        Set wb = ActiveWorkbook
        Set home = ActiveSheet
        'set boots page
            MsgBox ("boots_main.get_sheet_list need to add error report using boots_report saying something about how sheet boot cant be found")
            If (boots_main_v_alpha.sheet_exist(wb, "boots") = True) Then
                Set sht = wb.Sheets("boots")
            Else
                MsgBox ("boots_main.get_sheet_list error for finding boots")
                Stop
            End If
    'get size of arr
        ReDim arr(0 To wb.Sheets.count, 0 To boots_pos.i_sheet_count - 1)
        arr(0, 0) = "name"
        arr(0, 1) = "visible status"
        arr(0, 2) = "color status"
    'load arr
        Stop
        For i = 1 To wb.Sheets.count
            s = wb.Sheets(i).Name
            arr(i, 0) = s
            s = wb.Sheets(i).visible
            arr(i, 1) = s
            s = wb.Sheets(i).Tab.ColorIndex
            arr(i, 2) = s
        Next i
        'cleanup
            i = -1
            s = "empty"
    'get array fixed size
        i_2 = UBound(arr(), 1)
    'post values
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Stop
        For i = 0 To i_2
            'if not first row
                If (i <> 0) Then
                    'post name
                        sht.Cells(boots_pos.p_sheet_name_row + i, boots_pos.p_sheet_name_col).value = arr(i, 0)
                    'post visible
                        sht.Cells(boots_pos.p_sheet_visible_status_row + i, boots_pos.p_sheet_visible_status_col).value = arr(i, 1)
                    'post color
                        sht.Cells(boots_pos.p_sheet_color_row + i, boots_pos.p_sheet_color_col).value = arr(i, 2)
                End If
        Next i
        'cleanup
            i = -1
            i_2 = -1
            s = "empty"
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            get_sheet_list = True
            home.Activate
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'boots functions private

Private Sub boots_format()
    MsgBox ("<private>boots_main.boots_format needs green text & instructions")
    'define variables
        'address
            Dim wb As Workbook
            Dim sht As Worksheet
            Dim home As Worksheet
        'comtainers
            Dim i As Long
            Dim s As String
            Dim bool As Boolean
    'setup variables
        Set wb = ActiveWorkbook
        Set home = ActiveSheet
        Set sht = wb.Sheets("boots")
    'format
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sht.visible = -1
        'sheet index
            'name
                sht.Cells(boots_pos.p_sheet_name_row, boots_pos.p_sheet_name_col).value = "Sheet name"
            'visible
                sht.Cells(boots_pos.p_sheet_visible_status_row, boots_pos.p_sheet_visible_status_col).value = "Visible status"
            'color
                sht.Cells(boots_pos.p_sheet_color_row, boots_pos.p_sheet_color_col).value = "Color Status"
        'full sheet
            sht.Activate
            Cells.Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = -0.499984740745262
                    .PatternTintAndShade = 0
                End With
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
                Cells.EntireColumn.AutoFit
                Cells.Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Range("A1").Select
        'cleanup
        sht.visible = 2
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
End Sub




'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'alpha code


'https://stackoverflow.com/questions/55345116/excel-vba-determine-if-a-module-is-included-in-a-project

'http://www.vbaexpress.com/kb/getarticle.php?kb_id=250

Private Sub AddCode(newMacro As String, RandUniqStr As String, _
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
 
Private Sub delCode(MacroNm As String, RandUniqStr As String, _
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
 
Private Sub TestingIt()
    Dim prmtrs As String, toAdd As String
    prmtrs = "Key1:=Range(""C1""), Order1:=xlAscending, Header:=xlNo"
    toAdd = "Sub CreatedMacro()" & vbCrLf & " Cells.Sort " & prmtrs & vbCrLf & "End Sub"
    delCode "CreatedMacro", "a1b2c3d4e5f6g7h8i9", ThisWorkbook
    AddCode toAdd, "a1b2c3d4e5f6g7h8i9", ThisWorkbook
End Sub

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

'https://www.exceltip.com/modules-class-modules-in-vba/add-content-to-a-module-from-a-file-using-vba-in-microsoft-excel.html


Private Sub ImportModuleCode(ByVal wb As Workbook, ByVal ModuleName As String, ByVal ImportFromFile As String)
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


Private Sub test()
    Call boots_main_v_alpha.ImportModuleCode(ActiveWorkbook, "thisworkbook", "C:\tp\test.txt")
End Sub



