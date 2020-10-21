Attribute VB_Name = "Boots_Main_V_alpha"

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
                                                        'made for :boots_main
                                                                 ':Boots_report
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Purpose Case
                    'This Module is built to handle exporting of information
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Enum boots_pos
    'sheet list
        i_sheet_count = 3
        p_sheet_name_row = 3
            p_sheet_name_col = 2
                p_sheet_visible_status_row = boots_pos.p_sheet_name_row + 0
                    p_sheet_visible_status_col = boots_pos.p_sheet_name_col + 1
                        p_sheet_color_row = boots_pos.p_sheet_visible_status_row + 0
                            p_sheet_color_col = boots_pos.p_sheet_visible_status_col + 1
    'module tacker
        p_track_module_name_row = 7
            p_track_module_name_col = 6
                p_track_module_type_row = boots_pos.p_track_module_name_row + 0
                    p_track_module_type_col = boots_pos.p_track_module_name_col + 1
                        p_track_module_version_row = boots_pos.p_track_module_name_row + 0
                            p_track_module_version_col = boots_pos.p_track_module_type_col + 1
    'tracker
        'keystone
            p_tracker_keystone_row = 3
                p_tracker_keystone_col = 6
        'master pos def
            'rh (row header)
                p_rh_tracker_index_now_row = boots_pos.p_tracker_keystone_row + 1
                    p_rh_tracker_index_now_col = boots_pos.p_tracker_keystone_col + 0
                        p_rh_tracker_index_last_row = boots_pos.p_rh_tracker_index_now_row + 1
                            p_rh_tracker_index_last_col = boots_pos.p_rh_tracker_index_now_col + 0
            'ch (col header)
                p_ch_tracker_wb_row = boots_pos.p_tracker_keystone_row + 0
                    p_ch_tracker_wb_col = boots_pos.p_tracker_keystone_col + 1
                        p_ch_tracker_sht_row = boots_pos.p_ch_tracker_wb_row + 0
                            p_ch_tracker_sht_col = boots_pos.p_ch_tracker_wb_col + 1
                                p_ch_tracker_cell_row = boots_pos.p_ch_tracker_sht_row + 0
                                    p_ch_tracker_cell_col = boots_pos.p_ch_tracker_sht_col + 1
                                        p_ch_tracker_key_row = boots_pos.p_ch_tracker_cell_row + 0
                                            p_ch_tracker_key_col = boots_pos.p_ch_tracker_cell_col + 1
                                                p_ch_tracker_string_row = boots_pos.p_ch_tracker_key_row + 0
                                                    p_ch_tracker_string_col = boots_pos.p_ch_tracker_key_col + 1
            'workbook
                p_tracker_wb_now_row = boots_pos.p_rh_tracker_index_now_row + 0
                    p_tracker_wb_now_col = boots_pos.p_ch_tracker_wb_col + 0
                        p_tracker_wb_last_row = boots_pos.p_tracker_wb_now_row + 1
                            p_tracker_wb_last_col = boots_pos.p_tracker_wb_now_col + 0
            'sheet
                p_tracker_sheet_now_row = boots_pos.p_rh_tracker_index_now_row + 0
                    p_tracker_sheet_now_col = boots_pos.p_ch_tracker_sht_col + 0
                        p_tracker_sheet_last_row = boots_pos.p_tracker_sheet_now_row + 1
                            p_tracker_sheet_last_col = boots_pos.p_tracker_sheet_now_col + 0
            'cell
                p_tracker_cell_now_row = boots_pos.p_rh_tracker_index_now_row + 0
                    p_tracker_cell_now_col = boots_pos.p_ch_tracker_cell_col + 0
                        p_tracker_cell_last_row = boots_pos.p_tracker_cell_now_row + 1
                            p_tracker_cell_last_col = boots_pos.p_ch_tracker_cell_col + 0
            'key
                p_tracker_key_now_row = boots_pos.p_rh_tracker_index_now_row + 0
                    p_tracker_key_now_col = boots_pos.p_ch_tracker_key_col + 0
                        p_tracker_key_last_row = boots_pos.p_rh_tracker_index_last_row + 0
                            p_tracker_key_last_col = boots_pos.p_ch_tracker_key_col + 0
            'string
                p_tracker_string_now_row = boots_pos.p_rh_tracker_index_now_row + 0
                    p_tracker_string_now_col = boots_pos.p_ch_tracker_string_col + 0
                        p_tracker_string_last_row = boots_pos.p_rh_tracker_index_last_row + 0
                            p_tracker_string_last_col = boots_pos.p_ch_tracker_string_col + 0
End Enum

Enum get_project_files_choices
    na
    get_index
End Enum

Private Function global_get_project_files_not_tracked_filename() As String
    global_get_project_files_not_tracked_filename = "Na"
End Function


Private Function LOG_push_version(ByVal pos As Long) 'for version reporting
    Dim version As String
    Dim sht As Worksheet
    
    version = "Na"
    
    If (Boots_Main_V_alpha.sheet_exist(ActiveWorkbook, "Boots") = True) Then
        Set sht = ActiveWorkbook.Sheets("Boots")
    Else
        Exit Function
    End If
    sht.Cells(boots_pos.p_track_module_version_row + pos, boots_pos.p_track_module_version_col).value = version
End Function

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-
'first time boot code

Public Sub First_time_Run_only()
'adds needed refs
    'AddReference_part1_vba_app_extensibility_5_3
    'AddReference_part2_vbscript
'setup specific locations
    
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
    'initalize log
        Boots_Report_v_Alpha.Log_Initalize
    'does boots env exist
        bool = Boots_Main_V_alpha.sheet_exist(wb, "Boots")
        If (bool = False) Then
            Call Boots_Main_V_alpha.make_sheet(wb, "Boots", -1, True)
        End If
        'format boots
            Call Boots_Main_V_alpha.boots_format
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

Public Function sheet_exist(ByVal wb As Workbook, ByVal sheet As String, Optional more_instructions As String) As Boolean
    'function is designed to report back to the IDE if the specified sheet in the specified Enviorment exists and if it does report back as true
    'define variables
        Dim i As Long
        Dim s As String
        'convert to uppercase
            sheet = UCase(sheet)
    'iterate through all the sheets
        For i = 1 To wb.Sheets.count
            'store name in s
                s = UCase(wb.Sheets(i).Name)
            'if s and sheet are the same report as exist then exit else roll through
                If (s = sheet) Then
                    sheet_exist = True
                    Exit Function
                End If
        Next i
        'sheet not found so false will be returned
End Function

Public Function get_sheet_list() As Boolean
    'this function fetches the names of all the sheets in the workbook and posts them to the 'boots' page to make referencing them easy also it posts the color status
        'and the color of the tab.
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
            'check for boots page if not exist create it
                If (Boots_Main_V_alpha.sheet_exist(wb, "boots") = True) Then
                'exist
                    Set sht = wb.Sheets("boots")
                Else
                'dont exist
                    MsgBox ("boots_main.get_sheet_list error for finding boots")
                    Stop
                End If
    'get size of arr
        ReDim arr(0 To wb.Sheets.count, 0 To boots_pos.i_sheet_count - 1)
        'set namespace
            arr(0, 0) = "name"
            arr(0, 1) = "visible status"
            arr(0, 2) = "color status"
    'load arr
        For i = 1 To wb.Sheets.count
        'get sheet name
            s = wb.Sheets(i).Name
            'store sheet name
                arr(i, 0) = s
        'get visibility status
            s = wb.Sheets(i).visible
            'store vis status
                arr(i, 1) = s
        'get color index
            s = wb.Sheets(i).Tab.ColorIndex
            'store color
                arr(i, 2) = s
        Next i
        'cleanup
            i = -1
            s = "empty"
    'clear old values
        'hide updating
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
        'clear old values
            'find the size of the old list
                i = 1
                i_2 = 0
restart_get_sheet_clear_check:
                s = sht.Cells(boots_pos.p_sheet_name_row + i, boots_pos.p_sheet_name_col).value
                If (s <> "") Then
                'count and reset
                    i = i + 1
                    GoTo restart_get_sheet_clear_check
                End If
            'clear list
                i_2 = i
                For i = 1 To i_2
                    sht.Cells(boots_pos.p_sheet_name_row + i, boots_pos.p_sheet_name_col).value = ""
                    sht.Cells(boots_pos.p_sheet_visible_status_row + i, boots_pos.p_sheet_visible_status_col).value = ""
                    sht.Cells(boots_pos.p_sheet_color_row + i, boots_pos.p_sheet_color_col).value = ""
                Next i
        'cleanup
            i = -1
            i_2 = -1
            s = "empty"
    'get array fixed size
        i_2 = UBound(arr(), 1)
    'post values to boots
        'start loop to post
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
            get_sheet_list = True
End Function

Public Function get_project_files(Optional Optional_more_instructions As get_project_files_choices, Optional Optional_input As Variant) As Variant
    'this function fetches from the current workbook to find the workbook objects associated with the project as stores them on the boots page in the code name table. once this is done
        'get_project_files will also return the size of this table as its return variable.
        
        'if get index is selected the user must give some sort of numerical position in the list and will receice that module report
            'if the entry is not a number it will return the first position in the list
            'if the number is bigger than the index list it will return the last
            'once this is all determined a string of the asked for infomration will be returned in the return variable.
            
    'define varaibles
        Dim VBC                             'cursor selector to find and explore modules
        Dim type_ As vbext_ComponentType    'enumeration selection object
        'containers
            Dim i As Long
            Dim i_2 As Long
            Dim s As String
        'address positions
            Dim wb As Workbook
            Dim sht As Worksheet
    'set variables
        'set workbook obj to the active one
            Set wb = ActiveWorkbook
        'assign sht to boots if it exists
            If (Boots_Main_V_alpha.sheet_exist(ActiveWorkbook, "Boots") = True) Then
                Set sht = wb.Sheets("Boots")
            Else
                MsgBox ("get_project_files fatal error: unable to locate sheet boots")
            End If
    'clear the old list in boots only if optional instructions is set to 'na'
    If (Optional_more_instructions = get_project_files_choices.na) Then
        'find the size of the old list
            i = 1
            i_2 = 0
restart_get_project_file_clear_check:
            s = sht.Cells(boots_pos.p_track_module_name_row + i, boots_pos.p_track_module_name_col).value
            If (s <> "") Then
            'count and reset
                i = i + 1
                GoTo restart_get_project_file_clear_check
            End If
            'set get_project_files to the size of the table
                get_project_files = i
        'clear list
            i_2 = i
            For i = 1 To i_2
                sht.Cells(boots_pos.p_track_module_name_row + i, boots_pos.p_track_module_name_col).value = ""
                sht.Cells(boots_pos.p_track_module_type_row + i, boots_pos.p_track_module_type_col).value = ""
                sht.Cells(boots_pos.p_track_module_version_row + i, boots_pos.p_track_module_version_col).value = ""
            Next i
        'cleanup
            i = -1
            i_2 = -1
            s = "empty"
    End If
    'iterate throught all project files to get list
        i = 0
        On Error GoTo get_project_files_programatic_access_to_vb_model_failed
        For Each VBC In wb.VBProject.VBComponents
            'iterate row position
                i = i + 1
            'do only if option select is NA
            If (Optional_more_instructions = get_project_files_choices.na) Then
                'get type for this iteration
                    type_ = VBC.Type
                'find the type match and post
                    Select Case type_
                        Case vbext_ct_StdModule
                        'std module
                            sht.Cells(boots_pos.p_track_module_name_row + i, boots_pos.p_track_module_name_col).value = CStr(VBC.Name)
                            sht.Cells(boots_pos.p_track_module_type_row + i, boots_pos.p_track_module_type_col).value = CStr(VBC.Type)
                            Boots_Main_V_alpha.get_project_files_sub_LOG_push_version_module (i)
                        Case vbext_ct_Document
                        'sheets and insert page charts
                            sht.Cells(boots_pos.p_track_module_name_row + i, boots_pos.p_track_module_name_col).value = CStr(VBC.Name)
                            sht.Cells(boots_pos.p_track_module_type_row + i, boots_pos.p_track_module_type_col).value = CStr(VBC.Type)
                            sht.Cells(boots_pos.p_track_module_version_row + i, boots_pos.p_track_module_version_col).value = Boots_Main_V_alpha.global_get_project_files_not_tracked_filename
                        Case vbext_ct_MSForm
                        'this is a userform
                            sht.Cells(boots_pos.p_track_module_name_row + i, boots_pos.p_track_module_name_col).value = CStr(VBC.Name)
                            sht.Cells(boots_pos.p_track_module_type_row + i, boots_pos.p_track_module_type_col).value = CStr(VBC.Type)
                            sht.Cells(boots_pos.p_track_module_version_row + i, boots_pos.p_track_module_version_col).value = Boots_Main_V_alpha.global_get_project_files_not_tracked_filename
                        Case vbext_ct_ClassModule
                        'class module
                            sht.Cells(boots_pos.p_track_module_name_row + i, boots_pos.p_track_module_name_col).value = CStr(VBC.Name)
                            sht.Cells(boots_pos.p_track_module_type_row + i, boots_pos.p_track_module_type_col).value = CStr(VBC.Type)
                            sht.Cells(boots_pos.p_track_module_version_row + i, boots_pos.p_track_module_version_col).value = Boots_Main_V_alpha.global_get_project_files_not_tracked_filename
                        Case Else
                            MsgBox ("get_project_files non fatal error: is unable to identify the type of this object")
                            sht.Cells(boots_pos.p_track_module_name_row + i, boots_pos.p_track_module_name_col).value = CStr(VBC.Name)
                            sht.Cells(boots_pos.p_track_module_type_row + i, boots_pos.p_track_module_type_col).value = CStr(VBC.Type)
                            sht.Cells(boots_pos.p_track_module_version_row + i, boots_pos.p_track_module_version_col).value = Boots_Main_V_alpha.global_get_project_files_not_tracked_filename
                    End Select
            End If
            'if get_index Optional_more_instructions is 'get_index'
                If (Optional_more_instructions = get_project_files_choices.get_index) Then
                    'restart checkpoint
restart_get_project_files_get_index:
                    'check to see if optional input is a number
                        If (IsNumeric(Optional_input) = False) Then
                            'Optional_input is not a number report this and assign 1
                                MsgBox ("Error: get_project_files received a non numeric value for Optional_input when trying to find a get an index position will return the first position in the list")
                                Optional_input = 1
                                GoTo restart_get_project_files_get_index
                        End If
                    'check to see if the number given is bigger than the list
                        If (get_project_files <> Empty) Then
                            If (Optional_input > get_project_files) Then
                                Optional_input = get_project_files
                            End If
                        End If
                    'return pos to get_project_files for to return the requested information
                        If (i = Optional_input) Then
                            get_project_files = "Installed Project Object Report: Type: <" & sht.Cells(boots_pos.p_track_module_type_row + i, boots_pos.p_track_module_type_col).value & _
                            "> Name: <" & sht.Cells(boots_pos.p_track_module_name_row + i, boots_pos.p_track_module_name_col).value & "> Version: <" & sht.Cells(boots_pos.p_track_module_version_row + i, boots_pos.p_track_module_version_col).value & ">"
                            GoTo get_project_files_exit
                        End If
                End If
        Next VBC
    'cleanup
get_project_files_exit:
        'na
        Exit Function
    'error handle
        'get_project_files_programatic_access_to_vb_model_failed
get_project_files_programatic_access_to_vb_model_failed:
            'reset error caller
                'na
            'push variables to the log
                Call Boots_Report_v_Alpha.Push_Log(Error_, "")
                Call Boots_Report_v_Alpha.Push_Log(table_open, "")
                Call Boots_Report_v_Alpha.Push_Log(text, "SHOWING VARIABLES SNAPSHOT FROM... get_project_files")
                Call Boots_Report_v_Alpha.Push_Log(Variable, "VBC: '" & VBC & "' as variant/empty")
                Call Boots_Report_v_Alpha.Push_Log(Variable, "type_: '" & type_ & "' as vbext_ComponentType")
                Call Boots_Report_v_Alpha.Push_Log(Variable, "i: '" & i & "' as long")
                Call Boots_Report_v_Alpha.Push_Log(Variable, "i_2: '" & i_2 & "' as long")
                Call Boots_Report_v_Alpha.Push_Log(Variable, "s: '" & s & "' as string")
                Call Boots_Report_v_Alpha.Push_Log(Variable, "wb: '" & wb.path & "/" & wb.Name & "' as Workbook")
                Call Boots_Report_v_Alpha.Push_Log(Variable, "Sht: '" & sht.Name & "' as worksheet")
                Call Boots_Report_v_Alpha.Push_Log(Variable, "get_project_files: '" & get_project_files & "' as variant")
                Call Boots_Report_v_Alpha.Push_Log(Variable, "Optional_more_instructions: '" & Optional_more_instructions & "' as get_project_files_choices")
                Call Boots_Report_v_Alpha.Push_Log(table_close, "")
            'call the error handle
                Boots_Main_V_alpha.ERROR_programatic_access_to_vb_model_failed
End Function

Private Function get_project_files_sub_LOG_push_version_module(ByVal i As Long)
    Dim wb As Workbook
    Dim sht As Worksheet
    Dim VBC                             'cursor selector to find and explore modules
    Set wb = ActiveWorkbook 'set workbook obj to the active one
    Set VBC = wb.VBProject.VBComponents.Item(i)
    If (Boots_Main_V_alpha.sheet_exist(ActiveWorkbook, "Boots") = True) Then
        Set sht = wb.Sheets("Boots")
    Else
        MsgBox ("throw error saying boots page missing")
    End If
    On Error GoTo get_proj_files_sub_version_module_skip
        i_2 = Run(VBC.Name & Chr(46) & "LOG_push_version", i)
    If (1 <> 1) Then
get_proj_files_sub_version_module_skip:
        sht.Cells(boots_pos.p_track_module_version_row + i, boots_pos.p_track_module_version_col).value = Boots_Main_V_alpha.global_get_project_files_not_tracked_filename
    On Error GoTo 0
    End If
End Function

Public Function get_username(Optional dont_report As Boolean) As Variant
'currently NOT functional as of (8/7/2020) checked by: (Zachary Daugherty)
    'Created By (Zachary Daugherty)(8/7/2020)
    'Purpose Case & notes:
        'made to use functions to return the user name of the account in this session
    'Library Refrences required
        'workbook.object
    'Modules Required
        'Na
    'Inputs
        'Internal:
            'Na
        'required:
            'Na
        'optional:
            'Na
    'returned outputs
        'username as string
'code start
    get_username = (Environ$("Username"))
'code end
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
        'tracker
            'key
                sht.Cells(boots_pos.p_tracker_keystone_row, boots_pos.p_tracker_keystone_col).value = "Keystone"
            'headers
                sht.Cells(boots_pos.p_ch_tracker_wb_row, boots_pos.p_ch_tracker_wb_col).value = "Workbook"
                sht.Cells(boots_pos.p_ch_tracker_sht_row, boots_pos.p_ch_tracker_sht_col).value = "Sheet"
                sht.Cells(boots_pos.p_ch_tracker_cell_row, boots_pos.p_ch_tracker_cell_col).value = "Cells"
                sht.Cells(boots_pos.p_ch_tracker_key_row, boots_pos.p_ch_tracker_key_col).value = "Key"
                sht.Cells(boots_pos.p_ch_tracker_string_row, boots_pos.p_ch_tracker_string_col).value = "String"
            'now
                sht.Cells(boots_pos.p_rh_tracker_index_now_row, boots_pos.p_rh_tracker_index_now_col).value = "Now"
            'last
                sht.Cells(boots_pos.p_rh_tracker_index_last_row, boots_pos.p_rh_tracker_index_last_col).value = "Last"
        'module tracker
            sht.Cells(boots_pos.p_track_module_name_row, boots_pos.p_track_module_name_col).value = "Code Name"
            sht.Cells(boots_pos.p_track_module_type_row, boots_pos.p_track_module_type_col).value = "Type"
            sht.Cells(boots_pos.p_track_module_version_row, boots_pos.p_track_module_version_col).value = "Version"
        'full sheet
            sht.Activate
            Cells.Select
                With Selection.Font
                    .Color = -16711936
                    .TintAndShade = 0
                End With
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("A1").Select
            With ActiveWorkbook.Sheets("Boots").Tab
                .Color = 65280
                .TintAndShade = 0
            End With
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
    Call Boots_Main_V_alpha.ImportModuleCode(ActiveWorkbook, "thisworkbook", "C:\tp\test.txt")
End Sub


'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
                                                
                                                'Error Handler

'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Sub ERROR_programatic_access_to_vb_model_failed()
    Call Boots_Report_v_Alpha.Push_Log(Flag, "")
    Call Boots_Report_v_Alpha.Push_Log(text, "FATAL RUN-TIME ERROR '1004': Programmatic access to Visual basic project is not trusted")
    Call Boots_Report_v_Alpha.Push_Log(text, "To Fix this error in your excel environment click on 'File'; 'Options'")
    Call Boots_Report_v_Alpha.Push_Log(text, "then in the 'Options' window: click on 'Trust Center' then under 'Microsoft Excel Trust Center' select the 'trust center settings' button")
    Call Boots_Report_v_Alpha.Push_Log(text, "inside this menu click: 'Macro Settings' shown in the right bar.")
    Call Boots_Report_v_Alpha.Push_Log(text, "then inside this menu: under 'Developer macro settings' check the box displaying 'Trust access to the Vba Project Object model'")
    Call Boots_Report_v_Alpha.Push_Log(text, "then exit and restart the program")
    ActiveWorkbook.Save
    Call Boots_Report_v_Alpha.Push_Log(Display_now, "")
End Sub































