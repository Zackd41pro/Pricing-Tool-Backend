Attribute VB_Name = "Boots_Main_V_alpha"

Public Sub First_time_Run()
    'adds needed refs
    'calls in code from specified location
    'setup thisworkbook runtime
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

