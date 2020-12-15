Attribute VB_Name = "matrix_V2"
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
                                                        'made for :matrix_v2
                                                                 ':NA
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Purpose Case
                                            'this module is built to deal with matrix quick functions
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Enum bound_choice
    up_
    down_
End Enum

Sub status()
    Call MsgBox("matrix_V2 STATUS:" & Chr(10) & _
    "------------------------------------------------------------" & Chr(10) & _
    "Public functions: " & Chr(10) & _
    "matrix_dimensions_v1: Stable" & Chr(10) & Chr(10) & _
    "Private functions:" & Chr(10) & _
    "PVT_MATRIX_BOUND: Stable", , "Matrix_vX showing Status")
End Sub

Public Function matrix_dimensions_v1(ByVal matrix_ As Variant, Optional more_instructions As String) As Variant
'currently functional as of (9/10/2020) checked by: (Zachary Daughety)
    'Created By (Zachary Daugherty)(9/10/2020)
    'Purpose Case & notes:
'        If (dont_show_instructions = False) Then
'                Call MsgBox("showing instructions for Matrix_vX.matrix_dimensions_v1, function is built to return the position range of space a array is using as well as the distance between the start and the finish position of that specific range element." & Chr(10) & Chr(10) & _
'                "example:1" & Chr(10) & _
'                "arr(1 to 4)as string will return: '(<1><4><4>),(<empty><empty><empty>),' as a string" & Chr(10) & _
'                Chr(10) & "meaning the first dimension reports 4 positions starting at 1 and ending at 4 the second dimension reports empty as there is no defined space in this dimension", , "Showning instructions for matrix_v2.matrix_dimensions_v1:1-4")
'
'                Call MsgBox("example:2" & Chr(10) & _
'                "arr(1,0,-5 to 10)as string will return:'(<0><1><2>),(<0><0><0>),(<-5><10><16>),(<empty><empty><empty>),' as string" & Chr(10) & _
'                Chr(10) & "meaning the first dimension reports 2 positions starting at 0 and ending at 1. " & _
'                "the second dimension reports 1 position starting and ending at 0. " & _
'                "the third position reports 16 positions starting at -5 and ending at 10 " & _
'                "this is because 0 is counted as a position", , "Showning instructions for matrix_v2.matrix_dimensions_v1:2-4")
'
'                Call MsgBox("example:3" & Chr(10) & _
'                "arr() as string will return:'(<empty><empty><empty>),' as String" & Chr(10) & Chr(10) & _
'                "this is because there are no dimensions assigned to this matrix", , "Showning instructions for matrix_v2.matrix_dimensions_v1:3-4")
'
'                Call MsgBox("example:4" & Chr(10) & _
'                "I = 2 as long will return: '(<empty><empty><empty>),' as string " & Chr(10) & Chr(10) & _
'                "meaning since it is not an matrix it does not have any other elements as of its position", , "Showning instructions for matrix_v2.matrix_dimensions_v1:4-4")
'            Stop
'            Exit Function
'        End If
    'Library Refrences required
        'workbook.object
    'Modules Required
        'na
    'Inputs
        'Internal:
            'PVT_MATRIX_BOUND
        'required:
            'some matrix
        'optional:
            'na
    'returned outputs
        'string listing
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            matrix_dimensions_v1 = "matrix_dimensions_v1 - Public - Need Log report 10/28/20 - and help file"
            Exit Function
        End If
    'code start
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Matrix_v2.matrix_dimensions_v1 Starting...")
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
        'define variables
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "setup of variables...")
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            Dim i As Long                   'container for long
            Dim upper As Variant            'container storing the upper bound of the array
            Dim lower As Variant            'container storing the lower bound of the array
            Dim store As Long               'container for temp storage
            Dim size As Variant             'container for storing size
            Dim const_matrix_dims As Long   'storing the the max value of dims in a array for excel
            Dim neg_condition As Boolean    'container for marking check if there are neg dimensions in the array
            'set variables
                const_matrix_dims = 60 'max matrix dims as of 9/9/20
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        'run
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Finding Size of full dims available...")
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
            For i = 1 To const_matrix_dims
            'get size
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Get dim size info for dimension #: '" & i & "' ...")
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                'get bounds
                    If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "fetching upper bounds...")
                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            upper = PVT_MATRIX_BOUND(up_, matrix_, i, more_instructions)  'run private function pvt_matrix_bound to get size and pos
                    If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "fetching lower bounds...")
                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                            lower = PVT_MATRIX_BOUND(down_, matrix_, i, more_instructions) 'run private function pvt_matrix_bound to get size and pos
                    'check for fails
                        If (upper = "Failed to post" Or lower = "Failed to post") Then  'if pvt_matrix_bound returns 'failed to post' then
                            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Clearing out error codes...")
                                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                                    If (upper = "Failed to post") Then      'check and if so post as empty as there is no dim in this index
                                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Clearing out error codes in upper bound...")
                                        upper = "empty"
                                    End If
                                    If (lower = "Failed to post") Then      'check and if so post as empty as there is no dim in this index
                                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Clearing out error codes in lower bound...")
                                        lower = "empty"
                                    End If
                                'move to post                           'since one or both of the entrys failed to post the size is empty
                                    size = "empty"
                                    If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                    GoTo matrix_dimensions_v1_Alpha_post   'jumps over formating as there is no number information to format
                        End If
                'check for negative
                    If (upper < 0) Then     'offsets the position of the bounds to positive as we are only finding the distance between both points
                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Upper is negative running math.abs...")
                        store = Abs(upper)
                        upper = Abs(lower)
                        lower = store
                        store = 0
                        neg_condition = True
                    End If
                    If (lower < 0) Then     'offsets the position of the bounds to positive as we are only finding the distance between both points
                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "lower is negative running math.abs...")
                        lower = Abs(lower)
                        upper = upper + lower + 1
                        lower = 0
                        neg_condition = True
                    End If
                'post delta start to finsih
                    If ((upper - lower) <> 0) Then  'checks for position tracking including zero or not
                        'check for neg_condition which changes how the report is as the 0 position is already added in
                            
                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Calculating Delta...")
                            
                        If (neg_condition = True) Then
                            size = upper
                        Else
                            size = upper - lower + 1
                        End If
                    Else
                        size = 0
                    End If
            'get bounds for post as now we are concerned with the positions where before we were just taking the delta position.
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Get Positions...")
                    If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
                        upper = PVT_MATRIX_BOUND(up_, matrix_, i, "d_report")
                        lower = PVT_MATRIX_BOUND(down_, matrix_, i, "d_report")
                    If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
            'add entry
matrix_dimensions_v1_Alpha_post:
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Load dim: '" & i & "' with data...")
                If (i = 1) Then
                'first entry
                    matrix_dimensions_v1_Alpha = "(<" & lower & "><" & upper & "><" & size & ">)"
                Else
                'any post after the first
                    If (i < 60) Then
                        matrix_dimensions_v1 = matrix_dimensions_v1_Alpha + ",(<" & lower & "><" & upper & "><" & size & ">)"
                    End If
                End If
            'check for emptys to exit
                If ((upper = "empty") Or (lower = "empty") Or (size = "empty")) Then
                    If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Get dim size info for dimension #: '" & i & "' complete...")
                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Finding Size of full dims available complete...")
                                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                    If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Matrix_v2.matrix_dimensions_v1 Finished...")
                                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                                Exit Function
                End If
            'reset neg_condition
                neg_condition = False   'resets for next loop
            'reset log line
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Get dim size info for dimension #: '" & i & "' complete...")
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Next i
    'code end
        On Error GoTo 0
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Get dim size info for dimension #: '" & i & "' complete...")
            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Finding Size of full dims available complete...")
                    If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Matrix_v2.matrix_dimensions_v1 Finished...")
                            If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
                                    Exit Function
    'error handle
        'na
    'end error handle
End Function

Private Function PVT_MATRIX_BOUND(ByVal up_or_down As bound_choice, ByVal matrix_ As Variant, ByVal index As Long, Optional more_instructions As String) As Variant
'currently functional as of (9/9/2020) checked by: (Zachary Daugherty)
    'Created By (Zachary Daugherty)(9/11/2020)
    'Purpose Case & notes:
'        If (dont_show_instructions = False) Then
'            Call MsgBox("Showing Instructions for Matrix_VX:" & Chr(10) & Chr(10) & _
'            "this function is a private function that returns from the provided array and index the specified information" & Chr(10) & _
'            "possible answers are any number including zero as zero is a dimension or 'failed to post' meaning empty dimension" & Chr(10) & _
'            "if you select 'up' you will be given the upper bound of the array in the specified index dimension" & Chr(10) & _
'            "if you select 'down' you will be given the upper bound of the array in the specified index dimension" & Chr(10) & _
'            "if a dimension does not report anything the program will return back 'failed to post'", , "Showing instructions for Matrix_V2.PVT_MATRIX_BOUND")
'            Stop
'            Exit Function
'        End If
    'Library Refrences required
        'workbook.object
    'Modules Required
        'na
    'Inputs
        'Internal:
            'na
        'required:
            'what you want returned
            'array
            'dimension index
        'optional:
            'na
    'returned outputs
        'dimension position as long or 'Failed to post'
    'check for log reporting
        If (more_instructions = "Log_Report") Then
            PVT_MATRIX_BOUND = "PVT_MATRIX_BOUND - Private - Need Log report 10/28/20 - need help file"
            Exit Function
        End If
    'code start
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Matrix_v2.PVT_MATRIX_BOUND Start...")
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_S)
        'up
            If (up_or_down = up_) Then
                On Error GoTo pvt_matrix_bound_as_fail
                PVT_MATRIX_BOUND = UBound(matrix_, index)
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Returning upper bound...")
            End If
        'down
            If (up_or_down = down_) Then
                On Error GoTo pvt_matrix_bound_as_fail
                PVT_MATRIX_BOUND = LBound(matrix_, index)
                If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Returning lower bound...")
            End If
    'code end
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Matrix_v2.PVT_MATRIX_BOUND Finish...")
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Exit Function
    'error handle
pvt_matrix_bound_as_fail:
        PVT_MATRIX_BOUND = "Failed to post"
        On Error GoTo 0
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Returning ERROR FAILED TO POST...")
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(text, "Matrix_v2.PVT_MATRIX_BOUND Finish...")
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        If more_instructions <> "d_report" Then Call Boots_Report_v_Alpha.Log_Push(Trigger_e)
        Exit Function
    'end error handle
End Function
