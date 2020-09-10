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


Sub STATUS()
    MsgBox ("matrix_V2 STATUS:" & Chr(10) & _
    "------------------------------------------------------------" & Chr(10) & _
    "matrix_dimensions_Vx: Stable")
End Sub

Public Function matrix_dimensions_Alpha(ByVal matrix_ As Variant, Dont_show_instructions As Boolean) As String
'currently functional as of (9/10/2020) checked by: (Zachary Daughety)
    'Created By (Zachary Daugherty)(9/10/2020)
    'Purpose Case & notes:
        If (Dont_show_instructions = False) Then
                Call MsgBox("function is built to return the position range of space a array is using as well as the distance between the start and the finish position of that specific range element." & Chr(10) & Chr(10) & _
                "example:1" & Chr(10) & _
                "arr(1 to 4)as string will return: '(<1><4><4>),(<empty><empty><empty>),' as a string" & Chr(10) & _
                Chr(10) & "meaning the first dimension reports 4 positions starting at 1 and ending at 4 the second dimension reports empty as there is no defined space in this dimension", , "Showning instructions for matrix_v2.matrix_dimensions:1-4")
                
                Call MsgBox("example:2" & Chr(10) & _
                "arr(1,0,-5 to 10)as string will return:'(<0><1><2>),(<0><0><0>),(<-5><10><16>),(<empty><empty><empty>),' as string" & Chr(10) & _
                Chr(10) & "meaning the first dimension reports 2 positions starting at 0 and ending at 1. " & _
                "the second dimension reports 1 position starting and ending at 0. " & _
                "the third position reports 16 positions starting at -5 and ending at 10 " & _
                "this is because 0 is counted as a position", , "Showning instructions for matrix_v2.matrix_dimensions:2-4")
                
                Call MsgBox("example:3" & Chr(10) & _
                "arr() as string will return:'(<empty><empty><empty>),' as String" & Chr(10) & Chr(10) & _
                "this is because there are no dimensions assigned to this matrix", , "Showning instructions for matrix_v2.matrix_dimensions:3-4")
                
                Call MsgBox("example:4" & Chr(10) & _
                "I = 2 as long will return: '(<empty><empty><empty>),' as string " & Chr(10) & Chr(10) & _
                "meaning since it is not an matrix it does not have any other elements as of its position", , "Showning instructions for matrix_v2.matrix_dimensions:4-4")
            Stop
            Exit Function
        End If
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
    'code start
        'define variables
            Dim i As Long
            Dim upper As Variant
            Dim lower As Variant
            Dim store As Long
            Dim sz As Variant
            Dim matrix_dims As Long
            Dim neg_condition As Boolean
        'set variables
            matrix_dims = 60 'max matrix dims as of 9/9/20
        'run
            For i = 1 To matrix_dims
            'get size
                'get bounds
                    upper = PVT_MATRIX_BOUND(up_, matrix_, i)
                    lower = PVT_MATRIX_BOUND(down_, matrix_, i)
                    'check for fails
                        If (upper = "Failed to post" Or lower = "Failed to post") Then
                            If (upper = "Failed to post") Then
                                upper = "empty"
                            End If
                            If (lower = "Failed to post") Then
                                lower = "empty"
                            End If
                            'move to post
                                sz = "empty"
                                GoTo matrix_dimensions_Alpha_post
                        End If
                'check for negative
                    If (upper < 0) Then
                        store = Abs(upper)
                        upper = Abs(lower)
                        lower = store
                        store = 0
                        neg_condition = True
                    End If
                    If (lower < 0) Then
                        lower = Abs(lower)
                        upper = upper + lower + 1
                        lower = 0
                        neg_condition = True
                    End If
                'post size
                    If ((upper - lower) <> 0) Then
                    Stop
                        'check for neg_condition which changes how the report is as the 0 position is already added in
                        If (neg_condition = True) Then
                            sz = upper
                        Else
                            sz = upper - lower + 1
                        End If
                    Else
                        sz = 0
                    End If
            'get bounds for post
                upper = PVT_MATRIX_BOUND(up_, matrix_, i)
                lower = PVT_MATRIX_BOUND(down_, matrix_, i)
            'add entry
matrix_dimensions_Alpha_post:
                If (i = 1) Then
                'first entry
                    matrix_dimensions_Alpha = "(<" & lower & "><" & upper & "><" & sz & ">),"
                Else
                    If (i < 60) Then
                        matrix_dimensions_Alpha = matrix_dimensions_Alpha + "(<" & lower & "><" & upper & "><" & sz & ">),"
                    Else
                        Stop
                        matrix_dimensions_Alpha = matrix_dimensions_Alpha + "(<" & lower & "><" & upper & "><" & sz & ">)"
                    End If
                End If
            'check for emptys to exit
                If ((upper = "empty") Or (lower = "empty") Or (sz = "empty")) Then
                    Exit Function
                End If
            'reset neg_condition
                neg_condition = False
        Next i
    'code end
        On Error GoTo 0
        Exit Function
    'error handle
        'na
    'end error handle
End Function

Private Function PVT_MATRIX_BOUND(ByVal up_or_down As bound_choice, ByVal matrix_ As Variant, ByVal index As Long) As Variant
    If (up_or_down = up_) Then
        On Error GoTo pvt_matrix_bound_as_fail
        PVT_MATRIX_BOUND = UBound(matrix_, index)
    End If
    If (up_or_down = down_) Then
        On Error GoTo pvt_matrix_bound_as_fail
        PVT_MATRIX_BOUND = LBound(matrix_, index)
    End If
    Exit Function
pvt_matrix_bound_as_fail:
    PVT_MATRIX_BOUND = "Failed to post"
    On Error GoTo 0
End Function

Sub test()
Dim arr() As String
Dim x As String
Dim i As Long
i = 2
ReDim arr(1)
x = matrix_V2.matrix_dimensions_Alpha(arr(), False)
Stop
End Sub
