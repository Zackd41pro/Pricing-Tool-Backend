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
                                                        'made for :matrix_v1
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
    MsgBox ("matrix_V2 STATUS:" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "------------------------------------------------------------" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "matrix_dimensions_Vx: IN Development")
End Sub

Public Function matrix_dimensions_Alpha(ByVal matrix_ As Variant, Dont_show_instructions As Boolean) As String
    'MsgBox ("This Function will be designed to fetch the ubound and lbound of a matrix to make it easyier to determine its size quickly")
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ubound-function
    'Stop
    'Exit Function
    Dim i As Long
    Dim upper As Variant
    Dim lower As Variant
    Dim store As Long
    Dim sz As Variant
    Dim matrix_dims As Long
        matrix_dims = 60 'max matrix dims as of 9/9/20
    
    MsgBox ("DOES NOT DEAL WITH '-' matrix VALUES CORRECTLY")
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
                End If
                If (lower < 0) Then
                    Stop
                    lower = Abs(lower)
                    upper = upper + lower + 1
                    lower = 0
                    Stop
                End If
            'post size
                If ((upper - lower) <> 0) Then
                    sz = upper - lower + 1
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
    Next i
    On Error GoTo 0
    Exit Function
    Stop
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
ReDim arr(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0)
x = matrix_V2.matrix_dimensions_Alpha(arr, True)
Stop
End Sub
