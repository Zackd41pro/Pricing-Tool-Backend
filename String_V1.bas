Attribute VB_Name = "String_V1"
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
                                                        'made for :String_V1
                                                                 ':NA
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'Purpose Case
                    'this module is built to deal with tasks regarding usage with strings
                    
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        'CODE START
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Enum Shift_option
    up_s
    down_s
    left_s
    right_s
End Enum

Enum Disassociate_by_Char_left_or_right
    Left_C = 0
    Right_C = 1
End Enum

Enum Get_Char
    New_line = 10                   'next line'
    'break
    carriage_return = 13            'carriage return'
    'break
    Space = 32                      '' '
    Exclamation_mark = 33           '!
    Double_quotes = 34              '"
    Pound = 35                      '#
    Dollar = 36                     '$
    Percent = 37                    '%
    Ampersand = 38                  '&
    single_quote = 39               ''
    Open_parenthesis = 40           '(
    Close_parenthesis = 41          ')
    multiply = 42                   '*
    plus = 43                       '+
    comma = 44                      ',
    hyphen = 45                     '-
    Period = 46                     '.
    forward_slash = 47              '/
    'break
    colon = 58                      ':
    semicolon = 59                  ';
    less_than = 60                  '<
    equals = 61                     '=
    greater_than = 62               '>
    question_mark = 53              '?
    At = 64                         '@
    'break
    open_bracket = 91               '[
    back_slash = 92                  '\
    closing_bracket = 93            ']
    Power = 94                      '^
    underscore = 95                 '_
    grave_accent = 96               '`
    'break
    opening_brace = 123             '{
    vertial_bar = 124               '|
    Closing_brace = 125             '}
    Equivalency = 126               '~
    Delete = 127                    '
    Euro = 128                      '€
    'break
    Bullet = 149                    '•
    'break
    Trademark = 153                 '™
    'break
    inverted_exclamation_mark = 161 '¡
    cent = 162                      '¢
    'break
    copyright = 169                 '©
    'break
    Left_double_angle_quote = 171   '«
    'break
    Registered_TM = 174             '®
    'break
    Degrees = 176                   '°
    plus_or_minus = 177             '±
    'break
    Pilcrow_sign = 182              '¶
    middle_dot = 183                '·
    'break
    Division_sign = 247             '÷
End Enum

Sub STATUS()
    MsgBox ("STRING_V1 STATUS:" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "------------------------------------------------------------" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "is_same_V1: Functional" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "Has string inside_v1: Functional" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "Disassociate_by_Char_V1: Functional" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "get_Special_Char_V1: Functional" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "------------------------------------------------------------" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "Shift_V1: Needs Updating" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "------------------------------------------------------------" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "IndentText_VA: IN ALPHA" & String_V1.get_Special_Char_V1(carriage_return, True) & _
    "")
End Sub

Public Function is_same_V1(ByVal string_1 As String, ByVal string_2 As String, Optional Dont_show_instructions As Boolean) As Boolean
'currently functional as of (8/20/2020) checked by: (zachary Daugherty)
    'Created By (zachary daugherty)(8/20/20)
    'Purpose Case & notes:
        'string compair
    'Library Refrences required
        'na
    'Modules Required
        'na
    'Inputs
        'Internal:
            'na
        'required:
            'string_1 first string to compair
            'string_2 second string to compair
        'optional:
            'na
    'returned outputs
        'true if same
        'false if not
    'code start
        'check for show instructions check
            If (Dont_show_instructions = False) Then
                MsgBox ("function is_same_V1 takes string_1 and compairs string_2 to see if they are identical if so the function returns 1 or true if not it returns 0 or false")
                Stop
                Exit Function
            End If
        'compair
            If (string_1 = string_2) Then
                is_same = True
            Else
                Exit Function
            End If
    'code finish
End Function

Public Function has_string_inside_V1(ByVal value As String, ByVal sequence As String, _
    Optional give_pos As Boolean, Optional show_instructions As Boolean) As Long
'currently NOT functional as of (8/21/20) checked by: (Zachary Daugherty)
    'Created By (zachary Daugherty)(8/20/20)
    'Purpose Case & notes:
        'function has_string checks to see if values are in the sequence reports position if give_pos is true.
            'if give_pos is true and the value is not inside will return 0
    'Library Refrences required
        'na
    'Modules Required
        'na
    'Inputs
        'Internal:
            'na
        'required:
            'value: what we are searching for
            'sequence:were you are looking through
        'optional:
            '(information not filled out)
    'returned outputs
        '(information not filled out)
    'code start
        'define variables
            Dim s As String 'string storage
            Dim i As Long    'int storage
            Dim len_val As Long
            Dim len_seq As Long
            Dim pos As Long  'count thru
        'check for instructions post
            If (show_instructions = True) Then
                MsgBox ("function has_string_inside_V1 checks to see if values are in the sequence reports position if give_pos is true. if give_pos is true and the value is not inside will return 0")
                Stop
                Exit Function
            End If
        'setup variables
            s = sequence
            'find length of segment
                len_val = Len(value)
                len_seq = Len(sequence)
        'check if the value is bigger than the Sequence
            If (len_val > len_seq) Then
                'cleanup
                    has_string_inside_V1 = False
                    s = "empty"
                    len_val = -1
                    len_seq = -1
                    i = -1
                    Exit Function
            End If
        'scan
            For i = 1 To len_seq
                s = Mid(sequence, i, len_val)
                If (value = s) Then
                    'solution found
                        If (give_pos = True) Then
                            has_string_inside_V1 = i
                            Exit Function
                        Else
                            has_string_inside_V1 = True
                        End If
                        Exit For
                End If
            Next i
        'cleanup
            value = "empty"
            sequence = "empty"
            s = "empty"
            i = -1
            len_val = i
            len_seq = i
            pos = i
            Exit Function
    'code finish
End Function

Public Function Disassociate_by_Char_V1(ByVal seperator As String, ByVal sequence As String, _
    ByVal Left_or_Right As Disassociate_by_Char_left_or_right, Optional Dont_show_instructions As Boolean) As String
'currently functional as of (8/21/20) checked by: (Zachary Daugherty)
    'Created By (Zachary Daugherty)(8/21/20)
    'Purpose Case & notes:
        'takes input of sequence and returns the left or the right side of the seperator specified
    'Library Refrences required
        'workbook.object
    'Modules Required
        'na
    'Inputs
        'Internal:
            'String_V1.has_string_inside_V1
        'required:
            'seperator: what is seperating your text
            'sequence: the full string you want to pull text from
            'Left_or_right: specifys what side of the seperator to return
            
            'E.G.
            'sequence: 'zack-daugherty'
            'seperator: '-'
            'Left_or_Right: 'Right'
            'will return: 'daugherty'
            
        'optional:
            'show_instructions: gives a walkthrough on how to use
    'returned outputs
        'disassociated string text after or before seperator
    'code start
            'define variables
                Dim Seperator_len As Long
                Dim seperator_pos As Variant
                Dim Sequence_len As Long
                Dim i As Long
            'check for instructions
                If (Dont_show_instructions = False) Then
                    MsgBox ("Showning instructions for Disassociate_by_char_v1")
                    MsgBox ("Purpose Case & notes: takes input of sequence and returns the left or the right side of the seperator specified")
                    MsgBox ("Discription of variables:" & Chr(13) & _
                    "Internal: " & Chr(13) & _
                        "String_V1.has_string_inside_V1" & Chr(13) & _
                        "required: " & Chr(13) & _
                        "seperator: what is seperating your text" & Chr(13) & _
                        "sequence: the full string you want to pull text from" & Chr(13) & _
                        "Left_or_right: specifys what side of the seperator to return")
                    MsgBox ("Example of use:" & Chr(13) & _
                        "sequence: 'zack-daugherty'" & Chr(13) & _
                        "seperator: '-'" & Chr(13) & _
                        "Left_or_Right: 'Right'" & Chr(13) & _
                        "will return: 'daugherty'")
                    Stop
                    Exit Function
                End If
            'setup variables
                Sequence_len = Len(sequence)
                Seperator_len = Len(seperator)
                seperator_pos = String_V1.has_string_inside_V1(seperator, sequence, True, False)
            'check that Seperator has a value & seperator_pos is not 0 & sequence <> nothing
                If (Sequence_len > Seperator_len) Then
                    If (Seperator_len <= 0) Then
                        Disassociate_by_Char_V1 = sequence
                    End If
                    If (seperator_pos = 0) Then
                        Disassociate_by_Char_V1 = sequence
                        Exit Function
                    End If
                Else
                    Disassociate_by_Char_V1 = sequence
                    Exit Function
                End If
            'breakup
                If (Left_or_Right = Left_C) Then
                    If (Seperator_len > 1) Then
                        Stop
                    End If
                    Disassociate_by_Char_V1 = Mid(sequence, 1, seperator_pos - 1)
                Else
                    Disassociate_by_Char_V1 = Mid(sequence, seperator_pos + Seperator_len, Sequence_len)
                End If
    'code end
        Exit Function
    'error handler
        'na
End Function

Public Function get_Special_Char_V1(ByVal char As Get_Char, Optional Dont_show_instructions As Boolean) As String
'currently functional as of (8/21/20) checked by: (zachary daugherty)
    'Created By (Zachary Daugherty)(8/21/20)
    'Purpose Case & notes:
        'get_special_char fetches from memory not on normal ENG keyboard or allowed in code
    'Library Refrences required
        'chr.object
    'Modules Required
        'na
    'Inputs
        'Internal:
            'get_char
        'required:
            'char
        'optional:
            'show_instructions
    'returned outputs
        '(information not filled out)
    'code start
        'check for show_instructions
            If (Dont_show_instructions = False) Then
                MsgBox ("Showning instructions for get_Special_Char_V1")
                MsgBox ("get_special_char fetches from memory not on normal ENG keyboard or allowed in code")
                Stop
                Exit Function
            End If
        'get
            get_Special_Char_V1 = Chr(char)
    'code end
        Exit Function
    'error handler
        'na
End Function

Public Function Shift_V1(ByVal selection_ As Shift_option, ByVal overwrite As Boolean, ByVal pos_left_up As String, ByVal pos_right_bot As String, Optional wb As Workbook, Optional current_sht As Worksheet)
'    MsgBox ("uses disaccociate string function")
    MsgBox ("add green text: string_V1, shift")
    MsgBox ("need to add dont_show_instructions")
    MsgBox ("need to revise fuction to allow the shifting of formula inside the cell if option is selected")
    Stop
    'code start
        MsgBox ("need to add dont_show_instructions")
        Stop
        'define variables
            'locational data
                Dim home_pos As Worksheet
                Dim row As Long
                Dim col As Long
                Dim upper_row As Long
                Dim upper_col As Long
                Dim lower_row As Long
                Dim lower_col As Long
            'container
                Dim L As Long
                Dim L_2 As Long
                Dim anti_loop As Long
                Dim loop_size_width As Long
                Dim loop_size_high As Long
            
        'setup variables
            'setup of locational data
                'gloabl positional
                    If (wb Is Nothing) Then
                        Set wb = ActiveWorkbook
                    End If
                    Set home_pos = wb.ActiveSheet
                    If (current_sht Is Nothing) Then
                        Set current_sht = wb.ActiveSheet
                    End If
                'local pos
                    'decode the positions
                        row = String_V1.Disassociate_by_Char_V1(",", pos_right_bot, Left)
                        col = String_V1.Disassociate_by_Char_V1(",", pos_right_bot, Right)
                        lower_row = row
                        lower_col = col
                        row = String_V1.Disassociate_by_Char_V1(",", pos_left_up, Left)
                        col = String_V1.Disassociate_by_Char_V1(",", pos_left_up, Right)
                        upper_row = row
                        upper_col = col
                        row = 1
                        col = 1
                        Set current_sht = wb.ActiveSheet
                '...
                    loop_size_width = -1
                    loop_size_high = -1
                    L = -1
                        
        'restart check point
shift_restart:
            If (anti_loop > 30) Then
                MsgBox ("Anti Loop triggered please check code")
                Stop
            End If
        'check for definition errors
'            MsgBox ("String_v1: shift: need to check if at the extremes of the worksheet and throw errors if shifting thru")
'            Stop
            If ((lower_col < upper_col) Or (lower_row < upper_row)) Then
                'swap pos col
                    L = upper_col
                    upper_col = lower_col
                    lower_col = L
                    'cleanup
                        L = -1
                'swap pos row
                    L = upper_row
                    upper_row = lower_row
                    lower_row = L
                    'cleanup
                        L = -1
                'iterate anti_loop
                    anti_loop = anti_loop + 1
                    GoTo shift_restart
            End If
            If ((upper_col = 1) And (selection_ = left_)) Then
                MsgBox ("error: module string_vX: Function shift: was trying to shift left off the page")
                Stop
            End If
            If ((upper_row = 1) And (selection_ = up)) Then
                MsgBox ("error: module string_vX: Function shift: was trying to shift up off the page")
                Stop
            End If
            If ((lower_col = 16384) And (selection_ = right_)) Then
                MsgBox ("error: module string_vX: Function shift: was trying to shift right off the page")
                Stop
            End If
            If ((lower_row = 1048576) And (selection_ = down)) Then
                MsgBox ("error: module string_vX: Function shift: was trying to shift down off the page")
                Stop
            End If
        'set loop var
            loop_size_width = lower_col - upper_col
            If (loop_size_width = 0) Then
                loop_size_width = 0
            End If
            loop_size_high = lower_row - upper_row
            If (loop_size_high = 0) Then
                loop_size_high = 0
            End If
        'goto start position and start move
            Select Case selection_
                Case Shift_option.down
                    'start from the bottom right
                        row = lower_row
                        col = lower_col
                        'iterate thru loop move
                            For L = 0 To loop_size_high
                                For L_2 = 0 To loop_size_width
                                    If (overwrite = True) Then
shift_paste_override_down:
                                        current_sht.Cells(row + 1, col).Formula = current_sht.Cells(row, col).Formula
                                        current_sht.Cells(row, col).Formula = ""
                                    Else
                                        If (current_sht.Cells(row + 1, col).value = "") Then
                                            GoTo shift_paste_override_down
                                        End If
                                    End If
                                    col = col - 1
                                Next L_2
                                row = row - 1
                                col = lower_col
                            Next L
                             
                Case Shift_option.left_
                    'start from the top left
                        row = upper_row
                        col = upper_col
                        'iterate thru loop move
                            For L = 0 To loop_size_width
                                For L_2 = 0 To loop_size_high
                                    If (overwrite = True) Then
shift_paste_override_left:
                                        current_sht.Cells(row, col - 1).Formula = current_sht.Cells(row, col).Formula
                                        current_sht.Cells(row, col).Formula = ""
                                    Else
                                        If (current_sht.Cells(row, col - 1).value = "") Then
                                            GoTo shift_paste_override_left
                                        End If
                                    End If
                                col = col + 1
                                Next L_2
                                row = row + 1
                                col = upper_col
                            Next L
                Case Shift_option.right_
                    'start from the bottom right
                        row = lower_row
                        col = lower_col
                        'iterate thru loop move
                            For L = 0 To loop_size_high
                                For L_2 = 0 To loop_size_width
                                    If (overwrite = True) Then
shift_paste_override_right:
                                        current_sht.Cells(row, col + 1).Formula = current_sht.Cells(row, col).Formula
                                        current_sht.Cells(row, col).Formula = ""
                                    Else
                                        If (current_sht.Cells(row, col + 1).value = "") Then
                                            GoTo shift_paste_override_right
                                        End If
                                    End If
                                    col = col - 1
                                Next L_2
                                row = row - 1
                                col = lower_col
                            Next L
                Case Shift_option.up
                    'start from the top left
                        row = upper_row
                        col = upper_col
                        'iterate thru loop move
                            For L = 0 To loop_size_high
                                For L_2 = 0 To loop_size_width
                                    If (overwrite = True) Then
shift_paste_override_up:
                                        current_sht.Cells(row - 1, col).Formula = current_sht.Cells(row, col).Formula
                                        current_sht.Cells(row, col).Formula = ""
                                    Else
                                        If (current_sht.Cells(row - 1, col).value = "") Then
                                            GoTo shift_paste_override_up
                                        End If
                                    End If
                                    col = col + 1
                                Next L_2
                                row = row + 1
                                col = upper_col
                            Next L
                Case Else
                    'thrown error
                        MsgBox ("option unavailable error with enumeration please check")
                    Stop
            End Select
    'code end
        Exit Function
    'error handler
    
End Function

Public Function IndentText_VA(ByVal Text As String, ByVal Limit As Integer, ByVal StoreExtraRow As Integer, ByVal StoreExtraCol As Integer) As String
'THIS CODE WAS TAKEN FROM ZEDLIB MUST BE CHECKED BEFORE USED
    MsgBox ("THIS CODE IS NOT FUNCTIONAL AS OF VERSION 1 OF STRING")
    Stop
    Exit Function
'currently functional as of (11-27-18) Checked by: Zachary Daugherty

'this function is made to take String text that is to be formatted to be on multiple lines
    '(IMPORTANT READ)
        'Make sure you are on the page you want to paste your indent text on as i does not account other pages into consideration.
        'Also the program determines the end of a word as where it can see spaces, this is how it picks if a place to indent.
    'Library Refrences required
        'Na
    'Modules Required
        'Na
    'input values
        'input 1 is the entered string text
        'input 2 is the character limit for the line
        'input 3 is cell row store extra text
        'input 4 is cell col store extra text
    'what will be returned
        'line of text data to the specified length without cutting off words only accounting blank space as a new word.
        
    'example 1:
        'input: ["thisis my textline", 7, 1,1]
        'output: "thisis"
        'extra: cells(1,1) = "my textline"
'(CODE START)
    'define values
        Dim textstorage As String
        Dim LengthOfInput As Integer
        Dim DecodeText() As String
        Dim i As Integer
        Dim Sc As Integer
        Dim ReadyToExport As Boolean
        Dim NumOfSpaces As Integer
            textstorage = "Empty"
            LengthOfInput = 0
            i = 1
            Sc = 0
            ReadyToExport = False
            NumOfSpaces = 0
    'find length of input
        LengthOfInput = Len(Text)
        ReDim DecodeText(LengthOfInput)
    'import into string into array
        For i = 1 To LengthOfInput
            textstorage = (Left(Text, i))
            If (i > 1) Then
                textstorage = (Right(textstorage, (i - (i - 1))))
            End If
            DecodeText(i) = textstorage
        Next i
        textstorage = "Empty"
        i = 1
    'check to see if requested size is possible
        For i = 1 To LengthOfInput
            textstorage = DecodeText(i)
            If (textstorage <> " ") Then
                Sc = Sc + 1
            End If
            If (textstorage = " ") Then
                'if check is true then throw error
                If (Sc > Limit) Then
                    'set everything to nothing
                    textstorage = ""
                    IndentText = ""
                    MsgBox ("Command: Functions.Function____IndentText: Request Not Possible Row size Smaller than Biggest word")
                    Stop
                    End
                End If
                If (Sc <= Limit) Then
                    Sc = 0
                End If
            End If
        Next i
        i = 1
        textstorage = "Empty"
    'fill in the line to valid value
        textstorage = ""
        On Error Resume Next
        For i = 1 To Limit
            textstorage = textstorage + DecodeText(i)
        Next i
        i = i - 1
        Sc = i
        i = 1
    'check to see that words are not cut off
        If (DecodeText(Sc + 1) = " ") Then
            ReadyToExport = True
        End If
        If ((DecodeText(Sc) <> " ") And (DecodeText(Sc + 1) <> " ")) Then
            ReadyToExport = False
        Else
            ReadyToExport = True
        End If
    'if words are cutoff move back to last available space
        If (ReadyToExport = False) Then
            For i = 1 To Limit
                If (DecodeText(i) = " ") Then
                    NumOfSpaces = NumOfSpaces + 1
                End If
            Next i
            textstorage = ""
            i = 1
            LengthOfInput = NumOfSpaces
            NumOfSpaces = 0
            For i = 1 To Limit
                If (DecodeText(i) = " ") Then
                    NumOfSpaces = NumOfSpaces + 1
                End If
                If (NumOfSpaces = LengthOfInput) Then
                    Exit For
                End If
                textstorage = textstorage + DecodeText(i)
            Next i
            Sc = i
            i = 1
            NumOfSpaces = 0
            LengthOfInput = Len(Text)
            ReadyToExport = True
        End If
    If (ReadyToExport = True) Then
        'export to data
            IndentText = textstorage
            textstorage = ""
        'find leftovers
            For i = Sc To LengthOfInput
                If ((DecodeText(i) = " ") And (NumOfSpaces = 0)) Then
                    NumOfSpaces = 1
                Else
                    textstorage = textstorage + DecodeText(i)
                End If
            Next i
        'pushout extra data
            Cells(StoreExtraRow, StoreExtraCol) = textstorage
        Exit Function
    End If
End Function


Sub test()
    Dim i As Integer
    Call Shift(left_, False, "1,1", "4,4")
End Sub








