Attribute VB_Name = "Module2"
' My Notes:
' Encountered issues:
' 1) (obj is Nothing) // it's incorrect to use Null, since it is the term of database whereas Nothing relates to programming
' 2) Dim MyVar, AnotherVar As Integer  // MyVar is not Integer in this case, it is Variant
' 3) Set var= value // is used only for objects not for var types
' 4) r.SetRange  // redefine range bounaries
Sub Show_ascw()
    Dim kod As Long
    kod = AscW(Selection.Text)
    str1 = ChrW(kod)
    MsgBox "Selected text: " & Selection.Text & Chr(13) & "ASCW: " & kod
    'Selection.Range.InsertAfter (kod)
End Sub
Function mark_lines_pointed(ByVal m_pointer As String, ByVal m_new_tag As String) As Boolean
Dim exit_code As Boolean
Selection.HomeKey wdStory
exit_code = mark_line_containing(m_pointer, m_new_tag)
loop_limit = 0
    Do While r
        exit_code = mark_line_containing(m_pointer, m_new_tag)
        loop_limit = loop_checker(loop_limit, 1000, "main")
    Loop
mark_pointed_lines = exit_code
End Function

Function mark_line_containing(ByVal m_what, ByVal m_tag) As Boolean
    Dim r As Range
    Set r = find_str(m_what)
    If (r Is Nothing) = False Then
    
        Selection.HomeKey wdLine
        Selection.Range.InsertBefore "<" & m_tag & ">"
        Selection.EndKey wdLine
        Selection.Range.InsertAfter "</" & m_tag & ">"
        
        mark_line_containing = True
    Else
        mark_line_containing = False
    End If

End Function
Sub test_insert_at()
    Dim r As Range
    Do
        Set r = find_EOE_terminated_by_white_space(Selection.Range)
        If r Is Nothing Then
            MsgBox ("no EOE found! Good bay !")
            Exit Sub
        End If
        Set r = insert_at(r, r.Characters.Count - 1, "EOE")
        r.Select
        'Selection.range.SetRange Start:=r.Start, End:=r.End
        Selection.Collapse Direction:=wdCollapseEnd
    Loop
    
    
End Sub
Function insert_at(ByRef m_rng As Range, ByVal m_pos As Integer, ByVal m_what As String) As Range
    If m_pos > m_rng.Characters.Count Then
        Set insert_at = Nothing
        Exit Function
    End If
    
    Dim r As Range
    Set r = m_rng
    r.SetRange Start:=m_rng.Start + m_pos, End:=m_rng.End
    r.InsertBefore (m_what)
    m_rng.SetRange Start:=m_rng.Start, End:=m_rng.End + Len(m_what)
    
    Set insert_at = m_rng
End Function
Function find_str(ByVal m_what As String) As Range
Selection.find.ClearFormatting
    With Selection.find
        .Text = m_what
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    If Selection.find.found Then
        Set find_str = Selection.Range
    Else
        Set find_str = Nothing
    End If
    
End Function
Sub test_find_white_space_eoe()
    Dim r As Range
    Set r = find_EOE_terminated_by_white_space(Selection.Range)
End Sub

Function find_EOE_terminated_by_white_space(m_r As Range) As Range
' EOE - End of Entry
    Dim r As Range
    Dim flag As Boolean
    Dim i As Integer
    flag = False
    i = 0
    Do
        Set r = find_str(".^w^$")
        If r Is Nothing Then
            Set find_EOE_terminated_by_white_space = Nothing
            Exit Function
        End If
        If r.Characters.Last.Bold Then
            'MsgBox ("Last char of range is bold.")
            If is_in_one_line(r) = False Then
                MsgBox ("Found range captures two or more lines.EOE is found!")
                Set find_EOE_terminated_by_white_space = r
                Exit Do
            End If
        Else
            Set find_EOE_terminated_by_white_space = Nothing
        End If
        i = loop_checker(i, 100, "white_space_eof")
    Loop
    
    
End Function
Function loop_checker(ByVal m_counter As Integer, ByVal m_max_loop As Integer, ByVal m_func_name As String) As Integer
m_counter = m_counter + 1
loop_checker = m_counter
If m_counter Mod m_max_loop = 0 Then
        If MsgBox("Do you want to continue the loop in " & m_func_name, vbYesNo, "Debugging") = vbNo Then
           Stop
        End If
End If
End Function
Sub test_find_double_carret_return()

    Dim r As Range
    Do
        Set r = find_by_ascw(13)
        If r Is Nothing Then
            Exit Sub
        End If
        Selection.Collapse wdCollapseEnd
        Selection.MoveRight wdCharacter, 1, True
        code = AscW(Selection.Text)
        If code = 13 Then
            MsgBox "double par found!!"
            Exit Sub
        End If
        Selection.Collapse wdCollapseEnd
    Loop
End Sub
Function find_by_ascw(m_ascw As Long) As Range
   
    Dim r As Range
    str1 = ChrW(m_ascw)
    Set r = find_str(str1)
    Set find_by_ascw = r
    
End Function
Sub test_fs_ws_letter()
    Dim r As Range
    Set r = find_str(".^w^$")
    If r.Bold Then
        MsgBox ("Found range is bold.")
    End If
    If r.Characters.Last.Bold Then
        MsgBox ("Last char of range is bold.")
    End If
    If is_in_one_line(r) = False Then
        MsgBox ("Found range is not in one line.")
    End If
    
End Sub
Function is_in_one_line(r As Range) As Boolean
    Dim l1, l2 As Integer
    Dim r1 As Range
    Dim r2 As Range
    Set r1 = r.Characters.First
    Set r2 = r.Characters.Last
    l1 = GetLineNum(r1)
    l2 = GetLineNum(r2)
    
    If l1 = l2 Then
        is_in_one_line = True
    Else
        is_in_one_line = False
    End If
End Function
Sub WhereAmI()
    
    MsgBox "Paragraph number: " & GetParNum(Selection.Range) & vbCrLf & _
    "Absolute line number: " & GetAbsoluteLineNum(Selection.Range) & vbCrLf & _
    "Relative line number: " & GetLineNum(Selection.Range)
End Sub
 
 
Function GetParNum(r As Range) As Integer
    Dim rParagraphs As Range
    Dim CurPos As Long
     
    r.Select
    CurPos = ActiveDocument.Bookmarks("\startOfSel").Start
    Set rParagraphs = ActiveDocument.Range(Start:=0, End:=CurPos)
    GetParNum = rParagraphs.Paragraphs.Count
End Function
 
Function GetLineNum(r As Range) As Integer
     'relative to current page
    GetLineNum = r.Information(wdFirstCharacterLineNumber)
End Function
 
Function GetAbsoluteLineNum(r As Range) As Integer
    Dim i1 As Integer, i2 As Integer, Count As Integer, rTemp As Range
     
    r.Select
    Do
        i1 = Selection.Information(wdFirstCharacterLineNumber)
        Selection.GoTo what:=wdGoToLine, which:=wdGoToPrevious, Count:=1, Name:=""
         
        Count = Count + 1
        i2 = Selection.Information(wdFirstCharacterLineNumber)
    Loop Until i1 = i2
     
    r.Select
    GetAbsoluteLineNum = Count
End Function
