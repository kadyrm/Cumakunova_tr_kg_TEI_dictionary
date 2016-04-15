Attribute VB_Name = "Module2"
' My Notes:
' Encountered issues:
' 1) (obj is Nothing) // it's incorrect to use Null, since it is the term of database whereas Nothing relates to programming
' 2) Dim MyVar, AnotherVar As Integer  // MyVar is not Integer in this case, it is Variant
' 3) Set var= value // is used only for objects not for var types
' 4) r.SetRange  // redefine range bounaries
' 5) Chr(13)    // new line in VBA
Sub main()
    ' lets normalize paragraphs, Chr(13) corresponds to paragraph character
    m = replace_all_repeatedly(Chr(13) & Chr(13), Chr(13))
    ' lets normalize spaces
    m = replace_all_repeatedly(" " & " ", " ")
End Sub
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
        loop_limit = deadlock_saveguard(loop_limit, 1000, "main")
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

End Sub
Sub application_of_remove_tag_all()
    Dim r As Range
    'Set r = remove_tag_all("Sect")
    'Set r = remove_tag_all("Part")
    'Set r = remove_tag_all("LI")
    Set r = remove_tag_all("LI_Label")
End Sub
Function remove_tag_all(m_tag As String) As Range
' Removes all occurrencies of <m_tag>, </m_tag> and <m_tag/>
    Dim n As Long
    Dim r As Long
    Dim open_t As String
    Dim close_t As String
    Dim empty_t As String
    
    open_t = "<" & m_tag & ">"
    close_t = "</" & m_tag & ">"
    empty_t = "<" & m_tag & "/>"
    n = replace_all(open_t, "")
    MsgBox (n & " open tags were removed!")
    r = replace_all(close_t, "")
    If n <> r Then
        MsgBox "Number of close and open tags : ", , vbInformation
    End If
    MsgBox (n & " close tags were removed!")
    n = replace_all(empty_t, "")
    MsgBox (n & " empty tags were removed!")
    
    Set remove_tag_all = Selection.Range
End Function
Sub application_of_replace_tag_all()
    Dim r As Range
    'Set r = replace_tag_all("Sect")
    'Set r = replace_tag_all("H5", "P")
    'Set r = replace_tag_all("L", "P")
    Set r = replace_tag_all("LI_Title", "")
End Sub
Function replace_tag_all(m_tag As String, m_new_tag As String) As Range
' If m_new_tag is "" then removes all occurrencies of <m_tag>, </m_tag> and <m_tag/>
' Otherwise substitudes m_tag occurencies by m_new_tag value
    Dim n As Long
    Dim r As Long
    Dim open_t As String
    Dim close_t As String
    Dim empty_t As String
    Dim new_open_t As String
    Dim new_close_t As String
    Dim new_empty_t As String
    
    If m_new_tag = "" Then
        Set replace_tag_all = remove_tag_all(m_tag)
        Exit Function
    End If
    
    open_t = "<" & m_tag & ">"
    close_t = "</" & m_tag & ">"
    empty_t = "<" & m_tag & "/>"
    
    new_open_t = "<" & m_new_tag & ">"
    new_close_t = "</" & m_new_tag & ">"
    new_empty_t = "<" & m_new_tag & "/>"
    
    n = replace_all(open_t, new_open_t)
        MsgBox (n & " open tags were replaced!")
    r = replace_all(close_t, new_close_t)
        MsgBox (n & " close tags were replaced!")
    If n <> r Then
        MsgBox "Number of close and open tags don't match", , vbInformation
    End If
    n = replace_all(empty_t, new_empty_t)
        MsgBox (n & " empty tags were replaced!")
    
    Set replace_tag_all = Selection.Range
End Function

Sub apply_pattern1_1()
    Dim r As Range
    Do
        Set r = find_str(".^w^p</P>^p<P>^$")
        If r Is Nothing Then
            MsgBox ("no EOE found! Good bay !")
            Exit Sub
        End If
        Set r = insert_at(r, r.Characters.Count - 4, "</article>" & Chr(13) & "<article>" & Chr(13))
        'r.Select
        
        Selection.Range.SetRange Start:=r.Start, End:=r.End
        Selection.Collapse Direction:=wdCollapseEnd
    Loop
    
    
End Sub
Sub application_find_and_insert_at_all()
    'n = find_and_insert_at_all(".^w</P>^p<P>^$", -4, "<EOA>")
    'n = find_and_insert_at_all(".^w</P>^p<page><P>^#^#^w</P></page>^p<P>^$", -4, "<EOA>") ' when page_tag between two entries
    n = find_and_insert_at_all("?^w</P>^p<P>^$", -4, "</article>" & Chr(13) & "<article>" & Chr(13))
End Sub
Function find_and_insert_at_all(m_pattern As String, m_pos As Integer, m_insert_t As String) As Long
' if m_pos is negative the function will insert the text in the position set off from the end of the found range
    Dim r As Range
    Dim counter As Long
    counter = 0
    Do
        Set r = find_str(m_pattern)
        If r Is Nothing Then
            MsgBox (counter & " matching of the pattern were found! Good bay !")
            find_and_insert_at_all = counter
            Exit Function
        End If
        If m_pos < 0 Then
            Set r = insert_at(r, r.Characters.Count + m_pos, m_insert_t)
        Else
            Set r = insert_at(r, m_pos, m_insert_t)
        End If
        'r.Select
        
        Selection.Range.SetRange Start:=r.Start, End:=r.End
        Selection.Collapse Direction:=wdCollapseEnd
        counter = counter + 1
    Loop
    find_and_insert_at_all = counter
End Function
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
Function replace_all(ByVal m_find As String, ByVal m_replace As String) As Long
    'returns number of replacements
    replace_all = CountNoOfReplaces(m_find, m_replace)
End Function
Sub application_of_replace_all()
    'MsgBox "Number of replacements: " & replace_all(".^w</P>^p<P>^$", "EOE^&"), vbInformation
    'MsgBox "Number of replacements: " & replace_all("</article>", "</article>"), vbInformation
    'MsgBox "Number of replacements: " & replace_all("</article>", "</article>"), vbInformation
    MsgBox "Number of replacements: " & replace_all("</article>", "</article>"), vbInformation
End Sub
Function CountNoOfReplaces(StrFind As String, StrReplace As String) As Long

Dim NumCharsBefore As Long, NumCharsAfter As Long, LengthsAreEqual As Boolean

    Application.ScreenUpdating = False

    'Check whether the length of the Find and Replace strings are the same; _
    if they are, prefix the replace string with a hash (#)
    If Len(StrFind) = Len(StrReplace) Then
        LengthsAreEqual = True
        StrReplace = "#" & StrReplace
    End If

    'Get the number of chars in the doc BEFORE doing Find & Replace
    NumCharsBefore = ActiveDocument.Characters.Count

    'Do the Find and Replace
    With Selection.find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = StrFind
        .Replacement.Text = StrReplace
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute replace:=wdReplaceAll
    End With

    'Get the number of chars AFTER doing Find & Replace
    NumCharsAfter = ActiveDocument.Characters.Count

    'Calculate of the number of replacements,
    'and put the result into the function name variable
    CountNoOfReplaces = (NumCharsBefore - NumCharsAfter) / _
            (Len(StrFind) - Len(StrReplace))

    'If the lengths of the find & replace strings were equal at the start, _
    do another replace to strip out the #
    If LengthsAreEqual Then

        StrFind = StrReplace
        'Strip off the hash
        StrReplace = Mid$(StrReplace, 2)

        With Selection.find
            .Text = StrFind
            .Replacement.Text = StrReplace
            .Execute replace:=wdReplaceAll
        End With

    End If

    Application.ScreenUpdating = True
    'Free up memory
    ActiveDocument.UndoClear

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
Function find_EOE_fs_cr_l(m_r As Range) As Range
    
End Function
Function find_EOE_terminated_by_white_space(m_r As Range) As Range
' source file to be marked: Some Web Engine
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
        i = deadlock_saveguard(i, 100, "white_space_eof")
    Loop
    
    
End Function
Function deadlock_saveguard(ByVal m_counter As Integer, ByVal m_max_loop As Integer, ByVal m_func_name As String) As Integer
m_counter = m_counter + 1
deadlock_saveguard = m_counter
If m_counter Mod m_max_loop = 0 Then
        If MsgBox("Do you want to continue the loop in " & m_func_name, vbYesNo, "Debugging") = vbNo Then
           Stop
        End If
End If
End Function
Function replace_all_repeatedly(m_what As String, m_by As String)
    ' This procedure replaces "m_what" repeatedly by "m_by" while there are absolutely no new occurrencies
    
    ' lets define limits for deadlock_saveguard
    Dim max, current As Integer
    max = 2000
    current = 0
    ' lets define variable for number of replacements
    Dim l, m As Long
    l = 1 ' lets set it to 1 in order to enter the loop
    m = 0 ' this variable holds sum of all  replacements
    Do While l <> 0
        l = replace_all(m_what, m_by)
        m = m + l
        MsgBox ("Loop number: " & current + 1 & Chr(13) & "Replacements made: " & l)
        current = deadlock_saveguard(current, max, "replace_all_repeatedly")
    Loop
    replace_all_repeatedly = m
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
