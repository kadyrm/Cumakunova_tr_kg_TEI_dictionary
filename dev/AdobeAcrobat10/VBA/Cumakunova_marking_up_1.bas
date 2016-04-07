Attribute VB_Name = "Module2"
' Constructions used and reviewed:
' 1) [obj is Nothing] ; it's incorrect to use Null, since it is the term of database whereas Nothing relates to programming
Function mark_pointed_lines(ByVal m_pointer As String, ByVal m_new_tag As String) As Boolean
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
    Dim r As range
    Set r = find_str(m_what)
    If (r Is Nothing) = False Then
    
        Selection.HomeKey wdLine
        Selection.range.InsertBefore "<" & m_tag & ">"
        Selection.EndKey wdLine
        Selection.range.InsertAfter "</" & m_tag & ">"
        
        mark_line_containing = True
    Else
        mark_line_containing = False
    End If

End Function
Function find_str(ByVal m_what As String) As range
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
        Set find_str = Selection.range
    Else
        Set find_str = Nothing
    End If
    
End Function
