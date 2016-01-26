Attribute VB_Name = "Cumakunava TEI Markinup API version_1"
Sub MarkTerms()
' ----- Now. There is a dash defect in sub, that is, it may insert </term> before or after dash.
Application.Run "RemoveDoubleSpaces"
Selection.HomeKey unit:=wdStory
Selection.find.ClearFormatting
Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "^p "
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute replace:=wdReplaceAll
    Selection.HomeKey unit:=wdStory
    Selection.find.Text = "^p^?^?"
    Selection.find.Execute
Dim iLoop As Integer
iLoop = 0
Do While Selection.find.found
    Selection.Collapse Direction:=wdCollapseStart
    Selection.MoveRight unit:=wdCharacter, Count:=1
    ' *** Chek if there is a normal hyphen char
        Dim str As String
        Dim iFlag As Integer
        Selection.EndKey unit:=wdLine, Extend:=wdExtend
        str = Selection.Text
        iFlag = InStr(str, "–")
    ' ***
    Selection.HomeKey unit:=wdLine
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
    If iFlag <> 0 And Selection.Range.Bold Then
        Selection.InsertBefore ("<term>")
        ' * May be additional code
        Do While (Selection.Range.Bold And Selection.Text <> "–") Or (Not Selection.Range.Bold And Selection.Text = " ")
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Loop
        Selection.InsertBefore ("</term>")
    Else
        ' Do nothing
    End If
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.find.Execute
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 10 = 0 Then
            If MsgBox("Do you want to continue the loop? func name: MarkTerms", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop
End Sub
Sub MarkGraphicObjects()
' -----
Application.Run "RemoveDoubleSpace"
Selection.HomeKey unit:=wdStory
Selection.find.ClearFormatting
Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "^g"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    ' ***
    Selection.find.Execute replace:=wdReplaceAll
    Selection.HomeKey unit:=wdStory
    Selection.find.Text = "^p^?^?"
    Selection.find.Execute
Dim iLoop As Integer
iLoop = 0
Do While Selection.find.found
    Selection.Collapse Direction:=wdCollapseStart
    Selection.MoveRight unit:=wdCharacter, Count:=1
    ' *** Chek if there is a normal hyphen char
        Dim str As String
        Dim iFlag As Integer
        Selection.EndKey unit:=wdLine, Extend:=wdExtend
        str = Selection.Text
        iFlag = InStr(str, "–")
    ' ***
    Selection.HomeKey unit:=wdLine
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
    If iFlag <> 0 And Selection.Range.Bold Then
        Selection.InsertBefore ("<term>")
        ' * May be additional code
        Do While (Selection.Range.Bold And Selection.Text <> "–") Or (Not Selection.Range.Bold And Selection.Text = " ")
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Loop
        Selection.InsertBefore ("</term>")
    Else
        ' Do nothing
    End If
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.find.Execute
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 10 = 0 Then
            If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop
End Sub

Sub NormalizeDash()
'
'
'

Dim chLeft, chRight As String
Dim iLoop As Integer
iLoop = 0
'***
' 1. Replace long dashes to normal length ones
Selection.HomeKey unit:=wdStory
Selection.find.Forward = True
FindLongHyphen (wdFindStop)
Do While Selection.find.found
    Selection.Collapse Direction:=wdCollapseStart
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    chLeft = Selection.Text
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.MoveRight unit:=wdCharacter, Count:=1
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
    chRight = Selection.Text
    If chLeft = " " And chRight = " " Then
        Selection.MoveLeft unit:=wdCharacter, Count:=2, Extend:=wdExtend
        Selection.Text = "–"
    End If
    FindLongHyphen (wdFindStop)
    ' ---- Loop checker
    iLoop = iLoop + 1
    If iLoop Mod 10 = 0 Then
        If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
            Exit Do
        End If
    End If
    ' ----
Loop

Selection.HomeKey unit:=wdStory
'***
' 2.Replace short dashes to normal length ones
Selection.find.Forward = True
FindShortHyphen (wdFindStop)
Do While Selection.find.found
    Selection.Collapse Direction:=wdCollapseStart
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    chLeft = Selection.Text
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.MoveRight unit:=wdCharacter, Count:=1
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
    chRight = Selection.Text
    If chLeft = " " And chRight = " " Then
        Selection.MoveLeft unit:=wdCharacter, Count:=2, Extend:=wdExtend
        Selection.Text = "–"
    End If
    Selection.Collapse Direction:=wdCollapseEnd
    FindShortHyphen (wdFindStop)
    ' ---- Loop checker
    iLoop = iLoop + 1
    If iLoop Mod 50 = 0 Then
        If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
            Exit Do
        End If
    End If
    ' ----
    ' MsgBox "Hyphen's Found"
Loop

End Sub

Sub Move3StepsLeft1()
Attribute Move3StepsLeft1.VB_Description = "Макрос записан 23.02.2011 Customer"
Attribute Move3StepsLeft1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Move3StepLeft1"
'
' Move3StepLeft1 Макрос
' Макрос записан 23.02.2011 Customer
'
    Selection.MoveLeft unit:=wdCharacter, Count:=3
End Sub

Sub pick_term_all()
Attribute pick_term_all.VB_Description = "Макрос записан 19.04.2011 Customer"
Attribute pick_term_all.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос4"
'
' Макрос4 Макрос
' Макрос записан 19.04.2011 Customer
'
    Do While pick_term() = True
    
    
    Loop
    
End Sub
Sub MarkTables()
'copies all tables in a document to a new document
Dim oDoc1 As Document
Dim myRange As Range
Dim oTable As Table
If ActiveDocument.Tables.Count >= 1 Then
    Set oDoc1 = ActiveDocument
    
    'Set myRange = Documents.Add.Range(Start:=0, End:=0)
    For Each oTable In oDoc1.Tables
        Set myRange = ActiveDocument.Range(oTable.Range.Start - 1, oTable.Range.End + 1)
        myRange.InsertBefore ("<table>")
        myRange.InsertAfter ("</table>")
        'Set myRange = MsgBox(oTable.Range.Start)
        'oTable.Range.InsertParagraphBefore
        'oTable.Range.InsertBefore ("<table>")
        'oTable.Range.InsertParagraphAfter
        'oTable.Range.InsertAfter ("</table>")
    Next
End If
End Sub

Sub FindTables()
'copies all tables in a document to a new document
Dim oDoc1 As Document
Dim myRange As Range
Dim oTable As Table
If ActiveDocument.Tables.Count >= 1 Then
    Set oDoc1 = ActiveDocument
    Set myRange = Documents.Add.Range(Start:=0, End:=0)
    For Each oTable In oDoc1.Tables
        oTable.Range.Copy
        With myRange
            .Paste
            .Collapse Direction:=wdCollapseEnd
            .InsertParagraphAfter
            .Collapse Direction:=wdCollapseEnd
        End With
    Next
End If
End Sub

Sub FindNormalHyphen()
Attribute FindNormalHyphen.VB_Description = "Макрос записан 21.04.2011 Customer"
Attribute FindNormalHyphen.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос7"
'
' Макрос7 Макрос
' Макрос записан 21.04.2011 Customer
'
    Selection.find.ClearFormatting
    With Selection.find
        .Text = "–"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
End Sub
Sub RemoveSoftHyphen()
Attribute RemoveSoftHyphen.VB_Description = "Макрос записан 26.04.2011 Customer"
Attribute RemoveSoftHyphen.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос6"
'
' Макрос6 Макрос
' Макрос записан 26.04.2011 Customer
'
' ***Finds and replaces soft hyphens in document***

    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "^-"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    With Selection
        If .find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .find.Execute replace:=wdReplaceOne
        If .find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .find.Execute
    End With
    If Selection.find.found = True Then
        
    End If
    Selection.find.Execute replace:=wdReplaceAll
End Sub
Sub RemoveDoubleSpaces()
'
' User Макрос
' Макрос записан 26.04.2011 Customer
'
' *** Removes unnesessary spaces ***

    Selection.HomeKey unit:=wdStory
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    Do While Selection.find.found
        With Selection
            If .find.Forward = True Then
                .Collapse Direction:=wdCollapseStart
            Else
                .Collapse Direction:=wdCollapseEnd
            End If
            .find.Execute replace:=wdReplaceOne
            If .find.Forward = True Then
                .Collapse Direction:=wdCollapseEnd
            Else
                .Collapse Direction:=wdCollapseStart
            End If
            .find.Execute
        End With
        If Selection.find.found = True Then
            
        End If
        Selection.find.Execute replace:=wdReplaceAll
        Selection.HomeKey unit:=wdStory
        Selection.find.Execute
    Loop
    Selection.HomeKey unit:=wdStory
    
End Sub
Sub RemoveDoubleParagraph()
'
' User Макрос
' Макрос записан 26.04.2011 Customer
'
' *** Removes unnesessary paragraphes ***

    Selection.HomeKey unit:=wdStory
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    
    Do While Selection.find.found
        Selection.find.Execute replace:=wdReplaceAll
        'Selection.HomeKey unit:=wdStory
        Selection.find.Execute
        ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 5 = 0 Then
            If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
        ' ----
    Loop
    
    Selection.HomeKey unit:=wdStory
    
End Sub
Sub NormalizeTermParagraph_All()
'
' User Макрос
' Макрос записан 26.04.2011 Customer
'
' *** Replaces normal the paragraphs to SrarDict's format, " \n "***
    ' *** Mark the end of document.***
    Selection.EndKey unit:=wdStory
    Selection.InsertAfter "<End.>"
    ' ***
    Selection.HomeKey unit:=wdStory
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    Do While Selection.find.found
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.MoveRight unit:=wdCharacter, Count:=6, Extend:=wdExtend
        If Selection.Text = "<End.>" Then
            Exit Do
        End If
        If Selection.Text <> "<term>" And Selection.Tables.Count = 0 Then
            Selection.Collapse Direction:=wdCollapseStart
            Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Text = " \n   "
        End If
        
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.find.Execute
        ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
        ' ----
    Loop
    Selection.Delete
    Selection.HomeKey unit:=wdStory
    
End Sub
Sub NormalizeHyphen()
'
' User Макрос
' Макрос записан 26.04.2011 Customer
' Rules:
' "-" short length hyphen devides complex words
' "-" normal length hyphen's function is like dash
' "-" long hyphens are excepted from document
' *** ***
'
'
'

Dim chLeft, chRight As String
Dim iLoop As Integer
iLoop = 0
'***
' 1. Replace long hyphens to short length ones
Selection.HomeKey unit:=wdStory
Selection.find.Forward = True
Selection.find.ClearFormatting
    With Selection.find
        .Text = "^$—^$"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    
Do While Selection.find.found
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    If chLeft = "^$" And chRight = "^$" Then
        
    End If
    Selection.Text = "-"
    
    ' ---- Loop checker
    iLoop = iLoop + 1
    If iLoop Mod 10 = 0 Then
        If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
            Exit Do
        End If
    End If
    ' ----
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.find.Execute
Loop

Selection.HomeKey unit:=wdStory
'***
' 2.Replace normal hyphens to short length ones
Selection.find.ClearFormatting
    With Selection.find
        .Text = "^$–^$"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
Do While Selection.find.found
   Selection.Collapse Direction:=wdCollapseEnd
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Text = "-"
    Selection.Collapse Direction:=wdCollapseEnd
    ' ---- Loop checker
    iLoop = iLoop + 1
    If iLoop Mod 50 = 0 Then
        If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
            Exit Do
        End If
    End If
    ' ----
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.find.Execute
    
Loop



' 2.
End Sub

Sub FindLongHyphen(wdWrap As WdFindWrap)
Attribute FindLongHyphen.VB_Description = "Макрос записан 26.04.2011 Customer"
Attribute FindLongHyphen.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос6"
'
' Макрос6 Макрос
' Макрос записан 26.04.2011 Customer
'
    Selection.find.ClearFormatting
    With Selection.find
        .Text = "—"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdWrap
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
End Sub
Sub FindShortHyphen(wdWrap As WdFindWrap)
'
' Макрос6 Макрос
' Макрос записан 26.04.2011 Customer
'
    Selection.find.ClearFormatting
    With Selection.find
        .Text = "-"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdWrap
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
End Sub


Sub NormalizeTermHyphen()
Attribute NormalizeTermHyphen.VB_Description = "Макрос записан 28.04.2011 Customer"
Attribute NormalizeTermHyphen.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос7"
'
' Макрос7 Макрос
' Макрос записан 28.04.2011 Customer
' ***Constraint: A text considered as term when it's bold
'
    Dim iLoop As Integer
    iLoop = 0
    Selection.HomeKey unit:=wdStory
    Selection.find.ClearFormatting
    With Selection.find
            .Text = "^$ ^= ^$"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.find.Execute
    '***
    Do While Selection.find.found
        
        If Selection.Range.Bold = True Then
            MsgBox Selection.Range.Text + "  - Bold"
            Selection.Collapse Direction:=wdCollapseStart
            Selection.MoveRight unit:=wdCharacter, Count:=1
            Selection.MoveRight unit:=wdCharacter, Count:=3, Extend:=wdExtend
            Selection.Text = "-"
        End If
        iLoop = iLoop + 1
         ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 10 = 0 Then
            If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
        ' ----
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.find.Execute
    Loop
    
End Sub

Sub NormilizeReversion()
'
' * Tested *
' Макрос7 Макрос
' Макрос записан 28.04.2011 Customer
' Inserts dash char in reversion(reference) terms
'
    Dim iLoop As Integer
    iLoop = 0
    Selection.HomeKey unit:=wdStory
    Application.Run macroname:="RemoveDoubleSpaces"
    Selection.HomeKey unit:=wdStory
    FindReversion (wdFindStop)
    ' ***
    Do While Selection.find.found
        Selection.Collapse Direction:=wdCollapseStart
        'Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend

        If Selection.Text = ")" Or Selection.Range.Bold Then
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.InsertAfter (" – ")
            
        End If
        
         ' ---- Loop checker----
        iLoop = iLoop + 1
        If iLoop Mod 10 = 0 Then
            If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
        '      ----***-----
        
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.MoveRight unit:=wdCharacter, Count:=4
        ' ---- Now don't stop at end of document----
        FindReversion (wdFindStop)
    Loop
    
End Sub
Sub FindReversion(wdWrap As WdFindWrap)
'
' Макрос7 Макрос
' Макрос записан 28.04.2011 Customer
'
    Selection.find.ClearFormatting
    With Selection.find
        .Text = ", к."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdWrap
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    
End Sub
Sub ReplaceAll()
Attribute ReplaceAll.VB_Description = "Макрос записан 29.04.2011 Customer"
Attribute ReplaceAll.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос7"
'
' Макрос7 Макрос
' Макрос записан 29.04.2011 Customer
'
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "h"
        .Replacement.Text = "hh"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute replace:=wdReplaceOne
End Sub
Sub main_proc()
Application.Run macroname:="RemoveDoubleSpaces"
Application.Run macroname:="NormalizeParagraphs" ' removes spaces before paragraphs
Application.Run macroname:="NormalizeCommas"     ' removes spaces before commas
'0.
Application.Run macroname:="RemoveBigLetter"
'0.1
Application.Run macroname:="RemoveDoubleParagraph"
'1.
Application.Run macroname:="RemoveSoftHyphen"
'2.
Application.Run macroname:="NormalizeHyphen"
'3
Application.Run macroname:="NormalizeDash"
'4.
Application.Run macroname:="NormilizeReversion"
'5.
Application.Run macroname:="NormalizeTermHyphen"

'7.
Application.Run macroname:="MarkTerms"
'8
Application.Run macroname:="NormalizeTermParagraph_All"


 
End Sub
Sub RemoveBigLetter()
Attribute RemoveBigLetter.VB_Description = "Макрос записан 03.05.2011 Customer"
Attribute RemoveBigLetter.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос7"
'
' ! Note: This sub is not completed
' Макрос записан 03.05.2011 Customer
'
    Selection.HomeKey unit:=wdStory
    Selection.find.ClearFormatting
    With Selection.find
        .Text = "^p^$^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    Do While Selection.find.found
        
        If Selection.Range.font.Size >= 20 Then
            MsgBox "found"
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.MoveLeft unit:=wdCharacter, Count:=1
            Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Delete
        End If
        Selection.MoveRight unit:=wdCharacter, Count:=1
        Selection.find.Execute
        
        ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
    Loop
End Sub
Sub NormalizeParagraphs()
'
'removes spaces before paragraphs
'
Application.Run macroname:="RemoveDoubleSpaces"
Selection.find.ClearFormatting
    With Selection.find
            .Text = " ^p"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.find.Execute replace:=wdReplaceAll
End Sub
Sub RemoveHeaders()
Attribute RemoveHeaders.VB_Description = "Макрос записан 19.05.2011 Customer"
Attribute RemoveHeaders.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос7"
'
' Макрос7 Макрос
' Макрос записан 19.05.2011 Customer
' Now - not complete.
    Selection.find.ClearFormatting
    With Selection.find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = wdAlignParagraphCenter
    End With
    Selection.find.ParagraphFormat.Borders.Shadow = False
    With Selection.find
        .Text = "^?"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    Do While Selection.find.found
        Selection.Delete
    Loop
End Sub
Sub NormalizeCommas()
'
' removes space before commas
'
Selection.HomeKey unit:=wdStory
Selection.find.ClearFormatting
    With Selection.find
            .Text = " ,"
            .Replacement.Text = ","
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.find.Execute replace:=wdReplaceAll
        
End Sub
Sub used_new_code_elements()
Attribute used_new_code_elements.VB_Description = "Макрос записан 20.05.2011 Customer"
Attribute used_new_code_elements.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос7"
'
'
'
'   1. Character functions
    '    Dim str As String
    '    kod = AscW("X")
    '    kod = Asc("X")
    '    i = InStr(str, "–")
    '    c = Chr(65)
    '    c = ChrW$(kod)
    
'    2. how to invoke sub
'       Call MarkArticle_all  ' Equal to:
'       Application.Run "MarkArticle_all"
    
'    3. how determine end of document
'       ActiveDocument.Range.StoryLength ' used for determine the end of doc
'       ActiveDocument.Characters.Count  ' not recommended. Used for determine the end of doc, slowly
'       Selection.Style = wdStyleNormal  ' recommended
'
'   4. How to return an object from function
    '   Set is used when to assign an object
    '   Set sel = object_return()
    '   Set new_font = New font
    '   If m_font Is Nothing Then
'   5. how to input methods of vba
'       Inputbox() doesn't support characters except ANSI set.
'   6. program flow interruption fucntions
'       Stop ' to pause program progress
'       End  ' to stop program entirely
'       how to quit from the loop ?
'   7. how to shift selection to certain position
        ' dim pos as range
        ' pos.Select
    
End Sub
Sub ReplaceCharSet()
Dim target As String
Dim source As String
source = InputBox("Enter the character set to make conversion from", "Initial Character Set")
target = InputBox("Enter the character set to make conversion to", "Target Character Set")
Selection.InsertAfter (source & vbNewLine & target)


End Sub
Sub Show_ascw()
Dim kod As Long
kod = AscW(Selection.Text)
str1 = ChrW(kod)
MsgBox "Selected text: " & Selection.Text & Chr(13) & "ASCW: " & kod
'Selection.Range.InsertAfter (kod)
End Sub
Sub test_tmp()
Attribute test_tmp.VB_Description = "Макрос записан 16.06.2011 Customer"
Attribute test_tmp.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос8"
'
' Макрос8 Макрос
' Макрос записан 16.06.2011 Customer
'
    '<draft>
Dim str1, str2 As String
str1 = Selection.Text
i = find_and_mark_term()

str2 = "^$"
If is_term_char(Selection.Characters.Last) Then
    MsgBox "The condition's supported."
End If
Selection.Style = wdStyleNormal
kod = Asc(Selection.Text)
MsgBox Selection.Text + ""
MsgBox kod
kod = AscW(Selection.Text)
str1 = ChrW(kod)
MsgBox kod & str1
Selection.Range.InsertAfter ChrW(1186)

Dim c As String
Dim str As String
Dim code As Long
c = Left(Selection.Characters(1).Text, 1)
If is_alpha_kk(c) Then
    MsgBox "yes"
End If
'MsgBox wdMainTextStory
'MsgBox selection.Start equivalent to MsgBox selection.Range.Start
'selection.End = wdMainTextStory
b = select_article_tag()
If b Then
    b = mark_key_in_article_tag(Selection.Start, Selection.End)
End If

'</draft>

    
    Selection.find.ClearFormatting
    With Selection.find
        .Text = ChrW$(8221)
        .Replacement.Text = ChrW$(8221)
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .font.Bold = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
e1:     Selection.find.Execute replace:=wdReplaceAll
    Selection.Collapse Direction:=wdCollapseEnd
    GoTo e1
    If Selection.Range.Case = wdUpperCase And Selection.Range.Bold = True Then
        Selection.Range.Collapse Direction:=wdCollapseStart
        Selection.Move unit:=wdCharacter, Count:=1
        If Selection.Text = "^p" Then
         MsgBox "it's a term"
        End If
        
        MsgBox Selection.Range.Text
    End If
    
End Sub
'*********************************** <stardict> </stardict>***********************
Sub MarkArticles()

'*********************
' Mark shell - outer <article> tags
b = True
Do While b
    b = mark_article()
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 500 = 0 Then
            If MsgBox("Do you want to continue the loop in MarkArticle_all?", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop
' Mark inside - <key> and <def> articles
Selection.HomeKey unit:=wdStory

b = True
counter = 0
Do While b
    b = mark_key_and_def_in_article_tag()
    'Loop checker
    counter = loop_checker(counter, 1000, "MarkArticles")
Loop

End Sub
Sub Test_Searching()

Dim r As Range
Dim f As font
f = Selection.font
Set r = find("^?", , , f)
'r = select_article_tag()
MsgBox (r)
'Call select_uppercased_text(Selection)
End Sub
Sub Test_Scripts_Marking()

Dim r As Boolean
r = Mark_Script(4, "sup")

MsgBox (r)

End Sub
Function select_article_tag() As Boolean
Selection.find.ClearFormatting
    With Selection.find
        .Text = "<article>"
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
    Dim art_range As Range
    Dim s, e As Long
    s = Selection.Start
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.find.Text = "</article>"
    Selection.find.Execute
    If Not Selection.find.found Then
        MsgBox ("<article> without </article>. func_name: select_article_tag()")
        select_article_tag = False
        Exit Function
    End If
    e = Selection.End
    Set art_range = ActiveDocument.Range(s, e)
    Selection.Start = s
    Selection.End = e
    select_article_tag = True
Else
    select_article_tag = False
End If
End Function
Sub MarkArticle()

'Finds and marks term article with <article> </article> tag.
b = find_and_mark_term()
If b = True Then
    ActiveDocument.Undo (4)
    Selection.HomeKey unit:=wdLine
    Selection.Range.InsertAfter "<article>"
    Selection.MoveRight unit:=wdCharacter, Count:=Len("<article>"), Extend:=wdExtend
    Selection.Range.Case = wdLowerCase
    Selection.Range.InsertParagraphAfter
    Selection.MoveDown unit:=wdLine, Count:=1
    
    b = find_and_mark_term()
    If b Then
        ActiveDocument.Undo (4)
        Selection.HomeKey unit:=wdLine
        Selection.Range.InsertParagraphAfter
        Selection.Range.InsertBefore "</article>"
        Selection.MoveRight unit:=wdCharacter, Count:=Len("</article>"), Extend:=wdExtend
        Selection.Range.Case = wdLowerCase
        Selection.Collapse Direction:=wdCollapseEnd
    Else
        Selection.Range.InsertAfter "</article>"
        Selection.MoveRight unit:=wdCharacter, Count:=Len("<article>") + 1
        Selection.Range.InsertParagraphAfter
    End If
Else
    MsgBox "Cannot find article."
End If
'selection.MoveDown unit:=wdLine, Count:=1

End Sub
Function mark_article() As Boolean

'Finds <article> </article> tag.
b = find_and_mark_term()
If b = True Then
    ActiveDocument.Undo (4)
    Selection.HomeKey unit:=wdLine
    Selection.Range.InsertAfter "<article>"
    Selection.MoveRight unit:=wdCharacter, Count:=Len("<article>"), Extend:=wdExtend
    Selection.Range.Case = wdLowerCase
    Selection.Range.InsertParagraphAfter
    Selection.MoveDown unit:=wdLine, Count:=1
    
    b = find_and_mark_term()
    If b Then
        ActiveDocument.Undo (4)
        Selection.HomeKey unit:=wdLine
        Selection.Range.InsertAfter "</article>"
        Selection.MoveRight unit:=wdCharacter, Count:=Len("</article>"), Extend:=wdExtend
        Selection.Range.Case = wdLowerCase
        Selection.Range.InsertParagraphAfter
    Else
        Selection.Range.InsertAfter "</article>"
        Selection.MoveRight unit:=wdCharacter, Count:=Len("</article>"), Extend:=wdExtend
        Selection.Range.Case = wdLowerCase
        Selection.Range.InsertParagraphAfter
    End If
    mark_article = True
Else
    mark_article = False
    MsgBox "Cannot fine article."
End If
'selection.MoveDown unit:=wdLine, Count:=1

End Function
'**********************************<key></key>********************************
Function mark_key_in_article_tag(s_pos As Long, e_pos As Long) As Boolean
Selection.Start = s_pos
Selection.Collapse Direction:=wdCollapseStart
b = find_and_mark_term()
If Selection.End < e_pos Then
    'OK
    mark_key_in_article_tag = True
Else
    'Ups
    ActiveDocument.Undo (2)
    mark_key_in_article_tag = False
End If
End Function

Sub MarkKey()

'****************************************
'Discription
'Finds and marks <key> </key> tag.
'pos = find_any_letter()
'If selection.Range.Case = wdUpperCase And selection.Range.Bold Then
 '   sel_str = select_uppercase_bold(selection)
  '  selection.Range.InsertBefore "<key>"
   ' selection.Range.InsertAfter "</key>"
'Else
    Dim marked As Boolean
    marked = False
    Do While Not marked
        Selection.Collapse Direction:=wdCollapseEnd
        pos = find_paragraph()
        Selection.Collapse Direction:=wdCollapseEnd
        pos = find_any_letter(True, Selection)
        sel_str = select_uppercase_bold(Selection)
        Dim next_char As Range
        Set next_char = ActiveDocument.Range(Selection.End, Selection.End + 1)
        If sel_str <> "" And Not is_alpha(next_char.Text) And Selection.Tables.Count = 0 Then
            Selection.Range.InsertBefore "<key>"
            Selection.Range.InsertAfter "</key>"
            marked = True
        End If
        ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop in MarkKey?", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
    Loop
    
'End If
Selection.MoveRight unit:=wdCharacter, Count:=Len("</key>") + 1
End Sub
Function is_end(ByRef m_sel As Selection) As Boolean
'Cheks is the current selection reached the end of the document
'
If m_sel.Range.End > ActiveDocument.Range.StoryLength - 2 Then
    is_end = True
Else
    is_end = False
End If
End Function
Function mark_key() As Boolean

'****************************************
'Discription
'Finds and marks <key> </key> tag.
' 2nd version. Quote bag's removed.
Dim marked As Boolean
marked = False

Selection.Collapse Direction:=wdCollapseEnd
pos = find_paragraph()
Do While Not marked And Selection.Range.End <= ActiveDocument.Range.StoryLength - 2
    
    Selection.Collapse Direction:=wdCollapseEnd
    'pos = find_any_letter()
    Selection.MoveRight unit:=wdCharacter, Count:=2, Extend:=wdExtend
    sel_str = select_term(Selection)
    Dim tmp_range As Range
    Set tmp_range = ActiveDocument.Range(Selection.End, Selection.End + 1)
    If sel_str <> "" And Not is_alpha(tmp_range.Text) And Selection.Tables.Count = 0 Then
        '********** Marking***************
        Dim key_l As Integer
        key_l = Selection.Range.Characters.Count
        Selection.Range.InsertBefore "<key>"
        Selection.Collapse Direction:=wdCollapseStart
        Selection.MoveRight unit:=wdCharacter, Count:=Len("<key>"), Extend:=wdExtend
        Selection.Range.Case = wdLowerCase
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.MoveRight unit:=wdCharacter, Count:=key_l
        Selection.Range.InsertAfter "</key>"
        Selection.MoveRight unit:=wdCharacter, Count:=Len("</key>"), Extend:=wdExtend
        Selection.Range.Case = wdLowerCase
        Selection.Collapse Direction:=wdCollapseEnd
        mark_key = True
        marked = True
        Exit Function
    ElseIf Selection.Range.End >= ActiveDocument.Range.StoryLength - 1 Then
        mark_key = False
        Exit Function
    End If
    ' next step.
    Selection.Collapse Direction:=wdCollapseEnd
    pos = find_paragraph()

    ' ---- Loop checker
    iLoop = iLoop + 1
    If iLoop Mod 50 = 0 Then
        If MsgBox("Do you want to continue the loop in mark_key?", vbYesNo, "Debugging") = vbNo Then
            Exit Do
        End If
    End If
    '----***----
Loop
mark_key = False
'selection.MoveRight unit:=wdCharacter, Count:=Len("</key>") + 1
End Function

Function mark_selection(ByRef m_sel As Selection, ByVal m_open_tag As String, ByVal m_close_tag As String) As Boolean
If m_sel.Text = "" Then
    mark_selection = False
    Exit Function
End If
        Dim key_l As Integer
        key_l = Selection.Range.Characters.Count
        m_sel.Range.InsertBefore m_open_tag
        m_sel.Collapse Direction:=wdCollapseStart
        m_sel.MoveRight unit:=wdCharacter, Count:=Len(m_open_tag), Extend:=wdExtend
        m_sel.Range.Case = wdLowerCase
        m_sel.Collapse Direction:=wdCollapseEnd
        m_sel.MoveRight unit:=wdCharacter, Count:=key_l
        m_sel.Range.InsertAfter m_close_tag
        m_sel.MoveRight unit:=wdCharacter, Count:=Len(m_close_tag), Extend:=wdExtend
        m_sel.Range.Case = wdLowerCase
        m_sel.Collapse Direction:=wdCollapseEnd
        mark_selection = True
End Function
Sub MarkingTesting()
r = find_and_mark_term()

MsgBox (r)
'Call select_uppercased_text(Selection)
End Sub
Function find_and_mark_term() As Boolean
'
'Discription
'Finds the article term and marks it with <key> </key> tag.
'functionality similar to mark_key except of the format of the term which is just CAPITAL not bold
' modified: 05/02/2013
Dim marked As Boolean
Dim term As String
marked = False
' start searching loop untill the end of the document
counter = 0
Do While Not marked And is_end(Selection) = False
    'closer to potential term
    Selection.Collapse Direction:=wdCollapseEnd
    pos = find_paragraph()
    If is_end(Selection) Then
        marked = False
        Exit Do
    End If
    pos = find_any_letter(True, Selection)
    term = select_nonlowercased_text(Selection)
    term = trim_term(Selection)
    If term = "" Or Len(term) <= 1 Then
        ' attemtion failed; end of loop
        marked = False
    Else
        ' key term has been found
        ' remove extreme nonletter symbols
        ' term = trim_nonletter(Selection)
         'marking & forcing tags to lower case
        marked = mark_selection(Selection, "<key>", "</key>")
        ' loop checker
        counter = loop_checker(counter, 100, "find_and_mark_term")
    End If
Loop
' end searching
find_and_mark_term = marked

End Function
Function find_any_letter(ByVal m_direction As Boolean, ByRef m_sel As Selection) As Integer
' looks for any letter starting from m_sel selection to right or left
' returns 0 if anything is found and -1 if searching completed successfully
'
'
' depending on the search direction collapsing selection
If m_direction = True Then
    m_sel.Collapse Direction:=wdCollapseEnd
Else
    m_sel.Collapse Direction:=wdCollapseStart
End If
' start searching
m_sel.find.ClearFormatting
    With m_sel.find
        .Text = "^$"
        .Replacement.Text = ""
        .Forward = m_direction
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    m_sel.find.Execute
' result
find_any_letter = Selection.find.found
    
End Function
Function trim_nonletter(ByRef m_sel As Selection) As String
' removes last nonletter characters
' created: 05/02/2013
Dim term_start As Long
term_start = m_sel.Start
Do While case_checker(m_sel.Range.Characters.Last) = -1 And m_sel.Start >= term_start
    m_sel.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
Loop
trim_nonletter = Selection.Text
End Function
Function trim_term(ByRef m_sel As Selection) As String
' 05/02/2013
' remove unnessary last characters to select the correct term
Dim term As String
Dim next_c As Range
Set next_c = ActiveDocument.Range(1, 1)
term = trim_nonbold(m_sel)
term = trim_nonletter(m_sel)
next_c.SetRange Start:=m_sel.End, End:=m_sel.End + 1
If AscW(next_c.Text) = 187 Then
    m_sel.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
End If
trim_term = Selection.Text
End Function

Function trim_nonbold(ByRef m_sel As Selection) As String
' removes last nonbold characters
' 05/02/2013
Dim last_c As Range
Dim term_start As Long
Set last_c = ActiveDocument.Range(1, 1)
last_c.SetRange Start:=Selection.Range.End - 1, End:=Selection.Range.End
term_start = m_sel.Start
Do While last_c.font.Bold = False And last_c.Start >= term_start
    m_sel.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    last_c.SetRange Start:=Selection.Range.End - 1, End:=Selection.Range.End
Loop
trim_nonbold = Selection.Text
End Function

Function select_uppercased_text(m_current As Selection) As Integer
'****************************************
'Discription:
'Select the upper case part of the text beginning from the current selection position
'
If Selection.Range.Text = "" Then
    'If the range is emty select the next char
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
ElseIf Len(Selection.Range.Text) > 1 Then
    'if selection range contains more than 1 symbol decrease selection.
    Selection.Collapse Direction:=wdCollapseStart
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
ElseIf is_alpha(Selection.Range.Text) = False Then
    'if the range is not alphabetic char then exit.
    select_uppercased_text = 0
    Exit Function
End If
'Start loop
counter = 0
Do While case_checker(Selection.Range.Characters.Last) = 1
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
    counter = loop_checker(counter, 50, "select_uppercased_text")
Loop

End Function
Function select_nonlowercased_text(m_current As Selection) As String
'****************************************
'Discription:
'Select the upper case part of the text beginning from the current selection position
'
'pre checkings
If is_end(Selection) Then
    ' document's reached the end, nothing to select
    select_nonlowercased_text = ""
    Exit Function
ElseIf Len(Selection.Range.Text) > 1 Then
    'If the range is emty select the next char
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
ElseIf Len(Selection.Range.Text) > 1 Then
    'if selection range contains more than 1 symbol decrease selection.
    Selection.Collapse Direction:=wdCollapseStart
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
ElseIf is_alpha(Selection.Range.Text) = False Then
    'if the range is not alphabetic char then exit.
    select_nonlowercased_text = ""
    Exit Function
End If
'Start loop. Select text while character is not lower
i = 0
Do While case_checker(Selection.Range.Characters.Last) <> 0
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
    ' to prevent dead loop
    i = loop_checker(i, 100, "select_nonlowercased_text")
Loop
' Exclude last symbol, because it is lower cased
Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
select_nonlowercased_text = Selection.Text
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
Function is_upper(ByVal sym As String) As Boolean
' if empty
If sym = "" Then
    is_upper = False
    Exit Function
End If
' initialise en uppercased letters
en_alpha_str = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Dim kg_unicode(1 To 6) As Integer
Dim ru_unicode(1 To 33) As Integer
'initialize kg uppercased letters
kg_unicode(1) = 1186 ' kyrgyz Н
kg_unicode(2) = 1198 ' kyrgyz У
kg_unicode(3) = 1256 ' kyrgyz О

'initialise ru uppercased letters
For i = 1 To 32
    ru_unicode(i) = 1039 + i
Next
    ru_unicode(33) = 1025 ' big [yeo] Ё
    
' Check if the sym belongs to any group of above mentioned letter groups
'
' kg checking
sym_asc = AscW(sym)
For i = 1 To 3
    If AscW(sym) = kg_unicode(i) Then
        is_upper = True
        Exit Function
    End If
Next
' ru checking
For i = 1 To 33
    If AscW(sym) = ru_unicode(i) Then
        is_upper = True
        Exit Function
    End If
Next
' en checking
If InStr(ru_alpha_str, sym) Or InStr(en_alpha_str, sym) Then
    is_upper = True
    Exit Function
Else
    is_upper = False
    Exit Function
End If

End Function
Function case_checker(ByVal m_char As String) As Integer
' Returning values:
' -1    if the m_char is not a letter
' 0     if the m_char is a lower case letter
' 1     if the m_char is an upper case letter
If is_alpha(m_char) = False Then
' check if the sym isn't a letter
    case_checker = -1
    Exit Function
ElseIf is_upper(m_char) Then
' check if the sym is a capital letter
    case_checker = 1
    Exit Function
Else
' other way sym is a small letter
    case_checker = 0
End If
End Function

Function pick_term() As Boolean

'****************************************
'Discription
'
'
Dim marked As Boolean
marked = False

Selection.Collapse Direction:=wdCollapseEnd
pos = find(". ")
Do While Not marked And Selection.Range.End <= ActiveDocument.Range.StoryLength - 2
    
    Selection.Collapse Direction:=wdCollapseEnd
    'pos = find_any_letter()
    Selection.MoveRight unit:=wdCharacter, Count:=2, Extend:=wdExtend
    sel_str = select_term(Selection)
    Dim tmp_range As Range
    Set tmp_range = ActiveDocument.Range(Selection.End, Selection.End + 1)
    If sel_str <> "" And Not is_alpha(tmp_range.Text) And Selection.Tables.Count = 0 Then
        '********** Marking***************
        Dim key_l As Integer
        key_l = Selection.Range.Characters.Count
        Selection.Range.InsertParagraphBefore
        
        
        Selection.Collapse Direction:=wdCollapseEnd
        pick_term = True
        marked = True
        Exit Function
    ElseIf Selection.Range.End >= ActiveDocument.Range.StoryLength - 1 Then
        pick_term = False
        Exit Function
    End If
    ' next step.
    Selection.Collapse Direction:=wdCollapseEnd
    pos = find(". ")
    marked = Not pos

    ' ---- Loop checker
    iLoop = iLoop + 1
    If iLoop Mod 50 = 0 Then
        If MsgBox("Do you want to continue the loop in pick_term?", vbYesNo, "Debugging") = vbNo Then
            Exit Do
        End If
    End If
    '----***----
Loop
pick_term = False
'selection.MoveRight unit:=wdCharacter, Count:=Len("</key>") + 1
End Function

Function mark_key0() As Boolean

'****************************************
'Discription
'Finds and marks <key> </key> tag.
Dim marked As Boolean
marked = False
Do While Not marked
    Selection.Collapse Direction:=wdCollapseEnd
    pos = find_paragraph()
    Selection.Collapse Direction:=wdCollapseEnd
    pos = find_any_letter(True, Selection)
    sel_str = select_uppercase_bold(Selection)
    Dim tmp_range As Range
    Set tmp_range = ActiveDocument.Range(Selection.End, Selection.End + 1)
    If sel_str <> "" And Not is_alpha(tmp_range.Text) And Selection.Tables.Count = 0 Then
        '********** Marking***************
        Dim key_l As Integer
        key_l = Selection.Range.Characters.Count
        Selection.Range.InsertBefore "<key>"
        Selection.Collapse Direction:=wdCollapseStart
        Selection.MoveRight unit:=wdCharacter, Count:=Len("<key>"), Extend:=wdExtend
        Selection.Range.Case = wdLowerCase
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.MoveRight unit:=wdCharacter, Count:=key_l
        Selection.Range.InsertAfter "</key>"
        Selection.MoveRight unit:=wdCharacter, Count:=Len("</key>"), Extend:=wdExtend
        Selection.Range.Case = wdLowerCase
        Selection.Collapse Direction:=wdCollapseEnd
        mark_key0 = True
        marked = True
        Exit Function
    ElseIf Selection.Range.End >= ActiveDocument.Range.StoryLength - 1 Then
        mark_key0 = False
        Exit Function
    End If
    ' ---- Loop checker
    iLoop = iLoop + 1
    If iLoop Mod 50 = 0 Then
        If MsgBox("Do you want to continue the loop in mark_key?", vbYesNo, "Debugging") = vbNo Then
            Exit Do
        End If
    End If
    '----***----
Loop
Selection.MoveRight unit:=wdCharacter, Count:=Len("</key>") + 1
End Function
Function mark_key2() As Boolean

'****************************************
'Discription
'Finds and marks <key> </key> tag.
'This is second version, created to speed up previous.
Dim marked As Boolean
marked = False
Dim sel_str As String
Do While Not marked
e:  Selection.Collapse Direction:=wdCollapseEnd
    pos = find_paragraph()
    Selection.Collapse Direction:=wdCollapseEnd
    pos = find_any_letter(True, Selection)
    If Selection.Range.Bold = False Then
        If Selection.Range.End >= ActiveDocument.Range.StoryLength - 1 Then
            mark_key2 = False
            Exit Function
        Else
        'MsgBox ActiveDocument.Range.StoryLength
        GoTo e
        
        End If
    End If
    
    sel_str = select_uppercase_bold(Selection)
    '-------------------*select nonletter chars.
    Dim next_char As Range
    Set next_char = ActiveDocument.Range(Selection.End, Selection.End + 1)
    If next_char.Text <> " " And is_alpha(next_char.Text) = False Then
        If Selection.Range.End >= ActiveDocument.Range.StoryLength - 1 Then
            mark_key2 = False
            Exit Function
        End If
        Selection.Collapse Direction:=wdCollapseEnd
        GoTo e
    End If
    
    
    If sel_str <> "" And Not is_alpha(next_char.Text) And Selection.Tables.Count = 0 Then
        '********** Marking***************
        Dim key_l As Integer
        key_l = Selection.Range.Characters.Count
        Selection.Range.InsertBefore "<key>"
        Selection.Collapse Direction:=wdCollapseStart
        Selection.MoveRight unit:=wdCharacter, Count:=Len("<key>"), Extend:=wdExtend
        Selection.Range.Case = wdLowerCase
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.MoveRight unit:=wdCharacter, Count:=key_l
        Selection.Range.InsertAfter "</key>"
        Selection.MoveRight unit:=wdCharacter, Count:=Len("</key>"), Extend:=wdExtend
        Selection.Range.Case = wdLowerCase
        Selection.Collapse Direction:=wdCollapseEnd
        mark_key2 = True
        marked = True
        Exit Function
    ElseIf Selection.Range.End >= ActiveDocument.Range.StoryLength - 1 Then
        mark_key2 = False
        Exit Function
    End If
    ' ---- Loop checker
    iLoop = iLoop + 1
    If iLoop Mod 50 = 0 Then
        If MsgBox("Do you want to continue the loop in mark_key2?", vbYesNo, "Debugging") = vbNo Then
            Exit Do
        End If
    End If
    '----***----
Loop
Selection.MoveRight unit:=wdCharacter, Count:=Len("</key>") + 1
End Function
Function find_paragraph() As Integer
Selection.find.ClearFormatting
    With Selection.find
        .Text = "^p"
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
    ' <zaglushka>
    find_paragraph = 1
    ' </zaglushka>
End Function
Function test_object_return() As Range
Dim obj As Range
Set obj = Selection.Range
Selection.MoveRight unit:=wdCharacter, Count:=5, Extend:=wdExtend
Set object_return = obj
End Function
Function find0(ByVal m_what As String, Optional m_forward As Boolean = True, Optional m_wildcards As Boolean = False) As Range
'Function  symplifies searching
'Starts seaching from current position
'
'
Dim r_val As Range
'initializing
Selection.find.ClearFormatting
    With Selection.find
        .Text = m_what
        .Replacement.Text = ""
        .Forward = m_forward
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = m_wildcards
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
'executing
Selection.find.Execute
'return value
Set r_val = Selection.Range
Set find0 = r_val
End Function
Function find(ByVal m_what As String, Optional m_forward As Boolean = True, Optional m_wildcards As Boolean = False, Optional m_font As font = Nothing) As Range
'Function  symplifies searching
'Starts seaching from cursor position
'
'
Dim r_val As Range
'Selection.find.ClearFormatting
If Not (m_font Is Nothing) Then
   Selection.find.font = m_font
End If
    With Selection.find
        .Text = m_what
        .Replacement.Text = ""
        .Forward = m_forward
        .Wrap = wdFindStop
        .MatchWildcards = m_wildcards
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'Selection.find.ClearFormatting
    'Selection.find.font = m_font
    MsgBox (Selection.find.font.Superscript)
    Selection.find.Execute ' cannot find
If Selection.find.found Then
    Set r_val = Selection.Range
    Set find = r_val
Else
    Set find = Nothing
End If

End Function
Function find_txt(ByVal str As String) As Integer
Selection.find.ClearFormatting
    With Selection.find
        .Text = str
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
    ' <zaglushka>
    find_txt = Selection.find.found
    
    ' </zaglushka>
End Function
Function replace(ByVal src As String, ByVal trg As String) As Integer
Selection.find.ClearFormatting
    With Selection.find
        .Text = src
        .Replacement.Text = trg     ' cannot replace the quote
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'selection.find.Execute
    Selection.find.Execute
    If (Selection.find.found) Then
        Selection.Text = trg ' manually replacement needed
    End If
    ' <zaglushka>
    replace = Selection.find.found
    ' </zaglushka>
End Function
Sub Replace_all()
Dim src, trg As String
Dim q As Integer
q = 1
'src = InputBox("Find: ", Replacement)
'trg = InputBox("Replace with: ", Replacement)
src = ChrW(8217)
trg = ChrW(39)
Do While q
    q = replace(src, trg)
Loop


End Sub
Sub every_char_to_line()
Do
    Selection.MoveRight unit:=wdCharacter, Count:=2
    Selection.InsertAfter vbNewLine
Loop While Selection.Range.End < ActiveDocument.Range.StoryLength - 10
End Sub
Sub write_symbol_chars_to_begin()
Dim pos As Range
Dim lc, i As Integer
lc = 0
Do
    i = find_symbol_auto(True, Selection)
    Set pos = Selection.Range
    Selection.HomeKey wdStory
    Selection.InsertBefore (pos.Text)
    ' how to shift selection to certain position
    pos.Select
    Selection.MoveRight wdCharacter, 1
    Selection.Collapse wdCollapseEnd
    lc = loop_checker(lc, 1000, "writ_symbol_chars_to_begin")
Loop While i <> 0

End Sub
Function find_symbol_auto(ByVal m_direction As Boolean, ByRef m_sel As Selection) As Integer
' looks for any letter starting from m_sel selection to right or left
' returns unicode of found character if found and 0 if nothing was found
'
'
' depending on the search direction collapsing selection
If m_direction = True Then
    m_sel.Collapse Direction:=wdCollapseEnd
Else
    m_sel.Collapse Direction:=wdCollapseStart
End If
' start searching
m_sel.find.ClearFormatting
    With m_sel.find
        .Text = "^?"
        .Replacement.Text = ""
        .Forward = m_direction
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .font.Name = "Symbol"
        .font.ColorIndex = wdAuto
    End With
    m_sel.find.Execute
' result
If Selection.find.found = True Then
    find_symbol_auto = AscW(Selection.find.Text)
Else
    find_symbol_auto = 0
End If
    
End Function
        
Function select_uppercase_bold0(ByRef sel As Selection) As String
If Selection.Range.Case <> wdUpperCase Or Not Selection.Range.Bold Then
    select_uppercase_bold0 = ""
    Exit Function
End If
Do While Selection.Range.Case = wdUpperCase And Selection.Range.Bold
    'selection.Collapse direction:=wdCollapseEnd
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop in select_uppercase_bold", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop
Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend ' Critical point
iLoop = 0
off = Selection.Start
If Selection.Characters.Last.Case = wdUpperCase Then
    iLoop = iLoop + 1
End If
Do Until is_alpha(Selection.Characters.Last.Text) And Selection.Characters.Last.Bold And Selection.Characters.Last.Case = wdUpperCase
    If off > Selection.Start Then
        Exit Do
    End If
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop select_uppercase_bold", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop

select_uppercase_bold0 = Selection.Text
End Function

Function select_uppercase_bold(ByRef sel As Selection) As String
If Selection.Range.Case <> wdUpperCase Or Not Selection.Range.Bold Then
    select_uppercase_bold = ""
    Exit Function
End If
Do While Selection.Range.Case = wdUpperCase And Selection.Range.Bold
    'selection.Collapse direction:=wdCollapseEnd
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop in select_uppercase_bold", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop
Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend ' Critical point
iLoop = 0
off = Selection.Start
If Selection.Characters.Last.Case = wdUpperCase Then
    iLoop = iLoop + 1
End If
Do Until is_alpha(Selection.Characters.Last.Text) And Selection.Characters.Last.Bold And Selection.Characters.Last.Case = wdUpperCase
    If off > Selection.Start Then
        Exit Do
    End If
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop select_uppercase_bold", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop

select_uppercase_bold = Selection.Text
End Function

Function select_term(ByRef sel As Selection) As String
'
' This function created on base of select_uppercase_bold one, something like inheriting
' 29.11.2011 8:45

If Selection.Range.Case <> wdUpperCase Or Not Selection.Range.Bold Then
    select_term = ""
    Exit Function
End If
Do While Selection.Range.Case = wdUpperCase And Selection.Range.Bold
    'selection.Collapse direction:=wdCollapseEnd
    Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop in select_term", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop
Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend ' Critical point
iLoop = 0
off = Selection.Start
If Selection.Characters.Last.Case = wdUpperCase Then
    iLoop = iLoop + 1
End If
Do Until is_term_char(Selection.Characters.Last)
    
    If off > Selection.Start Then
        Exit Do
    End If
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop select_term", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop

select_term = Selection.Text
End Function


Public Function is_alpha_kk0(sym As String) As Boolean
ru_alpha_str = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧЩШЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчщшъыьэюя"
Dim kg_alpha_unicode(1 To 6) As Long
kg_alpha_unicode(1) = 1186 ' kyrgyz Н
kg_alpha_unicode(2) = 1198 ' kyrgyz у
kg_alpha_unicode(3) = 1256 ' kyrgyz О
kg_alpha_unicode(4) = 1187 ' kyrgyz н
kg_alpha_unicode(5) = 1199 ' kyrgyz у
kg_alpha_unicode(6) = 1257 ' kyrgyz о

' <bag> big kyrgyz characters seems to be missed
sym_asc = AscW(sym)
For i = 1 To 6
    If AscW(sym) = kg_alpha_unicode(i) Then
        is_alpha_kk0 = True
        Exit Function
    End If
Next
If InStr(ru_alpha_str, sym) Then
    is_alpha_kk0 = True
    Exit Function
Else
    is_alpha_kk0 = False
    Exit Function
End If

End Function
Public Function is_alpha_kk(sym As String) As Boolean
ru_alpha_str = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧЩШЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчщшъыьэюя"
Dim kg_alpha_unicode(1 To 6) As Long
' initialise kg letters
kg_alpha_unicode(1) = 1186 ' kyrgyz Н
kg_alpha_unicode(2) = 1198 ' kyrgyz у
kg_alpha_unicode(3) = 1256 ' kyrgyz О
kg_alpha_unicode(4) = 1187 ' kyrgyz н
kg_alpha_unicode(5) = 1199 ' kyrgyz у
kg_alpha_unicode(6) = 1257 ' kyrgyz о
sym_asc = AscW(sym)
For i = 1 To 6
    If AscW(sym) = kg_alpha_unicode(i) Then
        is_alpha_kk = True
        Exit Function
    End If
Next
If InStr(ru_alpha_str, sym) Then
    is_alpha_kk = True
    Exit Function
Else
    is_alpha_kk = False
    Exit Function
End If

End Function


Public Function is_alpha(sym As String) As Boolean
If sym = "" Then
    is_alpha = False
    Exit Function
End If
' this line were used in previos version of function, but it doesn't work properly on different platforms depending on locale settings
' ru_alpha_str = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧЩШЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчщшъыьэюя"
'
en_alpha_str = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz*"  'distortion of *
Dim kg_unicode(1 To 6) As Integer
Dim ru_unicode(1 To 66) As Integer
'initialize kg letters
kg_unicode(1) = 1186 ' kyrgyz Н
kg_unicode(2) = 1198 ' kyrgyz у
kg_unicode(3) = 1256 ' kyrgyz О
kg_unicode(4) = 1187 ' kyrgyz н
kg_unicode(5) = 1199 ' kyrgyz у
kg_unicode(6) = 1257 ' kyrgyz о

'initialise ru letters
For i = 1 To 64
    ru_unicode(i) = 1039 + i
Next
    ru_unicode(65) = 1025 ' big [yeo] Ё
    ru_unicode(66) = 1105 ' small [yeo] ё
    
' check if the sym belongs to any group of above mentioned letter groups
' kg checking
sym_asc = AscW(sym)
For i = 1 To 6
    If AscW(sym) = kg_unicode(i) Then
        is_alpha = True
        Exit Function
    End If
Next
' ru checking
For i = 1 To 66
    If AscW(sym) = ru_unicode(i) Then
        is_alpha = True
        Exit Function
    End If
Next
' en checking
If InStr(ru_alpha_str, sym) Or InStr(en_alpha_str, sym) Then
    is_alpha = True
    Exit Function
Else
    is_alpha = False
    Exit Function
End If

End Function
Public Function is_alpha0(sym As String) As Boolean
If sym = "" Then
    is_alpha0 = False
    Exit Function
End If
ru_alpha_str = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧЩШЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчщшъыьэюя"
en_alpha_str = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Dim kg_alpha_unicode(1 To 6) As Long
kg_alpha_unicode(1) = 1186 ' kyrgyz Н
kg_alpha_unicode(2) = 1198 ' kyrgyz у
kg_alpha_unicode(3) = 1256 ' kyrgyz О
kg_alpha_unicode(4) = 1187 ' kyrgyz н
kg_alpha_unicode(5) = 1199 ' kyrgyz у
kg_alpha_unicode(6) = 1257 ' kyrgyz о
sym_asc = AscW(sym)
For i = 1 To 6
    If AscW(sym) = kg_alpha_unicode(i) Then
        is_alpha0 = True
        Exit Function
    End If
Next
If InStr(ru_alpha_str, sym) Or InStr(en_alpha_str, sym) Then
    is_alpha0 = True
    Exit Function
Else
    is_alpha0 = False
    Exit Function
End If

End Function
' <bad>
Public Function is_punct(char As Characters) As Boolean
Dim punct_str As String
punct_str = ".,-"
space_str = "^t"

End Function
' </bad>
'******************************************<definition> </definition>******************************
Function mark_key_and_def_in_article_tag() As Boolean
b = select_article_tag()
If b Then
    b = mark_key_in_article_tag(Selection.Start, Selection.End)
    Selection.Collapse Direction:=wdCollapseEnd
    'selection.Find.Text = "</key>"
    'selection.Find.Execute
    Selection.Range.InsertAfter "<![CDATA["
    Selection.Range.InsertParagraphAfter
    Selection.Range.InsertAfter "<definition type='h'>"
    Selection.Range.InsertParagraphAfter
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.find.Text = "</article>"
    Selection.find.Execute
    Selection.Range.InsertParagraphBefore
    Selection.Range.InsertBefore "</definition>"
    Selection.Range.InsertParagraphBefore
    Selection.Range.InsertBefore "]]>"
    Selection.Collapse Direction:=wdCollapseEnd
    mark_key_and_def_in_article_tag = True
Else
    mark_key_and_def_in_article_tag = False
End If
End Function
' </bad>
Function mark_key_and_def_in_article_tag0() As Boolean
b = select_article_tag()
If b Then
    b = mark_key_in_article_tag(Selection.Start, Selection.End)
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.find.Text = "</key>"
    Selection.find.Execute
    Selection.Range.InsertAfter "<![CDATA["
    Selection.Range.InsertParagraphAfter
    Selection.Range.InsertAfter "<definition type='h'>"
    Selection.Range.InsertParagraphAfter
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.find.Text = "</article>"
    Selection.find.Execute
    Selection.Range.InsertParagraphBefore
    Selection.Range.InsertBefore "</definition>"
    Selection.Range.InsertParagraphBefore
    Selection.Range.InsertBefore "]]>"
    Selection.Collapse Direction:=wdCollapseEnd
    mark_key_and_def_in_article_tag0 = True
Else
    mark_key_and_def_in_article_tag0 = False
End If
End Function
Sub ConvertMS_Script_To_MS_Position()
' Description: Converts script property to position property in MS Word
'
Selection.HomeKey unit:=wdStory
'* When we use our own font object searching doesn't give any yeild
Dim new_font As font
Set new_font = Selection.font.Duplicate
new_font.Superscript = True
'Selection.find.font.Superscript = True
Set rg = find("^?", , , new_font)

Selection.find.ClearFormatting
    With Selection.find
        .Text = "^?"
        .font.Superscript = True
        .Replacement.Text = "^&"
        .Replacement.font.Superscript = False
        .Replacement.font.Position = 3
        .Replacement.font.ColorIndex = wdDarkBlue
        .Forward = True
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Selection.find.Execute replace:=wdReplaceAll
    End With

Selection.find.ClearFormatting
    With Selection.find
        .Text = "^?"
        .font.Subscript = True
        .Replacement.Text = "^&"
        .Replacement.font.Subscript = False
        .Replacement.font.Position = -3
        .Replacement.font.ColorIndex = wdViolet
        .Forward = True
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Selection.find.Execute replace:=wdReplaceAll
    End With

End Sub
Sub Mark_Scripts()
' marks the scripts realaized via MS Position property
'
    reply = MsgBox("To proper work document shouldn't contain any fraction valued super/subscripts. Please fix it manually if any. To continue press 'Yes', to terminate and quit 'NO' ", vbYesNo, "Sup/subscript acknowledgement.")
    If reply = vbNo Then
        Exit Sub
    End If
    For i = 1 To 6
        r = Mark_Script(i, "sup")
    Next
    For i = -6 To -1
        r = Mark_Script(i, "sub")
    Next
End Sub
Function Mark_Script(ByVal m_position As Integer, ByVal m_tag_name As String) As Boolean
    Dim open_tag, close_tag As String
    open_tag = "<" + m_tag_name + ">"
    close_tag = "</" + m_tag_name + ">"
    Selection.HomeKey unit:=wdStory
    With Selection.find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^?"
        .Replacement.Text = open_tag + "^&" + close_tag
        .Replacement.font.Color = wdColorBlue
        .font.Position = m_position           ' cannot assign the fraction number
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        Selection.find.Execute replace:=wdReplaceAll
    End With
    Mark_Script = Selection.find.found
    
   ' post clearing
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = close_tag + open_tag
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute replace:=wdReplaceAll
    
End Function
Sub NormalizeSyntax()

Application.Run macroname:="RemoveDoubleSpaces"
Application.Run macroname:="NormalizeParagraphs" ' removes spaces before paragraphs
Application.Run macroname:="NormalizeCommas"     ' removes spaces before commas
'0.
'Application.Run macroname:="RemoveBigLetter"     'ToDo:  it should be developed
'0.1
'Application.Run macroname:="RemoveDoubleParagraph"
'1.
Application.Run macroname:="RemoveSoftHyphen"

End Sub
Sub Normalize_KG_Letters()

' Managing fonts...
Selection.HomeKey unit:=wdStory

reply = MsgBox("Do you want to change font from JanyzakArial to WinKK?", vbYesNo, font)
If reply = vbYes Then
    Call ConvertJanyzakArialToWinKK
End If

reply = MsgBox("Do you want to change font from EncyclopedyCenterUnknownFont to WinKK?", vbYesNo, font)
If reply = vbYes Then
    Call ConvertEncyclopediaCenterFontToWinKK1
End If

reply = MsgBox("Do you want to change font from Times_Q2 Font to WinKK?", vbYesNo, font)
If reply = vbYes Then
    Call ConvertTimes_Q2ToWinKK1
End If
End Sub
Sub ConvertSymbolToUTF8_BasicLatin()


Dim i As Long
Dim Symbol(33) As Long
Dim UTF8(33) As Long

Symbol(1) = 32        '   space
Symbol(2) = 33        '   exclam
Symbol(3) = 35        '   numbersign
Symbol(4) = 37        '   percent
Symbol(5) = 38        '   ampersand
Symbol(6) = 40        '   parenleft
Symbol(7) = 41        '   parenright
Symbol(8) = 43        '   plus
Symbol(9) = 44        '   comma
Symbol(10) = 44       '   period
Symbol(11) = 47       '   slash
Symbol(12) = 48       '   zero
Symbol(13) = 49       '   one
Symbol(14) = 50       '   two
Symbol(15) = 51       '   three
Symbol(16) = 52       '   four
Symbol(17) = 53       '   five
Symbol(18) = 54       '   six
Symbol(19) = 55       '   seven
Symbol(20) = 56       '   eight
Symbol(21) = 57       '   nine
Symbol(22) = 58       '   colon
Symbol(23) = 59       '   semicolon
Symbol(24) = 60       '   less
Symbol(25) = 61       '   equal
Symbol(26) = 62       '   greater
Symbol(27) = 63       '   question
Symbol(28) = 91       '   bracketleft
Symbol(29) = 93       '   bracketright
Symbol(30) = 95       '   underscore
Symbol(31) = 123       '  braceleft
Symbol(32) = 124       '  bar
Symbol(33) = 125       '  braceright

UTF8(1) = 32          '   Space
UTF8(2) = 33          '   Exclamation mark
UTF8(3) = 35          '   Number sign
UTF8(4) = 37          '   Percent sign
UTF8(5) = 38          '   Ampersand
UTF8(6) = 40          '   Left parenthesis
UTF8(7) = 41          '   Right parenthesis
UTF8(8) = 43          '   Plus sign
UTF8(9) = 44          '   Comma
UTF8(10) = 46         '   Full stop
UTF8(11) = 47         '   Solidus
UTF8(12) = 48         '   Digit zero
UTF8(13) = 49         '   Digit one
UTF8(14) = 50         '   Digit two
UTF8(15) = 51         '   Digit three
UTF8(16) = 52         '   Digit four
UTF8(17) = 53         '   Digit five
UTF8(18) = 54         '   Digit six
UTF8(19) = 55         '   Digit seven
UTF8(20) = 56         '   Digit eight
UTF8(21) = 57         '   Digit nine
UTF8(22) = 58         '   Colon
UTF8(23) = 59         '   Semicolon
UTF8(24) = 60         '   Less-than sign
UTF8(25) = 61         '   Equals sign
UTF8(26) = 62         '   Greater-than sign
UTF8(27) = 63         '   Question mark
UTF8(28) = 91         '   Left square bracket
UTF8(29) = 93         '   Right square bracket
UTF8(30) = 95         '   Low line
UTF8(31) = 123        '   Left curly bracket
UTF8(32) = 124        '   Vertical line
UTF8(33) = 125        '   Right curly bracket


i = 1
Do While i <= 33
    Selection.find.Text = ChrW$(Symbol(i))
    Selection.find.MatchCase = True
    Selection.find.Replacement.Text = ChrW$(UTF8(i))
    Selection.find.Replacement.font.ColorIndex = wdBlue
    Selection.find.Execute replace:=wdReplaceAll
    Selection.HomeKey unit:=wdStory

    i = i + 1
Loop

'MsgBox Selection.find.found

End Sub
Sub ConvertSymbolToUTF8_Math()


Dim i As Long
Dim Symbol(39) As Long
Dim UTF8(39) As Long

Symbol(1) = 34        '   universal
Symbol(2) = 36        '   existential
Symbol(3) = 39        '   suchthat
Symbol(4) = 42        '   asteriskmath
Symbol(5) = 45        '   minus
Symbol(6) = 64        '   congruent
Symbol(7) = 92        '   therefore
Symbol(8) = 94        '   perpendicular
Symbol(9) = 126        '  similar
Symbol(10) = 163       '  lessequal
Symbol(11) = 165       '  infinity
Symbol(12) = 179       '  greaterequal
Symbol(13) = 181       '  proportional
Symbol(14) = 182       '  partialdiff
Symbol(15) = 183       '  bullet
Symbol(16) = 185       '  notequal
Symbol(17) = 186       '  equivalence
Symbol(18) = 187       '  approxequal
Symbol(19) = 196       '  circlemultiply
Symbol(20) = 197       '  circleplus
Symbol(21) = 198       '  emptyset
Symbol(22) = 199       '  intersection
Symbol(23) = 200       '  union
Symbol(24) = 201       '  propersuperset
Symbol(25) = 202       '  reflexsuperset
Symbol(26) = 203       '  notsubset
Symbol(27) = 204       '  propersubset
Symbol(28) = 205       '  reflexsubset
Symbol(29) = 206       '  element
Symbol(30) = 207       '  notelement
Symbol(31) = 208       '  angle
Symbol(32) = 209       '  gradient
Symbol(33) = 213       '  product
Symbol(34) = 214       '  radical
Symbol(35) = 215       '  dotmath
Symbol(36) = 217       '  logicaland
Symbol(37) = 218       '  logicalor
Symbol(38) = 229       '  summation
Symbol(39) = 242       '  integral


UTF8(1) = 8704          ' For all
UTF8(2) = 8707          ' There exists
UTF8(3) = 8717          ' Small contains as member
UTF8(4) = 8727          ' Asterisk operator
UTF8(5) = 8722          ' Minus sign
UTF8(6) = 8773          ' Approximately equal to
UTF8(7) = 8756          ' Therefore
UTF8(8) = 8869          ' Up tack
UTF8(9) = 8764          ' Tilde operator
UTF8(10) = 8804         ' Less-than or equal to
UTF8(11) = 8734         ' Infinity
UTF8(12) = 8805         ' Greater-than or equal to
UTF8(13) = 8733         ' Proportional to
UTF8(14) = 8706         ' Partial differential
UTF8(15) = 8729         ' Bullet operator
UTF8(16) = 8800         ' Not equal to
UTF8(17) = 8801         ' Identical to
UTF8(18) = 8776         ' Almost equal to
UTF8(19) = 8855         ' Circled times
UTF8(20) = 8853         ' Circled plus
UTF8(21) = 8709         ' Empty set
UTF8(22) = 8745         ' Intersection
UTF8(23) = 8746         ' Union
UTF8(24) = 8835         ' Superset of
UTF8(25) = 8839         ' Superset of or equal to
UTF8(26) = 8836         ' Not a subset of
UTF8(27) = 8834         ' Subset of
UTF8(28) = 8838         ' Subset of or equal to
UTF8(29) = 8712         ' Element of
UTF8(30) = 8713         ' Not an element of
UTF8(31) = 8736         ' Angle
UTF8(32) = 8711         ' Nabla
UTF8(33) = 8719         ' N-ary product
UTF8(34) = 8730         ' Square root
UTF8(35) = 8901         ' Dot operator
UTF8(36) = 8743         ' Logical and
UTF8(37) = 8744         ' Logical or
UTF8(38) = 8721         ' N-ary summation
UTF8(39) = 8747         ' Integral



i = 1
Do While i <= 39
    Selection.find.Text = ChrW$(Symbol(i))
    Selection.find.MatchCase = True
    Selection.find.Replacement.Text = ChrW$(UTF8(i))
    Selection.find.Replacement.font.ColorIndex = wdBlue
    Selection.find.Execute replace:=wdReplaceAll
    Selection.HomeKey unit:=wdStory

    i = i + 1
Loop

'MsgBox Selection.find.found

End Sub
Sub ConvertSymbolToUTF8_Misc()

Dim i As Long
Dim Symbol(32) As Long
Dim UTF8(32) As Long

Symbol(1) = 171        '      arrowboth
Symbol(2) = 172        '      arrowleft
Symbol(3) = 173        '      arrowup
Symbol(4) = 174        '      arrowright
Symbol(5) = 175        '      arrowdown
Symbol(6) = 191        '      carriagereturn
Symbol(7) = 219        '      arrowdblboth
Symbol(8) = 220        '      arrowdblleft
Symbol(9) = 221        '      arrowdblup
Symbol(10) = 222       '      arrowdblright
Symbol(11) = 223       '      arrowdbldown
Symbol(12) = 240       '      Euro
Symbol(13) = 192       '      aleph
Symbol(14) = 193       '      Ifraktur
Symbol(15) = 194       '      Rfraktur
Symbol(16) = 195       '      weierstrass
Symbol(17) = 212       '      trademarkserif
Symbol(18) = 228       '      trademarksans
Symbol(19) = 162       '      minute
Symbol(20) = 164       '      fraction
Symbol(21) = 178       '      second
Symbol(22) = 188       '      ellipsis
Symbol(23) = 176       '      degree
Symbol(24) = 177       '      plusminus
Symbol(25) = 180       '      multiply
Symbol(26) = 184       '      divide
Symbol(27) = 210       '      registerserif
Symbol(28) = 211       '      copyrightserif
Symbol(29) = 216       '      logicalnot
Symbol(30) = 226       '      registersans
Symbol(31) = 227       '      copyrightsans
Symbol(32) = 166       '      florin

UTF8(1) = 8596          '     Left right arrow
UTF8(2) = 8592          '     Leftwards arrow
UTF8(3) = 8593          '     Upwards arrow
UTF8(4) = 8594          '     Rightwards arrow
UTF8(5) = 8595          '     Downwards arrow
UTF8(6) = 8629          '     Downwards arrow with corner leftwards
UTF8(7) = 8660          '     Left right double arrow
UTF8(8) = 8656          '     Leftwards double arrow
UTF8(9) = 8657          '     Upwards double arrow
UTF8(10) = 8658         '     Rightwards double arrow
UTF8(11) = 8659         '     Downwards double arrow
UTF8(12) = 8364         '     Euro sign
UTF8(13) = 8501         '     Alef symbol
UTF8(14) = 8465         '     Black-letter capital I
UTF8(15) = 8476         '     Black-letter capital R
UTF8(16) = 8472         '     Script capital P
UTF8(17) = 8482         '     Trade mark sign (serif)
UTF8(18) = 8482         '     Trade mark sign (sans-serif)
UTF8(19) = 8242         '     Prime
UTF8(20) = 8260         '     Fraction slash
UTF8(21) = 8243         '     Double prime
UTF8(22) = 8230         '     Horizontal ellipsis
UTF8(23) = 176         '      Degree sign
UTF8(24) = 177         '      Plus-minus sign
UTF8(25) = 215         '      Multiplication sign
UTF8(26) = 247         '      Division sign
UTF8(27) = 174         '      Registered sign (serif)
UTF8(28) = 169         '      Copyright sign (serif)
UTF8(29) = 172         '      Not sign
UTF8(30) = 174         '      Registered sign (sans-serif)
UTF8(31) = 169         '      Copyright sign (sans-serif)
UTF8(32) = 402         '      Latin small letter f with hook




i = 1
Do While i <= 32
    Selection.find.Text = ChrW$(Symbol(i))
    Selection.find.MatchCase = True
    Selection.find.Replacement.Text = ChrW$(UTF8(i))
    Selection.find.Replacement.font.ColorIndex = wdBlue
    Selection.find.Execute replace:=wdReplaceAll
    Selection.HomeKey unit:=wdStory

    i = i + 1
Loop

'MsgBox Selection.find.found

End Sub
Sub ConvertSymbolGreekToUnicode()
'
'
' This function replaces Symbol font Greek characters to Unicode equivalents
'
'
Dim found As Boolean
Dim i As Long
Dim Symbol(53) As Long
Symbol(1) = 61505        '    Alpha
Symbol(2) = 61506        '    Beta
Symbol(3) = 61507        '    Chi
Symbol(4) = 61508        '    Delta
Symbol(5) = 61509        '    Epsilon
Symbol(6) = 61510        '    Phi
Symbol(7) = 61511        '    Gamma
Symbol(8) = 61512        '    Eta
Symbol(9) = 61513        '    Iota
Symbol(10) = 61514       '    theta1
Symbol(11) = 61515       '    Kappa
Symbol(12) = 61516       '    Lambda
Symbol(13) = 61517       '    Mu
Symbol(14) = 61518       '    Nu
Symbol(15) = 61519       '    Omicron
Symbol(16) = 61520       '    Pi
Symbol(17) = 61521       '    Theta
Symbol(18) = 61522       '    Rho
Symbol(19) = 61523       '    Sigma
Symbol(20) = 61524       '    Tau
Symbol(21) = 61525       '    Upsilon
Symbol(22) = 61526       '    sigma1
Symbol(23) = 61527       '    Omega
Symbol(24) = 61528       '    Xi
Symbol(25) = 61529       '    Psi
Symbol(26) = 61530       '    Zeta
Symbol(27) = 61537       '    alpha
Symbol(28) = 61538       '    beta
Symbol(29) = 61539       '    chi
Symbol(30) = 61540       '    delta
Symbol(31) = 61541       '    epsilon
Symbol(32) = 61542       '    phi
Symbol(33) = 61543       '    gamma
Symbol(34) = 61544       '    eta
Symbol(35) = 61545       '    iota
Symbol(36) = 61546       '    phi1
Symbol(37) = 61547       '    kappa
Symbol(38) = 61548       '    lambda
Symbol(39) = 61549       '    mu
Symbol(40) = 61550       '    nu
Symbol(41) = 61551       '    omicron
Symbol(42) = 61552       '    pi
Symbol(43) = 61553       '    theta
Symbol(44) = 61554       '    rho
Symbol(45) = 61555       '    sigma
Symbol(46) = 61556       '    tau
Symbol(47) = 61557       '    upsilon
Symbol(48) = 61558       '    omega1
Symbol(49) = 61559       '    omega
Symbol(50) = 61560       '    xi
Symbol(51) = 61561       '    psi
Symbol(52) = 61562       '    zeta
Symbol(53) = 61601       '    Upsilon1


Dim UTF8(53) As Long
UTF8(1) = 913          '  Alpha
UTF8(2) = 914          '  Beta
UTF8(3) = 935          '  Chi
UTF8(4) = 916          '  Delta
UTF8(5) = 917          '  Epsilon
UTF8(6) = 934          '  Phi
UTF8(7) = 915          '  Gamma
UTF8(8) = 919          '  Eta
UTF8(9) = 921          '  Iota
UTF8(10) = 977         '  theta1
UTF8(11) = 922         '  Kappa
UTF8(12) = 923         '  Lambda
UTF8(13) = 924         '  Mu
UTF8(14) = 925         '  Nu
UTF8(15) = 927         '  Omicron
UTF8(16) = 928         '  Pi
UTF8(17) = 920         '  Theta
UTF8(18) = 929         '  Rho
UTF8(19) = 931         '  Sigma
UTF8(20) = 932         '  Tau
UTF8(21) = 933         '  Upsilon
UTF8(22) = 962         '  sigma1
UTF8(23) = 937         '  Omega
UTF8(24) = 926         '  Xi
UTF8(25) = 936         '  Psi
UTF8(26) = 918         '  Zeta
UTF8(27) = 945         '  alpha
UTF8(28) = 946         '  beta
UTF8(29) = 967         '  chi
UTF8(30) = 948         '  delta
UTF8(31) = 949         '  epsilon
UTF8(32) = 966         '  phi
UTF8(33) = 947         '  gamma
UTF8(34) = 951         '  eta
UTF8(35) = 953         '  iota
UTF8(36) = 981         '  phi1
UTF8(37) = 954         '  kappa
UTF8(38) = 955         '  lambda
UTF8(39) = 956         '  mu
UTF8(40) = 957         '  nu
UTF8(41) = 959         '  omicron
UTF8(42) = 960         '  pi
UTF8(43) = 952         '  theta
UTF8(44) = 961         '  rho
UTF8(45) = 963         '  sigma
UTF8(46) = 964         '  tau
UTF8(47) = 965         '  upsilon
UTF8(48) = 982         '  omega1
UTF8(49) = 969         '  omega
UTF8(50) = 958         '  xi
UTF8(51) = 968         '  psi
UTF8(52) = 950         '  zeta
UTF8(53) = 978         '  Upsilon1


i = 1
Do While i <= 53
    Selection.find.Text = ChrW$(Symbol(i))
    Selection.find.MatchCase = True
    Selection.find.Replacement.Text = ChrW$(UTF8(i))
    Selection.find.Replacement.font.ColorIndex = wdBlue
    
    Selection.find.Execute replace:=wdReplaceAll
     
    
    Selection.HomeKey unit:=wdStory

    i = i + 1
Loop

'MsgBox Selection.find.found
'Tested.


End Sub
Sub ConvertSymbolPhenomenaToUnicode()
'
' This function replaces Symbol font chars to Unicode equivalents
'
Dim found As Boolean
Dim i As Long
Dim Symbol(59) As Long
Dim UTF8(59) As Long
Symbol(1) = 61472
Symbol(2) = 61485
Symbol(3) = 61619
Symbol(4) = 61566
Symbol(5) = 61655
Symbol(6) = 61485
Symbol(7) = 61487
Symbol(8) = 61501
Symbol(9) = 61630
Symbol(10) = 61616
Symbol(11) = 61602
Symbol(12) = 61472
Symbol(13) = 61625
Symbol(14) = 61484
Symbol(15) = 61627
Symbol(16) = 61481
Symbol(17) = 61614
Symbol(18) = 61617
Symbol(19) = 61488
Symbol(20) = 61489
Symbol(21) = 61500
Symbol(22) = 61502
Symbol(23) = 61533
Symbol(24) = 61531
Symbol(25) = 61491
Symbol(26) = 61483
Symbol(27) = 61490
Symbol(28) = 61480
Symbol(29) = 61603
Symbol(30) = 61620
Symbol(31) = 61534
Symbol(32) = 61618
Symbol(33) = 61682
Symbol(34) = 61604
Symbol(35) = 61486
Symbol(36) = 61605
Symbol(37) = 61482
Symbol(38) = 61606
Symbol(39) = 61498
Symbol(40) = 61493
Symbol(41) = 61495
Symbol(42) = 61496
Symbol(43) = 61499
Symbol(44) = 61492
Symbol(45) = 61564
Symbol(46) = 61504
Symbol(47) = 61629
Symbol(48) = 61648
Symbol(49) = 61622
Symbol(50) = 61649
Symbol(51) = 61587
Symbol(52) = 61669
Symbol(53) = 61624
Symbol(54) = 61628
Symbol(55) = 61657
Symbol(56) = 61632
Symbol(57) = 61477
Symbol(58) = 61611
Symbol(59) = 61671


UTF8(1) = 32            '  space
UTF8(2) = 8722          ' Minus sign
UTF8(3) = 8805          ' Greater-than or equal to
UTF8(4) = 8764          ' Tilde operator
UTF8(5) = 8901          ' Dot operator
UTF8(6) = 8722          ' Minus sign
UTF8(7) = 8260          '     Fraction slash
UTF8(8) = 61            '   Equals sign
UTF8(9) = 9135          '     Horizontal line extension
UTF8(10) = 176          '      Degree sign
UTF8(11) = 8242         '     Prime
UTF8(12) = 32           ' space
UTF8(13) = 8800         ' Not equal to
UTF8(14) = 44           '  is  ,
UTF8(15) = 8776         ' Almost equal to
UTF8(16) = 41           '   Right parenthesis
UTF8(17) = 8594         '     Rightwards arrow
UTF8(18) = 177         '      Plus-minus sign
UTF8(19) = 48         '   Digit zero
UTF8(20) = 49         '   Digit one
UTF8(21) = 60         '   Less-than sign
UTF8(22) = 62         '   Greater-than sign
UTF8(23) = 93         '   Right square bracket
UTF8(24) = 91         '   Left square bracket
UTF8(25) = 51         '   Digit three
UTF8(26) = 43         '   Plus sign
UTF8(27) = 50         '   Digit two
UTF8(28) = 40         '   Left parenthesis
UTF8(29) = 8804          ' LESS-THAN OR EQUAL TO
UTF8(30) = 215         '      Multiplication sign
UTF8(31) = 8869         '     UP TACK
UTF8(32) = 8243         '     Double prime
UTF8(33) = 8747         ' Integral
UTF8(34) = 8260         '     Fraction slash
UTF8(35) = 46         '   Full stop
UTF8(36) = 8734           'INFINITY
UTF8(37) = 8727           'ASTERISK OPERATOR
UTF8(38) = 402         '      Latin small letter f with hook
UTF8(39) = 58         '   Colon
UTF8(40) = 53         '   Digit five
UTF8(41) = 55         '   Digit five
UTF8(42) = 56         '   Digit five
UTF8(43) = 59         '   Semicolon
UTF8(44) = 52         '   Digit four
UTF8(45) = 9168         '     Vertical line extension
UTF8(46) = 8773           'APPROXIMATELY EQUAL TO
UTF8(47) = 8739           'divides
UTF8(48) = 8736           'ANGLE
UTF8(49) = 8706         ' Partial differential
UTF8(50) = 8711         ' Nabla
UTF8(51) = 8364         '     Euro sign
UTF8(52) = 8721         '     N-ary summation
UTF8(53) = 247         '      Division sign
UTF8(54) = 8230         '     Horizontal ellipsis
UTF8(55) = 8743         '     Logical and
UTF8(56) = 8501         '     Alef symbol
UTF8(57) = 37         '   Percent sign
UTF8(58) = 8596           'LEFT RIGHT ARROW
UTF8(59) = 9122           'LEFT SQUARE BRACKET EXTENSION



i = 1
Do While i <= 59
    Selection.find.Text = ChrW$(Symbol(i))
    Selection.find.MatchCase = True
    Selection.find.Replacement.Text = ChrW$(UTF8(i))
    Selection.find.Replacement.font.ColorIndex = wdBlue
    Selection.find.Execute replace:=wdReplaceAll
    Selection.HomeKey unit:=wdStory

    i = i + 1
Loop

'MsgBox Selection.find.found
'Tested.

End Sub
Sub main_proc2()

Call NormalizeSyntax
Call Normalize_KG_Letters
'Call ConvertToUnicode
Call MarkArticles
'Call Mark_Scripts
Call Insert_Header_Tag

End Sub

Function insert_img_tag() As Boolean
Selection.Collapse (wdCollapseStart)
Selection.find.ClearFormatting
    With Selection.find
        .Text = "[<]key[>]*[<][/]key[>]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute
    If Selection.find.found = False Then
        insert_img_tag = False
        Exit Function
    End If
    Dim m_key, tmp As String
    Dim beg, off As Integer
    tmp = Trim(Selection.Text)
    beg = Len("<key>") + 1
    off = Len(tmp) - Len("</key>") - Len("<key>")
    m_key = Trim(Mid(tmp, beg, off))
    
    Selection.Collapse (wdCollapseEnd)
    Selection.find.MatchWildcards = False
    Selection.find.Text = "]]>"
    Selection.find.Execute
    If Selection.find.found = False Then
        insert_img_tag = False
        Return
    End If
    Selection.InsertBefore ("<br><p><img src='" & m_key & ".jpg' /></p> ")
    Selection.Collapse (wdCollapseEnd)
    insert_img_tag = True
End Function
Public Sub insert_img_tag_all()
Dim flag As Boolean
flag = True
Do While flag
    flag = insert_img_tag()
    ' ---- Loop checker
        iLoop = iLoop + 1
        If iLoop Mod 50 = 0 Then
            If MsgBox("Do you want to continue the loop? func name: insert_img_tat_all", vbYesNo, "Debugging") = vbNo Then
                Exit Do
            End If
        End If
    '     ----***----
Loop
End Sub
Public Enum JanyzakArial
' Represents the symbols used in Janyzak Arial font
    soft_u = 1065
    soft_o = 1025
    ng = 1066
    soft_big_u = 1097
    soft_big_o = 1105
    big_ng = 1098
    
End Enum
Public Enum X_Arial
' Represents the unknown font the Encyclopedy Centre's used
    soft_u = 1199
    soft_o = 1257
    ng = 1187
    soft_big_u = 1198
    soft_big_o = 1256
    big_ng = 1186
    
End Enum
Public Enum Win_Khazah
' Represents the set of symbols used in WinXP on Kazakh(KK)keyboard
    soft_u = 1199
    soft_o = 1257
    ng = 1187
    soft_big_u = 1198
    soft_big_o = 1256
    big_ng = 1186
    
End Enum
Public Enum tmp
 a
 b
 c

End Enum

Public Sub ConvertJanyzakArialToWinKK()
'
'
' Replaces kyrgyz Janyzak Arial font characters to WinKK(Kazakh) ones
'
Dim found As Boolean
Dim i As Integer
Dim Janyzak(6) As Integer
Janyzak(1) = 1097 ' kg soft 'u'
Janyzak(2) = 1105 ' kg soft 'o'
Janyzak(3) = 1098 ' kg      'ng'
Janyzak(4) = 1065 ' kg soft 'U' (Upper Case)
Janyzak(5) = 1025  ' kg soft 'O' (Upper Case)
Janyzak(6) = 1066  ' kg      'NG'(Upper Case)

Dim WinKK(6) As Integer
WinKK(1) = 1199 ' kg soft 'u'
WinKK(2) = 1257 ' kg soft 'o'
WinKK(3) = 1187 ' kg      'ng'
WinKK(4) = 1198 ' kg soft 'U' (Upper Case)
WinKK(5) = 1256 ' kg soft 'O' (Upper Case)
WinKK(6) = 1186 ' kg      'NG'(Upper Case)

i = 1
Dim reply As Integer
Do While i <= 6
    Selection.find.Text = ChrW$(Janyzak(i))
    Selection.find.MatchCase = True
    Selection.find.Execute
    found = Selection.find.found
    Do While found
        Selection.Text = ChrW$(WinKK(i))
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.find.Execute
        found = Selection.find.found
        
        ' ---- Loop checker
            iLoop = iLoop + 1
            If iLoop Mod 500 = 0 Then
                If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
                    reply = vbNo
                    Exit Do
                End If
            End If
        ' ----
    Loop
    Selection.HomeKey unit:=wdStory
    If reply = vbNo Then
        Exit Do
    End If
    i = i + 1
Loop

'MsgBox selection.Text
End Sub
Public Sub JanyzakToWinKK()
'
' Replaces kyrgyz Janyzak Arial font characters
' to WinKK(Kazakh) ones
'
Dim i, kod As Integer
Dim strJnzk, strWinKK As String
Dim Jnzk(64) As Integer
Dim WinKK(64) As Integer
'Initialize Janizak symbols and codes
For i = 192 To 255
    Jnzk(i - 191) = i
    strJnzk = strJnzk & ChrW(i) & " "
Next i
'Initialize WinKK symbols and codes
For i = 1040 To 1103
    WinKK(i - 1039) = i
    strWinKK = strWinKK & ChrW(i) & " "
Next i

' Mismatch substitutions
WinKK(26) = 1065
WinKK(27) = 1066
WinKK(58) = 1097
WinKK(59) = 1098

' Start replacement
Dim found As Boolean
Dim reply As Integer
i = 1
Do While i <= 64
    Selection.HomeKey unit:=wdStory
    Selection.find.ClearFormatting
    With Selection.find
        .Text = ChrW$(Jnzk(i))
        .Replacement.Text = ChrW$(WinKK(i))
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute replace:=wdReplaceAll
    found = Selection.find.found
    'Selection.find.Text = ChrW$(Jnzk(i))
    'Selection.find.MatchCase = True
    'Selection.find.Replacement.Text = ChrW$(WinKK(i))
    'Selection.find.Execute
    'found = Selection.find.found
    'Selection.find.Execute replace:=wdReplaceAll
    i = i + 1
Loop

'MsgBox selection.Text
End Sub


Public Sub ConvertEncyclopediaCenterFontToWinKK()
'
'
' Replaces kyrgyz Encyclopedy Center Arial font characters to WinKK(Kazakh) ones
' Note: Not tested
'
Dim found As Boolean
Dim i As Integer
Dim EncCenter(6) As Integer
EncCenter(1) = 1108 ' kg soft 'u'
EncCenter(2) = 1113 ' kg soft 'o'
EncCenter(3) = 1111 ' kg      'ng'
EncCenter(4) = 1028 ' kg soft 'U' (Upper Case)
EncCenter(5) = 1033 ' kg soft 'O' (Upper Case)
EncCenter(6) = 1031 ' kg      'NG'(Upper Case)

Dim WinKK(6) As Integer
WinKK(1) = 1199 ' kg soft 'u'
WinKK(2) = 1257 ' kg soft 'o'
WinKK(3) = 1187 ' kg      'ng'
WinKK(4) = 1198 ' kg soft 'U' (Upper Case)
WinKK(5) = 1256 ' kg soft 'O' (Upper Case)
WinKK(6) = 1186 ' kg      'NG'(Upper Case)

i = 1
Dim quit As Boolean
quit = False
Do While i <= 6
    Selection.find.Text = ChrW$(EncCenter(i))
    Selection.find.MatchCase = True
    Selection.find.Execute
    found = Selection.find.found
    Do While found
        Selection.Text = ChrW$(WinKK(i))
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.find.Execute
        found = Selection.find.found
        
        ' ---- Loop checker
            iLoop = iLoop + 1
            If iLoop Mod 500 = 0 Then
                If MsgBox("Do you want to continue the loop", vbYesNo, "Debugging") = vbNo Then
                    quit = True
                    Exit Do
                End If
            End If
        ' ----
    Loop
    Selection.HomeKey unit:=wdStory
    If quit = True Then
        Exit Do
    End If
    i = i + 1
Loop

'MsgBox selection.Text
End Sub
Public Sub ConvertFineReaderKzToWinKK()
'
'
' Replaces kyrgyz Encyclopedy Center Arial font characters to WinKK(Kazakh) ones
' Note: Not tested
'
Dim found As Boolean
Dim i As Integer
Dim FReaderKz(6) As Integer
FReaderKz(1) = 1114 ' kg soft 'u'
FReaderKz(2) = 1113 ' kg soft 'o'
FReaderKz(3) = 1115 ' kg      'ng'
FReaderKz(4) = 1034 ' kg soft 'U' (Upper Case)
FReaderKz(5) = 1033 ' kg soft 'O' (Upper Case)
FReaderKz(6) = 1035 ' kg      'NG'(Upper Case)

Dim WinKK(6) As Integer
WinKK(1) = 1199 ' kg soft 'u'
WinKK(2) = 1257 ' kg soft 'o'
WinKK(3) = 1187 ' kg      'ng'
WinKK(4) = 1198 ' kg soft 'U' (Upper Case)
WinKK(5) = 1256 ' kg soft 'O' (Upper Case)
WinKK(6) = 1186 ' kg      'NG'(Upper Case)

i = 1
Do While i <= 6
    Selection.find.Text = ChrW$(FReaderKz(i))
    Selection.find.MatchCase = True
    Selection.find.Replacement.Text = ChrW$(WinKK(i))
    Selection.find.Execute replace:=wdReplaceAll
    
    Selection.HomeKey unit:=wdStory

    i = i + 1
Loop

'MsgBox selection.Text
'Tested.
End Sub
Public Sub ConvertEncyclopediaCenterFontToWinKK1()
'
'
' Replaces kyrgyz Encyclopedy Center Arial font characters to WinKK(Kazakh) ones
' Note: Not tested
'
Dim found As Boolean
Dim i As Integer
Dim EncCenter(6) As Integer
EncCenter(1) = 1108 ' kg soft 'u'
EncCenter(2) = 1113 ' kg soft 'o'
EncCenter(3) = 1111 ' kg      'ng'
EncCenter(4) = 1028 ' kg soft 'U' (Upper Case)
EncCenter(5) = 1033 ' kg soft 'O' (Upper Case)
EncCenter(6) = 1031 ' kg      'NG'(Upper Case)

Dim WinKK(6) As Integer
WinKK(1) = 1199 ' kg soft 'u'
WinKK(2) = 1257 ' kg soft 'o'
WinKK(3) = 1187 ' kg      'ng'
WinKK(4) = 1198 ' kg soft 'U' (Upper Case)
WinKK(5) = 1256 ' kg soft 'O' (Upper Case)
WinKK(6) = 1186 ' kg      'NG'(Upper Case)

i = 1
Do While i <= 6
    Selection.find.Text = ChrW$(EncCenter(i))
    Selection.find.MatchCase = True
    Selection.find.Replacement.Text = ChrW$(WinKK(i))
    Selection.find.Execute replace:=wdReplaceAll
    
    Selection.HomeKey unit:=wdStory

    i = i + 1
Loop

'MsgBox selection.Text
'Tested.
End Sub

Public Sub ConvertTimes_Q2ToWinKK1()
'
'
' Replaces kyrgyz Times_Q2 font characters to WinKK(Kazakh) ones
' Note: Not tested
'
Dim found As Boolean
Dim i As Integer
Dim Times_Q2(6) As Long
Times_Q2(1) = 61655 ' kg soft 'u'
Times_Q2(2) = 1108 ' kg soft 'o'
Times_Q2(3) = 1118 ' kg      'ng'
Times_Q2(4) = 1031 ' kg soft 'U' (Upper Case)
Times_Q2(5) = 1028 ' kg soft 'O' (Upper Case)
Times_Q2(6) = 1038 ' kg      'NG'(Upper Case)

Dim WinKK(6) As Integer
WinKK(1) = 1199 ' kg soft 'u'
WinKK(2) = 1257 ' kg soft 'o'
WinKK(3) = 1187 ' kg      'ng'
WinKK(4) = 1198 ' kg soft 'U' (Upper Case)
WinKK(5) = 1256 ' kg soft 'O' (Upper Case)
WinKK(6) = 1186 ' kg      'NG'(Upper Case)

i = 1
Do While i <= 6
    Selection.find.Text = ChrW$(Times_Q2(i))
    Selection.find.MatchCase = True
    Selection.find.Replacement.Text = ChrW$(WinKK(i))
    Selection.find.Execute 'replace:=wdReplaceAll
    
    Selection.HomeKey unit:=wdStory

    i = i + 1
Loop

'MsgBox selection.Text
'Tested.
End Sub
Public Sub Insert_Header_Tag()
Dim quote, nl, answer As String
quote = Chr(34)
nl = Chr(13)
Selection.HomeKey unit:=wdStory
answer = InputBox("Please enter the e-encyclopedy name:", "Encyclopedy Name")

Selection.InsertAfter "<?xml version=" & quote & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & nl
Selection.InsertAfter "<stardict xmlns:xi=" & quote & "http://www.w3.org/2003/XInclude" & quote & ">"
Selection.InsertAfter nl & "<info>" & nl
Selection.InsertAfter "<version>2.4.2</version>" & nl
Selection.InsertAfter "<bookname>" & answer & "</bookname>" & nl
Selection.InsertAfter "<author>Asanov A, Brimkulov U, Momunaliev K.</author>" & nl
Selection.InsertAfter "<email>unbrim@gmail.com, kadyr.momunaliev@gmail.com</email>" & nl
Selection.InsertAfter "<website>www.manas.kg</website>" & nl
Selection.InsertAfter "<description>Copyright: Kyrgyz Encyclopedia Editorial Board; Version: 1.0</description>" & nl
Selection.InsertAfter "<date>2012.10.30</date>" & nl
Selection.InsertAfter "<dicttype>Textual StarDict Dictionary</dicttype>" & nl
Selection.InsertAfter "</info>" & nl

Selection.EndKey unit:=wdStory
Selection.InsertAfter "</stardict>"
Selection.HomeKey unit:=wdStory
End Sub

Public Function set_doc_style_to(a_style As WdStyleType) As Integer
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory, Extend:=wdExtend
Selection.Style = a_style
set_doc_style_to = 0
End Function

Public Function is_term_char(ByVal c As Range) As Boolean

Dim extra, t_ch As String
Dim quote As String
extra = "()"
t_ch = extra & Chr(34) & Chr(187) & Chr(171) & ChrW$(8220) & ChrW$(8221) ' ["], [<<], [>>] - quotes added.
If is_alpha(c.Text) And c.Bold And c.Case = wdUpperCase Then
    is_term_char = True
    Exit Function
ElseIf InStr(t_ch, c.Text) > 0 And c.Bold Then
    is_term_char = True
Else
    is_term_char = False
End If

End Function
