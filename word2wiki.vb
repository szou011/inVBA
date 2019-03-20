Sub Word2TiddlyWiki()

    Application.ScreenUpdating = False

    ConvertH1
    ConvertH2
    ConvertH3
    ConvertH4
    ConvertH5

    ConvertItalic
    ConvertBold
    ConvertUnderline

    ConvertLists
    ' ConvertTables ... could not get it to work :(

    ' Copy to clipboard
    ActiveDocument.Content.Copy

    Application.ScreenUpdating = True
End Sub

Private Sub ConvertH1()
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)

    ActiveDocument.Select

    With Selection.Find

        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading1)
        .Text = ""

        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        .Forward = True
        .Wrap = wdFindContinue

        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "!"
                End If

                .Style = normalStyle
            End With
        Loop
    End With
End Sub

Private Sub ConvertH2()
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)

    ActiveDocument.Select

    With Selection.Find

        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading2)
        .Text = ""

        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        .Forward = True
        .Wrap = wdFindContinue

        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "!!"
                End If

                .Style = normalStyle
            End With
        Loop
    End With
End Sub

Private Sub ConvertH3()
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)

    ActiveDocument.Select

    With Selection.Find

        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading3)
        .Text = ""

        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        .Forward = True
        .Wrap = wdFindContinue

        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "!!!"
                End If

                .Style = normalStyle
            End With
        Loop
    End With
End Sub
Private Sub ConvertH4()
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)

    ActiveDocument.Select

    With Selection.Find

        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading3)
        .Text = ""

        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        .Forward = True
        .Wrap = wdFindContinue

        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline  characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                    ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "!!!!"
                End If

                .Style = normalStyle
            End With
        Loop
    End With
End Sub
Private Sub ConvertH5()
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)

    ActiveDocument.Select

    With Selection.Find

        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading3)
        .Text = ""

        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        .Forward = True
        .Wrap = wdFindContinue

        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "!!!!!"
                End If

                .Style = normalStyle
            End With
        Loop
    End With
End Sub

Private Sub ConvertBold()
    ActiveDocument.Select

    With Selection.Find

        .ClearFormatting
        .Font.Bold = True
        .Text = ""

        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        .Forward = True
        .Wrap = wdFindContinue

        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                          .Font.Bold = False
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "''"
                    .InsertAfter "''"
                End If

                .Font.Bold = False
            End With
        Loop
    End With
End Sub

Private Sub ConvertItalic()
    ActiveDocument.Select

    With Selection.Find

        .ClearFormatting
        .Font.Italic = True
        .Text = ""

        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        .Forward = True
        .Wrap = wdFindContinue

        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Font.Italic = False
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "//"
                    .InsertAfter "//"
                End If

                .Font.Italic = False
            End With
        Loop
    End With
End Sub

Private Sub ConvertUnderline()
    ActiveDocument.Select

    With Selection.Find

        .ClearFormatting
        .Font.Underline = True
        .Text = ""

        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        .Forward = True
        .Wrap = wdFindContinue

        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Font.Underline = False
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "__"
                    .InsertAfter "__"
                End If

                .Font.Underline = False
            End With
        Loop
    End With
End Sub

Private Sub ConvertLists()
   Dim para As Paragraph
    For Each para In ActiveDocument.ListParagraphs
        With para.Range
            .InsertBefore " "
            For i = 1 To .ListFormat.ListLevelNumber
                If .ListFormat.ListType = wdListBullet Then
                    .InsertBefore "*"
                Else
                    .InsertBefore "#"
                End If
            Next i
            .ListFormat.RemoveNumbers
        End With
    Next para
End Sub
