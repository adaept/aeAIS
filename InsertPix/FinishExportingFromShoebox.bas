Attribute VB_Name = "FinishExportingFromShoebox"
Public Declare Function sndPlaySound Lib "winmm.dll" _
Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long
Global gInitDir As String
Global gFileName As String
Global gCommonDialog As Control
Global glPath As String
Global glFileName As String

   
'FinishExportingFromShoebox macro for MDF in Shoebox 5.0
'2000-07-21 Mark R. Pedrotti
'2003-07-25 modified to use VBA routine for inserting graphics
Public Sub main()

Dim sEnvironment$
Dim bWindows
Dim sWordVersion$
Dim sFilePath$
Dim lenPath
Dim sMessage$
    UpdateStyles
    AlignColumnsAtSectionBreaks  '2000-07-21
    VBAInsertGraphic
    KeepHeadingWithNextEntry
    WordBasic.StartOfDocument

    sEnvironment$ = WordBasic.[AppInfo$](1)  '1997-10-01
    bWindows = (InStr(sEnvironment$, "Windows") <> 0)

    ' 1998-03-17 MRP: The distributed .DOT file has the binary format
    ' of Word 6 or 7. Word 8 (or later) want to save the template
    ' in its own binary format. Let the macro handle this situation
    ' so that the user isn't prompted to confirm saving it,
    ' because they probably won't understand why this is happening.
    If WordBasic.IsTemplateDirty() Then
        sWordVersion$ = WordBasic.[AppInfo$](2)
        If InStr(sWordVersion$, "6.0") <> 1 And InStr(sWordVersion$, "7.0") <> 1 Then
            WordBasic.SendKeys "%O"  ' Overwrite the same .dot file
            WordBasic.SaveTemplate
        End If
    End If

    ' 1997-09-17 MRP: Save .RTF file as a Word document
    sFilePath$ = WordBasic.[FileNameFromWindow$]()
    lenPath = Len(sFilePath$)
    If lenPath >= 4 Then
        If LCase(WordBasic.[Right$](sFilePath$, 4)) = ".rtf" Then
            sFilePath$ = WordBasic.[Left$](sFilePath$, lenPath - 4)
            If bWindows Then
                sFilePath$ = sFilePath$ + ".doc"
            End If
        End If
    End If
    WordBasic.FileSaveAs Name:=sFilePath$, Format:=0

    WordBasic.ScreenRefresh
    WordBasic.Beep
    sMessage$ = "Finished exporting from Shoebox."
    If bWindows Then
        sMessage$ = sMessage$ + Chr(10)
    Else
        sMessage$ = sMessage$ + Chr(13)
    End If
    sMessage$ = sMessage$ + "Saved the file as a Word Document."
    WordBasic.MsgBox sMessage$
End Sub  'MAIN

Private Sub UpdateStyles()
    'MDF in Shoebox converts markers to styles and does the
    'document and section formatting (Export Page Setup),
    'but all style formatting attributes come from the template.
    WordBasic.FileSummaryInfo Update:=1
    Dim dlg As Object: Set dlg = WordBasic.DialogRecord.DocumentStatistics(False)
    WordBasic.CurValues.DocumentStatistics dlg
    WordBasic.FileTemplates Template:=dlg.Template, LinkStyles:=1
End Sub

Private Sub AlignColumnsAtSectionBreaks()
Dim sParagraphStyle$
Dim sParagraphStylePrev$
    'Cause entries at the end of the double-column sections
    'to line up by deleting the paragraphs that Shoebox inserts.
    'This way, the section breaks directly follow the paragraph marks
    'of the last entry, indented, block, or finderlist paragraph.
    WordBasic.StartOfDocument
    WordBasic.EditFindStyle Style:=""
    WordBasic.EditFind find:="^p", Direction:=0, MatchCase:=0, WholeWord:=0, PatternMatch:=0, SoundsLike:=0, Format:=1, Wrap:=0
    While WordBasic.EditFindFound() = -1
        sParagraphStyle$ = WordBasic.[StyleName$]()
        If sParagraphStylePrev$ <> "" Then
            If sParagraphStyle$ = "Letter Section" Then WordBasic.WW6_EditClear
            If sParagraphStyle$ = "Single-column Section" Then WordBasic.WW6_EditClear
        End If
        sParagraphStylePrev$ = sParagraphStyle$
        WordBasic.RepeatFind
    Wend
End Sub


Private Sub KeepHeadingWithNextEntry()
Dim iViewNormal
Dim iPos
    '1996-01-22 MRP: I think this needs to be rewritten more carefully
    'Cause Letter Paragraphs to keep with the next (i.e. first)
    'Entry Paragraph of the section. The ordinary "Keep with next"
    'paragraph property is defeated by the section break after
    'the letter heading (when the lexical entries are double-column).
    'Insert a page break before any Letter Paragraph that is alone
    'at the bottom of a page (separated from its lexical entries).

    WordBasic.PrintStatusBar "Checking position of headings..."
    iViewNormal = WordBasic.ViewNormal()
    WordBasic.ViewPage
    WordBasic.ShowAll  'Cause Word to repaginate after styles are updated
    WordBasic.StartOfDocument
    iPos = WordBasic.GetSelStartPos()
    WordBasic.GoToNextPage
    While WordBasic.GetSelStartPos() <> iPos
        WordBasic.CharLeft 2
        If WordBasic.[StyleName$]() = "Double-column Section" Then
            WordBasic.ParaUp
            WordBasic.WW7_InsertPageBreak
        Else
            WordBasic.CharRight 2
        End If
        iPos = WordBasic.GetSelStartPos()
        WordBasic.GoToNextPage
    Wend
    WordBasic.ShowAll  'Return to the original Show/Hide PP setting
    If iViewNormal <> 0 Then WordBasic.ViewNormal
End Sub





Sub VBAInsertGraphic()
'rewrite of "InsertGraphic" macro to use VBA calls rather than WordBasic calls
'using VBA, it doesn't need to do a search to find the picture frames
'It also does not require dimensions or image type to be specified.
'More reliable in inserting long file or path names.
' Steve White, Language Software Support, Jaars

'The pc field data examples:
'no dimensions specified: .G.aehu.bmp
'    (if no dimensions specified, Word uses defaults, which are probably the native image size, or the space available on the page
'dimensions, no image type: .G.c:\my documents\my pictures\aehu.bmp;.5";2.5"
'dimensions and image type: .G.c:\documents and settings\white\my documents\my pictures\aehu.bmp;.5";2.5";BMP
'the image type info is not used by this macro, as it is not needed.
For Each aFrame In ActiveDocument.Frames
    On Error GoTo errorhandler
    sField$ = aFrame.Range
    aFrame.Range.Delete
    lenField = Len(sField$)  '23
    If lenField = 0 Then GoTo nextaframe
    'Sometimes when an image link is near the beginning of the entry, VBA detects
    'an extra frame with no contents.
    iAfterFilePath = InStr(1, sField$, ";") '9
    If (iAfterFilePath) > 0 Then
        sFilePath$ = Mid(sField$, 1, iAfterFilePath - 1)  '
        iAfterWidth = InStr(iAfterFilePath + 1, sField$, ";")  '15
        lenWidth = iAfterWidth - iAfterFilePath - 1  '5
        sWidth$ = Mid(sField$, iAfterFilePath + 1, lenWidth)  '1.25"
        iAfterHeight = InStr(iAfterWidth + 1, sField$, ";")  '20
        If (iAfterHeight) > 0 Then
            lenHeight = iAfterHeight - iAfterWidth - 1  '4
        Else
            lenHeight = lenField - iAfterWidth
        End If
        sHeight$ = Mid(sField$, iAfterWidth + 1, lenHeight)  '1.1"
    Else
        sFilePath$ = sField$
    End If
    If InStr(sFilePath$, ":") = 0 Then sFilePath$ = ActiveDocument.Path + Application.PathSeparator + sFilePath$
    aFrame.Range.InlineShapes.AddPicture sFilePath$, False, True
    If (iAfterFilePath) > 0 Then
        aFrame.Height = InchesToPoints(Val(sHeight$))
        aFrame.Range.InlineShapes(1).Height = InchesToPoints(Val(sHeight$))
        aFrame.Width = InchesToPoints(Val(sWidth$))
        aFrame.Range.InlineShapes(1).Width = InchesToPoints(Val(sWidth$))
    End If
nextaframe:
Next aFrame
Exit Sub
errorhandler:
MsgBox "Error inserting picture file " + sFilePath$ + Chr(13) + "Is the path name correct?"
On Error Resume Next
Resume Next
End Sub

Sub Hyperlink_separatefield()
'Makes sound file (or video file) names become hyperlinks. This version expects the sound file pathname
'to be in a separate field \sou in Shoebox.

    Selection.HomeKey Unit:=wdStory
    With Selection.find
        .Text = "[?? \sou"
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
loopagain:
    Selection.find.Execute
    If Selection.find.Found Then
        Selection.Delete
        Selection.Extend character:="]"
            ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        Left(Selection.Text, Len(Selection.Text) - 1), SubAddress:="", ScreenTip:="", TextToDisplay _
        :="Sound"
        GoTo loopagain
    End If
End Sub
Sub Hyperlink_inline_textaslabel()
'makes sound (or video) file names into hyperlinks.
'This version looks for the file name as some inline formatting in a regular MDF field
'and makes the text contents of that field into the hyperlink label
    Selection.HomeKey Unit:=wdStory
loopagain:
    With Selection.find
        .Text = "fh{"
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
    If Selection.find.Found Then
        sname$ = Selection.Style
        Selection.Delete
        Selection.Extend
        With Selection.find
            .Text = ""
            .Forward = True
            .Format = True
            .Style = sname$
        End With
        Selection.find.Execute
        slstring = Trim(Selection.Text)
        Selection.Delete
        Selection.Extend
        With Selection.find
            .Text = ""
            .Forward = False
            .Format = True
            .Style = sname$
        End With
        Selection.find.Execute
            ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        slstring, SubAddress:="", ScreenTip:="", TextToDisplay:=Selection.Text
        Selection.Collapse
        Selection.ExtendMode = False
        GoTo loopagain
    End If
For Each hLink In ActiveDocument.Hyperlinks
    hLink.Address = Trim(hLink.Address)
Next hLink
End Sub
Sub Hyperlink_inline_addlabel()
'makes sound (or video) file names into hyperlinks.
'This version looks for the file name as some inline formatting at the end of a regular MDF field
'and adds a hyperlink. All hyperlinks will have the same label.
'change text in quotes below to change the hyperlink label.
LabelText$ = "Hear It "
Options.AutoFormatAsYouTypeReplaceHyperlinks = False
    Selection.HomeKey Unit:=wdStory
loopagain:
    With Selection.find
        .Text = "fh{"
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
    If Selection.find.Found Then
        sname$ = Selection.Style
        Selection.Delete
        'With ActiveDocument.Bookmarks
        '.Add Range:=Selection.Range, Name:="x"
        '.DefaultSorting = wdSortByName
        '.ShowHidden = False
        'End With
        'Selection.TypeText Text:="x" 'formerly delete
        Selection.MoveLeft wdCharacter
        stemp$ = Selection.Text
        Selection.MoveRight wdCharacter
        If stemp$ <> " " Then Selection.TypeText Text:=" "
        Selection.Extend
        With Selection.find
            .Text = ""
            .Forward = True
            .Format = True
            .Style = sname$
        End With
        Selection.find.Execute
        slstring = Selection.Text
        Selection.Delete
        
            ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        slstring, SubAddress:="", ScreenTip:="", TextToDisplay:=LabelText$
        Selection.Collapse  ' formerly collapse
        Selection.TypeText Text:=" "
        'Add$ = ActiveDocument.Hyperlinks(1).Address
        'ActiveDocument.Bookmarks("x").Delete
        'With ActiveDocument.Bookmarks
        '    .DefaultSorting = wdSortByName
        '    .ShowHidden = False
        'End With
        Selection.ExtendMode = False
        GoTo loopagain
    End If
For Each hLink In ActiveDocument.Hyperlinks
    hLink.Address = Trim(hLink.Address)
Next hLink
    
End Sub

Sub Hyperlink_inline_textaslabel2()
'makes sound (or video) file names into hyperlinks.
'This version looks for the file name as some inline formatting in a regular MDF field
'and makes the text contents of that field into the hyperlink label
    Selection.HomeKey Unit:=wdStory
loopagain:
    With Selection.find
        .Text = " %"
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
    If Selection.find.Found Then
        sname$ = Selection.Style
        Selection.Delete
        Selection.Extend
        With Selection.find
            .Text = ""
            .Forward = True
            .Format = True
            .Style = sname$
        End With
        Selection.find.Execute
        slstring = Trim(Selection.Text)
        Selection.Delete
        Selection.Extend
        With Selection.find
            .Text = ""
            .Forward = False
            .Format = True
            .Style = sname$
        End With
        Selection.find.Execute
            ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        slstring, SubAddress:="", ScreenTip:="", TextToDisplay:=Selection.Text
        Selection.Collapse
        Selection.ExtendMode = False
        GoTo loopagain
    End If
For Each hLink In ActiveDocument.Hyperlinks
    hLink.Address = Trim(hLink.Address)
Next hLink
End Sub


