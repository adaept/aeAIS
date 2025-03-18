VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbLanguageListFromIni 
   Caption         =   "Select Current Language Project"
   ClientHeight    =   2772
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6924
   OleObjectBlob   =   "dbLanguageListFromIni.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbLanguageListFromIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LanguageNumber As String
Dim Number As Integer

Dim LanguageNameTemp As Variant

Dim vLanguageCode As String
Dim vLanguageProjectCode As String
Dim vLanguageProvince As String
Dim vLanguageCountry As String
Dim vLanguageFont As String
Dim vLanguageHeadingFont As String
Dim vLanguageLeading As String
Dim vLanguageSize As String
Dim vLanguageQuotesInProofPrintouts As String
Dim vLanguageDropCapChapterNumbers As String
Dim vLanguageHideNumberForEachVerse1 As String
Dim vLanguageHeaderOutside As String
Dim vLanguageHeaderOther As String
Dim vLanguageRestartFootnoteRefs As String
Dim vLanguageNoBreakHyphens As String
Dim vLanguageNoBreakSpaces As String
Dim vLanguageBrackets2HalfBrackets As String
Dim vLanguageBoldVerseNumbers As String
Dim vLanguageCheckingTable As String

Dim FileSystem As Object

Option Explicit
Private Sub ShowCurrentSettings()
Dim Msg As String, Style As Long, Title As String, Result As Integer
Dim sHideNumberForEachVerse1 As String

    If vLanguageHideNumberForEachVerse1 = "yes" Then
        sHideNumberForEachVerse1 = "Verse 1 in each chapter will be hidden." & vbCrLf
    Else
        sHideNumberForEachVerse1 = ""
    End If

    Msg = "The file " & Chr(34) & LanguageFileName & Chr(34) & _
            " contains these settings for this language:" & vbCrLf & vbCrLf & _
        "Language: " & vbTab & LanguageNameTemp & vbCrLf & _
        IIf(vLanguageCode <> "", "Language Code: " & vbTab & vLanguageCode & vbCrLf, "") & _
        IIf(vLanguageProjectCode <> "", _
            "ProjectCode: " & vbTab & vLanguageProjectCode & vbCrLf, "") & _
        "Province: " & vbTab & vLanguageProvince & vbCrLf & _
        "Country: " & vbTab & vbTab & vLanguageCountry & vbCrLf & _
        "Font: " & vbTab & vbTab & vLanguageFont & vbCrLf & _
        "   Font Size: " & vbTab & vLanguageSize & vbCrLf & _
        "   Line Spacing: " & vbTab & vLanguageLeading & vbCrLf & _
        "   Heading Font: " & vbTab & vLanguageHeadingFont & vbCrLf & _
        "Angle brackets (<< >>) changed to quotes in proof printouts:  " & _
            vLanguageQuotesInProofPrintouts & vbCrLf & _
        "Chapter numbers formatted as Drop Caps:  " & vLanguageDropCapChapterNumbers & vbCrLf & _
        sHideNumberForEachVerse1 & _
        "Headers will have " & vLanguageHeaderOutside & " at the outside edge." & vbCrLf & _
        "Footnote references will restart from 'a' each page: " & vLanguageRestartFootnoteRefs & vbCrLf & _
        "Make hyphens non-breaking in Proof printouts: " & vLanguageNoBreakHyphens & vbCrLf & _
        "Change hyphens into thin no-break spaces in Booklets: " & _
            vLanguageNoBreakSpaces & vbCr & _
        "Change brackets ([ and ]) to half brackets to mark implied information: " & _
            vLanguageBrackets2HalfBrackets & vbCrLf & _
        "Make verse numbers bold: " & vLanguageBoldVerseNumbers & vbCrLf & vbCrLf & _
        "Do you want to change any of these settings?"
    Style = vbQuestion + vbYesNo + vbDefaultButton2
    Title = "Current Language Settings 2"
    Result = MsgBox(Msg, Style, Title)
    
    If Result = vbYes Then
        ' This sets doc variables;
        ShowLanguageInfoForm
        If VariableGet("NoLanguageSet") = "True" Then Exit Sub
        DoFontChangesFromIni
    
    ElseIf Result = vbNo Then
        If VariableGet("NoLanguageSet") = "True" Then Exit Sub
        DoFontChangesFromIni
    End If
End Sub
Private Sub ListLanguages()
    If IniFile = "" Then SetIniFile
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    
    If Not FileSystem.FileExists(IniFile) Then
        Dim Msg As String, Style As Long, Result As Integer, Title As String
        Msg = "I can't find the file, " & LanguageFileName & ". " & _
            "Would you like me to create it for you?" & vbCrLf & vbCrLf & _
            "Click NO if you think " & _
            "the file should be there and you would like to look for it manually." & _
            vbCrLf & "Click YES if you're just starting to use this template, " & _
            "or if you want to reset the languages."
        Style = vbYesNo
        Title = "Languages.ini Not Found (ListLangs)"
        Result = MsgBox(Msg, Style, Title)
        If Result = vbYes Then
            CreateIniFile
            Number = 1
        ElseIf Result = vbNo Then
            Exit Sub
        End If
    End If

    Number = 1  ' Start with record Language_1
    Do
         LanguageNumber = "Language_" & Number
         
         LanguageNameTemp = System.PrivateProfileString( _
             FileName:=IniFile, Section:=LanguageNumber, _
             Key:="Name")
         
         ' This is so we can tell when we have processed all the languages.
         ' The result is that there can be no gap in the sequence of language numbers
         ' in Languages.ini file.
         '
         If LanguageNameTemp = "" Then Exit Do
         If LanguageNameTemp = " " Then Exit Do
         
         With cmboIniLanguages
             .AddItem LanguageNameTemp
         End With
         
         Number = Number + 1
     Loop While LanguageNameTemp <> ""
End Sub
Private Sub cmboIniLanguages_AfterUpdate()
    Number = cmboIniLanguages.ListIndex + 1
    LanguageNumber = "Language_" & Number
End Sub
Private Sub cmdAddNewLang2_Click()
    If cmboIniLanguages.Text <> "" Then
        VariableSet "tmpLang", cmboIniLanguages.Text
        VariableSet "LanguageName", cmboIniLanguages.Text
    End If
    VariableSet "NoLanguageSet", ""
    If VariableGet("WhenExported") = "" Then
        VariableSet "WhenExported", Format(Now, "dd-Mmm-yyyy") & " at " & Format(Now, "hh:mm")
    End If
    Me.Hide
    LangData.AddNewLanguageToIni
    Unload Me
End Sub
Private Sub cmdCancel_Click()
Dim myvar As Variable

For Each myvar In ActiveDocument.Variables
    myvar.Delete
    Next myvar
VariableSet "NoLanguageSet", "True"
Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim Answer As Integer
    Dim tempName As String
    Dim vPTNoFontChange As String
    
    If IniFile = "" Then SetIniFile
    
    'if they type in a language that doesn't match, tell them to click "Add a New"
    tempName = cmboIniLanguages.Text
    
    If tempName = "" Or tempName = " " Then GoTo NoLangsInIniFile 'ch11/28
    
    If cmboIniLanguages.ListCount = 0 Then GoTo NoLangsInIniFile
    
    For Each LanguageNameTemp In cmboIniLanguages.List
        If tempName = LanguageNameTemp Then GoTo OK
    Next LanguageNameTemp
    
NoLangsInIniFile:
    Answer = MsgBox("You have entered a new language for the settings file." & vbCrLf & _
            "Click 'OK' to add this as a new language," & vbCrLf & _
            "or click 'Cancel' to select another language from the list.", vbOKCancel)
    If Answer = vbOK Then
        GoTo EndNew
    Else: GoTo EndCancel
    End If
    
OK:
    VariableSet "NoLanguageSet", ""
    If VariableGet("WhenExported") = "" Then
        VariableSet "WhenExported", Format(Now, "dd-Mmm-yyyy") & " at " & Format(Now, "hh:mm")
    End If
    VariableSet "LanguageNumber", LanguageNumber
    LanguageNameTemp = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
            Key:="Name")
    VariableSet "LanguageName", LanguageNameTemp

    'now set the rest of the variables
    vLanguageCode = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="Code")
    
    vLanguageProjectCode = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="ProjectCode")
                    
    vLanguageProvince = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="Province")
            
    vLanguageCountry = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="Country")
            
    vLanguageFont = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="Font")
            
    vLanguageHeadingFont = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="HeadingFont")
                    
    vLanguageLeading = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="FontLeading")
    
    vLanguageSize = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="DefaultSizeInPoints")
        
    vLanguageQuotesInProofPrintouts = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="QuotesInProofPrintouts")
        
    vPTNoFontChange = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="PTNoFontChange")
    
    vLanguageDropCapChapterNumbers = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="DropCapChapterNumbers")

    vLanguageHideNumberForEachVerse1 = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="HideNumberForEachVerse1")

    vLanguageHeaderOutside = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="HeaderOutside")

    vLanguageHeaderOther = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="HeaderOther")

    vLanguageRestartFootnoteRefs = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="RestartFootnoteRefs")

    vLanguageNoBreakHyphens = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="NoBreakHyphens")

    vLanguageNoBreakSpaces = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="NoBreakSpaces")

    vLanguageBrackets2HalfBrackets = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="Brackets2HalfBrackets")

    vLanguageBoldVerseNumbers = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="BoldVerseNumbers")

    vLanguageCheckingTable = System.PrivateProfileString( _
        FileName:=IniFile, Section:=LanguageNumber, _
        Key:="CheckingTable")
        
    VariableSet "LanguageProvince", vLanguageProvince
        If VariableGet("LanguageProvince") = "" Then VariableSet "LanguageProvince", " "
    VariableSet "LanguageCountry", vLanguageCountry
        If VariableGet("LanguageCountry") = "" Then VariableSet "LanguageCountry", " "
    VariableSet "LanguageSize", vLanguageSize
        If VariableGet("LanguageSize") = "" Then VariableSet "LanguageSize", " "
    VariableSet "QuotesInProofPrintouts", vLanguageQuotesInProofPrintouts
        If VariableGet("QuotesInProofPrintouts") = "" Then VariableSet "QuotesInProofPrintouts", " "
    VariableSet "DropCapChapterNumbers", vLanguageDropCapChapterNumbers
        If VariableGet("DropCapChapterNumbers") = "" Then VariableSet "DropCapChapterNumbers", " "
    VariableSet "HideNumberForEachVerse1", vLanguageHideNumberForEachVerse1
        If VariableGet("HideNumberForEachVerse1") = "" Then VariableSet "HideNumberForEachVerse1", " "
    VariableSet "LanguageCode", vLanguageCode
        If VariableGet("LanguageCode") = "" Then VariableSet "LanguageCode", " "
    VariableSet "ProjectCode", vLanguageProjectCode
        If VariableGet("ProjectCode") = "" Then VariableSet "ProjectCode", " "
    VariableSet "LanguageFont", vLanguageFont
        If VariableGet("LanguageFont") = "" Then VariableSet "LanguageFont", " "
    VariableSet "HeadingFont", vLanguageHeadingFont
        If VariableGet("HeadingFont") = "" Then VariableSet "HeadingFont", " "
    VariableSet "Leading", vLanguageLeading
        If VariableGet("Leading") = "" Then VariableSet "Leading", " "
    VariableSet "HeaderOutside", vLanguageHeaderOutside
        If VariableGet("HeaderOutside") = "" Then VariableSet "HeaderOutside", " "
    VariableSet "HeaderOther", vLanguageHeaderOther
        If VariableGet("HeaderOther") = "" Then VariableSet "HeaderOther", " "
    VariableSet "RestartFootnoteRefs", vLanguageRestartFootnoteRefs
        If VariableGet("RestartFootnoteRefs") = "" Then VariableSet "RestartFootnoteRefs", " "
    VariableSet "NoBreakHyphens", vLanguageNoBreakHyphens
        If VariableGet("NoBreakHyphens") = "" Then VariableSet "NoBreakHyphens", " "
    VariableSet "NoBreakSpaces", vLanguageNoBreakSpaces
        If VariableGet("NoBreakSpaces") = "" Then VariableSet "NoBreakSpaces", " "
    VariableSet "Brackets2HalfBrackets", vLanguageBrackets2HalfBrackets
        If VariableGet("Brackets2HalfBrackets") = "" Then VariableSet "Brackets2HalfBrackets", " "
    VariableSet "BoldVerseNumbers", vLanguageBoldVerseNumbers
        If VariableGet("BoldVerseNumbers") = "" Then VariableSet "BoldVerseNumbers", " "
    VariableSet "CheckingTable", vLanguageCheckingTable
        If VariableGet("CheckingTable") = "" Then VariableSet "CheckingTable", " "
    VariableSet "PTNoFontChange", vPTNoFontChange
        If VariableGet("PTNoFontChange") = "" Then VariableSet "PTNoFontChange", " "

    ' now check font with PT output font
    Gen.SetDefaultFont
    LangData.DoFontChangesFromIni
    ShowCurrentSettings
    Me.Hide

EndNew:
   cmdAddNewLang2_Click
EndCancel:
    Unload Me
End Sub
Private Sub UserForm_Initialize()
ListLanguages

If cmboIniLanguages.ListCount = 0 Then Exit Sub

cmboIniLanguages.ListIndex = 0
cmboIniLanguages.SelStart = 0
cmboIniLanguages.SelLength = Len(cmboIniLanguages.Text)

End Sub
