VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbLanguageInfo 
   Caption         =   "Language Project Information Form"
   ClientHeight    =   8064
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11472
   OleObjectBlob   =   "dbLanguageInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbLanguageInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpFont As String
Dim tmpHFont As String
Dim tmpSpacing As String
Dim AddAsNewLanguage As Boolean

Option Explicit
Private Sub QuickSort(strArray() As String, intBottom As Integer, intTop As Integer)
' written by Steve White, 4 Nov 05

  Dim strPivot As String, strTemp As String
  Dim intBottomTemp As Integer, intTopTemp As Integer

  intBottomTemp = intBottom
  intTopTemp = intTop

  strPivot = strArray((intBottom + intTop) \ 2)

  While (intBottomTemp <= intTopTemp)
    
    While (strArray(intBottomTemp) < strPivot And intBottomTemp < intTop)
      intBottomTemp = intBottomTemp + 1
    Wend
    
    While (strPivot < strArray(intTopTemp) And intTopTemp > intBottom)
      intTopTemp = intTopTemp - 1
    Wend
    
    If intBottomTemp < intTopTemp Then
      strTemp = strArray(intBottomTemp)
      strArray(intBottomTemp) = strArray(intTopTemp)
      strArray(intTopTemp) = strTemp
    End If
    
    If intBottomTemp <= intTopTemp Then
      intBottomTemp = intBottomTemp + 1
      intTopTemp = intTopTemp - 1
    End If
  
  Wend

  'the function calls itself until everything is in good order
  If (intBottom < intTopTemp) Then QuickSort strArray, intBottom, intTopTemp
  If (intBottomTemp < intTop) Then QuickSort strArray, intBottomTemp, intTop

End Sub

Private Sub cbChangeHyphens_Click()

End Sub

Private Sub CheckBox2_Click()

End Sub

Private Sub cmboFontArray_AfterUpdate()
tmpFont = cmboFontArray.Text
End Sub

Private Sub cmboFontArray_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
tmpFont = cmboFontArray.Text
End Sub

Private Sub cmboFontArray_Click()
' VariableSet "LanguageFont", aFont
End Sub

Private Sub cmdAddNewLanguage_Click()
    VariableSet "tmpLang", txtLanguageNameIni
    AddAsNewLanguage = True
    cmdOK_Click
End Sub
Private Sub cmdAdvancedFont_Click()
    tmpFont = cmboFontArray.Text
    tmpHFont = VariableGet("HeadingFont")
    
    With dbAdvancedFontOptions
        If tmpHFont <> "" And tmpHFont <> " " Then
            .cmboFontArray = tmpHFont
        Else: .cmboFontArray = tmpFont
        End If
        .Show
    End With
End Sub
Private Sub cmdCancel_Click()
AddAsNewLanguage = False
Unload Me
End Sub
Private Sub cmdOK_Click()
'Copy current db values to ActiveDocument variables.
' Calls SendLanguageData2IniFile to set Ini values.
    Dim Response As Integer
    
    If VariableGet("LanguageName") = "" Then VariableSet "LanguageName", " "

    If dbLanguageInfo.txtLanguageNameIni = "" Then
        dbLanguageInfo.txtLanguageNameIni = " "
        MsgBox "Please specify a language name..."
        
    ElseIf dbLanguageInfo.txtLanguageNameIni.Value = ActiveDocument.Variables("LanguageName").Value Then
        AssignVariables
        VariableSet "NoLanguageSet", ""
        If VariableGet("WhenExported") = "" Then
            VariableSet "WhenExported", Format(Now, "dd-Mmm-yyyy") & " at " & Format(Now, "hh:mm")
        End If
        SendLanguageData2IniFile
        Unload Me
    ElseIf dbLanguageInfo.txtLanguageNameIni.Value <> ActiveDocument.Variables("LanguageName").Value Then
        Response = 0
        
        If AddAsNewLanguage = False Then
            Response = MsgBox("If you click 'Yes' here, you will change the name of this language. " & vbCrLf & _
                "If you want to add a new language to the list, please click 'No' here, then click " & vbCrLf & _
                "     'Add as new language.'" & vbCrLf & vbCrLf & _
                "Do you really want to change the name of this language?", vbQuestion + vbYesNo + vbDefaultButton2)
        End If
        
        If Response = vbYes _
        Or AddAsNewLanguage = True Then
            AssignVariables
            VariableSet "NoLanguageSet", ""
            If VariableGet("WhenExported") = "" Then
                VariableSet "WhenExported", Format(Now, "dd-Mmm-yyyy") & " at " & Format(Now, "hh:mm")
            End If
            SendLanguageData2IniFile
            Unload Me
        ElseIf Response = vbNo Then
            'stay here and wait for another click
        End If
    End If
End Sub
Private Sub ChangeFontComboBox()
' written by Steve White, 4 Nov 05
Dim afont As Variant
Dim fonts() As String
Dim i As Long

ReDim fonts(FontNames.Count)

For Each afont In FontNames
    fonts(i) = afont
    i = i + 1
Next afont

QuickSort fonts(), LBound(fonts), UBound(fonts)

For i = 0 To UBound(fonts)
    cmboFontArray.AddItem fonts(i)
Next i
End Sub
Private Sub AssignVariables()

    If dbLanguageInfo.optQuotesYes = True Then
        VariableSet "QuotesInProofPrintouts", "yes"
    Else: VariableSet "QuotesInProofPrintouts", "no"
    End If

    If dbLanguageInfo.optDropCapsYes = True Then
        VariableSet "DropCapChapterNumbers", "yes"
    Else
        VariableSet "DropCapChapterNumbers", "no"
        cbHideVerse1.Enabled = False
    End If

    If cbHideVerse1 = True Then
        VariableSet "HideNumberForEachVerse1", "yes"
    Else
        VariableSet "HideNumberForEachVerse1", "no"
    End If

    If dbLanguageInfo.optBookOutside = True Then
        VariableSet "HeaderOutside", "BookChapter"
    Else: VariableSet "HeaderOutside", "PageNumber"
    End If
    
    If dbLanguageInfo.optCenterHdr = True Then
        VariableSet "HeaderOther", "center"
    Else: VariableSet "HeaderOther", "inner"
    End If
    
    If cbRestartFootnoteRefs = True Then
        If VariableGet("RestartFootnoteRefs") <> "done" Then
            VariableSet "RestartFootnoteRefs", "yes"
        End If
    Else
        If VariableGet("RestartFootnoteRefs") = "yes" _
        Or VariableGet("RestartFootnoteRefs") = "done" Then
            FootnotesRestartingOff
        End If
        
        VariableSet "RestartFootnoteRefs", "no"
    End If
    
    If dbLanguageInfo.cbNoBreakHyphens = True Then
        VariableSet "NoBreakHyphens", "yes"
    Else: VariableSet "NoBreakHyphens", "no"
    End If
    
    If dbLanguageInfo.cbNoBreakSpaces = True Then
        VariableSet "NoBreakSpaces", "yes"
    Else: VariableSet "NoBreakSpaces", "no"
    End If
    
    If dbLanguageInfo.cbBrackets2HalfBrackets = True Then
        VariableSet "Brackets2HalfBrackets", "yes"
    Else: VariableSet "Brackets2HalfBrackets", "no"
    End If
    
    If dbLanguageInfo.cbBoldVerseNumbers = True Then
        VariableSet "BoldVerseNumbers", "yes"
    Else: VariableSet "BoldVerseNumbers", "no"
    End If
    
    VariableSet "LanguageName", dbLanguageInfo.txtLanguageNameIni
    VariableSet "LanguageCode", dbLanguageInfo.txtLangCodeIni
    VariableSet "ProjectCode", dbLanguageInfo.txtProjectCodeIni
    VariableSet "LanguageProvince", dbLanguageInfo.txtProvinceIni
    VariableSet "LanguageCountry", dbLanguageInfo.txtCountryIni
    VariableSet "LanguageFont", dbLanguageInfo.cmboFontArray
    VariableSet "LanguageSize", dbLanguageInfo.txtFontSizeIni
    VariableSet "LanguageNumber", dbLanguageInfo.txtIniNumber
    AddAsNewLanguage = False
    Unload Me
End Sub
Private Sub frmDropCaps_Click()
    If dbLanguageInfo.optDropCapsYes = True Then
        cbHideVerse1.Enabled = True
    Else
        cbHideVerse1.Enabled = False
    End If
End Sub
Private Sub lblCurrentLangSettings_Click()

End Sub


Private Sub optDropCapsNo_Click()
    cbHideVerse1 = False
    cbHideVerse1.Enabled = False
End Sub

Private Sub optDropCapsYes_Click()
        cbHideVerse1.Enabled = True
End Sub

Private Sub txtLanguageNameIni_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    txtLanguageNameIni.SelStart = 0
    txtLanguageNameIni.SelLength = Len(txtLanguageNameIni.Text)
End Sub

Private Sub UserForm_Initialize()
    If LangData.NewLanguage = True Then
        lblLangSettings.Caption = _
            "If you wish to add settings for a new language project to your settings file," & _
            " please fill in the blanks and then click 'OK'."
        AddAsNewLanguage = True
    Else
        lblLangSettings.Caption = _
            "These are the language project settings for your current Scripture portion."
        AddAsNewLanguage = False
    End If
    
    ChangeFontComboBox
End Sub
