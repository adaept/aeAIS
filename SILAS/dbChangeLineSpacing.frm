VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbChangeLineSpacing 
   Caption         =   "Change Line Spacing"
   ClientHeight    =   3120
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5484
   OleObjectBlob   =   "dbChangeLineSpacing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbChangeLineSpacing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentLineSpacing As String

Option Explicit
Private Sub CheckBox1_Click()

End Sub
Private Sub cmdCancelPrintout_Click()
    CancelFormatting = True
    Unload Me
End Sub
Private Sub cmdOK_Click()
' line spacing
    If optSglSp.Value = True Then
        VariableSet "CurrentLineSp", "single"
    ElseIf opt15sp.Value = True Then
        VariableSet "CurrentLineSp", "oneandhalf"
    ElseIf opt3sp.Value = True Then
        VariableSet "CurrentLineSp", "triple"
    Else: VariableSet "CurrentLineSp", "double"
    End If
    
    If cbAddWordSp Then  'set wide char.spacing, 4 sp between words, size 14
        ChangeLineSpacingPoints.SpaceBetweenWords
    End If
Unload Me
End Sub
Private Sub OptionButton3_Click()

End Sub
Private Sub CommandButton1_Click()

End Sub
Private Sub UserForm_Initialize()
    With ActiveDocument.Styles("_BodyText_Base").ParagraphFormat
       If .LineSpacingRule = 4 Or .LineSpacingRule = 0 Then _
       CurrentLineSpacing = "single"
    End With

    If CurrentLineSpacing = "single" Then
       optDblSp = True
    Else: optSglSp = True
    End If
End Sub

