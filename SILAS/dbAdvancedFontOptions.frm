VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbAdvancedFontOptions 
   Caption         =   "Advanced Font Options"
   ClientHeight    =   3312
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7812
   OleObjectBlob   =   "dbAdvancedFontOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbAdvancedFontOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub cmboFontArray_AfterUpdate()
'Dim tmpfont As String
'tmpfont = cmboFontArray.Text
End Sub

Private Sub cmboFontArray_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    cmboFontArray.SelStart = 0
    cmboFontArray.SelLength = Len(cmboFontArray.Value)

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'    If txtLineSpacing < 0.5 Or txtLineSpacing > 1.5 Then
'        MsgBox "Please enter a 2-decimal-place number between " & _
'            "0.50 and 1.49 into the 'Custom Line Spacing' box."
'    End If
    
    AssignVariables2
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ChangeFontComboBox
    ' cmboFontArray = tmpFont   ' What is this supposed to do?
    
'    If VariableGet("LineSpacing") <> "" Then
'        txtLineSpacing = VariableGet("LineSpacing")
'    Else: txtLineSpacing = "1.0"
'    End If
    
   If VariableGet("Leading") = "extratall" Then
      dbAdvancedFontOptions.optExtraTall = True
   ElseIf VariableGet("Leading") = "tall" Then
      dbAdvancedFontOptions.optTall = True
   Else:
      dbAdvancedFontOptions.optNormal = True
   End If

End Sub
Private Sub ChangeFontComboBox()
' written by Steve White, 4 Nov 05
Dim afont As Variant
Dim i As Long
Dim fonts() As String

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
Private Sub AssignVariables2()

    If cmboFontArray <> "" Then
        VariableSet "HeadingFont", dbAdvancedFontOptions.cmboFontArray
    Else:
        VariableSet "HeadingFont", VariableGet("LanguageFont")
    End If
        
'    If txtLineSpacing <> "" And txtLineSpacing <> " " Then
'        VariableSet "LineSpacing", dbAdvancedFontOptions.txtLineSpacing
'    Else: VariableSet "LineSpacing", "1.00"
'    End If

   If optTall = True Then
      VariableSet "Leading", "tall"
   ElseIf optExtraTall = True Then
      VariableSet "Leading", "extratall"
   Else: VariableSet "Leading", "normal"
   End If
   
   
End Sub
