VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbAskForHeaderAndTitle 
   Caption         =   "Information for ID and Headers"
   ClientHeight    =   5940
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7428
   OleObjectBlob   =   "dbAskForHeaderAndTitle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbAskForHeaderAndTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpVersion As String

Option Explicit
Private Sub cmdOK_Click()

'set book code
'Dim Text As String

If txtBookCode.TextLength = 3 Then
'   txtBookCode = txtBookCode & " "
'   With txtBookCode
'      With Selection.Find
'         .MatchWildcards = True
'         .Text = "([1-4a-zA-Z]{3})"
'      End With
'      Selection.Find.Execute
'      If Selection.Find.Found Then
         VariableSet "idBookCode", UCase(txtBookCode)
Else: VariableSet "idBookCode", "BUK"
'      End If
'   End With
End If

'set version info

If optFrstDraft = True Then
   tmpVersion = "[First draft]"
ElseIf optAdvCk = True Then
   tmpVersion = "[Advisor check done]"
ElseIf optVlgTsT = True Then
   tmpVersion = "[Testing done]"
ElseIf optOther = True And txtOther <> "" Then
   tmpVersion = "[" & txtOther & "]"
Else: tmpVersion = ""
End If

VariableSet "idVersionInfo", tmpVersion

'set header book name
'If txtHeaderName = "" Then txtHeaderName = " "
VariableSet "idBookName", txtHeaderName
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub txtOther_Change()
optOther = True
End Sub

