VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbPaperSize 
   Caption         =   "Select Paper Size"
   ClientHeight    =   3564
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5712
   OleObjectBlob   =   "dbPaperSize.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbPaperSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOK_Click()
    VariableSet "NoLanguageSet", ""
    If VariableGet("WhenExported") = "" Then
        VariableSet "WhenExported", Format(Now, "dd-Mmm-yyyy") & " at " & Format(Now, "hh:mm")
    End If

    Call SetSize2
End Sub

Sub SetSize2()
'
'for A4
Dim dbPaperSize As Variant
Dim tmpCountry As String

    If IniFile = "" Then SetIniFile

    If optA4.Value = True Then
        dbPaperSize = "A4"
    End If
    
'for Letter
    If optLetter.Value = True Then
        dbPaperSize = "Letter"
    End If

    tmpCountry = VariableGet("LanguageCountry")
    VariableSet "DefaultPaperSize", dbPaperSize
    System.PrivateProfileString( _
        FileName:=IniFile, Section:="DefaultPaperSize", _
        Key:=tmpCountry) = dbPaperSize
    Me.Hide
    MsgBox "Default Paper Size for " & tmpCountry & " is " & dbPaperSize & "." & vbCrLf _
        & "To change this default size later on, open the menu 'ScrLanguages' and do the job called " & vbCrLf & _
        vbTab & Chr(34) & "Change Default Paper Size For Your Country." & Chr(34)
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim tmpPaperSize As String
    
    If IniFile = "" Then SetIniFile

    txtCountry.Text = VariableGet("LanguageCountry")
    tmpPaperSize = System.PrivateProfileString( _
        FileName:=IniFile, _
        Section:="DefaultPaperSize", _
        Key:=txtCountry.Text)
    
    Select Case tmpPaperSize
        Case "A4"
            optA4.Value = True
        Case "Letter"
            optLetter.Value = True
    End Select
End Sub


