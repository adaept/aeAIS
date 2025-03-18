VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbPTFont 
   Caption         =   "Language Project Font Setup"
   ClientHeight    =   5376
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7860
   OleObjectBlob   =   "dbPTFont.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbPTFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileSystem As Object

Option Explicit
Private Sub cmdOK_Click()
    Dim vLanguageNumber As String
    
    If IniFile = "" Then SetIniFile
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    
    vLanguageNumber = VariableGet("LanguageNumber")
    VariableSet "PTNoFontChange", cbDontAskFont.Value
    System.PrivateProfileString(FileName:=IniFile, _
        Section:=vLanguageNumber, _
        Key:="PTNoFontChange") = VariableGet("PTNoFontChange")

    Unload Me
End Sub

