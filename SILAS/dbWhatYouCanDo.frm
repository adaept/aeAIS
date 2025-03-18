VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbWhatYouCanDo 
   Caption         =   "Next Step"
   ClientHeight    =   4908
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6960
   OleObjectBlob   =   "dbWhatYouCanDo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbWhatYouCanDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdDontShow_Click()
Dim fs As Object
If IniFile = "" Then SetIniFile
Set fs = CreateObject("Scripting.FileSystemObject")
System.PrivateProfileString( _
    FileName:=IniFile, Section:="Messages", _
    Key:="DontShowInstructions") = "True"
Unload Me
End Sub
Private Sub cmdOK_Click()
Unload Me
End Sub

