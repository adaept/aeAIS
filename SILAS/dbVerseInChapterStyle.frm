VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbVerseInChapterStyle 
   Caption         =   "Found Verse Number in a Chapter Paragraph"
   ClientHeight    =   3828
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7068
   OleObjectBlob   =   "dbVerseInChapterStyle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbVerseInChapterStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdIgnore_Click()
cmdOK = False
VariableSet "ChapterStyle", "continue"
Unload Me
End Sub
Private Sub cmdOK_Click()
   'cmdOK = True
VariableSet "ChapterStyle", "error"
Unload Me
End Sub

