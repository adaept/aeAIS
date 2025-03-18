VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbMsgBoxWithTimeout 
   Caption         =   "UserForm1"
   ClientHeight    =   3216
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5664
   OleObjectBlob   =   "dbMsgBoxWithTimeout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbMsgBoxWithTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub dbMsgBoxWithTimeout_Initialize( _
    Optional ByVal Title As String, Optional ByVal Text As String, _
    Optional ByVal DelayMsg As String, Optional ByVal TimeOut As Integer)
    ' Display a message for a specified time,
    ' with an OK button so the user can close it earlier.
    '
    Dim myText As String
    Dim myDelayMsg As String
    
    If Title = "" Then
        dbMsgBoxWithTimeout.Caption = "SILAS"
    Else
        dbMsgBoxWithTimeout.Caption = Title
    End If
    
    If Text = "" Then
        dbMsgBoxWithTimeout.strText = "The formatting task has been done."
    Else
        dbMsgBoxWithTimeout.strText = Text
    End If
    
    If DelayMsg = "" Then
        dbMsgBoxWithTimeout.strDelayMsg = "This dialog will appear for 5 seconds." & _
            vbCrLf & "You can close it sooner by operating the OK button."
    Else
        dbMsgBoxWithTimeout.strDelayMsg = DelayMsg
    End If
    
    If TimeOut = 0 Then TimeOut = 5

    TimeOut = TimeOut * 1000          ' milliseconds
    
    dbMsgBoxWithTimeout.Show
    
End Sub
Private Sub DelayMsg_Click()

End Sub
Private Sub OKbutton_Click()
    Unload Me
End Sub

