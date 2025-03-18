VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbBookletStyle 
   Caption         =   "Booklet Settings"
   ClientHeight    =   7596
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6948
   OleObjectBlob   =   "dbBookletStyle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbBookletStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iniPassage As String
Dim iniVlgName As String
Dim iniJustified As String
Dim iniCutFolded As String

Dim fs As Object

Option Explicit
Private Sub cmdCancelBooklet_Click()
    Dim Result As Integer

    CancelFormatting = True
    Result = MsgBox( _
        Prompt:="Are you sure you want to cancel this formatting job?", _
        Buttons:=vbYesNo + vbDefaultButton2)

    If Result = vbNo Then
        MsgBox "Please choose a style for your booklet, then click 'OK'."
        'stay here while they rethink the BookletStyle box
    Else
         CancelFormatting = True
         Unload Me
    End If
End Sub
Private Sub cmdOKBooklet_Click()
If IniFile = "" Then SetIniFile
Set fs = CreateObject("Scripting.FileSystemObject")

' booklet style
If optCutPages.Value = True Then
    VariableSet "BookletStyle", "cut"
ElseIf optFolded.Value = True Then
    VariableSet "BookletStyle", "folded"
End If

' margin trim
If optVillageMargins.Value = True Then
    VariableSet "MarginTrim", "no"
ElseIf optPrintshopMargins.Value = True Then
    VariableSet "MarginTrim", "yes"
End If

' text justified
If optJustifiedNo.Value = True Then
    VariableSet "Justified", "no"
ElseIf optJustifiedYes.Value = True Then
    VariableSet "Justified", "yes"
End If

' passage name
If txtPassage.Text = "" Then txtPassage.Text = " "
VariableSet "ScripturePassage", txtPassage.Text

' village name
If txtVillageName.Text = "" Then txtVillageName.Text = " "
VariableSet "VillageName", txtVillageName.Text

System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ScripturePassage") = ActiveDocument.Variables("ScripturePassage")
System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="VillageName") = ActiveDocument.Variables("VillageName")
System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="Justified") = ActiveDocument.Variables("Justified")
System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="BookletStyle") = ActiveDocument.Variables("BookletStyle")

Unload Me
End Sub
Private Sub cmdInstallBookletMacros_Click()
    Dim Response As Integer
    
    ' Check to see if this version of the booklet macros has been installed.
    '
    Dim MyInstalledBookletVersion As String
    
    On Error GoTo -1: On Error GoTo TryAndInstallBookletMacros
    MyInstalledBookletVersion = Exists_InstalledBookletVersion
    
    If InternationalCDbl(MyInstalledBookletVersion) < InternationalCDbl(InstallingBookletVersion) Then
        GoTo TryAndInstallBookletMacros
    Else
        Response = MsgBox( _
            Title:="Booklet macros up-to-date", _
            Prompt:="The booklet macros, version " & MyInstalledBookletVersion & _
                ", have been installed already." & vbCr & vbCr & _
                "If you like, I'll open the file that describes how to use the booklet macros." & _
                vbCr & "If you want to install the macros again, you can click the 'Install' button." & _
                vbCr & vbCr & _
                "Would you like me to open that file?", _
            Buttons:=vbYesNoCancel)
        
        If Response = vbYes Then
            GoTo TryAndInstallBookletMacros
        Else
            On Error GoTo 0
            Exit Sub
        End If
    End If
    ' Control should never come here.
    
TryAndInstallBookletMacros:
    On Error GoTo 0
    
    If Val(Application.Version) >= Wd2000 Then
        CancelFormatting = True
        Me.Hide
        BookletMacroInstall
    Else
        MsgBox _
            Title:="Version of Word too old", _
            Prompt:="Sorry, this version of the booklet macros doesn't work in Word 97."
    End If
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim MyInstalledBookletVersion As String
    
'first check Inifile for VillageName & ScripturePassage
' following statements are in (General) for this module (KB 06-11-24)
'Dim iniPassage As String
'Dim iniVlgName As String
If IniFile = "" Then SetIniFile
Set fs = CreateObject("Scripting.FileSystemObject")

iniPassage = System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ScripturePassage")
iniVlgName = System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="VillageName")
iniJustified = System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="Justified")
iniCutFolded = System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="BookletStyle")

If VariableGet("ScripturePassage") <> "" Then
    txtPassage = ActiveDocument.Variables("ScripturePassage")
ElseIf iniPassage <> "" And iniPassage <> " " Then
    txtPassage = iniPassage
Else: txtPassage = "Genesis 1-9"
End If

If VariableGet("VillageName") <> "" Then
    txtVillageName = ActiveDocument.Variables("VillageName")
ElseIf iniVlgName <> "" And iniVlgName <> " " Then
    txtVillageName = iniVlgName
Else: txtVillageName = "Hauna Village"
End If

If VariableGet("BookletStyle") = "" Then VariableSet "BookletStyle", iniCutFolded
If VariableGet("BookletStyle") = "cut" Then
    optCutPages.Value = True
    optFolded.Value = False
Else:
    optFolded.Value = True
    optCutPages.Value = False
End If

If VariableGet("MarginTrim") = "yes" Then
    optVillageMargins.Value = False
    optPrintshopMargins.Value = True
Else:
    optPrintshopMargins.Value = False
    optVillageMargins.Value = True
End If

If VariableGet("Justified") = "" Then VariableSet "Justified", iniJustified
If VariableGet("Justified") = "no" Then
    optJustifiedYes.Value = False
    optJustifiedNo.Value = True
Else:
    optJustifiedNo.Value = False
    optJustifiedYes.Value = True
End If

    On Error GoTo -1: On Error GoTo EndOfSub
    MyInstalledBookletVersion = Exists_InstalledBookletVersion
    
    If InternationalCDbl(MyInstalledBookletVersion) < InternationalCDbl(InstallingBookletVersion) Then
         With cmdInstallBookletMacros
            .Enabled = True
            .Caption = "Install current version of " & vbCr & _
               "Print Booklet Macros " & vbCr & _
               "to help you print folded booklets"
         End With
    Else
         With cmdInstallBookletMacros
            .Enabled = False
            .Caption = "Print Booklet Macro is installed. " & vbCr & _
               "See Scripture menu item: " & vbCr & _
               "'Read about printing booklets ...'"
         End With
    End If

EndOfSub:
    On Error GoTo 0
End Sub
