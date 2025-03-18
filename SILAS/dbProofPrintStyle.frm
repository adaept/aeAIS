VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbProofPrintStyle 
   Caption         =   "Proof Printout Settings"
   ClientHeight    =   6912
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5880
   OleObjectBlob   =   "dbProofPrintStyle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbProofPrintStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iniProofLineSp As String
Dim iniProofCkgTable As String
Dim iniProofBacked As String
Dim iniProofRtMargin As String

Option Explicit
Private Sub cmdCancelPrintout_Click()
Dim Result As Integer

    If CancelFormatting = False Then
        Result = MsgBox("Are you sure you want to cancel this formatting job?", vbYesNo + vbDefaultButton2)

        If Result = vbNo Then
            MsgBox "Please choose settings for your printout, then click 'OK'."
            'stay here while they rethink the ProofPrintStyle box
            Exit Sub
        Else
           CancelFormatting = True
        End If
    End If
    
    Unload Me
End Sub
Private Sub cmdOK_Click()
Dim fs As Object

If IniFile = "" Then SetIniFile
Set fs = CreateObject("Scripting.FileSystemObject")

' line spacing
If optSglSp.Value = True Then
    VariableSet "ProofLineSp", "single"
ElseIf opt15sp.Value = True Then
    VariableSet "ProofLineSp", "oneandhalf"
Else: VariableSet "ProofLineSp", "double"
End If

' footers
If optNoTables.Value = True Then
    VariableSet "ProofCkgTable", "False"
ElseIf optTables.Value = True Then
    VariableSet "ProofCkgTable", "True"
End If

' backed pages
If optOneSided.Value = True Then
    VariableSet "ProofBacked", "no"
ElseIf optBacked.Value = True Then
    VariableSet "ProofBacked", "yes"
End If

' right margin
If optMarginExtra.Value = True Then
    VariableSet "ProofRtMargin", "extrawide"
ElseIf optMarginWide.Value = True Then
    VariableSet "ProofRtMargin", "wide"
Else: VariableSet "ProofRtMargin", "even"
End If


System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ProofLineSpacing") = ActiveDocument.Variables("ProofLineSp")
System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ProofCheckingTable") = ActiveDocument.Variables("ProofCkgTable")
System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ProofBackedPages") = ActiveDocument.Variables("ProofBacked")
System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ProofRightMargin") = ActiveDocument.Variables("ProofRtMargin")

Unload Me
End Sub
Private Sub cmdTranslateTable_Click()
' Report that table is to be edited
' and close this dialog box.
'
    EditCheckingTable = True
    CancelFormatting = True
    cmdCancelPrintout_Click
End Sub

Private Sub UserForm_Initialize()

'first check Inifile for VillageName & ScripturePassage
' following statements are in (General) for this module (KB 06-11-24)
'Dim iniPassage As String
'Dim iniVlgName As String
Dim fs As Object

If IniFile = "" Then SetIniFile
Set fs = CreateObject("Scripting.FileSystemObject")

iniProofLineSp = System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ProofLineSpacing")
iniProofCkgTable = System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ProofCheckingTable")
iniProofBacked = System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ProofBackedPages")
iniProofRtMargin = System.PrivateProfileString( _
    FileName:=IniFile, Section:="LastBookletSettings", _
    Key:="ProofRightMargin")

' Set initial values for radio buttons
' line spacing
' If VariableGet("ProofLineSp") = "" Then VariableSet "ProofLineSp", iniProofLineSp
If VariableGet("ProofLineSp") = "" Then _
      ActiveDocument.Variables("ProofLineSp") = iniProofLineSp

If VariableGet("ProofLineSp") = "single" Then
    optSglSp.Value = True
    opt15sp.Value = False
    optDblSp.Value = False
    
ElseIf VariableGet("ProofLineSp") = "oneandhalf" Then
    optSglSp.Value = False
    opt15sp.Value = True
    optDblSp.Value = False
    
Else: optSglSp.Value = False
    opt15sp.Value = False
    optDblSp.Value = True
End If

' footers
If VariableGet("ProofCkgTable") = "" Then _
      ActiveDocument.Variables("ProofCkgTable") = iniProofCkgTable

If VariableGet("ProofCkgTable") = "False" Then
    optTables.Value = False
    optNoTables.Value = True
    
Else: optTables.Value = True
    optNoTables.Value = False
End If

' backed pages
If VariableGet("ProofBacked") = "" Then _
      ActiveDocument.Variables("ProofBacked") = iniProofBacked

If VariableGet("ProofBacked") = "yes" Then
    optBacked.Value = True
    optOneSided.Value = False
    
Else: optBacked.Value = False
    optOneSided.Value = True
End If

' right margin
If VariableGet("ProofRtMargin") = "" Then _
      ActiveDocument.Variables("ProofRtMargin") = iniProofRtMargin

If VariableGet("ProofRtMargin") = "extrawide" Then
    optMarginEven.Value = False
    optMarginWide.Value = False
    optMarginExtra.Value = True
    
ElseIf VariableGet("ProofRtMargin") = "wide" Then
    optMarginEven.Value = False
    optMarginWide.Value = True
    optMarginExtra.Value = False
  
Else: optMarginEven.Value = True
    optMarginWide.Value = False
    optMarginExtra.Value = False
End If

End Sub
