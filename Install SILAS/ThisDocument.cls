VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Code by Kim Blewett with portions from Mark Penny's Word Booklet Macros install routine, 28 Feb 2006
' Comments and suggestions for improvement are welcome!
' send them to: kblewett@gmail.com
' Last updated by KB on 28-2-06

Dim Source$
Dim InstallUninstallError As Boolean
Dim NormalDotFullName$
Private Sub InstallTemplate_2_Click()
    InstallScrTemplate.Main
End Sub
Private Sub Document_Open()
    Source$ = ActiveDocument.FullName
    NormalDotFullName$ = ActiveDocument.AttachedTemplate.FullName
End Sub
Private Sub InstallButton_Click()
    InstallScrTemplate.Main
End Sub
Private Sub UninstallButton_Click()
'
' Uninstalling macros code

    If MsgBox("Uninstalling macros for " & Application.Name & " (ver " & Application.Version & ")" & vbCrLf & _
        "in " & NormalDotFullName$, vbOKCancel, "Uninstall Macros") = vbCancel Then
            Exit Sub
    End If

    ActiveDocument.UpdateStylesOnOpen = False
    
    If InstallObjects(False) = True Then  ' go UNinstall (remove) them all
        MsgBox "The macros have been uninstalled.", vbInformation, "Uninstall Macros"
    Else
        MsgBox "There were one or more errors encountered while uninstalling the macros." & vbCrLf & _
        "Have you already manually removed some macro modules?" & vbCrLf & vbCrLf & _
        "You may need to contact technical support.", vbInformation, "Uninstall Macros"
    End If

End Sub
Function InstallObjects(ChangeMode As Boolean) As Boolean

    'wdOrganizerObject types: 0 = Styles
    '                         1 = AutoText
    '                         2 = CommandBars
    '                         3 = ProjectItems (macro modules & forms)
    
    ' Feel free to add more (or remove) objects here as needed.
    
    Modify ChangeMode, "Booklet Macros", 2
    Modify ChangeMode, "frmPrintBooklet", 3
    Modify ChangeMode, "Common", 3
    Modify ChangeMode, "ListWinPrinters", 3
    
    InstallObjects = Not InstallUninstallError
    
End Function
Private Sub Modify(AddObject As Boolean, ObjectName As String, ObjectType As Integer)
    
    On Error GoTo UnknownError ' trap errors here
    
    If AddObject Then ' install
        'first delete exising object, before copying the new object
        On Error Resume Next ' in case it isn't found, just continue
        Application.OrganizerDelete Source:=NormalDotFullName$, _
                                    Name:=ObjectName, _
                                    Object:=ObjectType
        On Error GoTo 0 ' if there's a problem we need to do something
        Application.OrganizerCopy Source:=Source$, _
                                  Destination:=NormalDotFullName$, _
                                  Name:=ObjectName, _
                                  Object:=ObjectType
    Else ' uninstall
        Application.OrganizerDelete Source:=NormalDotFullName$, _
                                    Name:=ObjectName, _
                                    Object:=ObjectType
    End If
    InstallUninstallError = False
    On Error GoTo 0 ' release error trap
    Exit Sub
    
UnknownError:
    If AddObject Then
        MsgBox "An unknown error has occurred while installing the " & ObjectName & " module." & vbCrLf & _
        "First uninstall the macros and then attempt to install them again.", vbExclamation, "Install Macros"
    Else
        MsgBox "An unknown error has occurred while attempting to uninstall the " & ObjectName & " module." & vbCrLf & _
        "You may have removed this module already.", vbExclamation, "Uninstall Macros"
    End If
    InstallUninstallError = True

End Sub

