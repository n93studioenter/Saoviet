Attribute VB_Name = "ModGuest"
Option Explicit

Public Wspace As Workspace
Public DBGuest As Database

Public pCurdir As String

Public Sub Main()
    If App.PrevInstance <> 0 Then
        MsgBox "Ch­¬ng tr×nh ®ang ho¹t ®éng !", vbExclamation, App.ProductName
        End
    End If
           
    pCurdir = App.Path
    If Right(pCurdir, 1) <> "\" Then pCurdir = pCurdir + "\"
    
    IniPath = Dir(App.Path, vbDirectory)
    
    FOpenDB.Show vbModal
    
    End
End Sub

Public Sub OpenDB1(pdatapath As String)
    If Dir(pdatapath) = "" Then Exit Sub
   Set Wspace = DBEngine.CreateWorkspace(CStr(Time), "Admin", "", dbUseJet)
    Workspaces.Append Wspace
    On Error Resume Next
    Set DBGuest = Wspace.OpenDatabase(pdatapath, True, False, ";PWD=" + Chr(8) + Chr(13) + Chr(27))
    DBGuest.NewPassword Chr(8) + Chr(13) + Chr(27), "unlock"
    Set DBGuest = Wspace.OpenDatabase(pdatapath, True, False, ";PWD=shark")
    DBGuest.NewPassword "shark", "unlock"
    On Error GoTo 0
    DBGuest.Close
KT:
    Wspace.Close
End Sub

