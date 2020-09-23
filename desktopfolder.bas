Attribute VB_Name = "Module3"
Option Explicit
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Global Const LVM_FIRST = &H1000
Global Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Global Const LVSCW_AUTOSIZE = -1
Global index
Public times As Boolean
Public Sub GetImageFiles()
Form1.ListView1.ListItems.Clear
Dim imtx2 As ListItem
Dim filenan As String
Dim bob()
Dim freak As Integer
freak = 0
If Dir("C:\Yar Desk\Imagess.lsd") = "" Then
    MsgBox ("No image files have been created, run the wizard."), vbExclamation
    Form5.Show
    Unload Form1
Else
Form1.Show
Open ("C:\Yar Desk\Imagess.lsd") For Input As #1
Do While Not EOF(1)
    Line Input #1, filenan
    freak = freak + 1
    
    Set imtx2 = Form1.ListView1.ListItems.Add(, , filenan)
Loop
Close #1
LockWindowUpdate Form1.ListView1.hWnd
SendMessage Form1.ListView1.hWnd, LVM_SETCOLUMNWIDTH, index, LVSCW_AUTOSIZE
    LockWindowUpdate 0
End If
End Sub

Public Sub GetMediaFiles()
Form1.ListView1.ListItems.Clear
Form5.Label4.Caption = "Creating the media files"
Dim imtx2 As ListItem
Dim filenan As String
Dim bob()
Dim freak As Integer
freak = 0
If Dir("C:\Yar Desk\media.lsd") = "" Then
    MsgBox ("No imedia files have been created, run the wizard."), vbExclamation
    Form5.Show
    Unload Form1
Else
Open ("C:\Yar Desk\media.lsd") For Input As #1
Do While Not EOF(1)
    Line Input #1, filenan
    freak = freak + 1
    
    Set imtx2 = Form1.ListView1.ListItems.Add(, , filenan)
Loop
Close #1
LockWindowUpdate Form1.ListView1.hWnd
SendMessage Form1.ListView1.hWnd, LVM_SETCOLUMNWIDTH, index, LVSCW_AUTOSIZE
    LockWindowUpdate 0
End If
Form1.Show
End Sub
Public Sub GetAppsFiles()
Form1.ListView1.ListItems.Clear
Form5.Label4.Caption = "Creating the app files"
Dim imtx2 As ListItem
Dim filenan As String
Dim bob()
Dim freak As Integer
freak = 0
If Dir("C:\Yar Desk\Apps.lsd") = "" Then
    MsgBox ("No application files have been created, run the wizard."), vbExclamation
    Form5.Show
    Unload Form1
Else
Open ("C:\Yar Desk\Apps.lsd") For Input As #1
Do While Not EOF(1)
    Line Input #1, filenan
    freak = freak + 1
    
    Set imtx2 = Form1.ListView1.ListItems.Add(, , filenan)
Loop
Close #1
LockWindowUpdate Form1.ListView1.hWnd
SendMessage Form1.ListView1.hWnd, LVM_SETCOLUMNWIDTH, index, LVSCW_AUTOSIZE
    LockWindowUpdate 0
End If
Form1.Show
End Sub
Public Sub GetTextFiles()
Form1.ListView1.ListItems.Clear

Dim imtx2 As ListItem
Dim filenan As String
Dim bob()
Dim freak As Integer
freak = 0
If Dir("C:\Yar Desk\text.lsd") = "" Then
    MsgBox ("No text files have been created, run the wizard."), vbInformation
    Form5.Show
    Unload Form1
Else
Form1.Show
Open ("C:\Yar Desk\text.lsd") For Input As #1
Do While Not EOF(1)
    Line Input #1, filenan
    freak = freak + 1
    Set imtx2 = Form1.ListView1.ListItems.Add(, , filenan)
Loop
LockWindowUpdate Form1.ListView1.hWnd
SendMessage Form1.ListView1.hWnd, LVM_SETCOLUMNWIDTH, index, LVSCW_AUTOSIZE
    LockWindowUpdate 0
Close #1
End If
End Sub
Public Sub GetOtherFiles()
Form1.ListView1.ListItems.Clear
Dim imtx2 As ListItem
Dim filenan As String
Dim bob()
Dim freak As Integer
freak = 0
If Dir("C:\Yar Desk\other.lsd") = "" Then
    MsgBox ("All files have not been created, run the wizard."), vbInformation
    Form5.Show
    Unload Form1
Else
Form1.Show
Open ("C:\Yar Desk\other.lsd") For Input As #1
Do While Not EOF(1)
    Line Input #1, filenan
    freak = freak + 1
        Set imtx2 = Form1.ListView1.ListItems.Add(, , filenan)

Loop
Close #1
Form1.Show
LockWindowUpdate Form1.ListView1.hWnd
SendMessage Form1.ListView1.hWnd, LVM_SETCOLUMNWIDTH, index, LVSCW_AUTOSIZE
    LockWindowUpdate 0
End If
End Sub

Public Sub KillFolderTree(Murder As String)
    Dim MurderDir As String
    MurderDir = Dir(Murder & "\*.*", vbDirectory)
    Do While MurderDir <> ""
        If MurderDir <> "." And MurderDir <> ".." Then
            If (GetAttr(Murder & "\" & MurderDir) And vbDirectory) = vbDirectory Then
                Call KillFolderTree(Murder & "\" & MurderDir)
                MurderDir = Dir(Murder & "\*.*", vbDirectory)
            Else
                On Error Resume Next
                Kill Murder & "\" & MurderDir
                On Error GoTo 0
                MurderDir = Dir
            End If
        Else
            MurderDir = Dir
        End If
    Loop
    On Error Resume Next
    RmDir Murder
End Sub

Public Function UnloadAll(FormA As Form, formB As Form, formC As Form, formD As Form, formE As Form)
Unload FormA
Unload formB
Unload formD
Unload formD
Unload formE
End
End Function





