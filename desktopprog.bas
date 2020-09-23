Attribute VB_Name = "Module1"
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public counter
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public i As Long
Public ani As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal index As Long) As Long
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Swapmouse, auto_arrange As Boolean
Const AC_SRC_OVER = &H0

Public Declare Function Pie Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public binmar1 As Boolean
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Global Const MINUTES = 15
    Public TimeOver
    Public MinUp
Public Const LB_SETHORIZONTALEXTENT = &H194
Public tempstring, tempstring2, filena, filenas, no As String
Dim start, lstindex
Public images As Boolean
Public other As Boolean
Public app As Boolean
Public Text As Boolean
Public media As Boolean
Public Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function
Public Sub wait()
Dim hol As Integer
For hol = 1 To 2243
DoEvents
Next hol
Exit Sub
End Sub
Public Sub wait2()
Dim hold As Single
hold = 0
For hold = 1 To 82000
DoEvents
Next hold
Exit Sub
End Sub

Public Function ChangeColor(label As label)
Dim t As Integer
For t = 255 To 0 Step -1
    DoEvents
    label.Refresh
    Call wait
    label.ForeColor = RGB(192, 192, 255)
    label.ForeColor = RGB(0, 0, t)
    label.Refresh
Next t
For u = 0 To 255
    DoEvents
    label.Refresh
    Call wait
    label.ForeColor = RGB(192, 192, 255)
    label.ForeColor = RGB(0, 0, u)
    label.Refresh
Next u
End Function
Public Function ChangeColor2(label As label)
Dim t As Integer
For t = 255 To 0 Step -1
    DoEvents
    label.Refresh
    Call wait
    label.ForeColor = RGB(192, 192, 255)
    label.ForeColor = RGB(t, 0, 0)
    label.Refresh
Next t
For u = 0 To 255
    DoEvents
    label.Refresh
    Call wait
    label.ForeColor = RGB(192, 192, 255)
    label.ForeColor = RGB(u, 0, 0)
    label.Refresh
Next u
End Function
Public Sub AddHScroll(list As listbox)
    Dim r As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
For r = 0 To list.ListCount - 1
        If Len(list.list(r)) > Len(list.list(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next r
    lngGreatestWidth = list.Parent.TextWidth(list.list(intGreatestLen) + Space(1))
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    SendMessage list.hWnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
End Sub
Public Sub RemoveListItem(listbox As listbox)
        Dim ListCount As Long
ListCount& = listbox.ListCount
Do While ListCount& > 0&
ListCount& = ListCount& - 1
If listbox.Selected(ListCount&) = True Then
listbox.RemoveItem (ListCount&)
End If
Loop
End Sub
Public Sub MurderList(list As listbox)
Dim gone As Integer
gone = 0
Do While gone < list.ListCount
        list.Text = list.list(gone)
    If list.ListIndex <> gone Then
        list.RemoveItem gone
    Else
        gone = gone + 1
    End If
Form5.Label4.Caption = "Deleting Duplicates..."

Loop
Call DoOperations
End Sub


Function FullPathName(FullName As String, CharL As Byte, Buffer As Byte)
    Buffer = Space(255)
    Ret = GetFullPathName(FullName, CharL, Buffer, "")
    Buffer = Left(Buffer, Ret)
End Function
Public Sub DoOperations()
filenas = "C:\Yar Desk\wizard.lsd"
On Error Resume Next
Open filenas For Input As #1
    On Error Resume Next
    Do Until EOF(1)
    DoEvents
    Line Input #1, filenas
    DoEvents
    tempstring = Dir(filenas & "*.*")
If tempstring <> "" Then
    tempstring2 = tempstring
End If
   tempstring = Dir
While Len(tempstring) > 0
DoEvents

tempstring2 = Dir$
DoEvents
Label4.Caption = "Load All files"
Wend
Loop
Close #1
End Sub
Public Function GetFileSize(FileName) As String
    On Error GoTo Gfserror
    Dim TempStr As String
    TempStr = FileLen(FileName)
    If TempStr >= "1024" Then
          TempStr = CCur(TempStr / 1024) & "KB"
    Else
        If TempStr >= "1048576" Then
               TempStr = CCur(TempStr / (1024 * 1024)) & "KB"
        Else
            TempStr = CCur(TempStr) & "B"
        End If
    End If
    GetFileSize = TempStr
    Exit Function
Gfserror:
    GetFileSize = "0B"
    Resume
End Function
Public Sub SetReadOnly(FileName As String)
    On Error Resume Next
    SetAttr FileName, vbReadOnly
End Sub
Public Function GetPath(ByVal fullfilepath As String) As String
    Dim l&, n%
    Dim char As String * 1
       GetPath = fullfilepath
    l = Len(fullfilepath)
    If l > 4 Then
        For n = l To 3 Step -1
            char = Mid$(fullfilepath, n, 1)
            Select Case char
                Case "\"
                GetPath = Left$(fullfilepath, n - 1)
                Exit For
            End Select
    Next
End If
End Function

Sub GiveListOfFiles()
Form5.cmdadd.Enabled = False
Form5.Command1.Enabled = True
Form5.List1.Clear
Form5.List1.AddItem (Form5.Dir1.Path)
2:
Call ListDir(Form5.Dir1.Path)
 Dim Count As Integer, Count2 As Integer, Num As Integer
 Dim tmpString As String
  Num% = Form5.ListAll.ListCount + 1
  For Count% = 0 To Form5.List1.ListCount - 1
   Form5.File1.Path = Form5.List1.list(Count%)
   For Count2% = 0 To Form5.File1.ListCount - 1
    DoEvents
    If UCase(Right(Form5.File1.list(Count2%), 4)) = "*" Then
           TempStr = (List1.list(Count%) & "\" & Form5.File1.list(Count2%))
            Do: DoEvents
       DoEvents
       If InStr(TempStr, "\") = False Then Exit Do
             TempStr = Mid(TempStr, InStr(TempStr, "\") + 1, Len(TempStr) - InStr(TempStr, "\") + 1)
      Loop
       Form5.ListAll.AddItem (Num% & ". " & TempStr)
       Num% = Num% + 1
    ElseIf UCase(Right(Form5.File1.list(Count2%), 4)) = "*" Then
      TempStr = (List1.list(Count%) & "\" & Form5.File1.list(Count2%))
      Form5.ListAll.AddItem Form5.File1.ListCount
      Do: DoEvents
       If InStr(TempStr, "\") = False Then Exit Do
       DoEvents
       TempStr = Mid(TempStr, InStr(TempStr, "\") + 1, Len(TempStr) - InStr(TempStr, "\") + 1)
      Loop
       Form5.ListAll.AddItem (Num% & ". " & TempStr)
      Num% = Num% + 1
    ElseIf UCase(Right(Form5.File1.list(Count2%), 4)) = "*" Then
      TempStr = (List1.list(Count%) & "\" & Form5.File1.list(Count2%))
      Form5.ListAll.AddItem Form5.File1.ListCount - 1
      Do: DoEvents
       If InStr(TempStr, "\") = False Then Exit Do
       TempStr = Mid(TempStr, InStr(TempStr, "\") + 1, Len(TempStr) - InStr(TempStr, "\") + 1)
      Loop
       Form5.ListAll.AddItem (Num% & ". " & TempStr)
      Num% = Num% + 1
    ElseIf UCase(Right(Form5.File1.list(Count2%), 4)) = "*" Then
      TempStr = (List1.list(Count%) & "\" & Form5.File1.list(Count2%))
      Form5.ListAll.AddItem Form5.File1.list(Count2%)
      Do: DoEvents
       If InStr(TempStr, "\") = False Then Exit Do
       TempStr = Mid(TempStr, InStr(TempStr, "\") + 1, Len(TempStr) - InStr(TempStr, "\") + 1)
      Loop
       Form5.ListAll.AddItem (Num% & ". " & TempStr)
      Num% = Num% + 1
    End If
   DoEvents
GetPath (Form5.File1.FileName)
   Next Count2%
  Call godir
  Next Count%
End Sub
Sub ListDir(Path)
Dim d(1000)
Dim lop, cnt, cur_depth
 Form5.Dir2.Path = Path
 For lop = 0 To Form5.Dir2.ListCount - 1
  d(cnt) = Form5.Dir2.list(lop)
  cnt = cnt + 1
 Next lop
 For lop = 0 To cnt - 1
  Form5.List1.AddItem (d(lop))
  cur_depth = cur_depth + 1
  ListDir d(lop)
 Next lop
  cur_depth = cur_depth - 1
End Sub
Sub godir()
1:

Dim Buffer As String, Ret As Long
Dim fullfilepath
        Buffer = Space(255)
For a = 0 To (Form5.File1.ListCount - 1)
Form5.File1.ListIndex = a
Ret = GetFullPathName(Form5.File1.FileName, 255, Buffer, "")
If Right(Form5.File1.Path, 1) = "\" Then
Form5.ListAll.AddItem Form5.File1.Path + Form5.File1.FileName
Else
Form5.ListAll.AddItem Form5.File1.Path + "\" + Form5.File1.FileName
End If
Next

End Sub
