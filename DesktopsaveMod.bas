Attribute VB_Name = "Module2"
Public Sub imagess()
Dim filena As String
'images = True
'Other = False
'app = False
'Text = False
'media = False
Form5.Label4.Caption = "Saving data for images."
   
   filena = "C:\Yar Desk\imagess.lsd"
Open filena For Output As #1
    DoEvents
    On Error Resume Next
    Dim items, i
      items = ListAll.ListCount
    For i = 1 To items
        DoEvents
        ListAll.ListIndex = -1 + i
        Form5.Text1 = Text1 & vbCrLf & Form5.ListAll.Text
    Next i
     Print #1, Form5.Text1.Text
Close #1
MsgBox ("Images are done moving to media when ready."), vbInformation
Form5.ListAll.Clear
Form5.Text1.Text = ""
Form5.ProgressBar1.Value = 20
End Sub
Public Sub others()
Dim filena As String
'images = True
'Other = False
'app = False
'Text = False
'media = False
Form5.Label4.Caption = "Saving data for all."
   filena = "C:\Yar Desk\other.lsd"
Open filena For Output As #1
    DoEvents
    On Error Resume Next
    Dim items, i
       items = ListAll.ListCount
    For i = 1 To items
        DoEvents
        ListAll.ListIndex = -1 + i
        Form5.Text1 = Text1 & vbCrLf & Form5.ListAll.Text
    Next i
     Print #1, Form5.Text1.Text
Close #1
Form5.ProgressBar1.Value = 100
MsgBox ("All Files are done. Pick a picture to tile on desktop."), vbInformation
Form1.CommonDialog1.FileName = ""
    Form1.CommonDialog1.ShowOpen
    openfilename = Form1.CommonDialog1.FileName
    Form1.Image6.Picture = LoadPicture(openfilename)
     Form1.Label14.Caption = Form1.CommonDialog1.FileName
    Call Form1.saving
    Call Form1.TileMe2
Form5.ListAll.Clear
Form5.Text1.Text = ""
Form1.Show
Unload Form5
End Sub

Public Sub Texts()
Form5.Show
Dim filena As String
'images = True
'Other = False
'app = False
'Text = False
'media = False
Form5.Label4.Caption = "Saving data for text."
   filena = "C:\Yar Desk\text.lsd"
Open filena For Output As #1
    DoEvents
    On Error Resume Next
    Dim items, i
        items = ListAll.ListCount
    For i = 1 To items
        DoEvents
        ListAll.ListIndex = -1 + i
        Form5.Text1 = Text1 & vbCrLf & Form5.ListAll.Text
    Next i
     Print #1, Form5.Text1.Text
Close #1
MsgBox ("Media files are done moving to Application files."), vbInformation
Form5.ListAll.Clear
Form5.Text1.Text = ""
Form5.ProgressBar1.Value = 60
End Sub

Public Sub Apps()
Form5.Show
Dim filena As String
'images = True
'Other = False
'app = False
'Text = False
'media = False
Form5.Label4.Caption = "Saving data for apps."
   filena = "C:\Yar Desk\Apps.lsd"
Open filena For Output As #1
    DoEvents
    On Error Resume Next
    Dim items, i
       items = ListAll.ListCount
    For i = 1 To items
        DoEvents
        ListAll.ListIndex = -1 + i
        Form5.Text1 = Text1 & vbCrLf & Form5.ListAll.Text
    Next i
     Print #1, Form5.Text1.Text
Close #1
MsgBox ("Apps are done moving to all files when ready."), vbInformation
Form5.ProgressBar1.Value = 80
End Sub

Public Sub medias()
Dim filena As String
'images = True
'Other = False
'app = False
'Text = False
'media = False
Form5.Label4.Caption = "Saving data for media."
   filena = "C:\Yar Desk\media.lsd"
Open filena For Output As #1
    DoEvents
    On Error Resume Next
    Dim items, i
        items = ListAll.ListCount
    For i = 1 To items
        DoEvents
        ListAll.ListIndex = -1 + i
        Form5.Text1 = Text1 & vbCrLf & Form5.ListAll.Text
    Next i
     Print #1, Form5.Text1.Text
Close #1
MsgBox ("Media files are done moving to text files when ready."), vbInformation
Form5.ListAll.Clear
Form5.Text1.Text = ""
Form5.ProgressBar1.Value = 40
End Sub
