VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Yar Systimization Wizard"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "desktopwizard.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   4920
      TabIndex        =   18
      Top             =   5640
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFC0C0&
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Text            =   "desktopwizard.frx":1042
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Text            =   "desktopwizard.frx":11AC
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFC0C0&
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   8295
      Begin VB.Timer Timer2 
         Left            =   6480
         Top             =   240
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   615
         Left            =   7560
         Picture         =   "desktopwizard.frx":11B2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   615
         Left            =   7560
         Picture         =   "desktopwizard.frx":21F4
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFC0C0&
         Height          =   2010
         Left            =   2280
         TabIndex        =   10
         Top             =   480
         Width           =   1680
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00FFC0C0&
         Height          =   1890
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2100
      End
      Begin VB.DirListBox Dir2 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   3480
         TabIndex        =   8
         Top             =   2520
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2115
      End
      Begin VB.ListBox ListAll 
         BackColor       =   &H00FFC0C0&
         Height          =   2010
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   3315
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H00FFC0C0&
         Height          =   480
         Left            =   5640
         Pattern         =   "All Image Files|*.jpg;*.bmp;*.gif;*.ico;*.cur"
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   7560
         Picture         =   "desktopwizard.frx":3236
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Folders"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Next >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   3600
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   7440
      Top             =   840
   End
   Begin VB.Image Image2 
      Height          =   1725
      Left            =   3000
      Picture         =   "desktopwizard.frx":3F00
      Top             =   960
      Width           =   2580
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   $"desktopwizard.frx":9510
      ForeColor       =   &H00FFC0C0&
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status: "
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5640
      Width           =   8295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Wizard"
      BeginProperty Font 
         Name            =   "911 Porscha"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdd_Click()
Label4.Caption = "Loading all Files..."
File1.Pattern = "All Image Files|*.jpg;*.bmp;*.gif;*.ico;*.cur;*.pcx;*.psp;*.pctp;*.png;*.tif"
Call GiveListOfFiles
Call ConvertFiles
Call imagess
Text1.Text = ""
ListAll.Clear
Call Savemediafiles
End Sub
Sub Savemediafiles()
Label4.Caption = "Loading all Files..."
File1.Pattern = "All Media Files|*.mp1;*.mp2;*.mp3*.mpeg;*.avi;*.asf;*.wmv;*.asx;*.vax;*.mid;*.midi;*.rmi;*.rma"
Call GiveListOfFiles
Call ConvertFiles
Call medias
Text1.Text = ""
List1.Clear
Call SaveTextfiles
End Sub
Sub SaveTextfiles()
Label4.Caption = "Loading all Files..."
File1.Pattern = "All Media Files|*.txt;*.doc;*.wri;*.rtf"
Call GiveListOfFiles
Call ConvertFiles
Call Texts
Text1.Text = ""
List1.Clear
Call SaveAppFile
End Sub
Sub SaveAppFile()
Label4.Caption = "Loading all Files..."
File1.Pattern = "*.exe"
Call GiveListOfFiles
Call ConvertFiles
Call Apps
Text1.Text = ""
List1.Clear
Call SaveAllFile
End Sub
Sub SaveAllFile()
File1.Pattern = "*.*"
Call GiveListOfFiles
Call ConvertFiles
Call others
Form1.Show
Form5.Hide
Unload Form5
End Sub

Sub ConvertFiles()
Label4.Caption = "Converting Files..."
Command1.Enabled = False
Command2.Enabled = True
Dim a As Long
    Dim b As String
        For a = 0 To (ListAll.ListCount - 1)
            b = b & ListAll.list(a) & vbCrLf
            Caption = "Processing File Number: " & a + 1
                Text1.Text = b
Next

End Sub
Private Sub Form_Load()
Dim YarFol
  Set YarFol = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
On Error GoTo errorfolder:
MyFoldersPath$ = ("C:\")
    If Not Right(MyFoldersPath$, 1) = "\" Then
       MyFoldersPPath$ = MyFoldersPath$ & "\" & "Yar Desk"
       YarFol.CreateFolder "C:\" & "\" & "Yar Desk"
       Else
       YarFol.CreateFolder "C:\Yar Desk"
       End If
errorfolder:

Exit Sub
MsgBox ("Please read instructions found in the upper left hand corner"), vbInformation
'Form1.Hide
If Dir("C:\Yar Desk\wizard.lsd") = "" Then
Open "C:\Yar Desk\wizard.lsd" For Output As #1
Close #1
End If
Kill ("C:\Yar Desk\wizard.lsd")
End Sub

Private Sub Timer1_Timer()
ChangeColor Label1
End Sub

