VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   3600
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      Picture         =   "desktopweb.frx":0000
      ScaleHeight     =   270
      ScaleWidth      =   5295
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin VB.Image Image1 
         Height          =   120
         Left            =   4800
         Picture         =   "desktopweb.frx":4ACA
         Top             =   75
         Width           =   150
      End
      Begin VB.Image Image3 
         Height          =   120
         Left            =   4800
         Picture         =   "desktopweb.frx":4C0C
         Top             =   75
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image Image2 
         Height          =   120
         Left            =   5040
         Picture         =   "desktopweb.frx":4D4E
         Top             =   75
         Width           =   150
      End
      Begin VB.Image Image4 
         Height          =   120
         Left            =   5040
         Picture         =   "desktopweb.frx":4E90
         Top             =   75
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Desktop "
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   15
         Width           =   1335
      End
      Begin VB.Image Image5 
         Height          =   225
         Left            =   20
         Picture         =   "desktopweb.frx":4FD2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   4680
      Picture         =   "desktopweb.frx":6014
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5285
      ExtentX         =   9322
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image Image17 
      Height          =   420
      Left            =   720
      Picture         =   "desktopweb.frx":7056
      Top             =   3480
      Width           =   555
   End
   Begin VB.Image Image18 
      Height          =   420
      Left            =   720
      Picture         =   "desktopweb.frx":7CD8
      Top             =   3480
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image19 
      Height          =   420
      Left            =   120
      Picture         =   "desktopweb.frx":88EA
      Top             =   3480
      Width           =   540
   End
   Begin VB.Image Image20 
      Height          =   420
      Left            =   120
      Picture         =   "desktopweb.frx":94FC
      Top             =   3480
      Width           =   540
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub DragForm(Form As Form)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, &O2, &O0
End Sub

Private Sub Command1_Click()
WebBrowser1.Navigate (Text1.Text)
End Sub

Private Sub Command11_Click()
Form1.Visible = True
Form4.Visible = False
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate ("www.planet-source-code.com/vb")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image17.Visible = True
Image19.Visible = True
Image18.Visible = False
Image20.Visible = False

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image3.Visible = True
Image1.Visible = False
End Sub

Private Sub Image17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image18.Visible = True
Image17.Visible = False
End Sub

Private Sub Image18_Click()
 WebBrowser1.GoForward
 End Sub


Private Sub Image19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image20.Visible = True
Image19.Visible = False
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Visible = True
    Image2.Visible = False
End Sub

Private Sub Image20_Click()
WebBrowser1.GoBack
End Sub

Private Sub Image3_Click()
Form4.WindowState = 1
End Sub

Private Sub Image4_Click()
Unload Form4
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image1.Visible = True
Image2.Visible = True
Image4.Visible = False
End Sub
