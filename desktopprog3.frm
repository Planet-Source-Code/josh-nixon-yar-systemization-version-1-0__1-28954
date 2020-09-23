VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "desktopprog3.frx":0000
   ScaleHeight     =   1785
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1080
      Picture         =   "desktopprog3.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Cursor and Icon Files|*.cur;*.ico"
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   240
      Picture         =   "desktopprog3.frx":1594
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      Picture         =   "desktopprog3.frx":225E
      ScaleHeight     =   225
      ScaleWidth      =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cursor"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1455
      End
      Begin VB.Image Image8 
         Height          =   120
         Left            =   1560
         Picture         =   "desktopprog3.frx":37B8
         Top             =   45
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image Image7 
         Height          =   120
         Left            =   1560
         Picture         =   "desktopprog3.frx":38FA
         Top             =   45
         Width           =   150
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   1095
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   1095
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub DragForm(Form As Form)
ReleaseCapture
SendMessage hwnd, &HA1, &O2, &O0
End Sub
Private Sub Command1_Click()
Dim filen As String
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    filen = CommonDialog1.FileName
   Picture3.Picture = LoadPicture(filen)
    Label1.Caption = CommonDialog1.FileName
End Sub
Private Sub Command2_Click()
Form1.Label12.Caption = CommonDialog1.FileName
Form1.MouseIcon = Form3.Picture3.Picture
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image8.Visible = True
    Image7.Visible = False
End Sub

Private Sub Image8_Click()
Form1.Image14.Visible = False
Unload Form3
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.Visible = True
    Image8.Visible = False
    DragForm Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image8.Visible = False
DragForm Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.Visible = True
    Image8.Visible = False
End Sub
