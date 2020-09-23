VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      Picture         =   "desktopcolor.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   1800
      Begin VB.Image Image7 
         Height          =   120
         Left            =   1560
         Picture         =   "desktopcolor.frx":155A
         Top             =   45
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image Image1 
         Height          =   120
         Left            =   1560
         Picture         =   "desktopcolor.frx":169C
         Top             =   45
         Width           =   150
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "BG Color"
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
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   885
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   3465
      Width           =   255
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1005
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   3585
      Width           =   255
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2000
      Left            =   0
      MouseIcon       =   "desktopcolor.frx":17DE
      MousePointer    =   99  'Custom
      Picture         =   "desktopcolor.frx":20A8
      ScaleHeight     =   1995
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click Desktop Background"
      BeginProperty Font 
         Name            =   "911 Porscha"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Image Image8 
      Height          =   120
      Left            =   1560
      Picture         =   "desktopcolor.frx":81F2
      Top             =   45
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "911 Porscha"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1080
      Top             =   1440
      Width           =   600
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim currcolpic As String
Sub DragForm(Form As Form)
ReleaseCapture
SendMessage hwnd, &HA1, &O2, &O0
End Sub


Private Sub Image1_Click()
Unload Form4
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image1.Visible = False
End Sub
Private Sub Image7_Click()
Unload Form6
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = True
Image7.Visible = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = True
Image7.Visible = False
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    colpressed = True                          'this function tells the computer to
    currcolpic = Picture5.Point(X, Y)
    If Button = 1 Then
    Call set_forecolor
    Picture1.ToolTipText = r + G + b
    End If
    If Button = 2 Then
    Call set_backcolor
    Let r = RGB(r, G, b)
    Picture5.ToolTipText = r + G + b
    End If
    
End Sub
Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If colpressed = True Then
    'same as above but this is here
    currcolpic = Picture5.Point(X, Y)
    If Button = 1 Then
    Call set_forecolor
    End If
    If Button = 2 Then
    Call set_backcolor
    End If
    End If
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form1.Picture4.BackColor = Shape3.FillColor
    colpressed = False          'stops the selecting of the color when the user 'unclicks'
End Sub
Public Sub set_forecolor()
    
    Picture4.BackColor = currcolpic
    Shape3.FillColor = Picture4.BackColor
    Picture3.BackColor = Shape3.FillColor
End Sub

Public Sub set_backcolor()
    Shape3.FillColor = Picture6.BackColor
    Picture6.BackColor = currcolpic
End Sub
