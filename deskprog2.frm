VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   LinkTopic       =   "Form2"
   ScaleHeight     =   2700
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   50
      Picture         =   "deskprog2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Swap Mouse Buttons"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   665
      Picture         =   "deskprog2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Wizard"
      Top             =   960
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1200
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files|*.*"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Wav Files|*.wav"
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Picture         =   "deskprog2.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Picture         =   "deskprog2.frx":16FE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   655
      Picture         =   "deskprog2.frx":1C68
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1275
      Picture         =   "deskprog2.frx":2532
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Help"
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   50
      Picture         =   "deskprog2.frx":2DFC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Time"
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1275
      Picture         =   "deskprog2.frx":36C6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Start Enterainment"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   665
      Picture         =   "deskprog2.frx":4390
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Organize"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   50
      Picture         =   "deskprog2.frx":4C5A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Start up Sound"
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      Picture         =   "deskprog2.frx":5524
      ScaleHeight     =   225
      ScaleWidth      =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Properties"
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
         TabIndex        =   1
         Top             =   0
         Width           =   1575
      End
      Begin VB.Image Image8 
         Height          =   120
         Left            =   1560
         Picture         =   "deskprog2.frx":6A7E
         Top             =   45
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image Image7 
         Height          =   120
         Left            =   1560
         Picture         =   "deskprog2.frx":6BC0
         Top             =   45
         Width           =   150
      End
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "No"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Preview    Wav"
      BeginProperty Font 
         Name            =   "911 Porscha"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Command8.Visible = True
    Command9.Visible = True
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    Label2.Caption = CommonDialog1.FileName
    MsgBox ("Large Wav Files may take a while to load depending on computer speed."), vbInformation
    Call save
End Sub

Private Sub Command10_Click()

Label7.Caption = True
If Swapmouse = True Then
SwapMouseButton (0)
Swapmouse = False
ElseIf Swapmouse = False Then
SwapMouseButton (1)
Swapmouse = True
End If

End Sub

Private Sub Command11_Click()
End Sub

Private Sub Command2_Click()
If auto_arrange = True Then
    auto_arrange = False
ElseIf auto_arrange = False Then
    auto_arrange = True
End If
End Sub
Private Sub Command3_Click()
    MsgBox ("This option will allow you to pick your favorite program to start up when Desktop is loaded"), vbInformation
    CommonDialog2.FileName = ""
    CommonDialog2.ShowOpen
    Label5.Caption = CommonDialog2.FileName
    MsgBox ("Data has been saved. Restart program for results"), vbInformation
 
End Sub

Private Sub Command4_Click()
If times = True Then
times = False
ElseIf times = False Then
times = True
End If

End Sub

Private Sub Command5_Click()
Form5.Show
End Sub

Private Sub Command6_Click()
ShellExecute hWnd, "open", "http://www.geocities.com/nit3shift", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Command7_Click()
Dim filenn As String
Dim filen As String
    filen = "C:\Yar Desk\desktopload2.lsd"
    Open filen For Output As #1
    On Error Resume Next
    filenn = Form1.Label12.Caption 'mouse icon
    Print #1, filenn
    filenn = Form1.Label14.Caption 'desktop background
    Print #1, filenn
    Print #1, Form2.Label2.Caption 'startupwav
    Print #1, auto_arrange  'organize
    Print #1, Form2.Label5.Caption  'Favorite startup program
    Print #1, "Timer=" & times 'Time on
    Print #1, "Last edited "; "Time"; Date
Close #1
Command
End Sub
Sub DragForm(Form As Form)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, &O2, &O0
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Command9.Visible = True
    Command8.Visible = False
End Sub
Private Sub Command9_Click()
If Label2.Caption = "" Then
    MsgBox ("You must load a wav file by pressing the sound button" + " This can be accessed by pressing the very first button at the top"), vbExclamation
Else
    e = sndPlaySound(Label2.Caption, 1)
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image7.Visible = True
    Image8.Visible = False
    Command9.Visible = False
    Command8.Visible = True
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image8.Visible = True
    Image7.Visible = False
End Sub

Private Sub Image8_Click()
Form1.Image15.Visible = False
Unload Form2
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image7.Visible = True
    Image8.Visible = False
    DragForm Me
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragForm Me
End Sub
Private Sub save()
'Dim filen As String
    'filen = "C:\Documents and Settings\Joshua Nixon\Desktop\myprogram\desktopload.txt"
    'Open filen For Output As #1
    'Print #1, "Start Up Wav=" & Label2.Caption
    'Print #1, auto_arrange
    'Print #1, "Favorite StartUp Program=" & Label5.Caption
    
'Close #1
End Sub
                                                                                                                            
                                                                                                                                                                            
