VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Desktop"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   Icon            =   "desktopprog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      Height          =   3615
      Left            =   0
      MouseIcon       =   "desktopprog.frx":1042
      ScaleHeight     =   3615
      ScaleWidth      =   5415
      TabIndex        =   4
      Top             =   240
      Width           =   5415
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tile Deskop"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1800
         Picture         =   "desktopprog.frx":1D0C
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Timer Timer5 
         Interval        =   1000
         Left            =   2640
         Top             =   840
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2280
         Top             =   840
      End
      Begin VB.PictureBox Picture8 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   4560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   1920
         Top             =   840
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         ScaleHeight     =   585
         ScaleWidth      =   2145
         TabIndex        =   18
         Top             =   2880
         Width           =   2175
         Begin VB.Label Label9 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Loading"
            BeginProperty Font 
               Name            =   "911 Porscha"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   19
            Top             =   240
            Width           =   1410
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   1560
         Top             =   840
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Load Picture"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2720
         Picture         =   "desktopprog.frx":2096
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1800
         Picture         =   "desktopprog.frx":2420
         ScaleHeight     =   225
         ScaleWidth      =   1800
         TabIndex        =   5
         Top             =   580
         Visible         =   0   'False
         Width           =   1800
         Begin VB.Image Image7 
            Height          =   120
            Left            =   1560
            Picture         =   "desktopprog.frx":397A
            Top             =   45
            Width           =   150
         End
         Begin VB.Image Image8 
            Height          =   120
            Left            =   1580
            Picture         =   "desktopprog.frx":3ABC
            Top             =   50
            Visible         =   0   'False
            Width           =   150
         End
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   1320
         Picture         =   "desktopprog.frx":3BFE
         ScaleHeight     =   1725
         ScaleWidth      =   2580
         TabIndex        =   43
         Top             =   960
         Width           =   2580
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   5325
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0100110000101110010101110110100001101001011101000110010100111101010010000110111101110100"
         Height          =   195
         Left            =   -960
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   6840
      End
      Begin VB.Image Image17 
         Height          =   240
         Left            =   3960
         Picture         =   "desktopprog.frx":920E
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4800
         TabIndex        =   27
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Here for mini brwoser"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   3120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Image Image16 
         Height          =   480
         Left            =   4320
         Picture         =   "desktopprog.frx":9598
         Top             =   2640
         Width           =   480
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   4080
         Picture         =   "desktopprog.frx":A5DA
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "All Folder"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Text Folder"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   4080
         Picture         =   "desktopprog.frx":AEA4
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   480
         Picture         =   "desktopprog.frx":B76E
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apps Folder"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Media Folder"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   480
         Picture         =   "desktopprog.frx":C038
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   480
         Picture         =   "desktopprog.frx":C902
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Image Folder"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Image Image6 
         Height          =   1215
         Left            =   1850
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   2175
         Left            =   1800
         Top             =   600
         Visible         =   0   'False
         Width           =   1790
      End
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
      Picture         =   "desktopprog.frx":D1CC
      ScaleHeight     =   270
      ScaleWidth      =   5295
      TabIndex        =   44
      Top             =   0
      Width           =   5295
      Begin VB.Image Image1 
         Height          =   120
         Left            =   4800
         Picture         =   "desktopprog.frx":11C96
         Top             =   75
         Width           =   150
      End
      Begin VB.Image Image2 
         Height          =   120
         Left            =   5020
         Picture         =   "desktopprog.frx":11DD8
         Top             =   75
         Width           =   150
      End
      Begin VB.Image Image3 
         Height          =   120
         Left            =   4800
         Picture         =   "desktopprog.frx":11F1A
         Top             =   75
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image Image4 
         Height          =   120
         Left            =   5040
         Picture         =   "desktopprog.frx":1205C
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
         TabIndex        =   46
         Top             =   15
         Width           =   1335
      End
      Begin VB.Image Image5 
         Height          =   225
         Left            =   20
         Picture         =   "desktopprog.frx":1219E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Left            =   2740
         TabIndex        =   45
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      Picture         =   "desktopprog.frx":131E0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   31
      Top             =   5640
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   30
      Top             =   4680
      Width           =   480
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   840
      Picture         =   "desktopprog.frx":1384A
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   29
      Top             =   5640
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Height          =   135
      Left            =   120
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Image Files|*.jpg;*.bmp;*.gif;*.ico;*.cur"
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3000
      Top             =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      Picture         =   "desktopprog.frx":1396C
      ScaleHeight     =   825
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   3840
      Width           =   5295
      Begin VB.CommandButton Command11 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   3000
         Picture         =   "desktopprog.frx":21D6A
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   4560
         Picture         =   "desktopprog.frx":22A34
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Quit"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   2520
         Picture         =   "desktopprog.frx":236FE
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Quick Launch"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   3840
         Picture         =   "desktopprog.frx":24740
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   2040
         Picture         =   "desktopprog.frx":2500A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00000040&
         Height          =   495
         Left            =   1080
         Picture         =   "desktopprog.frx":2604C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   120
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00000000&
         Caption         =   "Command5"
         Height          =   495
         Left            =   1560
         Picture         =   "desktopprog.frx":26916
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Clear Desktop"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   120
         Picture         =   "desktopprog.frx":271E0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   600
         Picture         =   "desktopprog.frx":27AAA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture9 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3510
      Left            =   0
      Picture         =   "desktopprog.frx":28374
      ScaleHeight     =   3510
      ScaleWidth      =   1545
      TabIndex        =   36
      Top             =   240
      Width           =   1545
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Image Image28 
         Height          =   855
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "double click to open."
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   42
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Image Image26 
         Height          =   480
         Left            =   0
         Picture         =   "desktopprog.frx":2A963
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image25 
         Height          =   480
         Left            =   0
         Picture         =   "desktopprog.frx":2B7A5
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image24 
         Height          =   480
         Left            =   0
         Picture         =   "desktopprog.frx":2C06F
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Left            =   0
         Picture         =   "desktopprog.frx":2C939
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Back"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   40
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Yar Interactive Desktop Version 1.0  "
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   38
         Top             =   480
         Width           =   1575
      End
      Begin VB.Image Image18 
         Height          =   480
         Left            =   0
         Picture         =   "desktopprog.frx":2D203
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Files"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   37
         Top             =   2160
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   1560
      TabIndex        =   39
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   6615
      EndProperty
      Picture         =   "desktopprog.frx":2DACD
   End
   Begin VB.Image Image20 
      Height          =   480
      Left            =   4560
      Picture         =   "desktopprog.frx":2F103
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image19 
      Height          =   480
      Left            =   4560
      Picture         =   "desktopprog.frx":2FD47
      Top             =   4680
      Width           =   480
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label19"
      Height          =   255
      Left            =   3600
      TabIndex        =   35
      Top             =   4920
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label18"
      Height          =   255
      Left            =   3360
      TabIndex        =   34
      Top             =   4920
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "911 Porscha"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   150
      Left            =   480
      TabIndex        =   28
      Top             =   4680
      Width           =   90
   End
   Begin VB.Image Image15 
      Height          =   240
      Left            =   4680
      Picture         =   "desktopprog.frx":30611
      Top             =   5250
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image14 
      Height          =   210
      Left            =   4995
      Picture         =   "desktopprog.frx":3078E
      Top             =   5250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   5205
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'|Joshua Nixon
'|16 Years Old in 10th
'|Yar Interactive
'|Desktop Version 1.0
'```````````````````'
'   Any Questions   '
'    Email me at    '
'JNixon21@excite.com'
'        Or         '
'Aim/Yahoo:Nit3shift'
'````````````````````
'This source is free to use by anyone thanks and vote please
'90% percent done application. Full should be done shortly
DefInt A-R
Dim j
Dim i
Dim auto_arrange As Boolean
Private Sub Command1_Click()
    Dim openfilename As String
    MsgBox ("Large pictures will take more time to process. After process is done folders will appear again."), vbInformation
    Command4.Visible = True
    Picture4.Visible = True
    Image6.Visible = True
    Command3.Visible = True
    Command4.Visible = True
    Shape1.Visible = True
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    openfilename = CommonDialog1.FileName
    Image6.Picture = LoadPicture(openfilename)
    Timer1.Enabled = True
    Picture5.Visible = True
    Label14.Caption = CommonDialog1.FileName
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label15.Caption = "Desktop Background"
End Sub

Private Sub Command10_Click()
    Unload Form1
    Unload Form2
    Unload Form3
    Unload Form4
    Unload Form5
    End
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label15.Caption = "Exit Program"
End Sub

Private Sub Command11_Click()
Call TileMe
End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.Caption = "Refresh Desktop"
End Sub

Private Sub Command2_Click()
    Image15.Visible = True
    Form2.Show
    Timer1.Enabled = True
    If Image14.Visible = False Then
    Image15.Left = 333
    Else
    Image15.Left = 312
    End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.Caption = "Properties"
End Sub

Private Sub Command3_Click()
Call TileMe2
End Sub
Sub TileMe2()
Call HideMe
    Call cleardesk
    Picture4.ScaleMode = 3
    Call hideme2
            Let Picture3.Picture = Image6.Picture
    Dim j
    Dim i
        For i = 0 To 35
        For j = 0 To 35
            DoEvents
                 Call HideMe
                 Call hideme2
                 
                 
                 Picture4.PaintPicture Picture3, j * Form1.Picture3.Width, i * Form1.Picture3.Height, Form1.Picture3.Width, Form1.Picture3.Height
                  
        Next j, i

Timer1.Enabled = True
Call ShowMe
Picture4.ScaleMode = 1
8:
End Sub



Sub TileMe()
Label9.Caption = "Loading Pictures"
Call cleardesk
    Call HideMe
            Picture4.ScaleMode = 3
   Call hideme2
        
        For i = 0 To 35
        For j = 0 To 35
            
            DoEvents
                 
                 Call HideMe
                 Call hideme2
                Picture4.PaintPicture Picture3, j * Form1.Picture3.Width, i * Form1.Picture3.Height, Form1.Picture3.Width, Form1.Picture3.Height
        
        Next j, i

Timer1.Enabled = True
Call ShowMe
Picture4.ScaleMode = 1
End Sub
Sub ShowMe()
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Image9.Visible = True
    Image10.Visible = True
    Image11.Visible = True
    Image12.Visible = True
    Image13.Visible = True
    Image16.Visible = True
    Image17.Visible = True
End Sub
Sub cleardesk()
    Form1.Picture4.PaintPicture Form1.Picture6.Image, 0, 0
    Picture4.BackColor = vbWhite
    Picture4.Cls
    Picture4.Refresh
End Sub
Sub hideme2()
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Image9.Visible = False
    Image10.Visible = False
    Image11.Visible = False
    Image12.Visible = False
    Image13.Visible = False
    Image16.Visible = False
    Image17.Visible = False
Picture4.ScaleMode = 3
Exit Sub
End Sub
Private Sub Command4_Click()
    Dim openfilename  As String
    Picture4.Visible = True
    Image6.Visible = True
    Command3.Visible = True
    Shape1.Visible = True
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    openfilename = CommonDialog1.FileName
    Image6.Picture = LoadPicture(openfilename)
End Sub

Private Sub Command5_Click()
 
    Call cleardesk
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.Caption = "Clear Background"
End Sub

Private Sub Command6_Click()
    Image14.Visible = True
    Form3.Show
If Image15.Visible = True + Image15.Left = 333 Then
    Image14.Left = 333 & Image15.Left = 312
Else
    Image15.Left = 312
End If
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.Caption = "Change Cursor"
End Sub

Private Sub Command7_Click()
Call autoarrange
End Sub
Sub autoarrange()
Call ShowMe
    Image9.Top = 2280
    Image9.Left = 480
    Image11.Top = 360
    Image11.Left = 480
    Image12.Top = 360
    Image12.Left = 4080
    Image10.Top = 1320
    Image10.Left = 480
    Image13.Top = 1320
    Label5.Top = 1800
    Label5.Left = 240
    Label8.Top = 1800
    Label7.Top = 840
    Label7.Left = 3840
    Label6.Top = 840
    Label6.Left = 240
    Label4.Top = 2760
    Label4.Left = 240
    Image13.Left = 4080
    Label8.Left = 3840
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.Caption = "Auto Arrange"
End Sub

Private Sub Command8_Click()
Call saving
End Sub
Sub saving()
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
Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.Caption = "Save All Data"
End Sub

Private Sub Command9_Click()
Let Label11.Caption = Form2.Label5.Caption
If Form2.Label5.Caption = "" Then
    MsgBox ("You must select a favorite program from the properties menu"), vbInformation
    Form2.Show
Else
    ShellExecute hWnd, "open", Form2.Label5.Caption, "", "", vbNormalFocus
End If
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.Caption = "Favorite Program"
End Sub

Private Sub Form_Load()
    times = True
    binmar1 = True
    counter = 0
    auto_arrange = False
    Swapmouse = True
    Form1.Show
    Image1.Visible = True
    Image3.Visible = False
    Image4.Visible = False
    Picture2.Visible = True
    Timer1.Enabled = True
    Image7.Visible = True
    Form1.MousePointer = 99
Call LoadPersonal
End Sub
Sub LoadPersonal()
Dim filen As String
If Dir("C:\Yar Desk\donttake.lsd") = "" Then
    MsgBox "This is your first time to open the program." + "You will need you go through a wizard, which will pop up right after you click okay.", vbInformation
    Form1.Hide
    Unload Form1
    Form5.Show
    Open "C:\Yar Desk\donttake.lsd" For Output As #1
                                                                                                                                 Print #1, "Dont Take Drugs. JUST SAY NO!!!!! :)!"
Close #1
End If
If Dir("C:\Yar Desk\desktopload2.lsd") = "" Then
    MsgBox "It appears the loading file does not exist. Go through the wizard."
    On Error Resume Next
    Open "C:\Yar Desk\desktopload2.lsd" For Output As #1
    Close #1
    GoTo 7
Else
    filen = "C:\Yar Desk\desktopload2.lsd"
    Open filen For Input As #1
    On Error Resume Next
    Line Input #1, filen
    Label12.Caption = filen
    Picture8.Picture = LoadPicture(Label12.Caption)
    Form1.MousePointer = Default
    Let Form1.MouseIcon = Picture8.Picture
    Line Input #1, filen
    Picture3.Picture = LoadPicture(filen)
    Call TileMe
    Line Input #1, filen
    retval = sndPlaySound(filen, 1)
    Line Input #1, filen
    'nothing yet
    Line Input #1, filen
    Form2.Label5.Caption = filen
    Line Input #1, filen
    Timer1.Enabled = filen
Form1.MousePointer = 11
    Load Form1
    Label9.Caption = "Loading."
    Label9.Refresh
    Load Form2
    Label9.Caption = "Loading.."
    Label9.Refresh
    Load Form3
    Label9.Caption = "Loading..."
    Label9.Refresh
    Load Form4
    Label9.Caption = "Loading...."
    Load Form6
    Label9.Caption = "Loading....."
    Label9.Refresh
    Label9.Caption = "Loading Files."
    Label9.Refresh
Close #1
    Picture13.Visible = False
    Picture7.Visible = False
    Form1.MousePointer = 99
End If
7:
On Error Resume Next

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image19.Visible = True
Image20.Visible = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Visible = False
    Image3.Visible = True
End Sub


Private Sub Image10_DblClick()
Call loadmediafile
End Sub
Private Sub Image11_DblClick()
Call loadappzfile
End Sub
Private Sub Image12_DblClick()
Call loadtextfile
End Sub
Private Sub Image13_DblClick()
Call loadotherfile
End Sub
Private Sub Image14_DblClick()
    Image14.Left = 4995
    Image14.Visible = False
    Unload Form3
End Sub
Private Sub Image15_DblClick()
    Image15.Visible = False
    Unload Form2
End Sub
Private Sub Image16_Click()
    Form4.Show
End Sub
Private Sub Image16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Visible = True
Label10.Caption = "Click Here for mini brwoser"
Timer4.Enabled = True
End Sub
Private Sub Image17_Click()
MsgBox ("Note: Colors are for temporary use, not for saving."), vbInformation
Load Form6
Form6.Show
End Sub

Private Sub Image17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label10.Visible = True
    Label10.Caption = "Click Here to change desktop backgorund"
    Timer4.Enabled = True

End Sub

Private Sub Image19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image19.Visible = False
Image20.Visible = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Visible = True
    Image2.Visible = False
End Sub



Private Sub Image27_Click()
If Dir("C:\Yar Desk\donttake.lsd") = "" Then
MsgBox ("There is nothing to delete"), vbInformation
Else
MsgBox ("Are you sure you want to delete all settings."), vbYesNo + vbCritical
Kill "C:\Yar Desk\donttake.lsd"
If Dir("C:\Yar Desk\wizard.lsd") = "" Then
MsgBox ("Deleted"), vbInformation
Else
Kill "C:\Yar Desk\wizard.lsd"
End If
End If
End Sub

Private Sub Image27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.Caption = "Delete all settings."
End Sub

Private Sub Image20_Click()
Dim res
res = MsgBox("Are you sure you want to delete all settings.", vbYesNo)
If res = 6 Then
KillFolderTree ("C:\Yar Desk\")
MsgBox ("Program must be restarted in order to function properlly."), vbInformation
UnloadAll Form1, Form2, Form3, Form4, Form5
Else
End If
End Sub


Private Sub Image3_Click()
Form1.WindowState = 1
End Sub

Private Sub Image4_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4

End
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub Image7_Click()
Unload Form2
End
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
Image8.Visible = True
End Sub

Private Sub Image8_Click()
Image6.Visible = False
    Image7.Visible = False
    Image8.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Shape1.Visible = False
    Picture5.Visible = False
    
End Sub

Private Sub Image9_DblClick()
Call loadImageFile
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

DragForm Me
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image2.Visible = True
End Sub

Private Sub Label22_Click()
Picture4.Visible = True
End Sub
Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label22.ForeColor = vbBlue
End Sub

Private Sub Label4_DblClick()
Call loadImageFile
End Sub
Sub loadImageFile()
Picture4.Visible = False
Label20.Caption = "Image Files"
Image28.Visible = False
Image23.Visible = False
Image28.Visible = True
Image18.Visible = True
Image24.Visible = False
Image25.Visible = False
Image26.Visible = False
Call GetImageFiles
End Sub
Private Sub Label5_DblClick()
Call loadmediafile
End Sub
Sub loadmediafile()
ListView1.ListItems.Clear
Dim imtx2
Picture4.Visible = False
Label20.Caption = "Media Files"
Image28.Visible = False
Image23.Visible = False
Image18.Visible = False
Image24.Visible = True
Image25.Visible = False
Image26.Visible = False
Call GetMediaFiles
End Sub
Private Sub Label6_DblClick()
Call loadappzfile
End Sub
Sub loadappzfile()
Picture4.Visible = False
Label20.Caption = "Apps Files"
Image28.Visible = False
Image23.Visible = True
Image18.Visible = False
Image24.Visible = False
Image25.Visible = False
Image26.Visible = False
Call GetAppsFiles
End Sub
Private Sub Label7_DblClick()
Call loadtextfile
End Sub
Sub loadtextfile()


Picture4.Visible = False
Label20.Caption = "Text Files"
Image28.Visible = False
Image23.Visible = False
Image18.Visible = False
Image24.Visible = False
Image25.Visible = False
Image26.Visible = True
Call GetTextFiles
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto_arrange = False Then
    Label7.Drag vbBeginDrag
    Let Image12.Top = Label7.Top - 500
    Let Image12.Left = Label7.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Else
End If
End Sub
Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto_arrange = False Then
    Label6.Drag vbEndDrag
    Let Image12.Top = Label7.Top - 500
    Let Image12.Left = Label7.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Else
End If
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto_arrange = False Then
    Label6.Drag vbBeginDrag
    Let Image11.Top = Label6.Top - 500
    Let Image11.Left = Label6.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Else
End If
End Sub
Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto_arrange = False Then
    Label6.Drag vbEndDrag
    Let Image11.Top = Label6.Top - 500
    Let Image11.Left = Label6.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Else
End If
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto_arrange = False Then
    Label5.Drag vbBeginDrag
    Let Image11.Top = Label6.Top - 500
    Let Image11.Left = Label6.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Else
End If
End Sub
Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto_arrange = False Then
    Label5.Drag vbEndDrag
    Let Image10.Top = Label5.Top - 500
    Let Image10.Left = Label5.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Else
End If
End Sub

Sub loadotherfile()
images = False
other = True
app = False
Text = False
media = False
Image28.Visible = False
Picture4.Visible = False
    Label20.Caption = "Other Files"
    Image28.Visible = True
    Image23.Visible = False
    Image18.Visible = False
    Image24.Visible = False
    Image25.Visible = True
    Image26.Visible = False
Call GetOtherFiles
End Sub

Private Sub Label8_DblClick()
Call loadotherfile
End Sub

Private Sub ListView1_DblClick()
ShellExecute hWnd, "open", ListView1.SelectedItem, "", "", vbNormalFocus
If Image28.Visible = True Then
Image28.Picture = LoadPicture(ListView1.SelectedItem)
Else
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Label24.Caption = GetFileSize(ListView1.SelectedItem)
If Image28.Visible = True Then
On Error GoTo 1
Image28.Picture = LoadPicture(ListView1.SelectedItem)
On Error GoTo 1
Else
End If
1:
End Sub

Private Sub picture4_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move (X - Source.Width / 2), (Y - Source.Height / 2)
    Let Image9.Top = Label4.Top - 500
    Let Image9.Left = Label4.Left + 265
    Let Image10.Top = Label5.Top - 500
    Let Image10.Left = Label5.Left + 265
    Let Image11.Top = Label6.Top - 500
    Let Image11.Left = Label6.Left + 265
    Let Image12.Top = Label7.Top - 500
    Let Image12.Left = Label7.Left + 265
    Let Image13.Top = Label8.Top - 500
    Let Image13.Left = Label8.Left + 265
Call ShowMe
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto_arrange = False Then
    Label8.Drag vbBeginDrag
    Let Image13.Top = Label8.Top - 500
    Let Image13.Left = Label8.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Call ShowMe
Else
End If
End Sub
Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto_arrange = False Then
    Label4.Drag vbEndDrag
    Let Image13.Top = Label8.Top - 500
    Let Image13.Left = Label8.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Call ShowMe
Else
End If
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto_arrange = False Then
    Label4.Drag vbBeginDrag
    Let Image9.Top = Label4.Top - 500
    Let Image9.Left = Label4.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Image9.Visible = True
Call ShowMe
Else
End If

End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.Drag vbEndDrag
    Let Image9.Top = Label4.Top - 500
    Let Image9.Left = Label4.Left + 265
    Form1.MouseIcon = Form3.Picture3.Picture
Call ShowMe
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
Timer1.Enabled = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image1.Visible = True
 Image2.Visible = True
 Image4.Visible = False
 Timer1.Enabled = True
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Timer1.Enabled = True
        Label15.Caption = ""
    End Sub
Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pt As POINTAPI
    Call GetCursorPos(pt)
    Label18.Caption = pt.X
    Label19.Caption = pt.Y
    Timer4.Enabled = False
    Label10.Visible = False
    Image1.Visible = True
    Image3.Visible = False
    Image4.Visible = False
    Picture2.Visible = True
    Timer1.Enabled = True
    Image7.Visible = True
    Image8.Visible = False
    binmar1 = False
    Timer5.Enabled = False
    Call ShowMe
End Sub

Sub HideMe()
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Command4.Visible = False
    Image6.Visible = False
    Command3.Visible = False
    Shape1.Visible = False
    Command4.Visible = False
    Image8.Visible = False
    Image7.Visible = False
    Picture5.Visible = False
    Image16.Visible = False
    Image17.Visible = False
End Sub
Public Sub DragForm(Form As Form)
     ReleaseCapture
     SendMessage Me.hWnd, &HA1, 2&, &O0
 Let Form2.Top = Form1.Top
 Timer1.Enabled = True
 End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label22.ForeColor = vbBlack
End Sub
Private Sub Timer1_Timer()
Label2.Caption = Time
End Sub
Private Sub Timer2_Timer()
X = GetSystemMetrics(0)
Y = GetSystemMetrics(1)
Label3.Caption = "Resoultion " & CStr(X) & " X " & CStr(Y)
    Picture10.Cls
  
    BitBlt Picture10.hdc, 0, 0, 27, 55, Picture12.hdc, counter, 0, SRCPAINT
    BitBlt Picture10.hdc, 0, 0, 27, 55, Picture11.hdc, counter, 0, SRCAND
    Picture10.Refresh 'refresh picture1
    counter = counter + 27 'add 64 to x
    Picture10.Refresh
    Let i = 54
    If counter >= i Then counter = 0 'checks to see
Picture10.Refresh
End Sub

Private Sub Timer3_Timer()
Call ShowMe
End Sub

Private Sub Timer4_Timer()
Call ChangeColor(Label10)
End Sub
Sub wait()
Dim hol As Integer
For hol = 1 To 2243
DoEvents
Next hol
Exit Sub
End Sub

