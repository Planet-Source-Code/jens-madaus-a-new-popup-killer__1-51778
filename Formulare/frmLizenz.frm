VERSION 5.00
Begin VB.Form frmLizenz 
   BackColor       =   &H00C6C7C6&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtLizenz 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   10
      Top             =   2160
      Width           =   3135
   End
   Begin Popup.MyButton MyButton1 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
      _extentx        =   3413
      _extenty        =   661
      btype           =   3
      tx              =   "&Freischalten"
      hov             =   0   'False
      hovres          =   0   'False
      enab            =   -1  'True
      prelit          =   0   'False
      font            =   "frmLizenz.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   13160660
      fcol            =   0
   End
   Begin VB.TextBox txtLizenz 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox txtLizenz 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtLizenz 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.PictureBox picTray 
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6960
      Width           =   495
   End
   Begin Popup.MyTopButton btnClose 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6960
      Width           =   315
      _extentx        =   556
      _extenty        =   556
      imagedown       =   "frmLizenz.frx":002C
      imagehot        =   "frmLizenz.frx":047E
      imagedisabled   =   "frmLizenz.frx":090A
      imageup         =   "frmLizenz.frx":0D30
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lizenzcode:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vorname:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Real Popup-Killer lizenzieren..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   7560
      Width           =   2490
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   2
      Left            =   3000
      Top             =   6840
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   0
      Left            =   2280
      Top             =   6840
      Width           =   285
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   2760
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   7680
      Width           =   195
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   3600
      Top             =   6840
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   3960
      Top             =   6840
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   1
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   285
   End
End
Attribute VB_Name = "frmLizenz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Unload Me
End Sub

Private Sub Form_Load()
    With frmPics
        imgTitle(0).Picture = .imgTitle(0).Picture
        imgTitle(1).Picture = .imgTitle(1).Picture
        imgTitle(2).Picture = .imgTitle(2).Picture
        imgWindowBottomLeft.Picture = .imgWindowBottomLeft.Picture
        imgWindowBottomRight.Picture = .imgWindowBottomRight.Picture
        imgWindowRight.Picture = .imgWindowRight.Picture
        imgWindowBottom.Picture = .imgWindowBottom.Picture
        imgWindowLeft.Picture = .imgWindowLeft.Picture
    End With
    
    MakeWindow Me, False
    SetForegroundWindow Me.hWnd
End Sub
