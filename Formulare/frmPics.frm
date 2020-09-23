VERSION 5.00
Begin VB.Form frmPics 
   ClientHeight    =   0
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   1320
      Picture         =   "frmPics.frx":0000
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   840
      Picture         =   "frmPics.frx":03D7
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   2
      Left            =   480
      Picture         =   "frmPics.frx":0779
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   1
      Left            =   360
      Picture         =   "frmPics.frx":0FB3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   0
      Left            =   0
      Picture         =   "frmPics.frx":12F7
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   2400
      Picture         =   "frmPics.frx":1793
      Stretch         =   -1  'True
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   3360
      Picture         =   "frmPics.frx":1EDD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   2880
      Picture         =   "frmPics.frx":2627
      Stretch         =   -1  'True
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   0
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   720
      Width           =   195
   End
End
Attribute VB_Name = "frmPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
