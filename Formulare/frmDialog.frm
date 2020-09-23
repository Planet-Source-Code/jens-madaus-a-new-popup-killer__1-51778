VERSION 5.00
Begin VB.Form frmDialog 
   BackColor       =   &H00C6C7C6&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   8265
   ClientTop       =   6210
   ClientWidth     =   1425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   1425
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstUrls 
      Height          =   1425
      ItemData        =   "frmDialog.frx":0000
      Left            =   480
      List            =   "frmDialog.frx":0002
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtDialogInput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin Popup.MyTopButton btnClose 
      Height          =   315
      Left            =   4320
      TabIndex        =   8
      ToolTipText     =   "Beenden"
      Top             =   5400
      Width           =   315
      _extentx        =   556
      _extenty        =   556
      imagedown       =   "frmDialog.frx":0004
      imagehot        =   "frmDialog.frx":0456
      imagedisabled   =   "frmDialog.frx":08E2
      imageup         =   "frmDialog.frx":0D08
   End
   Begin Popup.MyButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "&Ok"
      hov             =   -1
      hovres          =   0
      enab            =   -1
      prelit          =   0
      font            =   "frmDialog.frx":11A2
      coltype         =   1
      focusr          =   -1
      bcol            =   13160660
      fcol            =   0
   End
   Begin Popup.MyButton cmdNein 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "Nein"
      hov             =   -1
      hovres          =   0
      enab            =   -1
      prelit          =   0
      font            =   "frmDialog.frx":11CE
      coltype         =   1
      focusr          =   -1
      bcol            =   13160660
      fcol            =   0
   End
   Begin Popup.MyButton cmdJa 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "&Ja"
      hov             =   -1
      hovres          =   0
      enab            =   -1
      prelit          =   0
      font            =   "frmDialog.frx":11FA
      coltype         =   1
      focusr          =   -1
      bcol            =   13160660
      fcol            =   0
   End
   Begin Popup.MyButton cmdAbbruch 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "&Abbrechen"
      hov             =   -1
      hovres          =   0
      enab            =   -1
      prelit          =   0
      font            =   "frmDialog.frx":1226
      coltype         =   1
      focusr          =   -1
      bcol            =   13160660
      fcol            =   0
   End
   Begin VB.PictureBox picTray 
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   11235
      Width           =   135
   End
   Begin VB.Image imginformation 
      Height          =   480
      Left            =   5040
      Picture         =   "frmDialog.frx":1252
      Top             =   5400
      Width           =   480
   End
   Begin VB.Image imgquestion 
      Height          =   480
      Left            =   5880
      Picture         =   "frmDialog.frx":1F1C
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgcritical 
      Height          =   480
      Left            =   5880
      Picture         =   "frmDialog.frx":2BE6
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgexclamation 
      Height          =   480
      Left            =   5040
      Picture         =   "frmDialog.frx":38B0
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   60
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Achtung!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   480
      TabIndex        =   2
      Top             =   5880
      Width           =   750
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   1800
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   6075
      Width           =   195
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   0
      Left            =   360
      Top             =   8040
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   1
      Left            =   840
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   2
      Left            =   1320
      Top             =   8040
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   2280
      Top             =   8040
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   1800
      Top             =   8040
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   285
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbbruch_Click()
    cmdAnswer = 1
    Ende
End Sub

Private Sub cmdJa_Click()
    cmdAnswer = 2
    Ende
End Sub

Private Sub cmdNein_Click()
    cmdAnswer = 3
    Ende
End Sub

Private Sub cmdOk_Click()
    Select Case True
        Case (txtDialogInput.Visible)
            cmdAnswer = 0 & ";" & Me.txtDialogInput.Text
        Case (lstUrls.Visible)
            If lstUrls.ListCount > 0 Then
                Dim i As Long
                Dim tmpS As String
                For i = 0 To lstUrls.ListCount - 1
                    If lstUrls.Selected(i) Then
                        tmpS = lstUrls.List(i)
                        Exit For
                    End If
                Next
                cmdAnswer = 0 & ";" & tmpS
            End If
        Case Else
            cmdAnswer = 0
    End Select
    Ende
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
    
    btnClose.Enabled = False
    
    With Me
        Select Case lngLangIndex
            Case 0
                .cmdAbbruch.Caption = "&Abbrechen"
                .cmdJa.Caption = "&Ja"
                .cmdNein.Caption = "&Nein"
            Case 1
                .cmdAbbruch.Caption = "&Cancel"
                .cmdJa.Caption = "&Yes"
                .cmdNein.Caption = "&No"
        End Select
    End With
    
    lblMessage.Left = 200
    lblMessage.Top = 700
End Sub

Private Sub imgTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    DoDragMe
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoDragMe
End Sub

Private Sub DoDragMe()
    On Error GoTo PROC_ERR
    
    DoDrag Me
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    ShowDialog frmMainCode, Err.Description, IIf(lngLangIndex = 0, "Fehler ...!" & Err.Number, "Error...!" & Err.Number), OK, CRITICAL
    Resume PROC_EXIT
End Sub

Private Sub Ende()
    Unload Me
End Sub

Private Sub txtDialogInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOk_Click
End Sub
