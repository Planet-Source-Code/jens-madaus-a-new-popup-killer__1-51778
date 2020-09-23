VERSION 5.00
Begin VB.Form frmMainCode 
   Appearance      =   0  '2D
   BackColor       =   &H00C6C7C6&
   BorderStyle     =   0  'Kein
   Caption         =   "Real Popup-Killer"
   ClientHeight    =   5460
   ClientLeft      =   4185
   ClientTop       =   4800
   ClientWidth     =   2760
   ControlBox      =   0   'False
   Icon            =   "frmMainCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   2760
   StartUpPosition =   2  'Bildschirmmitte
   Begin Popup.MyButton btnReg 
      Height          =   285
      Left            =   140
      TabIndex        =   24
      Top             =   1350
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "&Registrieren"
      HOV             =   -1  'True
      HOVRES          =   0   'False
      PRELIT          =   0   'False
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin Popup.MyTopButton btnMinimizeTray 
      Height          =   315
      Left            =   1320
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Minimieren zum SysTray"
      Top             =   7800
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      ImageDown       =   "frmMainCode.frx":030A
      ImageHot        =   "frmMainCode.frx":073D
      ImageUp         =   "frmMainCode.frx":0BDC
   End
   Begin Popup.MyTopButton btnMinimize 
      Height          =   315
      Left            =   1680
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Minimieren"
      Top             =   7800
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      ImageDown       =   "frmMainCode.frx":1066
      ImageHot        =   "frmMainCode.frx":14A6
      ImageUp         =   "frmMainCode.frx":1942
   End
   Begin Popup.MyTopButton btnClose 
      Height          =   315
      Left            =   2040
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7800
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      ImageDown       =   "frmMainCode.frx":1DCC
      ImageHot        =   "frmMainCode.frx":221E
      ImageDisabled   =   "frmMainCode.frx":26A8
      ImageUp         =   "frmMainCode.frx":2ACD
   End
   Begin VB.PictureBox picTray 
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   600
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7800
      Width           =   495
   End
   Begin Popup.MyButton Command2 
      Height          =   285
      Left            =   1410
      TabIndex        =   2
      Top             =   1020
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "&Info"
      HOV             =   -1  'True
      HOVRES          =   0   'False
      PRELIT          =   0   'False
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin Popup.MyButton Command1 
      Height          =   285
      Left            =   1410
      TabIndex        =   3
      Top             =   1350
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "&>>"
      HOV             =   -1  'True
      HOVRES          =   0   'False
      PRELIT          =   0   'False
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin VB.TextBox txtTmplbl 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   165
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C6C7C6&
      Height          =   3450
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Frame Frame2 
         Height          =   25
         Index           =   3
         Left            =   0
         TabIndex        =   29
         Top             =   690
         Width           =   2520
      End
      Begin VB.CheckBox chkDial 
         BackColor       =   &H00C6C7C6&
         Caption         =   "Dialerdownload blocken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   100
         TabIndex        =   28
         Top             =   420
         Value           =   1  'Aktiviert
         Width           =   2355
      End
      Begin VB.Frame Frame2 
         Height          =   25
         Index           =   2
         Left            =   0
         TabIndex        =   23
         Top             =   1430
         Width           =   2520
      End
      Begin VB.ComboBox cmbLanguage 
         Height          =   315
         ItemData        =   "frmMainCode.frx":2F67
         Left            =   1320
         List            =   "frmMainCode.frx":2F71
         TabIndex        =   6
         Text            =   "cmbLanguage"
         Top             =   1070
         Width           =   1020
      End
      Begin Popup.MyButton Command5 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   3060
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "aktuelle &URL hinzufügen"
         HOV             =   -1  'True
         HOVRES          =   0   'False
         PRELIT          =   0   'False
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         FCOL            =   0
      End
      Begin VB.CheckBox chkStartOnWin 
         BackColor       =   &H00C6C7C6&
         Caption         =   "Starten mit &Windows"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   100
         TabIndex        =   5
         Top             =   720
         Width           =   2235
      End
      Begin VB.Frame Frame2 
         Height          =   25
         Index           =   1
         Left            =   0
         TabIndex        =   21
         Top             =   990
         Width           =   2520
      End
      Begin Popup.MyButton Command3 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2730
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "&Hinzufügen"
         HOV             =   -1  'True
         HOVRES          =   0   'False
         PRELIT          =   0   'False
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         FCOL            =   0
      End
      Begin Popup.MyButton Command4 
         Height          =   285
         Left            =   1305
         TabIndex        =   9
         Top             =   2730
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "&Löschen"
         HOV             =   -1  'True
         HOVRES          =   0   'False
         PRELIT          =   0   'False
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         FCOL            =   0
      End
      Begin VB.CheckBox chkSound 
         BackColor       =   &H00C6C7C6&
         Caption         =   "&Sound bei Popupkill"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   100
         TabIndex        =   4
         Top             =   120
         Value           =   1  'Aktiviert
         Width           =   2055
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E4D3BC&
         Height          =   840
         ItemData        =   "frmMainCode.frx":2F86
         Left            =   120
         List            =   "frmMainCode.frx":2F88
         TabIndex        =   7
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Height          =   25
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   390
         Width           =   2520
      End
      Begin VB.Label lblLanguage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sprache:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   255
         TabIndex        =   22
         Top             =   1100
         Width           =   780
      End
      Begin VB.Label lblSidePopup 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C6C7C6&
         BackStyle       =   0  'Transparent
         Caption         =   "Seiten mit Popupstart:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   12
         Top             =   1520
         Width           =   1935
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1100
      Left            =   0
      Top             =   5880
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   0
      Top             =   6360
   End
   Begin VB.Label txtHyperlink 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1755
      TabIndex        =   27
      Top             =   5160
      Width           =   540
   End
   Begin VB.Label txtEMail 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "email"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1200
      TabIndex        =   26
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label lblKontakt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kontakt:"
      Height          =   195
      Left            =   480
      TabIndex        =   25
      Top             =   5160
      Width           =   600
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E4D3BC&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Klicken = Aktiv/ inaktiv"
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Popups:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   15
      Top             =   1020
      Width           =   690
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1230
      TabIndex        =   14
      Top             =   1030
      Width           =   105
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Real Popup-Killer"
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
      Left            =   840
      TabIndex        =   16
      Top             =   8400
      Width           =   1395
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   0
      Left            =   2880
      Top             =   7680
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   1
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   285
   End
   Begin VB.Image imgTitle 
      Height          =   510
      Index           =   2
      Left            =   3600
      Top             =   7680
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   4560
      Top             =   7680
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   4200
      Top             =   7680
      Width           =   285
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   3360
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   8520
      Width           =   195
   End
End
Attribute VB_Name = "frmMainCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mo As Byte
Dim MA As Byte
Dim MS As Byte
Dim ListString As String

Dim WithEvents IEWin As cIEWindows
Attribute IEWin.VB_VarHelpID = -1
Public TIcon As New clsTray

Private Sub Check1_Click()
    txtTmplbl.SetFocus
End Sub

Private Sub btnClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Unload Me
End Sub

Private Sub btnMinimize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.WindowState = vbMinimized
End Sub

Private Sub btnMinimizeTray_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If Command1.Caption = "&<<" Then Command1_Click
        Me.WindowState = vbMinimized
        Me.Hide
    End If
End Sub

Function SoundOn(p)
    chkSound.Value = Val(p)
    bSound = (chkSound.Value)
End Function

Private Sub btnReg_Click()
    MakeNormal (Me.hwnd)
    frmLizenz.Show vbModal, Me
    MakeTopMost (Me.hwnd)
    txtTmplbl.SetFocus
End Sub

Private Sub chkDial_Click()
    On Error Resume Next
    
    Call ActivateDLKiller(chkDial.Value)
    txtTmplbl.SetFocus
End Sub

Private Sub chkSound_Click()
    On Error Resume Next
    
    bSound = (chkSound.Value)
    frmMenuForm.mSound.Checked = (chkSound.Value)
    txtTmplbl.SetFocus
End Sub

Private Sub chkStartOnWin_Click()
    On Error Resume Next
    
    Call SetAppVar
    
    If IsWinNT Then
        If chkStartOnWin.Value = 1 Then
            Call SetStringWert(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Real Popup-Killer", ap)
        Else
            Call RegFieldDelete(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Real Popup-Killer")
        End If
    Else
        If chkStartOnWin.Value = 1 Then
            Call SetStringWert(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", "Real Popup-Killer", ap)
        Else
            Call RegFieldDelete(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", "Real Popup-Killer")
        End If
    End If
    txtTmplbl.SetFocus
End Sub

Private Sub cmbLanguage_Click()
    On Error Resume Next
    txtTmplbl.SetFocus
    
    With cmbLanguage
        setLanguage (.ListIndex)
        CreateToolTips (.ListIndex)
        Select Case .ListIndex
            Case 0
                lngLangIndex = 0
                .Clear
                .AddItem "Deutsch", 0
                .AddItem "Englisch", 1
            Case 1
                lngLangIndex = 1
                .Clear
                .AddItem "German", 0
                .AddItem "English", 1
        End Select
        SaveSetting App.EXEName, "Config", "Language", lngLangIndex
    End With
End Sub

Private Sub cmbLanguage_LostFocus()
    cmbLanguage.Text = cmbLanguage.List(lngLangIndex)
End Sub

Private Sub Command1_Click()
    txtTmplbl.SetFocus
    Me.Enabled = False
    If Command1.Caption = "&>>" Then
        If lngLangIndex = 0 Then
            m_cTT.ToolText(frmMainCode.Command1) = "Programm-Optionen" & vbCrLf & "ausblenden"
        Else
            m_cTT.ToolText(frmMainCode.Command1) = "Hide Program-" & vbCrLf & "Options"
        End If
        Me.Height = 5460
        MakeWindow Me, False
        Frame1.Visible = True
        Command1.Caption = "&<<"
    Else
        If lngLangIndex = 0 Then
            m_cTT.ToolText(frmMainCode.Command1) = "Und noch mehr" & vbCrLf & "Einstellungs-" & vbCrLf & "möglichkeiten"
        Else
            m_cTT.ToolText(frmMainCode.Command1) = "And still more" & vbCrLf & "program options"
        End If
        Me.Height = 1770
        MakeWindow Me, False
        Frame1.Visible = False
        Command1.Caption = "&>>"
    End If
    Me.Enabled = True
End Sub

Private Sub Command2_Click()
    MakeNormal (Me.hwnd)
    mnuTip_Click
    MakeTopMost (Me.hwnd)
End Sub

Private Sub Command3_Click()
    Dim t As String
    Dim s As String
    Dim b As Boolean
    
    MakeNormal (Me.hwnd)
    If lngLangIndex = 1 Then
        s = "Please type in the URL that" & vbCrLf & "is supposed to be included" & vbCrLf & "than 'popupfree':"
    Else
        s = "Bitte tragen Sie hier die URL" & vbCrLf & "ein, die als 'popupfree'" & vbCrLf & "aufgenommen werden soll:"
    End If
    t = ShowDialog(Me, s, IIf(lngLangIndex = 0, "Popupfree-Seite hinzufügen...", "Add to 'popup white list'..."), OK_ABBRECHEN, noICON, False, , , True)
    MakeTopMost (Me.hwnd)
    
    If t <> "" Then
        If Left$(t, 1) = 0 And Len(t) > 2 Then
            Dim i As Long
            For i = 0 To List1.ListCount - 1
                If List1.List(i) = Mid$(t, 3) Then
                    b = True
                    Exit For
                End If
            Next
            If Not b Then
                List1.AddItem (Mid$(t, 3))
                Call addListBoxToolTip
            End If
        End If
    End If
    createURLString
    txtTmplbl.SetFocus
End Sub

Private Sub Command4_Click()
    Dim i As Long
    Dim z As Long
    Dim s As String
    Dim k As String
    
    If List1.ListCount = 0 Then
        txtTmplbl.SetFocus
        Exit Sub
    End If
    
    z = 1000000
    
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            z = i
            Exit For
        End If
    Next
    
    If z <> 1000000 Then
        DelEntry
    Else
        If lngLangIndex = 0 Then
            s = "Möchten Sie wirklich" & vbCrLf & "alle Einträge löschen?"
            k = "Löschen aller Popupfree-Einträge..."
        Else
            s = "Do you really want to delete all" & vbCrLf & "entries of the 'popup white list'?"
            k = "Delete all entries in 'popup white list'..."
        End If
        MakeNormal (frmMainCode.hwnd)
        If ShowDialog(Me, s, k, JA_NEIN, QUESTION) = 2 Then
            While List1.ListCount > 0
                List1.RemoveItem (0)
            Wend
        End If
        MakeTopMost (frmMainCode.hwnd)
    End If
    Call createURLString
    Call addListBoxToolTip
    txtTmplbl.SetFocus
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    
    Dim lIEWnd As Long
    
    lIEWnd = FindWindowEx(0&, 0&, "IEFrame", vbNullString)
    Call addAktURLs(lIEWnd)
    Call createURLString
    txtTmplbl.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Dim i       As Long
    Dim t       As String
    Dim u       As Long
    Dim TmpStr  As String
    Dim Result  As Long
    
    'Call AddAppToIEMenu
    
    Set IEWin = New cIEWindows
    
    LinkCreate txtHyperlink, "Internet", "http://www.pc-tool.de", Me
    LinkCreate txtEMail, "e-mail", "mailto:toolwork@web.de", Me
    LinkDisplay txtHyperlink
    LinkDisplay txtEMail
  
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
    
    If App.PrevInstance Then End
    
    u = FileLen(App.Path & "\" & App.EXEName & ".exe")
    
    Dim ctrl As Control
    
    With m_cTT
        Call .Create(Me)
        .MaxTipWidth = 240
        .DelayTime(ttDelayShow) = 5000
        For Each ctrl In Controls
            Call .AddTool(ctrl)
        Next
    End With
    
    Xp = Screen.TwipsPerPixelX
    yp = Screen.TwipsPerPixelY

    With Me
        .Width = 2790
        .Height = 1770
        MakeWindow Me, False
        MakeTopMost (.hwnd)
        .Hide
        RegisterHotKey .hwnd, MIN_HOTKEY, MOD_WIN, VK_Z
        RegisterHotKey .hwnd, RST_HOTKEY, MOD_WIN + MOD_SHIFT, VK_Z
        lngOldWindowProc = SetWindowLong(.hwnd, GWL_WNDPROC, AddressOf SubProc)
    End With
  
    Active = True
    bSound = True
    
    Frame1.BackColor = Me.BackColor
    
    i = 1
    
    t = GetSetting(App.EXEName, "Config", "Sure")
    
    If Len(t) > 0 Then
        t = DeCompress(t)
        chkURLString = t
        While t <> ""
            If InStr(1, t, ";") > 0 Then
                TmpStr = Left$(t, InStr(1, t, ";") - 1)
            Else
                TmpStr = t
            End If
            List1.AddItem (TmpStr)
            t = Mid$(t, Len(TmpStr) + 2)
        Wend
    Else
        If GetSetting(App.EXEName, "Config", "First") = "" Then
            List1.AddItem ("www.microsoft.com")
            List1.AddItem ("www.freenet.de")
            chkURLString = "www.microsoft.com;www.freenet.de"
        End If
    End If
    
    addListBoxToolTip
    lngLangIndex = Val(GetSetting(App.EXEName, "Config", "Language"))
    setLanguage (lngLangIndex)
    CreateToolTips (lngLangIndex)
    
    With cmbLanguage
        Select Case lngLangIndex
            Case 0
                .Clear
                .AddItem "Deutsch", 0
                .AddItem "Englisch", 1
                .Text = cmbLanguage.List(lngLangIndex)
            Case 1
                .Clear
                .AddItem "German", 0
                .AddItem "English", 1
                .Text = cmbLanguage.List(lngLangIndex)
        End Select
    End With
    
    chkSound.Value = Val(GetSetting(App.EXEName, "Config", "Sound"))
    chkDial.Value = Val(GetSetting(App.EXEName, "Config", "NoDial"))
    Call ActivateDLKiller(chkDial.Value)
    
    Me.icon = LoadResPicture(101, vbResIcon)
    
    With TIcon
        .RemoveIcon picTray
        .ShowIcon picTray
        .ChangeToolTip picTray, (IIf(lngLangIndex = 0, "Real Popup-Killer ist aktiv", _
                                 "Real Popup-Killer is active"))
    End With

    lblTitle = "Real Popup-Killer"
    LoadSound
  
    Result = RegValueGet(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\" & IIf(IsWinNT, "Run", "RunServices"), "Real Popup-Killer", ap)
    If Result = 0 Then chkStartOnWin.Value = 1
    
    If fIsFileDIR(App.Path & "\Icons", vbDirectory) <> -1 Then
        MkDir App.Path & "\Icons"
    End If
    SavePicture LoadResPicture(101, vbResIcon), App.Path & "\Icons\HT.ico"
    SavePicture LoadResPicture(103, vbResIcon), App.Path & "\Icons\IC.ico"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    If lngLangIndex = 0 Then
        If frmMenuForm.mHideIE.Caption = "IE-Fenster anzeigen (Shift+Winkey+Z)" Then Call RestoreIE
    Else
        If frmMenuForm.mHideIE.Caption = "Show all IE-Windows (Shift+Winkey+Z)" Then Call RestoreIE
    End If
    
    SetWindowLong hwnd, GWL_WNDPROC, origWndProc
    Set IEWin = Nothing
    
    TIcon.RemoveIcon picTray
    Call SaveSettings
    
    Unload frmPics
    Unload frmDialog
    Unload frmMenuForm
    Unload frmLizenz
End Sub

Private Sub Form_Terminate()
    Call ActivateDLKiller(0)
End Sub

Private Sub imgTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    DoDragMe
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then frmMenuForm.mAktivis
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    Dim z As Long
    
    z = 1000000
    
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            z = i
            Exit For
        End If
    Next
    
    Select Case Button
        Case 2
            frmMenuForm.mDelEntry.Enabled = (z <> 1000000)
            PopupMenu frmMenuForm.mDel
    End Select
End Sub

Private Sub mExit_Click()
    Unload Me
End Sub

Private Sub mnuTip_Click()
    Select Case lngLangIndex
        Case 0
            ShowDialog Me, "Es gibt tatsächlich Seiten, die mit einem" & vbCrLf & _
                           "eigenen Popup-Fenster (keine Werbung)" & vbCrLf & _
                           "starten! Sollten Sie auf eine solche Seite" & vbCrLf & _
                           "treffen, dann halten Sie beim Öffnen der" & vbCrLf & _
                           "Seite einfach die 'Strg-Taste' gedrückt!" & vbCrLf & _
                           "Somit signalisieren Sie dem Programm, daß" & vbCrLf & _
                           "dieses Popup mit geöffnet werden darf!" & vbCrLf & _
                           "Wenn Sie eine solche Seite öfter ansurfen," & vbCrLf & _
                           "dann fügen Sie diese in die Liste:" & vbCrLf & _
                           "Seiten mit Popupstart ein!", "Wichtiger Hinweis!", OK, INFO
        Case 1
            ShowDialog Me, "There are Internet-pages which start with" & vbCrLf & _
                           "an own popup-window (no advertising) in fact!" & vbCrLf & _
                           "If you meet such side you simply have to hold" & vbCrLf & _
                           "the 'Strg-button' pressedly during the opening" & vbCrLf & _
                           "of the page! Thus you signal to the program" & vbCrLf & _
                           "that this popup may be opened with! If you" & vbCrLf & _
                           "surf such a page more often you have to add" & vbCrLf & _
                           "these URLs to the list below:" & vbCrLf & _
                           "'popup white list'", "Important reference!", OK, INFO
    End Select
    txtTmplbl.SetFocus
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Fehler
    Dim LI&, LX&, LY&
    
    LX = CLng(x / Screen.TwipsPerPixelX)
    LY = CLng(y / Screen.TwipsPerPixelY)
    
    LI = SendMessage(List1.hwnd, LB_ITEMFROMPOINT, 0, _
                     ByVal ((LY * 65536) + LX))
                     
    With List1
        If LI > -1 And LI < .ListCount Then
            Select Case True
                Case .ListCount > 4
                    If TextWidth(ToolTip(LI + 1)) > .Width - 250 Then
                        .ToolTipText = ToolTip(LI + 1)
                    Else
                        .ToolTipText = ""
                    End If
                Case Else
                    If TextWidth(ToolTip(LI + 1)) > .Width - 50 Then
                        .ToolTipText = ToolTip(LI + 1)
                    Else
                        .ToolTipText = ""
                    End If
            End Select
        Else
            .ToolTipText = ""
        End If
    End With
Fehler:
    Err.Clear
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    
    Dim TmpStr As String
    
    TmpStr = GetInfo
    
    If GetAsyncKeyState(VK_LBUTTON) Then
        If InStr(1, UCase(TmpStr), "INTERNET") Or InStr(1, UCase(TmpStr), "#32768") Or UCase(Left$(TmpStr, 3)) = "ATL" Then
            LB = True
            Zähler = 0
            Timer2.Enabled = True
            Timer1.Enabled = False
        Else
            LB = False
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    If LB Then LB = False
    If GetAsyncKeyState(VK_LBUTTON) <> -32768 Then
        Timer1.Enabled = True
        Timer2.Enabled = False
    End If
End Sub

Private Sub DoDragMe()
    On Error GoTo PROC_ERR
    
    DoDrag Me
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    ShowDialog Me, Err.Description, IIf(lngLangIndex = 0, "Fehler ...!" & Err.Number, "Error...!" & Err.Number), OK, CRITICAL
    Resume PROC_EXIT
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If TIcon.bRunningInTray Then
        Dim lngMsg As Long
        Dim Result As Long
        
        With Me
            Select Case x
                Case 7725:
                    .WindowState = vbNormal
                    .Show
                Case 7755:
                    SetForegroundWindow .hwnd
                    GetIPs
                    PopupMenu frmMenuForm.mFile
            End Select
        End With
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoDragMe
End Sub

Private Sub txtTmplbl_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtEMail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    LinkHover txtEMail, x, y
End Sub

Private Sub txtEMail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        Me.WindowState = vbMinimized
        Me.Hide
        LinkGo txtEMail
    End If
End Sub

Private Sub txtHyperlink_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LinkHover txtHyperlink, x, y
End Sub

Private Sub txtHyperlink_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        Me.WindowState = vbMinimized
        Me.Hide
        LinkGo txtHyperlink
    End If
End Sub

Private Sub AddAppToIEMenu()
    On Error GoTo Fehler
    mnuAddIE
Fehler:
End Sub
