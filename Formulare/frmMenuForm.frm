VERSION 5.00
Begin VB.Form frmMenuForm 
   Caption         =   "Menu"
   ClientHeight    =   90
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   2070
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.Menu mFile 
      Caption         =   "Programm"
      Begin VB.Menu mnuShowMe 
         Caption         =   "Zeige Real Popup-Killer"
      End
      Begin VB.Menu mnuNix 
         Caption         =   "-"
      End
      Begin VB.Menu mAktiv 
         Caption         =   "aktiv"
         Checked         =   -1  'True
      End
      Begin VB.Menu mSound 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mNix0 
         Caption         =   "-"
      End
      Begin VB.Menu mHideIE 
         Caption         =   "IE-Fenster verstecken (Winkey+Z)"
      End
      Begin VB.Menu mNix1 
         Caption         =   "-"
      End
      Begin VB.Menu mAddAktURL 
         Caption         =   "aktuelle URL hinzufügen"
         Enabled         =   0   'False
      End
      Begin VB.Menu mNix2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIPs 
         Caption         =   "IP's"
         Begin VB.Menu mnuOwn 
            Caption         =   "OwnIP1"
            Index           =   0
         End
         Begin VB.Menu mnuOwn 
            Caption         =   "OwnIP2"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOwn 
            Caption         =   "OwnIP3"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOwn 
            Caption         =   "OwnIP4"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOwn 
            Caption         =   "OwnIP5"
            Index           =   4
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuNull2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Beenden"
      End
   End
   Begin VB.Menu mDel 
      Caption         =   "ListDel"
      Begin VB.Menu mDelEntry 
         Caption         =   "Eintrag löschen"
      End
      Begin VB.Menu mNix3 
         Caption         =   "-"
      End
      Begin VB.Menu mAktURL 
         Caption         =   "URL hinzufügen"
      End
   End
End
Attribute VB_Name = "frmMenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(p_intCancel As Integer)
    
    On Error GoTo PROC_ERR
    
    Set frmMenuForm = Nothing

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    ShowDialog frmMainCode, Err.Description, IIf(lngLangIndex = 0, "Fehler ...!" & Err.Number, "Error...!" & Err.Number), OK, CRITICAL
    Resume PROC_EXIT
End Sub

Private Sub mnu_Sound_Click()
    mSound.Checked = Not mSound.Checked
    bSound = (mSound.Checked)
    frmMainCode.chkSound.Value = Val(mSound.Checked)
End Sub

Private Sub mnuDelEntry_Click()
    DelEntry
End Sub

Function mAktivis()
    mAktiv_Click
End Function

Private Sub mAddAktURL_Click()
    Dim lIEWnd As Long
    
    lIEWnd = FindWindowEx(0&, 0&, "IEFrame", vbNullString)
    Call addAktURLs(lIEWnd)
    Call createURLString
End Sub

Private Sub mAktiv_Click()
    On Error Resume Next
    With mAktiv
        If .Checked Then
            .Checked = False
            Active = False
            With frmMainCode.TIcon
                '.RemoveIcon frmMainCode.picTray
                frmMainCode.icon = LoadResPicture(102, vbResIcon)
                .ShowIcon frmMainCode.picTray
                .ChangeToolTip frmMainCode.picTray, ("Real Popup-Killer " & IIf(lngLangIndex = 0, "ist inaktiv", "is inactive"))
            End With
            frmMainCode.lblInfo.Caption = "Status: " & IIf(lngLangIndex = 0, "inaktiv", "inactive")
            If frmMainCode.chkDial.Value = 1 Then Call ActivateDLKiller(0)
        Else
            .Checked = True
            Active = True
            With frmMainCode.TIcon
                '.RemoveIcon frmMainCode.picTray
                frmMainCode.icon = LoadResPicture(101, vbResIcon)
                .ShowIcon frmMainCode.picTray
                .ChangeToolTip frmMainCode.picTray, ("Real Popup-Killer " & IIf(lngLangIndex = 0, "ist aktiv", "is active"))
            End With
            frmMainCode.lblInfo.Caption = "Status: " & IIf(lngLangIndex = 0, "aktiv", "active")
            If frmMainCode.chkDial.Value = 1 Then Call ActivateDLKiller(1)
        End If
    End With
    With frmMainCode
        .Timer1.Enabled = Active
        .Timer2.Enabled = Active
    End With
End Sub

Private Sub mAktURL_Click()
    Call addURL
    Call createURLString
End Sub

Private Sub mDelEntry_Click()
    Call DelEntry
    Call createURLString
End Sub

Private Sub mExit_Click()
    Unload frmMainCode
End Sub

Private Sub mHideIE_Click()
    If lngLangIndex = 0 Then
        If mHideIE.Caption = "IE-Fenster verstecken (Winkey+Z)" Then
            Call MinimizeIE
        Else
            Call RestoreIE
        End If
    Else
        If mHideIE.Caption = "Hide all IE-Windows (Winkey+Z)" Then
            Call MinimizeIE
        Else
            Call RestoreIE
        End If
    End If
End Sub

Private Sub mnuShowMe_Click()
    On Error Resume Next
    
    With frmMainCode
        .Visible = True
        .WindowState = vbNormal
    End With
End Sub

Private Sub mSound_Click()
    mSound.Checked = Not (mSound.Checked)
    frmMainCode.SoundOn IIf(mSound.Checked, 1, 0)
End Sub
