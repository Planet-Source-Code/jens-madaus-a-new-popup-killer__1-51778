Attribute VB_Name = "modSkinForm"
Option Explicit

Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
    x As Long
    y As Long
End Type

Dim m_intResizable As Integer

Public Sub AlwaysOnTop(p_TheForm As Form, p_blnToggle As Boolean)
    On Error GoTo PROC_ERR
    
    If p_blnToggle Then
        SetWindowPos p_TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    Else
        SetWindowPos p_TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    ShowDialog frmMainCode, Err.Description, IIf(lngLangIndex = 0, "Fehler ...!" & Err.Number, "Error...!" & Err.Number), OK, CRITICAL
    Resume PROC_EXIT
End Sub

Public Sub DoDrag(p_TheForm As Form)
    On Error GoTo PROC_ERR
    
    If p_TheForm.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage p_TheForm.hwnd, &HA1, 2, 0&
    End If
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    ShowDialog frmMainCode, Err.Description, IIf(lngLangIndex = 0, "Fehler ...!" & Err.Number, "Error...!" & Err.Number), OK, CRITICAL
    Resume PROC_EXIT
End Sub

Public Sub DoTransparency(p_TheForm As Form)
    On Error GoTo PROC_ERR
    
    Dim alngTempRegions(6) As Long
    Dim lngFormWidthInPixels As Long
    Dim lngFormHeightInPixels As Long
    Dim varA
    
    lngFormWidthInPixels = p_TheForm.Width / Xp
    lngFormHeightInPixels = p_TheForm.Height / yp
    varA = CreateRoundRectRgn(0, 0, lngFormWidthInPixels, lngFormHeightInPixels, 24, 24)
    Call SetWindowRgn(p_TheForm.hwnd, varA, True)
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    ShowDialog frmMainCode, Err.Description, IIf(lngLangIndex = 0, "Fehler ...!" & Err.Number, "Error...!" & Err.Number), OK, CRITICAL
    Resume PROC_EXIT
End Sub

Public Sub MakeWindow(p_TheForm As Form, p_blnIsResizable As Boolean)
    On Error Resume Next
    
    Dim lngFormWidth As Long
    Dim lngFormHeight As Long
    Dim intTemp As Integer
    
    With p_TheForm
        m_intResizable = Val(p_blnIsResizable)
        lngFormWidth = (.Width / Xp)
        lngFormHeight = (.Height / yp)
        .BackColor = &HC6C7C6
        .Caption = p_TheForm!lblTitle.Caption
        !lblTitle.Left = 11 * Xp
        !lblTitle.Top = 9 * yp
        DoTransparency p_TheForm
        
        !imgTitle(0).Top = 0
        !imgTitle(0).Left = 0
        !imgTitle(2).Top = 0
        !imgTitle(2).Left = (lngFormWidth - 19) * Xp
        !imgTitle(1).Top = 0
        !imgTitle(1).Left = 19 * Xp
        !imgTitle(1).Width = Screen.Width
    
        !imgWindowLeft.Top = 30 * yp
        !imgWindowLeft.Left = 0
        !imgWindowLeft.Height = (lngFormHeight - 60) * yp
    
        !imgWindowBottomLeft.Top = (lngFormHeight - 30) * yp
        !imgWindowBottomLeft.Left = 0
    
        !imgWindowBottom.Top = (lngFormHeight - 30) * yp
        !imgWindowBottom.Left = 19 * Xp
        !imgWindowBottom.Width = (lngFormWidth - 38) * Xp
    
        !imgWindowBottomRight.Top = (lngFormHeight - 30) * yp
        !imgWindowBottomRight.Left = (lngFormWidth - 19) * Xp
    
        !imgWindowRight.Top = 30 * yp
        !imgWindowRight.Left = (lngFormWidth - 19) * Xp
        !imgWindowRight.Height = (lngFormHeight - 38) * yp
    
        !btnClose.Top = 7 * yp
        !btnClose.Left = (lngFormWidth - 29) * Xp
   
        !btnMinimize.Top = 7 * yp
        !btnMinimize.Left = (lngFormWidth - 52) * Xp
    
        !btnMinimizeTray.Top = 7 * yp
        !btnMinimizeTray.Left = (lngFormWidth - 75) * Xp
    End With
End Sub
