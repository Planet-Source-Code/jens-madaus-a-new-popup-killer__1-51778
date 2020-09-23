VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ttDelayTimeConstants
    ttDelayDefault = TTDT_AUTOMATIC '= 0
    ttDelayInitial = TTDT_INITIAL '= 3
    ttDelayShow = TTDT_AUTOPOP '= 2
    ttDelayReshow = TTDT_RESHOW '= 1
    ttDelayMask = 3
End Enum

Public Enum ttMarginConstants
    ttMarginLeft = 0
    ttMarginTop = 1
    ttMarginRight = 2
    ttMarginBottom = 3
End Enum

Private m_hwndTT As Long
Private m_cMaxTip As Long

Public Function Create(frm As Form) As Boolean
    If (m_hwndTT = 0) Then
        Call InitCommonControls
        m_hwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP, 0, 0, _
                                  0, 0, frm.hwnd, 0, App.hInstance, ByVal 0)
    End If
    Create = CBool(m_hwndTT)
End Function

Private Sub Class_Terminate()
    If m_hwndTT Then Call DestroyWindow(m_hwndTT)
End Sub

Public Function AddTool(ctrl As Control, Optional sText As String) As Boolean
    On Error GoTo Fehler
    
    Dim ti As TOOLINFO
    
    If (m_hwndTT = 0) Then Exit Function
    If (GetToolInfo(ctrl.hwnd, ti) = False) Then
        With ti
            .cbSize = Len(ti)
            .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
            .hwnd = ctrl.Container.hwnd
            .uId = ctrl.hwnd
            If Len(sText) Then
                .lpszText = sText
            Else
                .lpszText = ""
            End If
            m_cMaxTip = Max(m_cMaxTip, Len(.lpszText) + 1)
        End With
        AddTool = SendMessage(m_hwndTT, TTM_ADDTOOL, 0, ti)
    End If
Fehler:
End Function

Public Function RemoveTool(ctrl As Control) As Boolean
    Dim ti As TOOLINFO
    
    If (m_hwndTT = 0) Then Exit Function
    If GetToolInfo(ctrl.hwnd, ti) Then
        Call SendMessage(m_hwndTT, TTM_DELTOOL, 0, ti)
        RemoveTool = True
    End If
End Function

Public Property Get BackColor() As OLE_COLOR
    If (m_hwndTT = 0) Then Exit Property
    BackColor = SendMessage(m_hwndTT, TTM_GETTIPBKCOLOR, 0, 0)
End Property

Public Property Let BackColor(clr As OLE_COLOR)
    If (m_hwndTT = 0) Then Exit Property
    Call SendMessage(m_hwndTT, TTM_SETTIPBKCOLOR, clr, 0)
End Property

Public Property Get DelayTime(dwType As ttDelayTimeConstants) As Long
    If (m_hwndTT = 0) Then Exit Property
    DelayTime = SendMessage(m_hwndTT, TTM_GETDELAYTIME, (dwType And ttDelayMask), 0&)
End Property

Public Property Let DelayTime(dwType As ttDelayTimeConstants, dwMilliSecs As Long)
    If (m_hwndTT = 0) Then Exit Property
    Call SendMessage(m_hwndTT, TTM_SETDELAYTIME, (dwType And ttDelayMask), ByVal dwMilliSecs)
End Property

Public Property Get ForeColor() As OLE_COLOR
    If (m_hwndTT = 0) Then Exit Property
    ForeColor = SendMessage(m_hwndTT, TTM_SETTIPTEXTCOLOR, 0, 0)
End Property

Public Property Let ForeColor(clr As OLE_COLOR)
    If (m_hwndTT = 0) Then Exit Property
    Call SendMessage(m_hwndTT, TTM_SETTIPTEXTCOLOR, clr, 0)
End Property

Public Property Get hwnd() As Long
    hwnd = m_hwndTT
End Property

Public Property Get Margin(dwType As ttMarginConstants) As Long
    Dim rc As RECT
    
    If (m_hwndTT = 0) Then Exit Property
    Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, rc)
    Select Case dwType
        Case ttMarginLeft:      Margin = rc.Left
        Case ttMarginTop:       Margin = rc.Top
        Case ttMarginRight:     Margin = rc.Right
        Case ttMarginBottom:    Margin = rc.Bottom
    End Select
End Property

Public Property Let Margin(dwType As ttMarginConstants, cPixels As Long)
    Dim rc As RECT
    
    If (m_hwndTT = 0) Then Exit Property
    Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, rc)
    Select Case dwType
        Case ttMarginLeft:      rc.Left = cPixels
        Case ttMarginTop:       rc.Top = cPixels
        Case ttMarginRight:     rc.Right = cPixels
        Case ttMarginBottom:    rc.Bottom = cPixels
    End Select
    Call SendMessage(m_hwndTT, TTM_SETMARGIN, 0, rc)
End Property

Public Property Get MaxTipWidth() As Integer
    If (m_hwndTT = 0) Then Exit Property
    MaxTipWidth = LoWord(SendMessage(m_hwndTT, TTM_GETMAXTIPWIDTH, 0, 0))
End Property

Public Property Let MaxTipWidth(ByVal cx As Integer)
    If (m_hwndTT = 0) Then Exit Property
    If (cx < 1) Then cx = -1
    Call SendMessage(m_hwndTT, TTM_SETMAXTIPWIDTH, 0, ByVal CLng(cx))
End Property

Public Property Get ToolCount() As Integer
    If (m_hwndTT = 0) Then Exit Property
    ToolCount = SendMessage(m_hwndTT, TTM_GETTOOLCOUNT, 0, 0)
End Property

Public Property Get ToolText(ctrl As Control) As String
    Dim ti As TOOLINFO
    
    If (m_hwndTT = 0) Then Exit Property
    If GetToolInfo(ctrl.hwnd, ti, True) Then
        ToolText = GetStrFromBufferA(ti.lpszText)
    End If
End Property

Public Property Let ToolText(ctrl As Control, sText As String)
    Dim ti As TOOLINFO
    
    If (m_hwndTT = 0) Then Exit Property
    If GetToolInfo(ctrl.hwnd, ti) Then
        ti.lpszText = sText
        m_cMaxTip = Max(m_cMaxTip, Len(sText) + 1)
        Call SendMessage(m_hwndTT, TTM_UPDATETIPTEXT, 0, ti)
    End If
End Property

Private Function IsWindow(ctrl As Control) As Boolean
    On Error GoTo Out
    IsWindow = CBool(ctrl.hwnd)
Out:
End Function

Private Function GetToolInfo(hwndTool As Long, ti As TOOLINFO, Optional fGetText As Boolean = False) As Boolean
    Dim nItems As Integer
    Dim i As Integer
    
    ti.cbSize = Len(ti)
    If fGetText Then ti.lpszText = String$(m_cMaxTip, 0)
    nItems = ToolCount
    For i = 0 To nItems - 1
        If SendMessage(m_hwndTT, TTM_ENUMTOOLS, (i), ti) Then
            If (hwndTool = ti.uId) Then
                GetToolInfo = True
                Exit Function
            End If
        End If
    Next
End Function
