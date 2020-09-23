VERSION 5.00
Begin VB.UserControl MyButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   DefaultCancel   =   -1  'True
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   74
   ToolboxBitmap   =   "MyButton.ctx":0000
End
Attribute VB_Name = "MyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long

Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_CENTERABS = &H65

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF = 4


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum ButtonTypes
    [Windows 16-bit] = 1
    [Windows 32-bit] = 2
    [Windows XP chrome] = 3
    [Mac] = 4
    [Java metal] = 5
    [Netscape 6] = 6
    [Simple Flat] = 7
End Enum

Public Enum ColorTypes
    [Use Windows] = 1
    [Custom] = 2
    [Force Standard] = 3
End Enum

Public Event Click()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private MyButtonType As ButtonTypes
Private MyColorType As ColorTypes

Private He As Long
Private Wi As Long

Private BackC As Long
Private ForeC As Long

Private tmpOldBold As Boolean
Private tmpOldWidth As Long

Private XPFace As Long
Private bAccess As Boolean

Private elTex As String
Private TextFont As StdFont

Private rc As RECT, rc2 As RECT, rc3 As RECT
Private rgnNorm As Long

Private m_OnButt As Boolean

Private LastButton As Byte
Private isEnabled As Boolean
Private isHoverFontBold As Boolean
Private isHoverResized As Boolean
Private isHover As Boolean

Private hasFocus As Boolean, showFocusR As Boolean

Private cFace As Long, cLight As Long, cHighLight As Long, cShadow As Long, cDarkShadow As Long, cText As Long

Private lastStat As Byte, TE As String

Private Sub ButtonClicked()
    Dim tm As Long
    
    bAccess = True
    tm = Timer
    
    While Timer - tm < 0.15
        DoEvents
    Wend
    bAccess = False
    Call UserControl_Click
    If m_OnButt Then
        If Hover Then
            If HoverFontBold Then HoverOver
            If HoverResize Then
                If tmpOldWidth = 0 Then HoverResizeIt
            End If
        End If
        Call Redraw(1, False)
        UserControl.FontBold = tmpOldBold
    Else
        Call Redraw(0, False)
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    bAccess = True
    ButtonClicked
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Click()
    If (LastButton = 1) And (isEnabled) Then
        RaiseEvent Click
        If bAccess Then
            Call Redraw(2, False)
            UserControl.Refresh
        End If
    End If
End Sub

Private Sub UserControl_DblClick()
    If LastButton = 1 Then Call UserControl_MouseDown(1, 1, 1, 1)
End Sub

Private Sub UserControl_GotFocus()
    hasFocus = True
    Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 32
            ButtonClicked
        Case 39, 40
            SendKeys "{Tab}"
        Case 37, 38
            SendKeys "+{Tab}"
    End Select
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        Call UserControl_MouseUp(1, 1, 1, 1)
        LastButton = 1
        Call UserControl_Click
    End If
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    hasFocus = False
    Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Initialize()
    LastButton = 1
    rc2.Left = 2: rc2.Top = 2
    Call SetColors
End Sub

Private Sub UserControl_InitProperties()
    isEnabled = True
    isHover = False
    isHoverFontBold = False
    isHoverResized = False
    showFocusR = True
    Set TextFont = UserControl.Font
    MyButtonType = [Windows 32-bit]
    MyColorType = [Use Windows]
    BackC = GetSysColor(COLOR_BTNFACE)
    ForeC = GetSysColor(COLOR_BTNTEXT)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastButton = Button
    m_OnButt = False
    If Button <> 2 Then
        Call Redraw(2, False)
    End If
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoverIt Button, Shift, x, y
End Sub

Function HoverIt(Button As Integer, Shift As Integer, x As Single, y As Single)
    With UserControl
        If GetCapture() <> .hwnd Then
            m_OnButt = False
            Call SetCapture(.hwnd)
        End If
        
        If Not (x <= 0 Or x >= ScaleWidth Or y <= 0 Or y >= ScaleHeight) Then
            If m_OnButt = False Then
                m_OnButt = True
                Call SetCapture(UserControl.hwnd)
                Select Case Button
                    Case 0
                        If Hover Then
                            If HoverFontBold Then HoverOver
                            If HoverResize Then
                                If tmpOldWidth = 0 Then HoverResizeIt
                            End If
                        End If
                        Call Redraw(1, False)
                        .FontBold = tmpOldBold
                End Select
            End If
        Else
            m_OnButt = False
            If GetCapture() = .hwnd Then
                If tmpOldBold <> .FontBold Then .FontBold = tmpOldBold
                If tmpOldWidth <> .Width Then HoverUnResizeIt
                Call ReleaseCapture
                Call Redraw(0, True)
            End If
        End If
    End With
    RaiseEvent MouseMove(Button, Shift, x, y)
End Function

Private Sub HoverOver()
    tmpOldBold = UserControl.FontBold
    UserControl.FontBold = True
End Sub

Private Sub HoverUnResizeIt()
    If tmpOldWidth <> 0 Then
        UserControl.Width = tmpOldWidth
        tmpOldWidth = 0
    End If
End Sub

Private Sub HoverResizeIt()
    tmpOldWidth = UserControl.Width
    UserControl.Width = UserControl.Width * 1.1
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Call Redraw(0, False)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = BackC
End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)
    BackC = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "BCOL"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = ForeC
End Property

Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
    ForeC = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "FCOL"
End Property

Public Property Get ButtonType() As ButtonTypes
    ButtonType = MyButtonType
End Property

Public Property Let ButtonType(ByVal newValue As ButtonTypes)
    MyButtonType = newValue
    Call Redraw(0, True)
    PropertyChanged "BTYPE"
End Property

Public Property Get Caption() As String
    Caption = elTex
End Property

Public Property Let Caption(ByVal newValue As String)
    elTex = newValue
    Call SetAccessKeys
    Call Redraw(0, True)
    PropertyChanged "TX"
End Property

Public Property Get Enabled() As Boolean
    Enabled = isEnabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    isEnabled = newValue
    Call Redraw(0, True)
    UserControl.Enabled = isEnabled
    PropertyChanged "ENAB"
End Property

Public Property Get HoverResize() As Boolean
    HoverResize = isHoverResized
End Property

Public Property Let HoverResize(ByVal newValue As Boolean)
    isHoverResized = newValue
    PropertyChanged "HOVRES"
End Property

Public Property Get HoverFontBold() As Boolean
    HoverFontBold = isHoverFontBold
End Property

Public Property Let HoverFontBold(ByVal newValue As Boolean)
    isHoverFontBold = newValue
    PropertyChanged "PRELIT"
End Property

Public Property Get Hover() As Boolean
    Hover = isHover
    If Not (Hover) Then
        isHoverResized = False
        isHoverFontBold = False
    End If
End Property

Public Property Let Hover(ByVal newValue As Boolean)
    isHover = newValue
    PropertyChanged "HOV"
End Property

Public Property Get Font() As Font
    Set Font = TextFont
End Property

Public Property Set Font(ByRef newFont As Font)
    Set TextFont = newFont
    Set UserControl.Font = TextFont
    Call Redraw(0, True)
    PropertyChanged "FONT"
End Property

Public Property Get ColorScheme() As ColorTypes
    ColorScheme = MyColorType
End Property

Public Property Let ColorScheme(ByVal newValue As ColorTypes)
    MyColorType = newValue
    Call SetColors
    Call Redraw(0, True)
    PropertyChanged "COLTYPE"
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = showFocusR
End Property

Public Property Let ShowFocusRect(ByVal newValue As Boolean)
    showFocusR = newValue
    Call Redraw(lastStat, True)
    PropertyChanged "FOCUSR"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_Resize()
    He = UserControl.ScaleHeight
    Wi = UserControl.ScaleWidth
    rc.Bottom = He: rc.Right = Wi
    rc2.Bottom = He: rc2.Right = Wi
    rc3.Left = 4: rc3.Top = 4: rc3.Right = Wi - 4: rc3.Bottom = He - 4
    
    DeleteObject rgnNorm
    Call MakeRegion
    SetWindowRgn UserControl.hwnd, rgnNorm, True
    
    If m_OnButt Then Exit Sub
    Call Redraw(0, True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        MyButtonType = .ReadProperty("BTYPE", 2)
        elTex = .ReadProperty("TX", "")
        isHover = .ReadProperty("HOV", True)
        isHoverResized = .ReadProperty("HOVRES", False)
        isEnabled = .ReadProperty("ENAB", True)
        isHoverFontBold = .ReadProperty("PRELIT", False)
        Set TextFont = .ReadProperty("FONT", UserControl.Font)
        MyColorType = .ReadProperty("COLTYPE", 1)
        showFocusR = .ReadProperty("FOCUSR", True)
        BackC = .ReadProperty("BCOL", GetSysColor(COLOR_BTNFACE))
        ForeC = .ReadProperty("FCOL", GetSysColor(COLOR_BTNTEXT))
    End With
    
    UserControl.Enabled = isEnabled
    Set UserControl.Font = TextFont
    Call SetColors
    Call SetAccessKeys
    Call Redraw(0, True)
End Sub

Private Sub UserControl_Terminate()
    DeleteObject rgnNorm
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BTYPE", MyButtonType)
        Call .WriteProperty("TX", elTex)
        Call .WriteProperty("HOV", isHover)
        Call .WriteProperty("HOVRES", isHoverResized)
        Call .WriteProperty("PRELIT", isHoverFontBold)
        Call .WriteProperty("ENAB", isEnabled)
        Call .WriteProperty("FONT", TextFont)
        Call .WriteProperty("COLTYPE", MyColorType)
        Call .WriteProperty("FOCUSR", showFocusR)
        Call .WriteProperty("BCOL", BackC)
        Call .WriteProperty("FCOL", ForeC)
    End With
End Sub

Public Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)
    
    If (curStat = 1 And Not (Hover)) Then Exit Sub
    If Not (Force) Then
        If (curStat = lastStat) And (TE = elTex) Then Exit Sub
    End If
    
    If He = 0 Then Exit Sub
    
    lastStat = curStat
    TE = elTex
    
    Dim i As Long, stepXP1 As Single, stepXP2 As Single
    Dim preFocusValue As Boolean
    
    preFocusValue = hasFocus
    If hasFocus = True Then hasFocus = ShowFocusRect
    
    With UserControl
        .Cls
        DrawRectangle 0, 0, Wi, He, cFace
        
        If isEnabled = True Then
            SetTextColor .hdc, cText
            
            Select Case curStat
            Case Is = 0
                Select Case MyButtonType
                    Case 1
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        UserControl.Line (1, 0)-(Wi - 1, 0), cDarkShadow
                        UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cDarkShadow
                        UserControl.Line (0, 1)-(0, He - 1), cDarkShadow
                        UserControl.Line (Wi - 1, 1)-(Wi - 1, He - 1), cDarkShadow
                        DrawRectangle 1, 1, Wi - 2, He - 2, cHighLight, True
                        DrawRectangle 2, 2, Wi - 4, He - 4, cHighLight, True
                        UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
                        UserControl.Line (Wi - 3, 2)-(Wi - 3, He - 1), cShadow
                        UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cShadow
                        UserControl.Line (2, He - 3)-(Wi - 2, He - 3), cShadow
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                    Case 2
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        If Ambient.DisplayAsDefault = True Then
                            DrawRectangle 1, 1, Wi - 2, He - 2, cHighLight, True
                            DrawRectangle 2, 2, Wi - 4, He - 4, cLight, True
                            UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cDarkShadow
                            UserControl.Line (Wi - 3, 2)-(Wi - 3, He - 1), cShadow
                            UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cDarkShadow
                            UserControl.Line (2, He - 3)-(Wi - 2, He - 3), cShadow
                            If hasFocus = True Then DrawFocusRect .hdc, rc3
                            DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                        Else
                            DrawRectangle 0, 0, Wi - 1, He - 1, cHighLight, True
                            DrawRectangle 1, 1, Wi - 2, He - 2, cLight, True
                            UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cDarkShadow
                            UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
                            UserControl.Line (0, He - 1)-(Wi - 1, He - 1), cDarkShadow
                            UserControl.Line (1, He - 2)-(Wi - 2, He - 2), cShadow
                        End If
                    Case 3
                            stepXP1 = 25 / He
                            XPFace = ShiftColor(cFace, &H30, True)
                            For i = 1 To He
                                DrawLine 0, i, Wi, i, ShiftColor(XPFace, -stepXP1 * i, True)
                            Next
                            SetTextColor UserControl.hdc, cText
                            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                            DrawLine 2, 0, Wi - 2, 0, &H733C00
                            DrawLine 2, He - 1, Wi - 2, He - 1, &H733C00
                            DrawLine 0, 2, 0, He - 2, &H733C00
                            DrawLine Wi - 1, 2, Wi - 1, He - 2, &H733C00
                            mSetPixel 1, 1, &H7B4D10
                            mSetPixel 1, He - 2, &H7B4D10
                            mSetPixel Wi - 2, 1, &H7B4D10
                            mSetPixel Wi - 2, He - 2, &H7B4D10
                        
                        If (hasFocus) Or ((Ambient.DisplayAsDefault) And (showFocusR)) Then
                            DrawRectangle 1, 2, Wi - 2, He - 4, &HE7AE8C, True
                            DrawLine 2, He - 2, Wi - 2, He - 2, &HEF826B
                            DrawLine 2, 1, Wi - 2, 1, &HFFE7CE
                            DrawLine 1, 2, Wi - 1, 2, &HF7D7BD
                            
                            DrawLine 2, 3, 2, He - 3, &HF0D1B5
                            DrawLine Wi - 3, 3, Wi - 3, He - 3, &HF0D1B5
                        Else
                            DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPFace, -&H30, True)
                            DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPFace, -&H20, True)
                            DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPFace, -&H24, True)
                            DrawLine Wi - 3, 3, Wi - 3, He - 3, ShiftColor(XPFace, -&H18, True)
                            DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPFace, &H10, True)
                            DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPFace, &HA, True)
                            DrawLine 1, 2, 1, He - 2, ShiftColor(XPFace, -&H5, True)
                            DrawLine 2, 3, 2, He - 3, ShiftColor(XPFace, -&HA, True)
                        End If
    
                    Case 4
                        DrawRectangle 1, 1, Wi - 2, He - 2, cLight
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        UserControl.Line (2, 0)-(Wi - 2, 0), cDarkShadow
                        UserControl.Line (2, He - 1)-(Wi - 2, He - 1), cDarkShadow
                        UserControl.Line (0, 2)-(0, He - 2), cDarkShadow
                        UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), cDarkShadow
                        SetPixel .hdc, 1, 1, cDarkShadow
                        SetPixel .hdc, 1, He - 2, cDarkShadow
                        SetPixel .hdc, Wi - 2, 1, cDarkShadow
                        SetPixel .hdc, Wi - 2, He - 2, cDarkShadow
                        SetPixel .hdc, 1, 2, cFace
                        SetPixel .hdc, 2, 1, cFace
                        UserControl.Line (3, 2)-(Wi - 3, 2), cHighLight
                        UserControl.Line (2, 2)-(2, He - 3), cHighLight
                        SetPixel .hdc, 3, 3, cHighLight
                        UserControl.Line (Wi - 3, 1)-(Wi - 3, He - 3), cFace
                        UserControl.Line (1, He - 3)-(Wi - 3, He - 3), cFace
                        SetPixel .hdc, Wi - 4, He - 4, cFace
                        UserControl.Line (Wi - 2, 3)-(Wi - 2, He - 2), cShadow
                        UserControl.Line (3, He - 2)-(Wi - 2, He - 2), cShadow
                        SetPixel .hdc, Wi - 3, He - 3, cShadow
                        SetPixel .hdc, 2, He - 2, cFace
                        SetPixel .hdc, 2, He - 3, cLight
                        SetPixel .hdc, Wi - 2, 2, cFace
                        SetPixel .hdc, Wi - 3, 2, cLight
                    Case 5 'Java
                        .FontBold = True
                        DrawRectangle 1, 1, Wi - 1, He - 1, ShiftColor(cFace, &HC)
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        DrawRectangle 1, 1, Wi - 1, He - 1, cHighLight, True
                        DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                        SetPixel .hdc, 1, He - 2, ShiftColor(cShadow, &H1A)
                        SetPixel .hdc, Wi - 2, 1, ShiftColor(cShadow, &H1A)
                        If hasFocus = True Then DrawRectangle (Wi - UserControl.TextWidth(elTex)) \ 2 - 3, (He - UserControl.TextHeight(elTex)) \ 2 - 1, UserControl.TextWidth(elTex) + 6, UserControl.TextHeight(elTex) + 2, &HCC9999, True
                        .FontBold = TextFont.Bold
                    Case 6 'Netscape
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        DrawRectangle 0, 0, Wi, He, ShiftColor(cLight, &H8), True
                        DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cLight, &H8), True
                        UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cShadow
                        UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
                        UserControl.Line (0, He - 1)-(Wi, He - 1), cShadow
                        UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cShadow
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                     Case 7 'Flat
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        DrawRectangle 0, 0, Wi, He, cHighLight, True
                        UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cShadow
                        UserControl.Line (0, He - 1)-(Wi, He - 1), cShadow
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                End Select
            
            Case Is = 1
            
                Select Case MyButtonType
                    Case 1
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        UserControl.Line (1, 0)-(Wi - 1, 0), cDarkShadow
                        UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cDarkShadow
                        UserControl.Line (0, 1)-(0, He - 1), cDarkShadow
                        UserControl.Line (Wi - 1, 1)-(Wi - 1, He - 1), cDarkShadow
                        DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                        DrawRectangle 2, 2, Wi - 4, He - 4, cShadow, True
                        UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cHighLight
                        UserControl.Line (Wi - 3, 2)-(Wi - 3, He - 1), cHighLight
                        UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cHighLight
                        UserControl.Line (2, He - 3)-(Wi - 2, He - 3), cHighLight
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                    Case 2
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                        DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                    Case 3
                        stepXP1 = 70 / He '85
                        stepXP2 = 80 / He '60
                        For i = 2 To He
                            UserControl.Line (0, i)-(Wi, i), &HFFFFFF - RGB(stepXP1 * i, stepXP1 * i, stepXP2 * i)
                        Next
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        DrawRectangle 0, 0, Wi, He, RGB(0, 60, 115), True
                        DrawRectangle 1, 1, Wi - 2, He - 2, RGB(33, 85, 132), True
                        DrawRectangle 1, 2, Wi - 2, He - 4, RGB(255, 178, 49), True
                        UserControl.Line (2, He - 2)-(Wi - 2, He - 2), RGB(231, 150, 0)
                        UserControl.Line (2, 1)-(Wi - 2, 1), RGB(255, 243, 206)
                        UserControl.Line (1, 2)-(Wi - 1, 2), RGB(255, 219, 140)
                        UserControl.Line (2, 3)-(2, He - 3), RGB(255, 199, 90)
                        UserControl.Line (Wi - 3, 3)-(Wi - 3, He - 3), RGB(255, 199, 90)
                        
                    Case 4
                        DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                        SetTextColor .hdc, cLight
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        UserControl.Line (2, 0)-(Wi - 2, 0), cDarkShadow
                        UserControl.Line (2, He - 1)-(Wi - 2, He - 1), cDarkShadow
                        UserControl.Line (0, 2)-(0, He - 2), cDarkShadow
                        UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), cDarkShadow
                        DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H40), True
                        DrawRectangle 2, 2, Wi - 4, He - 4, ShiftColor(cShadow, -&H20), True
                        SetPixel .hdc, 2, 2, ShiftColor(cShadow, -&H40)
                        SetPixel .hdc, 3, 3, ShiftColor(cShadow, -&H20)
                        SetPixel .hdc, 1, 1, cDarkShadow
                        SetPixel .hdc, 1, He - 2, cDarkShadow
                        SetPixel .hdc, Wi - 2, 1, cDarkShadow
                        SetPixel .hdc, Wi - 2, He - 2, cDarkShadow
                        UserControl.Line (Wi - 3, 1)-(Wi - 3, He - 3), cShadow
                        UserControl.Line (1, He - 3)-(Wi - 2, He - 3), cShadow
                        SetPixel .hdc, Wi - 4, He - 4, cShadow
                        UserControl.Line (Wi - 2, 3)-(Wi - 2, He - 2), ShiftColor(cShadow, -&H10)
                        UserControl.Line (3, He - 2)-(Wi - 2, He - 2), ShiftColor(cShadow, -&H10)
                        SetPixel .hdc, Wi - 2, He - 3, ShiftColor(cShadow, -&H20)
                        SetPixel .hdc, Wi - 3, He - 2, ShiftColor(cShadow, -&H20)
        
                        SetPixel .hdc, 2, He - 2, ShiftColor(cShadow, -&H20)
                        SetPixel .hdc, 2, He - 3, ShiftColor(cShadow, -&H10)
                        SetPixel .hdc, 1, He - 3, ShiftColor(cShadow, -&H10)
                        SetPixel .hdc, Wi - 2, 2, ShiftColor(cShadow, -&H20)
                        SetPixel .hdc, Wi - 3, 2, ShiftColor(cShadow, -&H10)
                        SetPixel .hdc, Wi - 3, 1, ShiftColor(cShadow, -&H10)
                    Case 5 'Java
                        .FontBold = True
                        DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, &H10), False
                        DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                        UserControl.Line (Wi - 1, 1)-(Wi - 1, He), cHighLight
                        UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cHighLight
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        If hasFocus = True Then DrawRectangle (Wi - UserControl.TextWidth(elTex)) \ 2 - 3, (He - UserControl.TextHeight(elTex)) \ 2 - 1, UserControl.TextWidth(elTex) + 6, UserControl.TextHeight(elTex) + 2, &HCC9999, True
                        .FontBold = TextFont.Bold
                    Case 6 'Netscape
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        DrawRectangle 0, 0, Wi, He, cShadow, True
                        DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                        UserControl.Line (Wi - 1, 0)-(Wi - 1, He), ShiftColor(cLight, &H8)
                        UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), ShiftColor(cLight, &H8)
                        UserControl.Line (0, He - 1)-(Wi, He - 1), ShiftColor(cLight, &H8)
                        UserControl.Line (1, He - 2)-(Wi - 1, He - 2), ShiftColor(cLight, &H8)
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                     Case 7 'Flat
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        DrawRectangle 0, 0, Wi, He, cShadow, True
                        UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cHighLight
                        UserControl.Line (0, He - 1)-(Wi - 1, He - 1), cHighLight
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                End Select
            
            Case Is = 2
            
                Select Case MyButtonType
                    Case 1
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        UserControl.Line (1, 0)-(Wi - 1, 0), cDarkShadow
                        UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cDarkShadow
                        UserControl.Line (0, 1)-(0, He - 1), cDarkShadow
                        UserControl.Line (Wi - 1, 1)-(Wi - 1, He - 1), cDarkShadow
                        DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                        DrawRectangle 2, 2, Wi - 4, He - 4, cShadow, True
                        UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cHighLight
                        UserControl.Line (Wi - 3, 2)-(Wi - 3, He - 1), cHighLight
                        UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cHighLight
                        UserControl.Line (2, He - 3)-(Wi - 2, He - 3), cHighLight
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                    Case 2
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                        DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                    Case 3
                        DrawRectangle 1, 1, Wi - 2, He - 2, RGB(214, 207, 198), True
                        DrawRectangle 2, 2, Wi - 4, He - 4, RGB(222, 219, 210), True
                        DrawRectangle 3, 3, Wi - 6, He - 6, RGB(231, 231, 222), False
                        DrawRectangle 0, 0, Wi, He, RGB(0, 60, 115), True
                        UserControl.Line (2, He - 2)-(Wi - 2, He - 2), RGB(247, 243, 239)
                        UserControl.Line (2, He - 3)-(Wi - 2, He - 3), RGB(239, 235, 231)
                        UserControl.Line (1, He - 4)-(Wi - 1, He - 4), RGB(231, 227, 221)
                        
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        
                        SetPixel .hdc, 1, 1, &H7B4D10
                        SetPixel .hdc, 1, He - 2, &H7B4D10
                        SetPixel .hdc, Wi - 2, 1, &H7B4D10
                        SetPixel .hdc, Wi - 2, He - 2, &H7B4D10
    
                    Case 4
                        DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                        SetTextColor .hdc, cLight
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        UserControl.Line (2, 0)-(Wi - 2, 0), cDarkShadow
                        UserControl.Line (2, He - 1)-(Wi - 2, He - 1), cDarkShadow
                        UserControl.Line (0, 2)-(0, He - 2), cDarkShadow
                        UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), cDarkShadow
                        DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H40), True
                        DrawRectangle 2, 2, Wi - 4, He - 4, ShiftColor(cShadow, -&H20), True
                        SetPixel .hdc, 2, 2, ShiftColor(cShadow, -&H40)
                        SetPixel .hdc, 3, 3, ShiftColor(cShadow, -&H20)
                        SetPixel .hdc, 1, 1, cDarkShadow
                        SetPixel .hdc, 1, He - 2, cDarkShadow
                        SetPixel .hdc, Wi - 2, 1, cDarkShadow
                        SetPixel .hdc, Wi - 2, He - 2, cDarkShadow
                        UserControl.Line (Wi - 3, 1)-(Wi - 3, He - 3), cShadow
                        UserControl.Line (1, He - 3)-(Wi - 2, He - 3), cShadow
                        SetPixel .hdc, Wi - 4, He - 4, cShadow
                        UserControl.Line (Wi - 2, 3)-(Wi - 2, He - 2), ShiftColor(cShadow, -&H10)
                        UserControl.Line (3, He - 2)-(Wi - 2, He - 2), ShiftColor(cShadow, -&H10)
                        SetPixel .hdc, Wi - 2, He - 3, ShiftColor(cShadow, -&H20)
                        SetPixel .hdc, Wi - 3, He - 2, ShiftColor(cShadow, -&H20)
                        SetPixel .hdc, 2, He - 2, ShiftColor(cShadow, -&H20)
                        SetPixel .hdc, 2, He - 3, ShiftColor(cShadow, -&H10)
                        SetPixel .hdc, 1, He - 3, ShiftColor(cShadow, -&H10)
                        SetPixel .hdc, Wi - 2, 2, ShiftColor(cShadow, -&H20)
                        SetPixel .hdc, Wi - 3, 2, ShiftColor(cShadow, -&H10)
                        SetPixel .hdc, Wi - 3, 1, ShiftColor(cShadow, -&H10)
                    Case 5 'Java
                        .FontBold = True
                        DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, &H10), False
                        DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                        UserControl.Line (Wi - 1, 1)-(Wi - 1, He), cHighLight
                        UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cHighLight
                        DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                        If hasFocus = True Then DrawRectangle (Wi - UserControl.TextWidth(elTex)) \ 2 - 3, (He - UserControl.TextHeight(elTex)) \ 2 - 1, UserControl.TextWidth(elTex) + 6, UserControl.TextHeight(elTex) + 2, &HCC9999, True
                        .FontBold = TextFont.Bold
                    Case 6
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        DrawRectangle 0, 0, Wi, He, cShadow, True
                        DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                        UserControl.Line (Wi - 1, 0)-(Wi - 1, He), ShiftColor(cLight, &H8)
                        UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), ShiftColor(cLight, &H8)
                        UserControl.Line (0, He - 1)-(Wi, He - 1), ShiftColor(cLight, &H8)
                        UserControl.Line (1, He - 2)-(Wi - 1, He - 2), ShiftColor(cLight, &H8)
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                     Case 7
                        DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                        DrawRectangle 0, 0, Wi, He, cShadow, True
                        UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cHighLight
                        UserControl.Line (0, He - 1)-(Wi - 1, He - 1), cHighLight
                        If hasFocus = True Then DrawFocusRect .hdc, rc3
                End Select
            End Select
        Else
            Select Case MyButtonType
                Case 1
                    SetTextColor .hdc, cHighLight
                    DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                    SetTextColor .hdc, cShadow
                    DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                    UserControl.Line (1, 0)-(Wi - 1, 0), cDarkShadow
                    UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cDarkShadow
                    UserControl.Line (0, 1)-(0, He - 1), cDarkShadow
                    UserControl.Line (Wi - 1, 1)-(Wi - 1, He - 1), cDarkShadow
                    DrawRectangle 1, 1, Wi - 2, He - 2, cHighLight, True
                    DrawRectangle 2, 2, Wi - 4, He - 4, cHighLight, True
                    UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
                    UserControl.Line (Wi - 3, 2)-(Wi - 3, He - 1), cShadow
                    UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cShadow
                    UserControl.Line (2, He - 3)-(Wi - 2, He - 3), cShadow
                Case 2
                    SetTextColor .hdc, cHighLight
                    DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                    SetTextColor .hdc, cShadow
                    DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                    DrawRectangle 0, 0, Wi - 1, He - 1, cHighLight, True
                    DrawRectangle 1, 1, Wi - 2, He - 2, cLight, True
                    UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cDarkShadow
                    UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
                    UserControl.Line (0, He - 1)-(Wi - 1, He - 1), cDarkShadow
                    UserControl.Line (1, He - 2)-(Wi - 2, He - 2), cShadow
                Case 3
                    stepXP1 = 60 / He
                    stepXP2 = 40 / He
                    For i = 1 To He
                        UserControl.Line (0, i)-(Wi, i), &HFFFFFF - RGB(stepXP1 * i, stepXP1 * i, stepXP2 * i)
                    Next
                    SetTextColor .hdc, cHighLight
                    DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                    SetTextColor .hdc, cShadow
                    DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                    UserControl.Line (2, 0)-(Wi - 2, 0), &H733C00
                    UserControl.Line (2, He - 1)-(Wi - 2, He - 1), &H733C00
                    UserControl.Line (0, 2)-(0, He - 2), &H733C00
                    UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), &H733C00
                    SetPixel .hdc, 1, 1, &H7B4D10
                    SetPixel .hdc, 1, He - 2, &H7B4D10
                    SetPixel .hdc, Wi - 2, 1, &H7B4D10
                    SetPixel .hdc, Wi - 2, He - 2, &H7B4D10
                    DrawRectangle 1, 2, Wi - 2, He - 4, &HFDBA99, True
                    UserControl.Line (2, He - 2)-(Wi - 2, He - 2), &HFE8A71
                    UserControl.Line (2, 1)-(Wi - 2, 1), &HFFEAD0
                    UserControl.Line (1, 2)-(Wi - 1, 2), &HFAD9BF
                Case 4
                    DrawRectangle 1, 1, Wi - 2, He - 2, cLight
                    SetTextColor .hdc, cHighLight
                    DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                    SetTextColor .hdc, cShadow
                    DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                    UserControl.Line (2, 0)-(Wi - 2, 0), cDarkShadow
                    UserControl.Line (2, He - 1)-(Wi - 2, He - 1), cDarkShadow
                    UserControl.Line (0, 2)-(0, He - 2), cDarkShadow
                    UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), cDarkShadow
                    SetPixel .hdc, 1, 1, cDarkShadow
                    SetPixel .hdc, 1, He - 2, cDarkShadow
                    SetPixel .hdc, Wi - 2, 1, cDarkShadow
                    SetPixel .hdc, Wi - 2, He - 2, cDarkShadow
                    SetPixel .hdc, 1, 2, cFace
                    SetPixel .hdc, 2, 1, cFace
                    UserControl.Line (3, 2)-(Wi - 3, 2), cHighLight
                    UserControl.Line (2, 2)-(2, He - 3), cHighLight
                    SetPixel .hdc, 3, 3, cHighLight
                    UserControl.Line (Wi - 3, 1)-(Wi - 3, He - 3), cFace
                    UserControl.Line (1, He - 3)-(Wi - 3, He - 3), cFace
                    SetPixel .hdc, Wi - 4, He - 4, cFace
                    UserControl.Line (Wi - 2, 3)-(Wi - 2, He - 2), cShadow
                    UserControl.Line (3, He - 2)-(Wi - 2, He - 2), cShadow
                    SetPixel .hdc, Wi - 3, He - 3, cShadow
                    SetPixel .hdc, 2, He - 2, cFace
                    SetPixel .hdc, 2, He - 3, cLight
                    SetPixel .hdc, Wi - 2, 2, cFace
                    SetPixel .hdc, Wi - 3, 2, cLight
                Case 5
                    .FontBold = True
                    SetTextColor .hdc, cShadow
                    DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                    DrawRectangle 0, 0, Wi, He, cShadow, True
                    .FontBold = TextFont.Bold
                Case 6
                    SetTextColor .hdc, cShadow
                    DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                    DrawRectangle 0, 0, Wi, He, ShiftColor(cLight, &H8), True
                    DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cLight, &H8), True
                    UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cShadow
                    UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
                    UserControl.Line (0, He - 1)-(Wi, He - 1), cShadow
                    UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cShadow
                Case 7
                    SetTextColor .hdc, cHighLight
                    DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                    SetTextColor .hdc, cShadow
                    DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                    DrawRectangle 0, 0, Wi, He, cHighLight, True
                    UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cShadow
                    UserControl.Line (0, He - 1)-(Wi - 1, He - 1), cShadow
            End Select
        End If
weiter:
    End With
    hasFocus = preFocusValue
End Sub

Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    Dim pt As POINTAPI
    
    UserControl.ForeColor = Color
    MoveToEx UserControl.hdc, X1, Y1, pt
    LineTo UserControl.hdc, X2, Y2
End Sub

Private Sub mSetPixel(ByVal x As Long, ByVal y As Long, ByVal Color As Long)
    Call SetPixel(UserControl.hdc, x, y, Color)
End Sub

Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
    Dim bRect As RECT
    Dim hBrush As Long
    Dim ret As Long
    
    bRect.Left = x
    bRect.Top = y
    bRect.Right = x + Width
    bRect.Bottom = y + Height
    
    hBrush = CreateSolidBrush(Color)
    
    If OnlyBorder = False Then
        ret = FillRect(UserControl.hdc, bRect, hBrush)
    Else
        ret = FrameRect(UserControl.hdc, bRect, hBrush)
    End If
    
    ret = DeleteObject(hBrush)
End Sub

Private Sub SetColors()
    Select Case MyColorType
        Case Is = Custom
            cFace = BackC
            cText = ForeC
            cShadow = ShiftColor(cFace, -&H40)
            cLight = ShiftColor(cFace, &H1F)
            cHighLight = ShiftColor(cFace, &H2F)
            cDarkShadow = ShiftColor(cFace, -&HC0)
        Case Is = [Force Standard]
            cFace = &HC0C0C0
            cShadow = &H808080
            cLight = &HDFDFDF
            cDarkShadow = &H0
            cHighLight = &HFFFFFF
            cText = &H0
        Case Else
            cFace = GetSysColor(COLOR_BTNFACE)
            cShadow = GetSysColor(COLOR_BTNSHADOW)
            cLight = GetSysColor(COLOR_BTNLIGHT)
            cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
            cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
            cText = GetSysColor(COLOR_BTNTEXT)
    End Select
End Sub

Private Sub MakeRegion()
    Dim rgn1 As Long
    Dim rgn2 As Long
    
    DeleteObject rgnNorm
    rgnNorm = CreateRectRgn(0, 0, Wi, He)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
    
    Select Case MyButtonType
        Case 1
            rgn1 = CreateRectRgn(0, 0, 1, 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(0, He, 1, He - 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(Wi, He, Wi - 1, He - 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
        Case 3, 4
            rgn1 = CreateRectRgn(0, 0, 2, 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(0, He, 2, He - 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(0, 1, 1, 2)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(0, He - 1, 1, He - 2)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(Wi, 1, Wi - 1, 2)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(Wi, He - 1, Wi - 1, He - 2)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
        Case 5 'Java
            rgn1 = CreateRectRgn(0, He, 1, He - 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
    End Select

    DeleteObject rgn2
End Sub

Private Sub SetAccessKeys()
    Dim ampersandPos As Long
    
    If Len(elTex) > 1 Then
        ampersandPos = InStr(1, elTex, "&", vbTextCompare)
        If (ampersandPos < Len(elTex)) And (ampersandPos > 0) Then
            If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then
                UserControl.AccessKeys = LCase(Mid$(elTex, ampersandPos + 1, 1))
            Else
                ampersandPos = InStr(ampersandPos + 2, elTex, "&", vbTextCompare)
                If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then
                    UserControl.AccessKeys = LCase(Mid$(elTex, ampersandPos + 1, 1))
                Else
                    UserControl.AccessKeys = ""
                End If
            End If
        Else
            UserControl.AccessKeys = ""
        End If
    Else
        UserControl.AccessKeys = ""
    End If
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
    Dim Red As Long, Blue As Long, Green As Long
    
    If Not (isXP) Then
        Blue = ((Color \ &H10000) Mod &H100) + Value
    Else
        Blue = ((Color \ &H10000) Mod &H100)
        Blue = Blue + ((Blue * Value) \ &HC0)
    End If
    Green = ((Color \ &H100) Mod &H100) + Value
    Red = (Color And &HFF) + Value
    
    Select Case Red
        Case Is < 0
            Red = 0
        Case Is > 255
            Red = 255
    End Select
    
    Select Case Green
        Case Is < 0
            Green = 0
        Case Is > 255
            Green = 255
    End Select
   
    Select Case Blue
        Case Is < 0
            Blue = 0
        Case Is > 255
            Blue = 255
    End Select

    ShiftColor = RGB(Red, Green, Blue)
End Function
