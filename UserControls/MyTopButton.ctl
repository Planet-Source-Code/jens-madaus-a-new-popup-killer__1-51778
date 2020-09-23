VERSION 5.00
Begin VB.UserControl MyTopButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   HasDC           =   0   'False
   MaskColor       =   &H00C0C0C0&
   ScaleHeight     =   570
   ScaleWidth      =   405
   ToolboxBitmap   =   "MyTopButton.ctx":0000
   Begin VB.Image imDis 
      Height          =   315
      Left            =   1080
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imHot 
      Height          =   315
      Left            =   720
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imDown 
      Height          =   315
      Left            =   360
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imUp 
      Appearance      =   0  '2D
      Height          =   315
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "MyTopButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Const m_def_BackStyle = 1
Const m_def_MaskColor = &HC0C0C0
Const m_def_Enabled = True
Const m_def_Style = 0
Const m_def_Value = 0

Dim m_MaskColor As OLE_COLOR
Dim m_ImageUp As Picture
Dim m_Enabled As Boolean
Dim m_ImageDown As Picture
Dim m_ImageHot As Picture
Dim m_ImageDisabled As Picture

Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."

Public Enum VAL_U_P
    abUnPressed = 0
    abPressed = 1
End Enum

Private vval As VAL_U_P

Public Enum STYLE_B
    abCheckButton = 1
    abStandardButton = 0
End Enum

Private sstyle As STYLE_B

Public Enum BACKSTYLE_TO
    abTransparent = 0
    abOpaque = 1
End Enum

Private m_BackStyle As BACKSTYLE_TO

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    With UserControl
        If Not (m_Enabled) Then
            Set .Picture = imDis.Picture
            Set .MaskPicture = imDis.Picture
        Else
            Select Case vval
                Case Is = abUnPressed
                    Set .Picture = imUp.Picture
                    Set .MaskPicture = imUp.Picture
                Case Is = abPressed
                    Set .Picture = imDown.Picture
                    Set .MaskPicture = imDown.Picture
            End Select
        End If
    End With
    PropertyChanged "Enabled"
End Property

Public Property Get ImageDown() As Picture
    Set ImageDown = m_ImageDown
End Property

Public Property Set ImageDown(ByVal New_ImageDown As Picture)
    Set m_ImageDown = New_ImageDown
    Set imDown.Picture = New_ImageDown
    PropertyChanged "ImageDown"
End Property

Public Property Get ImageHot() As Picture
    Set ImageHot = m_ImageHot
End Property

Public Property Set ImageHot(ByVal New_ImageHot As Picture)
    Set m_ImageHot = New_ImageHot
    Set imHot.Picture = New_ImageHot
    PropertyChanged "ImageHot"
End Property

Public Property Get ImageDisabled() As Picture
    Set ImageDisabled = m_ImageDisabled
End Property

Public Property Set ImageDisabled(ByVal New_ImageDisabled As Picture)
    Set m_ImageDisabled = New_ImageDisabled
    Set imDis.Picture = New_ImageDisabled
    PropertyChanged "ImageDisabled"
End Property

Public Property Get Style() As STYLE_B
    Style = sstyle
End Property

Public Property Let Style(ByVal New_Style As STYLE_B)
    sstyle = New_Style
    PropertyChanged "Style"
    Set UserControl.Picture = imUp.Picture
    Set UserControl.MaskPicture = imUp.Picture
    vval = abUnPressed
End Property

Public Property Get Value() As VAL_U_P
    Value = vval
End Property

Public Property Let Value(ByVal New_Value As VAL_U_P)
    vval = New_Value
    PropertyChanged "Value"
    With UserControl
        Select Case vval
            Case Is = abUnPressed
                Set .Picture = imDown.Picture
                Set .MaskPicture = imDown.Picture
            Case Is = abUnPressed
                Set .Picture = imUp.Picture
                Set .MaskPicture = imUp.Picture
        End Select
        .Refresh
    End With
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_Enabled = False Then Exit Sub
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    With UserControl
        If Button = 1 Then
            If vval = abUnPressed Then
                Set .Picture = imDown.Picture
                Set .MaskPicture = imDown.Picture
            Else
                Set .Picture = imHot.Picture
                Set .MaskPicture = imHot.Picture
            End If
            .Refresh
        End If
    End With
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lp As POINTAPI
    Dim ret As Long
    Dim wn As Long
    
    If m_Enabled = False Then Exit Sub
    RaiseEvent MouseMove(Button, Shift, x, y)
    
    With UserControl
        Select Case Button
            Case Is = 1
                wn = GetCapture()
                If (wn <> .hwnd) And (wn <> 0) Then
                    If vval = abUnPressed Then
                        Set .Picture = imUp.Picture
                        Set .MaskPicture = imUp.Picture
                    Else
                        Set .Picture = imDown.Picture
                        Set .MaskPicture = imDown.Picture
                    End If
                    .Refresh
                    Exit Sub
                End If
                
                SetCapture .hwnd
                ret = GetCursorPos(lp)
                ScreenToClient .hwnd, lp
                lp.x = lp.x * Xp
                lp.y = lp.y * yp
                
                If lp.x > 0 And lp.x < .Width Then
                    If lp.y > 0 And lp.y < .Height Then
                        If vval = abUnPressed Then
                            Set .Picture = imDown.Picture
                            Set .MaskPicture = imDown.Picture
                        Else
                            Set .Picture = imHot.Picture
                            Set .MaskPicture = imHot.Picture
                        End If
                    Else
                        If vval = abUnPressed Then
                            Set .Picture = imUp.Picture
                            Set .MaskPicture = imUp.Picture
                        Else
                            Set .Picture = imDown.Picture
                            Set .MaskPicture = imDown.Picture
                        End If
                    End If
                Else
                    If vval = abUnPressed Then
                        Set .Picture = imUp.Picture
                        Set .MaskPicture = imUp.Picture
                    Else
                        Set .Picture = imDown.Picture
                        Set .MaskPicture = imDown.Picture
                    End If
                End If
            Case Is = 0
                If GetCapture() <> .hwnd Then SetCapture .hwnd
                
                ret = GetCursorPos(lp)
                ScreenToClient .hwnd, lp
                lp.x = lp.x * Xp
                lp.y = lp.y * yp
                
                If lp.x > 0 And lp.x < .Width Then
                    If lp.y > 0 And lp.y < .Height Then
                        If vval = abUnPressed Then
                            Set .Picture = imHot.Picture
                            Set .MaskPicture = imHot.Picture
                        End If
                    Else
                        ReleaseCapture
                        If vval = abUnPressed Then
                            Set .Picture = imUp.Picture
                            Set .MaskPicture = imUp.Picture
                        Else
                            Set .Picture = imDown.Picture
                            Set .MaskPicture = imDown.Picture
                        End If
                    End If
                Else
                    ReleaseCapture
                    If vval = abUnPressed Then
                        Set .Picture = imUp.Picture
                        Set .MaskPicture = imUp.Picture
                    Else
                        Set .Picture = imDown.Picture
                        Set .MaskPicture = imDown.Picture
                    End If
                End If
        End Select
        .Refresh
    End With
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Not (m_Enabled) Then Exit Sub
    RaiseEvent MouseUp(Button, Shift, x, y)
    
    With UserControl
        If Button = 1 Then
            If x > 0 And x < Extender.Width Then
                If y > 0 And y < Extender.Height Then
                    If sstyle = abCheckButton Then
                        If vval = abPressed Then vval = abUnPressed Else vval = abPressed
                        If vval = abUnPressed Then
                            Set .Picture = imHot.Picture
                            Set .MaskPicture = imHot.Picture
                        Else
                            Set .Picture = imDown.Picture
                            Set .MaskPicture = imDown.Picture
                        End If
                    Else
                        Set .Picture = imUp.Picture
                        Set .MaskPicture = imUp.Picture
                    End If
                Else
                    If vval = abUnPressed Then
                        Set .Picture = imUp.Picture
                        Set .MaskPicture = imUp.Picture
                    Else
                        Set .Picture = imDown.Picture
                        Set .MaskPicture = imDown.Picture
                    End If
                End If
            Else
                If vval = abUnPressed Then
                    Set .Picture = imUp.Picture
                    Set .MaskPicture = imUp.Picture
                Else
                    Set .Picture = imDown.Picture
                    Set .MaskPicture = imDown.Picture
                End If
            End If
        End If
    End With
    
    ReleaseCapture
    
End Sub

Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    Set m_ImageDown = LoadPicture("")
    Set m_ImageHot = LoadPicture("")
    Set m_ImageDisabled = LoadPicture("")
    Set m_ImageUp = LoadPicture("")
    sstyle = m_def_Style
    vval = m_def_Value
    m_MaskColor = m_def_MaskColor
    m_BackStyle = m_def_BackStyle
    Set UserControl.MaskPicture = LoadPicture("")
    UserControl.BackStyle = m_BackStyle
    UserControl.MaskColor = m_MaskColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        Set m_ImageDown = .ReadProperty("ImageDown", Nothing)
        Set m_ImageHot = .ReadProperty("ImageHot", Nothing)
        Set m_ImageDisabled = .ReadProperty("ImageDisabled", Nothing)
        sstyle = .ReadProperty("Style", m_def_Style)
        vval = .ReadProperty("Value", m_def_Value)
        Set m_ImageUp = .ReadProperty("ImageUp", Nothing)
        m_MaskColor = .ReadProperty("MaskColor", m_def_MaskColor)
        m_BackStyle = .ReadProperty("BackStyle", m_def_BackStyle)
    End With
        
    With UserControl
        .BackStyle = m_BackStyle
        Set imUp.Picture = m_ImageUp
        Set imDown.Picture = m_ImageDown
        Set imHot.Picture = m_ImageHot
        Set imDis.Picture = m_ImageDisabled
        If m_Enabled = True Then
            Select Case vval
                Case Is = abPressed
                    Set .Picture = imDown.Picture
                    Set .MaskPicture = imDown.Picture
                Case Is = abUnPressed
                    Set .Picture = imUp.Picture
                    Set .MaskPicture = imUp.Picture
                    Set imUp.Picture = m_ImageUp
            End Select
        Else
            Set .Picture = imDis.Picture
            Set .MaskPicture = imDis.Picture
        End If
    End With
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = imUp.Width
    UserControl.Height = imUp.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
        Call .WriteProperty("ImageDown", m_ImageDown, Nothing)
        Call .WriteProperty("ImageHot", m_ImageHot, Nothing)
        Call .WriteProperty("ImageDisabled", m_ImageDisabled, Nothing)
        Call .WriteProperty("ImageUp", m_ImageUp, Nothing)
        Call .WriteProperty("Style", sstyle, m_def_Style)
        Call .WriteProperty("Value", vval, m_def_Value)
        Call .WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)
        Call .WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    End With
End Sub

Public Property Get ImageUp() As Picture
Attribute ImageUp.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set ImageUp = m_ImageUp
End Property

Public Property Set ImageUp(ByVal New_ImageUp As Picture)
    Set m_ImageUp = New_ImageUp
    Set imUp.Picture = New_ImageUp
    PropertyChanged "ImageUp"
    UserControl.BackStyle = 1
    Set UserControl.Picture = imUp.Picture
    Set UserControl.MaskPicture = imUp.Picture
    DoEvents
    Extender.Width = imUp.Width
    Extender.Height = imUp.Height
    UserControl.BackStyle = m_BackStyle
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the Picture."
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
    UserControl.MaskColor = m_MaskColor
    Set UserControl.MaskPicture = m_ImageUp
End Property

Public Property Get BackStyle() As BACKSTYLE_TO
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BACKSTYLE_TO)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
    UserControl.BackStyle = m_BackStyle
    Set UserControl.MaskPicture = UserControl.Picture
    UserControl.MaskColor = m_MaskColor
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property
