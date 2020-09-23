VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
  
Private FormHandle As Long
Private mvarbRunningInTray As Boolean
Private SysIcon As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Property Let bRunningInTray(ByVal vData As Boolean)
    mvarbRunningInTray = vData
End Property

Property Get bRunningInTray() As Boolean
    bRunningInTray = mvarbRunningInTray
End Property

Public Sub ShowIcon(ByRef sysTrayForm)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = 512
    SysIcon.hIcon = frmMainCode.icon
    SysIcon.szTip = frmMainCode.Caption + Chr(0)
    Shell_NotifyIcon 0, SysIcon
    mvarbRunningInTray = True
End Sub

Public Sub RemoveIcon(sysTrayForm)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = vbNull
    SysIcon.hIcon = frmMainCode.icon
    SysIcon.szTip = Chr(0)
    Shell_NotifyIcon 2, SysIcon
    mvarbRunningInTray = False
End Sub

Public Sub ChangeIcon(sysTrayForm, icon As Image)
    If mvarbRunningInTray = True Then
        SysIcon.cbSize = Len(SysIcon)
        SysIcon.hwnd = sysTrayForm.hwnd
        SysIcon.hIcon = frmMainCode.icon
        Shell_NotifyIcon 1, SysIcon
    End If
End Sub

Public Sub ChangeToolTip(sysTrayForm, strNewTip As String)
    If mvarbRunningInTray = True Then
        SysIcon.cbSize = Len(SysIcon)
        SysIcon.hwnd = sysTrayForm.hwnd
        SysIcon.szTip = strNewTip & Chr(0)
        Shell_NotifyIcon 1, SysIcon
    End If
End Sub
