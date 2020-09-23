Attribute VB_Name = "modIEMenu"
Option Explicit

Type MENUITEMINFO
cbSize As Long
fMask As Long
fType As Long
fState As Long
wID As Long
hSubMenu As Long
hbmpChecked As Long
hbmpUnchecked As Long
dwItemData As Long
dwTypeData As String
cch As Long
End Type

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Public Const WM_MENUSELECT = &H11F
Public Const MF_SYSMENU = &H2000&
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20

Public origWndProc As Long

Public Sub SetHook(hwnd, bSet As Boolean)
    If bSet Then
        origWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf AppWndProc)
    ElseIf origWndProc Then
        Dim lRet As Long
        lRet = SetWindowLong(hwnd, GWL_WNDPROC, origWndProc)
        origWndProc = 0
    End If
End Sub

Public Function AppWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim iHi As Integer, iLo As Integer
    
    Select Case Msg
        Case WM_MENUSELECT
            CopyMemory iLo, wParam, 2
            CopyMemory iHi, ByVal VarPtr(wParam) + 2, 2
            If (iHi And MF_SYSMENU) = 0 Then
                Dim m As MENUITEMINFO, aCap As String
                m.dwTypeData = Space$(64)
                m.cbSize = Len(m)
                m.cch = 64
                m.fMask = MIIM_DATA Or MIIM_TYPE
                If GetMenuItemInfo(lParam, CLng(iLo), False, m) Then
                    aCap = m.dwTypeData & Chr$(0)
                    aCap = Left$(aCap, InStr(aCap, Chr$(0)) - 1)
                    If GetMenu(hwnd) <> lParam Then
                        Select Case aCap
                            Case "&Add to MSIE Tools Menu && Toolbar": 'Form1.Label1 = "Adds " & App.EXEName & " to Microsoft Internet Explorer's Tools Menu and an Icon on the Toolbar."
                            Case "&Delete from MSIE Tools Menu && Toolbar": 'Form1.Label1 = "Deletes " & App.EXEName & " From Microsoft Internet Explorer's Tools Menu and the Icon on the Toolbar."
                            Case "&Exit": 'Form1.Label1 = "Exit " & App.EXEName
                            Case Else: 'Form1.Label1 = aCap
                        End Select
                    Else
                    End If
               End If
           End If
    End Select
    AppWndProc = CallWindowProc(origWndProc, hwnd, Msg, wParam, lParam)
End Function

