Attribute VB_Name = "modHideIEWindows"
Option Explicit

Public Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Public Type IE_STATE_SAVE
    hwnd As Long
    wp As WINDOWPLACEMENT
End Type

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassNameA Lib "user32" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemoryH Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal Length As Long)

Global Const GWL_WNDPROC = (-4)
Global Const SW_HIDE = 0
Global Const WM_HOTKEY = &H312
Public Const MOD_SHIFT = &H4
Global Const MOD_WIN = &H8
Global Const VK_Z = &H5A

Global Const MIN_HOTKEY = &H5F
Global Const RST_HOTKEY = &H6F

Public lngOldWindowProc As Long
Private arrayIESS() As IE_STATE_SAVE

Public Function EnumWindowsIE() As Boolean
    Dim l As Long
    Dim sClsNm As String
    Dim ss As IE_STATE_SAVE
    Dim wp As WINDOWPLACEMENT
    
    wp.Length = Len(wp)
    
    For Each var In SWs
        If TypeOf var.Document Is HTMLDocument Then
            GetWindowPlacement var.hwnd, wp
            ss.hwnd = var.hwnd
            ss.wp = wp
            l = UBound(arrayIESS)
            arrayIESS(l) = ss
            ReDim Preserve arrayIESS(l + 1)
            wp.showCmd = SW_HIDE
            SetWindowPlacement var.hwnd, wp
        End If
    Next
    
    EnumWindowsIE = True
End Function

Public Function SubProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If wMsg = WM_HOTKEY Then
        If LoWord(lParam) = MOD_WIN And HiWord(lParam) = VK_Z And wParam = MIN_HOTKEY Then
            MinimizeIE
        End If
        If LoWord(lParam) = (MOD_WIN + MOD_SHIFT) And HiWord(lParam) = VK_Z And wParam = RST_HOTKEY Then
            RestoreIE
        End If
    End If
    SubProc = CallWindowProc(lngOldWindowProc, hwnd, wMsg, wParam, lParam)
End Function

Public Sub MinimizeIE()
    If lngLangIndex = 0 Then
        frmMenuForm.mHideIE.Caption = "IE-Fenster anzeigen (Shift+Winkey+Z)"
    Else
        frmMenuForm.mHideIE.Caption = "Show all IE-Windows (Shift+Winkey+Z)"
    End If
    
    ReDim arrayIESS(0)
    EnumWindowsIE
End Sub

Public Sub RestoreIE()
    On Error GoTo exit_RestoreIE
    
    Dim ieSS    As IE_STATE_SAVE
    Dim l       As Long

    If lngLangIndex = 0 Then
        frmMenuForm.mHideIE.Caption = "IE-Fenster verstecken (Winkey+Z)"
    Else
        frmMenuForm.mHideIE.Caption = "Hide all IE-Windows (Winkey+Z)"
    End If
    
    For l = UBound(arrayIESS) To LBound(arrayIESS) Step -1
        ieSS = arrayIESS(l)
        With ieSS
            If .hwnd > 0 Then
                SetWindowPlacement .hwnd, .wp
            End If
        End With
    Next
exit_RestoreIE:
    Exit Sub
End Sub

Private Function GetClassName(ByVal hwnd As Long) As String
    Dim lngReturn As Long
    Dim strReturn As String
    
    strReturn = Space(255)
    lngReturn = GetClassNameA(hwnd, strReturn, Len(strReturn))
    GetClassName = Left$(strReturn, lngReturn)
End Function

Public Function LoWord(ByVal dw As Long) As Integer
On Error GoTo Err_LoWord

    CopyMemoryH LoWord, ByVal VarPtr(dw), 2

Exit_LoWord:
    Exit Function

Err_LoWord:
    GoTo Exit_LoWord

End Function

Public Function HiWord(ByVal dw As Long) As Integer
On Error GoTo Err_HiWord

    CopyMemoryH HiWord, ByVal VarPtr(dw) + 2, 2

Exit_HiWord:
    Exit Function
Err_HiWord:
    GoTo Exit_HiWord
End Function
