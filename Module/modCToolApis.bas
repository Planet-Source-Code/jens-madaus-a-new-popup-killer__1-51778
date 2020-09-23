Attribute VB_Name = "modCToolApis"
Option Explicit

Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code  As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
                             ByVal wParam As Long, lParam As Any) As Long

Public Const WM_USER = &H400

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
                            (ByVal dwExStyle As Long, ByVal lpClassName As String, _
                             ByVal lpWindowName As String, ByVal dwStyle As Long, _
                             ByVal x As Long, ByVal y As Long, _
                             ByVal nWidth As Long, ByVal nHeight As Long, _
                             ByVal hwndParent As Long, ByVal hMenu As Long, _
                             ByVal hInstance As Long, lpParam As Any) As Long

Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Item()
Public ToolTip()

Public m_cTT As New cTooltip

Public Function Max(param1 As Long, param2 As Long) As Long
    If param1 > param2 Then Max = param1 Else Max = param2
End Function

Public Function GetStrFromBufferA(szA As String) As String
    If InStr(szA, vbNullChar) Then
        GetStrFromBufferA = Left$(szA, InStr(szA, vbNullChar) - 1)
    Else
        GetStrFromBufferA = szA
    End If
End Function

Function addListBoxToolTip()
    Dim A As Long
    Dim i As Long
    
    On Error Resume Next
    
    If frmMainCode.List1.ListCount > 0 Then
        ReDim Item(frmMainCode.List1.ListCount)
        ReDim ToolTip(frmMainCode.List1.ListCount)
        
        For A = 1 To UBound(Item)
            Item(A) = frmMainCode.List1.List(A - 1)
            ToolTip(A) = frmMainCode.List1.List(A - 1)
        Next
        frmMainCode.List1.Clear
        For i = 1 To (A - 1)
            frmMainCode.List1.AddItem (Item(i))
        Next
    End If
End Function
