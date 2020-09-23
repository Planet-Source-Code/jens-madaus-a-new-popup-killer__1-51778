Attribute VB_Name = "modApi"
Option Explicit

Public SWs As New SHDocVw.ShellWindows
Public var As SHDocVw.InternetExplorer
Public Const WM_CLOSE = &H10
Public Const PROCESS_TERMINATE = &H1

Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function DLLSelfRegister Lib "VB6STKIT.DLL" (ByVal lpDllName As String) As Integer
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage& Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal dwMessage As Long) As Integer

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
  
Private Declare Function GetShortPathName Lib "kernel32" _
        Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
        ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
  
'Left-click constants.
Global Const WM_LBUTTONDBLCLK = &H203   'Double-click
Global Const WM_LBUTTONDOWN = &H201     'Button down
Global Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Global Const WM_RBUTTONDBLCLK = &H206   'Double-click
Global Const WM_RBUTTONDOWN = &H204     'Button down
Global Const WM_RBUTTONUP = &H205       'Button up

Private Type POINTAPI
    x As Long
    y As Long
End Type

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'Abfrage des Systems ************************************
Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" _
    (LpVersionInformation As OSVERSIONINFO) As Long
                                  
Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformID        As Long
    szCSDVersion        As String * 128
End Type

Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

'*********************************************************

Public Const VK_LBUTTON = &H1
Public bIsRegistry  As Boolean
Public LB           As Boolean
Public Active       As Boolean
Public XY           As Long
Public bSound       As Boolean
Public sTimer       As Single
Public chkURLString As String
Public DLK As Object

Public LastState As String
Public DownState As String
Public MyArray As String
Public frmWB As frmMenuForm
Public TmpTimer As Byte
Public Antidown As Byte

Public Xp  As Single
Public yp As Single

Public ieURL As String
Public ap As String

Public Zähler As Long
Public lngLangIndex As Long

Public tempGum As Byte
Public chkClick As Boolean

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE 'TOPMOST_FLAGS
End Sub

Function GetInfo() As String
    On Error Resume Next
    
    Dim CursorPos As POINTAPI
    Dim HwndNow As Long
    Dim szText As String * 100
    
    GetCursorPos CursorPos
    HwndNow = WindowFromPoint(CursorPos.x, CursorPos.y)
    GetClassName HwndNow, szText, 100
    GetInfo = szText
End Function

Public Sub StartEMail(ByVal hwnd As Long, Optional ByVal Empfänger As String = "", _
                      Optional ByVal Betreff As String = "")
  On Error Resume Next
  
  Screen.MousePointer = 11
  Call ShellExecute(hwnd, "Open", "mailto:" + _
                    Empfänger + IIf(Betreff <> "", _
                    "?subject=" + Betreff, ""), "", "", 1)
  Screen.MousePointer = vbDefault
End Sub

Public Function IsWinNT() As Boolean
    Dim osvi As OSVERSIONINFO
    
    osvi.dwOSVersionInfoSize = Len(osvi)
    GetVersionEx osvi
    IsWinNT = (osvi.dwPlatformID = VER_PLATFORM_WIN32_NT)
End Function

Function setIEURLs() As Long
    Dim oDoc As Object
    Dim i As Long
    Dim z As Long
    Dim tmpS As String
    Dim tmpL As String
    Dim b As Boolean
    
    On Error Resume Next
    
    Set SWs = New ShellWindows
    
    For i = 0 To SWs.Count - 1
        Set var = SWs.Item(i)
        Set oDoc = var.Document
        
        If TypeName(oDoc) = "HTMLDocument" Then
            If InStr(1, var.LocationURL, "//") Then
                tmpS = Mid$(var.LocationURL, InStr(1, var.LocationURL, "//") + 2)
                If Right(tmpS, 1) = "/" Then tmpS = Left$(tmpS, Len(tmpS) - 1)
                z = z + 1
            Else
                tmpS = var.LocationURL
            End If
            Call frmDialog.lstUrls.AddItem(tmpS)
        End If
    Next
    setIEURLs = z
    Set SWs = Nothing
    Set var = Nothing
End Function

Function addAktURLs(handle As Long)
    Dim oDoc    As Object
    Dim i       As Long
    Dim z       As Long
    Dim tmpS    As String
    Dim tmpL    As String
    Dim b       As Boolean
    
    On Error Resume Next
    
    Set SWs = New ShellWindows
    
    For i = 0 To SWs.Count - 1
        Set var = SWs.Item(i)
        Set oDoc = var.Document
        
        If var.hwnd = handle Then
            If TypeName(oDoc) = "HTMLDocument" Then
                If InStr(1, var.LocationURL, "//") Then
                    tmpS = Mid$(var.LocationURL, InStr(1, var.LocationURL, "//") + 2)
                    If Right(tmpS, 1) = "/" Then tmpS = Left$(tmpS, Len(tmpS) - 1)
                    z = z + 1
                Else
                    tmpS = var.LocationURL
                End If
                For z = 0 To frmMainCode.List1.ListCount - 1
                    If frmMainCode.List1.List(z) = tmpS Then
                        b = True
                        Exit For
                    End If
                Next
                If Not b Then Call frmMainCode.List1.AddItem(tmpS)
                Exit For
            End If
        End If
    Next
    Set var = Nothing
    Set SWs = Nothing
End Function

Function addURL()
    Dim s As String
    Dim k As String
    Dim t As String
    Dim b As Boolean
    Dim z As Long
    
    With frmMainCode
        If setIEURLs > 0 Then
            MakeNormal (.hwnd)
            If lngLangIndex = 1 Then
                s = "Please select one of the entries below" & vbCrLf & "and click on Ok! This one will be included" & vbCrLf & "into the 'popup white list'!"
                k = "Add URL to the 'popup white list'"
            Else
                s = "Bitte wählen Sie eine der unten aufgelisteten" & vbCrLf & "URLs per Klick aus! Diese wird" & vbCrLf & "dann in die 'Popupfree-Liste' aufgenommen!"
                k = "URL zur Popupfree-Liste hinzufügen"
            End If
            
            t = ShowDialog(frmMainCode, s, k, OK_ABBRECHEN, , , , True)

            If Len(t) > 2 Then
                For z = 0 To frmMainCode.List1.ListCount - 1
                    If frmMainCode.List1.List(z) = Mid$(t, InStr(1, t, ";") + 1) Then
                        b = True
                        Exit For
                    End If
                Next
                If Not b Then .List1.AddItem (Mid$(t, InStr(1, t, ";") + 1))
                Call addListBoxToolTip
                createURLString
            End If
        Else
            MakeNormal (.hwnd)
            If lngLangIndex = 1 Then
                s = "Can't find any open" & vbCrLf & "Internet-Explorer window!"
                k = "Only with open IE-windows..."
            Else
                s = "Es sind keine Internet-" & vbCrLf & "Explorer-Fenster geöffnet!"
                k = "Nur bei geöffnetem IE..."
            End If
            
            Call ShowDialog(frmMainCode, s, k, OK, INFO)
        End If
    End With
    MakeTopMost (frmMainCode.hwnd)
End Function

Function SetAppVar()
    On Error Resume Next
    
    ap = App.Path
    ap = GetShortPath(ap)
    If Right(ap, 1) <> "\" Then ap = ap & "\"
    ap = ap & App.EXEName & ".exe"
End Function

Public Function fIsFileDIR(stPath As String, Optional lngType As Long) As Integer
    On Error Resume Next
    
    fIsFileDIR = Len(Dir(stPath, lngType)) > 0
End Function

Function GetShortPath(LongPath As String) As String
    Dim lngRes  As Long
    Dim strPath As String
    
    strPath = String$(165, 0)
    lngRes = GetShortPathName(LongPath, strPath, 164)

    If lngRes = 0 Then
        GetShortPath = LongPath
    Else
        GetShortPath = Left$(strPath, lngRes)
    End If
End Function

Function GetSysPath() As String
    Dim sysDir As String, Buffer As String * 260
    sysDir = Left$(Buffer, GetSystemDirectory(Buffer, Len(Buffer)))
    If Right(sysDir, 1) <> "\" Then sysDir = sysDir & "\"
    GetSysPath = sysDir
End Function

Public Function CreateFileFromResource(ByVal FilePath As String, ByVal ResID As Long, ByVal ResType As String) As Boolean
    Dim hFile As Long
    Dim FileData() As Byte
    
    FileData = LoadResData(ResID, ResType)
    On Error Resume Next
                         
    If Len(FileData(1)) > 0 Then
        hFile = FreeFile()
        Open FilePath For Binary Access Write As #hFile
        Put #hFile, 1, FileData
        Close #hFile
        CreateFileFromResource = True
    End If
End Function

Function ActivateDLKiller(chkd As Long)
    On Error GoTo Fehler
    Dim TmpStr As String
    If chkd = 1 Then
        Call AddDialKiller
        Set DLK = CreateObject("KillDial.KillClass")
    Else
        On Error Resume Next
        Set DLK = Nothing
        TmpStr = GetSysPath & "RPKAddon.dll"
        If fIsFileDIR(TmpStr) = -1 Then Kill (TmpStr)
    End If
    Exit Function
Fehler:
    frmMainCode.chkDial.Enabled = False
    Err.Clear
End Function

Function AddDialKiller()
    On Error GoTo Fehler
    Dim TmpStr As String
    TmpStr = GetSysPath & "RPKAddon.dll"
    If fIsFileDIR(TmpStr) <> -1 Then
        Call CreateFileFromResource(TmpStr, 102, "CUSTOM")
        DLLSelfRegister "RPKAddon.dll"
    End If
    TmpStr = ""
    Exit Function
Fehler:
    frmMainCode.chkDial.Value = 0
    frmMainCode.chkDial.Enabled = False
    Err.Clear
End Function
