Attribute VB_Name = "modEvents"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long

Public cIEWPtr As Long
Public bCancel As Boolean

Public Enum IDEVENTS
    ID_BeforeNavigate = 1
    ID_NavigationComplete = 2
    ID_DownloadBegin = 3
    ID_DownloadComplete = 4
    ID_DocumentComplete = 5
    ID_MouseDown = 6
    ID_MouseUp = 7
    ID_ContextMenu = 8
    ID_CommandStateChange = 9
End Enum

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Byte, ByVal uFlags As Long) As Long

Public Enum SoundID
  vbMusica = 101
  vbUtopia = 102
End Enum

Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Public SoundArray() As Byte

Private Function ResolvePointer(ByVal lpObj&) As cIEWindows
    Dim oIEW As cIEWindows
    CopyMemory oIEW, lpObj, 4&
    Set ResolvePointer = oIEW
    CopyMemory oIEW, 0&, 4&
End Function

Function LoadSound()
    On Error Resume Next
    SoundArray = LoadResData(101, "CUSTOM")
End Function

Function PlaySound()
    On Error Resume Next
    sndPlaySound SoundArray(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Function

Function SaveSettings()
    On Error Resume Next
    
    Dim i As Long
    Dim t As String
    
    With frmMainCode.List1
        For i = 0 To .ListCount - 1
            t = t & .List(i) & ";"
        Next
    End With
    
    If Len(t) > 0 Then
        t = Left$(t, Len(t) - 1)
        t = Compress(t)
    End If
    
    SaveSetting App.EXEName, "Config", "Sure", t
    SaveSetting App.EXEName, "Config", "Sound", (frmMainCode.chkSound.Value)
    SaveSetting App.EXEName, "Config", "NoDial", (frmMainCode.chkDial.Value)
    SaveSetting App.EXEName, "Config", "First", 0
End Function

Function DelEntry()
    Dim i As Long
    Dim s As String
    Dim k As String
    
    If lngLangIndex = 0 Then
        s = "Möchten Sie diesen" & vbCrLf & "Eintrag wirklich löschen?"
        k = "Löschen eines sicheren Eintrags..."
    Else
        s = "Do you really want to" & vbCrLf & "delete the selected item?"
        k = "Delete entry from 'popup white list'..."
    End If
        
    MakeNormal (frmMainCode.hwnd)
    i = ShowDialog(frmMainCode, s, k, JA_NEIN, QUESTION)
    MakeTopMost (frmMainCode.hwnd)
    
    If i = 2 Then
        i = 0
        While frmMainCode.List1.Selected(i) = False
            i = i + 1
        Wend
        frmMainCode.List1.RemoveItem (i)
    End If
End Function

Function createURLString()
    On Error Resume Next
    
    Dim i As Long
    
    chkURLString = ""
    
    With frmMainCode.List1
        While i <= .ListCount - 1
            chkURLString = chkURLString & .List(i)
            i = i + 1
        Wend
    End With
End Function

Sub Main()
    Dim Startart As Integer
    Startart = Val(Trim$(Command$))
    Load frmMainCode
End Sub

