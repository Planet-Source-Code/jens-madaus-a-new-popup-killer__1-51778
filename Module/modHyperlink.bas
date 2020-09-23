Attribute VB_Name = "modHyperlink"
Option Explicit

Private Declare Function FindExecutable Lib "shell32.dll" _
    Alias "FindExecutableA" (ByVal lpFile As String, _
    ByVal lpDirectory As String, ByVal lpResult As String) As Long
    
Private Declare Function SetCapture Lib "user32.dll" _
    (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" _
    () As Long

Private Declare Function GetDesktopWindow Lib "user32" _
    () As Long

Private Declare Function GetSystemDirectory Lib "kernel32" _
    Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_NOTFOUND = 2

Global Const Link_Normal = &H864820
Global Const Link_Hover = &H864820
Global Const resHand = 1000
Global Const resMAil = 1100

Public Sub LinkCreate(Link As Label, _
    ByVal Text As String, ByVal Aktion As String, _
    Optional Container As Variant)

    With Link
        .Caption = Text
        .Tag = Aktion
        .BorderStyle = 0
        
        If Not IsMissing(Container) Then
            .BackColor = &HC6C7C6
        End If
    End With
End Sub

Public Sub LinkDisplay(Link As Label, _
    Optional ByVal ColorNormal As Variant)
  
    With Link
        If IsMissing(ColorNormal) Then
            ' wenn keine "gesonderte" Textfarbe angegeben wurde,
            ' Standard-Textfarbe verwenden
            .ForeColor = Link_Normal
        Else
            ' Übergebene Textfarbe verwenden
            .ForeColor = ColorNormal
        End If
        .Font.Underline = True
        ' Fehlerbehandlung einschalten, falls das Icon
        ' in der Resource-Datei nicht gefunden wird
        On Local Error Resume Next
        If InStr(1, Link, "mail") > 0 Then
            .MouseIcon = LoadResPicture(resMAil, 1)
        Else
            .MouseIcon = LoadResPicture(resHand, 1)
        End If
        .MousePointer = 99
        On Local Error GoTo 0
    End With
End Sub

Public Sub LinkHover(Link As Label, x As Single, _
    y As Single, Optional ByVal ColorNormal As Variant, _
    Optional ByVal ColorHover As Variant)
  
    With Link
        If x >= 0 And y >= 0 And x <= .Width And y <= .Height Then
            If IsMissing(ColorHover) Then
                .ForeColor = Link_Hover
            Else
                .ForeColor = ColorHover
            End If
            .Font.Underline = True
        Else
            If IsMissing(ColorNormal) Then
                .ForeColor = Link_Normal
            Else
                .ForeColor = ColorNormal
            End If
            .Font.Underline = True
        End If
    End With
End Sub

Public Sub LinkGo(Link As Label, _
    Optional ByVal ColorNormal As Variant)
  
    Dim URL As String
    Dim SB As String
    
    LinkHover Link, -1, -1, ColorNormal
    
    Select Case True
        Case Left$(Link.Tag, 7) = "http://", Left$(Link.Tag, 4) = "www."
            URL = Link.Tag
            If Left$(URL, 7) <> "http://" Then URL = "http://" & URL
            SB = StandardBrowser("")
            If SB <> "" Then
                ShellExecute GetDesktopWindow(), "Open", SB, URL, App.Path, vbNormalFocus
            Else
                If lngLangIndex = 0 Then
                    ShowDialog frmMainCode, "Das Programm kann Ihren Internet-Browser nicht ausmachen!" & vbCrLf & "Bitte öffnen Sie die nachfolgende URL in Ihrem Browser:" & vbCrLf & "http://www.pc-tool.de", "Standardbrowser nicht verfügbar...", OK, INFO
                Else
                    ShowDialog frmMainCode, "The program cannot constitute your Internet Browser!" & vbCrLf & "Please open the following URL in your Browser:" & vbCrLf & "http://www.pc-tool.de", "Standardbrowser not available...", OK, INFO
                End If
            End If
        Case Left$(Link.Tag, 7) = "mailto:"
            Call StartEMail(GetDesktopWindow(), "toolwork@web.de", _
            IIf(lngLangIndex = 0, "Information/Anfrage", "Information/Inquiry"))
        Case Left$(Link.Tag, 4) = "App:"
            DocumentOpen Mid$(Link.Tag, 5)
    End Select
    SB = ""
    URL = ""
End Sub

Private Sub DocumentOpen(sFilename As String)
    Dim sDirectory As String
    Dim lRet As Long
    Dim DeskWin As Long
    
    DeskWin = GetDesktopWindow()
    lRet = ShellExecute(DeskWin, "open", sFilename, _
    vbNullString, vbNullString, vbNormalFocus)
    
    Select Case True
        Case lRet = SE_ERR_NOTFOUND
        Case lRet = SE_ERR_NOASSOC
            sDirectory = Space(260)
            lRet = GetSystemDirectory(sDirectory, Len(sDirectory))
            sDirectory = Left$(sDirectory, lRet)
            Call ShellExecute(DeskWin, vbNullString, _
            "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & _
            sFilename, sDirectory, vbNormalFocus)
    End Select
End Sub

Public Function StandardBrowser(Browser As String) As String
    Dim sExe As String
    Dim tmpFile As String
    Dim dNr As Integer
    
    tmpFile = App.Path + IIf(Right$(App.Path, 1) <> "\", "\", "") + "xxx.html"
    dNr = FreeFile
    
    Open tmpFile For Output As #dNr
    Close #dNr
    
    sExe = ExePfad(tmpFile)
    Kill tmpFile
    
    If sExe <> "" Then
        Select Case True
            Case InStr(LCase$(sExe), "iexplore") > 0
                Browser = "Microsoft Internet Explorer"
            Case InStr(LCase$(sExe), "netscape") > 0
                Browser = "Netscape Communicator"
            Case InStr(LCase$(sExe), "opera") > 0
                Browser = "Opera"
            Case Else
                Browser = ""
        End Select
    End If
    
    StandardBrowser = sExe
End Function

Public Function ExePfad(ByVal Datei As String) As String
    Dim Pfad As String
    
    Pfad = Space$(256)
    FindExecutable Datei, vbNullString, Pfad
    
    If Pfad <> "" Then
        Pfad = Left$(Pfad, InStr(Pfad, vbNullChar) - 1)
    End If
    
    If UCase$(Pfad) = UCase$(Datei) Then Pfad = ""
    ExePfad = Pfad
End Function



