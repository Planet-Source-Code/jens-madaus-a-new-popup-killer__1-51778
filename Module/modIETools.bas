Attribute VB_Name = "modIETools"
Option Explicit

Type DllVersionInfo
cbSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type

Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DllVersionInfo) As Long

Dim IEMV As DllVersionInfo
Dim CheckReg As String
Dim GetIEMajor As String
Dim Hico As String
Dim Ico As String
Dim Prog As String

Public Function DetectTitle() As String
    On Error Resume Next
    
    Dim Result As Long, Value As String
    
    Result = RegValueGet(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Window Title", Value)
    If Result = 0 And Value <> "" Then
        DetectTitle = Value
    Else
        DetectTitle = "Microsoft Internet Explorer"
    End If
End Function

Public Function mnuAddIE()
    Hico = App.Path & "\Icons\" & "HT.ico"
    Ico = App.Path & "\Icons\" & "IC.ico"
    Prog = App.Path & "\" & App.EXEName
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "ButtonText", "Real Popup-Killer"
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}"
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "Default Visible", "Yes"
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "Exec", Prog
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "HotIcon", Hico
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "Icon", Ico
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "MenuStatusBar", "Real Popup-Killer"
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "MenuText", "&Real Popup-Killer"
End Function

Public Function mnuDeleteIE()
    REGDeleteSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}"
End Function

