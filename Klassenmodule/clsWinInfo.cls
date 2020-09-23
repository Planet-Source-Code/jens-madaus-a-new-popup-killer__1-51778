VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Type DllVersionInfo
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type
    
'Private Konstanten
Private Const EM_UNDO As Long = &HC7
Private Const PROCESSOR_INTEL_386 As Integer = 386
Private Const PROCESSOR_INTEL_486 As Integer = 486
Private Const PROCESSOR_INTEL_PENTIUM As Integer = 586
Private Const PROCESSOR_MIPS_R4000 As Integer = 4000
Private Const PROCESSOR_ALPHA_21064 As Integer = 21064
Private Const ENTERPRISE_TYPE_LIVESTOCK As Long = 1
Private Const ENTERPRISE_TYPE_CROP As Long = 2

'Private Deklarationen
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (LpVersionInformation As OSVERSIONINFO) As Long

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As _
    MEMORYSTATUS)

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As _
    SYSTEM_INFO)

Private Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DllVersionInfo) As Long
    
Private mvarWinVersion As String
Private mvarWinName As String
Private mvarCPUVersion As String
Private mvarMemoryTotal As Long
Private mvarMemoryFree As Long
Private mvarVirtualMemoryTotal As Long
Private mvarVirtualMemoryFree As Variant
Private mvarIEVersion As String

Public Property Get VirtualMemoryFree() As Long
    VirtualMemoryFree = mvarVirtualMemoryFree
End Property

Public Property Get VirtualMemoryTotal() As Long
    VirtualMemoryTotal = mvarVirtualMemoryTotal
End Property

Public Property Get MemoryFree() As Long
    MemoryFree = mvarMemoryFree
End Property

Public Property Get MemoryTotal() As Long
    MemoryTotal = mvarMemoryTotal
End Property

Public Property Get WinVersion() As String
    WinVersion = mvarWinVersion
End Property

Public Property Get WinName() As String
    WinName = mvarWinName
End Property

Public Property Get CPUVersion() As String
    CPUVersion = mvarCPUVersion
End Property

Public Property Get IEVersion() As String
    IEVersion = mvarIEVersion
End Property

Private Function SystemInformation() As Boolean
    Dim strMsg As String
    Dim intRet As Integer
    Dim intVer_major As Integer
    Dim intVer_minor As Integer
    Dim lngBuild As Long
    Dim udtVerinfo As OSVERSIONINFO
    Dim udtSysInfo As SYSTEM_INFO
    Dim udtMemStatus As MEMORYSTATUS
    Dim lngMemory As Long
    Dim udtDVI As DllVersionInfo

    udtVerinfo.dwOSVersionInfoSize = Len(udtVerinfo)
    intRet = GetVersionEx(udtVerinfo)

    If intRet = 0 Then
        SystemInformation = False
    End If
    
    intVer_major = udtVerinfo.dwMajorVersion
    intVer_minor = udtVerinfo.dwMinorVersion

    Select Case intVer_major
        Case 3
            strMsg = IIf(intVer_minor = 5, "Windows NT ", "Windows ")
        Case 4
            Select Case intVer_minor
                Case 0
                strMsg = "Windows NT "
                Case 1
                strMsg = "Windows 98 "
                Case Else
                strMsg = "Windows ME "
            End Select
        Case 5
            strMsg = IIf(intVer_minor = 1, "Windows XP ", "Windows 2000 ")
        Case Else
            strMsg = "Windows Unknown Version"
    End Select
    
    lngBuild = udtVerinfo.dwBuildNumber

    mvarWinName = strMsg
    mvarWinVersion = intVer_major & "." & intVer_minor & " (Build " & lngBuild & ")"
    strMsg = ""
    
    GetSystemInfo udtSysInfo
    strMsg = strMsg + "CPU: "
    
    Select Case udtSysInfo.dwProcessorType
        Case PROCESSOR_INTEL_386
            strMsg = strMsg + "Intel 386"
        Case PROCESSOR_INTEL_486
            strMsg = strMsg + "Intel 486"
        Case PROCESSOR_INTEL_PENTIUM
            strMsg = strMsg + "Intel Pentium"
        Case PROCESSOR_MIPS_R4000
            strMsg = strMsg + "MIPS R4000"
        Case PROCESSOR_ALPHA_21064
            strMsg = strMsg + "DEC Alpha 21064"
        Case Else
            strMsg = strMsg + "(unknown)"
    End Select
    
    mvarCPUVersion = strMsg
    
    GlobalMemoryStatus udtMemStatus
    mvarMemoryTotal = udtMemStatus.dwTotalPhys / 1024
    mvarMemoryFree = udtMemStatus.dwAvailPhys / 1024
    mvarVirtualMemoryTotal = udtMemStatus.dwTotalVirtual / 1024
    mvarVirtualMemoryFree = udtMemStatus.dwAvailVirtual / 1024
    udtDVI.cbSize = Len(udtDVI)
    DllGetVersion udtDVI
    mvarIEVersion = "Internet Explorer " & udtDVI.dwMajorVersion & "." & _
    udtDVI.dwMinorVersion & "." & _
    udtDVI.dwBuildNumber
End Function

Private Sub Class_Initialize()
    SystemInformation
End Sub
