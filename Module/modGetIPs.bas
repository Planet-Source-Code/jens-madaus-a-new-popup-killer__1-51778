Attribute VB_Name = "modGetIPs"
Option Explicit
      
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal _
        wVersionRequired&, lpWSAData As WinSocketDataType) _
        As Long
        
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal _
        HostName$, ByVal HostLen%) As Long
        
Private Declare Function gethostbyname Lib "WSOCK32.DLL" _
        (ByVal HostName$) As Long
        
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" _
        (ByVal addr$, ByVal laenge%, ByVal typ%) As Long
        
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As _
        Any, ByVal hpvSource&, ByVal cbCopy&)
       
Private Type HostDeType
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Const WS_VERSION_REQD = &H101
Const MIN_SOCKETS_REQD = 1
Const SOCKET_ERROR = -1
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128

Private Type WinSocketDataType
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Public Function GetIPs() As String
    Dim IP As String, Host As String
    Dim x As Integer
    Dim i As Long
    Dim IPS As String
    
    Call InitSocketAPI
    Host = MyHostName
    
    frmMenuForm.mnuIPs.Visible = False
    frmMenuForm.mNix2.Visible = False
    
    For i = 1 To 4
        frmMenuForm.mnuOwn(i).Visible = False
    Next
    
    Do
        IP = HostByName(Host, x)
        If Len(IP) <> 0 Then
            frmMenuForm.mnuIPs.Visible = True
            frmMenuForm.mNix2.Visible = True
            frmMenuForm.mnuOwn(x).Caption = IP
            frmMenuForm.mnuOwn(x).Visible = True
        End If
        x = x + 1
    Loop While Len(IP) > 0
    Call CleanSockets
    GetIPs = IPS
End Function

Private Sub InitSocketAPI()
    Dim Result%
    Dim LoBy%, HiBy%
    Dim SocketData As WinSocketDataType
    
    Result = WSAStartup(WS_VERSION_REQD, SocketData)
    If Result <> 0 Then
        'MsgBox ("'winsock.dll' antwortet nicht !")
    End If
End Sub

Private Function MyHostName() As String
    Dim HostName As String * 256
    
    If gethostname(HostName, 256) = SOCKET_ERROR Then
        'MsgBox "Windows Sockets error " & str(WSAGetLastError())
        Exit Function
    Else
        MyHostName = NextChar(Trim$(HostName), Chr$(0))
    End If
End Function

Private Function HostByName(name$, Optional x% = 0) As String
    Dim MemIp() As Byte
    Dim y%
    Dim HostDeAddress&, HostIp&
    Dim IpAddress$
    Dim Host As HostDeType
    
    HostDeAddress = gethostbyname(name)
    If HostDeAddress = 0 Then
        HostByName = ""
        Exit Function
    End If
    
    Call RtlMoveMemory(Host, HostDeAddress, LenB(Host))
    
    For y = 0 To x
        Call RtlMoveMemory(HostIp, Host.hAddrList + 4 * y, 4)
        If HostIp = 0 Then
            HostByName = ""
            Exit Function
        End If
    Next
    
    ReDim MemIp(1 To Host.hLength)
    Call RtlMoveMemory(MemIp(1), HostIp, Host.hLength)
    
    IpAddress = ""
    
    For y = 1 To Host.hLength
        IpAddress = IpAddress & MemIp(y) & "."
    Next
    
    IpAddress = Left$(IpAddress, Len(IpAddress) - 1)
    HostByName = IpAddress
End Function

Private Sub CleanSockets()
    Dim Result&
    
    Result = WSACleanup()
    If Result <> 0 Then
        'MsgBox ("Socket Error " & Trim$(str$(Result)) & _
        " in Prozedur 'CleanSockets' aufgetreten !")
    End If
End Sub

Private Function NextChar(Text$, Char$) As String
    Dim pos%
    pos = InStr(1, Text, Char)
    If pos = 0 Then
        NextChar = Text
        Text = ""
    Else
        NextChar = Left$(Text, pos - 1)
        Text = Mid$(Text, pos + Len(Char))
    End If
End Function
