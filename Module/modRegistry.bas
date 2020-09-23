Attribute VB_Name = "modRegistry"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" (ByVal HKEY As Long, ByVal _
        lpSubKey As String, ByVal ulOptions As Long, ByVal _
        samDesired As Long, phkResult As Long) As Long
        
Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal HKEY As Long) As Long
        
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal _
        lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Any) As Long
        
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" (ByVal HKEY As Long, ByVal _
        lpSubKey As String, ByVal Reserved As Long, ByVal _
        lpClass As String, ByVal dwOptions As Long, ByVal _
        samDesired As Long, ByVal lpSecurityAttributes As Any, _
        phkResult As Long, lpdwDisposition As Long) As Long
        
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal _
        HKEY As Long) As Long
        
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, lpData As Long, ByVal cbData As Long) _
        As Long
        
Private Declare Function RegSetValueEx_Str Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, ByVal lpData As String, ByVal cbData As _
        Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
        "RegDeleteKeyA" (ByVal HKEY As Long, ByVal lpSubKey As _
        String) As Long
        
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
        "RegDeleteValueA" (ByVal HKEY As Long, ByVal lpValueName _
        As String) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias _
        "RegEnumValueA" (ByVal HKEY As Long, ByVal dwIndex As Long, _
        ByVal lpValueName As String, lpcbValueName As Long, _
        ByVal lpReserved As Long, lpType As Long, lpData As Byte, _
        lpcbData As Long) As Long 'Gibt ein existierendes Feld aus

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE Or _
                 KEY_ENUMERATE_SUB_KEYS _
                 Or KEY_NOTIFY
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE Or _
                       KEY_SET_VALUE Or _
                       KEY_CREATE_SUB_KEY Or _
                       KEY_ENUMERATE_SUB_KEYS Or _
                       KEY_NOTIFY Or _
                       KEY_CREATE_LINK
Const ERROR_SUCCESS = 0&

Const REG_OPTION_NON_VOLATILE = &H0

Public Enum eArt
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
End Enum

Public Type eEnum
    name As String
    art As eArt
End Type

Private m_art As Long
Private m_hkey As Long
Private m_sFeld As String

Function RegFieldDelete(Root&, Key$, Field$) As Long
    Dim Result&, HKEY&
    
    Result = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, HKEY)
    If Result = ERROR_SUCCESS Then
        Result = RegDeleteValue(HKEY, Field)
        Result = RegCloseKey(HKEY)
    End If
    RegFieldDelete = Result
End Function

Function RegValueSet(Root&, Key$, Field$, Value As Variant) As Long
    Dim Result&, HKEY&, s$, l&

    Result = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, HKEY)
    If Result = ERROR_SUCCESS Then
        Select Case VarType(Value)
            Case vbInteger, vbLong
            l = CLng(Value)
            Result = RegSetValueEx(HKEY, Field, 0, REG_DWORD, l, 4)
            Case vbString
            s = CStr(Value)
            Result = RegSetValueEx_Str(HKEY, Field, 0, REG_SZ, s, _
                                    Len(s) + 1)
        End Select
        Result = RegCloseKey(HKEY)
    End If
    RegValueSet = Result
End Function

Function RegKeyCreate(Root&, Newkey$) As Long
    Dim Result&, HKEY&, Back&

    Result = RegCreateKeyEx(Root, Newkey, 0, vbNullString, REG_OPTION_NON_VOLATILE, _
                            KEY_ALL_ACCESS, 0&, HKEY, Back)
    If Result = ERROR_SUCCESS Then
        Result = RegFlushKey(HKEY)
        If Result = ERROR_SUCCESS Then Call RegCloseKey(HKEY)
        RegKeyCreate = Back
    End If
End Function

Public Function Alle_Felder_auflisten(Optional Anzahl As Integer = 255) As eEnum()
    Dim cnt     As Long
    Dim Result  As Long
    Dim Length  As Long
    Dim i       As Integer
    Dim art     As eArt
    Dim Feld()  As eEnum
    Dim Feld2() As eEnum
    Dim name    As String * 255
    
    ReDim Feld(Anzahl)
    
    Feld(0).name = ""
    Feld(0).art = REG_SZ
    Do
        Length = Len(name)
        Result = RegEnumValue(m_hkey, cnt, ByVal name, Length, ByVal 0&, art, ByVal 0, ByVal 0&)
        If Result = ERROR_SUCCESS And Length <> 0 Then
            Feld(cnt + 1).name = Left$(name, Length)
            Feld(cnt + 1).art = art
        End If
        cnt = cnt + 1
    Loop Until Result <> ERROR_SUCCESS
    
    ReDim Preserve Feld(cnt - 2)
    Alle_Felder_auflisten = Feld
End Function

Function RegKeyExist(Root&, Key$) As Long
    Dim Result&, HKEY&
    Result = RegOpenKeyEx(Root, Key, 0, KEY_READ, HKEY)
    If Result = ERROR_SUCCESS Then Call RegCloseKey(HKEY)
    RegKeyExist = Result
End Function

Function SetLongWert(RegRoot As Long, sSchluessel As String, sFeld As String, LngWert As String)
    Dim Result As Long
    Dim LngInt As Long
    
    LngInt = CLng(Val(LngWert))
    Result = RegValueSet(RegRoot, sSchluessel, sFeld, LngInt)
    
    If Result = 0 Then
      'Label7.Caption = "Ok"
    Else
      'Label7.Caption = "Fehler"
    End If
End Function

Function SetStringWert(RegRoot As Long, sSchluessel As String, sFeld As String, Wert As String) As Long
    Dim Result As Long
    Dim strgVal As String
    
    strgVal = Trim(Wert)
    Result = RegValueSet(RegRoot, sSchluessel, sFeld, Wert)
    
    If Result = 0 Then bIsRegistry = True
End Function

Function RegValueGet(Root&, Key$, Field$, Value As Variant) As Long
    Dim Result&, HKEY&, dwType&, Lng&, Buffer$, l&

    Result = RegOpenKeyEx(Root, Key, 0, KEY_READ, HKEY)
    If Result = ERROR_SUCCESS Then
        Result = RegQueryValueEx(HKEY, Field, 0&, dwType, ByVal 0&, l)
        If Result = ERROR_SUCCESS Then
            Select Case dwType
                Case REG_SZ
                    Buffer = Space$(l + 1)
                    Result = RegQueryValueEx(HKEY, Field, 0&, _
                                     dwType, ByVal Buffer, l)
                    If Result = ERROR_SUCCESS Then
                        Value = GetStrFromBufferA(Buffer)
                    End If
                Case REG_DWORD
                    Result = RegQueryValueEx(HKEY, Field, 0&, dwType, Lng, l)
                    If Result = ERROR_SUCCESS Then Value = Lng
            End Select
        End If
    End If
    
    If Result = ERROR_SUCCESS Then Result = RegCloseKey(HKEY)
    RegValueGet = Result
End Function

