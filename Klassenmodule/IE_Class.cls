VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IE_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetAsyncKeyState% Lib "user32" (ByVal vKey As Long)
Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Private Declare Function ScreenToClient& Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI)
Private Declare Function GetKeyState% Lib "user32" (ByVal nVirtKey As Long)

Private WithEvents IE        As SHDocVw.InternetExplorer
Attribute IE.VB_VarHelpID = -1
Private WithEvents IEDoc     As MSHTML.HTMLDocument
Attribute IEDoc.VB_VarHelpID = -1

Private mvarIEHandle        As Long
Private mvarEnableBack      As Boolean
Private mvarEnableForward   As Boolean
Private bLDown              As Boolean
Private bRDown              As Boolean
Private timerS              As Single

Public Property Get EnableForward() As Boolean
    EnableForward = mvarEnableForward
End Property

Public Property Get EnableBack() As Boolean
    EnableBack = mvarEnableBack
End Property

Public Property Set iectl(IncomeIE As SHDocVw.InternetExplorer)
   
    On Error Resume Next
    
    Set IE = Nothing
    Set IE = IncomeIE
    
    mvarIEHandle = IncomeIE.hwnd
    
    Set IEDoc = Nothing
    Set IEDoc = IncomeIE.Document
    
    bCancel = True
    Err.Clear
    IE.GoForward
    
    If Err Then
        mvarEnableForward = False
        bCancel = False
    Else
        IE.ExecWB OLECMDID_STOPDOWNLOAD, OLECMDEXECOPT_DONTPROMPTUSER
    End If
    
End Property

Public Property Get iectl() As SHDocVw.InternetExplorer
    Set iectl = IE
End Property

Public Property Get IEHandle() As Long
    IEHandle = mvarIEHandle
End Property

Private Sub IE_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
    On Error Resume Next
    
    Select Case Command
        Case CSC_NAVIGATEBACK
            'LB = False
            mvarEnableBack = Enable
        Case CSC_NAVIGATEFORWARD
            'LB = False
            mvarEnableForward = Enable
    End Select
End Sub

Public Sub Class_Terminate()
    On Error Resume Next
    Set IE = Nothing
    Set IEDoc = Nothing
    Set iectl = Nothing
End Sub

Private Sub IE_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    On Error Resume Next
    
    Select Case True
        Case InStr(1, UCase(URL), "POPUP")
            LB = False
        Case LB And tempGum <> 1
            If Left$(UCase(URL), 10) <> "JAVASCRIPT" Then
                LB = False
            Else
                tempGum = 1
            End If
        Case Left$(UCase(URL), 10) = "JAVASCRIPT" And GetAsyncKeyState(VK_LBUTTON)
            LB = True
    End Select

    If bCancel Then
       Cancel = True
       bCancel = False
    End If
End Sub

Private Sub IE_NewWindow2(ppDisp As Object, Cancel As Boolean)
    On Error Resume Next
    
    Dim i       As Long
    Dim TmpStr  As String
    Dim URL     As String
    Dim taggi   As String
    
    Zähler = 0
    
    tempGum = 0
    IE.Silent = True
    
    If (GetAsyncKeyState(17) < 0) Or (Not (Active)) Then Exit Sub
    TmpStr = IE.LocationURL

    taggi = UCase(IEDoc.activeElement.tagName)
    
    If chkURLString <> "" Then
        If Right(TmpStr, 1) = "/" Then TmpStr = Left$(TmpStr, (Len(TmpStr) - 1))
        If InStr(1, TmpStr, "//") > 0 Then TmpStr = Mid$(TmpStr, InStr(1, TmpStr, "//") + 2)
        If InStr(1, UCase(chkURLString), UCase(TmpStr)) > 0 Then Exit Sub
    End If
    
    Select Case True
        Case taggi = "FRAME", taggi = "IFRAME", taggi = "A", taggi <> "", LB, UCase$(Left$(taggi, 3) = "ATL")
            Select Case True
                Case taggi = "FRAMESET", LB = False, GetAsyncKeyState(VK_LBUTTON), UCase$(taggi) = "BODY"
                    GoTo weiter
            End Select
            LB = False
            Cancel = False
            Exit Sub
        Case Else
weiter:
            LB = False
            XY = XY + 1
            Cancel = True
            If bSound Then PlaySound
            Call SetStringWert(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Disable Script Debugger", "yes")
            frmMainCode.lblCount.Caption = CStr(XY)
    End Select
End Sub
