VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIEWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"IE_Class"
Attribute VB_Ext_KEY = "Member0" ,"IE_Class"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents winShell As SHDocVw.ShellWindows
Attribute winShell.VB_VarHelpID = -1
Private mCol As Collection
Public Event IECommandStateChange(hwnd As Long, Button As CommandStateChangeConstants, Enable As Boolean)

Dim bRefreshing As Boolean
    
Public Property Get IE(vntIndexKey As Variant) As IE_Class
Attribute IE.VB_UserMemId = 0
    Do While bRefreshing
        DoEvents
        DoEvents
    Loop
    Set IE = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Private Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Refresh_Col
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    Set winShell = Nothing
End Sub

Private Sub winShell_WindowRegistered(ByVal lCookie As Long)
    Refresh_Col
    Do While bRefreshing
        DoEvents
    Loop
End Sub

Private Sub winShell_WindowRevoked(ByVal lCookie As Long)
    Refresh_Col
    Do While bRefreshing
        DoEvents
    Loop
End Sub

Private Sub Refresh_Col()
    On Error Resume Next
    
    bRefreshing = True
    
    Dim i As Long
    
    Set mCol = Nothing
    Set mCol = New Collection
    
    MyArray = ""
    
    For i = 0 To SWs.Count - 1
        Addit SWs(i), i
    Next
    
    If MyArray <> "" Then
        Call DLK.killdialer(MyArray, DetectTitle)
    End If
        
    i = (mCol.Count > 0)
    With frmMenuForm
        .mAddAktURL.Enabled = i
        .mAktURL.Enabled = i
        .mHideIE.Enabled = i
    End With
    frmMainCode.Command5.Enabled = i
    
    Set winShell = Nothing
    Set winShell = SWs
    Set SWs = Nothing
    Set var = Nothing
    bRefreshing = False
End Sub

Private Function Addit(iectl, i As Long)
    On Error Resume Next
    
    Dim ONH As Long
    Dim objNewMember As IE_Class
    
    If TypeOf iectl.Document Is HTMLDocument Then
        Set objNewMember = New IE_Class
        Set objNewMember.iectl = iectl
        ONH = objNewMember.IEHandle
        mCol.Add objNewMember, CStr(ONH)
        If Len(MyArray) > 0 Then
            MyArray = MyArray & ";" & ONH
        Else
            MyArray = ONH
        End If
    End If
    Set objNewMember = Nothing
End Function
