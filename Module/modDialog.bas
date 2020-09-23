Attribute VB_Name = "modDialog"
Option Explicit

Public Enum dlgIcon
    noICON = 0
    INFO = 1
    QUESTION = 2
    EXCLAMATION = 3
    CRITICAL = 3
End Enum

Public Enum dlgButton
    OK = 1
    JA = 2
    NEIN = 3
    JA_NEIN = 4
    ABBRECHEN = 5
    OK_ABBRECHEN = 6
End Enum

Private mButtonTyp As dlgButton
Private mIcon As dlgIcon

Public cmdAnswer As Variant

Public Function ShowDialog(frm As Form, msg As String, Title As String, Optional mButtonTyp As dlgButton, _
                           Optional mIcon As dlgIcon = 0, Optional BtnCloseEnabled As Boolean = False, _
                           Optional screenCenter As Boolean, Optional showList As Boolean, _
                           Optional showInput As Boolean) As Variant
    Dim Diff        As Long
    Dim imgWidth    As Long
    Dim im          As Control
    
    On Error Resume Next
    
    With frmDialog
        .Height = 100
        .Width = 100
        .imginformation.Visible = False
        .imgquestion.Visible = False
        .imgexclamation.Visible = False
        .imgcritical.Visible = False
        .lstUrls.Visible = False
        .txtDialogInput.Visible = False
        
        If showInput Then
            .lblTitle = Title
            .cmdOK.Default = False
            .cmdOK.Visible = True
            .cmdAbbruch.Visible = True
            .lblMessage.Width = .txtDialogInput.Width - .cmdOK.Width - 200
            .lblMessage = msg
            .txtDialogInput.Visible = True
            .txtDialogInput.Left = .lblMessage.Left
            .txtDialogInput.Top = .lblMessage.Top + .lblMessage.Height + 250
            .Width = .txtDialogInput.Width + .txtDialogInput.Left * 2
            .cmdOK.Top = .lblMessage.Top
            .cmdOK.Left = .txtDialogInput.Left + .txtDialogInput.Width - .cmdOK.Width
            .cmdAbbruch.Top = .cmdOK.Top + .cmdOK.Height + 30
            .cmdAbbruch.Left = .cmdOK.Left
            .Height = .txtDialogInput.Top + .txtDialogInput.Height + 200
            Call MakeWindow(frmDialog, False)
            CenterDialog
            .Show vbModal, frm
            ShowDialog = cmdAnswer
            Set im = Nothing
            Exit Function
        End If
    End With
    
    If mIcon > 0 Then
        Select Case mIcon
            Case 1
                Set im = frmDialog.imginformation
            Case 2
                Set im = frmDialog.imgquestion
            Case 3
                Set im = frmDialog.imgexclamation
            Case 4
                Set im = frmDialog.imgcritical
            Case Else
                imgWidth = 0
        End Select
        imgWidth = im.Width
        im.Left = 250
    End If
    
    With frmDialog
        .lblMessage = msg
        .lblTitle = Title
        .lblMessage.Left = .lblMessage.Left + (imgWidth) + im.Left
        .Width = .lblMessage.Width + (.lblMessage.Left * 2)
        .Left = frmMainCode.Left + frmMainCode.Width / 2 - .Width / 2
        .Top = frmMainCode.Top + frmMainCode.Height / 2 - .Height / 2
          
        If Not im Is Nothing Then
            im.Top = .lblMessage.Top + .lblMessage.Height / 2 - im.Height / 2
            im.Visible = True
            If (.Width - .lblMessage.Left / 2) >= (.lblTitle.Width + 800) Then
                .Width = .Width - .lblMessage.Left / 2
            End If
        End If
        
        If .Width < (.lblTitle.Width + 800) Then
            .Width = (.lblTitle.Width + 800)
        End If
            
        Select Case mButtonTyp
            Case 1
                .cmdOK.Visible = True
                If .Width < .cmdOK.Width * 1.5 Then .Width = .cmdOK.Width * 1.5
                .cmdOK.Top = .lblMessage.Top + .lblMessage.Height + 200
                .cmdOK.Left = .Width / 2 - .cmdOK.Width / 2
                .Height = .cmdOK.Top + .cmdOK.Height + 250
                Call MakeWindow(frmDialog, False)
                CenterDialog
                .Show vbModal, frm
                ShowDialog = cmdAnswer

            Case 2
                .cmdJa.Visible = True
                .Height = .Height + .cmdJa.Height
                .cmdJa.Top = .lblMessage.Top + .lblMessage.Height + .cmdJa.Height / 2
                .cmdJa.Left = .Width / 2 - .cmdJa.Width / 2
                
            Case 3
                .cmdNein.Visible = True
                .Height = .Height + .cmdOK.Height
                .cmdOK.Top = .lblMessage.Top + .lblMessage.Height + .cmdOK.Height / 2
                .cmdOK.Left = .Width / 2 - .cmdOK.Width / 2
                
            Case 4
                .cmdJa.Visible = True
                .cmdNein.Visible = True
                
                If .Width < (.cmdJa.Width + 100 + .cmdNein.Width) * 1.5 Then
                    .Width = (.cmdJa.Width + 100 + .cmdNein.Width) * 1.5
                Else
                    Diff = .Width - (.cmdJa.Width + 100 + .cmdNein.Width) * 1.5
                End If
                
                .cmdJa.Top = .lblMessage.Top + .lblMessage.Height + .cmdJa.Height / 2
                .cmdJa.Left = .Width / 2 - .cmdJa.Width - 30
                .cmdNein.Top = .cmdJa.Top
                .cmdNein.Left = .cmdJa.Left + .cmdJa.Width + 60
                .Height = .cmdJa.Top + .cmdJa.Height + 250
                
                Call MakeWindow(frmDialog, False)
                CenterDialog
                .Show vbModal, frm
                ShowDialog = cmdAnswer
            Case 5
                .cmdOK.Visible = True
                .Height = .Height + .cmdOK.Height
                .cmdOK.Top = .lblMessage.Top + .lblMessage.Height + .cmdOK.Height / 2
                .cmdOK.Left = .Width / 2 - .cmdOK.Width / 2
            
            Case 6
                .cmdAbbruch.Visible = True
                .cmdOK.Visible = True
                .lstUrls.Visible = True
                .lstUrls.Top = .lblMessage.Top + .lblMessage.Height + 100
                .lstUrls.Left = .lblMessage.Left
                .lstUrls.Width = .lblMessage.Width
                
                If .Width < (.cmdOK.Width + 100 + .cmdAbbruch.Width) * 1.5 Then
                    .Width = (.cmdOK.Width + 100 + .cmdAbbruch.Width) * 1.5
                Else
                    Diff = .Width - (.cmdOK.Width + 100 + .cmdAbbruch.Width) * 1.5
                End If
                
                .cmdOK.Top = .lstUrls.Top + .lstUrls.Height + 200
                .cmdOK.Left = .Width / 2 - .cmdOK.Width - 100 '+ Diff / 2
                .cmdAbbruch.Top = .cmdOK.Top
                .cmdAbbruch.Left = .cmdOK.Left + .cmdOK.Width + 100
                .Height = .cmdOK.Top + .cmdOK.Height + 250
                
                Call MakeWindow(frmDialog, False)
                CenterDialog
                .Show vbModal, frm
                ShowDialog = cmdAnswer
            
            Case Else
                Call MakeWindow(frmDialog, False)
                .Show , frm
        End Select

    End With
    Set im = Nothing
End Function

Function CenterDialog()
    frmDialog.Left = frmMainCode.Left + frmMainCode.Width / 2 - frmDialog.Width / 2
    frmDialog.Top = (frmMainCode.Top + frmMainCode.Height / 2) - frmDialog.Height / 2
End Function
