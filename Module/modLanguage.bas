Attribute VB_Name = "modLanguage"
Option Explicit

Function setLanguage(lLang As Long) As Boolean
    With frmMainCode
        Select Case lLang
            Case 0 'Deutsch
                frmMenuForm.mAktURL.Caption = "URL aus Liste hinzufügen"
                frmMenuForm.mDelEntry.Caption = "Eintrag löschen"
                .chkSound.Caption = Chr(38) & Chr(83) & Chr(111) & Chr(117) & Chr(110) & Chr(100) & Chr(32) & Chr(98) & Chr(101) & Chr(105) & Chr(32) & Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(107) & Chr(105) & Chr(108) & Chr(108)
                .chkDial.Caption = "Dialerdownload blocken"
                .chkStartOnWin.Caption = Chr(83) & Chr(116) & Chr(97) & Chr(114) & Chr(116) & Chr(101) & Chr(110) & Chr(32) & Chr(109) & Chr(105) & Chr(116) & Chr(32) & Chr(38) & Chr(87) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115)
                .lblLanguage.Caption = Chr(83) & Chr(112) & Chr(114) & Chr(97) & Chr(99) & Chr(104) & Chr(101) & Chr(58)
                .lblSidePopup.Caption = Chr(83) & Chr(101) & Chr(105) & Chr(116) & Chr(101) & Chr(110) & Chr(32) & Chr(109) & Chr(105) & Chr(116) & Chr(32) & Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(115) & Chr(116) & Chr(97) & Chr(114) & Chr(116) & Chr(58)
                .Command3.Caption = Chr(38) & Chr(72) & Chr(105) & Chr(110) & Chr(122) & Chr(117) & Chr(102) & Chr(252) & Chr(103) & Chr(101) & Chr(110)
                .Command4.Caption = Chr(38) & Chr(76) & Chr(246) & Chr(115) & Chr(99) & Chr(104) & Chr(101) & Chr(110)
                .Command5.Caption = Chr(97) & Chr(107) & Chr(116) & Chr(117) & Chr(101) & Chr(108) & Chr(108) & Chr(101) & Chr(32) & Chr(38) & Chr(85) & Chr(82) & Chr(76) & Chr(32) & Chr(104) & Chr(105) & Chr(110) & Chr(122) & Chr(117) & Chr(102) & Chr(252) & Chr(103) & Chr(101) & Chr(110)
                .lblInfo.Caption = IIf(Active, Chr(83) & Chr(116) & Chr(97) & Chr(116) & Chr(117) & Chr(115) & Chr(58) & Chr(32) & Chr(97) & Chr(107) & Chr(116) & Chr(105) & Chr(118), Chr(83) & Chr(116) & Chr(97) & Chr(116) & Chr(117) & Chr(115) & Chr(58) & Chr(32) & Chr(105) & Chr(110) & Chr(97) & Chr(107) & Chr(116) & Chr(105) & Chr(118))
                .btnReg.Caption = Chr(38) & Chr(82) & Chr(101) & Chr(103) & Chr(105) & Chr(115) & Chr(116) & Chr(114) & Chr(105) & Chr(101) & Chr(114) & Chr(101) & Chr(110)
                
            Case 1 'Englisch
                frmMenuForm.mAktURL.Caption = "add URL from list"
                frmMenuForm.mDelEntry.Caption = "delete entry"
                .chkSound.Caption = Chr(38) & Chr(83) & Chr(111) & Chr(117) & Chr(110) & Chr(100) & Chr(32) & Chr(111) & Chr(110) & Chr(32) & Chr(112) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(107) & Chr(105) & Chr(108) & Chr(108)
                .chkDial.Caption = "Stop Dialerdownload"
                .chkStartOnWin.Caption = Chr(83) & Chr(116) & Chr(97) & Chr(114) & Chr(116) & Chr(117) & Chr(112) & Chr(32) & Chr(119) & Chr(105) & Chr(116) & Chr(104) & Chr(32) & Chr(38) & Chr(87) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115)
                .lblLanguage.Caption = Chr(76) & Chr(97) & Chr(110) & Chr(103) & Chr(117) & Chr(97) & Chr(103) & Chr(101) & Chr(58)
                .lblSidePopup.Caption = Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(32) & Chr(119) & Chr(104) & Chr(105) & Chr(116) & Chr(101) & Chr(32) & Chr(108) & Chr(105) & Chr(115) & Chr(116) & Chr(58)
                .Command3.Caption = Chr(38) & Chr(65) & Chr(100) & Chr(100)
                .Command4.Caption = Chr(38) & Chr(68) & Chr(101) & Chr(108) & Chr(101) & Chr(116) & Chr(101)
                .Command5.Caption = Chr(65) & Chr(100) & Chr(100) & Chr(32) & Chr(97) & Chr(99) & Chr(116) & Chr(117) & Chr(97) & Chr(108) & Chr(32) & Chr(38) & Chr(85) & Chr(82) & Chr(76)
                .lblInfo.Caption = IIf(Active, Chr(83) & Chr(116) & Chr(97) & Chr(116) & Chr(117) & Chr(115) & Chr(58) & Chr(32) & Chr(97) & Chr(99) & Chr(116) & Chr(105) & Chr(118) & Chr(101), Chr(83) & Chr(116) & Chr(97) & Chr(116) & Chr(117) & Chr(115) & Chr(58) & Chr(32) & Chr(105) & Chr(110) & Chr(97) & Chr(99) & Chr(116) & Chr(105) & Chr(118) & Chr(101))
                .btnReg.Caption = Chr(38) & Chr(82) & Chr(101) & Chr(103) & Chr(105) & Chr(115) & Chr(116) & Chr(101) & Chr(114)
        End Select
    End With
End Function

Function CreateToolTips(lLang As Long)
    Select Case lLang
        Case 0
            With m_cTT
                frmMenuForm.mAddAktURL.Caption = "aktuelle URL hinzufügen"
                frmMenuForm.mnuShowMe.Caption = "Zeige Real Popup-Killer"
                .ToolText(frmMainCode.Command2) = Chr(72) & Chr(105) & Chr(101) & Chr(114) & Chr(32) & Chr(101) & Chr(114) & Chr(104) & Chr(97) & Chr(108) & Chr(116) & Chr(101) & Chr(110) & Chr(32) & Chr(83) & Chr(105) & Chr(101) & Chr(32) & Chr(107) & Chr(108) & Chr(101) & Chr(105) & Chr(110) & Chr(101) _
                                                  & vbCrLf & Chr(97) & Chr(98) & Chr(101) & Chr(114) & Chr(32) & Chr(119) & Chr(105) & Chr(99) & Chr(104) & Chr(116) & Chr(105) & Chr(103) & Chr(101) & Chr(32) & Chr(72) & Chr(105) & Chr(110) & Chr(119) & Chr(101) & Chr(105) & Chr(115) & Chr(101) & Chr(32) & Chr(97) & Chr(117) & Chr(102) _
                                                  & vbCrLf & Chr(100) & Chr(105) & Chr(101) & Chr(32) & Chr(80) & Chr(114) & Chr(111) & Chr(103) & Chr(114) & Chr(97) & Chr(109) & Chr(109) & Chr(102) & Chr(117) & Chr(110) & Chr(107) & Chr(116) & Chr(105) & Chr(111) & Chr(110) & Chr(101) & Chr(110)
                If frmMainCode.Command1.Caption = Chr(38) & Chr(62) & Chr(62) Then
                    .ToolText(frmMainCode.Command1) = Chr(85) & Chr(110) & Chr(100) & Chr(32) & Chr(110) & Chr(111) & Chr(99) & Chr(104) & Chr(32) & Chr(109) & Chr(101) & Chr(104) & Chr(114) _
                                                      & vbCrLf & Chr(69) & Chr(105) & Chr(110) & Chr(115) & Chr(116) & Chr(101) & Chr(108) & Chr(108) & Chr(117) & Chr(110) & Chr(103) & Chr(115) & Chr(45) _
                                                      & vbCrLf & Chr(109) & Chr(246) & Chr(103) & Chr(108) & Chr(105) & Chr(99) & Chr(104) & Chr(107) & Chr(101) & Chr(105) & Chr(116) & Chr(101) & Chr(110)
                Else
                    .ToolText(frmMainCode.Command1) = Chr(80) & Chr(114) & Chr(111) & Chr(103) & Chr(114) & Chr(97) & Chr(109) & Chr(109) & Chr(45) & Chr(79) & Chr(112) & Chr(116) & Chr(105) & Chr(111) & Chr(110) & Chr(101) & Chr(110) _
                                                      & vbCrLf & Chr(97) & Chr(117) & Chr(115) & Chr(98) & Chr(108) & Chr(101) & Chr(110) & Chr(100) & Chr(101) & Chr(110)
                End If
                .ToolText(frmMainCode.chkDial) = "Verhindert das" & vbCrLf & "Ausführen von Dialern"
                .ToolText(frmMainCode.chkSound) = Chr(83) & Chr(111) & Chr(117) & Chr(110) & Chr(100) & Chr(32) & Chr(97) & Chr(110) & Chr(47) & Chr(97) & Chr(117) & Chr(115) _
                                                  & vbCrLf & Chr(98) & Chr(101) & Chr(105) & Chr(32) & Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(107) & Chr(105) & Chr(108) & Chr(108)
                .ToolText(frmMainCode.Command3) = Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(115) & Chr(101) & Chr(105) & Chr(116) & Chr(101) _
                                                  & vbCrLf & Chr(104) & Chr(105) & Chr(110) & Chr(122) & Chr(117) & Chr(102) & Chr(252) & Chr(103) & Chr(101) & Chr(110)
                .ToolText(frmMainCode.Command4) = Chr(69) & Chr(105) & Chr(110) & Chr(101) & Chr(47) & Chr(97) & Chr(108) & Chr(108) & Chr(101) & Chr(32) & Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(115) & Chr(101) & Chr(105) & Chr(116) & Chr(101) & Chr(110) _
                                                  & vbCrLf & Chr(97) & Chr(117) & Chr(115) & Chr(32) & Chr(100) & Chr(101) & Chr(114) & Chr(32) & Chr(76) & Chr(105) & Chr(115) & Chr(116) & Chr(101) & Chr(32) & Chr(101) & Chr(110) & Chr(116) & Chr(102) & Chr(101) & Chr(114) & Chr(110) & Chr(101) & Chr(110)
                .ToolText(frmMainCode.btnClose) = Chr(66) & Chr(101) & Chr(101) & Chr(110) & Chr(100) & Chr(101) & Chr(110)
                .ToolText(frmMainCode.btnMinimize) = Chr(77) & Chr(105) & Chr(110) & Chr(105) & Chr(109) & Chr(105) & Chr(101) & Chr(114) & Chr(101) & Chr(110)
                .ToolText(frmMainCode.btnMinimizeTray) = Chr(77) & Chr(105) & Chr(110) & Chr(105) & Chr(109) & Chr(105) & Chr(101) & Chr(114) & Chr(101) & Chr(110) _
                                                         & vbCrLf & Chr(122) & Chr(117) & Chr(109) & Chr(32) & Chr(84) & Chr(114) & Chr(97) & Chr(121)
                .ToolText(frmMainCode.chkStartOnWin) = Chr(82) & Chr(101) & Chr(97) & Chr(108) & Chr(32) & Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(45) & Chr(75) & Chr(105) & Chr(108) & Chr(108) & Chr(101) & Chr(114) & Chr(32) & Chr(109) & Chr(105) & Chr(116) _
                                                       & vbCrLf & Chr(87) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115) & Chr(32) & Chr(115) & Chr(116) & Chr(97) & Chr(114) & Chr(116) & Chr(101) & Chr(110)
                .ToolText(frmMainCode.Command5) = Chr(85) & Chr(82) & Chr(76) & Chr(32) & Chr(101) & Chr(105) & Chr(110) & Chr(101) & Chr(115) & Chr(32) & Chr(111) & Chr(102) & Chr(102) & Chr(101) & Chr(110) & Chr(101) & Chr(110) & Chr(32) & Chr(73) & Chr(69) & Chr(32) & Chr(45) & Chr(70) & Chr(101) & Chr(110) & Chr(115) & Chr(116) & Chr(101) & Chr(114) & Chr(115) _
                                                  & vbCrLf & Chr(122) & Chr(117) & Chr(114) & Chr(32) & Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(102) & Chr(114) & Chr(101) & Chr(101) & Chr(45) & Chr(76) & Chr(105) & Chr(115) & Chr(116) & Chr(101) & Chr(32) & Chr(104) & Chr(105) & Chr(110) & Chr(122) & Chr(117) & Chr(102) & Chr(252) & Chr(103) & Chr(101) & Chr(110)
                .ToolText(frmMainCode.cmbLanguage) = Chr(87) & Chr(228) & Chr(104) & Chr(108) & Chr(101) & Chr(110) & Chr(32) & Chr(83) & Chr(105) & Chr(101) & Chr(32) & Chr(104) & Chr(105) & Chr(101) & Chr(114) & Chr(32) & Chr(73) & Chr(104) & Chr(114) & Chr(101) _
                                                     & vbCrLf & Chr(98) & Chr(101) & Chr(118) & Chr(111) & Chr(114) & Chr(122) & Chr(117) & Chr(103) & Chr(116) & Chr(101) & Chr(32) & Chr(83) & Chr(112) & Chr(114) & Chr(97) & Chr(99) & Chr(104) & Chr(101) & Chr(32) & Chr(97) & Chr(117) & Chr(115)
            End With
                frmMainCode.lblKontakt.Caption = "Kontakt:"
                frmMainCode.lblInfo.ToolTipText = Chr(75) & Chr(108) & Chr(105) & Chr(99) & Chr(107) & Chr(101) & Chr(110) & Chr(32) & Chr(61) & Chr(32) & Chr(97) & Chr(107) & Chr(116) & Chr(105) & Chr(118) & Chr(105) & Chr(101) & Chr(114) & Chr(101) & Chr(110) & Chr(47) & Chr(100) & Chr(101) & Chr(97) & Chr(107) & Chr(116) & Chr(105) & Chr(118) & Chr(105) & Chr(101) & Chr(114) & Chr(101) & Chr(110)
                frmMenuForm.mAktiv.Caption = Chr(97) & Chr(107) & Chr(116) & Chr(105) & Chr(118)
                frmMenuForm.mHideIE.Caption = "IE-Fenster verstecken (Winkey+Z)"
                frmMenuForm.mExit.Caption = Chr(66) & Chr(101) & Chr(101) & Chr(110) & Chr(100) & Chr(101) & Chr(110)
        Case 1
            With m_cTT
                frmMenuForm.mAddAktURL.Caption = "add actual URL to 'p-w-l'"
                frmMenuForm.mnuShowMe.Caption = "Show Real Popup-Killer"
                
                .ToolText(frmMainCode.Command2) = Chr(71) & Chr(101) & Chr(116) & Chr(32) & Chr(97) & Chr(32) & Chr(115) & Chr(109) & Chr(97) & Chr(108) & Chr(108) & Chr(44) _
                                                 & vbCrLf & Chr(104) & Chr(111) & Chr(119) & Chr(101) & Chr(118) & Chr(101) & Chr(114) & Chr(44) & Chr(32) & Chr(105) & Chr(109) & Chr(112) & Chr(111) & Chr(114) & Chr(116) & Chr(97) & Chr(110) & Chr(116) & Chr(32) & Chr(105) & Chr(110) & Chr(100) & Chr(105) & Chr(99) & Chr(97) & Chr(116) & Chr(105) & Chr(111) & Chr(110) _
                                                 & vbCrLf & Chr(111) & Chr(102) & Chr(32) & Chr(116) & Chr(104) & Chr(101) & Chr(32) & Chr(112) & Chr(114) & Chr(111) & Chr(103) & Chr(114) & Chr(97) & Chr(109) & Chr(32) & Chr(102) & Chr(117) & Chr(110) & Chr(99) & Chr(116) & Chr(105) & Chr(111) & Chr(110)
                If frmMainCode.Command1.Caption = Chr(38) & Chr(62) & Chr(62) Then
                    .ToolText(frmMainCode.Command1) = Chr(65) & Chr(110) & Chr(100) & Chr(32) & Chr(115) & Chr(116) & Chr(105) & Chr(108) & Chr(108) & Chr(32) & Chr(109) & Chr(111) & Chr(114) & Chr(101) _
                                                      & vbCrLf & Chr(112) & Chr(114) & Chr(111) & Chr(103) & Chr(114) & Chr(97) & Chr(109) & Chr(32) & Chr(111) & Chr(112) & Chr(116) & Chr(105) & Chr(111) & Chr(110) & Chr(115)
                Else
                    .ToolText(frmMainCode.Command1) = Chr(72) & Chr(105) & Chr(100) & Chr(101) & Chr(32) & Chr(112) & Chr(114) & Chr(111) & Chr(103) & Chr(114) & Chr(97) & Chr(109) & Chr(45) _
                                                      & vbCrLf & Chr(111) & Chr(112) & Chr(116) & Chr(105) & Chr(111) & Chr(110) & Chr(115)
                End If
                .ToolText(frmMainCode.chkDial) = "Stops the" & vbCrLf & "dialer-downloads"
                .ToolText(frmMainCode.chkSound) = Chr(83) & Chr(111) & Chr(117) & Chr(110) & Chr(100) & Chr(32) & Chr(111) & Chr(110) & Chr(47) & Chr(111) & Chr(102) & Chr(102) _
                                                  & vbCrLf & Chr(111) & Chr(110) & Chr(32) & Chr(112) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(107) & Chr(105) & Chr(108) & Chr(108)
                .ToolText(frmMainCode.Command3) = Chr(65) & Chr(100) & Chr(100) & Chr(32) & Chr(85) & Chr(82) & Chr(76) & Chr(32) & Chr(116) & Chr(111) & Chr(32) & Chr(116) & Chr(104) & Chr(101) _
                                                  & vbCrLf & Chr(39) & Chr(112) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(32) & Chr(119) & Chr(104) & Chr(105) & Chr(116) & Chr(101) & Chr(32) & Chr(108) & Chr(105) & Chr(115) & Chr(116) & Chr(39)
                .ToolText(frmMainCode.Command4) = Chr(82) & Chr(101) & Chr(109) & Chr(111) & Chr(118) & Chr(101) & Chr(32) & Chr(111) & Chr(110) & Chr(101) & Chr(47) & Chr(97) & Chr(108) & Chr(108) & Chr(32) & Chr(85) & Chr(82) & Chr(76) & Chr(115) _
                                                  & vbCrLf & Chr(102) & Chr(114) & Chr(111) & Chr(109) & Chr(32) & Chr(39) & Chr(112) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(32) & Chr(119) & Chr(104) & Chr(105) & Chr(116) & Chr(101) & Chr(32) & Chr(108) & Chr(105) & Chr(115) & Chr(116) & Chr(39)
                .ToolText(frmMainCode.btnClose) = Chr(69) & Chr(120) & Chr(105) & Chr(116)
                .ToolText(frmMainCode.btnMinimize) = Chr(77) & Chr(105) & Chr(110) & Chr(105) & Chr(109) & Chr(105) & Chr(122) & Chr(101)
                .ToolText(frmMainCode.btnMinimizeTray) = Chr(77) & Chr(105) & Chr(110) & Chr(105) & Chr(109) & Chr(105) & Chr(122) & Chr(101) _
                                                         & vbCrLf & Chr(116) & Chr(111) & Chr(32) & Chr(84) & Chr(114) & Chr(97) & Chr(121)
                .ToolText(frmMainCode.chkStartOnWin) = Chr(83) & Chr(116) & Chr(97) & Chr(114) & Chr(116) & Chr(32) & Chr(82) & Chr(101) & Chr(97) & Chr(108) & Chr(32) & Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(45) & Chr(75) & Chr(105) & Chr(108) & Chr(108) & Chr(101) & Chr(114) _
                                                       & vbCrLf & Chr(111) & Chr(110) & Chr(32) & Chr(87) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115) & Chr(45) & Chr(83) & Chr(116) & Chr(97) & Chr(114) & Chr(116) & Chr(117) & Chr(112)
                .ToolText(frmMainCode.Command5) = Chr(65) & Chr(100) & Chr(100) & Chr(32) & Chr(97) & Chr(110) & Chr(32) & Chr(97) & Chr(99) & Chr(116) & Chr(117) & Chr(97) & Chr(108) & Chr(108) & Chr(121) & Chr(32) & Chr(111) & Chr(112) & Chr(101) & Chr(110) & Chr(101) & Chr(100) & Chr(32) & Chr(73) & Chr(69) & Chr(45) & Chr(85) & Chr(82) & Chr(76) _
                                                  & vbCrLf & Chr(116) & Chr(111) & Chr(32) & Chr(116) & Chr(104) & Chr(101) & Chr(32) & Chr(39) & Chr(112) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(32) & Chr(119) & Chr(104) & Chr(105) & Chr(116) & Chr(101) & Chr(32) & Chr(108) & Chr(105) & Chr(115) & Chr(116) & Chr(39)
                .ToolText(frmMainCode.cmbLanguage) = Chr(83) & Chr(101) & Chr(108) & Chr(101) & Chr(99) & Chr(116) & Chr(32) & Chr(121) & Chr(111) & Chr(117) & Chr(114) _
                                                     & vbCrLf & Chr(102) & Chr(97) & Chr(118) & Chr(111) & Chr(117) & Chr(114) & Chr(101) & Chr(100) & Chr(32) & Chr(108) & Chr(97) & Chr(110) & Chr(103) & Chr(117) & Chr(97) & Chr(103) & Chr(101)
            End With
                frmMainCode.lblInfo.ToolTipText = Chr(67) & Chr(108) & Chr(105) & Chr(99) & Chr(107) & Chr(32) & Chr(61) & Chr(32) & Chr(97) & Chr(99) & Chr(116) & Chr(105) & Chr(118) & Chr(97) & Chr(116) & Chr(101) & Chr(47) & Chr(100) & Chr(101) & Chr(97) & Chr(99) & Chr(116) & Chr(105) & Chr(118) & Chr(97) & Chr(116) & Chr(101)
                frmMenuForm.mAktiv.Caption = Chr(97) & Chr(99) & Chr(116) & Chr(105) & Chr(118) & Chr(101)
                frmMenuForm.mHideIE.Caption = "Hide all IE-Windows (Winkey+Z)"
                frmMenuForm.mExit.Caption = Chr(69) & Chr(120) & Chr(105) & Chr(116)
                frmMainCode.lblKontakt.Caption = "Contact:"
    End Select
    If Active Then
        frmMainCode.TIcon.ChangeToolTip frmMainCode.picTray, ("Real Popup-Killer " & IIf(lngLangIndex = 1, _
                                                              "ist aktiv", "is active"))
    Else
        frmMainCode.TIcon.ChangeToolTip frmMainCode.picTray, (Chr(82) & Chr(101) & Chr(97) & Chr(108) & Chr(32) & Chr(80) & Chr(111) & Chr(112) & Chr(117) & Chr(112) & Chr(45) & Chr(75) & Chr(105) & Chr(108) & Chr(108) & Chr(101) & Chr(114) & Chr(32) & IIf(lngLangIndex = 1, _
                                                              Chr(105) & Chr(115) & Chr(116) & Chr(32) & Chr(105) & Chr(110) & Chr(97) & Chr(107) & Chr(116) & Chr(105) & Chr(118), Chr(105) & Chr(115) & Chr(32) & Chr(105) & Chr(110) & Chr(97) & Chr(99) & Chr(116) & Chr(105) & Chr(118) & Chr(101)))
    End If
End Function
