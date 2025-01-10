Attribute VB_Name = "MasUsadas"

'======================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
'======================================================
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"        ' COMO SI PULSARA ENTER
    End If
End Sub

'=====================================================================
'Private Sub txt_GotFocus(Index As Integer)
'=====================================================================
   
'txt(Index).SelStart = 0
'txt(Index).SelLength = Len(txt(Index).Text)
'End Sub
