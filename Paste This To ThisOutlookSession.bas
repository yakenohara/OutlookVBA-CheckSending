Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

    If MsgBox("送信してもいいですか？", vbYesNo + vbDefaultButton2) = vbNo Then
    
        Cancel = True
    
    End If

End Sub

