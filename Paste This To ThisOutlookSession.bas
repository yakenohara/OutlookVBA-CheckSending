Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

    If MsgBox("���M���Ă������ł����H", vbYesNo + vbDefaultButton2) = vbNo Then
    
        Cancel = True
    
    End If

End Sub

