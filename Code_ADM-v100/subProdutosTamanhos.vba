Option Compare Database



Private Sub cboTamanho_Click()
    
    Me.codTamanho = Me.cboTamanho.Column(1)
    Me.Tamanhos = Me.cboTamanho.Column(2)
    Me.codProduto = intProduto


End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click


    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
DoCmd.Maximize
End Sub


