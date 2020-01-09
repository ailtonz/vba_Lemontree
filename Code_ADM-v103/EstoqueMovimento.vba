Option Compare Database

Private Sub cbocodBarras_Click()
    Me.txtArtigo = Me.cbocodBarras.Column(1)
    Me.txtComposicao = Me.cbocodBarras.Column(2)
    Me.txtModelo = Me.cbocodBarras.Column(3)
    Me.txtCor = Me.cbocodBarras.Column(4)
    Me.txtTamanho = Me.cbocodBarras.Column(5)
End Sub

Private Sub Form_Close()
    DoCmd.Restore
End Sub

Private Sub Form_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub


