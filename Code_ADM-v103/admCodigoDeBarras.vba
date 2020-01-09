Option Compare Database

Private Sub cmdPlanilha_Click()
    ExportarCodigoDeBarras_PontoDeVenda
End Sub

Private Sub Form_Close()
    DoCmd.Restore
End Sub

Private Sub Form_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub
