Option Compare Database

Private Sub cmdEnviar_Click()
    Dim ctlOrigem As ListBox
    Dim ctlDestino As ListBox
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim strSQL As String
    
    Set ctlOrigem = Me.lstCoresDisponiveis
    Set ctlDestino = Me.lstCoresDoProduto
    
    For intCurrentRow = 0 To ctlOrigem.ListCount - 1
        If ctlOrigem.Selected(intCurrentRow) Then
            strSQL = "INSERT INTO ProdutosCores ( codProduto, codCor, Cor, Cores ) SELECT " & Me.codProduto & " as Produto, lstCoresDiponiveis.codCategoria, lstCoresDiponiveis.Categoria, lstCoresDiponiveis.Descricao01 FROM lstCoresDiponiveis where codCategoria = " & ctlOrigem.Column(0, intCurrentRow) & ""
            ExecutarSQL strSQL
            Me.lstCoresDoProduto.Requery
            Me.lstCoresDisponiveis.Requery
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing
End Sub


Private Sub cmdRemover_Click()
    Dim ctlOrigem As ListBox
    Dim ctlDestino As ListBox
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim strSQL As String
    
    Set ctlOrigem = Me.lstCoresDoProduto
    Set ctlDestino = Me.lstCoresDisponiveis
    
    For intCurrentRow = 0 To ctlOrigem.ListCount - 1
        If ctlOrigem.Selected(intCurrentRow) Then
            strSQL = "Delete * from ProdutosCores where codProduto = " & Me.codProduto & " and codCor = " & ctlOrigem.Column(0, intCurrentRow) & ""
            ExecutarSQL strSQL
            Me.lstCoresDoProduto.Requery
            Me.lstCoresDisponiveis.Requery
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing
End Sub

Private Sub lstCoresDisponiveis_DblClick(Cancel As Integer)
   Call cmdEnviar_Click
End Sub

Private Sub lstCoresDoProduto_DblClick(Cancel As Integer)
    Call cmdRemover_Click
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
