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
            strSQL = "UPDATE Vendas SET Vendas.Selecao = Yes WHERE (((Vendas.Dia)=Format(#" & ctlOrigem.Column(0, intCurrentRow) & "#,'mm/dd/yyyy')))"
            ExecutarSQL strSQL
            Me.lstCoresDoProduto.Requery
            Me.lstCoresDisponiveis.Requery
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing
End Sub


Private Sub cmdEnviarTodos_Click()
    Dim ctlOrigem As ListBox
    Dim ctlDestino As ListBox
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim strSQL As String
    
    Set ctlOrigem = Me.lstCoresDisponiveis
    Set ctlDestino = Me.lstCoresDoProduto
          
    For intCurrentRow = 0 To ctlOrigem.ListCount - 1
        If Not IsNull(ctlOrigem.Column(0, intCurrentRow)) Then
            strSQL = "UPDATE Vendas SET Vendas.Selecao = Yes WHERE (((Vendas.Dia)=Format(#" & ctlOrigem.Column(0, intCurrentRow) & "#,'mm/dd/yyyy')))"
            ExecutarSQL strSQL
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Me.lstCoresDoProduto.Requery
    Me.lstCoresDisponiveis.Requery

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing

End Sub

Private Sub cmdExportarRecebimentos_Click()
    ExportarRecebimentos
End Sub

Private Sub cmdExportarVendas_Click()
    ExportarVendas
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
            strSQL = "UPDATE Vendas SET Vendas.Selecao = No WHERE (((Vendas.Dia)=Format(#" & ctlOrigem.Column(0, intCurrentRow) & "#,'mm/dd/yyyy')))"
            ExecutarSQL strSQL
            Me.lstCoresDoProduto.Requery
            Me.lstCoresDisponiveis.Requery
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing
End Sub

Private Sub cmdRemoverTodos_Click()
    Dim ctlOrigem As ListBox
    Dim ctlDestino As ListBox
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim strSQL As String
    
    Set ctlOrigem = Me.lstCoresDoProduto
    Set ctlDestino = Me.lstCoresDisponiveis
          
    For intCurrentRow = 0 To ctlOrigem.ListCount - 1
        If Not IsNull(ctlOrigem.Column(0, intCurrentRow)) Then
            strSQL = "UPDATE Vendas SET Vendas.Selecao = No WHERE (((Vendas.Dia)=Format(#" & ctlOrigem.Column(0, intCurrentRow) & "#,'mm/dd/yyyy')))"
            ExecutarSQL strSQL
            ctlOrigem.Selected(intCurrentRow) = False
        End If
    Next intCurrentRow

    Me.lstCoresDoProduto.Requery
    Me.lstCoresDisponiveis.Requery

    Set ctlOrigem = Nothing
    Set ctlDestino = Nothing


End Sub

Private Sub Form_Close()
    Call cmdRemoverTodos_Click
End Sub

Private Sub lstCoresDisponiveis_DblClick(Cancel As Integer)
   Call cmdEnviar_Click
End Sub

Private Sub lstCoresDoProduto_DblClick(Cancel As Integer)
    Call cmdRemover_Click
End Sub
