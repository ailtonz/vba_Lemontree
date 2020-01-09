Option Compare Database
Dim cont As Integer

Private Sub CadastrarItensDaVenda()
Dim strSQL As String
Dim strTotal As String
Dim rstValorUnitario As DAO.Recordset
Dim rstProduto As DAO.Recordset

If Me.txtCodigoDeBarras <> "" Then
    Set rstValorUnitario = CurrentDb.OpenRecordset("Select ValorUnitario from CodigoDeBarras where codBarras = '" & Me.txtCodigoDeBarras & "'")
    Set rstProduto = CurrentDb.OpenRecordset("Select ([Artigo] & ' - ' & [Composicao] & ' - ' & [Modelo] & ' - ' & [Cor] & ' - ' & [Tamanho]) as Descricao from CodigoDeBarras where codBarras = '" & Me.txtCodigoDeBarras & "'")
    
    If rstProduto.EOF Then
    
        Me.txtDescricao = ""
        Me.txtCodigoDeBarras = ""
        Me.txtValorUnitario = ""
        Me.txtDescontoItem = 0
        Me.txtCodigoDeBarras.SetFocus
    
    Else
    
        strTotal = rstValorUnitario.Fields("ValorUnitario") - Me.txtDescontoItem
        strSQL = "INSERT INTO VendasItens (codVenda, codBarras, Quantidade, DescontoItem, ValorUnitario,ValorTotal,Produto) " & _
                    " values ( '" & Me.txtcodVenda & "', '" & Me.txtCodigoDeBarras & "', 1, '" & Me.txtDescontoItem & "', '" & rstValorUnitario.Fields("ValorUnitario") & "',' " & strTotal & "',' " & rstProduto.Fields("Descricao") & "')"
        
        ExecutarSQL strSQL
        Me.subVendasItens.Requery
        Me.txtDescricao = ""
        Me.txtCodigoDeBarras = ""
        Me.txtValorUnitario = ""
        Me.txtDescontoItem = 0
        Me.txtCodigoDeBarras.SetFocus
        Me.subVendasItens.Requery
        rstValorUnitario.Close
        rstProduto.Close
        
    End If
End If

End Sub


Private Sub cmdFinalizarVenda_Click()
On Error GoTo Err_cmdFinalizarVenda_Click

    CancelarVenda
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    DoCmd.GoToRecord , , acNewRec
    Call Form_Load
    Me.subVendasItens.Requery
    Me.txtCodigoDeBarras.SetFocus

Exit_cmdFinalizarVenda_Click:
    Exit Sub

Err_cmdFinalizarVenda_Click:
    MsgBox Err.Description
    Resume Exit_cmdFinalizarVenda_Click
End Sub

Private Sub Form_Close()
    CancelarVenda
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyF12
            cmdRecebimentos_Click
        
        Case vbKeyEscape
            cmdFinalizarVenda_Click
            
    End Select
    
End Sub

Private Sub Form_Load()
    
'    Me.KeyPreview = True
    Me.txtcodVenda = NovoCodigo(Me.RecordSource, Me.txtcodVenda.ControlSource)
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Me.PontoDeVenda = Left(CurrentMDB(), Len(CurrentMDB()) - 4)

End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.Caption = Left(CurrentMDB(), Len(CurrentMDB()) - 4)
'    DoCmd.GoToRecord   , , acNewRec
    
    
End Sub

Private Sub txtCodigoDeBarras_Exit(Cancel As Integer)
Dim rstProduto As DAO.Recordset

If Me.txtCodigoDeBarras <> "" Then
    Me.txtDescricao = ""
    Me.txtValorUnitario = ""
    Set rstProduto = CurrentDb.OpenRecordset("Select ([Artigo] & ' - ' & [Composicao] & ' - ' & [Modelo] & ' - ' & [Cor] & ' - ' & [Tamanho]) as Descricao, ValorUnitario from CodigoDeBarras where codBarras = '" & Me.txtCodigoDeBarras & "'")
    If rstProduto.EOF Then
        Me.txtDescricao = "PRODUTO NÃO CADASTRADO!!!"
    Else
        Me.txtDescricao = rstProduto.Fields("Descricao")
        Me.txtValorUnitario = rstProduto.Fields("ValorUnitario")
    End If
    rstProduto.Close
    Set rstProduto = Nothing
End If

End Sub

Private Sub txtCodigoDeBarras_GotFocus()
    Me.Recalc
End Sub

Private Sub txtDescontoItem_Exit(Cancel As Integer)
    
    CadastrarItensDaVenda

End Sub
Private Sub cmdRecebimentos_Click()
On Error GoTo Err_cmdRecebimentos_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "VendasRecebimentos"
    
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    stLinkCriteria = "[codVenda]=" & Me![txtcodVenda]
    
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
'    DoCmd.Close acForm, Me.Name, acSaveNo
    
    
Exit_cmdRecebimentos_Click:
    Exit Sub

Err_cmdRecebimentos_Click:
    MsgBox Err.Description
    Resume Exit_cmdRecebimentos_Click
    
End Sub

Private Sub CancelarVenda()
Dim strSQL As String

'Excluir Vendas Sem Produtos Ou Recebimentos
strSQL = "DELETE Vendas.codVenda, *" & _
         "FROM Vendas " & _
         "WHERE (((Vendas.codVenda) " & _
         "In (Select codVenda from " & _
         "(SELECT Vendas.codVenda FROM Vendas LEFT JOIN VendasItens ON Vendas.codVenda = VendasItens.codVenda " & _
         "WHERE (((VendasItens.codVenda) Is Null))) as tmpVendasXItens) " & _
         "Or (Vendas.codVenda) In (Select codVenda from (SELECT Vendas.codVenda " & _
         "FROM Vendas LEFT JOIN VendasRecebimentos ON Vendas.codVenda = VendasRecebimentos.codVenda " & _
         "WHERE (((VendasRecebimentos.codVenda) Is Null)))  as tmpVendasXRecebimentos)))"


ExecutarSQL strSQL
End Sub
