Option Compare Database

Private Sub cboEspecie_Click()
Dim strSQL As String
Dim ctlParcelamento As ComboBox
Dim ctlEspecie As ComboBox
    
Set ctlParcelamento = Me.cboParcelamento
Set ctlEspecie = Me.cboEspecie
   
ctlParcelamento.Value = ""
   
strSQL = "SELECT admCategorias.Categoria, admCategorias.Descricao01 " & _
         "FROM admCategorias WHERE (((admCategorias.codCategoria) In " & _
         "(Select codRelacao from admCategorias where Categoria = '" & ctlEspecie.Column(0) & "')));"

ctlParcelamento.RowSource = strSQL
ctlParcelamento.Requery


Set ctlParcelamento = Nothing
Set ctlEspecie = Nothing

Me.txtReceber = (Me.txtValorFinal - (Me.txtDescontoGeral + Me.txtRecebimentos))

End Sub

Private Sub Receber()

Dim strSQL As String

Dim ValorPago As Currency
Dim ValorRecebido As Currency

Dim ctlParcelamento As ComboBox
Dim ctlEspecie As ComboBox
    
Set ctlParcelamento = Me.cboParcelamento
Set ctlEspecie = Me.cboEspecie

ValorPago = Me.txtReceber

ValorRecebido = ValorPago - (ValorPago / 100 * ctlEspecie.Column(2))

GerarParcelamento Me.codVenda, Format(Now(), "dd/mm/yy"), ValorPago, ValorRecebido, IIf(IsNull(ctlParcelamento.Column(1)), 1, ctlParcelamento.Column(1)), ctlEspecie.Column(1)

Me.subVendasRecebimentos.Requery
Me.Recalc

End Sub

Private Sub cmdReceber_Click()
'    Me.txtDescontoGeral = 0
    Receber
End Sub
