Option Compare Database

Private Sub cmdRelacao_Click()
    Call Relacao
End Sub


Private Sub Relacao()

Dim ctlArtigos As ListBox
Dim ctlComposicoes As ListBox
Dim ctlModelos As ListBox
Dim ctlTamanhos As ListBox
Dim ctlCores As ListBox
Dim ctlCodigosDeBarras As SubForm

Dim ctlQuantidade As TextBox
Dim ctlValorUnitario As TextBox


Dim intArtigos As Integer
Dim intComposicoes As Integer
Dim intModelos As Integer
Dim intTamanhos As Integer
Dim intCores As Integer

Dim intQuantidade As Integer

Dim strSQL As String
Dim strCodigo As String
   
Set ctlArtigos = Me.lstArtigos
Set ctlComposicoes = Me.lstComposicoes
Set ctlModelos = Me.lstModelos
Set ctlTamanhos = Me.lstTamanhos
Set ctlCores = Me.lstCores
Set ctlCodigosDeBarras = Me.subCodigoDeBarras

Set ctlQuantidade = Me.txtQuantidade
Set ctlValorUnitario = Me.txtValorUnitario
    
 
For intArtigos = 0 To ctlArtigos.ListCount - 1
    If ctlArtigos.Selected(intArtigos) Then
    For intComposicoes = 0 To ctlComposicoes.ListCount - 1
        If ctlComposicoes.Selected(intComposicoes) Then
        For intModelos = 0 To ctlModelos.ListCount - 1
            If ctlModelos.Selected(intModelos) Then
            For intTamanhos = 0 To ctlTamanhos.ListCount - 1
                If ctlTamanhos.Selected(intTamanhos) Then
                For intCores = 0 To ctlCores.ListCount - 1
                    For intQuantidade = 1 To ctlQuantidade.Value
                        If ctlCores.Selected(intCores) Then
                        
                            strCodigo = ctlArtigos.Column(1, intArtigos) & _
                                        ctlComposicoes.Column(1, intComposicoes) & _
                                        ctlModelos.Column(1, intModelos) & _
                                        ctlTamanhos.Column(1, intTamanhos) & _
                                        ctlCores.Column(1, intCores)
                        
                            strSQL = "INSERT INTO tmpCodigoDeBarras ( codBarras,Artigo,Composicao,Modelo,Tamanho,Cor,ValorUnitario ) " & _
                                        " Values ('" & strCodigo & "','" & _
                                        ctlArtigos.Column(0, intArtigos) & "','" & _
                                        ctlComposicoes.Column(0, intComposicoes) & "','" & _
                                        ctlModelos.Column(0, intModelos) & "','" & _
                                        ctlTamanhos.Column(0, intTamanhos) & "','" & _
                                        ctlCores.Column(0, intCores) & "','" & _
                                        ctlValorUnitario.Value & "')"

                            ExecutarSQL strSQL
                            ctlCodigosDeBarras.Requery
                        End If
                    Next intQuantidade
                Next intCores
                End If
            Next intTamanhos
            End If
        Next intModelos
        End If
    Next intComposicoes
    End If
Next intArtigos




'    For intOrigem = 0 To ctlOrigem.ListCount - 1
'        If ctlOrigem.Selected(intOrigem) Then
'            For intDestino = 0 To ctlDestino.ListCount - 1
'                If ctlDestino.Selected(intDestino) Then
'                    strSQL = "INSERT INTO admCategorias ( codRelacao,Categoria,Descricao01,Descricao02 ) SELECT " & _
'                              ctlDestino.Column(0, intDestino) & " as Relacao,  " & _
'                              " codCategoria,'" & Me.cboCategoriasSecundario.Column(1) & "' as Descricao1 ,'" & ctlDestino.Column(1, intDestino) & "' as Descricao2 FROM admCategorias where codCategoria = " & ctlOrigem.Column(0, intOrigem) & ""
'                    ExecutarSQL strSQL, True
'                    ctlDestino.Selected(intDestino) = False
'                End If
'            Next intDestino
'            ctlOrigem.Selected(intOrigem) = False
'            Me.lstRelacionamentos.Requery
'        End If
'    Next intOrigem
'

Set ctlArtigos = Nothing
Set ctlComposicoes = Nothing
Set ctlModelos = Nothing
Set ctlTamanhos = Nothing
Set ctlCores = Nothing
Set ctlCodigosDeBarras = Nothing

Set ctlQuantidade = Nothing
Set ctlValorUnitario = Nothing


End Sub

Private Sub lstRelacionamentos_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim strSQL As String
'Dim ctlRelacao As ListBox
'
'Set ctlRelacao = Me.lstRelacionamentos
'
'    Select Case KeyCode
'
'        Case vbKeyDelete
'
'            strSQL = "Delete * from admCategorias where codCategoria = " & ctlRelacao.Column(0)
'            ExecutarSQL strSQL
'            ctlRelacao.Requery
'
'    End Select
End Sub

Private Sub cmdGerar_Click()
    Call Relacao
End Sub

Private Sub cmdPlanilha_Click()
    ExportarCodigoDeBarras_Etiquetas
End Sub

Private Sub cmdVendas_Click()
Dim sSQL As String

sSQL = "INSERT INTO CodigoDeBarras ( codBarras, Artigo, Composicao, Modelo, Cor, Tamanho, ValorUnitario )" & _
        " SELECT Distinct admNovosCodigosDeBarras.codBarras, admNovosCodigosDeBarras.Artigo, admNovosCodigosDeBarras.Composicao, " & _
        " admNovosCodigosDeBarras.Modelo, admNovosCodigosDeBarras.Cor, admNovosCodigosDeBarras.Tamanho, " & _
        " admNovosCodigosDeBarras.ValorUnitario FROM admNovosCodigosDeBarras;"


ExecutarSQL sSQL

sSQL = "Delete * from tmpCodigoDeBarras"

ExecutarSQL sSQL

Me.subCodigoDeBarras.Requery

End Sub
