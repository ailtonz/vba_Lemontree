Option Compare Database

Private Sub Artigo_Click()
    Me.Artigos = Me.Artigo.Column(1)
End Sub

Private Sub Artigo_NotInList(NewData As String, Response As Integer)
'Permite adicionar item à lista
Dim DB As DAO.Database
Dim rst As DAO.Recordset
Dim Resposta As Variant

On Error GoTo ErrHandler

'Pergunta se deseja acrescentar o novo item
Resposta = MsgBox("O Artigo : ' " & NewData & " '  não faz parte do cadastro." & vbCrLf & vbCrLf & "Deseja acrescentá-la?", vbYesNo + vbQuestion, "Artigos")
If Resposta = vbYes Then
    Set DB = CurrentDb()
    'Abre a tabela, adiciona o novo item e atualiza a combo
    Set rst = DB.OpenRecordset("admCategorias")
    With rst
        .AddNew
        !Descricao01 = NewData
        !codCategoria = NovoCodigo("admCategorias", "codCategoria")
        !Categoria = "Artigos"
        !Principal = 0
        .Update
        Response = acDataErrAdded
        .Close
    End With
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "subArtigos"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.GoToRecord , , acLast
    
End If

ExitHere:
Set rst = Nothing
Set DB = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description & vbCrLf & Err.Number & _
vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
Resume ExitHere

End Sub

Private Sub cmdCores_Click()
On Error GoTo Err_cmdCores_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    DoCmd.RunCommand acCmdSaveRecord
'    intProduto = Me.txtcodProduto
    stDocName = "subSelecaoDeCores"
    
    stLinkCriteria = "[codProduto]=" & Me![codProduto]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdCores_Click:
    Exit Sub

Err_cmdCores_Click:
    MsgBox Err.Description
    Resume Exit_cmdCores_Click
    
End Sub

Private Sub cmdTamanhos_Click()
On Error GoTo Err_cmdTamanhos_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    DoCmd.RunCommand acCmdSaveRecord
    intProduto = Me.txtcodProduto
    stDocName = "subProdutosTamanhos"
    
    stLinkCriteria = "[codProduto]=" & Me![txtcodProduto]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdTamanhos_Click:
    Exit Sub

Err_cmdTamanhos_Click:
    MsgBox Err.Description
    Resume Exit_cmdTamanhos_Click
    
End Sub

Private Sub Composicao_Click()
    Me.Composicoes = Me.Composicao.Column(1)
End Sub

Private Sub Composicao_NotInList(NewData As String, Response As Integer)
'Permite adicionar item à lista
Dim DB As DAO.Database
Dim rst As DAO.Recordset
Dim Resposta As Variant

On Error GoTo ErrHandler

'Pergunta se deseja acrescentar o novo item
Resposta = MsgBox("A Composição : ' " & NewData & " '  não faz parte do cadastro." & vbCrLf & vbCrLf & "Deseja acrescentá-la?", vbYesNo + vbQuestion, "Composiões")
If Resposta = vbYes Then
    Set DB = CurrentDb()
    'Abre a tabela, adiciona o novo item e atualiza a combo
    Set rst = DB.OpenRecordset("admCategorias")
    With rst
        .AddNew
        !Descricao01 = NewData
        !codCategoria = NovoCodigo("admCategorias", "codCategoria")
        !Categoria = "Composições"
        !Principal = 0
        .Update
        Response = acDataErrAdded
        .Close
    End With
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "subComposicoes"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.GoToRecord , , acLast
    
End If

ExitHere:
Set rst = Nothing
Set DB = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description & vbCrLf & Err.Number & _
vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
Resume ExitHere

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    If Me.NewRecord Then Me.txtcodProduto = NovoCodigo(Me.RecordSource, Me.txtcodProduto.ControlSource)
End Sub

Private Sub Form_Open(Cancel As Integer)
'    ExecutarSQL "Drop Table tmpProdutos"
    DoCmd.Maximize
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click
Dim strSQL As String
    
'    strSQL = "SELECT DISTINCT qryProdutosCodigoDeBarras.codBarras, qryProdutosCodigoDeBarras.Produto, Format([ValorUnitario],'Standard') AS Valor_Unitario INTO tmpProdutos FROM qryProdutosCodigoDeBarras"
'    ExecutarSQL strSQL
    DoCmd.Close


Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Private Sub Modelo_Click()
    Me.Modelos = Me.Modelo.Column(1)
End Sub

Private Sub Modelo_NotInList(NewData As String, Response As Integer)
'Permite adicionar item à lista
Dim DB As DAO.Database
Dim rst As DAO.Recordset
Dim Resposta As Variant

On Error GoTo ErrHandler

'Pergunta se deseja acrescentar o novo item
Resposta = MsgBox("O Modelo : ' " & NewData & " '  não faz parte do cadastro." & vbCrLf & vbCrLf & "Deseja acrescentá-lo?", vbYesNo + vbQuestion, "Composiões")
If Resposta = vbYes Then
    Set DB = CurrentDb()
    'Abre a tabela, adiciona o novo item e atualiza a combo
    Set rst = DB.OpenRecordset("admCategorias")
    With rst
        .AddNew
        !Descricao01 = NewData
        !codCategoria = NovoCodigo("admCategorias", "codCategoria")
        !Categoria = "Modelos"
        !Principal = 0
        .Update
        Response = acDataErrAdded
        .Close
    End With
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "subModelos"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.GoToRecord , , acLast
    
End If

ExitHere:
Set rst = Nothing
Set DB = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description & vbCrLf & Err.Number & _
vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
Resume ExitHere

End Sub

