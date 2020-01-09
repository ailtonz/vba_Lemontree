Option Compare Database

Global intProduto As Integer

Public Sub GerarPlanilhaCodigoDeBarrasAutomatico()

Dim strSQL As String
Dim rProdutos As DAO.Recordset

strSQL = "SELECT CodigoDeBarras.codBarras, CodigoDeBarras.Artigo, CodigoDeBarras.Composicao, CodigoDeBarras.Modelo, CodigoDeBarras.Cor, CodigoDeBarras.Tamanho, CodigoDeBarras.ValorUnitario, CodigoDeBarras.qtd" & _
         " FROM CodigoDeBarras where qtd > 0 ORDER BY CodigoDeBarras.Artigo, CodigoDeBarras.Composicao, CodigoDeBarras.Modelo, CodigoDeBarras.Cor, CodigoDeBarras.Tamanho;"


Set rProdutos = CurrentDb.OpenRecordset(strSQL)

While Not rProdutos.EOF

    For x = 1 To rProdutos.Fields("qtd")
        strSQL = "INSERT INTO tmpCodigoDeBarras ( codBarras,Artigo,Composicao,Modelo,Tamanho,Cor,ValorUnitario ) " & _
                    " Values ('" & rProdutos.Fields("codBarras") & "','" & _
                    rProdutos.Fields("Artigo") & "','" & _
                    rProdutos.Fields("Composicao") & "','" & _
                    rProdutos.Fields("Modelo") & "','" & _
                    rProdutos.Fields("Tamanho") & "','" & _
                    rProdutos.Fields("Cor") & "','" & _
                    rProdutos.Fields("ValorUnitario") & "')"
                    
        ExecutarSQL strSQL
    Next x
    rProdutos.MoveNext
Wend

rProdutos.Close

Set rProdutos = Nothing


End Sub

Public Function ImportarProdutos()
Dim fd As Office.FileDialog
Dim strArq As String
Dim strNomeDaTabela As String
Dim strCaminhoArquivo As String

On Error GoTo ErrHandler

Set fd = Application.FileDialog(msoFileDialogFilePicker)
strNomeDaTabela = "admPlanilhaProdutos"

' Open the file dialog
With fd
    .Title = "Importar >> Planilha de Produtos <<  do Ponto de Venda - ATENÇÃO: SELECIONAR PLANILHA DE PRODUTOS!"
    .Filters.Add "Arquivos Excel", "*.xls"
    .AllowMultiSelect = False

    If fd.Show = -1 Then
        strCaminhoArquivo = fd.SelectedItems(1)
    End If
    
    If strCaminhoArquivo <> "" Then
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel8, strNomeDaTabela, strCaminhoArquivo, True
        MsgBox "A Importação da Planilha de Produtos foi concluída.", vbOKOnly + vbInformation, "Importação de >> Planilha de Produtos << "
    End If

End With


ExitHere:
Exit Function

ErrHandler:
MsgBox Err.Description
Resume ExitHere

End Function

Public Function ImportarVendas()
Dim fd As Office.FileDialog
Dim strArq As String
Dim strNomeDaTabela As String
Dim strCaminhoArquivo As String

On Error GoTo ErrHandler

Set fd = Application.FileDialog(msoFileDialogFilePicker)
strNomeDaTabela = "admPlanilhaVendas"

' Open the file dialog
With fd
    .Title = "Importar >> Planilha de Vendas <<  do Ponto de Venda - ATENÇÃO: SELECIONAR PLANILHA DE VENDAS!"
    .Filters.Add "Arquivos Excel", "*.xls"
    .AllowMultiSelect = False

    If fd.Show = -1 Then
        strCaminhoArquivo = fd.SelectedItems(1)
    End If
    
    If strCaminhoArquivo <> "" Then
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel8, strNomeDaTabela, strCaminhoArquivo, True
        MsgBox "A Importação da Planilha de Vendas foi concluída.", vbOKOnly + vbInformation, "Importação de >> Planilha de Vendas << "
    End If

End With


ExitHere:
Exit Function

ErrHandler:
MsgBox Err.Description
Resume ExitHere

End Function

Public Function ImportarCodigoDeBarras()
Dim fd As Office.FileDialog
Dim strArq As String
Dim strNomeDaTabela As String
Dim strCaminhoArquivo As String

On Error GoTo ErrHandler

Set fd = Application.FileDialog(msoFileDialogFilePicker)
strNomeDaTabela = "CodigoDeBarras"

' Open the file dialog
With fd
    .Title = "Importar Código de Barras - ATENÇÃO: SELECIONAR PLANILHA DE CÓDIGO DE BARRAS!"
    .Filters.Add "Arquivos Excel", "*.xls"
    .AllowMultiSelect = False

    If fd.Show = -1 Then
        strCaminhoArquivo = fd.SelectedItems(1)
    End If
    
    If strCaminhoArquivo <> "" Then
        
'        ExecutarSQL "Delete from CodigoDeBarras"
        
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel8, strNomeDaTabela, strCaminhoArquivo, True
        
        MsgBox "A Importação dos Codigos De Barras foi concluída.", vbOKOnly + vbInformation, "Importação de Codigo De Barras"
    
    End If


End With


ExitHere:
Exit Function

ErrHandler:
MsgBox Err.Description
Resume ExitHere

End Function


Public Function ExportarCodigoDeBarras_PontoDeVenda()
    Dim sTemp As String
    Dim sSQL As String: sSQL = "admCodigoDeBarras-PontoDeVenda"
    
    sTemp = pathDesktopAddress & Format(Now, "yymmdd_hhnn") & "-CodigoDeBarras-PONTO_DE_VENDA.xls"
    
    DoCmd.Hourglass True
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel97, sSQL, sTemp, True
    
    DoCmd.Hourglass False
    MsgBox "A planilha foi gerada com êxito." & vbCrLf & vbCrLf & "Está em " & sTemp, vbInformation, "Exportação de Codigo De Barras"
    
End Function
 

Public Function ExportarCodigoDeBarras_Etiquetas()
    Dim sTemp As String
    Dim sSQL As String: sSQL = "admCodigoDeBarras-Etiquetas"
    
    sTemp = pathDesktopAddress & Format(Now, "yymmdd_hhnn") & "-CodigoDeBarras-ETIQUETAS.xls"
    
    DoCmd.Hourglass True
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel97, sSQL, sTemp, True
    
    DoCmd.Hourglass False
    MsgBox "A planilha foi gerada com êxito." & vbCrLf & vbCrLf & "Está em " & sTemp, vbInformation, "Exportação de Codigo De Barras - ETIQUETAS"
    
End Function

Public Function ExportarRecebimentos()
    Dim sTemp As String
    
    sTemp = pathDesktopAddress & Format(Now, "yymmdd_hhnn") & "-Vendas_" & Left(CurrentMDB(), Len(CurrentMDB()) - 4) & ".xls"
    
    DoCmd.Hourglass True
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel97, "admVendasRecebimentos", sTemp, True
    
    DoCmd.Hourglass False
    MsgBox "A planilha foi gerada com êxito." & vbCrLf & vbCrLf & "Está em " & sTemp, vbInformation, "Exportação de Vendas"
    
End Function


Public Function ExportarVendas()
    Dim sTemp As String
    
    sTemp = pathDesktopAddress & Format(Now, "yymmdd_hhnn") & "-Produtos_" & Left(CurrentMDB(), Len(CurrentMDB()) - 4) & ".xls"
    
    DoCmd.Hourglass True
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel97, "admVendasProdutos", sTemp, True
    
    DoCmd.Hourglass False
    MsgBox "A planilha foi gerada com êxito." & vbCrLf & vbCrLf & "Está em " & sTemp, vbInformation, "Exportação de Produtos"

End Function


Function GerarParcelamento(codVenda As Long, dtEmissao As Date, ValorPago As Currency, ValorRecebido As Currency, Parcelamento As String, strEspecie As String)

'Dim Valor As String
'Valor = "30"

Dim matriz As Variant
Dim x As Integer
Dim Parcelas As DAO.Recordset

Set Parcelas = CurrentDb.OpenRecordset("Select * from VendasRecebimentos")

matriz = Array()
matriz = Split(Parcelamento, ";")

BeginTrans

For x = 0 To UBound(matriz)
    y = x + 1
    Parcelas.AddNew
    Parcelas.Fields("codVenda") = codVenda
    Parcelas.Fields("Vencimento") = CalcularVencimento2(dtEmissao, CInt(matriz(x)))
    Parcelas.Fields("ValorPago") = ValorPago / (UBound(matriz) + 1)
    Parcelas.Fields("ValorRecebido") = ValorRecebido / (UBound(matriz) + 1)
    Parcelas.Fields("Parcelamento") = "" & y & "/" & (UBound(matriz) + 1) & ""
    Parcelas.Fields("Especie") = strEspecie
    Parcelas.Update
Next

CommitTrans

Parcelas.Close

End Function
