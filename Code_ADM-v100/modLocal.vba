Option Compare Database

Global intProduto As Integer


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


Public Function ExportarCodigoDeBarras()
    Dim sTemp As String
    
    sTemp = pathDesktopAddress & Format(Now, "yymmdd_hhnn") & "-CodigoDeBarras-PONTO_DE_VENDA.xls"
    
    DoCmd.Hourglass True
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel97, "admCodigoDeBarras", sTemp, True
    
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

'Public Function GerarEtiquetas(sSQL As String)
'
'Dim DB As DAO.Database
'Dim Rd As DAO.Recordset
'
'Dim XPlanilha As Object
'
'Dim iLinha As Integer
'Dim intCampos As Integer
'Dim i As Integer
'
'Dim Count As Long
'
'Dim sTemp As String
'Dim arqModelo As String
'
'Set DB = CurrentDb
'Set Rd = DB.OpenRecordset(sSQL)
'
'
'If Not Rd.EOF Then
'
'    Rd.MoveLast
'
'    Count = Rd.RecordCount
'
'    Rd.MoveFirst
'
'    If Count > 0 Then
'
'        DoEvents
'
'        Dim s As Variant
'        Dim C As Long
'
'        'Cria referencia ao EXCEL
'        Set XPlanilha = CreateObject("Excel.Application")
'
'        ''##################
'        ''Arquivo Modelo
'        ''##################
'
''            arqModelo = Application.CurrentProject.Path & "\" & Modelo
'
'            'Abre o arquivo modelo
'            XPlanilha.Workbooks.Add '.Open (arqModelo)
'
'            'Seleciona a primeira planilha
'            XPlanilha.Workbooks(1).Sheets(1).Select
'
'            'Incrementa a linha
'            iLinha = 1
'
'        ''##################
'        ''Transfere os dados
'        ''##################
'
'            DoCmd.Hourglass True
'
'            intCampos = Rd.Fields.Count
'
'            s = SysCmd(acSysCmdInitMeter, "Exportando " & Count & " Registros", Count)
'
'            Do While Not Rd.EOF
'
'                i = 0
'
'                XPlanilha.Cells(1, 1).Value = "Código de barras"
'                XPlanilha.Cells(1, 2).Value = "Artigo"
'                XPlanilha.Cells(1, 3).Value = "Cor"
'                XPlanilha.Cells(1, 4).Value = "Tamanho"
'                XPlanilha.Cells(1, 5).Value = "Valor Unitario"
'
''                For r = 1 To Rd.Fields("Saldo")
'                    iLinha = iLinha + 1 'incrementa a linha
'                    XPlanilha.Cells(iLinha, 1).Value = Rd.Fields("codBarras")
'                    XPlanilha.Cells(iLinha, 2).Value = Rd.Fields("Artigo")
'                    XPlanilha.Cells(iLinha, 3).Value = Rd.Fields("Cor")
'                    XPlanilha.Cells(iLinha, 4).Value = Rd.Fields("Tamanho")
'                    XPlanilha.Cells(iLinha, 5).Value = FormatCurrency(Rd.Fields("ValorUnitario"))
''                Next
'
'
'
'                s = SysCmd(acSysCmdUpdateMeter, iLinha)
'                Rd.MoveNext
'
'            Loop
'
'            s = SysCmd(acSysCmdRemoveMeter)
'
'
'        ''##################
'        ''Formata Arquivo
'        ''##################
'
'            'Formata novo nome da planilha
'            sTemp = pathDesktopAddress & Format(Now, "yymmdd_hhnn") & "-CodigoDeBarras.xls"
'            'Se o arquivo já existe, deleta
'            If Dir$(sTemp) <> "" Then Kill sTemp
'
'
'        ''##################
'        ''Salva
'        ''##################
'
'            XPlanilha.ActiveWorkbook.SaveAs Filename:=sTemp, _
'            FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
'            ReadOnlyRecommended:=False, CreateBackup:=False
'
'        ''##################
'        ''Fecha o Excel
'        ''##################
'            XPlanilha.Quit
'
'        ''######################
'        ''Descarrega da memória
'        ''######################
'            Set XPlanilha = Nothing
'
'        DoCmd.Hourglass False
'        MsgBox "A planilha foi gerada com êxito." & vbCrLf & vbCrLf & "Está em " & sTemp, vbInformation, "ATENÇÃO"
'
'    Else
'        DoCmd.Hourglass False
'        MsgBox "Não há dados para gerar a planilha.", vbInformation, "ATENÇÃO"
'
'    End If
'
'Else
'
'    MsgBox "Não há Registros!", vbOKOnly + vbInformation, "Exportar para Excel"
'
'End If
'
'Rd.Close
''rRelatorios.Close
'
'Set Rd = Nothing
'Set DB = Nothing
''Set rRelatorios = Nothing
'Set XPlanilha = Nothing
'
'
'End Function
'
