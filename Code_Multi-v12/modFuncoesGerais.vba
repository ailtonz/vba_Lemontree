Option Compare Database
Option Explicit

Public strTabela As String


Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = Tabela
       
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function

Public Function RedimencionaControle(frm As Form, ctl As Control)

Dim intAjuste As Integer
On Error Resume Next

intAjuste = frm.Section(acHeader).Height * frm.Section(acHeader).Visible

intAjuste = intAjuste + frm.Section(acFooter).Height * frm.Section(acFooter).Visible

On Error GoTo 0

intAjuste = Abs(intAjuste) + ctl.top

If intAjuste < frm.InsideHeight Then
    ctl.Height = frm.InsideHeight - intAjuste
'    ctl.Width = frm.InsideHeight + (intAjuste + intAjuste)
End If

End Function

Public Function EstaAberto(strName As String) As Boolean
On Error GoTo EstaAberto_Err
' Testa se o formulário está aberto

   Dim obj As AccessObject, dbs As Object
   Set dbs = Application.CurrentProject
   ' Procurar objetos AccessObject abertos na coleção AllForms.
   
   EstaAberto = False
   For Each obj In dbs.AllForms
        If obj.IsLoaded = True And obj.Name = strName Then
            ' Imprimir nome do obj.
            EstaAberto = True
            Exit For
        End If
   Next obj
    
EstaAberto_Fim:
  Exit Function
EstaAberto_Err:
  Resume EstaAberto_Fim
End Function

Public Function IsFormView(frm As Form) As Boolean
On Error GoTo IsFormView_Err
' Testa se o formulário está aberto em
' modo formulário (form view)

 IsFormView = False
 If frm.CurrentView = 1 Then
    IsFormView = True
 End If

IsFormView_Fim:
  Exit Function
IsFormView_Err:
  Resume IsFormView_Fim
End Function

Public Function pathDesktopAddress() As String
    pathDesktopAddress = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
End Function

Public Function ExecutarSQL(strSQL As String, Optional log As Boolean)
'Objetivo: Executar comandos SQL sem mostrar msg's do access.

'Desabilitar menssagens de execução de comando do access
DoCmd.SetWarnings False

If log Then GerarSaida strSQL, "sql.log"

'Executar a instrução SQL
DoCmd.RunSQL strSQL

'Abilitar menssagens de execução de comando do access
DoCmd.SetWarnings True

End Function

Public Function GerarSaida(strConteudo As String, strArquivo As String)

Open Application.CurrentProject.Path & "\" & strArquivo For Append As #1

Print #1, strConteudo

Close #1

End Function

Public Function NovoCodigo(Tabela, Campo)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")
If Not rstTabela.EOF Then
   NovoCodigo = rstTabela.Fields("CodigoNovo")
   If IsNull(NovoCodigo) Then
      NovoCodigo = 1
   End If
Else
   NovoCodigo = 1
End If
rstTabela.Close

End Function

Public Function CalcularVencimento(Dia As Integer, Optional MES As Integer, Optional Ano As Integer) As Date

If Month(Now) = 2 Then
    If Dia = 29 Or Dia = 30 Or Dia = 31 Then
        Dia = 1
        MES = MES + 1
    End If
End If

If MES > 0 And Ano > 0 Then
    CalcularVencimento = Format((DateSerial(Ano, MES, Dia)), "dd/mm/yyyy")
ElseIf MES = 0 And Ano > 0 Then
    CalcularVencimento = Format((DateSerial(Ano, Month(Now), Dia)), "dd/mm/yyyy")
ElseIf MES = 0 And Ano = 0 Then
    CalcularVencimento = Format((DateSerial(Year(Now), Month(Now), Dia)), "dd/mm/yyyy")
End If

End Function

Public Function CalcularVencimento2(dtInicio As Date, qtdDias As Integer, Optional ForaMes As Boolean) As Date

Dim MyDate

    If ForaMes Then
        MyDate = Format((DateSerial(Year(dtInicio), Month(dtInicio) + 1, qtdDias)), "dd/mm/yyyy")
        CalcularVencimento2 = MyDate
    Else
        MyDate = Format((DateSerial(Year(dtInicio), Month(dtInicio), Day(dtInicio) + qtdDias)), "dd/mm/yyyy")
        
        If Weekday(MyDate) = 1 Then ' Domingo
            CalcularVencimento2 = Format((DateSerial(Year(dtInicio), Month(dtInicio), Day(dtInicio) + qtdDias + 1)), "dd/mm/yyyy")
        ElseIf Weekday(MyDate) = 7 Then ' Sabado
            CalcularVencimento2 = Format((DateSerial(Year(dtInicio), Month(dtInicio), Day(dtInicio) + qtdDias + 2)), "dd/mm/yyyy")
        Else 'Dia da semana
            CalcularVencimento2 = MyDate
        End If
        
    End If

End Function


'Public Function DiaSemana(ByVal vsData As String) As String
'
'   Dim iDia As Integer
'
'   iDia = Weekday(vsData)
'
'   Select Case iDia
'
'   Case iDia = 1
'        DiaSemana = "Domingo"
'   Case iDia = 2
'        DiaSemana = "Segunda"
'   Case iDia = 3
'        DiaSemana = "Terça"
'   Case iDia = 4
'        DiaSemana = "Quarta"
'   Case iDia = 5
'        DiaSemana = "Quinta"
'   Case iDia = 6
'        DiaSemana = "Sexta"
'   Case iDia = 7
'        DiaSemana = "Sábado"
'
'   End Select
'
'End Function
