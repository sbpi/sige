<%@ Language=VBScript %>
<!--#INCLUDE FILE="Constants.inc" -->
<!--#INCLUDE FILE="jScript.asp" -->
<!--#INCLUDE FILE="Funcoes.asp" -->
<!--#INCLUDE FILE="FuncoesGR.asp" -->
<!--#INCLUDE FILE="Constants_ADO.inc"-->
<!--#INCLUDE FILE="esc.inc"-->
<%

Response.Expires = -1500
REM =========================================================================
REM  /GR_Log.asp
REM ------------------------------------------------------------------------
REM Nome     : Egisberto Vicente da Silva
REM Descricao: Gerencia o módulo de Log
REM Home     : http://www.sbpi.com.br/
REM Mail     : beto@sbpi.com.br
REM Criacao  : 01/10/2004 13:04
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM -------------------------------------------------------------------------
REM
REM Parâmetros recebidos:
REM    R (referência) = usado na rotina de gravação, com conteúdo igual ao parâmetro T
REM    O (operação)   = L   : Listagem
REM                   = P   : Filtragem
REM                   = W   : Geração de documento no formato MS-Word (Office 2003)

  Private w_EA
  Private w_IN
  Private w_EF, w_troca
  Private SQL, CL, DBMS, RS, p_agrega
  Private w_Data, w_pagina
  Dim w_tp, w_Disabled
  Dim p_regional, p_escola, p_bdados, p_operacao, p_inicio, p_fim, p_secretaria
  
  p_agrega   = uCase(Request("p_agrega"))
  p_regional = Request("p_regional")
  p_escola   = Request("p_escola")
  p_bdados   = Request("p_bdados")
  p_operacao = Request("p_operacao")
  p_inicio   = Request("p_inicio")
  p_fim      = Request("p_fim")
  w_troca    = Request("w_troca")
  

  
  w_EA       = ucase(Request("w_ea"))
  w_EW       = ucase(Request("w_ew"))
  w_IN       = ucase(Request("w_in"))
  w_EF       = ucase(Request("w_ef"))
  CL         = Request("CL")
  w_troca    = Request("w_troca")
  w_Disabled = "ENABLED"
  P4         = cDbl(Nvl(Request("P4"),30))
  P3         = cDbl(Nvl(Request("P3"),1))

  w_Data = Mid(100+Day(Date()),2,2) & "/" & Mid(100+Month(Date()),2,2) & "/" &Year(Date())
  w_Pagina = ExtractFileName(Request.ServerVariables("SCRIPT_NAME")) & "?w_ew="
  
  AbreSessao
    
  If w_ea = "A" Then w_ea = "P" End If

  Main
  
  FechaSessao

  Set w_EA      = Nothing
  Set w_IN      = Nothing
  Set w_EF      = Nothing
  Set CL        = Nothing
  
  Set SQL       = Nothing
  Set RS        = Nothing
  Set DBMS      = Nothing
  Set w_Data    = Nothing
  Set w_Pagina  = Nothing

REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main
  
  Server.ScriptTimeOut = conScriptTimeout
  Session.TimeOut      = conSessionTimeout
 
  MainBody

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Pesquisa gerencial
REM -------------------------------------------------------------------------
Sub Gerencial
  
  If Nvl(p_regional,"nulo") = "nulo" Then
     If Mid(Session("username"),1,2) = "RE" Then p_regional = replace(CL,"sq_cliente=","") Else p_regional = "" End If
  End If
    
  If w_ea = "L" or w_ea = "W" Then

     w_filtro = ""
     If p_regional > ""  Then 
        SQL = "select * from escCliente where sq_cliente = " & p_regional
        ConectaBD SQL
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Regional <td><font size=1>[<b>" & RS("ds_cliente") & "</b>]"
        DesconectaBD
     End If
     If p_escola > ""  Then 
        SQL = "select * from escCliente where sq_cliente = " & p_escola
        ConectaBD SQL
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Escola <td><font size=1>[<b>" & RS("ds_cliente") & "</b>]"
        DesconectaBD
     End If
     If p_bdados > ""  Then 
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Bloco de dados <td><font size=1>[<b>" & ExibeBlocoDados(p_bdados) & "</b>]"
     End If
     If p_operacao > ""  Then 
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Operação <td><font size=1>[<b>" & ExibeOperacao(p_operacao) & "</b>]"
     End If
     If p_inicio > ""  Then 
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Período <td><font size=1>[<b>" & p_inicio & " a " & p_fim & "</b>]"
     End If
     If w_filtro > "" Then w_filtro = "<table border=0><tr valign=""top""><td><font size=1><b>Filtro:</b><td nowrap><font size=1><ul>" & w_filtro & "</ul></tr></table>" End If

     SQL = "select a.*, " & VbCrLf & _
           "       case a.tipo " & VbCrLf & _
           "            when 0 then 'Consulta' " & VbCrLf & _
           "            when 1 then 'Inclusão' " & VbCrLf & _
           "            when 2 then 'Alteração' " & VbCrLf & _
           "            when 3 then 'Exclusão' " & VbCrLf & _
           "            else        'Erro' " & VbCrLf & _
           "       end nm_tipo, " & VbCrLf & _
           "       IsNull(b.nome,'Consulta') nm_funcionalidade,   IsNull(b.codigo,0) cd_funcionalidade, " & VbCrLf & _
           "       c.sq_cliente sq_escola,     c.ds_cliente nm_escola, " & VbCrLf & _
           "       d.sq_cliente sq_regional,   d.ds_cliente nm_regional, " & VbCrLf & _
           "       e.sq_cliente sq_secretaria, e.ds_cliente nm_secretaria " & VbCrLf & _
           "  from escCliente_Log                    a " & VbCrLf & _
           "       left outer join escFuncionalidade b on (a.sq_funcionalidade = b.sq_funcionalidade and " & VbCrLf & _
           "                                               b.tipo              = 2 " & VbCrLf & _
           "                                              ) " & VbCrLf & _
           "       inner      join escCliente        c on (a.sq_cliente        = c.sq_cliente) " & VbCrLf & _
           "         inner   join escCliente         d on (c.sq_cliente_pai    = d.sq_cliente) " & VbCrLf & _
           "           inner join escCliente         e on (d.sq_cliente_pai    = e.sq_cliente), " & VbCrLf & _
           "       escCliente                        f " & VbCrLf & _
           "         inner   join escTipo_Cliente    g on (f.sq_tipo_cliente   = g.sq_tipo_cliente) " & VbCrLf & _
           " where f.sq_cliente = " & replace(cl,"sq_cliente=","") & " " & VbCrLf & _
           "   and ((g.tipo = 2 and d.sq_cliente = " & replace(cl,"sq_cliente=","") & ") or " & VbCrLf & _
           "        (g.tipo <> 2) " & VbCrLf & _
           "       ) " & VbCrLf
     If p_regional > "" Then SQL = SQL & "   and d.sq_cliente = " & p_regional & VbCrLf End If
     If p_escola > ""   Then SQL = SQL & "   and c.sq_cliente = " & p_escola   & VbCrLf End If
     If p_bdados > ""   Then SQL = SQL & "   and b.codigo     = " & p_bdados   & VbCrLf End If
     If p_operacao > "" Then SQL = SQL & "   and a.tipo       = " & p_operacao & VbCrLf End If
     If p_inicio > ""   Then SQL = SQL & "   and a.data       between convert(datetime, '" & p_inicio & "', 103) and convert(datetime, '" & p_fim & "', 103) + 1 " & VbCrLf End If
     Select case p_agrega
        Case "SECRETARIA"
           SQL = SQL & "order by nm_secretaria" & VbCrLf
           ConectaBD SQL
           w_TP = TP & "Histórico de acessos por secretaria"
        Case "REGIONAL"
           SQL = SQL & "order by nm_regional" & VbCrLf
           ConectaBD SQL
           w_TP = TP & "Histórico de acessos por regional"
        Case "ESCOLA"
           SQL = SQL & "order by nm_escola" & VbCrLf
           ConectaBD SQL
           w_TP = TP & "Histórico de acessos por escola"
        Case "BLOCODADOS"
           SQL = SQL & "order by nm_funcionalidade" & VbCrLf
           ConectaBD SQL
           w_TP = TP & "Histórico de acessos por bloco de dados"
     End Select
  End If
  
  If w_ea = "W" Then
     HeaderWord
     w_pag   = 1
     w_linha = 0
     
     If w_filtro > "" Then ShowHTML w_filtro End If
  Else
     Cabecalho
     ShowHTML "<HEAD>"
     If w_ea = "P" Then
        ScriptOpen "Javascript"
        CheckBranco
        FormataData
        ValidateOpen "Validacao"
        Validate "p_inicio", "Recebimento inicial", "DATA", "", "10", "10", "", "0123456789/"
        Validate "p_fim", "Recebimento final", "DATA", "", "10", "10", "", "0123456789/"
        ShowHTML "  if ((theForm.p_inicio.value != '' && theForm.p_fim.value == '') || (theForm.p_inicio.value == '' && theForm.p_fim.value != '')) {"
        ShowHTML "     alert ('Informe ambas as datas de recebimento ou nenhuma delas!');"
        ShowHTML "     theForm.p_inicio.focus();"
        ShowHTML "     return false;"
        ShowHTML "  }"
        CompData "p_inicio", "Data inicial", "<=", "p_fim", "Data final"
        ValidateClose
        ScriptClose
     End If
     ShowHTML "</HEAD>"
     If w_Troca > "" Then ' Se for recarga da página
        BodyOpen "onLoad='document.Form." & w_Troca & ".focus();'"
     ElseIf InStr("P",w_ea) > 0 Then
        BodyOpen "onLoad='document.Form.p_agrega.focus()';"
     Else
        BodyOpenClean "onLoad=document.focus();"
     End If
     If w_ea = "L" Then
        ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
        ShowHTML "<HR>"
        If w_filtro > "" Then ShowHTML w_filtro End If
     Else
        ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
        ShowHTML "<HR>"
     End If
  End If

  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  If w_ea = "L" or w_ea = "W" Then
    If w_ea = "L" Then
       ShowHTML "<tr><td><font size=""1"">"
       If MontaFiltro("GET") > "" Then
          ShowHTML " <a accesskey=""F"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=P&P3=1&P4=" & P4 & "&TP=" & TP & "&SG=" & SG & MontaFiltro("GET") & """><u><font color=""#BC5100"">F</u>iltrar (Ativo)</font></a>"
        Else
          ShowHTML " <a accesskey=""F"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=P&p3=1&P4=" & P4 & "&TP=" & TP & "&SG=" & SG & MontaFiltro("GET") & """><u>F</u>iltrar (Inativo)</a>"
       End If
    End IF
    ImprimeCabecalho
    If RS.EOF Then
        ShowHTML "      <tr bgcolor=""" & conTrBgColor & """><td colspan=10 align=""center""><font size=""1""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      If w_ea = "L" Then
         ShowHTML "<SCRIPT LANGUAGE=""JAVASCRIPT"">"
         ShowHTML "  function lista (filtro, oper) {"
         ShowHTML "    if (filtro != -1) {"
         Select case p_agrega
            Case "SECRETARIA" ShowHTML "      document.Form.p_secretaria.value=filtro;"
            Case "REGIONAL"   ShowHTML "      document.Form.p_regional.value=filtro;"
            Case "ESCOLA"     ShowHTML "      document.Form.p_escola.value=filtro;"
            Case "BLOCODADOS" ShowHTML "      document.Form.p_bdados.value=filtro;"
         End Select
         ShowHTML "    }"
         Select case p_agrega
            Case "SECRETARIA" ShowHTML " else document.Form.p_secretaria.value='" & Request("p_secretaria") & "';"
            Case "REGIONAL"   ShowHTML " else document.Form.p_regional.value='"   & Request("p_regional")   & "';"
            Case "ESCOLA"     ShowHTML " else document.Form.p_escola.value='"     & Request("p_escola")     & "';"
            Case "BLOCODADOS" ShowHTML " else document.Form.p_bdados.value='"     & Request("p_bdados")     & "';"
         End Select
         ShowHTML "    if (oper != -1) document.Form.p_operacao.value=oper; else document.Form.p_operacao.value=''; "
         ShowHTML "    document.Form.submit();"
         ShowHTML "  }"
         ShowHTML "</SCRIPT>"
         AbreForm "Form", w_pagina & "Log&P3=1&P4=30&w_ea=L", "POST", "return(Validacao(this));", "Lista"
         ShowHTML MontaFiltro("POST")
         If Nvl(Request("p_operacao"),"nulo") = "nulo" Then
            ShowHTML "<input type=""Hidden"" name=""p_operacao"" value="""">" 
         End If
         Select case p_agrega
            Case "SECRETARIA" If Request("p_secretaria") = "" Then ShowHTML "<input type=""Hidden"" name=""p_secretaria"" value="""">" End If
            Case "REGIONAL"   If Request("p_regional") = ""   Then ShowHTML "<input type=""Hidden"" name=""p_regional"" value="""">"   End If
            Case "ESCOLA"     If Request("p_escola") = ""     Then ShowHTML "<input type=""Hidden"" name=""p_escola"" value="""">"     End If
            Case "BLOCODADOS" If Request("p_bdados") = ""     Then ShowHTML "<input type=""Hidden"" name=""p_bdados"" value="""">"     End If
            End Select
      End If
  
      RS.PageSize       = P4
      RS.AbsolutePage   = P3
      w_nm_quebra       = ""
      w_qt_quebra       = 0
      t_acesso          = 0
      t_consulta        = 0
      t_inc             = 0
      t_alt             = 0
      t_exc             = 0
      t_totacesso       = 0
      t_totconsulta     = 0
      t_totinc          = 0
      t_totalt          = 0
      t_totexc          = 0
      While Not RS.EOF
        Select Case p_agrega
           Case "SECRETARIA"
              If w_nm_quebra <> RS("nm_secretaria") Then
                 If w_qt_quebra > 0 Then
                    ImprimeLinha t_acesso, t_consulta, t_inc, t_alt, t_exc, w_chave
                    w_linha = w_linha + 2
                 End If
                 If O <> "W" or (w_ea = "W" and w_linha <= 25) Then
                    ' Se for geração de MS-Word, coloca a nova quebra somente se não estourou o limite
                    ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top""><td><font size=1><b>" & RS("nm_secretaria")
                 End If
                 w_nm_quebra       = RS("nm_secretaria")
                 w_chave           = RS("sq_regional")
                 w_qt_quebra       = 0
                 t_acesso          = 0
                 t_consulta        = 0
                 t_inc             = 0
                 t_alt             = 0
                 t_exc             = 0
                 w_linha           = w_linha + 1
              End If
           Case "REGIONAL"
              If w_nm_quebra <> RS("nm_regional") Then
                 If w_qt_quebra > 0 Then
                    ImprimeLinha t_acesso, t_consulta, t_inc, t_alt, t_exc, w_chave
                    w_linha = w_linha + 2
                 End If
                 If O <> "W" or (w_ea = "W" and w_linha <= 25) Then
                    ' Se for geração de MS-Word, coloca a nova quebra somente se não estourou o limite
                    ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top""><td><font size=1><b>" & RS("nm_regional")
                 End If
                 w_nm_quebra       = RS("nm_regional")
                 w_chave           = RS("sq_regional")
                 w_qt_quebra       = 0
                 t_acesso          = 0
                 t_consulta        = 0
                 t_inc             = 0
                 t_alt             = 0
                 t_exc             = 0
                 w_linha           = w_linha + 1
              End If
           Case "ESCOLA"
              If w_nm_quebra <> RS("nm_escola") Then
                 If w_qt_quebra > 0 Then
                    ImprimeLinha t_acesso, t_consulta, t_inc, t_alt, t_exc, w_chave
                    w_linha = w_linha + 2
                 End If
                 If O <> "W" or (w_ea = "W" and w_linha <= 25) Then
                    ' Se for geração de MS-Word, coloca a nova quebra somente se não estourou o limite
                    ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top""><td><font size=1><b>" & RS("nm_escola")
                 End If
                 w_nm_quebra       = RS("nm_escola")
                 w_chave           = RS("sq_escola")
                 w_qt_quebra       = 0
                 t_acesso          = 0
                 t_consulta        = 0
                 t_inc             = 0
                 t_alt             = 0
                 t_exc             = 0
                 w_linha           = w_linha + 1
              End If
           Case "BLOCODADOS"
              If w_nm_quebra <> RS("nm_funcionalidade") Then
                 If w_qt_quebra > 0 Then
                    ImprimeLinha t_acesso, t_consulta, t_inc, t_alt, t_exc, w_chave
                    w_linha = w_linha + 2
                 End If
                 If O <> "W" or (w_ea = "W" and w_linha <= 25) Then
                    ' Se for geração de MS-Word, coloca a nova quebra somente se não estourou o limite
                    ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top""><td><font size=1><b>" & RS("nm_funcionalidade")
                 End If
                 w_nm_quebra       = RS("nm_funcionalidade")
                 w_chave           = RS("cd_funcionalidade")
                 w_qt_quebra       = 0
                 t_acesso          = 0
                 t_consulta        = 0
                 t_inc             = 0
                 t_alt             = 0
                 t_exc             = 0
                 w_linha           = w_linha + 1
              End If
        End Select
        If w_ea = "W" and w_linha > 25 Then ' Se for geração de MS-Word, quebra a página
           ShowHTML "    </table>"
           ShowHTML "  </td>"
           ShowHTML "</tr>"
           ShowHTML "</table>"
           ShowHTML "</center></div>"
           ShowHTML "    <br style=""page-break-after:always"">"
           w_linha = 0
           w_pag   = w_pag + 1
           CabecalhoWord w_cliente, w_TP, w_pag
           If w_filtro > "" Then ShowHTML w_filtro End If
           ShowHTML "<div align=center><center>"
           ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
           ImprimeCabecalho
           Select Case p_agrega
              Case "SECRETARIA" ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top""><td><font size=1><b>" & RS("nm_secretaria")
              Case "REGIONAL"   ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top""><td><font size=1><b>" & RS("nm_regional")
              Case "ESCOLA"     ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top""><td><font size=1><b>" & RS("nm_escola")
              Case "BLOCODADOS" ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top""><td><font size=1><b>" & RS("nm_funcionalidade")
           End Select
           w_linha = w_linha + 1
        End If
        If  RS("tipo") = "0" Then
           t_consulta    = t_consulta + 1
           t_totconsulta = t_totconsulta + 1
        ElseIf  RS("tipo") = "1" Then
           t_inc    = t_inc + 1
           t_totinc = t_totinc + 1
        ElseIf  RS("tipo") = "2" Then
           t_alt    = t_alt + 1
           t_totalt = t_totalt + 1
        ElseIf RS("tipo") = "3" Then
           t_exc    = t_exc + 1
           t_totexc = t_totexc + 1
        End If
        t_acesso    = t_acesso + 1
        t_totacesso = t_totacesso + 1
        w_qt_quebra = w_qt_quebra + 1
        RS.MoveNext
      wend
      ImprimeLinha t_acesso, t_consulta, t_inc, t_alt, t_exc, w_chave
      ShowHTML "      <tr bgcolor=""#DCDCDC"" valign=""top"" align=""right"">"
      ShowHTML "          <td><font size=""1""><b>Totais</font></td>"
      ImprimeLinha t_totacesso, t_totconsulta, t_totinc, t_totalt, t_totexc, -1
    End If
    ShowHTML "      </center>"
    ShowHTML "    </table>"
    ShowHTML "  </td>"
    ShowHTML "</tr>"
  ElseIf Instr("P",w_ea) > 0 Then
    ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td><div align=""justify""><font size=2>Informe nos campos abaixo os valores que deseja filtrar e clique sobre o botão <i>Aplicar filtro</i>. Clicando sobre o botão <i>Remover filtro</i>, o filtro existente será apagado.</div><hr>"
    ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">" 
    ShowHTML "      <tr bgcolor=""" & conTrBgColor & """><td align=""center"" valign=""top""><table border=0 width=""90%"" cellspacing=0>"
    AbreForm "Form", w_Pagina & w_ew, "POST", "return(Validacao(this));", null
    ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""L"">"
    ShowHTML "<INPUT type=""hidden"" name=""w_troca"" value="""">"
    ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
    
    ' Exibe parâmetros de apresentação
    ShowHTML "         <tr><td colspan=""2"" align=""center"" bgcolor=""#D0D0D0"" style=""border: 2px solid rgb(0,0,0);""><font size=""1""><b>Parâmetros de Apresentação</td>"
    ShowHTML "         <tr valign=""top""><td colspan=2><table border=0 width=""100%"" cellpadding=0 cellspacing=0><tr valign=""top"">"
    ShowHTML "          <td><font size=""1""><b><U>A</U>gregar por:<br><SELECT ACCESSKEY=""O"" " & w_Disabled & " class=""STS"" name=""p_agrega"" size=""1"">"
    
    Select case p_agrega
       Case "SECRETARIA"  ShowHTML " <option value=""SECRETARIA"" selected>Secretaria <option value=""REGIONAL"">Regional          <option value=""ESCOLA"">Escola          <option value=""BLOCODADOS"">Bloco de Dados"
       Case "REGIONAL"    ShowHTML " <option value=""SECRETARIA"">Secretaria          <option value=""REGIONAL"" selected>Regional <option value=""ESCOLA"">Escola          <option value=""BLOCODADOS"">Bloco de Dados"
       Case "ESCOLA"      ShowHTML " <option value=""SECRETARIA"">Secretaria          <option value=""REGIONAL"">Regional          <option value=""ESCOLA"" selected>Escola <option value=""BLOCODADOS"">Bloco de Dados"
       Case "BLOCODADOS"  ShowHTML " <option value=""SECRETARIA"">Secretaria          <option value=""REGIONAL"">Regional          <option value=""ESCOLA"">Escola          <option value=""BLOCODADOS"" selected>Bloco de Dados"
       Case Else          ShowHTML " <option value=""SECRETARIA"" selected>Secretaria <option value=""REGIONAL"">Regional          <option value=""ESCOLA"">Escola          <option value=""BLOCODADOS"">Bloco de Dados"
    End Select
    
    ShowHTML "          </select></td>"
    ShowHTML "           </table>"
    ShowHTML "         </tr>"
    ShowHTML "         <tr><td valign=""top"" colspan=""2"" align=""center"" bgcolor=""#D0D0D0"" style=""border: 2px solid rgb(0,0,0);""><font size=""1""><b>Critérios de Busca</td>"
    ShowHTML "      <tr><td colspan=2><table border=0 width=""90%"" cellspacing=0><tr valign=""top"">"
    SelecaoRegional "<u>S</u>ubordinação:", "S", "Se desejar, indique um item para repuperar apenas as escolas subordinadas a ele.", p_regional, null, "p_regional", null, "onChange=""document.Form.w_ea.value='P'; document.Form.p_escola.value=''; document.Form.w_troca.value='p_escola'; document.Form.submit();"""
    SelecaoEscola   "<u>E</u>scola:"    , "E", "Selecione a Escola."            , p_escola, p_regional, "p_escola", null, null
    ShowHTML "         <tr valign=""top"">"
    SelecaoBlocoDados "<u>B</u>loco de Dados:", "B", "Selecione o bloco de dados.", p_bdados, null, "p_bdados", null, null
    SelecaoOperacao   "<u>O</u>peração:", "O", "Selecione a operação.", p_operacao, null, "p_operacao", null, null
    ShowHTML "          </table>"
    ShowHTML "      <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>P</u>eríodo:</b><br><input " & w_Disabled & " accesskey=""P"" type=""text"" name=""p_inicio"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(Nvl(p_inicio,Date()-30)) & """ onKeyDown=""FormataData(this,event);""> e <input " & w_Disabled & " accesskey=""C"" type=""text"" name=""p_fim"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(Nvl(p_fim,Date())) & """ onKeyDown=""FormataData(this,event);""></td>"
    ShowHTML "      <tr>"
    ShowHTML "      <tr><td align=""center"" colspan=""2"" height=""1"" bgcolor=""#000000"">"
    ShowHTML "      <tr><td align=""center"" colspan=""2"">"
    ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Exibir"" onClick=""javascript:document.Form.w_ea.value='L';"">"
    ShowHTML "          </td>"
    ShowHTML "      </tr>"
    ShowHTML "    </table>"
    ShowHTML "    </TD>"
    ShowHTML "</tr>"
    ShowHTML "</FORM>"
    ShowHTML "</table>"
  Else
    ScriptOpen "JavaScript"
    ShowHTML " alert('Opção não disponível');"
    ShowHTML " history.back(1);"
    ScriptClose
  End If
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape


End Sub
REM =========================================================================
REM Fim da rotina
REM -------------------------------------------------------------------------


REM =========================================================================
REM Rotina de impressao do cabecalho
REM -------------------------------------------------------------------------
Sub ImprimeCabecalho
    ShowHTML "<tr><td align=""center"">"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""#DCDCDC"" align=""center"">"
    Select case p_agrega
       Case "SECRETARIA" ShowHTML " <td><font size=""1""><b>Secretaria</font></td>"
       Case "REGIONAL"   ShowHTML " <td><font size=""1""><b>Regional de Ensino</font></td>"
       Case "ESCOLA"     ShowHTML " <td><font size=""1""><b>Escola</font></td>"
       Case "BLOCODADOS" ShowHTML " <td><font size=""1""><b>Bloco de Dados</font></td>"
    End Select
    ShowHTML "          <td><font size=""1""><b>Total</font></td>"
    ShowHTML "          <td><font size=""1""><b>Consulta</font></td>"
    ShowHTML "          <td><font size=""1""><b>Inclusões</font></td>"
    ShowHTML "          <td><font size=""1""><b>Alterações</font></td>"
    ShowHTML "          <td><font size=""1""><b>Exclusões</font></td>"
    ShowHTML "        </tr>"
End Sub
REM =========================================================================
REM Fim da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Log do sistema
REM -------------------------------------------------------------------------
Sub Log
  Dim w_chave, w_dt_ocorrencia, w_ds_ocorrencia, w_ano, P3, P4, w_data, w_ds_cliente
  
  Dim w_troca, i, w_texto
  
  w_Chave   = Request("w_Chave")
  w_troca   = Request("w_troca")
  P3        = Request("P3")
  P4        = Request("P4")
  
  If w_ea = "L" Then

     w_filtro = ""
     If p_regional > ""  Then 
        SQL = "select * from escCliente where sq_cliente = " & p_regional
        ConectaBD SQL
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Regional <td><font size=1>[<b>" & RS("ds_cliente") & "</b>]"
        DesconectaBD
     End If
     If p_escola > ""  Then 
        SQL = "select * from escCliente where sq_cliente = " & p_escola
        ConectaBD SQL
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Escola <td><font size=1>[<b>" & RS("ds_cliente") & "</b>]"
        DesconectaBD
     End If
     If p_bdados > ""  Then 
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Bloco de dados <td><font size=1>[<b>" & ExibeBlocoDados(p_bdados) & "</b>]"
     End If
     If p_operacao > ""  Then 
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Operação <td><font size=1>[<b>" & ExibeOperacao(p_operacao) & "</b>]"
     End If
     If p_inicio > ""  Then 
        w_filtro = w_filtro & "<tr valign=""top""><td align=""right""><font size=1>Período <td><font size=1>[<b>" & p_inicio & " a " & p_fim & "</b>]"
     End If
     If w_filtro > "" Then w_filtro = "<table border=0><tr valign=""top""><td><font size=1><b>Filtro:</b><td nowrap><font size=1><ul>" & w_filtro & "</ul></tr></table>" End If

     ' Recupera todos os registros para a listagem
     SQL = "select a.*, " & VbCrLf & _
           "       case a.tipo " & VbCrLf & _
           "            when 0 then 'Consulta' " & VbCrLf & _
           "            when 1 then 'Inclusão' " & VbCrLf & _
           "            when 2 then 'Alteração' " & VbCrLf & _
           "            when 3 then 'Exclusão' " & VbCrLf & _
           "            else        'Consulta' " & VbCrLf & _
           "       end nm_tipo, " & VbCrLf & _
           "       IsNull(b.nome,'Consulta') nm_funcionalidade,   IsNull(b.codigo,0) cd_funcionalidade, " & VbCrLf & _
           "       c.sq_cliente sq_escola,     c.ds_cliente nm_escola, " & VbCrLf & _
           "       d.sq_cliente sq_regional,   d.ds_cliente nm_regional, " & VbCrLf & _
           "       e.sq_cliente sq_secretaria, e.ds_cliente nm_secretaria " & VbCrLf & _
           "  from escCliente_Log                    a " & VbCrLf & _
           "       left outer join escFuncionalidade b on (a.sq_funcionalidade = b.sq_funcionalidade and " & VbCrLf & _
           "                                               b.tipo              = 2 " & VbCrLf & _
           "                                              ) " & VbCrLf & _
           "       inner      join escCliente        c on (a.sq_cliente        = c.sq_cliente) " & VbCrLf & _
           "         inner   join escCliente         d on (c.sq_cliente_pai    = d.sq_cliente) " & VbCrLf & _
           "           inner join escCliente         e on (d.sq_cliente_pai    = e.sq_cliente), " & VbCrLf & _
           "       escCliente                        f " & VbCrLf & _
           "         inner   join escTipo_Cliente    g on (f.sq_tipo_cliente   = g.sq_tipo_cliente) " & VbCrLf & _
           " where f.sq_cliente = " & replace(cl,"sq_cliente=","") & " " & VbCrLf & _
           "   and ((g.tipo = 2 and d.sq_cliente = " & replace(cl,"sq_cliente=","") & ") or " & VbCrLf & _
           "        (g.tipo <> 2) " & VbCrLf & _
           "       ) " & VbCrLf
     If p_regional > "" Then SQL = SQL & "   and d.sq_cliente = " & p_regional & VbCrLf End If
     If p_escola > ""   Then SQL = SQL & "   and c.sq_cliente = " & p_escola   & VbCrLf End If
     If cInt(Nvl(p_bdados,"-1")) = 0 Then
        SQL = SQL & "   and b.codigo     is null " & VbCrLf 
     ElseIf Nvl(p_bdados,"NULO") <> "NULO" Then 
        SQL = SQL & "   and b.codigo     = " & p_bdados   & VbCrLf 
     End If
     If p_operacao > "" Then SQL = SQL & "   and a.tipo       = " & p_operacao & VbCrLf End If
     If p_inicio > ""   Then SQL = SQL & "   and a.data       between convert(datetime, '" & p_inicio & "', 103) and convert(datetime, '" & p_fim & "', 103) + 1 " & VbCrLf End If
     SQL = SQL & "order by year(data) desc, month(data) desc, data desc"
     ConectaBD SQL
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  ShowHTML "<TITLE>Registro de ocorrências no site da escola</TITLE>"
  If InStr("IAEP",O) > 0 Then
     ScriptOpen "JavaScript"
     CheckBranco
     FormataData
     ValidateOpen "Validacao"
     If InStr("IA",O) > 0 Then
        Validate "w_dt_ocorrencia", "dt_ocorrencia", "DATA", "1", "10", "10", "1", "1"
        Validate "w_ds_ocorrencia", "ds_ocorrencia", "", "1", "2", "60", "1", "1"
     End If
     ShowHTML "  theForm.Botao[0].disabled=true;"
     ShowHTML "  theForm.Botao[1].disabled=true;"
     ValidateClose
     ScriptClose
  End If
  ShowHTML "</HEAD>"
  If w_troca > "" Then
     BodyOpen "onLoad='document.Form." & w_troca & ".focus()';"
  ElseIf w_ea = "I" or w_ea = "A" Then
     BodyOpen "onLoad='document.Form.w_dt_ocorrencia.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Registro de ocorrências</FONT></B>"
  ShowHTML "<HR>"
  If w_filtro > "" Then ShowHTML w_filtro End If
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  If w_ea = "L" Then
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML "<tr><td><font size=""2""><b>" & w_ds_cliente
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Data</font></td>"
    ShowHTML "          <td><font size=""1""><b>Hora</font></td>"
    ShowHTML "          <td><font size=""1""><b>Origem</font></td>"
    ShowHTML "          <td><font size=""1""><b>Escola</font></td>"
    ShowHTML "          <td><font size=""1""><b>Ocorrência</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=6 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      
      rs.PageSize     = P4
      rs.AbsolutePage = P3
      w_ano  = ""
      While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(P3)
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        If wAno <> year(RS("data")) & "/" & month(RS("data")) Then
           ShowHTML "      <tr bgcolor=""#C0C0C0"" valign=""top""><TD colspan=6 align=""center""><font size=2><B>" & month(RS("data")) & "/" & year(RS("data")) & "</b></font></td></tr>"
           wAno = year(RS("data")) & "/" & month(RS("data"))
        End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        If w_data <> FormataDataEdicao(FormatDateTime(RS("data"),2)) Then
           ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("data"),2)) & "</td>"
           w_data = FormataDataEdicao(FormatDateTime(RS("data"),2))
        Else
           ShowHTML "        <td align=""center""></td>"
        End If
        ShowHTML "        <td align=""center""><font size=""1"">" & FormatDateTime(RS("data"),3) & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("ip_origem") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("nm_escola") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("abrangencia") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_pagina & w_ew & "&R=" & w_Pagina & "Log&w_ea=V&CL=" & CL & "&w_chave=" & RS("sq_cliente_log") & """>Detalhes</A>&nbsp"
        ShowHTML "        </td>"
        ShowHTML "      </tr>"
        RS.MoveNext
      wend
    End If
    ShowHTML "      </center>"
    ShowHTML "    </table>"

    ShowHTML "<tr><td align=""center"" colspan=2>"
    MontaBarra w_pagina & w_ew & "&w_ea=L", cDbl(RS.PageCount), cDbl(P3), cDbl(P4), cDbl(RS.RecordCount)
    ShowHTML "</tr>"

    ShowHTML "  </td>"
    ShowHTML "</tr>"
    DesconectaBD
  ElseIf Instr("IAEV",O) > 0 Then
    ' Recupera os dados da ocorrência selecionada
     SQL = "select a.*, " & VbCrLf & _
           "       case a.tipo " & VbCrLf & _
           "            when 0 then 'Consulta' " & VbCrLf & _
           "            when 1 then 'Inclusão' " & VbCrLf & _
           "            when 2 then 'Alteração' " & VbCrLf & _
           "            when 3 then 'Exclusão' " & VbCrLf & _
           "            else        'Erro' " & VbCrLf & _
           "       end nm_tipo, " & VbCrLf & _
           "       IsNull(b.nome,'Consulta') nm_funcionalidade,   IsNull(b.codigo,0) cd_funcionalidade, " & VbCrLf & _
           "       c.sq_cliente sq_escola,     c.ds_cliente nm_escola, " & VbCrLf & _
           "       d.sq_cliente sq_regional,   d.ds_cliente nm_regional, " & VbCrLf & _
           "       e.sq_cliente sq_secretaria, e.ds_cliente nm_secretaria " & VbCrLf & _
           "  from escCliente_Log                    a " & VbCrLf & _
           "       left outer join escFuncionalidade b on (a.sq_funcionalidade = b.sq_funcionalidade and " & VbCrLf & _
           "                                               b.tipo              = 2 " & VbCrLf & _
           "                                              ) " & VbCrLf & _
           "       inner      join escCliente        c on (a.sq_cliente        = c.sq_cliente and " & VbCrLf & _
           "                                               c.sq_tipo_cliente   = 6 " & VbCrLf & _
           "                                              ) " & VbCrLf & _
           "         inner   join escCliente         d on (c.sq_cliente_pai    = d.sq_cliente) " & VbCrLf & _
           "           inner join escCliente         e on (d.sq_cliente_pai    = e.sq_cliente), " & VbCrLf & _
           "       escCliente                        f " & VbCrLf & _
           "         inner   join escTipo_Cliente    g on (f.sq_tipo_cliente   = g.sq_tipo_cliente) " & VbCrLf & _
           " where sq_cliente_log = " & w_chave
    ConectaBD SQL

    ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
    ShowHTML "    <table width=""95%"" border=""0"">"
    ShowHTML "        <tr><td colspan=2><font size=""2""><b>" & RS("nm_escola")
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td><font size=""1"">Data:<br><b>" & FormatDateTime(RS("data"),1) & ", " & FormatDateTime(RS("data"),3) & "</b></td>"
    ShowHTML "          <td><font size=""1"">IP de origem:<br><b>" & RS("ip_origem") & "</b></td>"
    ShowHTML "        <tr><td colspan=2><font size=""1"">&nbsp;"
    ShowHTML "        <tr><td colspan=2><font size=""1"">Tipo da ocorrencia: <b>"
    Select Case cDbl(RS("tipo"))
      Case 0 ShowHTML "            Acesso"
      Case 1 ShowHTML "            Inclusão"
      Case 2 ShowHTML "            Alteração"
      Case 3 ShowHTML "            Exclusão"
    End Select
    ShowHTML "            </td>"
    ShowHTML "        <tr><td colspan=2><font size=""1"">&nbsp;"
    ShowHTML "        <tr><td colspan=2><font size=""1"">Descrição:<br><b>" & RS("abrangencia") & "</b></td>"
    If RS("sql") > "" Then
       ShowHTML "        <tr><td colspan=2><font size=""1"">&nbsp;"
       ShowHTML "        <tr><td colspan=2><font size=""1"">Comando(s) executado(s):<br><b>" & replace(replace(RS("sql")," ","&nbsp;"),VbCrLf,"<br>") & "</b></td>"
    End If
    ShowHTML "      <tr><td colspan=2 height=1 bgcolor=""#000000"">"
    ShowHTML "      <tr><td colspan=2 align=""center""><input class=""STB"" type=""button"" onClick=""history.back(1);"" name=""Botao"" value=""Voltar""></td></tr>"
    ShowHTML "    </table>"
    ShowHTML "    </TD>"
    ShowHTML "</tr>"
    DesconectaBD
  Else
    ScriptOpen "JavaScript"
    ShowHTML " alert('Opção não disponível');"
    ShowHTML " history.back(1);"
    ScriptClose
  End If
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape

  Set w_ano             = Nothing 
  Set w_ds_ocorrencia   = Nothing 
  Set w_dt_ocorrencia   = Nothing 
  Set w_chave           = Nothing 
  Set w_troca           = Nothing 
  Set i                 = Nothing 
  Set w_texto           = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de calendário
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina de impressao da linha resumo
REM -------------------------------------------------------------------------
Sub ImprimeLinha (p_acesso, p_consulta, p_inc, p_alt, p_exc, p_chave)
    If w_ea = "L"                    Then ShowHTML "          <td align=""right""><font size=""1""><a class=""hl"" href=""javascript:lista('" & p_chave & "', -1);"" onMouseOver=""window.status='Exibe os registros de log.'; return true"" onMouseOut=""window.status=''; return true"">" & FormatNumber(p_acesso,0) & "</a>&nbsp;</font></td>"              Else ShowHTML "          <td align=""right""><font size=""1"">" & FormatNumber(p_acesso,0)   & "&nbsp;</font></td>" End If
    If p_consulta > 0 and w_ea = "L" Then ShowHTML "          <td align=""right""><a class=""hl"" href=""javascript:lista('" & p_chave & "', 0);"" onMouseOver=""window.status='Exibe os registros de log.'; return true"" onMouseOut=""window.status=''; return true""><font size=""1"">" & FormatNumber(p_consulta,0) & "</a>&nbsp;</font></td>"             Else ShowHTML "          <td align=""right""><font size=""1"">" & FormatNumber(p_consulta,0) & "&nbsp;</font></td>" End If
    If p_inc > 0      and w_ea = "L" Then ShowHTML "          <td align=""right""><a class=""hl"" href=""javascript:lista('" & p_chave & "', 1);"" onMouseOver=""window.status='Exibe os registros de log.'; return true"" onMouseOut=""window.status=''; return true""><font size=""1"">" & FormatNumber(p_inc,0) & "</a>&nbsp;</font></td>"                  Else ShowHTML "          <td align=""right""><font size=""1"">" & FormatNumber(p_inc,0)      & "&nbsp;</font></td>" End If
    If p_alt > 0      and w_ea = "L" Then ShowHTML "          <td align=""right""><a class=""hl"" href=""javascript:lista('" & p_chave & "', 2);"" onMouseOver=""window.status='Exibe os registros de log.'; return true"" onMouseOut=""window.status=''; return true""><font size=""1"">" & FormatNumber(p_alt,0) & "</a>&nbsp;</font></td>"                  Else ShowHTML "          <td align=""right""><font size=""1"">" & FormatNumber(p_alt,0)      & "&nbsp;</font></td>" End If
    If p_exc > 0      and w_ea = "L" Then ShowHTML "          <td align=""right""><a class=""hl"" href=""javascript:lista('" & p_chave & "', 3);"" onMouseOver=""window.status='Exibe os registros de log.'; return true"" onMouseOut=""window.status=''; return true""><font size=""1"">" & FormatNumber(p_exc,0) & "</a>&nbsp;</font></td>"                  Else ShowHTML "          <td align=""right""><font size=""1"">" & FormatNumber(p_exc,0)      & "&nbsp;</font></td>" End If
    ShowHTML "        </tr>"
End Sub
REM =========================================================================
REM Fim da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody
  
  Select Case uCase(w_EW)
     Case "GERENCIAL"   Gerencial
     Case "LOG"         Log
  End Select
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>

