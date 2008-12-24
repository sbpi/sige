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
REM Descricao: Liga a regional de ensino à sua respectiva escola
REM Home     : http://www.sbpi.com.br/
REM Mail     : beto@sbpi.com.br
REM Criacao  : 16/10/2004 15:20
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM -------------------------------------------------------------------------
REM
REM Parâmetros recebidos:
REM    R (referência) = usado na rotina de gravação, com conteúdo igual ao parâmetro T
REM w_ea (operação)   = P   : Filtragem
REM                   = A   : Alteração
REM                   = W   : Geração de documento no formato MS-Word (Office 2003)

  Private w_EA
  Private w_IN
  Private w_EF, w_troca
  Private SQL, CL, DBMS, RS, p_agrega
  Private w_Data, w_pagina
  Dim w_tp, w_Disabled
  Dim p_local, p_regiao, p_regional, p_escola, p_operacao, p_secretaria, p_tipo_escola
  
  p_regiao      = Request("p_regiao")
  p_regional    = Request("p_regional")
  p_escola      = Request("p_escola")
  p_local       = Request("p_local")
  p_tipo_escola = Request("p_tipo_escola")
  p_operacao    = Request("p_operacao")
  w_troca       = Request("w_troca")
  
  w_EW       = ucase(Request("w_ew"))
  w_IN       = ucase(Request("w_in"))
  w_EF       = ucase(Request("w_ef"))
  CL         = Request("CL")
  w_troca    = Request("w_troca")
  w_Disabled = "ENABLED"
  P4         = cDbl(Nvl(Request("P4"),30))
  P3         = cDbl(Nvl(Request("P3"),1))

  w_Data = Mid(100+Day(Date()),2,2) & "/" & Mid(100+Month(Date()),2,2) & "/" & year(Date())
  w_Pagina = ExtractFileName(Request.ServerVariables("SCRIPT_NAME")) & "?w_ew="
  
  AbreSessaoManut
    
  If w_ea = "A" Then w_ea = "L" End If

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
REM Procedimento que executa operações de BD
REM -------------------------------------------------------------------------
Public Sub Grava

  Dim w_chave, w_sql, w_funcionalidade, w_diretorio, w_imagem, w_arquivo
  
  if p_regiao <> "" then
    Cabecalho
    ShowHTML "</HEAD>"
    BodyOpen "onLoad=document.focus();"
  
    ' Recupera o código a ser alterado na tabela de clientes

    dbms.BeginTrans()
    SQL = "update sbpi.Cliente set " & VbCrLf & _
          "     sq_regiao_adm   = " & Request("p_regiao") & ", " & VbCrLf & _
          "     sq_tipo_cliente = " & Request("p_tipo_escola") & ", " & VbCrLf & _
          "     localizacao     = " & Request("p_local") & VbCrLf
    If nvl(p_regional,"") <> "" Then SQL = SQL & "     , sq_cliente_pai  = " & Request("p_regional") & VbCrLf Else SQL = SQL & "     , sq_cliente_pai  = null " & VbCrLf 
    SQL = SQL & _
          "   where sq_cliente  = " & Request("p_escola") & VbCrLf
    ExecutaSQL(SQL)
    dbms.CommitTrans()
            
    ScriptOpen "JavaScript"
    ShowHTML "  location.href='" & w_pagina & "Ligacao&w_ea=L&cl=" & cl & "&p_escola=" & p_escola & "&p_regional=" & p_regional & "&p_tipo_escola=" & p_tipo_escola &"';"
    ScriptClose
  Else
    ScriptOpen "JavaScript"
    ShowHTML "  alert('Bloco de dados não encontrado: " & SG & "');"
    ShowHTML "  history.back(1);"
    ScriptClose
  End If
  
  
  Set w_chave          = Nothing
  Set w_funcionalidade = Nothing
End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que executa operações de BD
REM =========================================================================

REM =========================================================================
REM Pesquisa gerencial
REM -------------------------------------------------------------------------
Sub LigaRegEsc
  
  If nvl(p_escola,"") <> "" Then
     SQL = "select * from sbpi.Cliente where sq_cliente = " & p_escola
     ConectaBD SQL
     p_regiao      = RS("sq_regiao_adm")
     p_regional    = RS("sq_cliente_pai")
     p_tipo_escola = RS("sq_tipo_cliente")
     p_local       = RS("localizacao")
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "JavaScript"
  ValidateOpen "Validacao"
  Validate "p_escola",      "Unidade de ensino",     "SELECT", "1", "1", "18", "", "1"
  Validate "p_regiao",      "Região administrativa", "SELECT", "1", "1", "18", "", "1"
  Validate "p_regional",    "Regional de ensino",    "SELECT", "1", "1", "18", "", "1"
  Validate "p_tipo_escola", "Tipo de unidade",       "SELECT", "1", "1", "18", "", "1"
  ShowHTML "  theForm.Botao.disabled=true;"
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  If w_Troca > "" Then ' Se for recarga da página
     BodyOpen "onLoad='document.Form." & w_Troca & ".focus();'"
  Else
     BodyOpen "onLoad=document.Form.p_escola.focus();"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
  ShowHTML "<B><FONT size=""2"" COLOR=""#000000"">Vinculação e tipologia</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">" 
  ShowHTML "      <tr bgcolor=""" & conTrBgColor & """><td align=""center"" valign=""top""><table border=0 width=""90%"" cellspacing=0>"
  AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
  ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"
  ShowHTML "<INPUT type=""hidden"" name=""w_troca"" value="""">"
  ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
  ShowHTML "      <tr><td colspan=2><table border=0 width=""90%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  SelecaoEscola         "Unidad<u>e</u> de ensino:", "E", "Selecione unidade." , p_escola, null, "p_escola", null, "onChange=""document.Form.action='" & w_pagina & w_ew & "'; document.Form.w_ea.value='P'; document.Form.p_regiao.value=''; document.Form.p_regional.value=''; document.Form.w_troca.value='p_regiao'; document.Form.submit();"""
  ShowHTML "        <tr valign=""top"">"
  SelecaoRegiaoAdm "Região a<u>d</u>ministrativa:", "D", "Indique a região administrativa.", p_regiao, p_escola, "p_regiao", null, null
  ShowHTML "        <tr valign=""top"">"
  SelecaoRegionalEscola "Regional de en<u>s</u>ino:", "S", "Indique a regional de ensino.", p_regional, p_escola, "p_regional", null, null
  ShowHTML "        <tr valign=""top"">"
  SelecaoTipoEscola     "<u>T</u>ipo de Escola:", "T", "Selecione o tipo da Escola.", p_tipo_escola, p_escola, "p_tipo_escola", null, null
  ShowHTML "         <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b>Localização</b><br>"
  If p_local = "1" Then
     ShowHTML "              <input  type=""radio"" name=""p_local"" value=""0""> Não informada <input  type=""radio"" name=""p_local"" value=""1"" checked> Urbana <input  type=""radio"" name=""p_local"" value=""2""> Rural "
  ElseIf p_local = "2" Then
     ShowHTML "              <input  type=""radio"" name=""p_local"" value=""0""> Não informada <input  type=""radio"" name=""p_local"" value=""1""> Urbana <input  type=""radio"" name=""p_local"" value=""2"" checked> Rural "
  Else
     ShowHTML "              <input  type=""radio"" name=""p_local"" value=""0"" checked> Não informada <input  type=""radio"" name=""p_local"" value=""1""> Urbana <input  type=""radio"" name=""p_local"" value=""2""> Rural "
  End If
  ShowHTML "          </table>"
  ShowHTML "      <tr valign=""top"">"
  ShowHTML "      <tr>"
  ShowHTML "      <tr><td align=""center"" colspan=""2"" height=""1"" bgcolor=""#000000"">"
  ShowHTML "      <tr><td align=""center"" colspan=""2"">"
  ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Gravar"">"
  ShowHTML "          </td>"
  ShowHTML "      </tr>"
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  ShowHTML "</FORM>"
  ShowHTML "</table>"
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape
End Sub
REM =========================================================================
REM Fim da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody
  
  Select Case uCase(w_EW)
     Case "LIGACAO"
        ligaRegEsc
      Case "GRAVA"
        Grava
  End Select
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>

