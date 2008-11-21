<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="Constants_ADO.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="Funcoes.asp"-->
<!--#INCLUDE FILE="JScript.asp"-->
<%
REM =========================================================================
REM  /newsletter.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: Página do portal 
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 13/04/2000 23:10
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by SBPI - Socriedade Brasileira para a Pesquisa em Informática
REM -------------------------------------------------------------------------

Private RS
Private RS1
Private sstrSN
Private sstrCL
Private strTitulo
Private marcado, i, dbms

sstrSN = "newsletter.asp"

Public w_EF
Public w_EW
Public w_IN
Public sstrDiretorio
Public sstrModelo
  
w_EF = "sq_cliente=" & Session("CL")
w_EW = uCase(Request("EW"))
w_IN = uCase(Request("IN"))

Set RS = Server.CreateObject("ADODB.RecordSet")
Set RS1 = Server.CreateObject("ADODB.RecordSet")

Server.ScriptTimeOut = Session("ScriptTimeOut")

ShowHTML "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
ShowHTML "<html xmlns=""http://www.w3.org/1999/xhtml"">"
ShowHTML "<head>"
ShowHTML "   <title>Secretaria de Estado de Educa&ccedil;&atilde;o</title>"
ShowHTML "   <meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> "
ShowHTML "   <link href=""css/estilo.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <link href=""css/print.css""  media=""print""  rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"

AbreSessao
Main
FechaSessao

ShowHTML "        </table>"
ShowHTML "    </div>"
ShowHTML "  </div> "

ShowHTML "  <br clear=""all"" />"
ShowHTML "</div>"

%><div id="rodape">
  <!-- Menu Rodape -->
  <ul>
    <li><a href="http://www.se.df.gov.br">Página Inicial</a></li>
    <li><a href="http://www.se.df.gov.br/300/30001005.asp">Fale conosco</a></li>
    <li class="ultimo"><a href="http://www.se.df.gov.br/300/30001009.asp">Mapa do site</a></li>
  </ul>
  <!-- Fim Menu Rodape -->
  <div class="clear"/>
  <!-- Endereço -->
  <div class="endereco">
  <div class="buriti">
  <p>Anexo do Palácio do Buriti - 9º andar  - Brasília </p>
  <p>Telefone (61) 3224 0016 (61) 3225 1266 | Fax (61) 3901 3171</p>
  </div>
  <div class="buritinga">
  <p>QNG AE Lote 22 bl 05 sala 03 - Taguatinga   Norte</p>
  <p>Telefone (61) 3355 86 30 | Fax 3355 86 94</p>
  </div>
  </div>
  
  <p class="copy">Copyright ® 2000/2008 - SE/GDF - Todos os Direitos Reservados</p>
  <!-- Fim Endereço -->
</div>
  <%
ShowHTML "</body>"
ShowHTML "</html>"

REM =========================================================================
REM Monta a tela de Cadastro
REM -------------------------------------------------------------------------
Public Sub GetCadastro

  Dim sql, sql2, wCont, sql1, wAtual, wIN
  
  ScriptOpen "JavaScript"
  ValidateOpen "Validacao"
  Validate "w_nome", "Nome", "", "1", "3", "60", "1", "1"
  Validate "w_email", "e-Mail", "", "1", "4", "60", "1", "1"
  ShowHTML "  if (theForm.w_tipo[0].checked==false && theForm.w_tipo[1].checked==false && theForm.w_tipo[2].checked==false) {"
  ShowHTML "     alert('Você deve selecionar uma das opções apresentadas no formulário!');"
  ShowHTML "     return false;"
  ShowHTML "  }"
  ShowHTML "  theForm.Botao[0].disabled=true;"
  ShowHTML "  theForm.Botao[1].disabled=true;"
  ValidateClose
  ScriptClose
  ShowHTML "</head>"
  ShowHTML "<body>"
  %>
<div id="pagina">
  <div id="topo"> </div>
  <div id="busca">
    <div class="data"> <% Response.Write(ExibeData(date())) %> </div>
    <div class="clear"></div>
  </div>
  <div id="menu">
    <div id="menuTop">
      <div class="esquerda">
        <ul>
          <li><a href="http://www.se.df.gov.br/300/30001001.asp">Secretaria de Educação</a></li>
          <li><a href="http://www.se.df.gov.br/300/30001009.asp">mapa do site</a></li>
          <li class="ultimo"><a class="ultimo" href="http://www.se.df.gov.br/300/30001005.asp">fale conosco</a></li>
        </ul>
      </div>
      <div class="direita">
        <ul>
          <li><a class="aluno" href="http://www.se.df.gov.br/300/30002001.asp">Aluno</a></li>
          <li><a class="educador" href="http://www.se.df.gov.br/300/30003001.asp">Educador</a></li>
          <li><a class="comunidade" href="http://www.se.df.gov.br/300/30004001.asp">Comunidade</a></li>
        </ul>
      </div>
    </div>
    <div id="menuMiddle"> </div>
  </div>
  <div class="clear"></div>
<div id="menuBottom">
  <div class="clear"></div>
</div>
<div id="conteudo"><h2>Solução Integrada de Gestão Educacional - SIGE</h2>
<%
  ShowHTML "    <div id=""texto""><!-- Conteúdo -->"
  ShowHTML "        <table width=""570"" border=""0"">"

  ShowHTML "        <FORM ACTION=""newsletter.asp"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"">"
  ShowHTML "        <input type=""hidden"" name=""EW"" value=""Grava"">"
  ShowHTML "        <input type=""hidden"" name=""IN"" value=""Cadastro"">"
  ShowHTML "          <TR><TD align=""left"" valign""middle"">Informe os dados solicitados no formulário abaixo e clique no botão ""Registrar"" para registrar-se em nossa lista de distribuição."
  ShowHTML "          <TR><TD>Informe seu nome completo:<br><input type=""text"" size=""60"" maxlength=""60"" name=""w_nome"" value="""" CLASS=""texto"">"
  ShowHTML "          <TR><TD>Informe o e-Mail que deseja receber a newsletter:<br><input type=""text"" size=""60"" maxlength=""60"" name=""w_email"" value="""" CLASS=""texto"">"
  ShowHTML "          <TR><TD>Selecione uma das opções abaixo: "
  ShowHTML "                <br><input type=""Radio"" name=""w_tipo"" value=""1""> Pai, mãe ou responsável por aluno da rede de ensino "
  ShowHTML "                <br><input type=""Radio"" name=""w_tipo"" value=""2""> Aluno da rede de ensino "
  ShowHTML "                <br><input type=""Radio"" name=""w_tipo"" value=""3""> Outro "
  ShowHTML "          <TR><TD align=""center""><font size=""2"" CLASS=""BTM"">"
  ShowHTML "               <input type=""submit"" name=""Botao"" value=""Registrar"" class=""botao"">"
  ShowHTML "               <input type=""button"" name=""Botao"" value=""Voltar"" class=""botao"" onClick=""javascript: history.back(1);"">"
  ShowHTML "        </FORM>" & VbCrLf
  ShowHTML "<br/><br/> Se desejar remover seu e-mail, clique <a class=""SS"" href=""http://www.gdfsige.df.gov.br/newsletter.asp?EW=Remover"" target=""_self"">aqui</a>."
 
End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Pesquisa
REM =========================================================================

REM =========================================================================
REM Monta a tela de Cadastro
REM -------------------------------------------------------------------------
Public Sub Remove

  Dim sql, sql2, wCont, sql1, wAtual, wIN
  
  ScriptOpen "JavaScript"
  ValidateOpen "Validacao"
  Validate "w_email", "e-Mail", "", "1", "4", "60", "1", "1"
  ShowHTML "  theForm.Botao[0].disabled=true;"
  ShowHTML "  theForm.Botao[1].disabled=true;"
  ValidateClose
  ScriptClose
  ShowHTML "</head>"
  ShowHTML "<body>"
    %>
<div id="pagina">
  <div id="topo"> </div>
  <div id="busca">
    <div class="data"> <% Response.Write(ExibeData(date())) %> </div>
    <div class="clear"></div>
  </div>
  <div id="menu">
    <div id="menuTop">
      <div class="esquerda">
        <ul>
          <li><a href="http://www.se.df.gov.br/300/30001001.asp">Secretaria de Educação</a></li>
          <li><a href="http://www.se.df.gov.br/300/30001009.asp">mapa do site</a></li>
          <li class="ultimo"><a class="ultimo" href="http://www.se.df.gov.br/300/30001005.asp">fale conosco</a></li>
        </ul>
      </div>
      <div class="direita">
        <ul>
          <li><a class="aluno" href="http://www.se.df.gov.br/300/30002001.asp">Aluno</a></li>
          <li><a class="educador" href="http://www.se.df.gov.br/300/30003001.asp">Educador</a></li>
          <li><a class="comunidade" href="http://www.se.df.gov.br/300/30004001.asp">Comunidade</a></li>
        </ul>
      </div>
    </div>
    <div id="menuMiddle"> </div>
  </div>
  <div class="clear"></div>
<div id="menuBottom">
  <div class="clear"></div>
</div>
<div id="conteudo"><h2>Solução Integrada de Gestão Educacional - SIGE</h2>
<%

  ShowHTML "        <table width=""570"" border=""0"">"
  ShowHTML "        <FORM ACTION=""newsletter.asp"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"">"
  ShowHTML "        <input type=""hidden"" name=""EW"" value=""Grava"">"
  ShowHTML "        <input type=""hidden"" name=""IN"" value=""Remove"">"
  ShowHTML "          <TR><TD>Informe seu e-mail no campo abaixo e clique no botão ""Remover"" para ser removido da nossa lista de distribuição."
  ShowHTML "          <TR><TD>Informe o e-Mail que deseja remover:<br><input type=""text"" size=""60"" maxlength=""60"" name=""w_email"" value="""" CLASS=""texto"">"
  ShowHTML "          <TR><TD align=""center"">"
  ShowHTML "               <input type=""submit"" name=""Botao"" value=""Remover"" class=""botao"">"
  ShowHTML "               <input type=""button"" name=""Botao"" value=""Voltar"" class=""botao"" onClick=""javascript: history.back(1);"">"
  ShowHTML "        </FORM>" & VbCrLf
End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Pesquisa
REM =========================================================================

REM =========================================================================
REM Rotina de preparação para envio de e-mail relativo a demandas eventuais
REM Finalidade: preparar os dados necessários ao envio automático de e-mail
REM Parâmetro: p_solic: número de identificação da solicitação. 
REM            p_tipo:  1 - Inclusão
REM                     2 - Remoção
REM -------------------------------------------------------------------------
Sub PreparaMail(p_nome, p_mail, p_tipo, p_evento)

  Dim w_html, w_destinatarios, w_assunto
  
  w_html = "<HTML>" & VbCrLf
  w_html = w_html & BodyOpenMail(null) & VbCrLf
  w_html = w_html & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & VbCrLf
  w_html = w_html & "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">" & VbCrLf
  w_html = w_html & "    <table width=""97%"" border=""0"">" & VbCrLf
  If p_evento = 1 Then
     w_html = w_html & "      <tr valign=""top""><td align=""center""><font size=2><b>INCLUSÃO NA LISTA DE DISTRIBUIÇÃO</b></font><br><br><td></tr>" & VbCrLf
     w_html = w_html & "      <tr align=""center""><td><font size=2><b><font color=""#BC3131"">ATENÇÃO</font>: Esta é uma mensagem de envio automático. Não responda.</b></font><br><br><td></tr>" & VbCrLf
     w_html = w_html & VbCrLf & "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
     w_html = w_html & VbCrLf & "      <tr><td><font size=2>Seu e-mail foi incluído na lista de distribuição de informativos da Secretaria de Estado de Educação do DF. A partir de agora, sempre que algum informativo for enviado para a lista, você será um dos destinatários.</b></font></td></tr>"
     w_html = w_html & "      <tr align=""center""><td><br><br><br><font size=1>" & VbCrLf
     w_html = w_html & "         Se desejar remover seu e-mail, clique <a class=""SS"" href=""http://www.gdfsige.df.gov.br/newsletter.asp?EW=Remover"" target=""_blank"">aqui</a>." & VbCrLf
     w_html = w_html & "      </font></td></tr>" & VbCrLf
  ElseIf p_evento = 2 Then
     w_html = w_html & "      <tr valign=""top""><td align=""center""><font size=2><b>SAÍDA DA LISTA DE DISTRIBUIÇÃO</b></font><br><br><td></tr>" & VbCrLf
     w_html = w_html & "      <tr align=""center""><td><font size=2><b><font color=""#BC3131"">ATENÇÃO</font>: Esta é uma mensagem de envio automático. Não responda.</b></font><br><br><td></tr>" & VbCrLf
     w_html = w_html & VbCrLf & "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
     w_html = w_html & VbCrLf & "      <tr><td><font size=2>Seu e-mail foi removido da lista de distribuição de informativos da Secretaria de Estado de Educação do DF.</b></font></td></tr>"
     w_html = w_html & "      <tr align=""center""><td><br><br><br><font size=1>" & VbCrLf
     w_html = w_html & "         Se desejar cadastrar novamente seu e-mail, clique <a class=""SS"" href=""http://www.gdfsige.df.gov.br/newsletter.asp?EW=i"" target=""_blank"">aqui</a>." & VbCrLf
  End IF
  w_html = w_html & "      <tr valign=""top""><td><br><br><br><font size=1>" & VbCrLf
  w_html = w_html & "         Dados da ocorrência:<br>" & VbCrLf
  w_html = w_html & "         <ul>" & VbCrLf
  w_html = w_html & "         <li>Data do servidor: <b>" & FormatDateTime(Date(),1) & ", " & Time() & "</b></li>" & VbCrLf
  w_html = w_html & "         <li>IP de origem: <b>" & Request.ServerVariables("REMOTE_HOST") & "</b></li>" & VbCrLf
  w_html = w_html & "         </ul>" & VbCrLf
  w_html = w_html & "      </font></td></tr>" & VbCrLf
  w_html = w_html & "    </table>" & VbCrLf
  w_html = w_html & "</td></tr>" & VbCrLf
  w_html = w_html & "</table>" & VbCrLf
  w_html = w_html & "</BODY>" & VbCrLf
  w_html = w_html & "</HTML>" & VbCrLf

  ' Prepara os dados necessários ao envio
  If p_evento = 1 Then ' Inclusão ou Conclusão
     w_assunto = "SEDF - Inclusão na lista de distribuição de informativos"
  ElseIf p_evento = 2 Then ' Tramitação
     w_assunto = "SEDF - Saída da lista de distribuição de informativos"
  End If

  ' Executa o envio do e-mail
  w_resultado = EnviaMail(w_assunto, w_html, p_mail, null)
        
  ' Se ocorreu algum erro, avisa da impossibilidade de envio
  If w_resultado > "" Then 
     ScriptOpen "JavaScript"
     ShowHTML "  alert('ATENÇÃO: não foi possível proceder o envio do e-mail.\n" & w_resultado & "');" 
     ScriptClose
  End If

  Set w_html                   = Nothing
  Set w_destinatarios          = Nothing
  Set w_assunto                = Nothing
End Sub
REM =========================================================================
REM Fim da rotina da preparação para envio de e-mail
REM -------------------------------------------------------------------------

REM =========================================================================
REM Procedimento que executa as operações de BD
REM -------------------------------------------------------------------------
Public Sub Grava

  Dim w_chave, w_sql, w_funcionalidade, w_diretorio, w_imagem, w_arquivo

  Cabecalho
  ShowHTML "</HEAD>"
  BodyOpen "onLoad=document.focus();"
  
  Select Case w_IN
    Case "CADASTRO"
       dbms.BeginTrans()
       
       SQL = "select envia_mail from escNewsletter where upper(email) = '" & uCase(Trim(Request("w_email"))) & "' " & VbCrLf
       ConectaBD SQL
       
       If RS.EOF Then
          SQL = "insert into escNewsletter (sq_cliente, nome, email, tipo, envia_mail, data_inclusao, data_alteracao) " & VbCrLf & _
                "  values ( 0," & VbCrLf & _
                "          '" & trim(Request("w_nome")) & "', " & VbCrLf & _
                "          '" & trim(Request("w_email")) & "', " & VbCrLf & _
                "          '" & Request("w_tipo") & "', " & VbCrLf & _
                "          'S', " & VbCrLf & _
                "          getdate(), " & VbCrLf & _
                "          null " & VbCrLf & _
                "         )" & VbCrLf
          ExecutaSQL(SQL)
          
          ' Envia e-mail comunicando a inclusão
          PreparaMail trim(Request("w_nome")), trim(Request("w_email")), Request("w_tipo"),1

          ScriptOpen "JavaScript"
          ShowHTML "  alert('Seu e-mail foi gravado com sucesso!\nA partir de agora você faz parte da lista de distribuição de nossa newsletter.');"
          ShowHTML "  location.href='default.asp';"
          ScriptClose
       Else
          If RS("envia_mail") = "N" Then
             SQL = "update escNewsletter set " & VbCrLf & _
                   "   envia_mail     = 'S', " & VbCrLf & _
                   "   nome           = '" & trim(Request("w_nome")) & "', " & VbCrLf & _
                   "   email          = '" & trim(Request("w_email")) & "', " & VbCrLf & _
                   "   tipo           = '" & Request("w_tipo") & "', " & VbCrLf & _
                   "   data_alteracao = getdate() " & VbCrLf & _
                   "where upper(email) = '" & uCase(Trim(Request("w_email"))) & "' " & VbCrLf
             ExecutaSQL(SQL)

             ' Envia e-mail comunicando a inclusão
             PreparaMail trim(Request("w_nome")), trim(Request("w_email")), Request("w_tipo"),1

             ScriptOpen "JavaScript"
             ShowHTML "  alert('Seu e-mail foi gravado com sucesso!\nA partir de agora você faz parte da lista de distribuição de nossa newsletter.');"
             ShowHTML "  location.href='default.asp';"
             ScriptClose
          Else
             ScriptOpen "JavaScript"
             ShowHTML "  alert('O e-mail informado já existe em nossa lista de distribuição!');"
             ShowHTML "  history.back(1);"
             ScriptClose
          End If
       End If
       dbms.CommitTrans()
  
    Case "REMOVE"
       dbms.BeginTrans()
       
       SQL = "select envia_mail from escNewsletter where upper(email) = '" & uCase(Trim(Request("w_email"))) & "' " & VbCrLf
       ConectaBD SQL
       
       If RS.EOF Then
          ScriptOpen "JavaScript"
          ShowHTML "  alert('O e-mail informado não existe em nossa lista de distribuição.');"
          ShowHTML "  location.href='default.asp';"
          ScriptClose
       Else
          If RS("envia_mail") = "S" Then
             SQL = "update escNewsletter set " & VbCrLf & _
                   "   envia_mail     = 'N', " & VbCrLf & _
                   "   data_alteracao = getdate() " & VbCrLf & _
                   "where upper(email) = '" & uCase(Trim(Request("w_email"))) & "' " & VbCrLf
             ExecutaSQL(SQL)

             ' Envia e-mail comunicando a inclusão
             PreparaMail trim(Request("w_nome")), trim(Request("w_email")), Request("w_tipo"),2

             ScriptOpen "JavaScript"
             ShowHTML "  alert('Seu e-mail foi removido com sucesso!\nA partir de agora você não faz parte da lista de distribuição de nossa newsletter.');"
             ShowHTML "  location.href='default.asp';"
             ScriptClose
          Else
             ScriptOpen "JavaScript"
             ShowHTML "  alert('O e-mail informado já foi removido da nossa lista de distribuição!');"
             ShowHTML "  history.back(1);"
             ScriptClose
          End If
       End If
       dbms.CommitTrans()
  
    Case Else
       ScriptOpen "JavaScript"
       ShowHTML "  alert('Bloco de dados não encontrado: " & SG & "');"
       ShowHTML "  history.back(1);"
       ScriptClose
  End Select
  
  Set w_chave          = Nothing
  Set w_funcionalidade = Nothing
End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que executa as operações de BD
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub Main

  strTitulo = "Cadastro"
  Select Case w_EW
     Case "GRAVA"
        Grava
     Case "REMOVER"
        Remove
     Case Else
        GetCadastro
  End Select

  Set RS = nothing
  Set RS1 = nothing

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

%>