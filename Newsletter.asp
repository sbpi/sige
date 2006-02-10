<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="Constants_ADO.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="Funcoes.asp"-->
<!--#INCLUDE FILE="JScript.asp"-->
<%
REM =========================================================================
REM  /newsletter.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papad�polis
REM Descricao: P�gina do portal 
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 13/04/2000 23:10
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Bras�lia - DF
REM Companhia: 2000 by SBPI - Socriedade Brasileira para a Pesquisa em Inform�tica
REM -------------------------------------------------------------------------

Private RS
Private RS1
Private sobjConn
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
Set sobjConn  = Server.CreateObject("ADODB.Connection")

sobjConn.ConnectionTimeout = 300
sobjConn.CommandTimeout = 300
sobjConn.Open conConnectionString
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

ShowHTML "<div id=""rodape"">"
ShowHTML "  <div id=""endereco"">"
ShowHTML "      <div></div>"
ShowHTML "      <div></div>"
ShowHTML "      <div><br /></div> "
ShowHTML "  </div>"
ShowHTML "</div>"
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
  ShowHTML "     alert('Voc� deve selecionar uma das op��es apresentadas no formul�rio!');"
  ShowHTML "     return false;"
  ShowHTML "  }"
  ShowHTML "  theForm.Botao[0].disabled=true;"
  ShowHTML "  theForm.Botao[1].disabled=true;"
  ValidateClose
  ScriptClose
  ShowHTML "</head>"
  ShowHTML "<body>"
  ShowHTML "<div id=""container"">"
  ShowHTML "  <div id=""cab"">"
  ShowHTML "    <div id=""cabtopo"">"
  ShowHTML "      <div id=""logoesq""><img src=""img/fundo_logoesq.gif"" border=0></div>"
  ShowHTML "      <div id=""logodir""><a href=""http://www.se.df.gov.br""><img src=""img/fundo_logodir.jpg"" border=0></a></div>"
  ShowHTML "    </div>"
  ShowHTML "    <div id=""cabbase"">"
  ShowHTML "      <div id=""busca"" valign=""center""><marquee width=""100%"" align=""middle""><font color=""white"" face=""Arial"" size=""2""><b>GDF - As grandes transforma��es tamb�m no ensino p�blico.</b></font></marquee></div> "
  ShowHTML "    </div>"
  ShowHTML "  </div>"
  ShowHTML "  <div id=""corpo"">"
  ShowHTML "    <div id=""menuesq"">"
  ShowHTML "      <div id=""logomenuesq""><img src=""img/fundo_logomenuesq.gif"" border=0></div>"
  ShowHTML "      <ul id=""menugov"">"
  ShowHTML "        <li>Registrando-se nesta tela voc� receber� periodicamente informativos preparados pela SEDF."
  ShowHTML "        <li>A qualquer momento voc� poder� retirar seu e-mail da lista de distribui��o, usando a <a href=""" & sstrSN & "?EW=Remover"">tela de remo��o</a>."
  ShowHTML "        <li><font color=""#FF0000"">ATEN��O:</font> voc� deve ter um e-mail v�lido para registrar-se."
  ShowHTML "      </ul>"
  ShowHTML "      <div id=""menusep""><hr /></div>"
  ShowHTML "      <div id=""menunav""></div>"
  ShowHTML "    </div>"
  ShowHTML "    <div id=""texto""><!-- Conte�do -->"
  ShowHTML "        <table width=""570"" border=""0"">"

  ShowHTML "        <FORM ACTION=""newsletter.asp"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"">"
  ShowHTML "        <input type=""hidden"" name=""EW"" value=""Grava"">"
  ShowHTML "        <input type=""hidden"" name=""IN"" value=""Cadastro"">"
  ShowHTML "          <TR><TD align=""left"" valign""middle"">Informe os dados solicitados no formul�rio abaixo e clique no bot�o ""Registrar"" para registrar-se em nossa lista de distribui��o."
  ShowHTML "          <TR><TD>Informe seu nome completo:<br><input type=""text"" size=""60"" maxlength=""60"" name=""w_nome"" value="""" CLASS=""texto"">"
  ShowHTML "          <TR><TD>Informe o e-Mail que deseja receber a newsletter:<br><input type=""text"" size=""60"" maxlength=""60"" name=""w_email"" value="""" CLASS=""texto"">"
  ShowHTML "          <TR><TD>Selecione uma das op��es abaixo: "
  ShowHTML "                <br><input type=""Radio"" name=""w_tipo"" value=""1""> Pai, m�e ou respons�vel por aluno da rede de ensino "
  ShowHTML "                <br><input type=""Radio"" name=""w_tipo"" value=""2""> Aluno da rede de ensino "
  ShowHTML "                <br><input type=""Radio"" name=""w_tipo"" value=""3""> Outro "
  ShowHTML "          <TR><TD align=""center""><font size=""2"" CLASS=""BTM"">"
  ShowHTML "               <input type=""submit"" name=""Botao"" value=""Registrar"" class=""botao"">"
  ShowHTML "               <input type=""button"" name=""Botao"" value=""Voltar"" class=""botao"" onClick=""javascript: history.back(1);"">"
  ShowHTML "        </FORM>" & VbCrLf
End Sub
REM -------------------------------------------------------------------------
REM Final da P�gina de Pesquisa
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
  ShowHTML "<div id=""container"">"
  ShowHTML "  <div id=""cab"">"
  ShowHTML "    <div id=""cabtopo"">"
  ShowHTML "      <div id=""logoesq""><img src=""img/fundo_logoesq.gif"" border=0></div>"
  ShowHTML "      <div id=""logodir""><a href=""http://www.se.df.gov.br""><img src=""img/fundo_logodir.jpg"" border=0></a></div>"
  ShowHTML "    </div>"
  ShowHTML "    <div id=""cabbase"">"
  ShowHTML "      <div id=""busca"" valign=""center""><marquee width=""100%"" align=""middle""><font color=""white"" face=""Arial"" size=""2""><b>GDF - As grandes transforma��es tamb�m no ensino p�blico.</b></font></marquee></div> "
  ShowHTML "    </div>"
  ShowHTML "  </div>"
  ShowHTML "  <div id=""corpo"">"
  ShowHTML "    <div id=""menuesq"">"
  ShowHTML "      <div id=""logomenuesq""><img src=""img/fundo_logomenuesq.gif"" border=0></div>"
  ShowHTML "      <ul id=""menugov"">"
  ShowHTML "        <li>Removendo seu e-mail voc� deixar� de fazer parte da lista de distribui��o de informativos da SEDF."
  ShowHTML "        <li>A qualquer momento voc� poder� cadastrar-se novamente, usando a <a href=""" & sstrSN & """>tela de cadastro</a>."
  ShowHTML "      </ul>"
  ShowHTML "      <div id=""menusep""><hr /></div>"
  ShowHTML "      <div id=""menunav""></div>"
  ShowHTML "    </div>"
  ShowHTML "    <div id=""texto""><!-- Conte�do -->"
  ShowHTML "        <table width=""570"" border=""0"">"
  ShowHTML "        <FORM ACTION=""newsletter.asp"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"">"
  ShowHTML "        <input type=""hidden"" name=""EW"" value=""Grava"">"
  ShowHTML "        <input type=""hidden"" name=""IN"" value=""Remove"">"
  ShowHTML "          <TR><TD>Informe seu e-mail no campo abaixo e clique no bot�o ""Remover"" para ser removido da nossa lista de distribui��o."
  ShowHTML "          <TR><TD>Informe o e-Mail que deseja remover:<br><input type=""text"" size=""60"" maxlength=""60"" name=""w_email"" value="""" CLASS=""texto"">"
  ShowHTML "          <TR><TD align=""center"">"
  ShowHTML "               <input type=""submit"" name=""Botao"" value=""Remover"" class=""botao"">"
  ShowHTML "               <input type=""button"" name=""Botao"" value=""Voltar"" class=""botao"" onClick=""javascript: history.back(1);"">"
  ShowHTML "        </FORM>" & VbCrLf
End Sub
REM -------------------------------------------------------------------------
REM Final da P�gina de Pesquisa
REM =========================================================================

REM =========================================================================
REM Rotina de prepara��o para envio de e-mail relativo a demandas eventuais
REM Finalidade: preparar os dados necess�rios ao envio autom�tico de e-mail
REM Par�metro: p_solic: n�mero de identifica��o da solicita��o. 
REM            p_tipo:  1 - Inclus�o
REM                     2 - Remo��o
REM -------------------------------------------------------------------------
Sub PreparaMail(p_nome, p_mail, p_tipo, p_evento)

  Dim w_html, w_destinatarios, w_assunto
  
  w_html = "<HTML>" & VbCrLf
  w_html = w_html & BodyOpenMail(null) & VbCrLf
  w_html = w_html & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & VbCrLf
  w_html = w_html & "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">" & VbCrLf
  w_html = w_html & "    <table width=""97%"" border=""0"">" & VbCrLf
  If p_evento = 1 Then
     w_html = w_html & "      <tr valign=""top""><td align=""center""><font size=2><b>INCLUS�O NA LISTA DE DISTRIBUI��O</b></font><br><br><td></tr>" & VbCrLf
     w_html = w_html & "      <tr align=""center""><td><font size=2><b><font color=""#BC3131"">ATEN��O</font>: Esta � uma mensagem de envio autom�tico. N�o responda.</b></font><br><br><td></tr>" & VbCrLf
     w_html = w_html & VbCrLf & "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
     w_html = w_html & VbCrLf & "      <tr><td><font size=2>Seu e-mail foi inclu�do na lista de distribui��o de informativos da Secretaria de Estado de Educa��o do DF. A partir de agora, sempre que algum informativo for enviado para a lista, voc� ser� um dos destinat�rios.</b></font></td></tr>"
     w_html = w_html & "      <tr align=""center""><td><br><br><br><font size=1>" & VbCrLf
     w_html = w_html & "         Se desejar remover seu e-mail, clique <a class=""SS"" href=""http://www.gdfsige.df.gov.br/newsletter.asp?EW=Remover"" target=""_blank"">aqui</a>." & VbCrLf
     w_html = w_html & "      </font></td></tr>" & VbCrLf
  ElseIf p_evento = 2 Then
     w_html = w_html & "      <tr valign=""top""><td align=""center""><font size=2><b>SA�DA DA LISTA DE DISTRIBUI��O</b></font><br><br><td></tr>" & VbCrLf
     w_html = w_html & "      <tr align=""center""><td><font size=2><b><font color=""#BC3131"">ATEN��O</font>: Esta � uma mensagem de envio autom�tico. N�o responda.</b></font><br><br><td></tr>" & VbCrLf
     w_html = w_html & VbCrLf & "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
     w_html = w_html & VbCrLf & "      <tr><td><font size=2>Seu e-mail foi removido da lista de distribui��o de informativos da Secretaria de Estado de Educa��o do DF.</b></font></td></tr>"
     w_html = w_html & "      <tr align=""center""><td><br><br><br><font size=1>" & VbCrLf
     w_html = w_html & "         Se desejar cadastrar novamente seu e-mail, clique <a class=""SS"" href=""http://www.gdfsige.df.gov.br/newsletter.asp?EW=i"" target=""_blank"">aqui</a>." & VbCrLf
  End IF
  w_html = w_html & "      <tr valign=""top""><td><br><br><br><font size=1>" & VbCrLf
  w_html = w_html & "         Dados da ocorr�ncia:<br>" & VbCrLf
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

  ' Prepara os dados necess�rios ao envio
  If p_evento = 1 Then ' Inclus�o ou Conclus�o
     w_assunto = "SEDF - Inclus�o na lista de distribui��o de informativos"
  ElseIf p_evento = 2 Then ' Tramita��o
     w_assunto = "SEDF - Sa�da da lista de distribui��o de informativos"
  End If

  ' Executa o envio do e-mail
  w_resultado = EnviaMail(w_assunto, w_html, p_mail, null)
        
  ' Se ocorreu algum erro, avisa da impossibilidade de envio
  If w_resultado > "" Then 
     ScriptOpen "JavaScript"
     ShowHTML "  alert('ATEN��O: n�o foi poss�vel proceder o envio do e-mail.\n" & w_resultado & "');" 
     ScriptClose
  End If

  Set w_html                   = Nothing
  Set w_destinatarios          = Nothing
  Set w_assunto                = Nothing
End Sub
REM =========================================================================
REM Fim da rotina da prepara��o para envio de e-mail
REM -------------------------------------------------------------------------

REM =========================================================================
REM Procedimento que executa as opera��es de BD
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
          
          ' Envia e-mail comunicando a inclus�o
          PreparaMail trim(Request("w_nome")), trim(Request("w_email")), Request("w_tipo"),1

          ScriptOpen "JavaScript"
          ShowHTML "  alert('Seu e-mail foi gravado com sucesso!\nA partir de agora voc� faz parte da lista de distribui��o de nossa newsletter.');"
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

             ' Envia e-mail comunicando a inclus�o
             PreparaMail trim(Request("w_nome")), trim(Request("w_email")), Request("w_tipo"),1

             ScriptOpen "JavaScript"
             ShowHTML "  alert('Seu e-mail foi gravado com sucesso!\nA partir de agora voc� faz parte da lista de distribui��o de nossa newsletter.');"
             ShowHTML "  location.href='default.asp';"
             ScriptClose
          Else
             ScriptOpen "JavaScript"
             ShowHTML "  alert('O e-mail informado j� existe em nossa lista de distribui��o!');"
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
          ShowHTML "  alert('O e-mail informado n�o existe em nossa lista de distribui��o.');"
          ShowHTML "  location.href='default.asp';"
          ScriptClose
       Else
          If RS("envia_mail") = "S" Then
             SQL = "update escNewsletter set " & VbCrLf & _
                   "   envia_mail     = 'N', " & VbCrLf & _
                   "   data_alteracao = getdate() " & VbCrLf & _
                   "where upper(email) = '" & uCase(Trim(Request("w_email"))) & "' " & VbCrLf
             ExecutaSQL(SQL)

             ' Envia e-mail comunicando a inclus�o
             PreparaMail trim(Request("w_nome")), trim(Request("w_email")), Request("w_tipo"),2

             ScriptOpen "JavaScript"
             ShowHTML "  alert('Seu e-mail foi removido com sucesso!\nA partir de agora voc� n�o faz parte da lista de distribui��o de nossa newsletter.');"
             ShowHTML "  location.href='default.asp';"
             ScriptClose
          Else
             ScriptOpen "JavaScript"
             ShowHTML "  alert('O e-mail informado j� foi removido da nossa lista de distribui��o!');"
             ShowHTML "  history.back(1);"
             ScriptClose
          End If
       End If
       dbms.CommitTrans()
  
    Case Else
       ScriptOpen "JavaScript"
       ShowHTML "  alert('Bloco de dados n�o encontrado: " & SG & "');"
       ShowHTML "  history.back(1);"
       ScriptClose
  End Select
  
  Set w_chave          = Nothing
  Set w_funcionalidade = Nothing
End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que executa as opera��es de BD
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
  Set sobjConn  = nothing

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

%>