<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="Funcoes.asp"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="esc.inc"-->
<%
Response.Expires = 0
REM =========================================================================
REM  /ControleMenu.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: TreeView da estrutura de controle
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 16/03/2000 13:35PM
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Informática
REM -------------------------------------------------------------------------

  Private SN
  Private CL
  Private w_imagem, w_tipo
  
  SN       = "Manut.Asp"
  CL       = Request("CL")
  w_tipo   = Request("w_tipo")
  w_Imagem = "img/SheetLittle.gif"

  Main

  Set w_tipo        = Nothing
  Set SN            = Nothing
  Set CL            = Nothing
  Set w_imagem      = Nothing
  
REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main

  Server.ScriptTimeOut = conScriptTimeout
  Session.TimeOut      = conSessionTimeout

  MainBody

  Server.ScriptTimeOut = conScriptTimeout

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

   ShowHTML "<HTML>"
   ShowHTML "<HEAD>"
   ShowHTML "<TITLE>" & "Controle Central" & "</TITLE>"
   ShowHTML "<style>"
   ShowHTML "<// a { color: ""#000000""; text-decoration: ""none""; } "
   ShowHTML "    a:hover { color:""#000000""; text-decoration: ""underline""; }"
   ShowHTML "    .SS{text-decoration:none;} "
   ShowHTML "    .SS:HOVER{text-decoration: ""underline"";} "
   ShowHTML "//></style>"
   ShowHTML "</HEAD>"
   ShowHTML "<BASEFONT FACE=""Verdana, Helvetica, Sans-Serif"" SIZE=""2"">"
   Response.Write "<BODY topmargin=0 bgcolor=""#FFFFFF"" BACKGROUND=""img/fundo.jpg"" BGPROPERTIES=""FIXED"" text=""#000000"" link=""#000000"" vlink=""#000000"" alink=""#FF0000""> "
   ShowHTML "  <table border=0 cellpadding=0 height=""80"" width=""100%""><tr><td nowrap><font size=1><b>"
   ShowHTML "  <TR><TD><font size=2><b>Atualização</TD></TR>"
   ShowHTML "  <TR><TD align=""center""><br><font size=1>Usuário:<b>" & Session("username") & "</TD></TR>"
   ShowHTML "  <TR><TD><font size=1><br>"
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""_blank"" CLASS=""SS"" HREF=""manuais/operacao/"" Title=""Exibe manual de operação do SIGE-WEB"">Manual SIGE-WEB</A><BR>"
   If w_tipo = 3 Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=A&w_ew=" & conWhatAdmin & "&w_ee=1"" Title=""Preenche dados administrativos da unidade de ensino!"">Administrativo</A><BR>"
   End If
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatSGE & "&w_ee=1"" Title=""Acessa a página de atualizações do SGE!"">Fotos</A><BR>"
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=A&w_ew=" & conWhatCliente & "&w_ee=1"" Title=""Altera senha de acesso e e-mail da unidade de ensino!"">Dados básicos</A><BR>"
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=A&w_ew=" & conWhatDadosAdicionais & "&w_ee=1"" Title=""Altera dados de registro e do contato da unidade exibido no site!"">Dados adicionais</A><BR>"
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=A&w_ew=" & conWhatSite & "&w_ee=1"" Title=""Altera textos e imagens exibidos no site da unidade!"">Dados do site</A><BR>"
   If w_tipo = 3 Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatEspecialidadeCliente & "&w_ee=1"" Title=""Cadastra modalidades de ensino oferecidas pela unidade!"">Áreas de atuação</A><br> "
   End If
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatDocumento & "&w_ee=1"" Title=""Cadastra arquivos para download!"">Arquivos</A><br> "
   If w_tipo = 3 Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatCalendario & "&w_ee=1"" Title=""Cadastra datas especiais da unidade!"">Calendário</A><br>"
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatMensagem & "&w_ee=1"" Title=""Cadastra mensagens da unidade dirigidas seus alunos!"">Mensagens</A><br>"
   End If
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatNotCliente & "&w_ee=1"" Title=""Cadastra notícias da unidade de ensino!"">Notícias</A><br>"
   If w_tipo = 2 Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=L&w_ew=" & conRelEscolas & "&w_ee=1"" Title=""Pesquisa unidades de ensino!"">Escolas</A><br> "
   End If
   If w_tipo = 3 Then
      ShowHTML "    <br><img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=Log&w_ee=1&P3=1&P4=30"" Title=""Exibe o registro de ocorrências do sistema!"">Log</A><br>"
   Else
      ShowHTML "    <br><img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""GR_Log.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=A&w_ew=Gerencial&w_ee=1"" Title=""Consulta ao Log!"">Log</A><BR>"
   End If
   ShowHTML "    <br><img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=ENCERRAR&w_ee=1&P3=1&P4=30"" Title=""Encerra a utilização da aplicação!"" onClick=""return(confirm('Confirma saída da aplicação?'));"">Encerrar</A>"
   ShowHTML "</TABLE>"

   ShowHTML "</BODY>"
   ShowHTML "</HTML>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do ControleMenu.asp
REM =========================================================================
%>