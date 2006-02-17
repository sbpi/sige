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
REM Criacao  : 20/01/2004 11:40
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Copyright: 2004 by SBPI Consultoria Ltda
REM -------------------------------------------------------------------------

  Private SN
  Private CL
  Private w_imagem
  
  SN       = "Controle.asp"
  CL       = Request("CL")
  w_Imagem = "img/SheetLittle.gif"

  Main

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
   ShowHTML "  <TR><TD align=""center""><font size=2><b>Atualização</TD></TR>"
   ShowHTML "  <TR><TD align=""center""><br><font size=1>Usuário:<b>" & Session("username") & "</TD></TR>"
   ShowHTML "  <TR><TD><font size=1><br>"

   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""_blank"" CLASS=""SS"" HREF=""manuais/operacao/"" Title=""Exibe manual de operação do SIGE-WEB"">Manual SIGE-WEB</A><BR>"
   If Session("username") = "ADMINISTRATIVO" or Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=A&w_ew=" & conWhatAdmin & "&w_ee=1"" Title=""Preenche dados administrativos da unidade de ensino!"">Administrativo</A><BR>"
   End If

   If Session("username") = "SEDF" or Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatDocumento & "&w_ee=1"" Title=""Cadastra arquivos para download!"">Arquivos</A><br> "
   End If

   If Session("username") = "GESTORVERSAO" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatSGE & "&w_ee=1"" Title=""Tela para gestão de versões!"">Sistemas</A><br> "
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=COMPONENTE&w_ee=1"" Title=""Tela para gestão de versões!"">Componentes</A><br> "
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=VERSAO&w_ee=1"" Title=""Tela para gestão de versões!"">Versões</A><br> "
   End If

   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=base&w_ee=1"" Title=""Cadastra calendário oficial!"">Calendário Base</A><br>"
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatCalendario & "&w_ee=1"" Title=""Cadastra datas especiais da rede de ensino!"">Calendário da rede</A><br>"
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""escRegional.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=L&w_ew=ligacao&w_ee=1"" Title=""Permite alterar a regional e o tipo da escola!"">Dados da escola</A><BR>"
   End If

   If Session("username") = "IMPRENSA" or Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Mensagem.asp?CL=" & CL & "&O=L&par=inicial&TP=e-Mail automático"" Title=""Permite o envio de e-mails através do sistema!"">Envio de e-Mail</A><br>"
   End If

   If Session("username") <> "ADMINISTRATIVO" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conRelEscolas & "&w_ee=1"" Title=""Pesquisa unidades de ensino!"">Escolas</A><br> "
   End If

   If Session("username") = "IMPRENSA" or Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=newsletter&w_ee=1"" Title=""Acessa a lista de distribuição da newsletter!"">Informativo</A><br>"
   End If

   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatEspecialidade & "&w_ee=1"" Title=""Cadastra modalidades de ensino oferecidas pela rede de ensino!"">Modalidades de ensino</A><br>"
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatMensagem & "&w_ee=1"" Title=""Cadastra mensagens dirigidas seus alunos!"">Mensagens a alunos</A><br>"
   End If

   If Session("username") = "SEDF" OR Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatNotCliente & "&w_ee=1"" Title=""Cadastra notícias da rede de ensino!"">Notícias</A><br>"
   End If

   If Session("username") = "SEDF" or Session("username") = "SBPI" then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""controle.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=L&w_ew=TIPOCLIENTE&w_ee=1"" Title=""Tipos de Cliente!"">Tipo de Instituição</A><BR>"
   End If

   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=A&w_ew=" & conWhatCliente & "&w_ee=1"" Title=""Altera a senha de acesso!"">Senha</A><BR>"

   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=senhaesp&w_ee=1"" Title=""Exibe a lista de senhas de regionais e outros usuários (menos escolas)!"">Senhas especiais</A><br>"
   End If

   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""GR_Log.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=A&w_ew=Gerencial&w_ee=1"" Title=""Consulta ao Log!"">Log</A><BR>"
   End If
      
   ShowHTML "    <br><img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & encerrar & "&w_ee=1"" Title=""Encerra a utilização da aplicação!"">Encerrar</A> "
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