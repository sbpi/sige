<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../Funcoes.asp"-->
<%
Response.Expires = 0
REM =========================================================================
REM  /Default.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papad�polis
REM Descricao: Exibi��o da p�gina da escola
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 31/03/2000 09:30
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Bras�lia - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Inform�tica
REM -------------------------------------------------------------------------

  MainBody

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  ShowHTML "<html>"
  ShowHTML "<head>"
  ShowHTML "	<title>Secretaria de Estado da Educa��o</title>"
  ShowHTML "</head>"
  ShowHTML "<frameset framespacing=""0"" border=""0"" rows=""*"" frameborder=""0"">"
  ShowHTML "	<frame name=""content"" src=""http://www.itcweb.com.br/cursosige/"" target=""_self"" marginwidth=""0"" marginheight=""0"" frameborder=""NO"">"
  ShowHTML "  <noframes>"
  ShowHTML "  <body topmargin=""0"">"
  ShowHTML "  <p>Seu navegador n�o aceita quadros. Utilize o Netscape Navigator 4.x ou o Internet Explorer 5.x</p>"
  ShowHTML "  </body>"
  ShowHTML "  </noframes>"
  ShowHTML "</frameset>"
  ShowHTML "</html>"

End Sub
%>
