<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/Constants_ADO.inc"-->
<!--#INCLUDE VIRTUAL="/Modelos/Constants.inc"-->
<!--#INCLUDE VIRTUAL="/esc.inc"-->
<!--#INCLUDE VIRTUAL="/Funcoes.asp"-->
<%
Response.Expires = 0
REM =========================================================================
REM  /Default.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: Exibição da página da escola
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 31/03/2000 09:30
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Informática
REM -------------------------------------------------------------------------

  Private sobjRS
  Private sobjConn

  Private sstrSN
  Private sstrCL

  sstrSN = "/Modelos/Mod15/Default.Asp"

  Public sstrEW
  Public sstrIN
  Public CL

  sstrEW = Request.QueryString("EW")
  sstrIN = Request.QueryString("IN")

  CL = *%*        ' Informa o código do cliente
  
  Main

REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main

  Set sobjRS = Server.CreateObject("ADODB.RecordSet")
  Set sobjConn  = Server.CreateObject("ADODB.Connection")

  sobjConn.ConnectionTimeout = 300
  sobjConn.CommandTimeout = 300
  sobjConn.Open conConnectionString

  MainBody

  Server.ScriptTimeOut = Session("ScriptTimeOut")

  Set sobjRS = nothing
  Set sobjConn  = nothing
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  Dim sql

  sql = "SELECT ds_cliente FROM escCliente WHERE sq_cliente=" & CL

  sobjRS.Open sql, sobjConn, adOpenForwardOnly

  Session("BodyWidth") = "620"

  ShowHTML "<html>"
  ShowHTML "<head>"
  ShowHTML "	<title>" & sobjRS("ds_cliente") & "</title>"
  ShowHTML "</head>"
  ShowHTML "<frameset framespacing=""0"" border=""0"" rows=""*"" frameborder=""0"">"
  ShowHTML "	<frame name=""content"" src=""" & sstrSN & "?EW=110&EF=sq_cliente=" & CL & "&CL=" & CL & """ target=""_self"" marginwidth=""0"" marginheight=""0"" frameborder=""NO"">"
  ShowHTML "  <noframes>"
  ShowHTML "  <body topmargin=""0"">"
  ShowHTML "  <p>Seu navegador não aceita quadros. Utilize o Netscape Navigator 4.x ou o Internet Explorer 5.x</p>"
  ShowHTML "  </body>"
  ShowHTML "  </noframes>"
  ShowHTML "</frameset>"
  ShowHTML "</html>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>
