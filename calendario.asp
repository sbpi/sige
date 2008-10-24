<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="Funcoes.asp"-->
<!--#INCLUDE FILE="FuncoesGR.asp"-->
<!--#INCLUDE FILE="jScript.asp"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="Constants_ADO.inc"-->
<!--#INCLUDE FILE="esc.inc"-->
<%
Response.Expires = -1500
REM =========================================================================
REM  /privada.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: Manutenção dos dados das escolas
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 28/03/2000 20:20
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Informática
REM -------------------------------------------------------------------------

  Private w_EA
  Private w_IN
  Private w_EF
  Private w_R
  Private CL, DBMS, SQL, RS, RS1
  Private w_Data, w_pagina, w_diretorio, w_ds_diretorio, w_tipo
  
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
    
  If InStr(uCase(Request.ServerVariables("http_content_type")),"MULTIPART/FORM-DATA") > 0 Then
     ' Cria o objeto de upload
     w_EW = Request("w_ew")
     Dim ul, File
     Set ul = Nothing
     Set ul = Server.CreateObject("Dundas.Upload.2")
     ul.SaveToMemory
     w_EA   = ul.Form("w_ea")
     w_R    = ul.Form("R")
     CL     = ul.Form("CL")
  Else
     w_R  = Request("R")
     w_EA = ucase(Request("w_ea"))
     w_EW = ucase(Request("w_ew"))
     w_IN = ucase(Request("w_in"))
     w_EF = ucase(Request("w_ef"))
     CL   = Request("CL")
     p_tipo = Nvl(Request("p_tipo"),"H")
     p_campos = Request("p_campos")
  End If

  w_Data = Mid(100+Day(Date()),2,2) & "/" & Mid(100+Month(Date()),2,2) & "/" &Year(Date())
  w_Pagina = ExtractFileName(Request.ServerVariables("SCRIPT_NAME")) & "?w_ew="
  
  If w_ew = "LOGON" Then
     AbreSessaoManut
     Logon
  ElseIf w_ew = "VALIDA" Then
     AbreSessaoManut
     Valida
  ElseIf w_ew = "ENCERRAR" Then
     Session.Abandon()
     ShowHTML "<HTML>" & chr(13) & chr(10)
     ShowHTML "<BODY>" & chr(13) & chr(10)
     ShowHTML "<SCRIPT LANGUAGE='JAVASCRIPT'>" & chr(13) & chr(10)
     ShowHTML "  top.location.href='calendario.asp?w_ew=logon';" & chr(13) & chr(10)
     ShowHTML "</SCRIPT>" & chr(13) & chr(10)
     ShowHTML "</BODY>" & chr(13) & chr(10)
     ShowHTML "</HTML>" & chr(13) & chr(10)
  ElseIf Session("Username") > "" and  CL > "" Then
     AbreSessaoManut
     
     SQL = "select b.tipo from escCliente a inner join escTipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente and a." & CL & ")"
     ConectaBD SQL
     w_tipo = RS("tipo")
     DesconectaBD
     Main
  Else
     AbreSessaoManut
     ShowHTML "<HTML>" & chr(13) & chr(10)
     ShowHTML "<BODY>" & chr(13) & chr(10)
     ShowHTML "<SCRIPT LANGUAGE='JAVASCRIPT'>" & chr(13) & chr(10)
     ShowHTML "  top.location.href='calendario.asp?w_ew=logon';" & chr(13) & chr(10)
     ShowHTML "</SCRIPT>" & chr(13) & chr(10)
     ShowHTML "</BODY>" & chr(13) & chr(10)
     ShowHTML "</HTML>" & chr(13) & chr(10)
  End If
  
  Set w_tipo        = Nothing
  Set w_R           = Nothing
  Set w_EA          = Nothing
  Set w_IN          = Nothing
  Set w_EF          = Nothing
  Set CL            = Nothing
  Set SQL           = Nothing
  Set RS            = Nothing
  Set DBMS          = Nothing
  Set w_Data        = Nothing
  Set w_Pagina      = Nothing
  Set w_Diretorio   = Nothing
  Set w_ds_diretorio= Nothing

REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main

  Server.ScriptTimeOut = conScriptTimeout
  Session.TimeOut      = conSessionTimeout

  MainBody

End Sub

REM =========================================================================
REM Monta a Frame
REM -------------------------------------------------------------------------
Public Sub ShowFrames

  Session("BodyWidth") = "620"

  ShowHTML "<html>"
  ShowHTML "<head>"
  ShowHTML "    <title>Controle Central</title>"
  ShowHTML "</head>"
  ShowHTML "<frameset cols=""200,*"">"
  ShowHTML "    <frame id=""Menu"" name=""Menu"" src=""calendario.asp?CL=" & CL & "&w_tipo=" & w_tipo & "&w_ew=SHOWMENU"" scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
  ShowHTML "    <frame id=""Body"" name=""Body"" src=""calendario.asp?CL=" & CL & "&w_ee=1&w_ea=L&w_ew=" & conWhatCalendario & """ scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
  ShowHTML "</frameset>"
  ShowHTML "</html>"

End Sub

REM =========================================================================
REM Rotina de montagem do menu
REM -------------------------------------------------------------------------
Sub showArquivos

   Dim w_imagem
   
   w_imagem = "img/SheetLittle.gif"

   ShowHTML "<HTML>"
   ShowHTML "<HEAD>"
   ShowHTML "<TITLE>" & "Controle Central" & "</TITLE>"
   ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
   ShowHTML "<style>"
   ShowHTML "<// a { color: ""#000000""; text-decoration: ""none""; } "
   ShowHTML "    a:hover { color:""#000000""; text-decoration: ""underline""; }"
   ShowHTML "    .SS{text-decoration:none;} "
   ShowHTML "    .SS:HOVER{text-decoration: ""underline"";} "
   ShowHTML "//></style>"
   ShowHTML "</HEAD>"
   ShowHTML "<BASEFONT FACE=""Verdana, Helvetica, Sans-Serif"" SIZE=""2"">"
   Response.Write "<BODY topmargin=0 bgcolor=""#FFFFFF"" BACKGROUND=""img/background.gif"" BGPROPERTIES=""FIXED"" text=""#000000"" link=""#000000"" vlink=""#000000"" alink=""#FF0000""> "


   ShowHTML "</BODY>"
   ShowHTML "</HTML>"
   
   

End Sub

REM =========================================================================
REM Rotina de montagem do menu
REM -------------------------------------------------------------------------
Sub showMenu

   Dim w_imagem
   
   w_imagem = "img/SheetLittle.gif"

   ShowHTML "<HTML>"
   ShowHTML "<HEAD>"
   ShowHTML "<TITLE>" & "Controle Central" & "</TITLE>"
   ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
   ShowHTML "   <link href=""/css/impressao.css"" media=""print"" rel=""stylesheet"" type=""text/css"" />"
   ShowHTML "<style>"
   ShowHTML "<// a { color: ""#000000""; text-decoration: ""none""; } "
   ShowHTML "    a:hover { color:""#000000""; text-decoration: ""underline""; }"
   ShowHTML "    .SS{text-decoration:none;} "
   ShowHTML "    .SS:HOVER{text-decoration: ""underline"";} "
   ShowHTML "//></style>"
   ShowHTML "</HEAD>"
   ShowHTML "<BASEFONT FACE=""Verdana, Helvetica, Sans-Serif"" SIZE=""2"">"
   Response.Write "<BODY topmargin=0 bgcolor=""#FFFFFF"" BACKGROUND=""img/background.gif"" BGPROPERTIES=""FIXED"" text=""#000000"" link=""#000000"" vlink=""#000000"" alink=""#FF0000""> "
   ShowHTML "  <table id=""menu"" border=0 cellpadding=0 height=""80"" width=""100%""><tr><td nowrap><font size=1><b>"
   ShowHTML "  <TR><TD><font size=2><b>Conferência</TD></TR>"
   ShowHTML "  <TR><TD><font size=1><br>"
   'ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=A&w_ew=" & conWhatDadosAdicionais & "&w_ee=1"">Dados cadastrais</A><BR>"
   'ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatDocumento & "&w_ee=1"">Cursos</A><br> "
   'ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatMensagem & "&w_ee=1"">Ordens de serviço</A><br>"
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=L&w_ew=CALS&w_ee=1"">Calendários</A><br>"
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatCalendario & "&w_ee=1"">Datas de calendário</A><br>"
   'ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatNotCliente & "&w_ee=1"">Portarias</A><br>"
   'ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=L&w_ew=" & "formulario" & "&w_ee=1"">Formulário de Atualização</A><br>"
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=L&w_ew=" & "formcal" & "&w_ee=1"">Formulário de Calendário</A><br>"
   ShowHTML "    <br><img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=A&w_ew=" & conWhatCliente & "&w_ee=1"" Title=""Altera senha de acesso e e-mail da unidade de ensino!"">Alteração de senha</A><BR>"

   ShowHTML "    <br><img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" HREF=""calendario.asp?CL=" & CL & "&w_ea=L&w_ew=ENCERRAR&w_ee=1&P3=1&P4=30"" Title=""Encerra a utilização da aplicação!"" onClick=""return(confirm('Confirma saída da aplicação?'));"">Encerrar</A>"
   ShowHTML "</TABLE>"

   ShowHTML "</BODY>"
   ShowHTML "</HTML>"
   
   Set w_imagem = Nothing

End Sub

REM =========================================================================
REM Rotina de criação da tela de logon
REM -------------------------------------------------------------------------
Sub LogOn
  
  Dim wAno
  
  ShowHTML "<HTML>"
  ShowHTML "<HEAD>"
  ShowHTML "<TITLE>SigeWeb - Autenticação</TITLE>"
  ScriptOpen "JavaScript"
  ValidateOpen "Validacao"
  'Validate "CL"        , "Unidade"         , ""  , "1" , "1" , "14" ,  "" , "1"
  Validate "Login1"    , "Nome de usuário" , ""  , "1" , "2" , "14" , "1" , "1"
  Validate "Password1" , "Senha"           , "1" , "1" , "3" , "19" , "1" , "1"
  ShowHTML "  theForm.Password.value = theForm.Password1.value; "
  ShowHTML "  theForm.Password1.value = """"; "
  ShowHTML "  theForm.Botao.disabled = true; "
  ShowHTML "  theForm.Login.value = theForm.Login1.value; "
  ShowHTML "  theForm.Login1.value = """"; "
  ValidateClose
  ScriptClose
  ShowHTML "<BASEFONT FACE=""Verdana"" SIZE=""2""> "
  ShowHTML "<style> "
  ShowHTML " .SS{text-decoration:none;font:bold 8pt} "
  ShowHTML " .SS:HOVER{text-decoration: underline;} "
  ShowHTML " .HL{text-decoration:none;font:Arial;color=""#0000FF""} "
  ShowHTML " .HL:HOVER{text-decoration: underline;} "
  ShowHTML " .TTM{font: 10pt Arial}"
  ShowHTML " .BTM{font: 8pt Verdana}"
  ShowHTML " .XTM{font: 12pt Verdana}"
  ShowHtml " .STI {font-size: 8pt; border: 1px solid #000000; background-color: #F5F5F5}"  & VbCrLf
  ShowHtml " .STB {font-size: 8pt; color: #FFFFFF; border: 1px solid #000000; background-color: #808080; }"  & VbCrLf
  ShowHTML " .sts {font-size: 8pt; border-top: 1px solid #000000; background-color: #F5F5F5}"
  ShowHTML "</style> "
  ShowHTML "</HEAD>"
  If wNoUsuario = "" Then
     ShowHTML "<body topmargin=0 leftmargin=10 onLoad=""document.Form.Login1.focus();"">"
  Else
     ShowHTML "<body topmargin=0 leftmargin=10 onLoad=""document.focus();"">"
  End If
  ShowHTML "<CENTER>"
  ShowHTML "<form method=""post"" action=""calendario.asp?w_ew=Valida"" onsubmit=""return(Validacao(this));"" name=""Form""> "
  ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""Login"" VALUE=""""> "
  ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""Password"" VALUE=""""> "
  ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""w_ew"" VALUE=""Valida""> "
  ShowHTML "<TABLE cellSpacing=0 cellPadding=0 width=""760"" height=550 border=1  background=""files/" & p_cliente & "/img/fundo.jpg"" bgproperties=""fixed""><tr><td width=""100%"" valign=""top"">"
  ShowHTML "  <TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0 background=""img/cabecalho.gif"">"
  ShowHTML "    <TR>"
  ShowHTML "      <TD vAlign=TOP align=middle width=""65%"">"
  ShowHTML "        <B><FONT face=Arial size=5 color=#000088>Secretaria de Estado de Educação"
  ShowHTML "        <br><br>SIGE - Módulo Web</B></FONT>"
  ShowHTML "     </TD>"
  ShowHTML "    </TR>"
  ShowHTML "    <TR><TD colspan=3 borderColor=#ffffff height=22><HR align=center color=#808080></TD></TR>"
  ShowHTML "  </TABLE>"
  ShowHTML "  <table width=""100%"" border=""0"">"
  ShowHTML "    <tr><td valign=""middle"" width=""100%"" height=""100%"">"
  ShowHTML "        <table width=""100%"" height=""100%"" border=""0"">"
  ShowHTML "          <tr><td align=""center"" colspan=2><font size=""1"" color=""#990000""><b>Esta aplicação é de uso interno da Secretaria de Estado de Educação.<br>As informações contidas nesta aplicação são restritas e de uso exclusivo.<br>O uso indevido acarretará ao infrator penalidades de acordo com a legislação em vigor.<br><br>Informe os dados solicitados e clique no botão <i>OK</i> para ser autenticado pela aplicação.</b></font>"
  ShowHTML "          <TR><TD colspan=2 height=19 align=center><IMG height=2 src=""img/linha.jpg"" width=551></TD></TR>"
  ShowHTML "          <tr><td align=""right"" width=""43%""><font size=""2""><b>Nome de usuário:<td><input class=""sti"" name=""Login1"" size=""14"" maxlength=""14"">"
  ShowHTML "          <tr><td align=""right""><font size=""2""><b>Senha:<td><input class=""sti"" type=""Password"" name=""Password1"" size=""19"">"
  ShowHTML "          <tr><td align=""right""><td><font size=""2""><b><input class=""stb"" type=""submit"" value=""OK"" name=""Botao""> "
  ShowHTML "          </font></td> </tr> "
  ShowHTML "          <TR><TD colspan=2 align=""center""><br><table border=0 cellpadding=0 cellspacing=0><tr><td>"
  ShowHTML "              <P><IMG height=37 src=""img/ajuda.jpg"" width=629><br>"
  ShowHTML "              <font face=""Arial"" size=1><b>PARA ACESSAR A PÁGINA DE ATUALIZAÇÃO</b></font>"
  ShowHTML "              <FONT face=""Verdana, Arial, Helvetica, sans-serif"" size=1>"
  ShowHTML "              <li>Nome de usuário - Informe o nome de usuário da sua escola ou da sua regional de ensino"
  ShowHTML "              <li>Senha - Informe sua senha de acesso"
  ShowHTML "              <li>Se esqueceu ou não foi informado dos dados acima, favor entrar em contato com a SEDF / SUBIP / Diretoria de Sistemas de Informação Educacional - DSIE<br>"
  ShowHTML "              </FONT></P>"
  ShowHTML "              <P><font face=""Arial"" size=1><b>DOCUMENTAÇÃO - LEIA COM ATENÇÃO</b></font><br>"
  ShowHTML "              <FONT face=""Verdana"" size=1>"
  ShowHTML "              . <a class=""SS"" href=""sedf/Orientacoes_Acesso.pdf"" target=""_blank"" title=""Abre arquivo que descreve as novas características e funcionalidades do SIGE-WEB."">Apresentação da nova versão do SIGE-WEB (PDF - 130KB - 4 páginas)</a><BR>"
  ShowHTML "              . <a class=""SS"" href=""manuais/operacao/"" target=""_blank"" title=""Exibe manual de operação do SIGE-WEB"">Manual SIGE-WEB (HTML)</A><BR>"
  ShowHTML "              </FONT></P></TD></TR>"
  ShowHTML "</table>"
  ShowHTML "              </TD></TR>"
  ShowHTML "          </table> "  
  wAno =  Request.QueryString("wAno")
  
  If wAno = "" Then
     wAno = Year(Date())
  End If

'  sql = "SELECT case in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end AS in_destinatario, " & VbCrLf & _
'        "       dt_arquivo, ds_titulo, nr_ordem, ds_arquivo, ln_arquivo, 'SEDF' AS Origem, x.ln_internet diretorio " & VbCrLf & _
'        "From escCliente_Arquivo AS a INNER JOIN escCliente AS x ON (a.sq_site_cliente = x.sq_cliente)" & VbCrLf & _
'        "WHERE in_ativo = 'Sim'" & VbCrLf & _
'        "  AND x.sq_cliente = 0" & VbCrLf & _
'        "  AND in_destinatario = 'E'" & VbCrLf & _
'        "  and YEAR(a.dt_arquivo) = " & wAno & VbCrLf & _
'        "ORDER BY origem, nr_ordem, dt_arquivo desc, in_destinatario " & VbCrLf 
'  ConectaBD SQL
'  ShowHTML " <tr><td><table width=""100%"" border=""0"">"
'  ShowHTML "          <TR><TD align=""center""><br><table border=0 cellpadding=0 cellspacing=0><tr><td>"
'  ShowHTML "              <P><font face=""Arial"" size=1><b>ARQUIVOS INSERIDOS PELA SEDF</b></font><br>"  
'  ShowHTML "<tr><td><TABLE border=0 cellSpacing=5 width=""95%"">"
'  ShowHTML "  <TR>"
'  ShowHTML "    <TD><FONT face=""Verdana"" size=1><b>Origem"
'  ShowHTML "    <TD><FONT face=""Verdana"" size=1><b>Alvo"
'  ShowHTML "    <TD><FONT face=""Verdana"" size=1><b>Data"
'  ShowHTML "    <TD><FONT face=""Verdana"" size=1><b>Componente curricular"
'  ShowHTML "  </TR>"
'  ShowHTML "  <TR>"
'  ShowHTML "    <TD COLSPAN=""4"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
'  ShowHTML "  </TR>"
'  wCont = 0
'     
'  If Not RS.EOF Then
'     wCont = 1
'     Do While Not RS.EOF
'        ShowHTML "  <TR valign=""top"">"
'        ShowHTML "    <TD><FONT face=""Verdana"" size=1>" & RS("origem")
'        ShowHTML "    <TD><FONT face=""Verdana"" size=1>" & RS("in_destinatario")
'        ShowHTML "    <TD><FONT face=""Verdana"" size=1>" & Mid(100+Day(RS("dt_arquivo")),2,2) & "/" & Mid(100+Month(RS("dt_arquivo")),2,2) & "/" &Year(RS("dt_arquivo"))
'   REM  ShowHTML "    <TD><FONT face=""Verdana"" size=1><a href=""http://" & replace(RS("diretorio"),"http://","") & "//" & RS("ln_arquivo") & """ target=""_blank"">" & RS("ds_titulo") & "</a><br><div align=""justify""><font size=1>.:. " & RS("ds_arquivo") & "</div>" Original
'		ShowHTML "    <TD><FONT face=""Verdana"" size=1><a href=""http://" & replace(replace(RS("diretorio"),"http://","") , "se.df.gov.br" , "gdfsige.df.gov.br") & "/sedf/sedf/" & RS("ln_arquivo") & """ target=""_blank"">" & RS("ds_titulo") & "</a><br><div align=""justify""><font size=1>.:. " & RS("ds_arquivo") & "</div>"
'        ShowHTML "  </TR>"
'
'        RS.MoveNext
'     Loop
'  Else
'     ShowHTML "  <TR><TD COLSPAN=4 ALIGN=CENTER><FONT face=""Verdana"" size=1>Não há arquivos disponíveis no momento para o ano de " & wAno & " </TR>"
'  End If
'  RS.Close

'  SQL = "SELECT year(dt_arquivo) ano " & VbCrLf & _
'        "From escCliente_Arquivo a INNER JOIN escCliente AS x ON (a.sq_site_cliente = x.sq_cliente)" & VbCrLf & _
'        "WHERE in_ativo = 'Sim'" & VbCrLf & _
'        "  AND x.sq_cliente = 0" & VbCrLf & _
'        "  AND in_destinatario = 'E'" & VbCrLf & _
'        "  and YEAR(a.dt_arquivo) <> " & wAno & VbCrLf & _
'        "ORDER BY year(dt_arquivo) desc " & VbCrLf 
'  ConectaBD SQL
'    If Not RS.EOF Then
'       ShowHTML "  <TR><TD COLSPAN=4 ><FONT face=""Verdana"" size=1><b>Arquivos de outros anos</b><br>"
'       While Not RS.EOF
'          ShowHTML "     <li><a href=""" & w_dir & "calendario.asp?wAno=" & RS("ano") & """ >Arquivos de " & RS("ano") & "</a></TR>"
'          RS.MoveNext
'       Wend
'       ShowHTML "  </TD></TR>"
'    End If
'    RS.Close
'  ShowHTML "</table>"
  ShowHTML "              </FONT></P>"
  ShowHTML "              </TD></TR>"
  ShowHTML "          </table> "  
  ShowHTML "        </table> "
  ShowHTML "    </tr> "
  ShowHTML "  </table>"
  ShowHTML "</table>"   
  ShowHTML "</form> "  
  ShowHTML "</CENTER>"
  ShowHTML "</body>"
  ShowHTML "</html>"
End Sub

REM =========================================================================
REM Rotina de autenticação dos usuários
REM -------------------------------------------------------------------------
Sub Valida

  Dim w_Erro, w_cliente, w_uid, w_pwd
  
  w_uid = replace(replace(Trim(uCase(Request("Login"))),"'", ""), """", "")
  w_pwd = replace(replace(Trim(uCase(Request("Password"))),"'", ""), """", "")

  w_Erro = 0
  SQL = "select count(*) existe from escCliente where upper(ds_username) = upper('" & w_uid & "') and PUBLICA = 'N'"
  ConectaBD SQL
  If RS("existe") = 0 Then
     w_Erro = 1
  Else
     DesconectaBD
     SQL = "select count(*) existe from escCliente where upper(ds_username) = upper('" & w_uid & "') and ds_senha_acesso = '" & w_pwd & "'"
     
     
     ConectaBD SQL
     If RS("existe") = 0 Then
        w_Erro = 2
     End If
  End If
  DesconectaBD
  ScriptOpen "JavaScript"
  If w_erro > 0 Then
     If w_Erro = 1 Then
        ShowHTML "  alert('Combinação Usuário e Senha inexistente!');"
     Else
        ShowHTML "  alert('Senha inválida!');"
     End If
     ShowHTML "  history.back(1);"
  Else
     ' Recupera informações a serem usadas na montagem das telas para o usuário
     SQL = "select * from escCliente where upper(ds_username) = upper('" & w_uid & "') and ds_senha_acesso = '" & w_pwd & "'"
     ConectaBD SQL
     Session("username")     = uCase(RS("ds_username"))
     Session("nome_cliente") = RS("ds_cliente")
     w_cliente               = RS("sq_cliente")
     DesconectaBD     
     
     '   o acesso na tabela de log
     SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql) " & VbCrLf & _
           "values ( " & VbCrLf & _
           "         " & w_cliente & ", " & VbCrLf & _
           "         getdate(), " & VbCrLf & _
           "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
           "         0, " & VbCrLf & _
             "         'Acesso à tela de atualização da escola.', " & VbCrLf & _
           "         null " & VbCrLf & _
           "       ) " & VbCrLf
     ExecutaSQL(SQL)

     ShowHTML "  location.href='calendario.asp?&w_in=1&cl=sq_cliente=" & w_cliente & "';"
  End If
  ScriptClose

  Set w_Erro    = Nothing
  Set w_Cliente = Nothing

End Sub

Public Sub Grava

  Dim sql
    
  Cabecalho
  ShowHTML "</HEAD>"
  BodyOpen "onLoad=document.focus();"
  
  If w_R <> conWhatSGE AND w_R <> "CALS" Then
     SQL = "select sq_funcionalidade from escFuncionalidade where tipo = 2 and codigo = '" & w_R & "'"
     ConectaBD SQL
     w_funcionalidade = RS("sq_funcionalidade")
     DesconectaBD
  End If
  
Select Case w_R
  
    Case conWhatCliente
       dbms.BeginTrans()
       SQL = "update escCliente set " & VbCrLf & _
            "   ds_senha_acesso = '" & trim(Request("w_ds_senha_acesso")) & "', " & VbCrLf & _
            "   dt_alteracao    = getdate() " & VbCrLf & _
            "   where " & CL
       ExecutaSQL(SQL)

       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=A&CL=" & CL & "';"
       ScriptClose

  Case conWhatCalendario
       dbms.BeginTrans()

       If w_ea = "I" Then
        Dim Item
        For each Item in Request("checkbox")
            ' Recupera o valor da próxima chave primária
            SQL = "select IsNull(max(sq_ocorrencia),0)+1 chave from escCalendario_Cliente" & VbCrLf
            ConectaBD SQL
            w_chave = RS("chave")
            DesconectaBD
          
            ' Insere o arquivo
            SQL = " insert into escCalendario_Cliente " & VbCrLf & _
                "    (sq_ocorrencia, sq_site_cliente, dt_ocorrencia, ds_ocorrencia, sq_tipo_data, sq_particular_calendario) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf & _
                "     " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "     convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_ocorrencia"),2)) & "',103), " & VbCrLf & _
                "     '" & Request("w_ds_ocorrencia") & "', " & Request("w_tipo") & ", " & Item & VbCrLf & _
                " )" & VbCrLf
            ExecutaSQL(SQL)
                              
            ' Grava o acesso na tabela de log
            w_sql = SQL
            SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         1, " & VbCrLf & _
                "         'Inclusão de data no calendário da escola.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
            ExecutaSQL(SQL)                            
        Next     
        
       ElseIf w_ea = "A" Then

       
          SQL = " update escCalendario_Cliente set " & VbCrLf & _
                "     dt_ocorrencia            = convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_ocorrencia"),2)) & "',103), " & VbCrLf & _
                "     ds_ocorrencia            = '" & Request("w_ds_ocorrencia") & "', " & VbCrLf & _
                "     sq_tipo_data             = "  & Request("w_tipo") & VbCrLf & _
                "where sq_ocorrencia = " & Request("w_chave") & VbCrLf          
          ExecutaSQL(SQL)

          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         2, " & VbCrLf & _
                "         'Alteração de data no calendário da escola.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "E" Then
          SQL = " delete escCalendario_Cliente where sq_ocorrencia = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         3, " & VbCrLf & _
                "         'Exclusão de data no calendário da escola.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       End If

       SQL = "update escCliente set dt_alteracao = getdate() where " & CL & VbCrLf
       ExecutaSQL(SQL)
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
Case "CALS"
       dbms.BeginTrans()

       If w_ea = "I" Then
          ' Recupera o valor da próxima chave primária
          SQL = "select IsNull(max(sq_particular_calendario),0)+1 chave from escParticular_Calendario" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD

          ' Insere o arquivo
          SQL = " insert into escParticular_Calendario " & VbCrLf & _
                "    (sq_particular_calendario, sq_cliente, nome, ativo, ordem, observacao) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf & _
                "     " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "     '" & Request("w_nome") & "' ," & VbCrLf & _
                "     '" & UCase(Request("w_ativo")) & "', " & Request("w_ordem") & ", " &  VbCrLf & _
                "     '" & Request("w_obs") &  "'" & VbCrLf & _
                " )" & VbCrLf
          ExecutaSQL(SQL)

       ElseIf w_ea = "A" Then
          SQL = " update escParticular_Calendario set " & VbCrLf & _
                "  nome       = '" & Request("w_nome") & "', " & VbCrLf & _
                "  ativo      = '" & UCase(Request("w_ativo")) & "', " & VbCrLf & _
                "  ordem      = "  & Request("w_ordem") &", " & VbCrLf & _   
                "  observacao = '"  & Request("w_obs") & "'" & VbCrLf & _
                "  where sq_particular_calendario = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

       ElseIf w_ea = "E" Then
          SQL = " select count(*) as registros from escCalendario_Cliente where sq_particular_calendario = " & Request("w_chave")
          ConectaBD SQL
          If RS("registros") = 0 Then          
            SQL = " delete escParticular_Calendario where sq_particular_calendario = " & Request("w_chave") & VbCrLf
            ExecutaSQL(SQL)
          Else
            %>
            <script>alert('Existem datas atribuídas a este calendário.\nExclua a data e em seguida, o calendário.');</script>            
            <%
          End If
       End If

       SQL = "update escCliente set dt_alteracao = getdate() where " & CL & VbCrLf
       ExecutaSQL(SQL)
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
End Select    
End Sub

REM =========================================================================
REM Tela de exibicao de cursos
REM -------------------------------------------------------------------------
Sub ShowCurso
  Dim sql, strNome, wCont, wRede, wAno
  
  wAno =  Request.QueryString("wAno")
  
  If wAno = "" Then
     wAno = Year(Date())
  End If

    sql = "SELECT b.ds_curso,  " & VbCrLf & _
           "       case a.categoria " & VbCrLf & _ 
           "          when 1 then 'Educação Profissional Técnica de Nível Médio' " & VbCrLf & _ 
           "          else 'Sem informação' " & VbCrLf & _ 
           "       end as nm_categoria " & VbCrLf & _ 
           "  FROM escParticular_curso a " & VbCrLf & _ 
           "       inner join escCurso b on (a.sq_curso = b.sq_curso)" & VbCrLf & _ 
           " WHERE a." & CL
           
    ConectaBD sql  
    
    Cabecalho
    ShowHTML "<head>"
    ShowHTML "   <meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> "
    ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
    ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
    ShowHTML "</head>"
    ShowHTML "<BASE HREF=""" & conSite & "/"">"
    ShowHTML "<body>"
     wCont = 0
     If Not RS.EOF Then
        ShowHTML "<p align=""justify"">Cursos profissionalizantes oferecidos pela instituição de ensino:</p>"

        ShowHTML "<TABLE width=""100%"" border=0 cellSpacing=3>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD width=""2%"" >&nbsp;</TD>"
        ShowHTML "    <TD width=""45%""><b>Curso</TD>"
        ShowHTML "    <TD width=""48%""><b>Categoria</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
        ShowHTML "    <TD COLSPAN=""6"" HEIGHT=""1"" BGCOLOR=""#DAEABD""></TD>"
        ShowHTML "  </TR>"
        wCont = 1
        Do While Not RS.EOF
           ShowHTML "  <TR valign=""top"">"
           ShowHTML "    <TD>&nbsp;</TD>"
           ShowHTML "    <TD>" & RS("ds_curso") & "</TD>"
           ShowHTML "    <TD>" & RS("nm_categoria") & "</TD>"
           ShowHTML "  </TR>"
           wCont = wCont + 1
           RS.MoveNext
        Loop
        
     Else
        ShowHTML "<p align=""justify"">Nenhum curso profissionalizante registrado para esta instituição de ensino:</p>"
     End If
     RS.Close    
     ShowHTML "    </TABLE>"
 
End Sub

REM =========================================================================
REM Tela de dados cadastrais
REM -------------------------------------------------------------------------
Sub ShowCadastro

' Novo Bloco de Dados


  Dim sql, strNome, authabsecretario, zona
       
  sql = "SELECT b.ds_cliente, a.cnpj_escola, a.mantenedora, a.mantenedora_endereco, a.cnpj_executora, a.codinep, a.diretor, a.secretario, a.aut_hab_secretario, " & VbCrLf & _
        "       coalesce(convert(varchar,a.vencimento,103),'Sem informação') as vencimento, " & VbCrLf & _
        "       a.endereco, b.no_municipio, b.sg_uf,  " & VbCrLf & _
        "       '(' + substring(telefone_1,0,3)+')'+substring(telefone_1,3,110) as telefone_1, " & VbCrLf & _
        "       '(' + substring(telefone_2,0,3)+')'+substring(telefone_2,3,110) as telefone_2, " & VbCrLf & _
        "       a.email_1, a.email_2, a.cep, " & VbCrLf & _
        "       '(' + substring(fax,0,3)+')'+substring(fax,3,110) as fax, " & VbCrLf & _ 
        "       b.localizacao, c.no_regiao, a.url " & VbCrLf & _ 
        "FROM esccliente_particular                 a " & VbCrLf & _
        "     INNER   JOIN escCliente               b on (a.sq_cliente   = b.sq_cliente) " & VbCrLf & _
        "       INNER JOIN escRegiao_Administrativa c on (b.sq_regiao_adm = c.sq_regiao_adm) " & VbCrLf & _
        " where a." & CL
  
  ConectaBD SQL

  
'Fim do Novo Bloco

    Cabecalho
    ShowHTML "<head>"
    ShowHTML "   <title>" & RS("DS_CLIENTE") & "</title>"
    ShowHTML "   <meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> "
    ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
    ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
    ShowHTML "</head>"
    ShowHTML "<BASE HREF=""" & conSite & "/"">"
    ShowHTML "<body>"
  
  If Not RS.EOF Then
    Do While Not RS.EOF
        ShowHTML "<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""2"">"
        ShowHTML "<tr id=""trheader""><td colspan=""2"" width=""100%"">" & uCase("Identificação:") & "</td></tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td align=""right"" width=""30%"" nowrap=""nowrap""><b>Nome da Instiuição:"
        ShowHTML "    <td align=""left"" width=""70%"" >" & RS("ds_cliente") 
        ShowHTML "  </tr>"
        'ShowHTML "  <tr valign=""top"">"
        'ShowHTML "    <td width=""10%"">"
        'ShowHTML "    <td align=""right"" width=""20%"" nowrap=""nowrap""><b>CNPJ da Instituição:"
        'ShowHTML "    <td align=""left"" width=""70%"" >" & Nvl(RS("cnpj_escola"),"Sem informação") 
        'ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td align=""right"" width=""30%"" nowrap=""nowrap""><b>Mantenedora:"
        ShowHTML "    <td align=""left"" width=""70%"" >" & Nvl(RS("mantenedora"),"Sem informação") 
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td align=""right"" width=""30%""><b>Diretor(a):"
        ShowHTML "    <td align=""left"" width=""70%"">" & Nvl(RS("diretor"),"Sem informação") 
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td align=""right"" width=""30%""><b>Secretário(a):"
        ShowHTML "    <td align=""left"" width=""70%"">" & Nvl(RS("secretario"),"Sem informação")
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td align=""right"" width=""30%""><b>Situação Secretário(a):"
        Select Case cint(RS("aut_hab_secretario"))
            case 0 authabsecretario = "Sem Informação"
            Case 1 authabsecretario = "Autorizado"
            Case 2 authabsecretario = "Habilitado"
            Case 3 authabsecretario = "Irregular"            
        End Select
        ShowHTML "    <td align=""left"" width=""70%"">" & authabsecretario & "</td>"
        ShowHTML "  </tr>"
        
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td align=""right"" width=""30%"" nowrap=""nowrap""><b>Código INEP:"
        ShowHTML "    <td align=""left"" width=""70%"" >" & RS("codinep") 
        ShowHTML "  </tr>"       

        ShowHTML "  <tr>"
        ShowHTML "    <td align=""right"" nowrap=""nowrap"" width=""30%"" ><b>Validade credenciamento:</b>"
        ShowHTML "    <td align=""left"" width=""70%"">" & RS("vencimento") & "</td>"
        ShowHTML "  </tr>"
        ShowHTML "  </TABLE>"
        
        ShowHTML "<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""2"">"
        ShowHTML "<tr id=""trheader""><td colspan=""2"" width=""100%"">" & uCase("LOCALIZAÇÃO E CONTATOS:") & "</td></tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""30%"" align=""right""><b>Zona:"
        Select Case cint(RS("localizacao"))
            case 0  zona = "Urbana"
            case 1  zona = "Rural"
            default zona = "Sem Informação"
        End Select
        ShowHTML "    <td width=""70%"" align=""left"">" & zona & "</td>"
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""30%"" align=""right""><b>Endereço:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("endereco")' & " - " & RS("cep")
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""30%"" align=""right""><b>Endereço da Mantenedora:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("mantenedora_endereco")' & " - " & RS("cep")
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""30%"" align=""right""><b>CEP:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("cep") & "</td>"
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""30%"" align=""right""><b>Região Administrativa:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_regiao")
        ShowHTML "  </tr>"
        ShowHTML "  <tr>"
        ShowHTML "    <td width=""30%"" align=""right"" valign=""top""><b>Telefones:"
        If(Nvl(RS("telefone_1"),"") <> "" and Nvl(RS("telefone_2"),"") <> "") Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("telefone_1") & " / " & RS("telefone_2")
        ElseIf Nvl(RS("telefone_1"),"") <> "" and Nvl(RS("telefone_2"),"")  = "" Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("telefone_1")
        ElseIf(Nvl(RS("telefone_2"),"") <> "" and Nvl(RS("telefone_1"),"")  = "") Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("telefone_2")
        Else
          ShowHTML "    <td width=""70%"" align=""left"">Sem informação"
        End If 
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""30%"" align=""right""><b>Fax:"
        ShowHTML "    <td width=""70%"" align=""left"">" & nvl(RS("fax"),"Sem informação")
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""30%"" align=""right""><b>E-mails:"
        If Nvl(RS("email_1"),"")  <> "" and Nvl(RS("email_2"),"") <> "" Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("email_1") & " / " & RS("email_2")
        ElseIf(Nvl(RS("email_1"),"")  <> "" and Nvl(RS("email_2"),"")  = "") Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("email_1")
        ElseIf(Nvl(RS("email_2"),"")  <> "" and Nvl(RS("email_1"),"")  = "") Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("email_2")
        Else
          ShowHTML "    <td width=""70%"" align=""left"">Sem informação"
        End If
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""30%"" align=""right""><b>Site:"
        If Nvl(RS("url"),"") <> 0 Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("url")
        Else
          ShowHTML "    <td width=""70%"" align=""left"">Sem informação"
        End If   
        ShowHTML "  </tr>"  

  ShowHTML "  </TABLE>"
  
  RS.MoveNext

  Loop

  End If
  
  If wAno = "" Then
     wAno = Year(Date())
  End If

     sql = "SELECT INFANTIL, FUNDAMENTAL, EJA, MEDIO, DISTANCIA, PROFISSIONAL " &_
           "FROM esccliente_particular a " & _
           "WHERE a." & CL
     dim modalidade
     
        

     ConectaBD sql
     wCont = 0     
     
     If Not RS.EOF Then
        Select Case cint(RS("INFANTIL"))
            Case 0 modalidadeinf = "Não oferece"
            Case 1 modalidadeinf = "Creche"
            Case 2 modalidadeinf = "Pré-Escola"
            Case 3 modalidadeinf = "Creche e Pré-Escola"
        End Select
        
        Select Case cint(RS("FUNDAMENTAL"))
            Case 0  modalidadefund = "Não oferece"            
            Case 1  modalidadefund = "1º ao 9º ano"            
            Case 2  modalidadefund = "1º ao 9º ano e 1ª a 4ª	"            
            Case 3  modalidadefund = "1º ao 9º ano e 1ª a 8ª"            
            Case 4  modalidadefund = "1º ao 5º ano e 1ª a 4ª"            
            Case 5  modalidadefund = "1º ao 5º ano e 1ª a 8ª"            
            Case 6  modalidadefund = "1º ao 5º ano"            
            Case 7  modalidadefund = "3ª e 4ª c/ext. prog. e 1º ao 9º ano"            
            Case 8  modalidadefund = "3ª e 4ª c/ext. prog. e 1º ao 5º ano"            
            Case 9  modalidadefund = "1ª a 4ª Séries"            
            Case 10 modalidadefund = "1ª a 8ª Séries"            
            Case 11 modalidadefund = "5ª a 8ª Séries"            
        End Select
        
        Select Case cint(RS("MEDIO"))
            Case 0 modalidademed = "Não oferece"            
            Case 1  modalidademed = "Oferece"
        End Select
        
        Select Case cint(RS("EJA"))
            Case 0 modalidadeeja = "Não oferece"            
            Case 1 modalidadeeja = "1º e 2º Segmento"            
            Case 2 modalidadeeja = "1º e 3º Segmento"
            Case 3 modalidadeeja = "2º e 3º Segmento"                
            Case 4 modalidadeeja = "1º,2º e 3º Segmento"
            Case 5 modalidadeeja = "2º Segmento"
            Case 6 modalidadeeja = "3º Segmento"
            Case 7 modalidadeeja = "1º Segmento"
        End Select
        
        Select Case cint(RS("DISTANCIA"))
            Case 0 modalidadedist = "Não oferece"            
            Case 1 modalidadedist = "Oferece"
        End Select
        Select Case cint(RS("PROFISSIONAL"))
            Case 0 modalidadeprof = "Não oferece"            
            Case 1 modalidadeprof = "Oferece"
        End Select

        
        wCont = 1
         
        Do While Not RS.EOF
        
        ShowHTML "<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""2"">"
        ShowHTML "<tr id=""trheader""><td colspan=""2"" width=""100%"">" & uCase("Ofertas da instituição de ensino:") & "</td></tr>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD align=""right"" width=""30%""><b>Educação Infantil: </b></TD>"
        ShowHTML "    <TD align ""left"" width=""70%"">" & modalidadeinf & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD align=""right""><b>Ensino Fundamental: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadefund & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD align=""right""><b>Ensino Médio: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidademed & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD align=""right""><b>Educação de Jovens e Adultos: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadeeja & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD align=""right""><b>Educação à Distância: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadedist & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD align=""right""><b>Educação Profissional: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadeprof & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "</TABLE>"
        wCont = wCont + 1
        RS.MoveNext
        Loop
        
     Else
        ShowHTML "    <p><BR>Até o presente momento, a Instituição não oferece cursos profissionalizantes. </p>"
     End If
  
End Sub

REM =========================================================================
REM Tela de dados básicos
REM -------------------------------------------------------------------------
Sub GetCliente

  Dim w_sq_cliente, w_ds_senha_acesso, w_ds_email, w_ds_apelido

  If w_ea = "A" Then
     SQL = "select * from escCliente where " & CL
     ConectaBD SQL
     w_sq_cliente       = RS("sq_cliente")
     w_ds_senha_acesso  = RS("ds_senha_acesso")
     w_ds_email         = RS("ds_email")
     w_ds_apelido       = RS("ds_apelido")
     
  End If
  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "Javascript"
  ValidateOpen "Validacao"
  If w_ea = "A" Then
     Validate "w_ds_senha_acesso", "Senha de acesso", "1", "1", "4", "14", "1", "1"
     Validate "w_ds_email", "e-Mail", "1", "", "6", "60", "1", "1"
     'Validate "w_ds_apelido", "Apelido", "1", "", "6", "30", "1", "1"
  End If
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.Form.w_ds_senha_acesso.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Atualização de dados básicos</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
  ShowHTML MontaFiltro("POST")
  ShowHTML "<input type=""hidden"" name=""R"" value=""" & w_ew & """>"

  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Dados básicos da Unidade</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Os dados deste bloco são utilizados para acesso à página de atualização do site, para envio de mensagens ao responsável técnico e para pesquisa por apelido no mecanismo de busca.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>S</u>enha de acesso:</b><br><INPUT ACCESSKEY=""S"" " & w_Disabled & " class=""STI"" type=""password"" name=""w_ds_senha_acesso"" size=""14"" maxlength=""14"" value=""" & w_ds_senha_acesso & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a senha desejada para acessar a tela de atualização dos dados do site.','white')""; ONMOUSEOUT=""kill()""></td>"
  'ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>e</u>-Mail:</b><br><INPUT ACCESSKEY=""E"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_email"" size=""60"" maxlength=""60"" value=""" & w_ds_email & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o e-mail de sua unidade para contatos com a equipe técnica. Este e-mail não será divulgado no site.','white')""; ONMOUSEOUT=""kill()""></td>"
  'ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>A</u>pelido:</b><br><INPUT ACCESSKEY=""A"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_apelido"" size=""30"" maxlength=""30"" value=""" & w_ds_apelido & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o apelido pelo qual sua unidade é mais conhecida.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000""></TD></TR>"
  ' Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
  ShowHTML "      <tr><td align=""center"" colspan=""3"">"
  ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Gravar"">"
  ShowHTML "          </td>"
  ShowHTML "      </tr>"
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  ShowHTML "</FORM>"
  ShowHTML "</table>"
  Rodape

  Set w_sq_cliente      = Nothing
  Set w_ds_senha_acesso = Nothing
  Set w_ds_email        = Nothing
  Set w_ds_apelido      = Nothing
      
End Sub

REM =========================================================================
REM Tela de portarias
REM -------------------------------------------------------------------------
Sub ShowPortaria
  Dim sql, strNome, wCont, wRede, wAno
  
    sql = "SELECT numero, data, dodf, dodf_pagina, dodf_data, observacao FROM escParticular_Portaria a WHERE a." & CL
    ConectaBD sql
    Cabecalho
    ShowHTML "<head>"
    ShowHTML "   <meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> "
    ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
    ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
    ShowHTML "</head>"
    ShowHTML "<BASE HREF=""" & conSite & "/"">"
    ShowHTML "<body>"
    
    If Not RS.EOF Then
        ShowHTML "<p align=""justify"">Informações sobre Portarias :</p>"
        ShowHTML "<TABLE width=""100%"" cellSpacing=""0"" cellpadding=""0"">"
        ShowHTML "  <TR id=""trheader"" valign=""center"">"
        ShowHTML "    <TD width=""2%"" >&nbsp;</TD>"
        ShowHTML "    <TD width=""5%""><b>Número</TD>"
        ShowHTML "    <TD width=""10%""><b>Data</TD>"
        ShowHTML "    <TD width=""10%""><b>DODF</TD>"
        ShowHTML "    <TD width=""15%""><b>Página DODF</TD>"
        ShowHTML "    <TD width=""10%""><b>Data DODF</TD>"
        ShowHTML "    <TD width=""40%""><b>Observação</TD>"
        ShowHTML "  </TR>"
        wCont = 1
        Do While Not RS.EOF
           ShowHTML "  <TR valign=""top"">"
           ShowHTML "    <TD>&nbsp;</TD>"
           ShowHTML "    <TD>" & RS("numero") & "</TD>"
           ShowHTML "    <TD>" & RS("data") & "</TD>"
           ShowHTML "    <TD>" & RS("dodf") & "</TD>"
           ShowHTML "    <TD>" & RS("dodf_pagina") & "</TD>"
           ShowHTML "    <TD>" & RS("dodf_data") & "</TD>"
           ShowHTML "    <TD>" & RS("observacao") & "</TD>"
           ShowHTML "  </TR>"
           wCont = wCont + 1
           RS.MoveNext
        Loop
        
     Else
        ShowHTML "<p align=""justify"">Nenhum curso profissionalizante registrado para esta instituição de ensino:</p>"
     End If
     DesconectaBD    
     ShowHTML "    </TABLE>"
End Sub

REM =========================================================================
REM Tela de ordens de serviço
REM -------------------------------------------------------------------------
Sub ShowOS


  Dim sql, strNome, wCont, wRede, wAno
  
    sql = "select  numero, data, dodf, dodf_pagina, dodf_data, observacao FROM escparticular_os a WHERE a." & CL
    ConectaBD sql
    Cabecalho
    ShowHTML "<head>"
    ShowHTML "   <meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> "
    ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
    ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
    ShowHTML "</head>"
    ShowHTML "<BASE HREF=""" & conSite & "/"">"
    ShowHTML "<body>"
    
    If Not RS.EOF Then
        ShowHTML "<p align=""justify"">Informações sobre as Ordens de Serviço da Instituição:</p>"
        ShowHTML "<TABLE id=""OrdemServico"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
        ShowHTML "  <TR id=""trheader"" valign=""center"">"
        ShowHTML "    <TD width=""2%"">&nbsp;</TD>"
        ShowHTML "    <TD width=""5%""><b>Numero</TD>"
        ShowHTML "    <TD width=""10%""><b>Data</TD>"
        ShowHTML "    <TD width=""10%""><b>DODF</TD>"
        ShowHTML "    <TD width=""10%""><b>Página DODF</TD>"
        ShowHTML "    <TD width=""10%""><b>Data DODF</TD>"
        ShowHTML "    <TD width=""40%""><b>Observação</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
'        ShowHTML "    <TD COLSPAN=""6"" HEIGHT=""1"" BGCOLOR=""#DAEABD""></TD>"
        ShowHTML "  </TR>"
        wCont = 1
        Do While Not RS.EOF
           ShowHTML "  <TR valign=""top"">"
           ShowHTML "    <TD>&nbsp;</TD>"
           ShowHTML "    <TD>" & RS("numero") & "</TD>"
           ShowHTML "    <TD>" & RS("data") & "</TD>"
           ShowHTML "    <TD>" & RS("dodf") & "</TD>"
           ShowHTML "    <TD>" & RS("dodf_pagina") & "</TD>"
           ShowHTML "    <TD>" & RS("dodf_data") & "</TD>"
           ShowHTML "    <TD>" & RS("observacao") & "</TD>"
           ShowHTML "  </TR>"
           wCont = wCont + 1
           RS.MoveNext
        Loop
        
     Else
        ShowHTML "<p align=""justify"">Nenhum curso profissionalizante registrado para esta instituição de ensino:</p>"
     End If
     DesconectaBD 
     ShowHTML "    </TABLE>"
End Sub


Public Sub FormCal
  Dim sql, strNome, wCont, wRede, wAno, wDatas(31,12,10), wCores(31,12,10), wImagem(31,12,10)
  Dim w_ini, w_fim, w_dias, w_inicial, w_final, w_mes, w_fim1, w_ini2
  Dim w_cliente
  
  w_cliente = replace(CL,"sq_cliente=","")
  wAno = Request("w_ano")
  
   ShowHTML "<HTML>"
   ShowHTML "<HEAD>"
   ShowHTML "<TITLE>" & "Controle Central" & "</TITLE>"
   ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
   ShowHTML "   <link href=""/css/impressao.css"" media=""print"" rel=""stylesheet"" type=""text/css"" />"
   ShowHTML "<style>"
   ShowHTML "<// a { color: ""#000000""; text-decoration: ""none""; } "
   ShowHTML "    a:hover { color:""#000000""; text-decoration: ""underline""; }"
   ShowHTML "    .SS{text-decoration:none;} "
   ShowHTML "    .SS:HOVER{text-decoration: ""underline"";} "
   ShowHTML "//></style>"
   ShowHTML "</HEAD>"
   ShowHTML "<BASEFONT FACE=""Verdana, Helvetica, Sans-Serif"" SIZE=""2"">"
   
   
        Response.Write "<BODY topmargin=0 bgcolor=""#FFFFFF"" BGPROPERTIES=""FIXED"" text=""#000000"" link=""#000000"" vlink=""#000000"" alink=""#FF0000""></body>"        
        ShowHTML "<div id=""campopesquisa"">"
        ShowHTML "<H3>" & Session("nome_cliente") & "</H3>"
        ShowHTML "<HR>"
        ShowHTML "<br/><br/><p><font size=""1""><b>Por favor, insira os critérios da pesquisa:</b></p>"

        AbreForm "Form", w_Pagina & "FormCal", "POST", null, null
        ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & w_ew & """>"
        ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
        ShowHTML "<INPUT type=""hidden"" name=""w_chave"" value=""" & w_chave & """>"
        ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente"" value=""" & replace(CL,"sq_cliente=","") & """>"
        ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"
       
            SQL = "select distinct(year(dt_ocorrencia)) as ano from escCalendario_Cliente where " & replace(CL,"sq_cliente","sq_site_cliente") & "order by year(dt_ocorrencia) desc"
            ConectaBD SQL
            ShowHTML" <select name=""w_ano"" onChange=""document.Form.w_calendario.selectedIndex = 0; document.Form.submit();""> "
            ShowHTML" <option value="""">Selecione o ano...</option>"
            while not RS.Eof                
                If trim(Request("w_ano")) = trim(RS("ano")) Then
                    ShowHTML" <option selected value=""" &  RS("ano") & """> " & RS("ano") & "</option>"
                else
                    ShowHTML" <option value=""" & RS("ano") & """> " & RS("ano") & "</option>"
                end if
                RS.MoveNext            
            wend
            ShowHTML" </select> "       
            DesconectaBD
        If 1=1 then'Request("w_ano") <> "" Then           
            SQL = "select distinct(b.nome), a.sq_particular_calendario, b.ordem from escCalendario_Cliente a left join escParticular_Calendario b on (a.sq_particular_calendario = b.sq_particular_calendario) where " & Replace(CL, "sq_cliente", "sq_site_cliente") & " and year(dt_ocorrencia) = " & nvl(Request("w_ano"),0) & "order by ordem"
            ConectaBD SQL
            ShowHTML" <select name=""w_calendario""> "
            ShowHTML" <option value=""0"">Selecione o calendário...</option>"
            while not RS.Eof
                If trim(Request("w_calendario")) = trim(RS("sq_particular_calendario")) Then            
                    ShowHTML" <option selected value="""& RS("sq_particular_calendario") &"""> " & RS("nome") & "</option>"
                Else
                    ShowHTML" <option value="""& RS("sq_particular_calendario") &"""> " & RS("nome") & "</option>"
                End If
                RS.MoveNext            
            wend
            ShowHTML" </select> "
            DesconectaBD 
        end if    
        If Request("w_ano") <> "" Then                   
            ShowHTML "<INPUT type=""hidden"" name=""w_print"" value=""s"">"
            ShowHTML" <input type=submit name=""showCalendar"" value=""Exibir Calendário"">"
            If Request("w_print") = "s" Then
                ShowHTML "<div id=""impressora""><TABLE align=""center"" width=""100%"" border=0><TR align=""right""><TD> <img src=""img/impressora.gif"" onclick=""window.print();""/></TD></TR></TABLE></div>"
            End If     
        End If
        ShowHTML" </form>"
        ShowHTML" </div>"
        
    if (nvl(trim(Request("w_calendario")),"0") = "0") then 
        Response.End
    end if
   Response.Write "<BODY topmargin=0 bgcolor=""#FFFFFF"" BACKGROUND=""img/background.gif"" BGPROPERTIES=""FIXED"" text=""#000000"" link=""#000000"" vlink=""#000000"" alink=""#FF0000""> "   
   ShowHTML "<TABLE align=""center"" width=""100%"" border=0>"
   ShowHTML "<tr><td align=""center""><TABLE align=""center"" border=0 cellSpacing=0>"

  sql = "SELECT b.ds_cliente, a.cnpj_escola, a.mantenedora, a.mantenedora_endereco, a.cnpj_executora, a.codinep, a.diretor, a.secretario, a.aut_hab_secretario, " & VbCrLf & _
        "       coalesce(convert(varchar,a.vencimento,103),'Sem informação') as vencimento, " & VbCrLf & _
        "       a.endereco, b.no_municipio, b.sg_uf,  " & VbCrLf & _
        "       '(' + substring(telefone_1,0,3)+')'+substring(telefone_1,3,110) as telefone_1, " & VbCrLf & _
        "       '(' + substring(telefone_2,0,3)+')'+substring(telefone_2,3,110) as telefone_2, " & VbCrLf & _
        "       a.email_1, a.email_2, a.cep, " & VbCrLf & _
        "       '(' + substring(fax,0,3)+')'+substring(fax,3,110) as fax, " & VbCrLf & _ 
        "       b.localizacao, c.no_regiao, a.url, d.nome as nomecal, d.observacao " & VbCrLf & _ 
        "FROM esccliente_particular                 a " & VbCrLf & _
        "     INNER   JOIN escCliente               b on (a.sq_cliente   = b.sq_cliente) " & VbCrLf & _
        "       INNER JOIN escRegiao_Administrativa c on (b.sq_regiao_adm = c.sq_regiao_adm) " & VbCrLf & _
        "       INNER JOIN escParticular_Calendario d on (b.sq_cliente = d.sq_cliente) " & VbCrLf & _
        " where a." & CL & " AND d.sq_particular_calendario = " & Request("w_calendario")
  
  ConectaBD SQL
  
  Dim wObs
  wObs = RS("observacao")
   

  ShowHTML "<tr><td align=""center""><font size=""2""><b>"
  ShowHTML "  <p><font size=3>"  & RS("ds_cliente") & "</font></p>"
  ShowHTML "  <p>"  & RS("endereco")
  ShowHTML "  <p>"  & RS("mantenedora")
  ShowHTML "  <p>Credenciamento: " & RS("vencimento")
  ShowHTML "  <p>Oferta: " & RS("nomecal")  
  ShowHTML "  <p><font size=2>CALENDÁRIO ESCOLAR " & Request("w_ano") & "</font></p>"
  
  DesconectaBD

  sql = "SELECT '' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'B' AS origem FROM escCalendario_base a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE YEAR(dt_ocorrencia)=" & Request("w_ano") & " " & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT '#99CCFF' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE sq_site_cliente = " & w_cliente & "  AND YEAR(dt_ocorrencia)= " & Request("w_ano") & " AND sq_particular_calendario = " & Request("w_calendario") & VbCrLf & _
        "ORDER BY data, origem desc, ocorrencia" & VbCrLf 
  ConectaBD sql
  
  If sstrIN = 0 then
  
     If Not RS.EOF Then
        Do While Not RS.EOF
           If RS("origem") = "E" then
              wDatas(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("ocorrencia") & " (Origem: Escola)"
              wImagem(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2))= RS("imagem")
              wCores(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("cor")
           Else
              wDatas(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("ocorrencia") & " (Origem: Oficial)"
              wImagem(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("imagem")
           End If
           DataOcc = RS("data")
           RS.MoveNext                      
        Loop        
        RS.MoveFirst
     End If
     ShowHTML "<tr><td><TABLE align=""center"" border=1>"
     ShowHTML "<tr><td><TABLE border=0 cellSpacing=0>"
     ShowHTML "<tr valign=""top"">"     
     ShowHTML "  <td>" & MontaCalendario("01" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("02" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("03" & wAno, wDatas, wCores, wImagem)
     ShowHTML "<tr valign=""top"">"
     ShowHTML "  <td>" & MontaCalendario("04" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("05" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("06" & wAno, wDatas, wCores, wImagem)
     ShowHTML "<tr valign=""top"">"
     ShowHTML "  <td>" & MontaCalendario("07" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("08" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("09" & wAno, wDatas, wCores, wImagem)
     ShowHTML "<tr valign=""top"">"
     ShowHTML "  <td>" & MontaCalendario("10" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("11" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("12" & wAno, wDatas, wCores, wImagem)
     ShowHTML "</table>"
     ShowHTML "  <td>"

     ' Recupera o ano letivo e o período de recesso
     sql = "  select * from " & VbCrLf & _
           "        (select dt_ocorrencia w_let_ini " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  Request("w_ano") &  " and b.sigla = 'IA') " & VbCrLf & _
           "          where sq_site_cliente = " & w_cliente & VbCrLf & _
           "        ) a, " & VbCrLf & _
           "        (select dt_ocorrencia w_let_fim " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  Request("w_ano") &  " and b.sigla = 'TA') " & VbCrLf & _
           "          where sq_site_cliente = " & w_cliente & VbCrLf & _
           "        ) b, " & VbCrLf & _
           "        (select dt_ocorrencia w_let2_ini " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  Request("w_ano") &  " and b.sigla = 'I2') " & VbCrLf & _
           "          where sq_site_cliente = " & w_cliente & VbCrLf & _
           "        ) c, " & VbCrLf & _
           "        (select dt_ocorrencia w_let1_fim " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  Request("w_ano") &  " and b.sigla = 'T1') " & VbCrLf & _
           "          where sq_site_cliente = " & w_cliente & VbCrLf & _
           "        ) d " & VbCrLf
     ConectaBD sql
     Do While Not RS.EOF
       w_fim1 = RS("w_let1_fim")
       w_ini2 = RS("w_let2_ini")
       RS.MoveNext
     Loop
     DesconectaBD

     For w_cont = 1 to 2
        If w_cont = 1 Then
           w_ini = 1
           w_fim = 7
        Else
           w_ini = 7
           w_fim = 12
           ShowHTML "  <br><br>"
        End If
        w_dias = 0
        ShowHTML "  <TABLE align=""center"" border=1 cellSpacing=1>"
        ShowHTML "  <TR ALIGN=""CENTER""><TD COLSPAN=""4""><b>" & w_cont & "º Semestre</b></td></tr>"
        ShowHTML "  <TR valign=""top"" ALIGN=""CENTER"">"
        ShowHTML "    <TD><b>MÊS"
        ShowHTML "    <TD><b>DL"
        ShowHTML "    <TD><b>SLE"
        ShowHTML "    <TD><b>TOT"
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
        ShowHTML "    <TD COLSPAN=""2"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
        ShowHTML "  </TR>"
        For wCont = w_ini to w_fim
           If month(w_ini2) = wCont and w_cont = 2Then
              w_inicial = formataDataEdicao(w_ini2)
           Else
              w_inicial = "01/" & mid(100+wCont,2,2) & "/" & wAno
           End If
           If month(w_fim1) = wCont and w_cont = 1 Then
              w_final = formataDataEdicao(w_fim1)
           Else
              sql = "SELECT convert(varchar,dbo.lastDay(convert(datetime,'01/" & mid(100+wCont,2,2) & "/" & wAno & "' ,103)),103) fim" & VbCrLf 
              ConectaBD SQL
              w_final = RS("fim")
              DesconectaBD
           End If
           
           sql = "SELECT dbo.diasLetivos('" & w_inicial & "', '" & w_final & "'," & w_cliente & ") qtd" & VbCrLf 
           ConectaBD SQL
           If RS("qtd") > 0 Then
              ShowHTML "  <TR>"
              ShowHTML "    <TD>" & nomeMes(wCont)
              ShowHTML "    <TD ALIGN=""CENTER"">" & RS("qtd")
              w_mes = RS("qtd")

              DesconectaBD
              sql = "SELECT count(*) qtd FROM escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla ='SL') WHERE sq_site_cliente = " & w_cliente & "  AND dt_ocorrencia between convert(datetime, '" & w_inicial & "',103) and convert(datetime, '" & w_final & "',103)" & VbCrLf
              ConectaBD sql
              ShowHTML "    <TD ALIGN=""CENTER"">" & RS("qtd")
              w_mes = w_mes + RS("qtd")
              DesconectaBD
              ShowHTML "    <TD ALIGN=""CENTER"">" & w_mes
              ShowHTML "  </TR>"
              w_dias = w_dias + w_mes
              
           End If
        Next
        ShowHTML "  <TR>"
        ShowHTML "    <TD COLSPAN=""3"" nowrap>Dias Letivos - Parcial</TD>"
        ShowHTML "    <TD ALIGN=""CENTER"">" & w_dias & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "    </TABLE>"
        
        If( w_cont = 1) Then
            w_total1 = w_dias
        else
            w_total2 = w_dias
        end if
     Next
     if(w_total1 and w_total2) Then
        ShowHTML "  <br/><br/>"
        w_total = cInt(w_total1) + cInt(w_total2)
        if(cInt(w_total) < 200)Then
        alerta = "<span style=""font-size:14px; font-decoration: bold; color:#ff3737"">*</span>"
        legenda= "<span style=""font-size:11px;"">O total de dias letivos <br/>é inferior a 200 dias.</span>"
        else
        alerta = ""
        legenda= ""
        End If
        ShowHTML " <TABLE cellspacing=""1"" border=""0"" align=""center""><TR valign=""top""><TD>Total de Dias Letivos</TD><TD><b>" & w_total & alerta & "</b></TD></TR><TR><TD colspan=""2"">" & alerta & legenda & "</TD></TR></TABLE>"
        
     End If
     

     ShowHTML "<tr><td colspan=""2""><TABLE width=""100%"" align=""center"" border=1 cellSpacing=1>"
     ShowHTML "<tr valign=""middle"" align=""center"">"
     ShowHTML "  <td><font size=1><b>LEGENDA</b></font>"
     ShowHTML "  <td><font size=1><b>FERIADOS</b></font>"
     ShowHTML "  <td><font size=1><b>RECESSOS</b></font>"
     ShowHTML "  <td><font size=1><b>SÁBADOS<br>LETIVOS<br>ESPECIAIS</b></font>"
     ShowHTML "</tr>"
     ShowHTML "<tr valign=""top"">"
     'Legenda
     sql = "SELECT * from escTipo_Data where abrangencia <> 'P' order by nome " & VbCrLf
     ConectaBD sql
     If Not RS.EOF Then
        ShowHTML "  <td><TABLE width=""90%"" align=""center"" border=0 cellSpacing=1>"
        wCont = 1
        Do While Not RS.EOF
           ShowHTML "  <TR VALIGN=""TOP"">"
           ShowHTML "    <TD>&nbsp;"
           ShowHTML "    <TD><img src=""/img/" & RS("imagem") & """ align=""center"">"
           ShowHTML "    <TD>" & RS("nome")
           ShowHTML "  </TR>"
           RS.MoveNext
        Loop
        ShowHTML "    </TABLE>"
     Else
        ShowHTML "  <td>Sem informação"
     End If
     DesconectaBD
     'Feriados
     sql = "SELECT '' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'B' AS origem FROM escCalendario_base a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE YEAR(dt_ocorrencia)=" & Request("w_ano") & " AND b.sigla <> 'CN' " & VbCrLf & _
           "ORDER BY data, origem desc, ocorrencia" & VbCrLf
     ConectaBD sql
     If Not RS.EOF Then
        ShowHTML "  <td><TABLE width=""90%"" align=""center"" border=0 cellSpacing=1>"
        wCont = 1
        Do While Not RS.EOF
           ShowHTML "  <TR VALIGN=""TOP"">"
           ShowHTML "    <TD>&nbsp;"
           ShowHTML "    <TD>" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2)
           ShowHTML "    <TD>" & RS("ocorrencia")
           ShowHTML "  </TR>"
           RS.MoveNext
        Loop
        ShowHTML "    </TABLE>"
     Else
        ShowHTML "  <td>Sem informação"
     End If
     DesconectaBD
     'Recesso
     sql = "SELECT '#99CCFF' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla in ('RE','RA')) WHERE sq_site_cliente = " & w_cliente & " AND YEAR(dt_ocorrencia)= " & Request("w_ano") & " AND sq_particular_calendario = " & Request("w_calendario") &  VbCrLf & _
           "ORDER BY data, origem desc, ocorrencia" & VbCrLf
     ConectaBD sql
  
     If Not RS.EOF Then
        ShowHTML "  <td><TABLE width=""90%"" align=""center"" border=0 cellSpacing=1>"
        Do While Not RS.EOF

           ShowHTML "  <TR VALIGN=""TOP"">"
           ShowHTML "    <TD>&nbsp;"
           ShowHTML "    <TD>" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2)
           ShowHTML "    <TD><font color=""#0000FF"">" & RS("ocorrencia")
           ShowHTML "  </TR>"
           RS.MoveNext
        Loop
        ShowHTML "    </TABLE>"
     Else
        ShowHTML "  <td>Sem informação"
     End If
     DesconectaBD
     
     wCont = 0
     sql = "SELECT '#99CCFF' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla ='SL') WHERE sq_site_cliente = " & w_cliente & "  AND YEAR(dt_ocorrencia)= " & Request("w_ano") & " " & VbCrLf & _
           "ORDER BY data, origem desc, ocorrencia" & VbCrLf 
     ConectaBD sql
       
     If Not RS.EOF Then
        ShowHTML "  <td><TABLE width=""90%"" align=""center"" border=0 cellSpacing=1>"
        Do While Not RS.EOF

           ShowHTML "  <TR VALIGN=""TOP"">"
           ShowHTML "    <TD>&nbsp;"
           ShowHTML "    <TD>" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2)
           ShowHTML "  </TR>"
           RS.MoveNext
        Loop
        ShowHTML "    </TABLE>"
     Else
        ShowHTML "  <td>Sem informação"
     End If
     DesconectaBD
     
     ShowHTML " </TABLE>"

  End If
  ShowHTML "</TABLE>"
  ShowHTML "<H2>Observação</H2>"
  ShowHTML "<textarea name=""w_obs"" style=""width:450px; height:200px;"">" & wObs & "</textarea>"
  ShowHTML "</BODY>"
  ShowHTML "</HTML>"
End Sub

 Sub GetArquivos


    Cabecalho
    ShowHTML "<head>"
    ShowHTML "   <meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> "
    ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
    ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
    ShowHTML "</head>"
    ShowHTML "<BASE HREF=""" & conSite & "/"">"
    ShowHTML "<body>"
    ShowHTML("<table style=""border:solid thin #333333;"" border=""0"" cellpadding=""0"" cellspacing=""0"" align=""left"" width=""20%"">")
    ShowHTML("  <tr>")
    ShowHTML("    <td width=""50%"">")
    ShowHTML("      <img style=""border: none"" src=""img/word.gif""/>")
    ShowHTML("    </td>")
    ShowHTML("    <td>")
    ShowHTML("<a style=""font: 14px Verdana; color: #333333; font-weight: bold;"" href=""files/Formulário%20Recadastramento.doc"">Formulário de Recadastramento <span style=""color: #6699CC"">(download)</span></a>")
    ShowHTML("    </td>")
    ShowHTML("  </tr>")
    ShowHTML("</table>")
    ShowHTML "</body>"
    ShowHTML "</html>"
    
    

  
End Sub

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------

Sub GetCalendario
  Dim w_chave, w_dt_ocorrencia, w_ds_ocorrencia, w_ano, w_sle

  Dim w_troca, i, w_texto

  w_Chave           = Request("w_Chave")
  w_troca           = Request("w_troca")

SQL = "select sq_tipo_data from escTipo_Data where sigla = 'SL'"
ConectaBD SQL
w_sle   = RS("sq_tipo_data")
DesconectaBD
     
  If w_troca > "" Then ' Se for recarga da página
     w_dt_ocorrencia      = Request("w_dt_ocorrencia")
     w_ds_ocorrencia      = Request("w_ds_ocorrencia")
     w_tipo               = Request("w_tipo")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     'SQL = "select * from escCalendario_Cliente where " & replace(CL,"sq_cliente","sq_site_cliente") & " order by year(dt_ocorrencia) desc, dt_ocorrencia"
     SQL = "select a.*, b.nome, c.nome as nomecal from escCalendario_Cliente a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) left join escParticular_Calendario c on (a.sq_particular_calendario = c.sq_particular_calendario) where " & replace(CL,"sq_cliente","sq_site_cliente") & " order by year(dt_ocorrencia) desc, c.nome, dt_ocorrencia"
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escCalendario_Cliente where sq_ocorrencia = " & w_chave
     ConectaBD SQL
     w_dt_ocorrencia   = RS("dt_ocorrencia")
     w_ds_ocorrencia   = RS("ds_ocorrencia")
     w_tipo            = RS("sq_tipo_data")
     DesconectaBD
  End If
  

  Cabecalho
  ShowHTML "<HEAD>"
  If InStr("IAEP",O) > 0 Then
     ScriptOpen "JavaScript"
     CheckBranco
     FormataData
     ValidateOpen "Validacao"
     If InStr("IA",O) > 0 Then
        Validate "w_dt_ocorrencia", "data", "DATA", "1", "10", "10", "1", "1"
        Validate "w_ds_ocorrencia", "ocorrência", "", "1", "2", "60", "1", "1"
        ShowHTML "  var w_data, w_data1, w_data2;"
        ShowHTML "  w_data = theForm.w_dt_ocorrencia.value;"
        ShowHTML "  w_data = w_data.substr(3,2) + '/' + w_data.substr(0,2) + '/' + w_data.substr(6,4);"
        ShowHTML "  w_data1  = new Date(Date.parse(w_data));"
        ShowHTML "  if (w_data1.getDay() == 0 && theForm.w_tipo[theForm.w_tipo.selectedIndex].value != " & w_sle & ") {"
        ShowHTML "     alert('Exceto nos dias de Sábados/Domingos Letivos Especiais,\ndata não pode ser cadastrada em dia de Domingo!');"
        ShowHTML "     theForm.w_dt_ocorrencia.focus();"
        ShowHTML "     return false;"
        ShowHTML "  }"
        ShowHTML "  if (theForm.w_tipo[theForm.w_tipo.selectedIndex].value == " & w_sle & ") {"
        ShowHTML "      if  (w_data1.getDay() != 6 && w_data1.getDay() != 0 ) {"
        ShowHTML "          alert('Sábado/Domingo Letivo Especial só pode ser atribuído a sábados ou domingos!');"
        ShowHTML "          theForm.w_dt_ocorrencia.focus();"
        ShowHTML "          return false;"
        ShowHTML "      }"
        ShowHTML "  }else{"
        ShowHTML "      if  (w_data1.getDay() == 6 || w_data1.getDay() == 0) {"
        ShowHTML "          alert('Em sábados e domingos, só pode ser atribuído Sábado/Domingo Letivo Especial!');"
        ShowHTML "          theForm.w_dt_ocorrencia.focus();"
        ShowHTML "          return false;"
        ShowHTML "      }"        
        ShowHTML "  }"
        
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
  ShowHTML "<H3>" & Session("nome_cliente") & "</H3>"
  ShowHTML "<HR>"
  ShowHTML "<B><FONT COLOR=""#000000"">Datas de Calendários</FONT></B>"
  ShowHTML "<BR/><BR/>"
  
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  If w_ea = "L" Then
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML "<tr><td><font size=""2""><a accesskey=""I"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & par & "&w_ea=I&CL=" & CL & """><u>I</u>ncluir</a>&nbsp;"
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Data</font></td>"
    ShowHTML "          <td><font size=""1""><b>Tipo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Ocorrência</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      w_ano  = ""
      dim Calendario
      
      Calendario = ""
      While Not RS.EOF        
        'While Calendario = RS("nomecal")
            If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
            If wAno <> year(RS("dt_ocorrencia")) Then
               ShowHTML "      <tr bgcolor=""#C0C0C0"" valign=""top""><TD colspan=4 align=""center""><font size=2><B>" & year(RS("dt_ocorrencia")) & "</b></font></td></tr>"
               wAno = year(RS("dt_ocorrencia"))
               Calendario = ""
            End If
            if Calendario <> RS("nomecal") Then
                ShowHTML "      <tr bgcolor=""#fffff1"" valign=""top""><TD colspan=4 align=""left""><font size=2><B>" & RS("nomecal") & "</b></font></td></tr>"
                Calendario = RS("nomecal")
            End If
            ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
            ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("dt_ocorrencia"),2)) & "</td>"
            ShowHTML "        <td><font size=""1"">" & nvl(RS("nome"),"---") & "</td>"
            ShowHTML "        <td><font size=""1"">" & RS("ds_ocorrencia") & "</td>"
            ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
            ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_ocorrencia") & """>Alterar</A>&nbsp"
            ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_chave=" & RS("sq_ocorrencia") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
            ShowHTML "        </td>"
            ShowHTML "      </tr>"
            RS.MoveNext        
      wend
    End If
    ShowHTML "      </center>"
    ShowHTML "    </table>"
    ShowHTML "  </td>"
    ShowHTML "</tr>"
    DesconectaBD
  ElseIf Instr("IAEV",O) > 0 Then
    If InStr("EV",w_ea) Then
       w_Disabled = " DISABLED "
    End If
    AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
    ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & w_ew & """>"
    ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_chave"" value=""" & w_chave & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente"" value=""" & replace(CL,"sq_cliente=","") & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"

    ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
    ShowHTML "    <table width=""95%"" border=""0"">"
    ShowHTML " <tr><td style=""border: 1px; background: #ffffff; border:#cccccc; font: 9px; color:#666666""><STRONG><font color=""#CC0000"">Atenção:</font></STRONG><STRONG></STRONG><ul><li><STRONG>Apresentação dos Professores</STRONG> não deve ser em data posterior ao início do Semestre/Ano Letivo.</li><li><STRONG>Recesso Escolar</STRONG> só deve ser marcado após término do Semestre Letivo.</li><li><b>Avaliação Final</b> só pode ser marcada dentro do Semestre/Ano Letivo</li><li><b>Férias Coletivas dos Professores</b> deve ser de 30 dias corridos.</li></ul></td></tr>"
    ShowHTML "        <tr valign=""top""></tr>"
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>D</u>ata:</b><br><input accesskey=""D"" type=""text"" name=""w_dt_ocorrencia"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(FormatDateTime(Nvl(w_dt_ocorrencia,Date()),2)) & """ onKeyDown=""FormataData(this,event);"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a data de ocorrência. O sistema colocará as barras automaticamente.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "          <td valign=""top""><font size=""1""><b>D<u>e</u>scrição:</b><br><input " & w_Disabled & " accesskey=""E"" type=""text"" name=""w_ds_ocorrencia"" class=""STI"" SIZE=""60"" MAXLENGTH=""60"" VALUE=""" & w_ds_ocorrencia & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva a ocorrência.','white')""; ONMOUSEOUT=""kill()""></td>"
        ShowHTML "          <td><font size=""1""><b>Tipo da ocorrência:</b><br><SELECT CLASS=""STI"" NAME=""w_tipo"">"
    ShowHTML "          <option value=""""> ---"
    SQL = "SELECT * FROM escTipo_Data a WHERE a.abrangencia <> 'P' AND a.sigla <> 'CN' AND a.sigla <> 'FF' ORDER BY a.nome" & VbCrLf
    ConectaBD SQL
    While Not RS.EOF
       If cDbl(nvl(RS("sq_tipo_data"),0)) = cDbl(nvl(w_tipo,0)) Then
          ShowHTML "          <option value=""" & RS("sq_tipo_data") & """ SELECTED>" & RS("nome")
       Else
          ShowHTML "          <option value=""" & RS("sq_tipo_data") & """>" & RS("nome")
       End If
       RS.MoveNext
    Wend
    ShowHTML "          </select></tr>"
    If w_ea <> "A" Then
        ShowHTML "          <tr><td>"      
        montaCheckBoxCal replace(CL,"sq_cliente=","")
        ShowHTML "          </td></tr>"
    Else
        SQL = " select c.nome as nomecal " & VbCrLf & _ 
              " from escCalendario_Cliente a " & VbCrLf & _ 
              " left join  escParticular_Calendario c on (a.sq_particular_calendario = c.sq_particular_calendario)" & VbCrLf & _ 
              " where sq_site_cliente=" & replace(CL,"sq_cliente=","") & " and sq_ocorrencia = " & Request("w_chave") 
        ConectaBD SQL
        ShowHTML "          <tr><td>"
            ShowHTML "<label><font size=""1""><b>Atribuição:</b></font></label>"
            ShowHTML " <br><font size=""1""><input name=""checkbox"" type=""checkbox"" checked=""checked""disabled>" &  RS("nomecal")
            DesconectaBD
        ShowHTML "          </td></tr>"
    End If
    ShowHTML "        </table>"
    ShowHTML "      <tr>"
    ShowHTML "      <tr><td align=""center"" colspan=4><hr>"
    If w_ea = "E" Then
       ShowHTML "   <input class=""STB"" type=""submit"" name=""Botao"" value=""Excluir"">"
    Else
       If w_ea = "I" Then
          ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Incluir"">"
       Else
          ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Atualizar"">"
       End If
    End If
    ShowHTML "            <input class=""STB"" type=""button"" onClick=""location.href='" & w_Pagina & w_ew & "&CL=" & CL & "&w_ea=L';"" name=""Botao"" value=""Cancelar"">"
    ShowHTML "          </td>"
    ShowHTML "      </tr>"
    ShowHTML "    </table>"
    ShowHTML "    </TD>"
    ShowHTML "</tr>"
    ShowHTML "</FORM>"
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

Sub GetCalendarios
  Dim w_chave, w_dt_ocorrencia, w_ds_ocorrencia, w_ano

  Dim w_troca, i, w_texto

  w_Chave           = Request("w_Chave")
  w_troca           = Request("w_troca")

  If w_troca > "" Then ' Se for recarga da página
     w_nome               = Request("w_nome")
     w_ativo              = UCase(Request("w_ativo"))
     w_ordem              = Request("w_ordem")
     
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select * from escParticular_Calendario where " & CL & "order by ordem, nome"
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escParticular_Calendario where sq_particular_calendario = " & w_chave
     ConectaBD SQL
     w_nome   = RS("nome")
     w_ativo  = UCase(RS("ativo"))
     w_ordem  = RS("ordem")
     w_obs    = RS("observacao")
     DesconectaBD
  End If

  Cabecalho
  ShowHTML "<HEAD>"
  If InStr("IAEP",O) > 0 Then
     ScriptOpen "JavaScript"
     CheckBranco
     FormataData
     ValidateOpen "Validacao"
     ShowHTML "  theForm.Botao[0].disabled=true;"
     ShowHTML "  theForm.Botao[1].disabled=true;"
     ValidateClose
     ScriptClose
  End If
  ShowHTML "</HEAD>"
  If w_troca > "" Then
     BodyOpen "onLoad='document.Form." & w_troca & ".focus()';"
  ElseIf w_ea = "I" or w_ea = "A" Then
     BodyOpen "onLoad='document.Form.w_nome.focus()';"
  Else
     BodyOpen "onLoad='document.Form.w_nome.focus()';"
  End If
  ShowHTML "<H3>" & Session("nome_cliente") & "</H3>"
  ShowHTML "<HR>"
  ShowHTML "<B><FONT COLOR=""#000000"">Calendários da Unidade</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  If w_ea = "L" Then
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML "<tr><td><font size=""2""><a accesskey=""I"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & par & "&w_ea=I&CL=" & CL & """><u>I</u>ncluir</a>&nbsp;"
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Título</font></td>"
    ShowHTML "          <td><font size=""1""><b>Ativo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operação</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      w_ano  = ""
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & nvl(RS("nome"),"---") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("ativo") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_particular_calendario") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_chave=" & RS("sq_particular_calendario") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
        ShowHTML "        </td>"
        ShowHTML "      </tr>"

        RS.MoveNext
      wend
    End If
    ShowHTML "      </center>"
    ShowHTML "    </table>"
    ShowHTML "  </td>"
    ShowHTML "</tr>"
    DesconectaBD
  ElseIf Instr("IAEV",O) > 0 Then
    If InStr("EV",w_ea) Then
       w_Disabled = " DISABLED "
    End If
    AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
    ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & w_ew & """>"
    ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_chave"" value=""" & w_chave & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente"" value=""" & replace(CL,"sq_cliente=","") & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"

ShowHTML " <script> function Numerics(e) { "
ShowHTML "     var tecla=(window.event)?event.keyCode:e.which;"
ShowHTML "     if((tecla > 47 && tecla < 58)) return true;"
ShowHTML "     else{"
ShowHTML "     if (tecla != 8) return false;"
ShowHTML "     else return true;"
ShowHTML "     }"
ShowHTML " } "
ShowHTML " </script> "

    ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
    ShowHTML "    <table width=""95%"" border=""0"">"
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b<u>T</u>ítulo:</b><br><input " & w_Disabled & " accesskey=""T"" type=""text"" name=""w_nome"" class=""STI"" SIZE=""60"" MAXLENGTH=""60"" VALUE=""" & w_nome & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva a ocorrência.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "          <td valign=""top""><font size=""1""><b<u>O</u>rdenação:</b><br><input " & w_Disabled & " accesskey=""O"" type=""text"" name=""w_ordem"" class=""STI"" VALUE=""" & w_ordem & """ maxlength=""4"" size=""5"" onKeyPress=""return Numerics(event);"" ></td>"
    MontaRadioNS "<u>A</u>tivo?", w_ativo, "w_ativo"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b<u>O</u>bservação:</b><br><textarea " & w_Disabled & " accesskey=""T"" name=""w_obs"" class=""STI"" COLS=""40"" ROWS=""5"" VALUE=""" & w_obs & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva a ocorrência.','white')""; ONMOUSEOUT=""kill()""></textarea></td>"
    ShowHTML "        </tr>"
    ShowHTML "        </table>"
    ShowHTML "      <tr>"
    ShowHTML "      <tr><td align=""center"" colspan=4><hr>"
    If w_ea = "E" Then
       ShowHTML "   <input class=""STB"" type=""submit"" name=""Botao"" value=""Excluir"">"
    Else
       If w_ea = "I" Then
          ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Incluir"">"
       Else
          ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Atualizar"">"
       End If
    End If
    ShowHTML "            <input class=""STB"" type=""button"" onClick=""location.href='" & w_Pagina & w_ew & "&CL=" & CL & "&w_ea=L';"" name=""Botao"" value=""Cancelar"">"
    ShowHTML "          </td>"
    ShowHTML "      </tr>"
    ShowHTML "    </table>"
    ShowHTML "    </TD>"
    ShowHTML "</tr>"
    ShowHTML "</FORM>"
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
REM Fim do cadastro de calendários personalizados
REM -------------------------------------------------------------------------

Private Sub MainBody

  If Nvl(w_IN,0) = 1 then
    ShowFrames
  ElseIf CL > "" Then
    Select Case uCase(w_EW)
      Case "SHOWMENU"                   showMenu
      Case conWhatDadosAdicionais       ShowCadastro

      Case conWhatNotCliente            ShowPortaria
      Case conWhatDocumento             ShowCurso
      Case conWhatMensagem              ShowOS
      Case conWhatCalendario            GetCalendario
      Case conWhatCliente               GetCliente
      Case "CALS"                       GetCalendarios
      Case "FORMULARIO"                 GetArquivos
      Case "FORMCAL"                    FormCal
      Case "GRAVA"                      Grava
      Case Else

           If ( Not Request.QueryString( conToMakeSystem ) > "" ) Then
              ShowFrames
           End If
    End Select

  End If

End Sub
%>