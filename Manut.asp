<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="Funcoes.asp"-->
<!--#INCLUDE FILE="jScript.asp"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="Constants_ADO.inc"-->
<!--#INCLUDE FILE="esc.inc"-->
<%
Response.Expires = -1500
REM =========================================================================
REM  /Manut.asp
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
  End If

  w_Data = Mid(100+Day(Date()),2,2) & "/" & Mid(100+Month(Date()),2,2) & "/" &Year(Date())
  w_Pagina = ExtractFileName(Request.ServerVariables("SCRIPT_NAME")) & "?w_ew="
  
  If w_ew = "LOGON" Then
     AbreSessao
     Logon
  ElseIf w_ew = "VALIDA" Then
     AbreSessao
     Valida
  ElseIf w_ew = "ENCERRAR" Then
     Session.Abandon()
     ShowHTML "<HTML>" & chr(13) & chr(10)
     ShowHTML "<BODY>" & chr(13) & chr(10)
     ShowHTML "<SCRIPT LANGUAGE='JAVASCRIPT'>" & chr(13) & chr(10)
     ShowHTML "  top.location.href='Manut.asp?w_ew=logon';" & chr(13) & chr(10)
     ShowHTML "</SCRIPT>" & chr(13) & chr(10)
     ShowHTML "</BODY>" & chr(13) & chr(10)
     ShowHTML "</HTML>" & chr(13) & chr(10)
  ElseIf CL > "" Then
     AbreSessao
     
     SQL = "select b.tipo from escCliente a inner join escTipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente and a." & CL & ")"
     ConectaBD SQL
     w_tipo = RS("tipo")
     DesconectaBD
     Main
  Else
     AbreSessao
     ShowHTML "<HTML>" & chr(13) & chr(10)
     ShowHTML "<BODY>" & chr(13) & chr(10)
     ShowHTML "<SCRIPT LANGUAGE='JAVASCRIPT'>" & chr(13) & chr(10)
     ShowHTML "  top.location.href='Manut.asp?w_ew=logon';" & chr(13) & chr(10)
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
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

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
  ShowHTML "    <frame name=""Menu"" src=""Manut.asp?CL=" & CL & "&w_tipo=" & w_tipo & "&w_ew=SHOWMENU"" scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
  ShowHTML "    <frame name=""Body"" src=""Manut.asp?CL=" & CL & "&w_ee=1&w_ea=A&w_ew=" & conWhatSite & """ scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
  ShowHTML "</frameset>"
  ShowHTML "</html>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Frame
REM =========================================================================

REM =========================================================================
REM Rotina de montagem do menu
REM -------------------------------------------------------------------------
Sub showMenu

   Dim w_imagem
   
   w_imagem = "img/SheetLittle.gif"

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
   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatDocumento & "&w_ee=1"" Title=""Cadastra arquivos para download!"">Arquivos (<i>download</i>)</A><br> "
   If w_tipo = 3 Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatCalendario & "&w_ee=1"" Title=""Cadastra datas especiais da unidade!"">Calendário</A><br>"
      SQL = "select b.ds_especialidade from escEspecialidade_Cliente a inner join escEspecialidade b on (a.sq_codigo_espec = b.sq_especialidade and a." & CL & ")"
      ConectaBD SQL
      If Not RS.EOF Then
         If uCase(RS("ds_especialidade")) <> uCase("Biblioteca") Then
            ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Manut.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatMensagem & "&w_ee=1"" Title=""Cadastra mensagens da unidade dirigidas seus alunos!"">Mensagens</A><br>"
         End If
      End If
      DesconectaBD
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
  ShowHTML "<form method=""post"" action=""manut.asp?w_ew=Valida"" onsubmit=""return(Validacao(this));"" name=""Form""> "
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

  sql = "SELECT case in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end AS in_destinatario, " & VbCrLf & _
        "       dt_arquivo, ds_titulo, nr_ordem, ds_arquivo, ln_arquivo, 'SEDF' AS Origem, x.ln_internet diretorio " & VbCrLf & _
        "From escCliente_Arquivo AS a INNER JOIN escCliente AS x ON (a.sq_site_cliente = x.sq_cliente)" & VbCrLf & _
        "WHERE in_ativo = 'Sim'" & VbCrLf & _
        "  AND x.sq_cliente = 0" & VbCrLf & _
        "  AND in_destinatario = 'E'" & VbCrLf & _
        "  and YEAR(a.dt_arquivo) = " & wAno & VbCrLf & _
        "ORDER BY origem, nr_ordem, dt_arquivo desc, in_destinatario " & VbCrLf 
  ConectaBD SQL
  ShowHTML " <tr><td><table width=""100%"" border=""0"">"
  ShowHTML "          <TR><TD align=""center""><br><table border=0 cellpadding=0 cellspacing=0><tr><td>"
  ShowHTML "              <P><font face=""Arial"" size=1><b>ARQUIVOS INSERIDOS PELA SEDF</b></font><br>"  
  ShowHTML "<tr><td><TABLE border=0 cellSpacing=5 width=""95%"">"
  ShowHTML "  <TR>"
  ShowHTML "    <TD><FONT face=""Verdana"" size=1><b>Origem"
  ShowHTML "    <TD><FONT face=""Verdana"" size=1><b>Alvo"
  ShowHTML "    <TD><FONT face=""Verdana"" size=1><b>Data"
  ShowHTML "    <TD><FONT face=""Verdana"" size=1><b>Componente curricular"
  ShowHTML "  </TR>"
  ShowHTML "  <TR>"
  ShowHTML "    <TD COLSPAN=""4"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
  ShowHTML "  </TR>"
  wCont = 0
     
  If Not RS.EOF Then
     wCont = 1
     Do While Not RS.EOF
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD><FONT face=""Verdana"" size=1>" & RS("origem")
        ShowHTML "    <TD><FONT face=""Verdana"" size=1>" & RS("in_destinatario")
        ShowHTML "    <TD><FONT face=""Verdana"" size=1>" & Mid(100+Day(RS("dt_arquivo")),2,2) & "/" & Mid(100+Month(RS("dt_arquivo")),2,2) & "/" &Year(RS("dt_arquivo"))
        ShowHTML "    <TD><FONT face=""Verdana"" size=1><a href=""http://" & replace(RS("diretorio"),"http://","") & "/" & RS("ln_arquivo") & """ target=""_blank"">" & RS("ds_titulo") & "</a><br><div align=""justify""><font size=1>.:. " & RS("ds_arquivo") & "</div>"
        ShowHTML "  </TR>"

        RS.MoveNext
     Loop
  Else
     ShowHTML "  <TR><TD COLSPAN=4 ALIGN=CENTER><FONT face=""Verdana"" size=1>Não há arquivos disponíveis no momento para o ano de " & wAno & " </TR>"
  End If
  RS.Close

  SQL = "SELECT year(dt_arquivo) ano " & VbCrLf & _
        "From escCliente_Arquivo a INNER JOIN escCliente AS x ON (a.sq_site_cliente = x.sq_cliente)" & VbCrLf & _
        "WHERE in_ativo = 'Sim'" & VbCrLf & _
        "  AND x.sq_cliente = 0" & VbCrLf & _
        "  AND in_destinatario = 'E'" & VbCrLf & _
        "  and YEAR(a.dt_arquivo) <> " & wAno & VbCrLf & _
        "ORDER BY year(dt_arquivo) desc " & VbCrLf 
  ConectaBD SQL
    If Not RS.EOF Then
       ShowHTML "  <TR><TD COLSPAN=4 ><FONT face=""Verdana"" size=1><b>Arquivos de outros anos</b><br>"
       While Not RS.EOF
          ShowHTML "     <li><a href=""" & w_dir & "Manut.asp?wAno=" & RS("ano") & """ >Arquivos de " & RS("ano") & "</a></TR>"
          RS.MoveNext
       Wend
       ShowHTML "  </TD></TR>"
    End If
    RS.Close
  ShowHTML "</table>"
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
REM Fim da rotina de criação da tela de logon
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina de autenticação dos usuários
REM -------------------------------------------------------------------------
Sub Valida

  Dim w_Erro, w_cliente, w_uid, w_pwd
  
  w_uid = replace(replace(Trim(uCase(Request("Login"))),"'", ""), """", "")
  w_pwd = replace(replace(Trim(uCase(Request("Password"))),"'", ""), """", "")

  w_Erro = 0
  SQL = "select count(*) existe from escCliente where upper(ds_username) = upper('" & w_uid & "')"
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
     Session("username") = uCase(RS("ds_username"))
     w_cliente           = RS("sq_cliente")
     DesconectaBD
     
     ' Grava o acesso na tabela de log
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

     ShowHTML "  location.href='manut.asp?&w_in=1&cl=sq_cliente=" & w_cliente & "';"
  End If
  ScriptClose

  Set w_Erro    = Nothing
  Set w_Cliente = Nothing

End Sub
REM =========================================================================
REM Fim da rotina de autenticação de usuários
REM -------------------------------------------------------------------------

REM =========================================================================
REM Cadastro de arquivos
REM -------------------------------------------------------------------------
Sub GetDocumento
  Dim w_chave, w_ds_titulo, w_in_ativo, w_ds_arquivo, w_ln_arquivo
  Dim w_dt_arquivo, w_in_destinatario, w_nr_ordem

  Dim w_troca, i, w_texto

  w_Chave           = Request("w_Chave")
  w_troca           = Request("w_troca")

  If w_troca > "" Then ' Se for recarga da página
     w_dt_arquivo      = Request("w_dt_arquivo")
     w_ds_titulo       = Request("w_ds_titulo")
     w_in_ativo        = Request("w_in_ativo")
     w_ds_arquivo      = Request("w_ds_arquivo")
     w_ln_arquivo      = Request("w_ln_arquivo")
     w_in_destinatario = Request("w_in_destinatario")
     w_nr_ordem        = Request("w_nr_ordem")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select * from escCliente_Arquivo where " & replace(CL,"sq_cliente","sq_site_cliente") & " order by nr_ordem, ltrim(upper(ds_titulo))"
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select b.ds_diretorio, a.* from escCliente_Arquivo a inner join escCliente_Site b on (a.sq_site_cliente = b.sq_site_cliente) where a.sq_arquivo = " & w_chave
     ConectaBD SQL
     w_dt_arquivo      = RS("dt_arquivo")
     w_ds_titulo       = RS("ds_titulo")
     w_in_ativo        = RS("in_ativo")
     w_ds_arquivo      = RS("ds_arquivo")
     w_ln_arquivo      = RS("ln_arquivo")
     w_in_destinatario = RS("in_destinatario")
     w_nr_ordem        = RS("nr_ordem")
     w_ds_diretorio    = RS("ds_diretorio")
     DesconectaBD
  End If

  Cabecalho
  ShowHTML "<HEAD>"
  If InStr("IAEP",w_ea) > 0 Then
     ScriptOpen "JavaScript"
     ValidateOpen "Validacao"
     If InStr("IA",w_ea) > 0 Then
        Validate "w_ds_titulo", "Título", "", "1", "2", "50", "1", "1"
        Validate "w_ds_arquivo", "Descrição", "", "1", "2", "200", "1", "1"
        If w_ea = "I" Then
           Validate "w_ln_arquivo", "link", "", "1", "2", "100", "1", "1"
        End If
        Validate "w_nr_ordem", "Nr. de ordem", "", "1", "1", "2", "1", "0123546789"
        ShowHTML " if (theForm.w_ln_arquivo.value > """"){"
        ShowHTML "    if((theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.DLL')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.SH')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.BAT')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.EXE')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.ASP')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.PHP')!=-1)) {"
        ShowHTML "       alert('Tipo de arquivo não permitido!');"
        ShowHTML "       theForm.w_ln_arquivo.focus(); "
        ShowHTML "       return false;"
        ShowHTML "    }"
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
     BodyOpen "onLoad='document.Form.w_ds_titulo.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de arquivos da unidade</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  ShowHTML "  <tr><td align=""center"" colspan=""2""><font size=""1"" color=""red""><b>IMPORTANTE: <a href=""sedf/orientacoes_word.pdf"" class=""hl"" target=""_blank"">Para documentos Word, clique aqui para ler as orientações sobre a formatação e a proteção do texto</a></b></font>."
  If w_ea = "L" Then
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML "<tr><td><font size=""2""><a accesskey=""I"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & par & "&w_ea=I&CL=" & CL & """><u>I</u>ncluir</a>&nbsp;"
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Ordem</font></td>"
    ShowHTML "          <td><font size=""1""><b>Arquivo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Descrição</font></td>"
    ShowHTML "          <td><font size=""1""><b>Ativo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("nr_ordem") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("ds_titulo") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("ds_arquivo") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("in_ativo") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_arquivo") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=E&CL=" & CL & "&w_chave=" & RS("sq_arquivo") & """>Excluir</A>&nbsp"
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
    ShowHTML "<FORM action=""" & w_pagina & "Grava"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"" enctype=""multipart/form-data"">"
    ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & w_ew & """>"
    ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_chave"" value=""" & w_chave & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente"" value=""" & replace(CL,"sq_cliente=","") & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"

    ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
    ShowHTML "    <table width=""95%"" border=""0"">"
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>T</u>ítulo:</b><br><input " & w_Disabled & " accesskey=""T"" type=""text"" name=""w_ds_titulo"" class=""STI"" SIZE=""50"" MAXLENGTH=""50"" VALUE=""" & w_ds_titulo & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe um título para o arquivo.','white')""; ONMOUSEOUT=""kill()""></td>"
    If w_ea = "I" Then
       ShowHTML "          <td valign=""top""><font size=""1""><b>Cadastramento:</b><br><input disabled type=""text"" name=""w_dt_arquivo"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(FormatDateTime(date(),2)) & """ ONMOUSEOVER=""popup('Data de inclusão do arquivo.','white')""; ONMOUSEOUT=""kill()""></td>"
    Else
       ShowHTML "          <td valign=""top""><font size=""1""><b>Última alteração:</b><br><input disabled type=""text"" name=""w_dt_arquivo"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(FormatDateTime(w_dt_arquivo,2)) & """ ONMOUSEOVER=""popup('Data da última alteração deste arquivo.','white')""; ONMOUSEOUT=""kill()""></td>"
    End If
    ShowHTML "        </table>"
    ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>D</u>escrição:</b><br><textarea " & w_Disabled & " accesskey=""D"" name=""w_ds_arquivo"" class=""STI"" ROWS=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva a finalidade do arquivo.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_arquivo & "</TEXTAREA></td>"
    ShowHTML "      <tr>"
    ShowHTML "      </tr>"
    ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>L</u>ink:</b><br><input " & w_Disabled & " accesskey=""L"" type=""file"" name=""w_ln_arquivo"" class=""STI"" SIZE=""80"" MAXLENGTH=""100"" VALUE="""" ONMOUSEOVER=""popup('OBRIGATÓRIO. Clique no botão ao lado para localizar o arquivo. Ele será transferido automaticamente para o servidor.','white')""; ONMOUSEOUT=""kill()"">"
    If w_ln_arquivo > "" Then
       ShowHTML "              <b><a class=""SS"" href=""http://" & Replace(lCase(w_ds_diretorio),lCase("http://"),"") & "/" & w_ln_arquivo & """ target=""_blank"" title=""Clique para exibir o arquivo atual."">Exibir</a></b>"
    End If
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td><font size=""1""><b><u>N</u>r. de ordem:</b><br><input " & w_Disabled & " accesskey=""N"" type=""text"" name=""w_nr_ordem"" class=""STI"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & w_nr_ordem & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a posição em que este arquivo deve aparecer na lista de arquivos disponíveis. Ex: 1, 2, 3 etc.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "          <td><font size=""1""><b><u>D</u>estinatários:</b><br><select " & w_Disabled & " accesskey=""D"" name=""w_in_destinatario"" class=""STS"" SIZE=""1"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o público ao qual o arquivo destina-se.','white')""; ONMOUSEOUT=""kill()"">"
    If w_in_destinatario = "A" Then
       ShowHTML "            <OPTION VALUE=""A"" SELECTED>Apenas alunos <OPTION VALUE=""P"">Apenas professores <OPTION VALUE=""T"">Professores e alunos"
    ElseIf w_in_destinatario = "E" Then
       ShowHTML "            <OPTION VALUE=""A"">Apenas alunos <OPTION VALUE=""P"">Apenas professores <OPTION VALUE=""T"">Professores e alunos"
    ElseIf w_in_destinatario = "P" Then
       ShowHTML "            <OPTION VALUE=""A"">Apenas alunos <OPTION VALUE=""P"" SELECTED>Apenas professores <OPTION VALUE=""T"">Professores e alunos"
    Else
       ShowHTML "            <OPTION VALUE=""A"">Apenas alunos <OPTION VALUE=""P"">Apenas professores <OPTION VALUE=""T"" SELECTED>Professores e alunos"
    End If
    ShowHTML "            </SELECTED></TD>"
    MontaRadioSN "<b>Exibir no site?</b>", w_in_ativo, "w_in_ativo"
    ShowHTML "        </table>"
    ShowHTML "      <tr>"
    ShowHTML "      <tr><td align=""center"" colspan=4><hr>"
    If w_ea = "E" Then
       ShowHTML "   <input class=""STB"" type=""submit"" name=""Botao"" value=""Excluir"" onClick=""return confirm('Confirma a exclusão do registro?');"">"
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
    'ShowHTML " history.back(1);"
    ScriptClose
  End If
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape

  Set w_ds_titulo       = Nothing
  Set w_in_ativo        = Nothing
  Set w_ds_arquivo      = Nothing
  Set w_ln_arquivo      = Nothing

  Set w_chave           = Nothing

  Set w_troca           = Nothing
  Set i                 = Nothing
  Set w_texto           = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de arquivos
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina de alteração dos recursos da etapa
REM -------------------------------------------------------------------------
Sub GetEspecialidadeCliente

  Dim w_chave, w_cont, wAtual

  If w_ea = "L" Then
     SQL = "select a.sq_especialidade, a.ds_especialidade, a.tp_especialidade, b.sq_codigo_cli " & VbCrLf & _
           "  from escEspecialidade a " & VbCrLf & _
           "       left outer join escEspecialidade_Cliente b on (a.sq_especialidade = b.sq_codigo_espec and b." & CL & ") " & VbCrLf & _
           "order by a.nr_ordem, a.ds_especialidade " & VbCrLf
     ConectaBD SQL
  End If

  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "Javascript"
  ValidateOpen "Validacao"
  If w_ea = "L" Then
     ShowHTML "  for (i = 0; i < theForm.w_chave.length; i++) {"
     ShowHTML "      if (theForm.w_chave[i].checked) break;"
     ShowHTML "      if (i == theForm.w_chave.length-1) {"
     ShowHTML "         alert('Você deve selecionar pelo menos uma área de atuação!');"
     ShowHTML "         return false;"
     ShowHTML "      }"
     ShowHTML "  }"
     ShowHTML "  theForm.Botao.disabled=true;"
  End If
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Atualização de áreas de atuação</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
  ShowHTML MontaFiltro("POST")
  ShowHTML "<input type=""hidden"" name=""R"" value=""" & w_ew & """>"
  ShowHTML "<input type=""hidden"" name=""w_sq_cliente"" value=""" & replace(cl,"sq_cliente=","") & """>"
  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Áreas de atuação</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Selecione as áreas de atuação oferecidas pela unidade.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  wAtual = ""
  While Not RS.EOF
    If wAtual = "" or wAtual <> RS("tp_especialidade") Then
       wAtual = RS("tp_especialidade")
       If wAtual = "M" Then
          ShowHTML "          <TR><TD><font size=1><b>Etapas / Modalidades de ensino</b>:"
       ElseIf wAtual = "R" Then
          ShowHTML "          <TR><TD><font size=1><b>Em Regime de Intercomplementaridade</b>:"
       Else
          ShowHTML "          <TR><TD><font size=1><b>Outras</b>:"
       End If
    End If
    If cDbl(Nvl(RS("sq_codigo_cli"),0)) > 0 Then
       ShowHTML "      <tr valign=""top""><td><font size=""1""><input type=""checkbox"" name=""w_chave"" value=""" & RS("sq_especialidade") & """ checked>" & RS("ds_especialidade") & "</td></tr>"
    Else
       ShowHTML "      <tr valign=""top""><td><font size=""1""><input type=""checkbox"" name=""w_chave"" value=""" & RS("sq_especialidade") & """>" & RS("ds_especialidade") & "</td></tr>"
    End If
    RS.MoveNext
  wend
  ShowHTML "      </center>"
  ShowHTML "    </table>"
  ShowHTML "  </td>"
  ShowHTML "</tr>"
  DesConectaBD
  ShowHTML "      <tr><td align=""center""><font size=1>&nbsp;"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000"">"
  ShowHTML "      <tr><td align=""center"">"
  ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Gravar"">"
  ShowHTML "          </td>"
  ShowHTML "      </tr>"
  ShowHTML "</table>"
  ShowHTML "</center>"
  ShowHTML "</FORM>"
  Rodape

  Set w_chave       = Nothing
  Set w_cont        = Nothing


End Sub
REM =========================================================================
REM Fim da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Tela de dados adicionais
REM -------------------------------------------------------------------------
Sub GetDadosAdicionais

  Dim w_sq_cliente, w_nr_cnpj, w_tp_registro, w_ds_ato, w_nr_ato, w_dt_ato, w_ds_orgao
  Dim w_ds_logradouro, w_no_bairro, w_nr_cep, w_no_contato, w_ds_email_contato
  Dim w_nr_fone_contato, w_nr_fax_contato, w_no_diretor, w_no_secretario

  If w_ea = "A" Then
     SQL = "select * from escCliente_Dados where " & CL
     ConectaBD SQL
     If Not RS.EOF Then
        w_sq_cliente       = RS("sq_cliente")
        w_nr_cnpj          = RS("nr_cnpj")
        w_tp_registro      = RS("tp_registro")
        w_ds_ato           = RS("ds_ato")
        w_nr_ato           = RS("nr_ato")
        w_dt_ato           = RS("dt_ato")
        w_ds_orgao         = RS("ds_orgao")
        w_ds_logradouro    = RS("ds_logradouro")
        w_no_bairro        = RS("no_bairro")
        w_nr_cep           = trim(replace(Nvl(RS("nr_cep"),""),".",""))
        w_no_contato       = RS("no_contato")
        w_ds_email_contato = RS("ds_email_contato")
        w_nr_fone_contato  = RS("nr_fone_contato")
        w_nr_fax_contato   = RS("nr_fax_contato")
        w_no_diretor       = RS("no_diretor")
        w_no_secretario    = RS("no_secretario")
     Else
        w_sq_cliente       = ""
        w_nr_cnpj          = ""
        w_tp_registro      = ""
        w_ds_ato           = ""
        w_nr_ato           = ""
        w_dt_ato           = ""
        w_ds_orgao         = ""
        w_ds_logradouro    = ""
        w_no_bairro        = ""
        w_nr_cep           = ""
        w_no_contato       = ""
        w_ds_email_contato = ""
        w_nr_fone_contato  = ""
        w_nr_fax_contato   = ""
        w_no_diretor       = ""
        w_no_secretario    = ""
     End If
  End If

  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "Javascript"
  CheckBranco
  FormataDATA
  FormataCEP
  ValidateOpen "Validacao"
  If w_ea = "A" Then
     Validate "w_ds_logradouro", "Logradouro", "1", "1", "4", "60", "1", "1"
     Validate "w_no_bairro", "Bairro", "1", "", "2", "30", "1", "1"
     Validate "w_nr_cep", "CEP", "1", "1", "9", "9", "", "0123456789-"
     Validate "w_no_diretor", "Nome do(a) Diretor(a)", "1", "", "2", "40", "1", "1"
     If w_tipo = 3 Then
        Validate "w_no_secretario", "Nome do(a) Secretário(a)", "1", "", "2", "40", "1", "1"
     End If
     'Validate "w_nr_cnpj", "CNPJ", "1", "", "14", "14", "", "012356789"
     'Validate "w_tp_registro", "Tipo do registro", "1", "", "6", "60", "1", "1"
     'Validate "w_ds_ato", "Ato", "1", "", "1", "30", "1", "1"
     'Validate "w_nr_ato", "Número", "1", "", "1", "15", "1", "1"
     'Validate "w_dt_ato", "Data", "DATA", "", "10", "10", "1", "1"
     'Validate "w_ds_orgao", "Órgão Emissor", "1", "", "1", "30", "1", "1"
     Validate "w_no_contato", "Nome", "1", "1", "2", "35", "1", "1"
     Validate "w_ds_email_contato", "e-Mail", "1", "", "6", "60", "1", "1"
     Validate "w_nr_fone_contato", "Telefone", "1", "1", "6", "20", "1", "1"
     Validate "w_nr_fax_contato", "Fax", "1", "", "6", "20", "1", "1"
  End If
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.Form.w_ds_logradouro.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Atualização de dados adicionais da unidade</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
  ShowHTML MontaFiltro("POST")
  ShowHTML "<input type=""hidden"" name=""R"" value=""" & w_ew & """>"

  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Endereço</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Informe o endereço da unidade, a ser exibido na seção ""Quem Somos"" do site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>L</u>ogradouro:</b><br><INPUT ACCESSKEY=""L"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_logradouro"" size=""60"" maxlength=""60"" value=""" & w_ds_logradouro & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o logradouro de funcionamento da unidade.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b><u>B</u>airro:</b><br><INPUT ACCESSKEY=""B"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_no_bairro"" size=""30"" maxlength=""30"" value=""" & w_no_bairro & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o bairro de funcionamento da unidade.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b>C<u>E</u>P:</b><br><INPUT ACCESSKEY=""C"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_cep"" size=""9"" maxlength=""9"" value=""" & w_nr_cep & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o CEP da unidade.','white')""; ONMOUSEOUT=""kill()"" onKeyDown=""FormataCEP(this,event);""></td>"
  ShowHTML "        </table>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Dirigentes</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Os dados deste bloco são exibidos na seção ""Quem Somos"" do site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b>Nome do(a) <u>D</u>iretor:</b><br><INPUT ACCESSKEY=""D"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_no_diretor"" size=""40"" maxlength=""40"" value=""" & w_no_diretor & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o nome do(a) diretor(a) da unidade.','white')""; ONMOUSEOUT=""kill()""></td>"
  If w_tipo = 3 Then
     ShowHTML "          <td><font size=""1""><b>Nome do(a)<u>S</u>ecretário(a):</b><br><INPUT ACCESSKEY=""S"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_no_secretario"" size=""40"" maxlength=""40"" value=""" & w_no_secretario & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o nome do(a) secretário(a) da unidade.','white')""; ONMOUSEOUT=""kill()""></td>"
  End If
  ShowHTML "        </table>"
'  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
'  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
'  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Registro</td></td></tr>"
'  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
'  ShowHTML "      <tr><td><font size=1>Os dados deste bloco não são exibidos no site, servindo apenas para memória das informações relativas ao registro da unidade.</font></td></tr>"
'  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
'  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
'  ShowHTML "        <tr valign=""top"">"
'  ShowHTML "          <td><font size=""1""><b><u>C</u>NPJ:</b><br><INPUT ACCESSKEY=""C"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_cnpj"" size=""14"" maxlength=""14"" value=""" & w_nr_cnpj & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o CNPJ da unidade.','white')""; ONMOUSEOUT=""kill()""></td>"
'  ShowHTML "          <td><font size=""1""><b><u>T</u>ipo do registro:</b><br><INPUT ACCESSKEY=""T"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_tp_registro"" size=""10"" maxlength=""10"" value=""" & w_tp_registro & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o tipo do registro da unidade. Ex: autorizado, reconhecido etc.','white')""; ONMOUSEOUT=""kill()""></td>"
'  ShowHTML "          <td><font size=""1""><b><u>A</u>to:</b><br><INPUT ACCESSKEY=""A"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_ato"" size=""30"" maxlength=""30"" value=""" & w_ds_ato & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o ato de criação da unidade. Ex: Portaria MEC.','white')""; ONMOUSEOUT=""kill()""></td>"
'  ShowHTML "        <tr valign=""top"">"
'  ShowHTML "          <td><font size=""1""><b><u>N</u>úmero:</b><br><INPUT ACCESSKEY=""N"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_ato"" size=""15"" maxlength=""15"" value=""" & w_nr_ato & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o número do ato de criação da unidade. Ex: 3029/1965','white')""; ONMOUSEOUT=""kill()""></td>"
'  ShowHTML "          <td><font size=""1""><b><u>D</u>ata:</b><br><INPUT ACCESSKEY=""D"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_dt_ato"" size=""10"" maxlength=""10"" value=""" & FormataDataEdicao(w_dt_ato) & """ ONMOUSEOVER=""popup('OPCIONAL. Informe a data do ato de criação da unidade. O sistema coloca automaticamente as barras separadoras.','white')""; ONMOUSEOUT=""kill()"" onKeyDown=""FormataData(this,event);""></td>"
'  ShowHTML "          <td><font size=""1""><b><u>O</u>rgão emissor:</b><br><INPUT ACCESSKEY=""O"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_orgao"" size=""30"" maxlength=""30"" value=""" & w_ds_orgao & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o órgão emissor do registro da unidade.','white')""; ONMOUSEOUT=""kill()""></td>"
'  ShowHTML "        </table>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Contato técnico da unidade</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Informe os dados da pessoa de contato técnico para uso interno, não sendo utilizado para divulgação no site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b>No<u>m</u>e:</b><br><INPUT ACCESSKEY=""M"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_no_contato"" size=""35"" maxlength=""35"" value=""" & w_no_contato & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o nome da pessoa a ser contactada pela equipe técnica. Os dados deste contato não serão exibidos no site.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b><u>e</u>-Mail:</b><br><INPUT ACCESSKEY=""E"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_email_contato"" size=""40"" maxlength=""60"" value=""" & w_ds_email_contato & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o e-mail da pessoa de contato técnico.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b>Tele<u>f</u>one:</b><br><INPUT ACCESSKEY=""F"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_fone_contato"" size=""20"" maxlength=""20"" value=""" & w_nr_fone_contato & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o telefone da pessoa de contato técnico.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b>Fa<u>x</u>:</b><br><INPUT ACCESSKEY=""X"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_fax_contato"" size=""20"" maxlength=""20"" value=""" & w_nr_fax_contato & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o fax da pessoa de contato técnico.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        </table>"
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
  Set w_nr_cnpj         = Nothing
  Set w_tp_registro     = Nothing
  Set w_ds_ato          = Nothing
  Set w_nr_ato          = Nothing
  Set w_dt_ato          = Nothing
  Set w_ds_orgao        = Nothing
  Set w_ds_logradouro   = Nothing
  Set w_no_bairro       = Nothing
  Set w_nr_cep          = Nothing
  Set w_no_contato      = Nothing
  Set w_ds_email_contato= Nothing
  Set w_nr_fone_contato = Nothing
  Set w_nr_fax_contato  = Nothing
  Set w_no_diretor      = Nothing
  Set w_no_secretario   = Nothing
End Sub
REM =========================================================================
REM Fim da tela de dados adicionais
REM -------------------------------------------------------------------------

REM =========================================================================
REM Tela de dados administrativos
REM -------------------------------------------------------------------------
Sub GetAdmin

  Dim w_sq_cliente, w_limpeza, w_merenda, w_banheiro, w_equipamento, w_quantidade

  SQL = "select * from escCliente_Admin where " & CL
  ConectaBD SQL
  If RS.EOF Then
     w_sq_cliente       = replace(CL,"sq_cliente=","")
  Else
     w_sq_cliente       = RS("sq_cliente")
     w_limpeza          = RS("limpeza_terceirizada")
     w_merenda          = RS("merenda_terceirizada")
     w_banheiro         = RS("banheiro")
  End If
  DesconectaBD

  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "Javascript"
  CheckBranco
  FormataDATA
  FormataCEP
  ValidateOpen "Validacao"
  Validate "w_banheiro", "Quantidade de quadros magneticos", "1", "1", "1", "2", "", "0123456789"
  ShowHTML "  for (i = 0; i < theForm.w_equipamento.length; i++) {"
  ShowHTML "      if (isNaN(parseInt(theForm.w_equipamento[i].value))) {"
  ShowHTML "         alert('Informe apenas números na quantidade dos equipamentos!');"
  ShowHTML "         theForm.w_equipamento[i].focus();"
  ShowHTML "         return false;"
  ShowHTML "      }"
  ShowHTML "  }"
  ShowHTML "  theForm.Botao.disabled=true;"
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.Form.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Atualização de dados administrativos da unidade</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
  ShowHTML MontaFiltro("POST")
  ShowHTML "<input type=""hidden"" name=""R"" value=""" & w_ew & """>"

  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center""><font size=1 color=""#BC3131""><b>ATENÇÃO: Após inserir ou alterar os dados abaixo, clique sobre o botão ""<i>Gravar</i>"", localizado no final desta página; caso contrário, seus dados serão perdidos.</td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Serviços</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Informe os dados solicitados abaixo, sobre a terceirização de serviços e a quantidade de quadros magnéticos.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr>"
  MontaRadioNS "<b>O serviço de limpeza da escola é feita por empresa terceirizada?", w_limpeza, "w_limpeza"
  ShowHTML "      <tr>"
  MontaRadioNS "<b>A escola oferece merenda?", w_merenda, "w_merenda"
  ShowHTML "      <tr><td><font size=""1""><b>Informe a quantidade total de QUADROS MAGNÉTICOS da escola:</b><br><INPUT " & w_Disabled & " class=""STI"" type=""text"" name=""w_banheiro"" size=""2"" maxlength=""2"" value=""" & w_banheiro & """></td>"
  ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000""></TD></TR>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Equipamentos</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Verifique a lista abaixo e informe a quantidade de equipamentos existentes na escola.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  SQL = "select a.nome, b.sq_equipamento, b.nome nm_equipamento, IsNull(c.quantidade, 0) quantidade " & VbCrLf & _
        "  from escTipo_Equipamento                    a " & VbCrLf & _
        "       inner      join escEquipamento         b on (a.sq_tipo_equipamento = b.sq_tipo_equipamento) " & VbCrLf & _
        "       left outer join escCliente_Equipamento c on (b.sq_equipamento      = c.sq_equipamento and " & VbCrLf & _
        "                                                    c.sq_cliente          = " & replace(CL,"sq_cliente=","") & " " & VbCrLf & _
        "                                                   ) " & VbCrLf & _
        "order by a.codigo, b.nome " & VbCrLf
  ConectaBD SQL
  If RS.EOF Then
     ShowHTML "      <tr><td><font size=""1""><b>Não há equipamentos cadastrados.</td>"
  Else
     ShowHTML "      <tr><td><table border=0 width=""100%"">"
     w_cont  = 0
     w_atual = ""
     While not RS.EOF
        If w_atual <> RS("nome") Then
           ShowHTML "        <TR><TD colspan=2><font size=2><b>" & RS("nome") & "</b>"
           w_atual = RS("nome")
           w_cont  = 0
        End If
        If w_cont Mod 2 = 0 Then
           ShowHTML "        <TR valign=""top"">"
           w_cont = 0
        End If
        ShowHTML "        <td><font size=1><INPUT " & w_Disabled & " class=""STI"" type=""text"" name=""w_equipamento"" size=""3"" maxlength=""3"" value=""" & RS("quantidade") & """> "
        If cDbl(RS("quantidade")) > 0 Then
           ShowHTML "        <b>" & RS("nm_equipamento") & "</b>"
        Else
           ShowHTML "        " & RS("nm_equipamento")
        End If
        ShowHTML "                         <INPUT " & w_Disabled & " type=""hidden"" name=""w_codigo"" value=""" & RS("sq_equipamento") & """> " & "</td>"
        w_cont = w_cont + 1
        RS.MoveNext
     Wend
     ShowHTML "      </table>"
  End If
  DesconectaBD

  ' Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
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
  Set w_limpeza         = Nothing
  Set w_merenda         = Nothing
  Set w_banheiro        = Nothing
  Set w_equipamento     = Nothing
  Set w_quantidade      = Nothing
End Sub
REM =========================================================================
REM Fim da tela de dados administrativos
REM -------------------------------------------------------------------------

REM =========================================================================
REM Tela de cadastro de fotos
REM -------------------------------------------------------------------------
Sub GetSGE
  Dim w_chave, w_ds_foto, w_tp_foto, w_ln_foto
  Dim w_nr_ordem

  Dim w_troca, i, w_texto

  w_Chave           = Request("w_Chave")
  w_troca           = Request("w_troca")
  w_tp_foto         = Request("w_tp_foto")

  If w_troca > "" Then ' Se for recarga da página
     w_ds_foto         = Request("w_ds_foto")
     w_ln_foto         = Request("w_ln_foto")
     w_nr_ordem        = Request("w_nr_ordem")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select * from escCliente_Foto where " & CL & " and tp_foto = 'P' order by nr_ordem"
     ConectaBD SQL
     SQL = "select * from escCliente_Foto where " & CL & " and tp_foto = 'Q' order by nr_ordem"
     RS1.Open SQL, dbms, adOpenStatic
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select b.ds_diretorio, a.* from escCliente_Foto a inner join escCliente_Site b on (a.sq_cliente = b.sq_site_cliente) where a.sq_cliente_foto = " & w_chave
     ConectaBD SQL
     w_ds_foto         = RS("ds_foto")
     w_tp_foto         = RS("tp_foto")
     w_ln_foto         = RS("ln_foto")
     w_nr_ordem        = RS("nr_ordem")
     w_ds_diretorio    = RS("ds_diretorio")
     DesconectaBD
  End If

  Cabecalho
  ShowHTML "<HEAD>"
  If InStr("IAEP",w_ea) > 0 Then
     ScriptOpen "JavaScript"
     ValidateOpen "Validacao"
     If InStr("IA",w_ea) > 0 Then
        Validate "w_ds_foto", "Título", "", "1", "2", "100", "1", "1"
        If w_ea = "I" Then
           Validate "w_ln_foto", "link", "1", "", "5", "100", "1", "1"
           ShowHTML " if (theForm.w_ln_foto.value > """"){"
           ShowHTML "    if((theForm.w_ln_foto.value.lastIndexOf('.jpg')==-1) && (theForm.w_ln_foto.value.lastIndexOf('.gif')==-1) && (theForm.w_ln_foto.value.lastIndexOf('.JPG')==-1) && (theForm.w_ln_foto.value.lastIndexOf('.GIF')==-1)) {"
           ShowHTML "       alert('Só é possível escolher arquivos com a extensão \'.jpg\' ou \'.gif\'!');"
           ShowHTML "       theForm.w_ln_foto.focus(); "
           ShowHTML "       return false;"
           ShowHTML "    }"
           ShowHTML "  }"           
        End If
        Validate "w_nr_ordem", "Nr. de ordem", "", "1", "1", "2", "1", "0123546789"
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
     BodyOpen "onLoad='document.Form.w_ds_foto.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de fotos da unidade</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  If w_ea = "L" Then
    
    ShowHTML "      <tr><td colspan=""3"" align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
    ShowHTML "      <tr><td colspan=""3"" align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
    ShowHTML "      <tr><td colspan=""3"" valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Fotos da página de abertura do site</td></td></tr>"
    ShowHTML "      <tr><td colspan=""3"" align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
    ShowHTML "      <tr><td colspan=""3""><font size=1>Informe até 5 imagens a serem colocadas na página de abertura do site. A exibição será feita no formato 302px (largura) por 206px (altura). Recomenda-se imagens com extensão JPG, de até 40KB de tamanho.</font></td></tr>"    
    ShowHTML "      <tr><td colspan=""3"" align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"

    
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    If cDbl(RS.RecordCount) < 5 Then
       ShowHTML "<tr><td><font size=""2""><a accesskey=""I"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & par & "&w_ea=I&w_tp_foto=P&CL=" & CL & """><u>I</u>ncluir</a>&nbsp;"
    Else
       ShowHTML "<tr><td><font size=""2"">&nbsp;"
    End If
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Ordem</font></td>"
    ShowHTML "          <td><font size=""1""><b>Descrição</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("nr_ordem") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("ds_foto") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_cliente_foto") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=E&CL=" & CL & "&w_chave=" & RS("sq_cliente_foto") & """>Excluir</A>&nbsp"
        ShowHTML "        </td>"
        ShowHTML "      </tr>"

        RS.MoveNext
      wend
    End If
    ShowHTML "      </center>"
    ShowHTML "    </table>"
    ShowHTML "  </td>"
    ShowHTML "</tr>"
    ShowHTML "<tr><td colspan=""3""><br><br></td></tr>"
    ShowHTML "      <tr><td colspan=""3"" align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
    ShowHTML "      <tr><td colspan=""3"" align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
    ShowHTML "      <tr><td colspan=""3"" valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Fotos da página ""Quem somos""</td></td></tr>"
    ShowHTML "      <tr><td colspan=""3"" align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
    ShowHTML "      <tr><td colspan=""3""><font size=1>Informe até 10 imagens a serem colocadas na página ""Quem somos"" do site. A exibição será feita no formato 201px (largura) por 153px (altura). Recomenda-se imagens com extensão JPG, de até 40KB de tamanho.</font></td></tr>"    
    ShowHTML "      <tr><td colspan=""3"" align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"    
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    If cDbl(RS1.RecordCount) < 10 Then
       ShowHTML "<tr><td><font size=""2""><a accesskey=""I"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & par & "&w_ea=I&w_tp_foto=Q&CL=" & CL & """><u>I</u>ncluir</a>&nbsp;"
    Else
       ShowHTML "<tr><td><font size=""2"">&nbsp;"
    End If    
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS1.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Ordem</font></td>"
    ShowHTML "          <td><font size=""1""><b>Descrição</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS1.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS1.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS1("nr_ordem") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS1("ds_foto") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS1("sq_cliente_foto") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=E&CL=" & CL & "&w_chave=" & RS1("sq_cliente_foto") & """>Excluir</A>&nbsp"
        ShowHTML "        </td>"
        ShowHTML "      </tr>"

        RS1.MoveNext
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
    ShowHTML "<FORM action=""" & w_pagina & "Grava"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"" enctype=""multipart/form-data"">"
    ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & w_ew & """>"
    ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_chave"" value=""" & w_chave & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente"" value=""" & replace(CL,"sq_cliente=","") & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_tp_foto"" value=""" & w_tp_foto & """>"

    ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
    ShowHTML "    <table width=""95%"" border=""0"">"
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>T</u>ítulo:</b><br><input " & w_Disabled & " accesskey=""T"" type=""text"" name=""w_ds_foto"" class=""STI"" SIZE=""60"" MAXLENGTH=""100"" VALUE=""" & w_ds_foto & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe um título para o arquivo.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "        </table>"
    ShowHTML "      <tr>"
    ShowHTML "      </tr>"
    ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>L</u>ink:</b><br><input " & w_Disabled & " accesskey=""L"" type=""file"" name=""w_ln_foto"" class=""STI"" SIZE=""80"" MAXLENGTH=""100"" VALUE="""" ONMOUSEOVER=""popup('OBRIGATÓRIO. Clique no botão ao lado para localizar o arquivo. Ele será transferido automaticamente para o servidor.','white')""; ONMOUSEOUT=""kill()"">"
    If w_ln_foto > "" Then
       ShowHTML "              <b><a class=""SS"" href=""http://" & Replace(lCase(w_ds_diretorio),lCase("http://"),"") & "/" & w_ln_foto & """ target=""_blank"" title=""Clique para exibir o arquivo atual."">Exibir</a></b>"
    End If
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td><font size=""1""><b><u>N</u>r. de ordem:</b><br><input " & w_Disabled & " accesskey=""N"" type=""text"" name=""w_nr_ordem"" class=""STI"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & w_nr_ordem & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a posição em que este arquivo deve aparecer na lista de arquivos disponíveis. Ex: 1, 2, 3 etc.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "        </table>"
    ShowHTML "      <tr>"
    ShowHTML "      <tr><td align=""center"" colspan=4><hr>"
    If w_ea = "E" Then
       ShowHTML "   <input class=""STB"" type=""submit"" name=""Botao"" value=""Excluir"" onClick=""return confirm('Confirma a exclusão do registro?');"">"
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
    'ShowHTML " history.back(1);"
    ScriptClose
  End If
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape

  Set w_ds_foto         = Nothing
  Set w_tp_foto         = Nothing
  Set w_ln_foto         = Nothing

  Set w_chave           = Nothing

  Set w_troca           = Nothing
  Set i                 = Nothing
  Set w_texto           = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de fotos
REM -------------------------------------------------------------------------

REM =========================================================================
REM Tela de atualizações do SGE
REM -------------------------------------------------------------------------
Sub GetSGE1
  Cabecalho
  ShowHTML "<HEAD>"
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Página de controle das atualizações do SGE</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Atualizações do SGE</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>"
  ShowHTML "         Abaixo são listadas as atualizações do SGE, da mais nova até a mais antiga. Para atualizar sua versão, siga os passos abaixo:"
  ShowHTML "         <ol>"
  ShowHTML "           <li>Verifique qual a versão instalada em sua escola;"
  ShowHTML "           <li>Verifique se há alguma versão listada abaixo, mais recente que a instalada em sua escola;"
  ShowHTML "           <li>Clique sobre o link da versão desejada;"
  ShowHTML "           <li>Irá ser exibida uma janela perguntando se você deseja abrir ou salvar o arquivo. Clique sobre o botão &quot;<b><i>Abrir</i></b>&quot;. O processo de transferência do arquivo será iniciado;"
  ShowHTML "           <li>Repita este passo caso a transferência não seja concluída corretamente;"
  ShowHTML "           <li>Quando a transferência for concluída com sucesso, será exibida uma janela. Clique sobre o botão &quot;<b>Unzip</b>&quot;;"
  ShowHTML "           <li>Quando for exibida a mensagem &quot;<b>1 file(s) unzipped successfully</b>&quot;, a atualização estará concluída. Clique em &quot;<b>Ok</b>&quot; e, em seguida no botão &quot;<b>Close</b>&quot;."
  ShowHTML "         <p align=""center""><b>ATENÇÃO: O SGE deve ser encerrado durante o processo de atualização.</b></p>"
  ShowHTML "         <p align=""center""><font size=4 color=""#0000FF""> ESTAS ATUALIZA&Ccedil;&Otilde;ES S&Oacute; PODEM SER APLICADAS NA VERS&Atilde;O 5.0.0.1 OU SUPERIOR DO SGE M&OacuteDULO ESCOLA </font><br>			"
  ShowHTML "         <font size=2 color=""#0000FF"">Caso a sua vers&atilde;o seja inferior &agrave; 5.0.0.1, solicite ao Help-Desk pelo número 342-2229 a instalação em sua escola.</font></p> "
  ShowHTML "         </ol>"
  ShowHTML "      </font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>"
  ShowHTML "         <p align=""center""><b>ATUALIZAÇÕES DISPONÍVEIS:</b></p>"
  ShowHTML "         <table border=""1"">"
  ShowHTML "           <tr>"
  ShowHTML "             <td colspan=""5"" align=""center""><font size=2><b>SGE</b></font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr align=""center"">"
  ShowHTML "             <td><font size=1><b>Arquivo </b></font></td>"
  ShowHTML "             <td><font size=1><b>Altera&ccedil;&otilde;es </b></font></td>"
  ShowHTML "             <td><font size=1><b>Tamanho </b></font></td>"
  ShowHTML "             <td><font size=1><b>Data </b></font></td>"
  ShowHTML "             <td><font size=1><b>Vers&atilde;o </b></font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td colspan=""5""><font size=1><b>Arquivos disponibilizados em 05/04/2005 </td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a class=""hl"" href=""sge/SGE_Censo.exe"">SGE_Censo.dll </a></td>"
  ShowHTML "             <td><font size=1>Conclus&atilde;o e corre&ccedil;&otilde;es dos relat&oacute;rios do censo 2005 </font></td>"
  ShowHTML "             <td align=""right""><font size=1>813k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>05/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.2 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_Passivo.exe"" class=""hl"">SGE_Passivo.dll </a></td>"
  ShowHTML "             <td><font size=1>Corre&ccedil;&atilde;o no banco de dados </font></td>"
  ShowHTML "             <td align=""right""><font size=1>796k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>05/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.2 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td colspan=""5""><font size=1><b>Arquivos disponibilizados em 08/04/2005 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE.exe"" class=""hl"">SGE.EXE </a></td>"
  ShowHTML "             <td><font size=1>Inclus&atilde;o do relat&oacute;rio de capacidade f&iacute;sica </font></td>"
  ShowHTML "             <td align=""right""><font size=1>2305k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>08/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.2 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_Certificado.exe"" class=""hl"">SGE_Certificado.dll </a></td>"
  ShowHTML "             <td><font size=1>Melhorias no certificado do Ensino M&eacute;dio </font></td>"
  ShowHTML "             <td align=""right""><font size=1>746k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>08/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.3 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_Declaracoes.exe"" class=""hl"">SGE_Declaracoes.dll </a></td>"
  ShowHTML "             <td><font size=1>Corre&ccedil;&atilde;o do relat&oacute;rio declara&ccedil;&atilde;o de escolaridade - per&iacute;odo de recesso escolar </font></td>"
  ShowHTML "             <td align=""right""><font size=1>776k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>08/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.3 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_Graficos.exe"" class=""hl"">SGE_Graficos.dll </a></td>"
  ShowHTML "             <td><font size=1>Corre&ccedil;&atilde;o do gr&aacute;fico de aproveitamento de turma </font></td>"
  ShowHTML "             <td align=""right""><font size=1>795k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>08/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.3 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_HistCad.exe"" class=""hl"">SGE_HistCad.dll </a></td>"
  ShowHTML "             <td><font size=1>Corre&ccedil;&atilde;o do hist&oacute;rico de 80 colunas e do relat&oacute;rio de ensino m&eacute;dio de 80 colunas </font></td>"
  ShowHTML "             <td align=""right""><font size=1>864k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>08/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.3 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_MapRes.exe"" class=""hl"">SGE_MapRes.dll </a></td>"
  ShowHTML "             <td><font size=1>Melhoria da performance na gera&ccedil;&atilde;o do mapa de resultados </font></td>"
  ShowHTML "             <td align=""right""><font size=1>708k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>08/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.3 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td colspan=""5""><font size=1><b>Arquivos disponibilizados em 13/04/2005 </td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_CadAlunos.exe"" class=""hl"">SGE_CadAlunos.dll </a></td>"
  ShowHTML "             <td><font size=1>Altera&ccedil;&atilde;o da sele&ccedil;&atilde;o de alunos com a inclus&atilde;o da coluna <em>passivo </em></font></td>"
  ShowHTML "             <td align=""right""><font size=1>977k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>13/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.3 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_Script.exe"" class=""hl"">SGE_Script.dll </a></td>"
  ShowHTML "             <td><font size=1>Corre&ccedil;&atilde;o no banco de dados </font></td>"
  ShowHTML "             <td align=""right""><font size=1>889k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>13/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.4 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_NotasFaltas.exe"" class=""hl"">SGE_NotasFaltas.dll </a></td>"
  ShowHTML "             <td><font size=1>Altera&ccedil;&atilde;o no procedimento de fechamento da tela de digita&ccedil;&atilde;o de notas/faltas </font></td>"
  ShowHTML "             <td align=""right""><font size=1>691k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>13/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.3 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td colspan=""5""><font size=1><b>Atualiza&ccedil;&otilde;es obrigat&oacute;rias apenas para escolas que possuem a modalidade de ensino EJA </td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_EJACad.exe"" class=""hl"">SGE_EJACad.dll </a></td>"
  ShowHTML "             <td><font size=1>Altera&ccedil;&atilde;o da sele&ccedil;&atilde;o de alunos com a inclus&ccedil;&atilde;o da coluna <em>passivo </em></font></td>"
  ShowHTML "             <td align=""right""><font size=1>1122k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>08/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.3 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_HistEJA.exe"" class=""hl"">SGE_HistEJA.dll </a></td>"
  ShowHTML "             <td><font size=1>Corre&ccedil;&atilde;o do hist&oacute;rico de 80 colunas </font></td>"
  ShowHTML "             <td align=""right""><font size=1>721k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>08/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.3 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "           <tr>"
  ShowHTML "             <td><font size=1><a href=""sge/SGE_EJARel.exe"" class=""hl"">SGE_EJARel.dll </a></td>"
  ShowHTML "             <td><font size=1>Implementa&ccedil;&otilde;es dos relat&oacute;rios do EJA conforme solicita&ccedil;&otilde;es </font></td>"
  ShowHTML "             <td align=""right""><font size=1>1302k </font></td>"
  ShowHTML "             <td align=""center""><font size=1>13/04/2005 </font></td>"
  ShowHTML "             <td><font size=1>5.0.0.4 </font></td>"
  ShowHTML "           </tr>"
  ShowHTML "         </table>"
  ShowHTML "         <br>"
  ShowHTML "         <table border=1 width=""100%"">"
  ShowHTML "         <tr valign=""top"" align=""center"">"
  ShowHTML "           <td colspan=5><font size=1><b>OUTROS PROGRAMAS ÚTEIS</b></font></td>"
  ShowHTML "           <tr valign=""top"" align=""center"">"
  ShowHTML "             <td><font size=1><b>Arquivo</b></font></td>"
  ShowHTML "             <td><font size=1><b>Descrição</b></font></td>"
  ShowHTML "             <td><font size=1><b>Tamanho</b></font></td>"
  ShowHTML "             <td><font size=1><b>Data</b></font></td>"
  ShowHTML "           <tr valign=""top"">"
  ShowHTML "             <td><font size=1><a class=""hl"" href=""sge/CORRIGEGDB.EXE"">CORRIGEGDB.EXE</a></font></td>"
  ShowHTML "             <td><font size=1>Realiza a reestruturação da base</font></td>"
  ShowHTML "             <td align=""right""><font size=1>349k</font></td>"
  ShowHTML "             <td align=""center""><font size=1>08/04/2005</font></td>"
  ShowHTML "           <tr valign=""top"">"
  ShowHTML "             <td><font size=1><a class=""hl"" href=""sge/Avaliacao.EXE"">Avaliacao.exe</a></font></td>"
  ShowHTML "             <td><font size=1>Módulo Avaliação</font></td>"
  ShowHTML "             <td align=""right""><font size=1>688k</font></td>"
  ShowHTML "             <td align=""center""><font size=1>05/04/2005</font></td>"
  ShowHTML "           <tr valign=""top"">"
  ShowHTML "             <td><font size=1><a class=""hl"" href=""sge/BackupSGE.EXE"">BackupSGE.EXE</a></font></td>"
  ShowHTML "             <td><font size=1>Aplicativo de backup/restauração da base de dados</font></td>"
  ShowHTML "             <td align=""right""><font size=1>455k</font></td>"
  ShowHTML "             <td align=""center""><font size=1>05/04/2005</font></td>"
  ShowHTML "           <tr valign=""top"">"
  ShowHTML "             <td><font size=1><a class=""hl"" href=""sge/textos.EXE"">Textos.EXE</a></font></td>"
  ShowHTML "             <td><font size=1>Módulo de integração com o MS Word</font></td>"
  ShowHTML "             <td align=""right""><font size=1>550k</font></td>"
  ShowHTML "             <td align=""center""><font size=1>05/04/2005</font></td>"
  ShowHTML "         </table>"
  ShowHTML "      </font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr>"
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  ShowHTML "</table>"
  Rodape
End Sub
REM =========================================================================
REM Fim da tela de dados administrativos
REM -------------------------------------------------------------------------

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
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>S</u>enha de acesso:</b><br><INPUT ACCESSKEY=""S"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_senha_acesso"" size=""14"" maxlength=""14"" value=""" & w_ds_senha_acesso & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a senha desejada para acessar a tela de atualização dos dados do site.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>e</u>-Mail:</b><br><INPUT ACCESSKEY=""E"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_email"" size=""60"" maxlength=""60"" value=""" & w_ds_email & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o e-mail de sua unidade para contatos com a equipe técnica. Este e-mail não será divulgado no site.','white')""; ONMOUSEOUT=""kill()""></td>"
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
REM Fim da tela de dados básicos
REM -------------------------------------------------------------------------

REM =========================================================================
REM Tela de dados do site
REM -------------------------------------------------------------------------
Sub GetSite

  Dim w_sq_cliente, w_no_contato_internet, w_ds_email_internet
  Dim w_nr_fone_internet, w_nr_fax_internet, w_ds_texto_abertura, w_ds_institucional, w_ds_mensagem
  Dim w_pedagogica

  If w_ea = "A" Then
     SQL = "select * from escCliente_Site where " & CL
     ConectaBD SQL
     w_sq_cliente           = RS("sq_cliente")
     w_no_contato_internet  = RS("no_contato_internet")
     w_ds_email_internet    = RS("ds_email_internet")
     w_nr_fone_internet     = RS("nr_fone_internet")
     w_nr_fax_internet      = RS("nr_fax_internet")
     w_ds_texto_abertura    = RS("ds_texto_abertura")
     w_ds_institucional     = RS("ds_institucional")
     w_ds_mensagem          = RS("ds_mensagem")
     w_ds_diretorio         = RS("ds_diretorio")
     w_pedagogica           = RS("ln_prop_pedagogica")
  End If

  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "Javascript"
  ValidateOpen "Validacao"
  If w_ea = "A" Then
     Validate "w_no_contato_internet", "Nome", "1", "1", "2", "35", "1", "1"
     Validate "w_ds_email_internet", "e-Mail", "1", "1", "6", "60", "1", "1"
     Validate "w_nr_fone_internet", "Telefone", "1", "1", "6", "20", "1", "1"
     Validate "w_nr_fax_internet", "Fax", "1", "", "6", "20", "1", "1"
     Validate "w_ds_texto_abertura", "Texto de abertura", "1", "1", "4", "8000", "1", "1"
     Validate "w_ds_institucional", "Texto da seção \""Quem somos\""", "1", "1", "4", "8000", "1", "1"
     If w_tipo = 2 Then ' Se for regional
        Validate "w_pedagogica", "Composição administrativa", "1", "", "5", "100", "1", "1"
     Else
        Validate "w_pedagogica", "Projeto", "1", "", "5", "100", "1", "1"
     End If
     ShowHTML " if (theForm.w_pedagogica.value > """"){"
     ShowHTML "    if((theForm.w_pedagogica.value.lastIndexOf('.PDF')==-1) && (theForm.w_pedagogica.value.lastIndexOf('.pdf')==-1) && (theForm.w_pedagogica.value.lastIndexOf('.DOC')==-1) && (theForm.w_pedagogica.value.lastIndexOf('.doc')==-1)) {"
     ShowHTML "       alert('Esolha arquivos com as extensões \'.doc\' ou \'.pdf\'!');"
     ShowHTML "       theForm.w_pedagogica.value=''; "
     ShowHTML "       theForm.w_pedagogica.focus(); "
     ShowHTML "       return false;"
     ShowHTML "    }"
     ShowHTML "  }"
     Validate "w_ds_mensagem", "\""Texto da mensagem em destaque\""", "1", "1", "4", "80", "1", "1"
  End If
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.Form.w_no_contato_internet.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Atualização de dados do site da unidade</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  ShowHTML "<FORM action=""Manut.asp?w_ew=Grava"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"" enctype=""multipart/form-data"">"
  ShowHTML MontaFiltro("POST")
  ShowHTML "<input type=""hidden"" name=""R"" value=""" & w_ew & """>"
  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Contato da unidade para divulgação no site</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Informe os dados da pessoa a ser exibida como contato da unidade para divulgação no site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b>No<u>m</u>e:</b><br><INPUT ACCESSKEY=""M"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_no_contato_internet"" size=""35"" maxlength=""35"" value=""" & w_no_contato_internet & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o nome da pessoa a ser exibida na seção \'Quem somos\' do site da unidade.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b><u>e</u>-Mail:</b><br><INPUT ACCESSKEY=""E"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_email_internet"" size=""40"" maxlength=""60"" value=""" & w_ds_email_internet & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o e-mail da pessoa.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b>Tele<u>f</u>one:</b><br><INPUT ACCESSKEY=""F"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_fone_internet"" size=""20"" maxlength=""20"" value=""" & w_nr_fone_internet & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o telefone da pessoa.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b>Fa<u>x</u>:</b><br><INPUT ACCESSKEY=""X"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_fax_internet"" size=""20"" maxlength=""20"" value=""" & w_nr_fax_internet & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o fax da pessoa.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        </table>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Página de abertura do site</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Informe o texto a ser colocado na página de abertura do site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=""1""><b>Texto da <u>p</u>ágina de abertura:</b><br><TEXTAREA ACCESSKEY=""P"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_texto_abertura"" rows=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o texto a ser exibido na página de abertura do site.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_texto_abertura & "</TEXTAREA></td>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Página ""Quem somos""</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Informe o texto ser colocado na página ""Quem somos"" do site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "      <tr><td><font size=""1""><b>Texto da seção ""<u>Q</u>uem somos"":</b><br><TEXTAREA ACCESSKEY=""Q"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_institucional"" rows=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o texto a ser exibido na seção \'Quem somos\' do site.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_institucional & "</TEXTAREA></td>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  If w_tipo = 2 Then ' Se for regional
     ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Página ""Composição administrativa""</td></td></tr>"
     ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
     ShowHTML "      <tr><td><font size=1>Informe o arquivo Word ou PDF a ser exibido na página ""Composição administrativa"" do site."
     ShowHTML "          <br><font color=""red""><b>IMPORTANTE: <a href=""sedf/orientacoes_word.pdf"" class=""hl"" target=""_blank"">Para documentos Word, clique aqui para ler as orientações sobre a formatação e a proteção do texto</a></b></font>."
     ShowHTML "      </font></td></tr>"
     ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
     ShowHTML "      <tr><td><font size=""1""><b>Composição adminis<u>t</u>rativa (arquivo Word ou PDF):</b><br><INPUT ACCESSKEY=""T"" " & w_Disabled & " class=""STI"" type=""file"" name=""w_pedagogica"" size=""60"" maxlength=""100"" value="""" ONMOUSEOVER=""popup('OPCIONAL. Clique no botão ao lado para localizar o arquivo que contém a composição administrativa da regional. Ele será transferido automaticamente para o servidor.','white')""; ONMOUSEOUT=""kill()"">"
  Else
     SQL = "select b.ds_especialidade from escEspecialidade_Cliente a inner join escEspecialidade b on (a.sq_codigo_espec = b.sq_especialidade and a." & CL & ")"
     ConectaBD SQL
     If Not RS.EOF Then
        If uCase(RS("ds_especialidade")) <> uCase("Biblioteca") Then
           ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Página ""Projeto""</td></td></tr>"
           ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
           ShowHTML "      <tr><td><font size=1>Informe o arquivo Word ou PDF a ser exibido na página ""Projeto"" do site."
           ShowHTML "          <br><font color=""red""><b>IMPORTANTE: <a href=""sedf/orientacoes_word.pdf"" class=""hl"" target=""_blank"">Para documentos Word, clique aqui para ler as orientações sobre a formatação e a proteção do texto</a></b></font>."
           ShowHTML "      </font></td></tr>"
           ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
           ShowHTML "      <tr><td><font size=""1""><b>Proje<u>t</u>o (arquivo Word):</b><br><INPUT ACCESSKEY=""T"" " & w_Disabled & " class=""STI"" type=""file"" name=""w_pedagogica"" size=""60"" maxlength=""100"" value="""" ONMOUSEOVER=""popup('OPCIONAL. Clique no botão ao lado para localizar o arquivo que contém o projeto da escola. Ele será transferido automaticamente para o servidor.','white')""; ONMOUSEOUT=""kill()"">"
        End If
     End If
     DesconectaBD
  End If
  If w_pedagogica > "" Then
     ShowHTML "              <b><a class=""SS"" href=""http://" & Replace(lCase(w_ds_diretorio),lCase("http://"),"") & "/" & w_pedagogica & """ target=""_blank"" title=""Clique para abrir o arquivo atual."">Exibir</a></b>"
  End If
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Diversos</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Informe os dados abaixo de exibição geral no site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=""1""><b>Texto da men<u>s</u>agem em destaque:</b><br><INPUT ACCESSKEY=""S"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_mensagem"" size=""80"" maxlength=""80"" value=""" & w_ds_mensagem & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe um texto que será exibido na parte superior do site, numa barra rolante.','white')""; ONMOUSEOUT=""kill()""></td>"
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

  Set w_pedagogica          = Nothing
  Set w_sq_cliente          = Nothing
  Set w_no_contato_internet = Nothing
  Set w_ds_email_internet   = Nothing
  Set w_nr_fone_internet    = Nothing
  Set w_nr_fax_internet     = Nothing
  Set w_ds_texto_abertura   = Nothing
  Set w_ds_institucional    = Nothing
  Set w_ds_mensagem         = Nothing

End Sub
REM =========================================================================
REM Fim da tela de dados do site
REM -------------------------------------------------------------------------

REM =========================================================================
REM Cadastro de notícias
REM -------------------------------------------------------------------------
Sub GetNoticiaCliente
  Dim w_chave, w_ds_titulo, w_in_ativo, w_ds_noticia
  Dim w_dt_noticia, w_in_exibe

  Dim w_troca, i, w_texto

  w_Chave           = Request("w_Chave")
  w_troca           = Request("w_troca")

  If w_troca > "" Then ' Se for recarga da página
     w_dt_noticia      = Request("w_dt_noticia")
     w_ds_titulo       = Request("w_ds_titulo")
     w_ds_noticia      = Request("w_ds_noticia")
     w_ln_noticia      = Request("w_ln_noticia")
     w_in_ativo        = Request("w_in_ativo")
     w_in_exibe        = Request("w_in_exibe")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select * from escNoticia_Cliente where " & replace(CL,"sq_cliente","sq_site_cliente") & " order by dt_noticia desc"
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escNoticia_Cliente where sq_noticia = " & w_chave
     ConectaBD SQL
     w_dt_noticia      = RS("dt_noticia")
     w_ds_titulo       = RS("ds_titulo")
     w_ds_noticia      = RS("ds_noticia")
     w_in_ativo        = RS("in_ativo")
     w_in_exibe        = RS("in_exibe")
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
        Validate "w_dt_noticia", "data", "DATA", "1", "10", "10", "1", "1"
        Validate "w_ds_titulo", "título", "", "1", "2", "50", "1", "1"
        Validate "w_ds_noticia", "descrição", "", "1", "2", "4000", "1", "1"
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
     BodyOpen "onLoad='document.Form.w_dt_noticia.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de notícias da unidade</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Data</font></td>"
    ShowHTML "          <td><font size=""1""><b>Título</font></td>"
    ShowHTML "          <td><font size=""1""><b>Ativo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("dt_noticia"),2)) & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("ds_titulo") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("in_ativo") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_noticia") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_chave=" & RS("sq_noticia") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>D</u>ata:</b><br><input accesskey=""D"" type=""text"" name=""w_dt_noticia"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(FormatDateTime(Nvl(w_dt_noticia,Date()),2)) & """ onKeyDown=""FormataData(this,event);"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a data da notícia. O sistema colocará as barras automaticamente.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>T</u>ítulo:</b><br><input " & w_Disabled & " accesskey=""T"" type=""text"" name=""w_ds_titulo"" class=""STI"" SIZE=""50"" MAXLENGTH=""50"" VALUE=""" & w_ds_titulo & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe um título para o noticia.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "        </table>"
    ShowHTML "      <tr><td valign=""top""><font size=""1""><b>D<u>e</u>scrição:</b><br><textarea " & w_Disabled & " accesskey=""E"" name=""w_ds_noticia"" class=""STI"" ROWS=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva a notícia.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_noticia & "</TEXTAREA></td>"
    ShowHTML "      <tr>"
    ShowHTML "      </tr>"
    ShowHTML "      <tr>"
    MontaRadioSN "<b>Exibir no site?</b>", w_in_ativo, "w_in_ativo"
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
    'ShowHTML " history.back(1);"
    ScriptClose
  End If
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape

  Set w_ds_titulo       = Nothing
  Set w_in_ativo        = Nothing
  Set w_ds_noticia      = Nothing
  Set w_ln_noticia      = Nothing

  Set w_chave           = Nothing

  Set w_troca           = Nothing
  Set i                 = Nothing
  Set w_texto           = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de notícias
REM -------------------------------------------------------------------------

REM =========================================================================
REM Cadastro de calendario
REM -------------------------------------------------------------------------
Sub GetCalendario
  Dim w_chave, w_dt_ocorrencia, w_ds_ocorrencia, w_ano

  Dim w_troca, i, w_texto

  w_Chave           = Request("w_Chave")
  w_troca           = Request("w_troca")

  If w_troca > "" Then ' Se for recarga da página
     w_dt_ocorrencia      = Request("w_dt_ocorrencia")
     w_ds_ocorrencia      = Request("w_ds_ocorrencia")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select * from escCalendario_Cliente where " & replace(CL,"sq_cliente","sq_site_cliente") & " order by year(dt_ocorrencia) desc, dt_ocorrencia"
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escCalendario_Cliente where sq_ocorrencia = " & w_chave
     ConectaBD SQL
     w_dt_ocorrencia   = RS("dt_ocorrencia")
     w_ds_ocorrencia   = RS("ds_ocorrencia")
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
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de calendário da unidade</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Data</font></td>"
    ShowHTML "          <td><font size=""1""><b>Ocorrência</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      w_ano  = ""
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        If wAno <> year(RS("dt_ocorrencia")) Then
           ShowHTML "      <tr bgcolor=""#C0C0C0"" valign=""top""><TD colspan=3 align=""center""><font size=2><B>" & year(RS("dt_ocorrencia")) & "</b></font></td></tr>"
           wAno = year(RS("dt_ocorrencia"))
        End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("dt_ocorrencia"),2)) & "</td>"
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
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>D</u>ata:</b><br><input accesskey=""D"" type=""text"" name=""w_dt_ocorrencia"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(FormatDateTime(Nvl(w_dt_ocorrencia,Date()),2)) & """ onKeyDown=""FormataData(this,event);"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a data de ocorrência. O sistema colocará as barras automaticamente.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "          <td valign=""top""><font size=""1""><b>D<u>e</u>scrição:</b><br><input " & w_Disabled & " accesskey=""E"" type=""text"" name=""w_ds_ocorrencia"" class=""STI"" SIZE=""60"" MAXLENGTH=""60"" VALUE=""" & w_ds_ocorrencia & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva a ocorrência.','white')""; ONMOUSEOUT=""kill()""></td>"
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

REM =========================================================================
REM Log do sistema
REM -------------------------------------------------------------------------
Sub ShowLog
  Dim w_chave, w_dt_ocorrencia, w_ds_ocorrencia, w_ano, P3, P4, w_data, w_ds_cliente

  Dim w_troca, i, w_texto

  w_Chave   = Request("w_Chave")
  w_troca   = Request("w_troca")
  P3        = Request("P3")
  P4        = Request("P4")

  ' Recupera o nome da escola
  SQL = "select * from escCliente where " & CL & ""
  ConectaBD SQL
  w_ds_cliente = RS("ds_cliente")
  DesconectaBD
  
  If w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select * from escCliente_Log where " & CL & " order by year(data) desc, month(data) desc, data desc"
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
        Validate "w_dt_ocorrencia", "data", "DATA", "1", "10", "10", "1", "1"
        Validate "w_ds_ocorrencia", "descriçao", "", "1", "2", "60", "1", "1"
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
  ShowHTML "<B><FONT COLOR=""#000000"">" & w_ds_cliente & " - Registro de ocorrências</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  If w_ea = "L" Then
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Data</font></td>"
    ShowHTML "          <td><font size=""1""><b>Hora</font></td>"
    ShowHTML "          <td><font size=""1""><b>Ocorrência</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=4 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else

      rs.PageSize     = P4
      rs.AbsolutePage = P3
      w_ano  = ""
      While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(P3)
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        If wAno <> year(RS("data")) & "/" & month(RS("data")) Then
           ShowHTML "      <tr bgcolor=""#C0C0C0"" valign=""top""><TD colspan=4 align=""center""><font size=2><B>" & month(RS("data")) & "/" & year(RS("data")) & "</b></font></td></tr>"
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
        ShowHTML "        <td><font size=""1"">" & RS("abrangencia") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "Log&R=" & w_Pagina & "Log&w_ea=V&CL=" & CL & "&w_chave=" & RS("sq_cliente_log") & """>Detalhes</A>&nbsp"
        ShowHTML "        </td>"
        ShowHTML "      </tr>"

        RS.MoveNext
      wend
    End If
    ShowHTML "      </center>"
    ShowHTML "    </table>"

    ShowHTML "<tr><td colspan=2><br><hr>"

    ShowHTML "<tr><td align=""center"" colspan=2>"
    MontaBarra "Manut.asp?w_ea=L&w_ew=" & w_ew, cDbl(RS.PageCount), cDbl(P3), cDbl(P4), cDbl(RS.RecordCount)
    ShowHTML "</tr>"

    ShowHTML "  </td>"
    ShowHTML "</tr>"
    DesconectaBD
  ElseIf Instr("IAEV",O) > 0 Then
    ' Recupera os dados da ocorrência selecionada
    SQL = "select * from escCliente_Log where sq_cliente_log = " & w_chave
    ConectaBD SQL

    ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
    ShowHTML "    <table width=""95%"" border=""0"">"
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
REM Cadastro de mensagens ao aluno
REM -------------------------------------------------------------------------
Sub GetMensagem
  Dim w_chave, w_dt_mensagem, w_ds_mensagem, w_sq_aluno

  Dim w_troca, i, w_texto

  w_Chave           = Request("w_Chave")
  w_texto           = Request("w_texto")
  w_troca           = Request("w_troca")

  If w_troca > "" Then ' Se for recarga da página
     w_dt_mensagem  = Request("w_dt_mensagem")
     w_ds_mensagem  = Request("w_ds_mensagem")
     w_sq_aluno     = Request("w_sq_aluno")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select a.*, b.no_aluno, b.nr_matricula " & VbCrLf & _
           "  from escMensagem_Aluno a " & VbCrLf & _
           "       inner join escAluno b on (a.sq_aluno = b.sq_aluno and b." & replace(CL,"sq_cliente","sq_site_cliente") & ") " & VbCrLf & _
           "order by a.dt_mensagem desc, b.no_aluno, a.in_lida " & VbCrLf
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escMensagem_Aluno where sq_mensagem = " & w_chave
     ConectaBD SQL
     w_dt_mensagem   = RS("dt_mensagem")
     w_ds_mensagem   = RS("ds_mensagem")
     w_sq_aluno      = RS("sq_aluno")
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
        Validate "w_dt_mensagem", "Data", "DATA", "1", "10", "10", "1", "1"
        Validate "w_ds_mensagem", "Mensagem", "", "1", "2", "4000", "1", "1"
        If w_ea = "I" Then
           Validate "w_sq_aluno", "Aluno", "SELECT", "1", "1", "10", "", "1"
        End If
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
     BodyOpen "onLoad='document.Form.w_dt_mensagem.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Mensagens a alunos da unidade</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Data</font></td>"
    ShowHTML "          <td><font size=""1""><b>Matrícula</font></td>"
    ShowHTML "          <td><font size=""1""><b>Aluno</font></td>"
    ShowHTML "          <td><font size=""1""><b>Lida</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"" ONMOUSEOVER=""popup('" & replace(replace(ExibeTexto(RS("ds_mensagem")),"'","\'"),"""","\'") & "','white')""; ONMOUSEOUT=""kill()"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("dt_mensagem"),2)) & "</td>"
        ShowHTML "        <td align=""center"" nowrap><font size=""1"">" & RS("nr_matricula") & "</td>"
        ShowHTML "        <td><font size=""1"">" & lcase(RS("no_aluno")) & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("in_lida") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_mensagem") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_chave=" & RS("sq_mensagem") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
    ShowHTML "<INPUT type=""hidden"" name=""w_troca"" value="""">"

    ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
    ShowHTML "    <table width=""95%"" border=""0"">"
    ShowHTML "      <tr><td><font size=""1""><b><u>D</u>ata:</b><br><input accesskey=""D"" type=""text"" name=""w_dt_mensagem"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(FormatDateTime(Nvl(w_dt_mensagem,Date()),2)) & """ onKeyDown=""FormataData(this,event);"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a data da mensagem. O sistema colocará as barras automaticamente.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "      <tr><td><font size=""1""><b>M<u>e</u>nsagem:</b><br><textarea " & w_Disabled & " accesskey=""E"" name=""w_ds_mensagem"" class=""STI"" ROWS=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva a mensagem.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_mensagem & "</TEXTAREA>"
    If w_ea = "I" Then
       ShowHTML "      <tr><td><font size=""1""><b><u>P</u>rocurar por:</b><br><input accesskey=""P"" type=""text"" name=""w_texto"" class=""STI"" SIZE=""50"" MAXLENGTH=""50"" VALUE=""" & w_texto & """ ONMOUSEOVER=""popup('Informe parte do nome ou da matrícula do aluno destinatário da mensagem.','white')""; ONMOUSEOUT=""kill()"">"
       ShowHTML "          <input type=""Button"" class=""STB"" name=""Pesquisa"" value=""Procurar"" onClick=""document.Form.w_troca.value='w_sq_aluno'; document.Form.action='" & w_pagina & w_ew & "'; document.Form.submit();""></td>"
       ShowHTML "      <tr><td><font size=""1""><b><u>A</u>luno:</b><br><select accesskey=""A"" name=""w_sq_aluno"" class=""STS"" SIZE=""1"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Selecione o aluno destinatário da mensagem na lista.','white')""; ONMOUSEOUT=""kill()"">"
       ShowHTML "          <OPTION VALUE="""">---"
       If w_texto > "" Then
          SQL = "select sq_aluno, nr_matricula, no_aluno " & VbCrLf & _
                "  from escAluno " & VbCrLf & _
                " where (upper(no_aluno) like '%" & uCase(w_texto) & "%' or " & VbCrLf & _
                "        nr_matricula like '%" & w_texto & "%') " & VbCrLf & _
                "   and " & replace(CL,"sq_cliente","sq_site_cliente") & " " & VbCrLf & _
                "order by no_aluno, nr_matricula" & VbCrLf
          ConectaBD SQL
          While Not RS.EOF
             ShowHTML "          <OPTION VALUE=""" & RS("sq_aluno") & """>" & RS("no_aluno") & " (" & RS("nr_matricula") & ")"
             RS.MoveNext
          Wend
          DesconectaBD
       End If
       ShowHTML "          </select>"
    Else
       SQL = "select nr_matricula, no_aluno from escAluno where sq_aluno = " & w_sq_aluno
       ConectaBD SQL
       ShowHTML "      <tr><td><font size=""1""><b>Aluno:<br>" & RS("no_aluno") & " (" & RS("nr_matricula") & ")</td>"
       DesconectaBD
    End If
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

  Set w_ds_mensagem = Nothing
  Set w_dt_mensagem = Nothing
  Set w_chave       = Nothing
  Set w_troca       = Nothing
  Set i             = Nothing
  Set w_texto       = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de mensagens ao aluno
REM -------------------------------------------------------------------------

REM =========================================================================
REM Procedimento que executa as operações de BD
REM -------------------------------------------------------------------------
Public Sub Grava

  Dim w_chave, w_imagem, w_file, w_sql, w_complemento, w_funcionalidade

  Cabecalho
  ShowHTML "</HEAD>"
  BodyOpen "onLoad=document.focus();"

  ' Recupera o código a ser gravado na tabela de log
  
  If w_R <> conWhatSGE Then
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
             "   dt_alteracao    = getdate(), " & VbCrLf
       'If Request("w_ds_apelido") > "" Then SQL = SQL & "   ds_apelido      = '" & trim(Request("w_ds_apelido")) & "', " & VbCrLf  Else SQL = SQL & "   ds_apelido      = null, " & VbCrLf  End If
       If Request("w_ds_email") > ""   Then SQL = SQL & "   ds_email        = '" & trim(Request("w_ds_email")) & "' " & VbCrLf   Else SQL = SQL & "   ds_email        = null " & VbCrLf End If
       SQL = SQL & _
             "where " & CL & VbCrLf
       ExecutaSQL(SQL)

       ' Grava o acesso na tabela de log
       w_sql = SQL
       SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
             "values ( " & VbCrLf & _
             "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
             "         getdate(), " & VbCrLf & _
             "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
             "         2, " & VbCrLf & _
             "         'Atualização de dados básicos.', " & VbCrLf & _
             "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
             "         " & w_funcionalidade & " " & VbCrLf & _
             "       ) " & VbCrLf
       ExecutaSQL(SQL)
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=A&CL=" & CL & "';"
       ScriptClose

    Case conWhatDadosAdicionais
       dbms.BeginTrans()
       
       SQL = "select count(*) existe from escCliente_Dados where " & CL
       ConectaBD SQL
       If RS("existe") = 0 Then
          w_chave = replace(CL,"sq_cliente=","")
          'Criação das escolas na tabela que contém dados complementares'
          SQL = "insert into escCliente_Dados ( " & VbCrLf &_ 
                "     sq_cliente_dados, sq_cliente,       nr_cnpj,        tp_registro, " & VbCrLf &_ 
                "     ds_ato,           nr_ato,           dt_ato,         ds_orgao, " & VbCrLf &_ 
                "     no_bairro,        ds_email_contato, nr_fax_contato, no_diretor, " & VbCrLf &_ 
                "     no_secretario,    ds_logradouro,    nr_cep,         no_contato," & VbCrLf &_ 
                "     nr_fone_contato" & VbCrLf &_ 
                " ) values ( " & VbCrLf &_   
                "     " & w_chave & "," & VbCrLf & _
                "     " & w_chave & "," & VbCrLf
          If Request("w_nr_cnpj") > ""          Then SQL = SQL & "   '" & Request("w_nr_cnpj") & "', "                 & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_tp_registro") > ""      Then SQL = SQL & "   '" & Request("w_tp_registro") & "', "             & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_ds_ato") > ""           Then SQL = SQL & "   '" & trim(Request("w_ds_ato")) & "', "            & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_nr_ato") > ""           Then SQL = SQL & "   '" & trim(Request("w_nr_ato")) & "', "            & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_dt_ato") > ""           Then SQL = SQL & "   '" & cDate(trim(Request("w_dt_ato"))) & "', "     & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_ds_orgao") > ""         Then SQL = SQL & "   '" & trim(Request("w_ds_orgao")) & "', "          & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_no_bairro") > ""        Then SQL = SQL & "   '" & trim(Request("w_no_bairro")) & "', "         & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_ds_email_contato") > "" Then SQL = SQL & "   '" & trim(Request("w_ds_email_contato")) & "', "  & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_nr_fax_contato") > ""   Then SQL = SQL & "   '" & trim(Request("w_nr_fax_contato")) & "', "    & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_no_diretor") > ""       Then SQL = SQL & "   '" & trim(Request("w_no_diretor")) & "', "        & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          If Request("w_no_secretario") > ""    Then SQL = SQL & "   '" & trim(Request("w_no_secretario")) & "', "     & VbCrLf Else SQL = SQL & "   null, " & VbCrLf End If
          sql = sql & _
                "   '" & trim(Request("w_ds_logradouro")) & "', " & VbCrLf & _
                "   '" & Request("w_nr_cep") & "', " & VbCrLf & _
                "   '" & trim(Request("w_no_contato")) & "', " & VbCrLf & _
                "   '" & trim(Request("w_nr_fone_contato")) & "' " & VbCrLf & _
                ")" & VbCrLf
          ExecutaSQL(SQL)

       Else
          SQL = "update escCliente_Dados set " & VbCrLf & _
                "   ds_logradouro  = '" & trim(Request("w_ds_logradouro")) & "', " & VbCrLf & _
                "   nr_cep           = '" & Request("w_nr_cep") & "', " & VbCrLf & _
                "   no_contato     = '" & trim(Request("w_no_contato")) & "', " & VbCrLf & _
                "   nr_fone_contato= '" & trim(Request("w_nr_fone_contato")) & "', " & VbCrLf
          If Request("w_no_bairro") > ""        Then SQL = SQL & "   no_bairro          = '" & trim(Request("w_no_bairro")) & "', "      & VbCrLf Else SQL = SQL & "   no_bairro        = null, " & VbCrLf End If
          If Request("w_no_diretor") > ""       Then SQL = SQL & "   no_diretor       = '" & trim(Request("w_no_diretor")) & "', "       & VbCrLf Else SQL = SQL & "   no_diretor       = null, " & VbCrLf End If
          If Request("w_no_secretario") > ""    Then SQL = SQL & "   no_secretario    = '" & trim(Request("w_no_secretario")) & "', "    & VbCrLf Else SQL = SQL & "   no_secretario    = null, " & VbCrLf End If
          If Request("w_nr_cnpj") > ""          Then SQL = SQL & "   nr_cnpj            = '" & Request("w_nr_cnpj") & "', "              & VbCrLf Else SQL = SQL & "   nr_cnpj          = null, " & VbCrLf End If
          If Request("w_tp_registro") > ""      Then SQL = SQL & "   tp_registro      = '" & Request("w_tp_registro") & "', "            & VbCrLf Else SQL = SQL & "   tp_registro      = null, " & VbCrLf End If
          If Request("w_ds_ato") > ""           Then SQL = SQL & "   ds_ato             = '" & trim(Request("w_ds_ato")) & "', "         & VbCrLf Else SQL = SQL & "   ds_ato           = null, " & VbCrLf End If
          If Request("w_nr_ato") > ""           Then SQL = SQL & "   nr_ato             = '" & trim(Request("w_nr_ato")) & "', "         & VbCrLf Else SQL = SQL & "   nr_ato           = null, " & VbCrLf End If
          If Request("w_dt_ato") > ""           Then SQL = SQL & "   dt_ato             = '" & cDate(trim(Request("w_dt_ato"))) & "', "  & VbCrLf Else SQL = SQL & "   dt_ato           = null, " & VbCrLf End If
          If Request("w_ds_orgao") > ""         Then SQL = SQL & "   ds_orgao           = '" & trim(Request("w_ds_orgao")) & "', "       & VbCrLf Else SQL = SQL & "   ds_orgao         = null, " & VbCrLf End If
          If Request("w_ds_email_contato") > "" Then SQL = SQL & "   ds_email_contato = '" & trim(Request("w_ds_email_contato")) & "', " & VbCrLf Else SQL = SQL & "   ds_email_contato = null, " & VbCrLf End If
          If Request("w_nr_fax_contato") > ""   Then SQL = SQL & "   nr_fax_contato   = '" & trim(Request("w_nr_fax_contato")) & "' "    & VbCrLf Else SQL = SQL & "   nr_fax_contato   = null "  & VbCrLf End If
          SQL = SQL & _
                "where " & CL & VbCrLf
          ExecutaSQL(SQL)
       End If
       w_sql = SQL
       SQL = "update escCliente set dt_alteracao = getdate() where " & CL & VbCrLf
       ExecutaSQL(SQL)

       ' Grava o acesso na tabela de log
       SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
             "values ( " & VbCrLf & _
             "           " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
             "           getdate(), " & VbCrLf & _
             "           '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
             "           2, " & VbCrLf & _
             "           'Atualização de dados adicionais.', " & VbCrLf & _
             "           '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
             "           " & w_funcionalidade & " " & VbCrLf & _
             "       ) " & VbCrLf
       ExecutaSQL(SQL)
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=A&cl=" & cl & "';"
       ScriptClose

    Case conWhatSite

       SQL = "select a.ds_username, b.ds_username pai " & VbCrLf & _
             "  from escCliente a, " & VbCrLf & _
             "       escCliente b " & VbCrLf & _
             " where b.sq_cliente_pai is null " & VbCrLf & _
             "   and a." & CL
       ConectaBD SQL
       w_diretorio = replace(conFilePhysical & "\" & RS("pai") &  "\" & RS("ds_username") & "\","\\","\")

       dbms.BeginTrans()
       SQL = "update escCliente_Site set " & VbCrLf & _
             "   no_contato_internet    = '" & trim(ul.Form("w_no_contato_internet")) & "', " & VbCrLf & _
             "   ds_email_internet      = '" & trim(ul.Form("w_ds_email_internet")) & "', " & VbCrLf & _
             "   nr_fone_internet       = '" & trim(ul.Form("w_nr_fone_internet")) & "', " & VbCrLf & _
             "   ds_texto_abertura      = '" & trim(ul.Form("w_ds_texto_abertura")) & "', " & VbCrLf & _
             "   ds_institucional       = '" & trim(ul.Form("w_ds_institucional")) & "', " & VbCrLf
       If ul.Form("w_nr_fax_internet") > ""   Then SQL = SQL & "   nr_fax_internet        = '" & trim(ul.Form("w_nr_fax_internet")) & "', "       & VbCrLf Else SQL = SQL & "   nr_fax_internet        = null, "  & VbCrLf End If
       If ul.Form("w_ds_mensagem") > ""       Then SQL = SQL & "   ds_mensagem            = '" & trim(ul.Form("w_ds_mensagem")) & "' "            & VbCrLf Else SQL = SQL & "   ds_mensagem            = null "   & VbCrLf End If
       SQL = SQL & _
             "where " & CL & VbCrLf
       ExecutaSQL(SQL)
       w_sql = SQL

       SQL = "update escCliente set dt_alteracao = getdate() where " & CL & VbCrLf
       ExecutaSQL(SQL)

       If ul.Files("w_pedagogica").Size > 0 Then
          ul.Files("w_pedagogica").SaveAs(w_diretorio & extractFileName(ul.Files("w_pedagogica").OriginalPath))
          w_imagem = extractFileName(ul.Files("w_pedagogica").OriginalPath)

          SQL = "update escCliente_Site set " & VbCrLf & _
                "   ln_prop_pedagogica = '" & w_imagem & "' " & VbCrLf & _
                "where " & CL & VbCrLf
          ExecutaSQL(SQL)

       End If
       ' Grava o acesso na tabela de log
       SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
             "values ( " & VbCrLf & _
             "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
             "         getdate(), " & VbCrLf & _
             "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
             "         2, " & VbCrLf & _
             "         'Atualização de dados do site.', " & VbCrLf & _
             "         '" & replace(Mid(w_sql,1,2000),"'", "''") & "', " & VbCrLf & _
             "         " & w_funcionalidade & " " & VbCrLf & _
             "       ) " & VbCrLf
       ExecutaSQL(SQL)
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & w_R & "&w_ea=A&cl=" & cl & "';"
       ScriptClose

    Case conWhatAdmin
       dbms.BeginTrans()
       ' Apaga os equipamentos existentes para a unidade
       SQL = "delete escCliente_Equipamento where " & CL & VbCrLf
       w_sql = SQL & VbCrLf
       ExecutaSQL(SQL)

       ' Atualiza os dados em escCliente_Admin
       SQL = "select count(*) existe from escCliente_Admin where " & CL & VbCrLf
       ConectaBD SQL
       If RS("existe") = 0 Then
          w_oper = 1
          SQL = "insert into escCliente_Admin (sq_cliente, limpeza_terceirizada, merenda_terceirizada, banheiro) " & VbCrLf & _
                " values (" & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         '" & Request("w_limpeza") & "', " & VbCrLf & _
                "         '" & Request("w_merenda") & "', " & VbCrLf & _
                "         '" & Request("w_banheiro") & "' " & VbCrLf & _
                "        )" & VbCrLf
          w_sql = SQL & VbCrLf
       Else
          w_oper = 2
          SQL = "update escCliente_Admin set " & VbCrLf & _
                "   limpeza_terceirizada = '" & Request("w_limpeza") & "', " & VbCrLf & _
                "   merenda_terceirizada = '" & Request("w_merenda") & "', " & VbCrLf & _
                "   banheiro             = '" & Request("w_banheiro") & "' " & VbCrLf & _
                " where " & CL & " " & VbCrLf
          w_sql = SQL & VbCrLf
       End If
       ExecutaSQL(SQL)

       ' Em seguida, grava os equipamentos com quantidade maior que zero
       For w_cont = 1 To Request.Form("w_equipamento").Count
          If Nvl(Request("w_equipamento")(w_cont),0) > 0 Then
             SQL = "insert into escCliente_Equipamento (sq_cliente, sq_equipamento, quantidade) " & VbCrLf & _
                   "values (" & VbCrLf & _
                   "   " & Replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                   "   " & Request("w_codigo")(w_cont) & ", " & VbCrLf & _
                   "   " & Request("w_equipamento")(w_cont) & " " & VbCrLf & _
                   ")"
             ExecutaSQL(SQL)
             w_sql = w_sql & SQL & VbCrLf
          End If
       Next

       ' Grava o acesso na tabela de log
       SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
             "values ( " & VbCrLf & _
             "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
             "         getdate(), " & VbCrLf & _
             "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
             "         " & w_oper & ", " & VbCrLf & _
             "         'Informações administrativas da escola.', " & VbCrLf & _
             "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
             "         " & w_funcionalidade & " " & VbCrLf & _
             "       ) " & VbCrLf
       ExecutaSQL(SQL)

       SQL = "update escCliente set dt_alteracao = getdate() where " & CL & VbCrLf
       ExecutaSQL(SQL)

       dbms.CommitTrans()
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
    Case conWhatEspecialidadeCliente
       dbms.BeginTrans()
       ' Apaga as especialidades existentes para a unidade
       SQL = "delete escEspecialidade_Cliente where " & CL & VbCrLf
       w_sql = SQL & VbCrLf
       ExecutaSQL(SQL)
       ' Em seguida, cria apenas as especialidades indicadas pelo cliente
       For w_cont = 1 To Request.Form("w_chave").Count
          SQL = "select IsNull(max(sq_codigo_cli),0)+1 chave from escEspecialidade_Cliente" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD
          SQL = "insert into escEspecialidade_Cliente (sq_codigo_cli, sq_cliente, sq_codigo_espec) " & VbCrLf & _
                "values (" & VbCrLf & _
                "   " & w_chave & ", " & VbCrLf & _
                "   " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "   " & Request("w_chave")(w_cont) & " " & VbCrLf & _
                ")"
          ExecutaSQL(SQL)
          w_sql = w_sql & SQL & VbCrLf
       Next

       SQL = "update escCliente set dt_alteracao = getdate() where " & CL & VbCrLf
       ExecutaSQL(SQL)

       ' Grava o acesso na tabela de log

       SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
             "values ( " & VbCrLf & _
             "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
             "         getdate(), " & VbCrLf & _
             "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
             "         2, " & VbCrLf & _
             "         'Atualização das modalidades de ensino.', " & VbCrLf & _
             "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
             "         " & w_funcionalidade & " " & VbCrLf & _
             "       ) " & VbCrLf
       ExecutaSQL(SQL)
       dbms.CommitTrans()
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
    Case conWhatDocumento
       SQL = "select a.ds_username, b.ds_username pai " & VbCrLf & _
             "  from escCliente a, " & VbCrLf & _
             "       escCliente b " & VbCrLf & _
             " where b.sq_cliente_pai is null " & VbCrLf & _
             "   and a." & CL
       ConectaBD SQL
       w_diretorio = replace(conFilePhysical & "\" & RS("pai") &  "\" & RS("ds_username") & "\","\\","\")
       DesconectaBD

       dbms.BeginTrans()
       If w_ea = "I" Then
          ' Verifica se o nome desse arquivo já existe
          SQL = "select count(*) existe from escCliente_Arquivo where ln_arquivo = '" & extractFileName(ul.Files("w_ln_arquivo").OriginalPath) & "'" & VbCrLf
          ConectaBD SQL
          If RS("existe") > 0 Then
             ScriptOpen "JavaScript"
             ShowHTML "  alert('Nome do arquivo já existe!');"
             ShowHTML "  history.back(1);"
             ScriptClose
             Exit Sub
          End If
          DesconectaBD
          
          ul.Files("w_ln_arquivo").SaveAs(w_diretorio & extractFileName(ul.Files("w_ln_arquivo").OriginalPath))
          w_imagem = extractFileName(ul.Files("w_ln_arquivo").OriginalPath)

          ' Recupera o valor da próxima chave primária
          SQL = "select IsNull(max(sq_arquivo),0)+1 chave from escCliente_Arquivo" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD

          ' Insere o arquivo
          SQL = " insert into escCliente_Arquivo " & VbCrLf & _
                "    (sq_arquivo, sq_site_cliente, dt_arquivo, ds_arquivo, ln_arquivo, " & VbCrLf & _
                "     in_ativo,   in_destinatario, nr_ordem,   ds_titulo) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf & _
                "     " & ul.Form("w_sq_cliente") & ", " & VbCrLf & _
                "     convert(datetime, '" & FormataDataEdicao(FormatDateTime(Date(),2)) & "',103), " & VbCrLf & _
                "     '" & ul.Form("w_ds_arquivo") & "', " & VbCrLf & _
                "     '" & w_imagem & "', " & VbCrLf & _
                "     '" & ul.Form("w_in_ativo") & "', " & VbCrLf & _
                "     '" & ul.Form("w_in_destinatario") & "', " & VbCrLf & _
                "     '" & ul.Form("w_nr_ordem") & "', " & VbCrLf & _
                "     '" & ul.Form("w_ds_titulo") & "' " & VbCrLf & _
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
                "         'Usuário """ & uCase(Session("username")) & """ - inclusão de arquivo.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
             "         " & w_funcionalidade & " " & VbCrLf & _
             "       ) " & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "A" Then
          ' Verifica se o nome desse arquivo já existe
          SQL = "select count(*) existe from escCliente_Arquivo where ln_arquivo = '" & extractFileName(ul.Files("w_ln_arquivo").OriginalPath) & "' and sq_arquivo <> " & ul.Form("w_chave") & VbCrLf
          ConectaBD SQL
          If RS("existe") > 0 Then
             ScriptOpen "JavaScript"
             ShowHTML "  alert('Nome do arquivo já existe!');"
             ShowHTML "  history.back(1);"
             ScriptClose
             Exit Sub
          End If
          DesconectaBD
                    
          ' Remove o arquivo físico
          SQL = "select ln_arquivo arquivo from escCliente_Arquivo where sq_arquivo = " & ul.Form("w_chave")
          ConectaBD SQL
          w_arquivo = RS("arquivo")
          DesconectaBD

          SQL = " update escCliente_Arquivo set " & VbCrLf & _
                "     ds_titulo       = '" & ul.Form("w_ds_titulo") & "', " & VbCrLf & _
                "     dt_arquivo      = convert(datetime, '" & FormataDataEdicao(FormatDateTime(Date(),2)) & "',103), " & VbCrLf & _
                "     ds_arquivo      = '" & ul.Form("w_ds_arquivo") & "', " & VbCrLf & _
                "     in_ativo        = '" & ul.Form("w_in_ativo") & "', " & VbCrLf & _
                "     in_destinatario = '" & ul.Form("w_in_destinatario") & "', " & VbCrLf & _
                "     nr_ordem        = '" & ul.Form("w_nr_ordem") & "' " & VbCrLf & _
                "where sq_arquivo = " & ul.Form("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         2, " & VbCrLf & _
                "         'Alteração de arquivo.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)

          If ul.Files("w_ln_arquivo").Size > 0 Then
             ' Remove o arquivo físico
             DeleteAFile w_diretorio & w_arquivo

             ul.Files("w_ln_arquivo").SaveAs(w_diretorio & extractFileName(ul.Files("w_ln_arquivo").OriginalPath))
             w_imagem = extractFileName(ul.Files("w_ln_arquivo").OriginalPath)

             SQL = " update escCliente_Arquivo set " & VbCrLf & _
                   "     ln_arquivo      = '" & w_imagem & "' " & VbCrLf & _
                   "where sq_arquivo = " & ul.Form("w_chave") & VbCrLf
             ExecutaSQL(SQL)
          End If

       ElseIf w_ea = "E" Then
          ' Remove o arquivo físico
          SQL = "select ln_arquivo from escCliente_Arquivo where sq_arquivo = " & ul.Form("w_chave")
          ConectaBD SQL
          DeleteAFile w_diretorio & RS("ln_arquivo")
          DesconectaBD

          SQL = " delete escCliente_Arquivo where sq_arquivo = " & ul.Form("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         3, " & VbCrLf & _
                "         'Exclusão de arquivo.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       End If

       SQL = "update escCliente set dt_alteracao = getdate() where " & CL & VbCrLf
       ExecutaSQL(SQL)
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & w_R & "&w_ea=L&cl=" & cl & "';"
       ScriptClose

    Case conWhatCalendario
       dbms.BeginTrans()

       If w_ea = "I" Then
          ' Recupera o valor da próxima chave primária
          SQL = "select IsNull(max(sq_ocorrencia),0)+1 chave from escCalendario_Cliente" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD

          ' Insere o arquivo
          SQL = " insert into escCalendario_Cliente " & VbCrLf & _
                "    (sq_ocorrencia, sq_site_cliente, dt_ocorrencia, ds_ocorrencia) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf & _
                "     " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "     convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_ocorrencia"),2)) & "',103), " & VbCrLf & _
                "     '" & Request("w_ds_ocorrencia") & "' " & VbCrLf & _
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
       ElseIf w_ea = "A" Then
          SQL = " update escCalendario_Cliente set " & VbCrLf & _
                "     dt_ocorrencia  = convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_ocorrencia"),2)) & "',103), " & VbCrLf & _
                "     ds_ocorrencia  = '" & Request("w_ds_ocorrencia") & "' " & VbCrLf & _
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

    Case conWhatNotCliente
       dbms.BeginTrans()

       If w_ea = "I" Then
          ' Recupera o valor da próxima chave primária
          SQL = "select IsNull(max(sq_noticia),0)+1 chave from escNoticia_Cliente" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD

          ' Insere o arquivo
          SQL = " insert into escNoticia_Cliente " & VbCrLf & _
                "    (sq_noticia, sq_site_cliente, dt_noticia, ds_titulo, ds_noticia, in_ativo) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf & _
                "     " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "     convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_noticia"),2)) & "',103), " & VbCrLf & _
                "     '" & Request("w_ds_titulo") & "', " & VbCrLf & _
                "     '" & Request("w_ds_noticia") & "', " & VbCrLf & _
                "     '" & Request("w_in_ativo") & "' " & VbCrLf & _
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
                "         'Inclusão de notícia da escola.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "A" Then
          SQL = " update escNoticia_Cliente set " & VbCrLf & _
                "     dt_noticia  = convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_noticia"),2)) & "',103), " & VbCrLf & _
                "     ds_titulo   = '" & Request("w_ds_titulo") & "', " & VbCrLf & _
                "     ds_noticia  = '" & Request("w_ds_noticia") & "', " & VbCrLf & _
                "     in_ativo    = '" & Request("w_in_ativo") & "' " & VbCrLf & _
                "where sq_noticia = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         2, " & VbCrLf & _
                "         'Alteração de notícia da escola.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "E" Then
          SQL = " delete escNoticia_Cliente where sq_noticia = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         3, " & VbCrLf & _
                "         'Exclusão de notícia da escola.', " & VbCrLf & _
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

    Case conWhatMensagem
       dbms.BeginTrans()

       If w_ea = "I" Then
          ' Recupera o valor da próxima chave primária
          SQL = "select IsNull(max(sq_mensagem),0)+1 chave from escMensagem_Aluno" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD

          ' Insere o arquivo
          SQL = " insert into escMensagem_Aluno " & VbCrLf & _
                "    (sq_mensagem, sq_aluno, dt_mensagem, ds_mensagem, in_lida) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf & _
                "     " & Request("w_sq_Aluno") & ", " & VbCrLf & _
                "     convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_mensagem"),2)) & "',103), " & VbCrLf & _
                "     '" & Request("w_ds_mensagem") & "', " & VbCrLf & _
                "     'Não' " & VbCrLf & _
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
                "         'Inclusão de mensagem para aluno da escola.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "A" Then
          SQL = " update escMensagem_Aluno set " & VbCrLf & _
                "     dt_mensagem  = convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_mensagem"),2)) & "',103), " & VbCrLf & _
                "     ds_mensagem  = '" & Request("w_ds_mensagem") & "' " & VbCrLf & _
                "where sq_mensagem = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         2, " & VbCrLf & _
                "         'Alteração de mensagem para aluno da escola.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "E" Then
          SQL = " delete escMensagem_Aluno where sq_mensagem = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & replace(CL,"sq_cliente=","") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         3, " & VbCrLf & _
                "         'Exclusão de mensagem para aluno da escola.', " & VbCrLf & _
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
    
    Case conWhatSGE
       SQL = "select a.ds_username, b.ds_username pai " & VbCrLf & _
             "  from escCliente a, " & VbCrLf & _
             "       escCliente b " & VbCrLf & _
             " where b.sq_cliente_pai is null " & VbCrLf & _
             "   and a." & CL
       ConectaBD SQL
       w_diretorio = replace(conFilePhysical & "\" & RS("pai") &  "\" & RS("ds_username") & "\","\\","\")
       DesconectaBD
       dbms.BeginTrans()
       If w_ea = "I" Then
          ' Verifica se o nome desse arquivo já existe
          SQL = "select count(*) existe from escCliente_Foto where ln_foto = '" & extractFileName(ul.Files("w_ln_foto").OriginalPath) & "'" & VbCrLf
          ConectaBD SQL
          If RS("existe") > 0 Then
             ScriptOpen "JavaScript"
             ShowHTML "  alert('Nome do arquivo já existe!');"
             ShowHTML "  history.back(1);"
             ScriptClose
             Exit Sub
          End If
          DesconectaBD
          ul.Files("w_ln_foto").SaveAs(w_diretorio & extractFileName(ul.Files("w_ln_foto").OriginalPath))
          w_imagem = extractFileName(ul.Files("w_ln_foto").OriginalPath)

          ' Recupera o valor da próxima chave primária
          SQL = "select IsNull(max(sq_cliente_foto),0)+1 chave from escCliente_Foto" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD

          ' Insere o arquivo
          SQL = " insert into escCliente_Foto " & VbCrLf & _
                "    (sq_cliente_foto, sq_cliente, ds_foto, ln_foto, tp_foto, nr_ordem) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf & _
                "     " & ul.Form("w_sq_cliente") & ", " & VbCrLf & _
                "     '" & ul.Form("w_ds_foto") & "', " & VbCrLf & _
                "     '" & w_imagem & "', " & VbCrLf & _
                "     '" & ul.Form("w_tp_foto") & "', " & VbCrLf & _
                "     '" & ul.Form("w_nr_ordem") & "' " & VbCrLf & _
                " )" & VbCrLf
          ExecutaSQL(SQL)

       ElseIf w_ea = "A" Then
          ' Verifica se o nome desse arquivo já existe
          SQL = "select count(*) existe from escCliente_Foto where ln_foto = '" & extractFileName(ul.Files("w_ln_foto").OriginalPath) & "' and sq_cliente_foto <> " & ul.Form("w_chave") & VbCrLf
          ConectaBD SQL
          If RS("existe") > 0 Then
             ScriptOpen "JavaScript"
             ShowHTML "  alert('Nome do arquivo já existe!');"
             ShowHTML "  history.back(1);"
             ScriptClose
             Exit Sub
          End If
          DesconectaBD
          
          ' Remove o arquivo físico
          SQL = "select ln_foto foto from escCliente_Foto where sq_cliente_foto = " & ul.Form("w_chave")
          ConectaBD SQL
          w_foto = RS("foto")
          DesconectaBD

          SQL = " update escCliente_Foto set " & VbCrLf & _
                "     ds_foto       = '" & ul.Form("w_ds_foto") & "', " & VbCrLf & _
                "     nr_ordem      = '" & ul.Form("w_nr_ordem") & "' " & VbCrLf & _
                "where sq_cliente_foto = " & ul.Form("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If ul.Files("w_ln_foto").Size > 0 Then
             ' Remove o arquivo físico
             DeleteAFile w_diretorio & w_foto

             ul.Files("w_ln_foto").SaveAs(w_diretorio & extractFileName(ul.Files("w_ln_foto").OriginalPath))
             w_imagem = extractFileName(ul.Files("w_ln_foto").OriginalPath)

             SQL = " update escCliente_Foto set " & VbCrLf & _
                   "     ln_foto      = '" & w_imagem & "' " & VbCrLf & _
                   "where sq_cliente_foto = " & ul.Form("w_chave") & VbCrLf
             ExecutaSQL(SQL)
          End If

       ElseIf w_ea = "E" Then
          ' Remove o arquivo físico
          SQL = "select ln_foto from escCliente_Foto where sq_cliente_foto = " & ul.Form("w_chave")
          ConectaBD SQL
          DeleteAFile w_diretorio & RS("ln_foto")
          DesconectaBD

          SQL = " delete escCliente_Foto where sq_cliente_foto = " & ul.Form("w_chave") & VbCrLf
          ExecutaSQL(SQL)

       End If

       SQL = "update escCliente set dt_alteracao = getdate() where " & CL & VbCrLf
       ExecutaSQL(SQL)
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & w_R & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
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
Private Sub MainBody

  If Nvl(w_IN,0) = 1 then
    ShowFrames
  ElseIf CL > "" Then
    Select Case uCase(w_EW)
      Case "SHOWMENU"                   showMenu
      Case conWhatAdmin                 GetAdmin
      Case conWhatSGE                   GetSGE
      Case conWhatCliente               GetCliente
      Case conWhatDadosAdicionais       GetDadosAdicionais
      Case conWhatSite                  GetSite
      Case conWhatNotCliente            GetNoticiaCliente
      Case conWhatCalendario            GetCalendario
      Case conWhatEspecialidadeCliente  GetEspecialidadeCliente
      Case conWhatDocumento             GetDocumento
      Case conWhatCalendario            GetCalendario
      Case conWhatMensagem              GetMensagem
      Case "LOG"                        ShowLog
      Case "TEST"                       Response.Write "ok"
      Case "GRAVA"                      Grava
      Case Else
           If ( Not Request.QueryString( conToMakeSystem ) > "" ) Then
              ShowFrames
           End If
    End Select

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do Manut.asp
REM =========================================================================
%>