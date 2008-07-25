<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/Constants_ADO.inc"-->
<!--#INCLUDE VIRTUAL="/modelos/Constants.inc"-->
<!--#INCLUDE VIRTUAL="/esc.inc"-->
<!--#INCLUDE VIRTUAL="/Funcoes.asp"-->
<!--#INCLUDE VIRTUAL="/JScript.asp"-->
<%
Response.Expires = -1500
REM =========================================================================
REM  /Default.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: Exibição dos dados da escola
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 16/03/2000 14:10PM
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Informática
REM -------------------------------------------------------------------------

Private RS, RS1, sql
Private RSCabecalho
Private RSRodape

Private dbms

Private sstrSN
Private sstrCL

sstrSN = "Default.Asp"
sstrCL = "Aluno.Asp"

Public sstrEF
Public sstrEW
Public sstrIN
Public CL

CL = Request("CL")
sstrEF = "sq_cliente=" & CL
sstrEW = Request("EW")
sstrIN = Request("IN")
w_dir  = "Modelos/Mod15/"

Private sstrData

Set RS          = Server.CreateObject("ADODB.RecordSet")
Set RS1         = Server.CreateObject("ADODB.RecordSet")
Set RSCabecalho = Server.CreateObject("ADODB.RecordSet")
Set RSRodape    = Server.CreateObject("ADODB.RecordSet")
Set dbms        = Server.CreateObject("ADODB.Connection")

dbms.ConnectionTimeout = 300
dbms.CommandTimeout = 300
dbms.Open conConnectionString

Server.ScriptTimeOut = conScriptTimeout

ShowHTML "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
ShowHTML "<html xmlns=""http://www.w3.org/1999/xhtml"">"
ShowHTML "<head>"
ShowHTML "   <title>Secretaria de Estado de Educa&ccedil;&atilde;o</title>"
ShowHTML "   <meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> "
ShowHTML "   <link href=""/css/estilo.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <link href=""/css/print.css""  media=""print""  rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
ShowHTML "</head>"

ShowHTML "<BASE HREF=""" & conSite & "/"">"
If (Request("H") > "") or (Request("Q") > "") or (Request("p_regional") > "") or (wIN > 0) Then
   ShowHTML "<body onLoad=""location.href='#resultado'"">"
Else
   ShowHTML "<body>"
End If
ShowHTML "<div id=""container"">"
ShowHTML "  <div id=""cab"">"
ShowHTML "    <div id=""cabtopo"">"
ShowHTML "      <div id=""logoesq""><img src=""img/fundo_logoesq.gif"" border=0></div>"
ShowHTML "      <div id=""logodir""><a href=""http://www.se.df.gov.br""><img src=""img/fundo_logodir.jpg"" border=0></a></div>"
ShowHTML "    </div>"

sql = "SELECT a.sq_cliente, a.ds_cliente, b.ds_mensagem FROM escCliente a inner join esccliente_site b on (a.sq_cliente = b.sq_cliente)WHERE a." & sstrEF
ConectaBD SQL
ShowHTML "    <div id=""cabbase"">"
ShowHTML "      <div id=""busca"" valign=""center""><marquee width=""100%"" align=""middle""><font color=""white"" face=""Arial"" size=""2""><b>" & RS("ds_mensagem") & "</b></font></marquee></div> "
ShowHTML "    </div>"
ShowHTML "  </div>"
ShowHTML "  <div id=""corpo"">"
ShowHTML "    <div id=""menuesq"">"
ShowHTML "      <div id=""logomenuesq""><img src=""img/fundo_logomenuesq.gif"" border=0></div>"
ShowHTML "      <ul id=""menugov"">"
ShowHTML "      <script language=""JavaScript"" src=""inc/mm_menu.js"" type=""text/JavaScript""></script>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=110&CL=" & replace(CL,"sq_cliente=","") & """ >Inicial</a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=113&CL=" & replace(CL,"sq_cliente=","") & """ id=""link1"">Quem somos</a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=117&CL=" & replace(CL,"sq_cliente=","") & """ >Fale conosco </a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.Asp?EW=116&EF=" & CL & "&CL=" & replace(CL,"sq_cliente=","") & "&IN=1"" >Composiçao administrativa</a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & """ id=""link2"">Notícias</a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=115&CL=" & replace(CL,"sq_cliente=","") & """ id=""link5"">Calendário</a>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & """ >Arquivos (<i>download</i>)</a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=116&IN=0&CL=" & replace(CL,"sq_cliente=","") & """ >Escolas</a></li>"
ShowHTML "	    </ul>"
ShowHTML "      <div id=""menusep""><hr /></div>"
ShowHTML "      <div id=""menunav"">"
ShowHTML "      <ul id=""menunav"">"
ShowHTML "      <li><a href=""newsletter.asp?ew=i"">Clique aqui para receber informativos da SEDF</a>"
ShowHTML "	    </ul>"
ShowHTML "      </div>"
ShowHTML "    </div>"
ShowHTML "	  <div id=""menutxt"">"
ShowHTML "      <ul id=""menutexto"">"
ShowHTML "      <li><b>" & RS("ds_cliente") & "</b></li>"
Select Case sstrEW
  Case conWhatPrincipal   ShowHTML "      <li><b>Inicial</b></li>"
  Case conWhatManut       ShowHTML "      <li><b>Alunos</b></li>"
  Case conWhatQuem        ShowHTML "      <li><b>Quem somos</b></li>"
  Case conWhatExFale      ShowHTML "      <li><b>Fale conosco</b></li>"
  Case conWhatArquivo     ShowHTML "      <li><b>Arquivos</b></li>"
  Case conWhatExNoticia   ShowHTML "      <li><b>Notícias</b></li>"
  Case conWhatExCalend    ShowHTML "      <li><b>Calendário</b></li>"
  Case conWhatValida
     If sstrIN > 0 Then
        ShowHTML "      <li><b>Composição administrativa</b></li>"
     Else
        ShowHTML "      <li><b>Alunos</b></li>"
     End If
End Select

ShowHTML "	    </ul>"
ShowHTML "	  </div>"
ShowHTML "    <div id=""texto""><!-- Conteúdo -->"
ShowHTML "        <table width=""570"" border=""0"">"
RS.Close

Main

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

Set dbms          = nothing
Set RS1           = nothing
Set RS            = nothing
Set RSCabecalho   = nothing
Set RSRodape      = nothing

REM =========================================================================
REM Monta a tela de Arquivos
REM -------------------------------------------------------------------------
Public Sub ShowArquivo

  Dim sql, strNome, wCont, wAno
  
  wAno =  Request.QueryString("wAno")
  
  If wAno = "" Then
     wAno = Year(Date())
  End If

  sql = "SELECT case in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end AS in_destinatario, " & VbCrLf & _
        "       dt_arquivo, ds_titulo, nr_ordem, ds_arquivo, ln_arquivo, 'Regional' AS Origem, x.ln_internet diretorio " & VbCrLf & _
        "From escCliente_Arquivo a, escCliente x " & VbCrLf & _
        "WHERE in_ativo = 'Sim'" & VbCrLf & _
        "  AND x." & sstrEF & " " & VbCrLf & _
        "  AND a." & replace(sstrEF,"sq_cliente","sq_site_cliente") & " " & VbCrLf & _
        "  and in_destinatario <> 'E'" & VbCrLf & _
        "  and YEAR(a.dt_arquivo) = " & wAno & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT case a.in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end AS in_destinatario, " & VbCrLf & _
        "       a.dt_arquivo, a.ds_titulo, a.nr_ordem, ds_arquivo, ln_arquivo, 'SEDF' AS Origem, d.ds_diretorio diretorio " & VbCrLf & _
        "From escCliente_Arquivo AS a INNER JOIN escCliente AS c ON (a.sq_site_cliente = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente AS e ON (e.sq_cliente_pai  = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente_Site AS d ON (c.sq_cliente = d.sq_cliente) " & VbCrLf & _
        "WHERE a.in_ativo = 'Sim'" & VbCrLf & _
        "  AND e." & sstrEF & " " & VbCrLf & _
        "  and in_destinatario <> 'E'" & VbCrLf & _        
        "  and YEAR(a.dt_arquivo) = " & wAno & VbCrLf & _
        "ORDER BY origem, nr_ordem, dt_arquivo desc, in_destinatario " & VbCrLf 
  RS.Open sql, dbms, adOpenForwardOnly

  ShowHTML "<tr><td><TABLE border=0 cellSpacing=5 width=""95%"">"
  ShowHTML "  <TR>"
  ShowHTML "    <TD><b>Origem"
  ShowHTML "    <TD><b>Alvo"
  ShowHTML "    <TD><b>Data"
  ShowHTML "    <TD><b>Componente curricular"
  ShowHTML "  </TR>"
  ShowHTML "  <TR>"
  ShowHTML "    <TD COLSPAN=""4"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
  ShowHTML "  </TR>"
  wCont = 0
     
  If Not RS.EOF Then
     wCont = 1
     Do While Not RS.EOF
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD>" & RS("origem")
        ShowHTML "    <TD>" & RS("in_destinatario")
        ShowHTML "    <TD>" & Mid(100+Day(RS("dt_arquivo")),2,2) & "/" & Mid(100+Month(RS("dt_arquivo")),2,2) & "/" &Year(RS("dt_arquivo"))
        ShowHTML "    <TD><a href=""" & RS("diretorio") & "/" & RS("ln_arquivo") & """ target=""_blank"">" & RS("ds_titulo") & "</a><br><div align=""justify""><font size=1>.:. " & RS("ds_arquivo") & "</div>"
        ShowHTML "  </TR>"

        RS.MoveNext
     Loop
  Else
     ShowHTML "  <TR><TD COLSPAN=4 ALIGN=CENTER>Não há arquivos disponíveis no momento para o ano de " & wAno & " </TR>"
  End If
  RS.Close

  sql = "SELECT year(dt_arquivo) ano " & VbCrLf & _
        "From escCliente_Arquivo a, escCliente x " & VbCrLf & _
        "WHERE in_ativo = 'Sim'" & VbCrLf & _
        "  AND x." & sstrEF & " " & VbCrLf & _
        "  AND a." & replace(sstrEF,"sq_cliente","sq_site_cliente") & " " & VbCrLf & _
        "  and in_destinatario <> 'E'" & VbCrLf & _
        "  and YEAR(a.dt_arquivo) <> " & wAno & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT year(dt_arquivo) ano " & VbCrLf & _
        "From escCliente_Arquivo AS a INNER JOIN escCliente AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente AS c ON (c.sq_cliente_pai  = b.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente AS e ON (e.sq_cliente_pai  = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente_Site AS d ON (b.sq_cliente = d.sq_site_cliente) " & VbCrLf & _
        "WHERE a.in_ativo = 'Sim'" & VbCrLf & _
        "  AND e." & sstrEF & " " & VbCrLf & _
        "  and YEAR(a.dt_arquivo) <> " & wAno & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT year(dt_arquivo) ano " & VbCrLf & _
        "From escCliente_Arquivo AS a INNER JOIN escCliente AS c ON (a.sq_site_cliente = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente AS e ON (e.sq_cliente_pai  = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente_Site AS d ON (c.sq_cliente = d.sq_cliente) " & VbCrLf & _
        "WHERE a.in_ativo = 'Sim'" & VbCrLf & _
        "  AND e." & sstrEF & " " & VbCrLf & _
        "  and in_destinatario <> 'E'" & VbCrLf & _
        "  and YEAR(a.dt_arquivo) <> " & wAno & VbCrLf & _
        "ORDER BY year(dt_arquivo) desc " & VbCrLf 
    RS.Open sql, dbms, adOpenForwardOnly
    If Not RS.EOF Then
       ShowHTML "  <TR><TD COLSPAN=4 ><b>Arquivos de outros anos</b><br>"
       While Not RS.EOF
          ShowHTML "     <li><a href=""" & w_dir & "Default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & "&wAno=" & RS("ano") & """ >Arquivos de " & RS("ano") & "</a></TR>"
          RS.MoveNext
       Wend
       ShowHTML "  </TD></TR>"
    End If
    RS.Close
  ShowHTML "</table></center>"
End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Arquivos
REM =========================================================================

REM =========================================================================
REM Monta a tela principal
REM -------------------------------------------------------------------------
Public Sub ShowPrincipal

  Dim sql, strNome

  sql = "SELECT a.*, b.im_logo, b.im_foto_abertura1, b.im_foto_abertura2, b.ds_diretorio, b.ds_texto_abertura, d.*, e.sq_cliente_foto, e.ln_foto, e.ds_foto " & _
        "From escCliente as a INNER JOIN escCLIENTE_SITE  as b on (a.sq_cliente = b.sq_cliente) " & _
        "                     INNER JOIN escModelo        as c on (b.sq_modelo = c.sq_modelo) " & _
        "                     LEFT  JOIN escCliente_Dados as d on (a.sq_cliente = d.sq_cliente) " & _
        "                     LEFT  JOIN escCliente_Foto  as e on (a.sq_cliente = e.sq_cliente and e.tp_foto = 'P') " & _
        "WHERE a." & sstrEF & " " & VbCrLf & _
        "ORDER BY e.nr_ordem"

  RS.Open sql, dbms, adOpenForwardOnly

  ShowHTML "    <tr valign=""top""><td> "

  If Not RS.EOF Then
    If RS("ds_texto_abertura") > "" Then
       ShowHTML "        <P align=justify>" & replace(RS("ds_texto_abertura"),chr(13)&chr(10), "<br>")
    Else
       ShowHTML "        <P align=justify>"
    End If
    ShowHTML "        </td>"
    If RS("sq_cliente_foto") > "" Then
       ShowHTML "        <td align=""right"">"
       Do While Not RS.EOF
          ShowHTML "        <img class=""foto"" src=""" & RS("diretorio") & "/" & RS("ln_foto") & """ width=""302"" height=""206""><br>" & RS("ds_foto") & "<br>"
          RS.MoveNext
       Loop
    End If
  End If
  ShowHTML "        </td>"
  ShowHTML "    </tr>"

  RS.Close
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Página Principal
REM =========================================================================

REM =========================================================================
REM Monta a tela Quem Somos
REM -------------------------------------------------------------------------
Public Sub ShowQuem
  Dim sql, strNome

   sql = "SELECT b.ds_institucional, b.ds_diretorio, e.sq_cliente_foto, e.ln_foto, e.ds_foto " & _
        "From escCliente as a INNER JOIN escCLIENTE_SITE  as b on (a.sq_cliente = b.sq_cliente) " & _
        "                     INNER JOIN escModelo        as c on (b.sq_modelo = c.sq_modelo) " & _
        "                     LEFT JOIN escCliente_Dados as d on (a.sq_cliente = d.sq_cliente) " & _
        "                     LEFT  JOIN escCliente_Foto  as e on (a.sq_cliente = e.sq_cliente and e.tp_foto = 'Q') " & _
        "WHERE a." & sstrEF & " " & VbCrLf & _
        "ORDER BY e.nr_ordem"
  RS.Open sql, dbms, adOpenForwardOnly

  ShowHTML "    <tr valign=""top""><td> "

  If Not RS.EOF Then
    If RS("ds_institucional") > "" Then
       ShowHTML "        <p align=""left"" style=""margin-top: 0; margin-bottom: 0"">" & replace(RS("ds_institucional"),chr(13)&chr(10), "<br>")
    Else
       ShowHTML "        <P align=justify>"
    End If
    ShowHTML "        </td>"
    If RS("sq_cliente_foto") > "" Then
       ShowHTML "  <td align=""right""><font size=""1""><b>Fotografias</b><font size=1> (Clique sobre a imagem para ampliar)</b><br>"
       Do While Not RS.EOF
          ShowHTML "   <a href=""" & RS("diretorio") & "/" & RS("ln_foto") & """ target=""_blank""><img align=""top"" class=""foto"" src=""" & RS("diretorio") & "/" & RS("ln_foto") & """ width=""201"" height=""153""> " & RS("ds_foto") & "</a>"
          RS.MoveNext
       Loop
    End If
  End If
  ShowHTML "        </td>"
  ShowHTML "    </tr>"    
  RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Página Quem Somos
REM =========================================================================

REM =========================================================================
REM Monta a tela Fale Conosco
REM -------------------------------------------------------------------------
Public Sub ShowExFale

  Dim sql, strNome

  sql = "SELECT b.*, c.*, a.no_municipio, a.sg_uf " & _
        "From escCliente as a INNER JOIN escCLIENTE_SITE as b on (a.sq_cliente = b.sq_cliente) " & _
        "                     LEFT JOIN escCLIENTE_Dados as c on (a.sq_cliente = c.sq_cliente) " & _
        "WHERE a." & sstrEF 

  RS.Open sql, dbms, adOpenForwardOnly
  
  'Validação dos campos do formulário de envio de email
  'ScriptOpen "JavaScript"
  'ValidateOpen "Validacao"
  'Validate "w_nome", "Nome", "1", "1", "3", "60", "1", "1"
  'Validate "w_email", "e-Mail", "1", "1", "3", "60", "1", "1"
  'Validate "w_endereco", "Endereço", "1", "1", "3", "60" , "1", "1"
  'Validate "w_telefone", "Telefone", "1", "1", "7", "30" , "1", "1"
  'Validate "w_assunto", "Assunto", "1", "1", "5", "80" , "1", "1"
  'Validate "w_mensagem", "Mensagem", "1", "1", "5", "1000" , "1", "1"
  'ShowHTML "  theForm.Botao.disabled=true;"
  'ValidateClose
  'ScriptClose
  
  
  ShowHTML "<tr><td>"

  ShowHTML "<p align=""justify""><li>Informações, sugestões e reclamações podem ser feitas utilizando os dados abaixo:"

  If Not RS.EOF Then

  ShowHTML "<table border=""0"" cellspacing=""3"">"
  ShowHTML "  <tr>"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td width=""20%"" align=""right""><b>Diretor(a):"
  ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_diretor") 
  ShowHTML "  </tr>"
  'ShowHTML "  <tr>"
  'ShowHTML "    <td width=""10%"">"
  'ShowHTML "    <td width=""20%"" align=""right""><b>Secretário(a):"
  'ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_secretario") 
  'ShowHTML "  </tr>"
  ShowHTML "  <tr>"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td width=""20%"" align=""right""><b>Contato:</b>"
  ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_contato_internet")
  if RS("ds_email_internet") > "" then
     ShowHTML "  </tr>"
     ShowHTML "  <tr>"
     ShowHTML "    <td width=""10%"">"
     ShowHTML "    <td width=""20%"" align=""right"">"
     ShowHTML "    <td width=""70%"" align=""left"">(<a href=""mailto:" & RS("ds_email_internet") & """>" & RS("ds_email_internet") & "</a>)"
  end if
  ShowHTML "  </tr>"
  ShowHTML "  <tr>"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td width=""20%"" align=""right""><b>Telefone:"
  ShowHTML "    <td width=""70%"" align=""left"">" & RS("nr_fone_internet")
  if RS("nr_fax_internet") > "" then
     ShowHTML "    <b> Fax: </b>" & RS("nr_fax_internet")
  end if
  ShowHTML "  </tr>"
  ShowHTML "  <tr>"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td width=""20%"" align=""right""><b>Endereço:"
  ShowHTML "    <td width=""70%"" align=""left"">" & RS("ds_logradouro")
  ShowHTML "  </tr>"
  If RS("no_bairro") > "" Then
     ShowHTML "  <tr>"
     ShowHTML "    <td width=""10%"">"
     ShowHTML "    <td width=""20%"" align=""right""><b>Bairro:"
     ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_bairro")
     ShowHTML "  </tr>"
  End If
  ShowHTML "  <tr>"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td width=""20%"" align=""right""><b>Município:"
  ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_municipio") & "-" & RS("sg_uf") & "&nbsp;&nbsp;<b>CEP:</b> " & RS("nr_cep")
  ShowHTML "  </tr>"
  ShowHTML "</table>"
  
  'If Nvl(RS("ds_email_internet"),"") > "" Then
  '   ShowHTML "<FORM action=""" & w_dir & sstrSN & "?EW=PreparaMail"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"">"
  '   ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
  '   ShowHTML "<p align=""justify""><li>Se desejar, envie uma mensagem de e-mail usando o formulário abaixo:"
  '   ShowHTML "<table border=""1"" bgcolor=""#DAEABD"" style=""border: 1px solid rgb(0,0,0);""><tr><td>"
  '   ShowHTML "  <table border=""0"" cellspacing=""3"">"
  '   ShowHTML "    <tr><td><font size=""1""><b>Nome: </b><br><input type=""text"" name=""w_nome"" class=""STI"" size=""60"" maxlength=""60"" value="""" title=""Informe o nome do remetente"">"
  '   ShowHTML "    <tr><td><font size=""1""><b>e-Mail: </b><br><input type=""text"" name=""w_email"" class=""STI"" size=""60"" maxlength=""60"" value="""" title=""Informe o email do remetente"">"
  '   ShowHTML "    <tr><td><font size=""1""><b>Endereço: </b><br><input type=""text"" name=""w_endereco"" class=""STI"" size=""60"" maxlength=""60"" value="""" title=""Informe o endereco do remetente"">"
  '   ShowHTML "    <tr><td><font size=""1""><b>Telefone: </b><br><input type=""text"" name=""w_telefone"" class=""STI"" size=""13"" maxlength=""30"" value="""" title=""Informe o telefone do remetente"">"
  '   ShowHTML "    <tr><td><font size=""1""><b>Assunto: </b><br><input type=""text"" name=""w_assunto"" class=""STI"" size=""66"" maxlength=""80"" value="""" title=""Informe o assunto do E-mail"">"
  '   ShowHTML "    <tr><td><font size=""1""><b>Mensagem: </b><br><TEXTAREA class=""STI"" type=""text"" name=""w_mensagem"" ROWS=4 COLS=50></TEXTAREA ></td>"
  '   ShowHTML "    <tr><td align=""center""><font size=""1""><input class=""STB"" type=""submit"" name=""Botao"" value=""Enviar"">"
  '   ShowHTML "  </table>"
  '   ShowHTML "</table>"
  '   ShowHTML "</FORM>"  
  'End If
     
  ShowHTML "</tr>"
  RS.Close
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Página Fale Conosco
REM =========================================================================

REM =========================================================================
REM Monta a tela Notícias
REM -------------------------------------------------------------------------
Public Sub ShowNoticia

  Dim sql, strNome, wCont, wRede, wAno
  
  wAno =  Request.QueryString("wAno")
  
  If wAno = "" Then
     wAno = Year(Date())
  End If

  If sstrIN = "" then
     sql = "SELECT a.sq_noticia, a.dt_noticia AS data, a.ds_titulo AS ocorrencia, 'G' AS origem " & VbCrLf & _
           "  FROM escNoticia_Cliente  a" & VbCrLf & _
           " WHERE sq_site_cliente   = " & CL & VbCrLf & _
           "  and a.in_ativo         = 'Sim'" & VbCrLf & _
           "  and a.in_exibe         = 'Sim'" & VbCrLf & _
           "  and YEAR(a.dt_noticia) = " & wAno & VbCrLf & _
           "UNION " & VbCrLf & _
           "SELECT a.sq_noticia, dt_noticia AS data, ds_titulo AS ocorrencia, 'R' AS origem " & VbCrLf & _
           "FROM escNoticia_Cliente    AS a" & VbCrLf & _
           "     INNER JOIN escCliente AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
           "     INNER JOIN escCliente AS c ON (b.sq_cliente      = c.sq_cliente_pai) " & VbCrLf & _
           "WHERE c.sq_cliente       = " & CL & VbCrLf & _
           "  and a.in_ativo         = 'Sim'" & VbCrLf & _
           "  and a.in_exibe         = 'Sim'" & VbCrLf & _
           "  and YEAR(a.dt_noticia) = " & wAno & VbCrLf & _
           "ORDER BY data desc, ocorrencia" & VbCrLf
     RS.Open sql, dbms, adOpenForwardOnly
     
     ShowHTML "<tr><td><TABLE align=""center"" width=""95%""border=0 cellSpacing=1>"
     ShowHTML "  <TR valign=""top"" align=""center"">"
     ShowHTML "    <TD width=""2%"">&nbsp;"
     ShowHTML "    <TD width=""15%""><b>Data"
     ShowHTML "    <TD width=""10%""><b>Origem"
     ShowHTML "    <TD width=""15%""><b>Ocorrência"
     ShowHTML "    <TD width=""48%"">"
     ShowHTML "  </TR>"
     ShowHTML "  <TR>"
     ShowHTML "    <TD COLSPAN=""6"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
     ShowHTML "  </TR>"
     wCont = 0
     If Not RS.EOF Then
        wCont = 1
        Do While Not RS.EOF
           ShowHTML "  <TR valign=""top"">"
           ShowHTML "    <TD>&nbsp;"
           If RS("origem") = "R" then
              ShowHTML "    <TD align=""center"">" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2) & "/" &Year(RS("data"))
              ShowHTML "    <TD align=""center"">SEDF"
              ShowHTML "    <TD colspan=2>" & RS("ocorrencia") & "&nbsp;&nbsp;&nbsp;<a href=""" & w_dir & sstrSN & "?EW=" & conWhatExNoticia & "&CL=" & CL & "&EF=" & sstrEF & "&IN=" & RS("sq_noticia") & """><b>[Ler]</a>"
           Else
              ShowHTML "    <TD align=""center"">" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2) & "/" &Year(RS("data"))
              ShowHTML "    <TD align=""center"">Regional"
              ShowHTML "    <TD colspan=2>" & RS("ocorrencia") & "&nbsp;&nbsp;&nbsp;<a href=""" & w_dir & sstrSN & "?EW=" & conWhatExNoticia & "&CL=" & CL & "&EF=" & sstrEF & "&IN=" & RS("sq_noticia") & """><b>[Ler]</a>"
           End If
           ShowHTML "  </TR>"
           wCont = wCont + 1
           RS.MoveNext
        Loop
        
     Else
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD colspan=""5""><BR>Nenhuma notícia foi encontrada para o ano " & wAno & "</TD>"
        ShowHTML "  </TR>"
     End If
     RS.Close
     sql = "SELECT YEAR(a.dt_noticia) ano " & VbCrLf & _
           "  FROM escNoticia_Cliente  a" & VbCrLf & _
           " WHERE sq_site_cliente   = " & CL & VbCrLf & _
           "  and a.in_ativo         = 'Sim'" & VbCrLf & _
           "  and a.in_exibe         = 'Sim'" & VbCrLf & _
           "  and YEAR(a.dt_noticia) <> " & wAno & VbCrLf & _
           "UNION " & VbCrLf & _
           "SELECT YEAR(a.dt_noticia) ano " & VbCrLf & _
           "FROM escNoticia_Cliente    AS a" & VbCrLf & _
           "     INNER JOIN escCliente AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
           "     INNER JOIN escCliente AS c ON (b.sq_cliente      = c.sq_cliente_pai) " & VbCrLf & _
           "WHERE c.sq_cliente       = " & CL & VbCrLf & _
           "  and a.in_ativo         = 'Sim'" & VbCrLf & _
           "  and a.in_exibe         = 'Sim'" & VbCrLf & _
           "  and YEAR(a.dt_noticia) <> " & wAno & VbCrLf & _
           "ORDER BY YEAR(dt_noticia) desc" & VbCrLf
     RS.Open sql, dbms, adOpenForwardOnly
     If Not RS.EOF Then
        ShowHTML "  <TR valign=""top""><TD colspan=""5""><br><b>Notícias de outros anos</b><br>"     
        While Not RS.EOF
           ShowHTML "    <li><a href=""" & w_dir & "Default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & "&wAno=" & RS("ano") & """ id=""link2"">Notícias de " & RS("ano") & "</a> </TD>"
           RS.MoveNext
        Wend
        ShowHTML "  </TD></TR>"
     End If
     RS.Close
     ShowHTML "    </TABLE>"
 
  Else
     sql = "SELECT * " & _
           "From escNoticia_Cliente " & _
           "WHERE sq_noticia = " & sstrIN & " "
     RS.Open sql, dbms, adOpenForwardOnly

     ShowHTML "<tr><td><p align=""center""><font size=""4""><b>" & RS("ds_titulo") & "</b></font><p align=""left"">"
     ShowHTML replace(RS("ds_noticia"),chr(13)&chr(10), "<br>")
     ShowHTML "<p><center><img src=""Img/bt_voltar.gif"" border=0 valign=""center"" onClick=""history.go(-1)"" alt=""Volta"">"
     RS.Close
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Página Notícias
REM =========================================================================

REM =========================================================================
REM Monta a tela Calendário
REM -------------------------------------------------------------------------
Public Sub ShowCalend

  Dim sql, strNome, wCont, wRede, wAno, wDatas(31,12,10), wCores(31,12,10)
  wAno = Year(Date())
  
  sql = "SELECT '' as cor, dt_ocorrencia AS data, ds_ocorrencia AS ocorrencia, 'B' AS origem FROM escCalendario_base WHERE YEAR(dt_ocorrencia)=" & Year(Date()) & " " & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT '#99CCFF' as cor, dt_ocorrencia AS data, ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente WHERE sq_site_cliente = " & CL & "  AND YEAR(dt_ocorrencia)= " & Year(Date()) & " " & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT '#FFFF99' as cor, dt_ocorrencia AS data, ds_ocorrencia AS ocorrencia, 'R' AS origem " & VbCrLf & _
        "FROM escCalendario_Cliente AS a" & VbCrLf & _
        "     INNER JOIN escCliente AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
        "     INNER JOIN escCliente AS c ON (b.sq_cliente = c.sq_cliente_pai) " & VbCrLf & _
        "WHERE c.sq_cliente = " & CL & VbCrLf & _
        "  AND YEAR(dt_ocorrencia)= " & wAno & " " & VbCrLf & _
        "ORDER BY data, origem desc, ocorrencia" & VbCrLf 

  RS.Open sql, dbms, adOpenForwardOnly
  
  If sstrIN = "" then
     If Not RS.EOF Then
        Do While Not RS.EOF
           If RS("origem") = "E" then
              wDatas(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("ocorrencia") & " (Origem: Escola)"
              wCores(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("cor")
           ElseIf RS("origem") = "R" then
              wDatas(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("ocorrencia") & " (Origem: SEDF)"
              wCores(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("cor")
           Else
              wDatas(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("ocorrencia") & " (Origem: Oficial)"
           End If
           RS.MoveNext
        Loop
        RS.MoveFirst
     End If
     ShowHTML "<tr><td><TABLE align=""center"" width=""100%"" border=0 cellSpacing=1>"
     ShowHTML "<tr valign=""top"">"
     ShowHTML "  <td>" & MontaCalendario("01" & wAno, wDatas, wCores)
     ShowHTML "  <td>" & MontaCalendario("02" & wAno, wDatas, wCores)
     ShowHTML "  <td>" & MontaCalendario("03" & wAno, wDatas, wCores)
     ShowHTML "  <td>" & MontaCalendario("04" & wAno, wDatas, wCores)
     ShowHTML "<tr valign=""top"">"
     ShowHTML "  <td>" & MontaCalendario("05" & wAno, wDatas, wCores)
     ShowHTML "  <td>" & MontaCalendario("06" & wAno, wDatas, wCores)
     ShowHTML "  <td>" & MontaCalendario("07" & wAno, wDatas, wCores)
     ShowHTML "  <td>" & MontaCalendario("08" & wAno, wDatas, wCores)
     ShowHTML "<tr valign=""top"">"
     ShowHTML "  <td>" & MontaCalendario("09" & wAno, wDatas, wCores)
     ShowHTML "  <td>" & MontaCalendario("10" & wAno, wDatas, wCores)
     ShowHTML "  <td>" & MontaCalendario("11" & wAno, wDatas, wCores)
     ShowHTML "  <td>" & MontaCalendario("12" & wAno, wDatas, wCores)
     ShowHTML "</table>"

     ShowHTML "<tr><td><TABLE align=""center"" width=""95%""border=0 cellSpacing=1>"
     ShowHTML "  <TR valign=""top"">"
     ShowHTML "    <TD width=""2%"">&nbsp;"
     ShowHTML "    <TD width=""15%""><b>Origem"
     ShowHTML "    <TD width=""15%""><b>Data"
     ShowHTML "    <TD width=""15%""><b>Ocorrência"
     ShowHTML "    <TD width=""43%"">"
     ShowHTML "  </TR>"
     ShowHTML "  <TR>"
     ShowHTML "    <TD COLSPAN=""6"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
     ShowHTML "  </TR>"
     wCont = 0
     If Not RS.EOF Then
        wCont = 1
        Do While Not RS.EOF

           ShowHTML "  <TR>"
           ShowHTML "    <TD>&nbsp;"
           If RS("origem") = "E" then
              ShowHTML "    <TD>Escola"
              ShowHTML "    <TD>" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2) & "/" &Year(RS("data"))
              ShowHTML "    <TD colspan=2><font color=""#0000FF"">" & RS("ocorrencia")
              wDatas(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("ocorrencia") & " (Origem: Escola)"
           ElseIf RS("origem") = "R" then
              ShowHTML "    <TD>SEDF"
              ShowHTML "    <TD>" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2) & "/" &Year(RS("data"))
              ShowHTML "    <TD colspan=2>" & RS("ocorrencia")
              wDatas(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("ocorrencia") & " (Origem: SEDF)"
           Else
              ShowHTML "    <TD>Oficial"
              ShowHTML "    <TD>" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2) & "/" &Year(RS("data"))
              ShowHTML "    <TD colspan=2>" & RS("ocorrencia")
              wDatas(Day(RS("data")), Month(RS("data")), Mid(Year(RS("data")),4,2)) = RS("ocorrencia") & " (Origem: Oficial)"
           End If
           ShowHTML "  </TR>"
           wCont = wCont + 1
           RS.MoveNext
        Loop

     End If

     ShowHTML "    </TABLE>"
 
  End If
  RS.Close
End Sub
REM -------------------------------------------------------------------------
REM Final da Página Calendário
REM =========================================================================

REM =========================================================================
REM Monta a tela de escolas
REM -------------------------------------------------------------------------
Public Sub ShowEscolas

  Dim sql2, wCont, sql1, wAtual, wIN
  
  If sstrIN > 0 Then
     sql = "SELECT b.ds_diretorio, b.ln_prop_pedagogica " & _
           "From escCliente as a INNER JOIN escCLIENTE_SITE as b on (a.sq_cliente = b.sq_cliente) " & _
           "                     INNER JOIN escModelo as c on (b.sq_modelo = c.sq_modelo) " & _
           "                     INNER JOIN escCliente_Dados as d on (a.sq_cliente = d.sq_cliente) " & _
           "WHERE a." & sstrEF 
     ConectaBD SQL

     ShowHTML "<tr><td>"
     If Not RS.EOF Then
        ShowHTML "    <font size=1><b>"
        If InStr(lCase(RS("ln_prop_pedagogica")),lCase(".pdf")) > 0 Then
           ShowHTML "        A exibição do arquivo exige que o Acrobat Reader tenha sido instalado em seu computador."
           ShowHTML "        <br>Se o arquivo não for exibido no quadro abaixo, clique <a href=""http://www.adobe.com.br/products/acrobat/readstep2.html"" target=""_blank"">aqui</a> para instalar ou atualizar o Acrobat Reader."
        Else
           ShowHTML "        A exibição do arquivo exige o editor de textos Word ou equivalente. "
           ShowHTML "        <br>Se o arquivo não for exibido no quadro abaixo, verifique se o Word foi corretamente instalado em seu computador."        
        End If
        ShowHTML "<table align=""center"" width=""100%"" cellspacing=0 style=""border: 1px solid rgb(0,0,0);""><tr><td>"
        ShowHTML "    <iframe src=""" & RS("diretorio") & "/" & RS("ln_prop_pedagogica") & """ width=""100%"" height=""510"">"
        ShowHTML "    </iframe>"
        ShowHTML "</table>"
     Else
        ShowHTML "    <font size=1><b>Composição administrativa não informada."
     End If
  Else
    ShowHTML "        <FORM ACTION=""" & w_dir & "default.asp"" id=form1 name=form1 METHOD=""POST"">"
    ShowHTML "        <input type=""Hidden"" name=""EW"" value=""" & sstrEW & """>"
    ShowHTML "        <input type=""Hidden"" name=""IN"" value=""" & sstrIN & """>"
    ShowHTML "        <input type=""Hidden"" name=""CL"" value=""" & CL & """>"
    ShowHTML "        <input type=""Hidden"" name=""htBT"" value=""1"">"
    ShowHTML "        <input type=""Hidden"" name=""str2"" value=""10"">"
    ShowHTML "          <TR><TD colspan=""2"" align=""left"" valign""middle""><p align=""justify"">Selecione as formas de busca desejadas para listar as escolas:</p></td></tr>"
    SQL = "SELECT sq_tipo_cliente, ds_tipo_cliente FROM escTipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
    ConectaBD SQL
    ShowHTML "          <tr><td colspan=2><b>Tipo de escola:</b><br><SELECT class=""texto"" NAME=""Q"">"
    If RS.RecordCount > 1 Then ShowHTML "          <option value="""">Todos" End If
    While Not RS.EOF
       If cDbl(nvl(RS("sq_tipo_cliente"),0)) = cDbl(nvl(Request("Q"),0)) Then
          ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """ SELECTED>" & RS("ds_tipo_cliente")
       Else
          ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """>" & RS("ds_tipo_cliente")
       End If
       RS.MoveNext
    Wend
    ShowHTML "          </select>"
    DesconectaBD
    wCont = 0
    wIN = 0
    sql1 = ""

    sql = "SELECT DISTINCT a.* " & _ 
          "from escEspecialidade AS a " & _ 
          "     INNER JOIN escEspecialidade_cliente AS c ON (a.sq_especialidade = c.sq_codigo_espec) " & _
          "     INNER JOIN escCliente AS d ON (c.sq_cliente = d.sq_cliente) " & _
          "ORDER BY a.nr_ordem, a.ds_especialidade "
    ConectaBD SQL
  
    If Not RS.EOF Then
	    wCont = 0
	    wAtual = ""

	    Do While Not RS.EOF
             
         If wAtual = "" or wAtual <> RS("tp_especialidade") Then
            wAtual = RS("tp_especialidade")
            If wAtual = "M" Then
               ShowHTML "          <TR><TD colspan=2><b>Etapas / Modalidades de ensino</b>:"
            ElseIf wAtual = "R" Then
               ShowHTML "          <TR><TD colspan=2><b>Em Regime de Intercomplementaridade</b>:"
            End If
         End If
         wCont = wCont + 1
         marcado = "N"
         For i = 1 to Request("H").Count
             If cDbl(RS("sq_especialidade")) = cDbl(Request("H")(i)) Then marcado = "S" End If
         Next
         
         If marcado = "S" Then
            ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""H"" value=""" & RS("sq_especialidade") & """ checked> " & RS("ds_especialidade")
            sql1 = Request("H")
            wIN = 1
         Else
            ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""H"" value=""" & RS("sq_especialidade") & """> " & RS("ds_especialidade")
         End If
	       RS.MoveNext

         If (wCont Mod 2) = 0 Then 
           wCont = 0
           'ShowHTML "          <TR>"
	       End If
	    Loop
    End If
    ShowHTML "          <TR><TD colspan=2 valign""middle"">"
    ShowHTML "           <input type=""button"" name=""Botao"" value=""Pesquisar"" class=""botao"" onClick=""javascript: document.form1.Botao.disabled=true; document.form1.submit();"">"
    ShowHTML "        </tr>"
    ShowHTML "        </form>"

    if Request("Q") > "" or wIN > 0 or CL > 0 then
      ShowHTML "        <tr><td colspan=2><table width=""570"" border=""0"" align=""left"" cellpadding=0 cellspacing=0>"
      ShowHTML "          <TR><TD><hr>"
      ShowHTML "          <TR><TD><b><a name=""resultado"">Resultado da pesquisa:</a><br><br>"

      sql = "SELECT DISTINCT 0 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf,e.ds_logradouro,e.no_bairro,e.nr_cep " & VbCrLf & _
            "  from escCliente                               d " & VbCrLf & _
            "       INNER      JOIN escTipo_Cliente          b ON (d.sq_tipo_cliente  = b.sq_tipo_cliente and " & VbCrLf & _
            "                                                      b.tipo             = 2 " & VbCrLf & _
            "                                                     ) " & VbCrLf & _
            "       LEFT OUTER JOIN escCliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
            " where d.sq_cliente = " & CL & " " & VbCrLf & _
            "UNION " & VbCrLf & _
            "SELECT DISTINCT 0 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf,e.ds_logradouro,e.no_bairro,e.nr_cep " & VbCrLf & _
            "  from escCliente                               d " & VbCrLf & _
            "       LEFT OUTER JOIN escCliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
            "       INNER      JOIN escCliente_Site          a ON (d.sq_cliente       = a.sq_cliente) " & VbCrLf & _
            "       INNER      JOIN escTipo_Cliente          b ON (d.sq_tipo_cliente  = b.sq_tipo_cliente and " & VbCrLf & _
            "                                                      b.tipo             = 2 " & VbCrLf & _
            "                                                     ) " & VbCrLf & _
            "       INNER      JOIN escCliente               c ON (d.sq_cliente       = c.sq_cliente_pai) " & VbCrLf & _
            "         INNER    JOIN escEspecialidade_cliente f ON (c.sq_cliente       = f.sq_cliente), " & VbCrLf & _
            "       escCliente                               g " & VbCrLf & _
            " where g.sq_cliente = " & CL & " " & VbCrLf
      If CL > ""                    Then sql = sql + "    and c.sq_cliente_pai = " & CL & VbCrLf End If
      If Request("q") > ""          Then sql = sql + "    and c.sq_tipo_cliente= " & Request("q")          & VbCrLf End If
      if sql1 > "" then
         sql = sql + "  and f.sq_codigo_espec in (" + Request("H") + ")" & VbCrLf 
      end if
      sql = sql & _      
            "UNION " & VbCrLf & _ 
            "SELECT DISTINCT 1 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf,e.ds_logradouro,e.no_bairro,e.nr_cep " & VbCrLf & _ 
            "  from escEspecialidade                           a " & VbCrLf & _ 
            "       INNER        JOIN escEspecialidade_cliente c ON (a.sq_especialidade = c.sq_codigo_espec) " & VbCrLf & _
            "       INNER        JOIN escCliente               d ON (c.sq_cliente       = d.sq_cliente) " & VbCrLf & _
            "         LEFT OUTER JOIN escCliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
            " where 1=1 " & VbCrLf 

      If CL > ""                    Then sql = sql + "    and d.sq_cliente_pai = " & CL & VbCrLf End If
      If Request("q") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("q")          & VbCrLf End If

      if sql1 > "" then
         sql = sql + "  and a.sq_especialidade in (" + Request("H") + ")" & VbCrLf 
      end if
      sql = sql + "ORDER BY tipo, d.ds_cliente " & VbCrLf
      ConectaBD SQL
      
      If Not RS.EOF Then

        If Request("str2") > "" Then RS.PageSize = cDbl(Request("str2")) Else RS.PageSize = 10 End If
        
        rs.AbsolutePage = Nvl(Request("htBT"),1)

        ShowHTML "          <TR><TD><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
        ShowHTML "            <tr><td width=""5%""><td width=""95%"">"

        While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(Nvl(Request("htBT"),1))

          ShowHTML "            <TR><TD valign=""top""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center""> "
          ShowHTML "                <td align=""left""><font size=""2""><b>"
          If Not IsNull (RS("LN_INTERNET")) Then
             ShowHTML "                <a href=""http://" & replace(RS("LN_INTERNET"),"http://","") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b>"
          Else
             ShowHTML RS("DS_CLIENTE") & "</b>"
          End If
          ShowHTML "<br>Endereço: " & RS("ds_logradouro")
          If Nvl(RS("no_bairro"),"nulo") <> "nulo" Then ShowHTML " - Bairro: " & RS("no_bairro") End If
          If Nvl(RS("nr_cep"),"nulo") <> "nulo" and len(RS("nr_cep")) > 5 Then ShowHTML " - CEP: " & RS("nr_cep") End If
          ShowHTML "<br>Localização: " & RS("no_municipio") & "-" & RS("sg_uf")

	        If RS("tipo") <> "0" Then
	           sql2 = "SELECT f.ds_especialidade " & _ 
	                 "from escEspecialidade_cliente AS e " & _
                   "     INNER JOIN escEspecialidade AS f ON (e.sq_codigo_espec = f.sq_especialidade) " & _
                   "WHERE e.sq_cliente = " & RS("sq_cliente") & " " & _
                   "ORDER BY f.ds_especialidade "

             RS1.Open sql2, dbms, adOpenForwardOnly

             wCont = 0
	           While Not RS1.EOF
	             If wCont = 0 Then
                 ShowHTML "<br>" & RS1("ds_especialidade")
                 wCont = 1
               Else
                 ShowHTML ", " & RS1("ds_especialidade")
               End if
	             RS1.MoveNext
	           Wend

             RS1.Close
          End If

    	    RS.MoveNext

    	  Wend
	
        ShowHTML "            </table>"
        ShowHTML "            <tr><td><td align=""center""><hr>"

        ShowHTML "<tr><td>"
        MontaBarra w_dir&"default.asp?EW="&sstrEW&"&IN="&sstrIN&"&CL="&CL, cDbl(RS.PageCount), cDbl(Nvl(Request("htBT"),1)), cDbl(Request("str2")), cDbl(RS.RecordCount)
        ShowHTML "</tr>"


      Else

        ShowHTML "            <TR><TD colspan=2><p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<b>Nenhuma ocorrência encontrada para as opções acima."

      End If

      ShowHTML "          </table>"
    End If
  End If
End Sub

REM =========================================================================
REM Rotina de preparação para envio de e-mail do fale conosco
REM -------------------------------------------------------------------------
Sub PreparaMail

  Dim w_fc_nome, w_fc_email, w_fc_endereco, w_fc_telefone, w_fc_assunto
  Dim w_fc_mensagem, w_cliente
  Dim RSM, w_destinatario, w_assunto, w_resultado, w_html
  
  Set RSM = Server.CreateObject("ADODB.RecordSet")
  
  w_fc_nome     = Request("w_nome")
  w_fc_email    = Request("w_email")
  w_fc_endereco = Request("w_endereco")
  w_fc_telefone = Request("w_telefone")
  w_fc_assunto  = Request("w_assunto")
  w_fc_mensagem = Request("w_mensagem")   
    
  SQL = "SELECT a.ds_cliente, b.ds_email_internet, c.no_diretor, c.no_secretario " & VbCrLf &_
        "  FROM escCliente                       a " & VbCrLf &_
        "       LEFT OUTER JOIN escCliente_Site  b on (a.sq_cliente = b.sq_cliente) " & VbCrLf &_
        "       LEFT OUTER JOIN escCliente_Dados c on (a.sq_cliente = c.sq_cliente) " & VbCrLf &_
        " WHERE a.sq_cliente = " & CL & VbCrLf
  RSM.Open SQL, dbms, adOpenForwardOnly
  
  'w_destinatario = RS("ds_email_internet")&";"&w_fc_email
  w_destinatario = "celso@sbpi.com.br"&";"&w_fc_email
  w_resultado = ""

  w_html = "<HTML>" & VbCrLf
  w_html = w_html & BodyOpenMail(null) & VbCrLf
  w_html = w_html & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & VbCrLf
  w_html = w_html & "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">" & VbCrLf
  w_html = w_html & "    <table width=""97%"" border=""0"">" & VbCrLf
  w_html = w_html & "      <tr><td align=""center"" bgcolor=""#D0D0D0"" style=""border: 2px solid rgb(0,0,0);""><font size=""1""><b>FALE CONOSCO (" & RSM("ds_cliente") & ")</b></td>"
  w_html = w_html & "      <tr><td><font size=2><b><font color=""#BC3131"">ATENÇÃO</font>: Esta mensagem foi enviada pelo site da regional de ensino. Favor não responder.</b></font><br><br><td></tr>" & VbCrLf
  
  'Cabeçalho do e-Mail
  w_html = w_html & VbCrLf & "      <tr><td align=""left""><font size=""1"">&nbsp;</td>"
  w_html = w_html & VbCrLf & "      <tr bgcolor=""" & conTrBgColor & """><td align=""center""><table width=""99%"" border=""0"">"
  w_html = w_html & VbCrLf & "        <tr><td align=""center"" bgcolor=""#D0D0D0"" style=""border: 2px solid rgb(0,0,0);""><font size=""1""><b>REMETENTE</b></td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left"" bgcolor=""#FDFDFD"" style=""border: 1px solid rgb(0,0,0);""><font size=""1"">Nome: <b>" & w_fc_nome & "</td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left"" bgcolor=""#FDFDFD"" style=""border: 1px solid rgb(0,0,0);""><font size=""1"">e-Mail: <b><a href=""mailto:" & w_fc_email & """>" & w_fc_email & "</td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left"" bgcolor=""#FDFDFD"" style=""border: 1px solid rgb(0,0,0);""><font size=""1"">Endereço: <b>" & w_fc_endereco & "</td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left"" bgcolor=""#FDFDFD"" style=""border: 1px solid rgb(0,0,0);""><font size=""1"">Telefone: <b>" & w_fc_telefone & "</td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left""><font size=""1"">&nbsp;</td>"

  ' Mensagem do e-Mail
  w_html = w_html & VbCrLf & "        <tr><td align=""center"" bgcolor=""#D0D0D0"" style=""border: 2px solid rgb(0,0,0);""><font size=""1""><b>MENSAGEM</b></td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left"" bgcolor=""#FDFDFD"" style=""border: 1px solid rgb(0,0,0);""><font size=1>Assunto: <b>" & w_fc_assunto & "</b></font></td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left"" bgcolor=""#FDFDFD"" style=""border: 1px solid rgb(0,0,0);""><font size=1>Texto:<br><b>" & CRLF2BR(w_fc_mensagem) & "</td>"
  w_html = w_html & VbCrLf & "      </table>"
  w_html = w_html & VbCrLf & "</tr>"
  
  'Dados finais do email
  w_html = w_html & VbCrLf & "        <tr><td align=""left""><font size=""1"">&nbsp;</td>"
  w_html = w_html & VbCrLf & "      <tr><td align=""center"" bgcolor=""#D0D0D0"" style=""border: 2px solid rgb(0,0,0);""><font size=""1""><b>DADOS DA OCORRÊNCIA</b></td>"
  w_html = w_html & "      <tr valign=""top""><td><font size=2>" & VbCrLf
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
  w_assunto = "Encaminhamento de mensagem pelo site"
  
  RSM.Close
  
  
  If w_destinatario > "" Then
     ' Executa o envio do e-mail
     w_resultado = EnviaMail1(w_assunto, w_html, w_destinatario)
  End If
  
  ' Se ocorreu algum erro, avisa da impossibilidade de envio, senão volta para pagina do fale conosco
  If w_resultado > "" Then 
     ScriptOpen "JavaScript"
     ShowHTML "  alert('ATENÇÃO: não foi possível proceder o envio do e-mail.\n" & w_resultado & "');" 
     ShowHTML "  history.back(1);"
     ScriptClose
  Else
     ScriptOpen "JavaScript"
     ShowHTML "  alert('Mensagem enviada com sucesso, uma cópia foi enviada à você a título de registro!');" 
     ShowHTML "  location.href='" & conVirtualPath & w_dir & "Default.asp?EW=" & conWhatExFale & "&cl=" & cl & "';"
     ScriptClose
  End If

  Set w_html                   = Nothing
  Set w_destinatarios          = Nothing
  Set w_assunto                = Nothing
  Set RSM                      = Nothing
End Sub
REM =========================================================================
REM Fim da rotina da preparação para envio de e-mail
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da barra de navegação de recordsets
REM -------------------------------------------------------------------------
Sub MontaBarra (p_link, p_PageCount, p_AbsolutePage, p_PageSize, p_RecordCount)

  ShowHTML "<SCRIPT LANGUAGE=""JAVASCRIPT"">" & VbCrLf
  ShowHTML "  function pagina (pag) {" & VbCrLf
  ShowHTML "    document.Barra.htBT.value = pag;" & VbCrLf
  ShowHTML "    document.Barra.submit();" & VbCrLf
  ShowHTML "  }" & VbCrLf
  ShowHTML "</SCRIPT>" & VbCrLf
  ShowHTML "<center><FORM ACTION=""" & p_link &""" METHOD=""POST"" name=""Barra"">" & VbCrLf
  ShowHTML "<input type=""Hidden"" name=""str2"" value=""" & p_pagesize & """>" & VbCrLf
  ShowHTML "<input type=""Hidden"" name=""htBT"" value="""">" & VbCrLf
  For i = 1 to Request("H").Count
     ShowHTML "<input type=""Hidden"" name=""H"" value=""" & Request("H")(i) & """>" & VbCrLf
  Next
  If p_PageCount = p_AbsolutePage Then
     ShowHTML "<br>" & (p_RecordCount-((p_PageCount-1)*p_PageSize)) & " linhas apresentadas de " & p_RecordCount & " linhas" & VbCrLf
  Else
     ShowHTML "<br>" & p_PageSize & " linhas apresentadas de " & p_RecordCount & " linhas" & VbCrLf
  End If
  If p_PageSize < p_RecordCount Then
     ShowHTML "<br>na página " & p_AbsolutePage & " de " & p_PageCount & " páginas" & VbCrLf
     If p_AbsolutePage > 1 Then
        ShowHTML "<br>[<A TITLE=""Primeira página"" HREF=""javascript:pagina(1)"">Primeira</A>]&nbsp;" & VbCrLf
        ShowHTML "[<A TITLE=""Página anterior"" HREF=""javascript:pagina(" & p_AbsolutePage-1 & ")"">Anterior</A>]&nbsp;" & VbCrLf
     Else
        ShowHTML "<br>[Primeira]&nbsp;" & VbCrLf
        ShowHTML "[Anterior]&nbsp;" & VbCrLf
     End If
     If p_PageCount = p_AbsolutePage Then
        ShowHTML "[Próxima]&nbsp;" & VbCrLf
        ShowHTML "[Última]" & VbCrLf
     Else
        ShowHTML "[<A TITLE=""Página seguinte"" HREF=""javascript:pagina(" & p_AbsolutePage+1 & ")"">Próxima</A>]&nbsp;" & VbCrLf
        ShowHTML "[<A TITLE=""Última página"" HREF=""javascript:pagina(" & p_PageCount & ")"">Última</A>]" & VbCrLf
     End If
  Else
     ShowHTML "<br>na página " & p_AbsolutePage & " de " & p_PageCount & " página" & VbCrLf
  End If
  ShowHTML "</FORM></center>" & VbCrLf

End Sub

REM =========================================================================
REM Corpo Principal do programa
REM -------------------------------------------------------------------------
Private Sub Main

  If Request.QueryString("EW") = conWhatSenha Then
      ShowSenha
  End If

  Select Case sstrEW
    Case conWhatPrincipal   ShowPrincipal
    Case conWhatManut       ShowAluno
    Case conWhatQuem        ShowQuem
    Case conWhatExFale      ShowExFale
    Case conWhatArquivo     ShowArquivo
    Case conWhatExNoticia   ShowNoticia
    Case conWhatExCalend    ShowCalend
    Case conWhatValida      ShowEscolas
    Case "PreparaMail"      PreparaMail
  End Select
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>
