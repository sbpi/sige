<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/Constants_ADO.inc"-->
<!--#INCLUDE VIRTUAL="/modelos/Constants.inc"-->
<!--#INCLUDE VIRTUAL="/esc.inc"-->
<!--#INCLUDE VIRTUAL="/Funcoes.asp"-->
<%
Response.Expires = -1500
REM =========================================================================
REM  /default.asp
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

Private dbms, sobjConn

Private sstrSN
Private sstrCL

sstrSN = "default.asp"
sstrCL = "Aluno.asp"

Public sstrEF
Public sstrEW
Public sstrIN
Public CL

CL = int(Request.QueryString("CL"))
sstrEF = "sq_cliente=" & CL
sstrEW = Request.QueryString("EW")
sstrIN = int(Request.QueryString("IN"))
w_dir  = "Modelos/Mod16/"

Private sstrData

Set RS = Server.CreateObject("ADODB.RecordSet")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
Set RSCabecalho = Server.CreateObject("ADODB.RecordSet")
Set RSRodape = Server.CreateObject("ADODB.RecordSet")
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
ShowHTML "   <link href=""/css/estilo.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <link href=""/css/print.css""  media=""print""  rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
ShowHTML "</head>"
ShowHTML "<BASE HREF=""" & conSite & "/"">"
ShowHTML "<body>"

ShowHTML "    <div id=""pagina"">"
ShowHTML "    <div id=""topo""></div>"
%>
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
          <li><a target="_blank" class="aluno" href="http://www.se.df.gov.br/300/30002001.asp">Aluno</a></li>
          <li><a target="_blank" class="educador" href="http://www.se.df.gov.br/300/30003001.asp">Educador</a></li>
          <li><a target="_blank" class="comunidade" href="http://www.se.df.gov.br/300/30004001.asp">Comunidade</a></li>
        </ul>
      </div>
    </div>
    <div id="menuMiddle"> </div>
  </div>
  <div class="clear"></div>
<div id="menuBottom"> 
<%

ShowHTML "      <ul>"
ShowHTML "      <script language=""JavaScript"" src=""inc/mm_menu.js"" type=""text/JavaScript""></script>"
if(Request.QueryString("EW") = "110") Then
ShowHTML "      <li><a class=""selected"" href=""" & w_dir & "default.asp?EW=110&CL=" & replace(CL,"sq_cliente=","") & """ >Inicial</a> </li>"
else
ShowHTML "      <li><a href=""" & w_dir & "default.asp?EW=110&CL=" & replace(CL,"sq_cliente=","") & """ >Inicial</a> </li>"
end if
if(Request.QueryString("EW") = "113") Then
    ShowHTML "      <li><a class=""selected"" href=""" & w_dir & "default.asp?EW=113&CL=" & replace(CL,"sq_cliente=","") & """ id=""link1"">Quem somos</a> </li>"
else
    ShowHTML "      <li><a href=""" & w_dir & "default.asp?EW=113&CL=" & replace(CL,"sq_cliente=","") & """ id=""link1"">Quem somos</a> </li>"
end if
if(Request.QueryString("EW") = "117") Then
    ShowHTML "      <li><a class=""selected"" href=""" & w_dir & "default.asp?EW=117&CL=" & replace(CL,"sq_cliente=","") & """ >Fale conosco </a></li>"
else
    ShowHTML "      <li><a href=""" & w_dir & "default.asp?EW=117&CL=" & replace(CL,"sq_cliente=","") & """ >Fale conosco </a></li>"
end if
if(Request.QueryString("EW") = "116" AND Request.QueryString("IN") = "1") Then
    ShowHTML "      <li><a class=""selected"" href=""" & w_dir & "default.asp?EW=116&EF=" & CL & "&CL=" & replace(CL,"sq_cliente=","") & "&IN=1"" >Projeto</a></li>"
else
    ShowHTML "      <li><a href=""" & w_dir & "default.asp?EW=116&EF=" & CL & "&CL=" & replace(CL,"sq_cliente=","") & "&IN=1"" >Projeto</a></li>"
end if
if(Request.QueryString("EW") = "114") Then
    ShowHTML "      <li><a class=""selected"" href=""" & w_dir & "default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & """ id=""link2"">Notícias</a> </li>"
else
    ShowHTML "      <li><a href=""" & w_dir & "default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & """ id=""link2"">Notícias</a> </li>"
end if
if(Request.QueryString("EW") = "115") Then
    ShowHTML "      <li><a class=""selected"" href=""" & w_dir & "default.asp?EW=115&CL=" & replace(CL,"sq_cliente=","") & """ id=""link5"">Calendário</a>"
else    
    ShowHTML "      <li><a href=""" & w_dir & "default.asp?EW=115&CL=" & replace(CL,"sq_cliente=","") & """ id=""link5"">Calendário</a>"
end if
if(Request.QueryString("EW") = "143") Then
    ShowHTML "      <li><a class=""selected"" href=""" & w_dir & "default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & """ >Arquivos (<i>download</i>)</a></li>"
else
    ShowHTML "      <li><a href=""" & w_dir & "default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & """ >Arquivos (<i>download</i>)</a></li>"
end if
if(Request.QueryString("EW") = "116" AND Request.QueryString("IN") = "0") Then
    ShowHTML "      <li><a class=""selected"" href=""" & w_dir & "default.asp?EW=116&IN=0&CL=" & replace(CL,"sq_cliente=","") & """ >Alunos</a></li>"
else
    ShowHTML "      <li><a href=""" & w_dir & "default.asp?EW=116&IN=0&CL=" & replace(CL,"sq_cliente=","") & """ >Alunos</a></li>"
end if
if(Request.QueryString("EW") = "oferta") Then
    ShowHTML "      <li><a class=""selected"" href=""" & w_dir & "default.asp?EW=oferta&EF=" & CL & "&CL=" & replace(CL,"sq_cliente=","") & "&IN=1"" >Oferta</a></li>"
else
    ShowHTML "      <li><a href=""" & w_dir & "default.asp?EW=oferta&EF=" & CL & "&CL=" & replace(CL,"sq_cliente=","") & "&IN=1"" >Oferta</a></li>"
end if
ShowHTML "      <li><a target=""_blank"" href=""http://siade.cesgranrio.org.br"" >Questionário SIADE</a></li>"
ShowHTML "	    </ul>"
ShowHTML "	    <div class=""clear""></div>"

ShowHTML "      </div>"

sql = "SELECT a.sq_cliente, a.ds_cliente, b.ds_mensagem FROM escCliente a inner join esccliente_site b on (a.sq_cliente = b.sq_cliente)WHERE a." & sstrEF
RS.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly
ShowHTML "  <div id=""conteudo"">"
ShowHTML "      <h2>" & RS("ds_cliente") & "</h2>"
Select Case sstrEW
  Case conWhatPrincipal   ShowHTML "      <h3>Inicial</h3>"
  Case conWhatManut       ShowHTML "      <h3>Alunos</h3>"
  Case conWhatQuem        ShowHTML "      <h3>Quem somos</h3>"
  Case conWhatExFale      ShowHTML "      <h3>Fale conosco</h3>"
  Case conWhatArquivo     ShowHTML "      <h3>Arquivos</h3>"
  Case conWhatExNoticia   ShowHTML "      <h3>Notícias</h3>"
  Case conWhatExCalend    ShowHTML "      <h3>Calendário</h3>"
  Case "oferta"           ShowHTML "      <h3>Oferta</h3>"
  Case conWhatVah3da
     ShowHTML "<ul>"
     If sstrIN > 0 Then
        ShowHTML "      <li><b>Projeto</b></li>"
     Else
        ShowHTML "      <li><b>Alunos</b></li>"
     End If
     ShowHTML "</ul>"
End Select

RS.Close

Main

ShowHTML "  </div> "
ShowHTML "  <div class=""clear""></div>"

ShowHTML "  </div> "
%>
<div id="rodape">
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

Set dbms              = nothing
Set RS            = nothing
Set RSCabecalho   = nothing
Set RSRodape      = nothing
Set sobjConn          = nothing

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
        "       dt_arquivo, ds_titulo, nr_ordem, ds_arquivo, ln_arquivo, 'Escola' AS Origem, x.ln_internet diretorio " & VbCrLf & _
        "From escCliente_Arquivo a, escCliente x " & VbCrLf & _
        "WHERE in_ativo = 'Sim'" & VbCrLf & _
        "  AND x." & sstrEF & " " & VbCrLf & _
        "  AND a." & replace(sstrEF,"sq_cliente","sq_site_cliente") & " " & VbCrLf & _
        "  and in_destinatario <> 'E'" & VbCrLf & _
        "  and YEAR(a.dt_arquivo) = " & wAno & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT case a.in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end AS in_destinatario, " & VbCrLf & _
        "       a.dt_arquivo, a.ds_titulo, a.nr_ordem, ds_arquivo, ln_arquivo, 'SEDF' AS Origem, d.ds_diretorio diretorio " & VbCrLf & _
        "From escCliente_Arquivo AS a INNER JOIN escCliente AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente AS c ON (c.sq_cliente_pai  = b.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente AS e ON (e.sq_cliente_pai  = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente_Site AS d ON (b.sq_cliente = d.sq_site_cliente) " & VbCrLf & _
        "WHERE a.in_ativo = 'Sim'" & VbCrLf & _
        "  AND e." & sstrEF & " " & VbCrLf & _
        "  and in_destinatario <> 'E'" & VbCrLf & _
        "  and YEAR(a.dt_arquivo) = " & wAno & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT case a.in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end AS in_destinatario, " & VbCrLf & _
        "       a.dt_arquivo, a.ds_titulo, a.nr_ordem, ds_arquivo, ln_arquivo, 'Regional' AS Origem, d.ds_diretorio diretorio " & VbCrLf & _
        "From escCliente_Arquivo AS a INNER JOIN escCliente AS c ON (a.sq_site_cliente = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente AS e ON (e.sq_cliente_pai  = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente_Site AS d ON (c.sq_cliente = d.sq_cliente) " & VbCrLf & _
        "WHERE a.in_ativo = 'Sim'" & VbCrLf & _
        "  AND e." & sstrEF & " " & VbCrLf & _
        "  and in_destinatario <> 'E'" & VbCrLf & _
        "  and YEAR(a.dt_arquivo) = " & wAno & VbCrLf & _
        "ORDER BY dt_arquivo desc, origem, nr_ordem, in_destinatario " & VbCrLf 
  RS.Open sql, sobjConn, adOpenForwardOnly

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
  ShowHTML "  </TABLE>"

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
        "  and in_destinatario <> 'E'" & VbCrLf & _
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
    RS.Open sql, sobjConn, adOpenForwardOnly
    If Not RS.EOF Then
       ShowHTML "  <TR><TD><b>Arquivos de outros anos</b><br>"
       While Not RS.EOF
          ShowHTML "     <li><a href=""" & w_dir & "default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & "&wAno=" & RS("ano") & """ >Arquivos de " & RS("ano") & "</a>"
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
        "                     INNER JOIN escCliente_Dados as d on (a.sq_cliente = d.sq_cliente) " & _
        "                     LEFT  JOIN escCliente_Foto  as e on (a.sq_cliente = e.sq_cliente and e.tp_foto = 'P') " & _
        "WHERE a." & sstrEF & " " & VbCrLf & _
        "ORDER BY e.nr_ordem"
        
  RS.Open sql, sobjConn, adOpenForwardOnly

  If Not RS.EOF Then
    If RS("ds_texto_abertura") > "" Then
       ShowHTML "        <p class=""chamada"">" & replace(RS("ds_texto_abertura"),chr(13)&chr(10), "<br/>") & "</p>"
    Else
       ShowHTML "        <p class=""chamada"">"
    End If
  If RS("sq_cliente_foto") > "" Then
       ShowHTML "    <ul class=""fotos"">"   
       Do While Not RS.EOF
          ShowHTML "    <li><a href=""" & RS("ds_diretorio") & "/" & RS("ln_foto") & """ target=""_blank"" title=""Clique sobre a imagem para ampliar""><img align=""top"" class=""foto"" src=""" & RS("ds_diretorio") & "/" & RS("ln_foto") & """ width=""302"" height=""206""><br>" & RS("ds_foto")& "</a></li>"
          RS.MoveNext
       Loop
       ShowHTML "         </ul>"
  End If
    
  End If

  RS.Close
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Página Principal
REM =========================================================================

REM =========================================================================
REM Monta a tela Quem Somos
REM -------------------------------------------------------------------------
Public Sub ShowQuem
  Dim sql, strNome, itemmenu
  itemmenu = "Quem Somos"

   sql = "SELECT b.ds_institucional, b.ds_diretorio, e.sq_cliente_foto, e.ln_foto, e.ds_foto " & _
        "From escCliente as a INNER JOIN escCLIENTE_SITE  as b on (a.sq_cliente = b.sq_cliente) " & _
        "                     INNER JOIN escModelo        as c on (b.sq_modelo = c.sq_modelo) " & _
        "                     LEFT  JOIN escCliente_Dados as d on (a.sq_cliente = d.sq_cliente) " & _
        "                     LEFT  JOIN escCliente_Foto  as e on (a.sq_cliente = e.sq_cliente and e.tp_foto = 'Q') " & _
        "WHERE a." & sstrEF & " " & VbCrLf & _
        "ORDER BY e.nr_ordem"

  RS.Open sql, sobjConn, adOpenForwardOnly
  
  ShowHTML "    <tr valign=""top""><td> "
  
  If Not RS.EOF Then
    If RS("ds_institucional") > "" Then
       ShowHTML "        <p align=""left"" style=""margin-top: 0; margin-bottom: 0"">" & replace(RS("ds_institucional"),chr(13)&chr(10), "<br>")
    Else
       ShowHTML "        <P align=justify>"
    End If
    ShowHTML "        </td>"
    If RS("sq_cliente_foto") > "" Then
       ShowHTML "    <td align=""right""><font size=""1""><br/><b>Fotografias</b><font size=1></b><br>"
       ShowHTML "    <ul class=""fotos"">"   
       Do While Not RS.EOF
          ShowHTML "    <li><a href=""" & RS("ds_diretorio") & "/" & RS("ln_foto") & """ target=""_blank"" title=""Clique sobre a imagem para ampliar""><img align=""top"" class=""foto"" src=""" & RS("ds_diretorio") & "/" & RS("ln_foto") & """ width=""302"" height=""206""><br>" & RS("ds_foto")& "</a></li>"
          RS.MoveNext
       Loop
       ShowHTML "    </ul>"
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
        "                     INNER JOIN escCLIENTE_Dados as c on (a.sq_cliente = c.sq_cliente) " & _
        "WHERE a." & sstrEF 

  RS.Open sql, sobjConn, adOpenForwardOnly

  ShowHTML "<tr><td>"

  ShowHTML "<p align=""justify"">Informações, sugestões e reclamações podem ser feitas utilizando os dados abaixo:</p>"

  If Not RS.EOF Then
    Do While Not RS.EOF

  ShowHTML "<table border=""0"" cellspacing=""3"">"
  ShowHTML "  <tr>"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td width=""20%"" align=""right""><b>Diretor(a):"
  ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_diretor") 
  ShowHTML "  </tr>"
  ShowHTML "  <tr>"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td width=""20%"" align=""right""><b>Secretário(a):"
  ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_secretario") 
  ShowHTML "  </tr>"
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

      RS.MoveNext

    Loop

  End If

  ShowHTML "</tr>"

  RS.Close

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

  If sstrIN = 0 then
     sql = "SELECT a.sq_noticia, a.dt_noticia AS data, a.ds_titulo AS ocorrencia, 'Escola' AS origem " & VbCrLf & _
           "  FROM escNoticia_Cliente  a" & VbCrLf & _
           " WHERE sq_site_cliente   = " & CL & VbCrLf & _
           "  and a.in_ativo         = 'Sim'" & VbCrLf & _
           "  and a.in_exibe         = 'Sim'" & VbCrLf & _
           "  and YEAR(a.dt_noticia) = " & wAno & VbCrLf & _
           "UNION " & VbCrLf & _
           "SELECT a.sq_noticia, dt_noticia AS data, ds_titulo AS ocorrencia, 'Regional' AS origem " & VbCrLf & _
           "FROM escNoticia_Cliente    AS a" & VbCrLf & _
           "     INNER JOIN escCliente AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
           "     INNER JOIN escCliente AS c ON (b.sq_cliente      = c.sq_cliente_pai) " & VbCrLf & _
           "WHERE c.sq_cliente       = " & CL & VbCrLf & _
           "  and a.in_ativo         = 'Sim'" & VbCrLf & _
           "  and a.in_exibe         = 'Sim'" & VbCrLf & _
           "  and YEAR(a.dt_noticia) = " & wAno & VbCrLf & _
           "UNION " & VbCrLf & _
           "SELECT a.sq_noticia, dt_noticia AS data, ds_titulo AS ocorrencia, 'SEDF' AS origem " & VbCrLf & _
           "FROM escNoticia_Cliente    AS a" & VbCrLf & _
           "     INNER JOIN escCliente AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
           "     INNER JOIN escCliente AS c ON (b.sq_cliente      = c.sq_cliente_pai) " & VbCrLf & _
           "     INNER JOIN escCliente AS d ON (c.sq_cliente      = d.sq_cliente_pai) " & VbCrLf & _
           "WHERE d.sq_cliente       = " & CL & VbCrLf & _
           "  and a.in_ativo         = 'Sim'" & VbCrLf & _
           "  and a.in_exibe         = 'Sim'" & VbCrLf & _
           "  and YEAR(a.dt_noticia) = " & wAno & VbCrLf & _
           "ORDER BY data desc, ocorrencia" & VbCrLf
     RS.Open sql, sobjConn, adOpenForwardOnly
     
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
           ShowHTML "    <TD align=""center"">" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2) & "/" &Year(RS("data"))
           ShowHTML "    <TD align=""center"">" & RS("origem")
           ShowHTML "    <TD colspan=2>" & RS("ocorrencia") & "&nbsp;&nbsp;&nbsp;<a href=""" & w_dir & sstrSN & "?EW=" & conWhatExNoticia & "&CL=" & CL & "&EF=" & sstrEF & "&IN=" & RS("sq_noticia") & """><b>[Ler]</a>"
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
           "UNION " & VbCrLf & _
           "SELECT YEAR(a.dt_noticia) ano " & VbCrLf & _
           "FROM escNoticia_Cliente    AS a" & VbCrLf & _
           "     INNER JOIN escCliente AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
           "     INNER JOIN escCliente AS c ON (b.sq_cliente      = c.sq_cliente_pai) " & VbCrLf & _
           "     INNER JOIN escCliente AS d ON (c.sq_cliente      = d.sq_cliente_pai) " & VbCrLf & _
           "WHERE d.sq_cliente       = " & CL & VbCrLf & _
           "  and a.in_ativo         = 'Sim'" & VbCrLf & _
           "  and a.in_exibe         = 'Sim'" & VbCrLf & _
           "  and YEAR(a.dt_noticia) <> " & wAno & VbCrLf & _
           "ORDER BY YEAR(dt_noticia) desc" & VbCrLf
     RS.Open sql, sobjConn, adOpenForwardOnly
     If Not RS.EOF Then
        ShowHTML "  <TR valign=""top""><TD colspan=""5""><br><b>Notícias de outros anos</b><br>"     
        While Not RS.EOF
           ShowHTML "    <li><a href=""" & w_dir & "default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & "&wAno=" & RS("ano") & """ id=""link2"">Notícias de " & RS("ano") & "</a> </TD>"
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
     RS.Open sql, sobjConn, adOpenForwardOnly

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

  Dim sql, strNome, wCont, wRede, wAno, wDatas(31,12,10), wCores(31,12,10), wImagem(31,12,10)
  wAno = Year(Date())
  
  sql = "SELECT '' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'B' AS origem FROM escCalendario_base a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE YEAR(dt_ocorrencia)=" & Year(Date()) & " " & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT '#99CCFF' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE sq_site_cliente = " & CL & "  AND YEAR(dt_ocorrencia)= " & Year(Date()) & " " & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT '#FFFF99' as cor, e.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'R' AS origem " & VbCrLf & _
        "FROM escCalendario_Cliente   AS a" & VbCrLf & _
        "     INNER JOIN escCliente   AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
        "     INNER JOIN escCliente   AS c ON (b.sq_cliente      = c.sq_cliente_pai) " & VbCrLf & _
        "     INNER JOIN escCliente   AS d ON (c.sq_cliente      = d.sq_cliente_pai) " & VbCrLf & _
        "     LEFT  JOIN escTipo_Data AS e ON (a.sq_tipo_data    = e.sq_tipo_data) " & VbCrLf & _
        "WHERE d.sq_cliente = " & CL & VbCrLf & _
        "  AND YEAR(dt_ocorrencia)= " & wAno & " " & VbCrLf & _
        "ORDER BY data, origem desc, ocorrencia" & VbCrLf 

  RS1.Open sql, sobjConn, adOpenForwardOnly
  
  
  If sstrIN = 0 then
     If Not RS1.EOF Then
        Do While Not RS1.EOF
           If RS1("origem") = "E" then
              wDatas(Day(RS1("data")), Month(RS1("data")), Mid(Year(RS1("data")),4,2)) = RS1("ocorrencia") & " (Origem: Escola)"
              wImagem(Day(RS1("data")), Month(RS1("data")), Mid(Year(RS1("data")),4,2))= RS1("imagem")
              wCores(Day(RS1("data")), Month(RS1("data")), Mid(Year(RS1("data")),4,2)) = RS1("cor")
           ElseIf RS1("origem") = "R" then
              wDatas(Day(RS1("data")), Month(RS1("data")), Mid(Year(RS1("data")),4,2)) = RS1("ocorrencia") & " (Origem: SEDF)"
              wImagem(Day(RS1("data")), Month(RS1("data")), Mid(Year(RS1("data")),4,2)) = RS1("imagem")
              wCores(Day(RS1("data")), Month(RS1("data")), Mid(Year(RS1("data")),4,2)) = RS1("cor")
           Else
              wDatas(Day(RS1("data")), Month(RS1("data")), Mid(Year(RS1("data")),4,2)) = RS1("ocorrencia") & " (Origem: Oficial)"
              wImagem(Day(RS1("data")), Month(RS1("data")), Mid(Year(RS1("data")),4,2)) = RS1("imagem")
           End If
           RS1.MoveNext
        Loop
        RS1.MoveFirst
     End If
     ShowHTML "<tr><td><TABLE align=""center"" width=""100%"" border=0 cellSpacing=0>"
     ShowHTML "<tr valign=""top"">"
     ShowHTML "  <td>" & MontaCalendario("01" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("02" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("03" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("04" & wAno, wDatas, wCores, wImagem)
     ShowHTML "<tr valign=""top"">"
     ShowHTML "  <td>" & MontaCalendario("05" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("06" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("07" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("08" & wAno, wDatas, wCores, wImagem)
     ShowHTML "<tr valign=""top"">"
     ShowHTML "  <td>" & MontaCalendario("09" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("10" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("11" & wAno, wDatas, wCores, wImagem)
     ShowHTML "  <td>" & MontaCalendario("12" & wAno, wDatas, wCores, wImagem)
     ShowHTML "</table>"
     
     ShowHTML "<tr><td colspan=""2""><TABLE width=""100%"" align=""center"" border=0 cellSpacing=1>"
     ShowHTML "<tr valign=""middle"" align=""center"">"
     ShowHTML "  <td><font size=1><b>LEGENDA</b></font>"
     ShowHTML "  <td><font size=1><b>FERIADOS</b></font>"
     ShowHTML "  <td><font size=1><b>RECESSOS</b></font>"
     ShowHTML "  <td><font size=1><b>DIAS LETIVOS</b></font>"
     ShowHTML "</tr>"
     ShowHTML "  <TR>"
     ShowHTML "    <TD CHEIGHT=""1"" BGCOLOR=""#DAEABD"">"
     ShowHTML "  </TR>"
     
     ShowHTML "<tr valign=""top"">"
     'Legenda
     sql = "SELECT * from escTipo_Data where abrangencia <> 'U' order by nome " & VbCrLf
     RS.Open sql, sobjConn, adOpenForwardOnly

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
        ShowHTML "     </TABLE>"
     Else
        ShowHTML "  <td>Sem informação"
     End If
     
     RS.Close
     'Feriados
     sql = "SELECT '' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'B' AS origem FROM escCalendario_base a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE YEAR(dt_ocorrencia)=" & wAno & " AND b.sigla <> 'CN' " & VbCrLf & _
           "ORDER BY data, origem desc, ocorrencia" & VbCrLf
     RS.Open sql, sobjConn, adOpenForwardOnly
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
     RS.Close
     
     ' Exibe recessos
     sql = "SELECT '#99CCFF' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla in ('RE','RA')) WHERE sq_site_cliente = " & CL & "  AND YEAR(dt_ocorrencia)= " & Year(Date()) & " " & VbCrLf & _
           "UNION " & VbCrLf & _
           "SELECT '#FFFF99' as cor, e.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'R' AS origem " & VbCrLf & _
           "FROM escCalendario_Cliente   AS a" & VbCrLf & _
           "     INNER JOIN escCliente   AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
           "     INNER JOIN escCliente   AS c ON (b.sq_cliente      = c.sq_cliente_pai) " & VbCrLf & _
           "     INNER JOIN escCliente   AS d ON (c.sq_cliente      = d.sq_cliente_pai) " & VbCrLf & _
           "     INNER  JOIN escTipo_Data AS e ON (a.sq_tipo_data    = e.sq_tipo_data and e.sigla in ('RE','RA')) " & VbCrLf & _
           "WHERE d.sq_cliente = " & CL & VbCrLf & _
           "  AND YEAR(dt_ocorrencia)= " & wAno & " " & VbCrLf & _
           "ORDER BY data, origem desc, ocorrencia" & VbCrLf

            
     RS.Open sql, sobjConn, adOpenForwardOnly
  
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
        ShowHTML "    </TABLE></td>"
     Else
        ShowHTML "  <td>Sem informação"
     End If
     RS.Close
     ShowHTML "  <td valign=""top""><TABLE align=""center"" border=0 cellSpacing=1>"
     
      ' Recupera o ano letivo e o período de recesso
     sql = "  select * from " & VbCrLf & _
           "        (select dt_ocorrencia w_let_ini " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  wAno &  "and b.sigla = 'IA') " & VbCrLf & _
           "          where sq_site_cliente = " & 0 & VbCrLf & _
           "        ) a, " & VbCrLf & _
           "        (select dt_ocorrencia w_let_fim " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  wAno &  " and b.sigla = 'TA') " & VbCrLf & _
           "          where sq_site_cliente = " & 0 & VbCrLf & _
           "        ) b, " & VbCrLf & _
           "        (select dt_ocorrencia w_let2_ini " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  wAno &  " and b.sigla = 'I2') " & VbCrLf & _
           "          where sq_site_cliente = " & 0 & VbCrLf & _
           "        ) c, " & VbCrLf & _
           "        (select dt_ocorrencia w_let1_fim " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  wAno &  " and b.sigla = 'T1') " & VbCrLf & _
           "          where sq_site_cliente = " & 0 & VbCrLf & _
           "        ) d " & VbCrLf
     RS.Open sql, sobjConn, adOpenForwardOnly
     Do While Not RS.EOF
       w_fim1 = RS("w_let1_fim")
       w_ini2 = RS("w_let2_ini")
       RS.MoveNext
     Loop
     Rs.Close
     
     For w_cont = 1 to 2
        If w_cont = 1 Then
           w_ini = 1
           w_fim = 7
        Else
           w_ini = 7
           w_fim = 12
           ShowHTML "  <TR valign=""top"" ALIGN=""CENTER""><TD COLSPAN=""4""><br/><br/></td></tr>"
        End If
        w_dias = 0
      
        ShowHTML "  <TR valign=""top"" ALIGN=""CENTER""><TD COLSPAN=""4""><b>" & w_cont & "º Semestre</b></td></tr>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD><b>MÊS"
        ShowHTML "    <TD><b>D.L."
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
        ShowHTML "    <TD COLSPAN=""2"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
        ShowHTML "  </TR>"
        For wCont = w_ini to w_fim
           If month(w_ini2) = wCont and w_cont = 2Then
              w_inicial = formataDataEdicao(w_ini2)
           Else
              w_inicial = "01/" & mid(100+wCont,2,2) & "/" & year(date())
           End If
           If month(w_fim1) = wCont and w_cont = 1 Then
              w_final = formataDataEdicao(w_fim1)
           Else
              sql = "SELECT convert(varchar,dbo.lastDay(convert(datetime,'01/" & mid(100+wCont,2,2) & "/'+cast(year(getDate()) as varchar),103)),103) fim" & VbCrLf 
              RS.Open sql, sobjConn, adOpenForwardOnly
              w_final = RS("fim")
              RS.Close 
           End If
           
           sql = "SELECT dbo.diasLetivos('" & w_inicial & "', '" & w_final & "'," & CL & ",null) qtd" & VbCrLf 
           RS.Open sql, sobjConn, adOpenForwardOnly
           If RS("qtd") > 0 Then
              ShowHTML "  <TR>"
              ShowHTML "    <TD>" & nomeMes(wCont)
              ShowHTML "    <TD ALIGN=""CENTER"">" & RS("qtd")
              w_mes = RS("qtd")
              ShowHTML "  </TR>"
              w_dias = w_dias + w_mes
           End If
           RS.Close
           
        Next
        ShowHTML "  <TR>"
        ShowHTML "    <TD nowrap>Dias Letivos"
        ShowHTML "    <TD ALIGN=""CENTER"">" & w_dias
        ShowHTML "  </TR>"
        
     Next        
     ShowHTML "    </TABLE></td>"
     ShowHTML "    </TABLE>"
%><!--
     ShowHTML "<tr><td valign=""top""><TABLE width=""100%"" align=""center"" border=0 cellSpacing=1>"
     ShowHTML "<tr valign=""top""><td><TABLE width=""90%"" align=""center"" border=0 cellSpacing=1>"
     ShowHTML "  <TR valign=""top"">"
     ShowHTML "    <TD width=""2%"">&nbsp;"
     ShowHTML "    <TD width=""15%""><b>Origem"
     ShowHTML "    <TD width=""15%""><b>Data"
     ShowHTML "    <TD width=""68%""><b>Ocorrência"
     ShowHTML "  </TR>"
     ShowHTML "  <TR>"
     ShowHTML "    <TD COLSPAN=""4"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
     ShowHTML "  </TR>"
     wCont = 0
     If Not RS1.EOF Then
        wCont = 1
        Do While Not RS1.EOF

           ShowHTML "  <TR VALIGN=""TOP"">"
           ShowHTML "    <TD>&nbsp;"
           If RS1("origem") = "E" then
              ShowHTML "    <TD>Escola"
              ShowHTML "    <TD>" & Mid(100+Day(RS1("data")),2,2) & "/" & Mid(100+Month(RS1("data")),2,2)
              ShowHTML "    <TD><font color=""#0000FF"">" & RS1("ocorrencia")
           ElseIf RS1("origem") = "R" then
              ShowHTML "    <TD>SEDF"
              ShowHTML "    <TD>" & Mid(100+Day(RS1("data")),2,2) & "/" & Mid(100+Month(RS1("data")),2,2)
              ShowHTML "    <TD>" & RS1("ocorrencia")
           Else
              ShowHTML "    <TD>Oficial"
              ShowHTML "    <TD>" & Mid(100+Day(RS1("data")),2,2) & "/" & Mid(100+Month(RS1("data")),2,2)
              ShowHTML "    <TD>" & RS1("ocorrencia")
           End If
           ShowHTML "  </TR>"
           wCont = wCont + 1
           RS1.MoveNext
        Loop

     End If

     ShowHTML "    </TABLE>"
 
     RS1.Close
     ShowHTML "  <td><TABLE width=""10%"" align=""center"" border=0 cellSpacing=1>"
     ShowHTML "  <TR valign=""top"" ALIGN=""CENTER"">"
     ShowHTML "    <TD><b>Mês"
     ShowHTML "    <TD NOWRAP><b>Dias letivos"
     ShowHTML "  </TR>"
     ShowHTML "  <TR>"
     ShowHTML "    <TD COLSPAN=""2"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
     ShowHTML "  </TR>"
     For wCont = 1 to 12
        sql = "SELECT dbo.diasLetivos('01/" & mid(100+wCont,2,2) & "/'+cast(year(getDate()) as varchar),convert(varchar,dbo.lastDay(convert(datetime,'01/" & mid(100+wCont,2,2) & "/'+cast(year(getDate()) as varchar),103)),103)," & CL & ") qtd" & VbCrLf 
        RS1.Open sql, sobjConn, adOpenForwardOnly
        
        If RS1("qtd") > 0 Then
           ShowHTML "  <TR>"
           ShowHTML "    <TD ALIGN=""CENTER"">" & wCont & "/" & year(date())
           ShowHTML "    <TD ALIGN=""CENTER"">" & RS1("qtd")
           ShowHTML "  </TR>"
        End If

        RS1.Close
     Next

     ShowHTML "    </TABLE>"
     ShowHTML " </TABLE>"
-->
<%
  End If
End Sub
REM -------------------------------------------------------------------------
REM Final da Página Calendário
REM =========================================================================

REM =========================================================================
REM Monta a tela de Validação de Senha
REM -------------------------------------------------------------------------
Public Sub ShowValida

  Dim strLabel, strTexto, strAction, strCampo
  
  If sstrIN > 0 Then
     sql = "SELECT b.ds_diretorio, b.ln_prop_pedagogica " & _
           "From escCliente as a INNER JOIN escCLIENTE_SITE as b on (a.sq_cliente = b.sq_cliente) " & _
           "                     INNER JOIN escModelo as c on (b.sq_modelo = c.sq_modelo) " & _
           "                     INNER JOIN escCliente_Dados as d on (a.sq_cliente = d.sq_cliente) " & _
           "WHERE a." & sstrEF 

     RS.Open sql, sobjConn, adOpenForwardOnly
     ShowHTML "<tr><td>"
     If Nvl(RS("ln_prop_pedagogica"),"") > "" Then
        ShowHTML "    <font size=1><b>"
        If InStr(lCase(RS("ln_prop_pedagogica")),lCase(".pdf")) > 0 Then
           ShowHTML "        A exibição do arquivo exige que o Acrobat Reader tenha sido instalado em seu computador."
           ShowHTML "        <br>Se o arquivo não for exibido no quadro abaixo, clique <a href=""http://www.adobe.com.br/products/acrobat/readstep2.html"" target=""_blank"">aqui</a> para instalar ou atualizar o Acrobat Reader."
        Else
           ShowHTML "        A exibição do arquivo exige o editor de textos Word ou equivalente. "
           ShowHTML "        <br>Se o arquivo não for exibido no quadro abaixo, verifique se o Word foi corretamente instalado em seu computador."        
        End If
        ShowHTML "<table align=""center"" width=""100%"" cellspacing=0 style=""border: 1px solid rgb(0,0,0);""><tr><td>"
        ShowHTML "    <iframe src=""" & RS("ds_diretorio") & "/" & RS("ln_prop_pedagogica") & """ width=""100%"" height=""510"">"
        ShowHTML "    </iframe>"
        ShowHTML "</table>"
     Else
        ShowHTML "    <font size=1><b>Projeto não informado."
     End If
  Else
     strLabel = "Alunos"
     strTexto = "Informe nos campos abaixo sua Matrícula e Senha de Acesso, conforme informado pela escola. Se você não recebeu esses dados, clique no botão <i>Fale Conosco</I>, acima, para ver como entrar em contato com a escola e consegui-los."
     strAction = sstrSN & "?EW=" & conWhatSenha & "&EF=" & sstrEF & "&CL=" & CL & "&IN=" & sstrIN
     strCampo = "Matrícula"

     ShowHTML "<script Language=""JavaScript"">" & chr(13)
     ShowHTML "<!--" & chr(13)
     ShowHTML "function Validacao(theForm)" & chr(13)
     ShowHTML "{" & chr(13)
     ShowHTML "  var checkOK = ""ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-"";" & chr(13)
     ShowHTML "  var checkStr = theForm.UID.value;" & chr(13)
     ShowHTML "  var allValid = true;" & chr(13)
     ShowHTML "  for (i = 0;  i < checkStr.length;  i++)" & chr(13)
     ShowHTML "  {" & chr(13)
     ShowHTML "    ch = checkStr.charAt(i);" & chr(13)
     ShowHTML "    for (j = 0;  j < checkOK.length;  j++)" & chr(13)
     ShowHTML "      if (ch == checkOK.charAt(j))" & chr(13)
     ShowHTML "        break;" & chr(13)
     ShowHTML "    if (j == checkOK.length)" & chr(13)
     ShowHTML "    {" & chr(13)
     ShowHTML "      allValid = false;" & chr(13)
     ShowHTML "      break;" & chr(13)
     ShowHTML "    }" & chr(13)
     ShowHTML "  }" & chr(13)
     ShowHTML "  if (!allValid)" & chr(13)
     ShowHTML "  {" & chr(13)
     ShowHTML "    alert(""Informe apenas letras e números no campo \""Nome de Usuário\""."");" & chr(13)
     ShowHTML "    theForm.UID.focus();" & chr(13)
     ShowHTML "    return (false);" & chr(13)
     ShowHTML "  }" & chr(13)
     ShowHTML "  var checkOK = ""ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-"";" & chr(13)
     ShowHTML "  var checkStr = theForm.PWD.value;" & chr(13)
     ShowHTML "  var allValid = true;" & chr(13)
     ShowHTML "  for (i = 0;  i < checkStr.length;  i++)" & chr(13)
     ShowHTML "  {" & chr(13)
     ShowHTML "    ch = checkStr.charAt(i);" & chr(13)
     ShowHTML "    for (j = 0;  j < checkOK.length;  j++)" & chr(13)
     ShowHTML "      if (ch == checkOK.charAt(j))" & chr(13)
     ShowHTML "        break;" & chr(13)
     ShowHTML "    if (j == checkOK.length)" & chr(13)
     ShowHTML "    {" & chr(13)
     ShowHTML "      allValid = false;" & chr(13)
     ShowHTML "      break;" & chr(13)
     ShowHTML "    }" & chr(13)
     ShowHTML "  }" & chr(13)
     ShowHTML "  if (!allValid)" & chr(13)
     ShowHTML "  {" & chr(13)
     ShowHTML "    alert(""Informe apenas letras e números no campo \""Senha\""."");" & chr(13)
     ShowHTML "    theForm.PWD.focus();" & chr(13)
     ShowHTML "    return (false);" & chr(13)
     ShowHTML "  }" & chr(13)
     ShowHTML "  return (true);" & chr(13)
     ShowHTML "}" & chr(13)
     ShowHTML "//-->" & chr(13)
     ShowHTML "</script>" & chr(13)
     ShowHTML "<form method=""POST"" action=""" & w_dir & strAction & """ onsubmit=""return Validacao(this)"" name=""Form1"">"

     ShowHTML "<tr><td><table align=""center"" width=""95%"">"
     ShowHTML "<tr><td colspan=""3""><p align=""justify"">" & strTexto & "<p>&nbsp"
     ShowHTML "<tr><td align=""right""><b>" & strCampo & ":</b><td>&nbsp;<td><input type=""text"" class=""texto"" lenght=""14"" maxsize=""14"" name=""UID"" value="""">"
     ShowHTML "<tr><td align=""right""><b>Senha de acesso:</b><td>&nbsp;<td><input type=""password"" class=""texto"" lenght=""14"" maxsize=""14"" name=""PWD"" value="""">"
     ShowHTML "<tr><td align=""right"">&nbsp;<td>&nbsp;<td><input type=""submit"" name=""BTN"" class=""botao"" value=""Encontrar"">&nbsp;<input type=""reset"" class=""botao"" name=""CLR"" value=""Limpar"">"
     ShowHTML "</table>"

     ShowHTML "</form>"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Validação de Senha
REM =========================================================================

REM =========================================================================
REM Monta a tela de Verificação de Senha
REM -------------------------------------------------------------------------
Public Sub ShowSenha

  Dim sql, strNome, strDestino, w_uid, w_pwd
  
  w_uid = replace(replace(Trim(uCase(Request("UID"))),"'", ""), """", "")
  w_pwd = replace(replace(Trim(uCase(Request("PWD"))),"'", ""), """", "")

  If sstrIN = 0 Then
     sql = "SELECT * FROM escAluno " & VbCrLf & _
           "WHERE sq_site_cliente = " & CL & VbCrLf & _
           "  AND NR_MATRICULA   ='" & w_uid & "'" & VbCrLf & _
           "  AND DS_SENHA_ACESSO='" & w_pwd & "'" & VbCrLf
  Else
     sql = "SELECT * from escCliente " & VbCrLf & _
           "WHERE sq_cliente     = " & CL & VbCrLf & _
           "  AND DS_USERNAME    ='" & w_uid & "'" & VbCrLf & _
           "  AND DS_SENHA_ACESSO='" & w_pwd & "'" & VbCrLf
  End If

  RS.Open sql, sobjConn, adOpenForwardOnly

  If RS.EOF Then

    ShowHTML "<tr><td align=""right""><b><font size=2>Validação</font></b></td></tr>"

    ShowHTML "<tr><td><p align=""justify"">Nome de usuário ou senha de acesso inválida. Volte à página anterior para tentar novamente.</p>"
    ShowHTML "<p><center><img src=""Img/bt_voltar.gif"" border=0 valign=""center"" onClick=""history.go(-1)"" alt=""Volta"">"
    ShowHTML "</tr>"

  Else
    If sstrIN = 0 Then
       ' Grava o acesso na tabela de log
       w_chave = RS("sq_aluno")
       SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql) " & VbCrLf & _
             "values ( " & VbCrLf & _
             "         " & CL & ", " & VbCrLf & _
             "         getdate(), " & VbCrLf & _
             "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
             "         0, " & VbCrLf & _
             "         'Acesso à tela do aluno " & RS("no_aluno") & " (" & RS("nr_matricula") & ").', " & VbCrLf & _
             "         null " & VbCrLf & _
             "       ) " & VbCrLf
       AbreSessao
       ExecutaSQL(SQL)
       ShowHTML "<SCRIPT LANGUAGE=""JAVASCRIPT"">" & VbCrLf
       ShowHTML "   window.open('Aluno.asp?EW=118&EF=" & sstrEF & "&EA=sq_aluno=" & w_chave & "&CL=" & CL & "&IN=10" & "', 'aluno', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=780,height=500,left=10,top=10');" & VbCrLf
       ShowHTML "   history.go(-1);" & VbCrLf
       ShowHTML "</SCRIPT>" & VbCrLf
    Else
       ' Grava o acesso na tabela de log
       SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql) " & VbCrLf & _
             "values ( " & VbCrLf & _
             "         " & CL & ", " & VbCrLf & _
             "         getdate(), " & VbCrLf & _
             "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
             "         0, " & VbCrLf & _
             "         'Acesso à tela de atualização da escola.', " & VbCrLf & _
             "         null " & VbCrLf & _
             "       ) " & VbCrLf
       AbreSessao
       ExecutaSQL(SQL)
       ShowHTML "<SCRIPT LANGUAGE=""JAVASCRIPT"">" & VbCrLf
       ShowHTML "   window.open('/Manut.asp?CL=" & sstrEF & "&w_in=" & sstrIN & "', 'cliente', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=780,height=580,left=10,top=10');" & VbCrLf
       ShowHTML "   history.go(-1);" & VbCrLf
       ShowHTML "</SCRIPT>" & VbCrLf
    End If
    FechaSessao

  End If

  ShowHTML "</tr></center>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Verificação de Senha
REM =========================================================================

REM =========================================================================
REM Monta a tela de ofertas da escola
REM -------------------------------------------------------------------------
Public Sub ShowOferta

Dim sql2, biblioteca, linguas
Dim oferta(50), w_modalidade, w_serie, w_turno, w_mod_atual, w_ser_atual, w_tur_atual
Dim w_texto, w_numero, w_pos, i, w_cont, w_temp

    w_cont = 0
    
    sql2 = "SELECT f.ds_especialidade " & VbCrLf & _ 
           "  from escEspecialidade_cliente AS e " & VbCrLf & _
           "       INNER JOIN escEspecialidade AS f ON (e.sq_codigo_espec = f.sq_especialidade and " & VbCrLf & _
           "                                            'J'               = f.tp_especialidade " & VbCrLf & _
           "                                           ) " & VbCrLf & _
           " WHERE e.sq_cliente = " & CL & " " & VbCrLf & _
           "ORDER BY ds_especialidade "
    RS1.Open sql2, sobjConn, adOpenForwardOnly
    
    ' Recupera a oferta a partir das turmas da escola
    sql  = "select distinct a.sq_site_cliente, b.nome ds_serie, b.curso ds_modalidade, case rtrim(ltrim(a.ds_turno)) when 'M' then 1 when 'V' then 2 else 3 end ds_turno " & VbCrLf & _ 
           "  from escTurma a inner join escTurma_Modalidade b on (upper(rtrim(ltrim(a.ds_serie)))=upper(rtrim(ltrim(b.serie))))" & VbCrLf & _
           " where sq_site_cliente = " & CL & " " & VbCrLf & _
           "order by 3,2 "
    RS.Open sql, sobjConn, adOpenForwardOnly

    If (Not RS.EOF) or (Not RS1.EOF) Then
        ShowHTML("<p><b>Etapas / Modalidades de ensino oferecidas:</b>")
        ShowHTML("<dl>")
        
        If Not RS1.EOF Then
           While Not RS1.EOF
             ShowHTML("<dt>" & RS1("ds_especialidade") & "</dt>")
             RS1.MoveNext
           Wend
           RS1.Close
        End If

        If Not RS.EOF Then
            w_cont = 1
            While Not RS.EOF
              oferta(w_cont) = RS("ds_modalidade") & "}" & RS("ds_serie") & "|" & exibeTurno(RS("ds_turno"))
              w_cont = w_cont + 1
              RS.MoveNext
            Wend
            Sort oferta

            w_mod_atual = ""
            w_ser_atual = ""
            w_tur_atual = ""
            For Each i In oferta
              If i > "" Then
                w_modalidade = mid(i,1,instr(i,"}")-1)
                w_serie      =mid(i,instr(i,"}")+1,(instr(i,"|")-instr(i,"}")-1))
                w_turno      = mid(i,instr(i,"|")+1,10)
                if w_modalidade <> w_mod_atual then
                   ShowHTML("<dt><li>" & w_modalidade) 'exibeModal(w_modalidade))
                   w_mod_atual = w_modalidade
                   w_ser_atual = ""
                   w_tur_atual = ""
                end if
                if w_serie <> w_ser_atual then
                   if w_serie = "0" then
                      ShowHTML("<dd>")
                   else
                      ShowHTML("<dd>" & w_serie & ": ") 'exibeSerie(w_modalidade,w_serie))
                   end if
                   w_ser_atual = w_serie
                   w_tur_atual = ""
                   w_cont      = 0
                end if
                if w_turno <> w_tur_atual then
                   if w_cont > 0 then
                      Response.Write ", " & w_turno 'exibeTurno(w_turno)
                   else
                      Response.Write " " & w_turno 'exibeTurno(w_turno)
                   end if
                   w_cont = w_cont + 1
                end if
              End If
            Next
        End If
        ShowHTML("</dl>")
    End If
    RS1.Close
    
    sql2 = "SELECT f.ds_especialidade " & VbCrLf & _ 
           "  from escEspecialidade_cliente AS e " & VbCrLf & _
           "       INNER JOIN escEspecialidade AS f ON (e.sq_codigo_espec  = f.sq_especialidade and " & VbCrLf & _
           "                                            f.tp_especialidade not in ('M','J') " & VbCrLf & _
           "                                           ) " & VbCrLf & _
           " WHERE e.sq_cliente = " & CL & " " & VbCrLf & _
           "ORDER BY ds_especialidade "
    RS1.Open sql2, sobjConn, adOpenForwardOnly
    
    If Not RS1.EOF Then
       ShowHTML("<p><b>Em Regime de Intercomplementaridade: </b></p><ul>")
       While Not RS1.EOF
         ShowHTML("<li>" & RS1("ds_especialidade") & "</li>")
         RS1.MoveNext
       Wend
       ShowHTML("</ul>")
       RS1.Close
    End If
    
End Sub

REM =========================================================================
REM Monta a tela de ofertas da escola
REM -------------------------------------------------------------------------
Public Sub ShowOferta1

Dim sql2, biblioteca, linguas

    sql2 = "SELECT f.ds_especialidade " & VbCrLf & _ 
    "  from escEspecialidade_cliente AS e " & VbCrLf & _
    "       INNER JOIN escEspecialidade AS f ON (e.sq_codigo_espec = f.sq_especialidade) " & VbCrLf & _
    " WHERE e.sq_cliente = " & CL & " " & VbCrLf & _
    "UNION " & VbCrLf & _
    "SELECT f.ds_especialidade " & VbCrLf & _ 
    "  from escEspecialidade f " & VbCrLf & _
    " WHERE upper(ds_especialidade) like '%PROFISSIONAL%' " & VbCrLf & _
    "   and 0 < (select count(sq_cliente) from escParticular_curso where sq_cliente = " & CL & ") " & VbCrLf & _
    "ORDER BY ds_especialidade "
    RS1.Open sql2, sobjConn, adOpenForwardOnly
    w_cont = 0
    
    
    
    ShowHTML("<p><b>Etapas / Modalidades de ensino oferecidas:</b></p>")
    ShowHTML("<ul>")
    While Not RS1.EOF
            If uCase(RS1("ds_especialidade")) = "BIBLIOTECA" Then
                biblioteca = RS1("ds_especialidade")
            End If
            If uCase(RS1("ds_especialidade")) = "LÍNGUAS" Then
                linguas =  RS1("ds_especialidade")
            End If
            If uCase(RS1("ds_especialidade")) <> "BIBLIOTECA" AND uCase(RS1("ds_especialidade")) <> "LÍNGUAS" Then
                ShowHTML "<li>" & RS1("ds_especialidade") & "</li>"
                w_cont = 1            
            End If
    RS1.MoveNext
    Wend
    ShowHTML("</ul>")
    if(biblioteca <> "" OR linguas <> "") Then
        ShowHTML("<p><b>Em Regime de Intercomplementaridade: </b></p>")
        ShowHTML("<ul>")
        If(biblioteca <> "") Then
        ShowHTML("<li>" & biblioteca & "</li>")
        End If
        If(linguas <> "") Then
        ShowHTML("<li>" & linguas & "</li>")
        ShowHTML("</ul>")
        End If
    End If
    RS1.Close
    
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
    Case conWhatValida      ShowValida
    Case "oferta"           ShowOferta
  End Select
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>