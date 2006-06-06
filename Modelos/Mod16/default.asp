<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../../Constants_ADO.inc"-->
<!--#INCLUDE FILE="../Constants.inc"-->
<!--#INCLUDE FILE="../../esc.inc"-->
<!--#INCLUDE FILE="../../Funcoes.asp"-->
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

Private RS, sql
Private RSCabecalho
Private RSRodape

Private dbms, sobjConn

Private sstrSN
Private sstrCL

sstrSN = "Default.Asp"
sstrCL = "Aluno.Asp"

Public sstrEF
Public sstrEW
Public sstrIN
Public CL

CL = Request.QueryString("CL")
sstrEF = "sq_cliente=" & CL
sstrEW = Request.QueryString("EW")
sstrIN = Request.QueryString("IN")
w_dir  = "Modelos/Mod16/"

Private sstrData

Set RS = Server.CreateObject("ADODB.RecordSet")
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
ShowHTML "   <link href=""../../css/estilo.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <link href=""../../css/print.css""  media=""print""  rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
ShowHTML "</head>"

ShowHTML "<BASE HREF=""http://" & Request.ServerVariables("server_name") & conVirtualPath & """>"
ShowHTML "<body>"
ShowHTML "<div id=""container"">"
ShowHTML "  <div id=""cab"">"
ShowHTML "    <div id=""cabtopo"">"
ShowHTML "      <div id=""logoesq""><img src=""img/fundo_logoesq.gif"" border=0></div>"
ShowHTML "      <div id=""logodir""><a href=""http://www.se.df.gov.br""><img src=""img/fundo_logodir.jpg"" border=0></a></div>"
ShowHTML "    </div>"

sql = "SELECT a.sq_cliente, a.ds_cliente, b.ds_mensagem FROM escCliente a inner join esccliente_site b on (a.sq_cliente = b.sq_cliente)WHERE a." & sstrEF
RS.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly
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
ShowHTML "      <li><a href=""" & w_dir & "Default.Asp?EW=116&EF=" & CL & "&CL=" & replace(CL,"sq_cliente=","") & "&IN=1"" >Projeto</a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & """ id=""link2"">Notícias</a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=115&CL=" & replace(CL,"sq_cliente=","") & """ id=""link5"">Calendário</a>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & """ >Arquivos (<i>download</i>)</a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=116&IN=0&CL=" & replace(CL,"sq_cliente=","") & """ >Alunos</a></li>"
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
        ShowHTML "      <li><b>Projeto</b></li>"
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
        "ORDER BY origem, nr_ordem, dt_arquivo desc, in_destinatario " & VbCrLf 
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
        ShowHTML "    <TD><a href=""http://" & replace(RS("diretorio"),"http://","") & "/" & RS("ln_arquivo") & """ target=""_blank"">" & RS("ds_titulo") & "</a><br><div align=""justify""><font size=1>.:. " & RS("ds_arquivo") & "</div>"
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
        "                     INNER JOIN escCliente_Dados as d on (a.sq_cliente = d.sq_cliente) " & _
        "                     LEFT  JOIN escCliente_Foto  as e on (a.sq_cliente = e.sq_cliente and e.tp_foto = 'P') " & _
        "WHERE a." & sstrEF & " " & VbCrLf & _
        "ORDER BY e.nr_ordem"

  RS.Open sql, sobjConn, adOpenForwardOnly

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
          ShowHTML "        <img class=""foto"" src=""http://" & Replace(lCase(RS("ds_diretorio")),lCase("http://"),"") & "/" & RS("ln_foto") & """ width=""302"" height=""206""><br>" & RS("ds_foto") & "<br>"
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
       ShowHTML "    <td align=""right""><font size=""1""><b>Fotografias</b><font size=1></b><br>"
       Do While Not RS.EOF
          ShowHTML "     <a href=""http://" & Replace(lCase(RS("ds_diretorio")),"http://","") & "/" & RS("ln_foto") & """ target=""_blank"" title=""Clique sobre a imagem para ampliar""><img align=""top"" class=""foto"" src=""http://" & Replace(lCase(RS("ds_diretorio")),"http://","") & "/" & RS("ln_foto") & """ width=""201"" height=""153""><br>" & RS("ds_foto")& "</a><br>"
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

  If sstrIN = "" then
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
        "     INNER JOIN escCliente AS d ON (c.sq_cliente = d.sq_cliente_pai) " & VbCrLf & _
        "WHERE d.sq_cliente = " & CL & VbCrLf & _
        "  AND YEAR(dt_ocorrencia)= " & wAno & " " & VbCrLf & _
        "ORDER BY data, origem desc, ocorrencia" & VbCrLf 

  RS.Open sql, sobjConn, adOpenForwardOnly
  
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
     RS.MoveFirst
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
        ShowHTML "    <iframe src=""http://" & replace(RS("ds_diretorio"),"http://","") & "/" & RS("ln_prop_pedagogica") & """ width=""100%"" height=""510"">"
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

  If sstrIN = "0" Then
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
    If sstrIN = "0" Then
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
       ShowHTML "   window.open('aluno.asp?EW=118&EF=" & sstrEF & "&EA=sq_aluno=" & w_chave & "&CL=" & CL & "&IN=10" & "', 'aluno', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=780,height=500,left=10,top=10');" & VbCrLf
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
       ShowHTML "   window.open('../../Manut.asp?CL=" & sstrEF & "&w_in=" & sstrIN & "', 'cliente', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=780,height=580,left=10,top=10');" & VbCrLf
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
  End Select
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>
