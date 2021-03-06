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
REM Nome     : Alexandre Vinhadelli Papad�polis 
REM Descricao: Exibi��o dos dados da escola
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 16/03/2000 14:10PM
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Bras�lia - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Inform�tica
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
ShowHTML "   <script language=""javascript"" src=""/inc/scripts.js""> </script>"
ShowHTML "   <script language=""javascript"" src=""/inc/jquery-1.3.2.js""> </script>"
ShowHTML "   <script language=""javascript"" src=""/inc/jquery.hoverIntent.minified.js""> </script>"
ShowHTML "   <script language=""javascript"" src=""/inc/jquery.bgiframe.min.js""> </script>"
ShowHTML "<!--[if IE]><script src=""/inc/excanvas.js"" type=""text/javascript"" charset=""utf-8""></script><![endif]-->"
ShowHTML "   <script language=""javascript"" src=""/inc/jquery.bt.js""> </script>"
ShowHTML "   <script language=""javascript"" src=""/inc/jquery.treeview.js""> </script>"

ShowHTML "</head>"

ShowHTML "<BASE HREF=""" & conSite & "/"">"
If (Request("H") > "") or (Request("Q") > "") or (Request("p_regional") > "") or (wIN > 0) Then
   ShowHTML "<body onLoad=""location.href='#resultado'"">"
Else
   ShowHTML "<body>"
End If
ShowHTML "    <div id=""pagina"">"
ShowHTML "    <div id=""topo""></div>"
%>
<script>
$(document).ready(function(){
  $('a[title]').bt({
  fill: '#F7F7F7', 
  strokeStyle: '#B7B7B7', 
  spikeLength: 10, 
  spikeGirth: 10, 
  padding: 8, 
  cornerRadius: 0, 
  cssStyles: {
    fontFamily: '"lucida grande",tahoma,verdana,arial,sans-serif', 
    fontSize: '13px',
    width: 'auto'
  }
});
$("#arvore").treeview();
});
</script>
<div id="busca">
    <div class="data"> <% Response.Write(ExibeData(date())) %> </div>
    <div class="clear"></div>
  </div>
  <div id="menu">
    <div id="menuTop">
      <div class="esquerda">
        <ul>
          <li><a href="http://www.se.df.gov.br/300/30001001.asp">Secretaria de Educa��o</a></li>
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

sql = "SELECT a.sq_cliente, a.ds_cliente, b.ds_mensagem FROM escCliente a inner join esccliente_site b on (a.sq_cliente = b.sq_cliente)WHERE a." & sstrEF
ConectaBD SQL
ShowHTML "      <ul>"
ShowHTML "      <script language=""JavaScript"" src=""inc/mm_menu.js"" type=""text/JavaScript""></script>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=110&CL=" & replace(CL,"sq_cliente=","") & """ >Inicial</a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=113&CL=" & replace(CL,"sq_cliente=","") & """ id=""link1"">Quem somos</a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=117&CL=" & replace(CL,"sq_cliente=","") & """ >Fale conosco </a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.Asp?EW=composicao&EF=" & CL & "&CL=" & replace(CL,"sq_cliente=","") & "&IN=1"" >Composi�ao administrativa</a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & """ id=""link2"">Not�cias</a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=115&CL=" & replace(CL,"sq_cliente=","") & """ id=""link5"">Calend�rio</a>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & """ >Arquivos (<i>download</i>)</a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=116&IN=0&CL=" & replace(CL,"sq_cliente=","") & """ >Escolas</a></li>"
ShowHTML "	    </ul>"
ShowHTML "	    <div class=""clear""></div>"
ShowHTML "    </div>"
ShowHTML "	  <div id=""conteudo"">"
ShowHTML "      <h2>" & RS("ds_cliente") & "</h2>"
Select Case sstrEW
  Case conWhatPrincipal   ShowHTML "      <h3>Inicial</h3>"
  Case conWhatManut       ShowHTML "      <h3>Alunos</h3>"
  Case conWhatQuem        ShowHTML "      <h3>Quem somos</h3>"
  Case conWhatExFale      ShowHTML "      <h3>Fale conosco</h3>"
  Case conWhatArquivo     ShowHTML "      <h3>Arquivos</h3>"
  Case conWhatExNoticia   ShowHTML "      <h3>Not�cias</h3>"
  Case conWhatExCalend    ShowHTML "      <h3>Calend�rio</h3>"
  Case conWhatValida
End Select
RS.Close

Main
ShowHTML "    <div class=""clear""/>"
ShowHTML "    </div>"

ShowHTML "    </div>"
%>
<div id="rodape">
  <!-- Menu Rodape -->
  <ul>
    <li><a href="http://www.se.df.gov.br">P�gina Inicial</a></li>
    <li><a href="http://www.se.df.gov.br/300/30001005.asp">Fale conosco</a></li>
    <li class="ultimo"><a href="http://www.se.df.gov.br/300/30001009.asp">Mapa do site</a></li>
  </ul>
  <!-- Fim Menu Rodape -->
  <div class="clear"/>
  <!-- Endere�o -->
  <div class="endereco">
  <div class="buriti">
  <p>Anexo do Pal�cio do Buriti - 9� andar  - Bras�lia </p>
  <p>Telefone (61) 3224 0016 (61) 3225 1266 | Fax (61) 3901 3171</p>
  </div>
  <div class="buritinga">
  <p>QNG AE Lote 22 bl 05 sala 03 - Taguatinga   Norte</p>
  <p>Telefone (61) 3355 86 30 | Fax 3355 86 94</p>
  </div>
  </div>
  
  <p class="copy">Copyright � 2000/2008 - SE/GDF - Todos os Direitos Reservados</p>
  <!-- Fim Endere�o -->
</div>
<%
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

  sql = "SELECT case a.pasta when '1' then (a.pasta*1000000)+(12-month(a.dt_arquivo)*10000)+a.nr_ordem else (a.pasta*1000000)+(12-month(a.dt_arquivo)*10000)+a.nr_ordem end as ordena, case a.in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end AS in_destinatario, " & VbCrLf & _
        "       dt_arquivo, ds_titulo, nr_ordem, a.pasta, case a.pasta when '1' then 'Meses' when '2' then 'Formul�rios' when '3' then 'Diversos' else '' end as nmpasta, ds_arquivo, ln_arquivo, 'Regional' AS Origem, x.ln_internet diretorio " & VbCrLf & _
        "From escCliente_Arquivo a, escCliente x " & VbCrLf & _
        "WHERE in_ativo = 'Sim'" & VbCrLf & _
        "  AND x." & sstrEF & " " & VbCrLf & _
        "  AND a." & replace(sstrEF,"sq_cliente","sq_site_cliente") & " " & VbCrLf & _
        "  and in_destinatario <> 'E'" & VbCrLf & _
        "  and YEAR(a.dt_arquivo) = " & wAno & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT case a.pasta when '1' then (a.pasta*1000000)+(12-month(a.dt_arquivo)*10000)+a.nr_ordem else (a.pasta*1000000)+(12-month(a.dt_arquivo)*10000)+a.nr_ordem end as ordena, case a.in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end AS in_destinatario, " & VbCrLf & _
        "       a.dt_arquivo, a.ds_titulo, a.nr_ordem, a.pasta, case a.pasta when '1' then 'Meses' when '2' then 'Formul�rios' when '3' then 'Diversos' else '' end as nmpasta, ds_arquivo, ln_arquivo, 'SEDF' AS Origem, d.ds_diretorio diretorio " & VbCrLf & _
        "From escCliente_Arquivo AS a INNER JOIN escCliente AS c ON (a.sq_site_cliente = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente AS e ON (e.sq_cliente_pai  = c.sq_cliente) " & VbCrLf & _
        "                             INNER JOIN escCliente_Site AS d ON (c.sq_cliente = d.sq_cliente) " & VbCrLf & _
        "WHERE a.in_ativo = 'Sim'" & VbCrLf & _
        "  AND e." & sstrEF & " " & VbCrLf & _
        "  and in_destinatario <> 'E'" & VbCrLf & _        
        "  and YEAR(a.dt_arquivo) = " & wAno & VbCrLf & _
        " ORDER BY 1 asc " & VbCrLf 
  RS.Open sql, dbms, adOpenForwardOnly
  'response.write sql
  wCont = 0
  Dim meses
  If Not RS.EOF Then
    wCont = 1
    ShowHTML "<ul id=""arvore"" class=""filetree"">"
    pasta = 0    
    meses = 0
    formularios = 0
    diversos = 0
    Do While Not RS.EOF
      for i = 1 to 3 step 1
        If(cInt(RS("pasta")) <> cInt(pasta) AND cInt(RS("pasta"))=1)Then
          meses = meses + 1
        elseIf(cInt(RS("pasta")) <> cInt(pasta) AND cInt(RS("pasta"))=2)Then
          formularios = formularios + 1
        elseIf(cInt(RS("pasta")) <> cInt(pasta) AND cInt(RS("pasta"))=3)Then
          diversos = diversos + 1
        End If
        'response.write "1: "&meses&"<br/>"
        'response.write "2: "&formularios&"<br/>"
        'response.write "3: "&diversos&"<br/>"
        If(cInt(RS("pasta")) <> cInt(pasta) AND cInt(RS("pasta"))=1)Then
          ShowHTML "<li><span class=""folder"">"&RS("nmpasta")&"</span><ul>"
        ElseIf(cInt(RS("pasta")) <> cInt(pasta) AND cInt(RS("pasta")) = 2)Then
          If(meses > 0 OR diversos > 0) Then
            ShowHTML "</ul></li></ul></li><li class=""closed""><span class=""folder"">"&RS("nmpasta")&"</span><ul>"
          else
            ShowHTML "<li><span class=""folder"">"&RS("nmpasta")&"</span><ul>"
          End If          
        ElseIf(cInt(RS("pasta")) <> cInt(pasta) AND cInt(RS("pasta")) = 3)Then
          If(meses > 0 OR formularios > 0) Then
            ShowHTML "</ul></li></li><li class=""closed""><span class=""folder"">"&RS("nmpasta")&"</span><ul>"
          else
            ShowHTML "<li><span class=""folder"">"&RS("nmpasta")&"</span><ul>"
          End If          
        End If
            If(cInt(RS("pasta")) = 1 )Then
              If(cInt(month(RS("dt_arquivo")))<>mes) Then
                If(cInt(RS("pasta")) = cInt(pasta))Then
                  ShowHTML "</ul></li>"
                End If                
                ShowHTML " <li class=""closed""><span class=""folder"">"
                ShowHTML exibeMes(RS("dt_arquivo"))&"/"&year(RS("dt_arquivo"))
                ShowHTML "</span><ul>"
              End If                
            End If
          If(cInt(RS("pasta")) = i)Then
            If (instr(1,RS("ln_arquivo"),".pdf",1) > 0)Then
              icone = "pdf.gif"
            ElseIf (instr(1,RS("ln_arquivo"),".xls",1) > 0)Then
              icone = "xls.gif"
            ElseIf (instr(1,RS("ln_arquivo"),".txt",1) > 0)Then
              icone = "txt.gif"
            ElseIf (instr(1,RS("ln_arquivo"),".doc",1) > 0)Then
              icone = "doc.gif"
            ElseIf (instr(1,RS("ln_arquivo"),".ppt",1) > 0)Then
              icone = "ppt.gif"
            ElseIf (instr(1,RS("ln_arquivo"),".jpeg",1) > 0 OR instr(1,RS("ln_arquivo"),".jpg",1) > 0 OR instr(1,RS("ln_arquivo"),".gif",1) > 0 OR instr(1,RS("ln_arquivo"),".png",1) > 0)Then
              icone = "picture.gif"
            Else
              icone = "undefined.gif"
            End If
              ShowHTML "<LI><span class=""file"" style=""background: url(../img/"&icone&") 0 0 no-repeat;"">&nbsp;<a href=""" & RS("diretorio") & "/" & RS("ln_arquivo") & """ title="""&Nvl(RS("ds_arquivo"),"Descri��o n�o informada.")&""" target=""_blank"">"&RS("dt_arquivo")&" � "&RS("ds_titulo")&"&nbsp;("&RS("origem")&")</a></span></LI>"
          End If
        pasta = RS("pasta")
        mes   = cInt(Month(RS("dt_arquivo")))
      Next
    RS.MoveNext
    Loop
    ShowHTML "  </UL>"
    ShowHTML "  </UL>"
  Else
   ShowHTML "  <TR><TD COLSPAN=4 ALIGN=CENTER>N�o h� arquivos dispon�veis no momento para o ano de " & wAno & " </TR>"
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
      ShowHTML "  </UL>"
      ShowHTML "  <TR><br /><TD COLSPAN=4 ><b>Arquivos de outros anos</b><br>"
      While Not RS.EOF
        ShowHTML "     <li><a href=""" & w_dir & "Default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & "&wAno=" & RS("ano") & """ >Arquivos de " & RS("ano") & "</a></li>"
        RS.MoveNext
      Wend
      ShowHTML "  </TD></TR>"
    End If
    RS.Close
  ShowHTML "</table></center>"
End Sub
REM -------------------------------------------------------------------------
REM Final da P�gina de Arquivos
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
  
  'Response.Write sql
  'Response.End 

  If Not RS.EOF Then
    If RS("ds_texto_abertura") > "" Then
       ShowHTML "        <p class=""chamada"">" & replace(RS("ds_texto_abertura"),chr(13)&chr(10), "<br/>")
    Else
       ShowHTML "        <p class=""chamada"">"
    End If
    If RS("sq_cliente_foto") > "" Then
         ShowHTML "    <ul class=""fotos"">"   
         Do While Not RS.EOF
            ShowHTML "    <li><a href=""" & RS("ds_diretorio") & "/" & RS("ln_foto") & """ target=""_blank"" title=""Clique sobre a imagem para ampliar""><img align=""top"" src=""" & RS("ds_diretorio") & "/" & RS("ln_foto") & """ width=""302"" height=""206"">" & RS("ds_foto")& "</a></li>"
            RS.MoveNext
         Loop
         ShowHTML "         </ul>"
    End If    
  End If

  RS.Close
  
End Sub
REM -------------------------------------------------------------------------
REM Final da P�gina Principal
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

  If Not RS.EOF Then
    If RS("ds_institucional") > "" Then
       ShowHTML "        <p class=""chamada"">" & replace(RS("ds_institucional"),chr(13)&chr(10), "<br/>") & "</p>"
    Else
       ShowHTML "        <P class=""chamada"">"
    End If
    If RS("sq_cliente_foto") > "" Then
       ShowHTML "    <ul class=""fotos"">"   
       Do While Not RS.EOF
          ShowHTML "    <li><a href=""" & RS("ds_diretorio") & "/" & RS("ln_foto") & """ target=""_blank"" title=""Clique sobre a imagem para ampliar""><img align=""top"" class=""foto"" src=""" & RS("ds_diretorio") & "/" & RS("ln_foto") & """ width=""302"" height=""206""><br>" & RS("ds_foto")& "</a></li>"
          RS.MoveNext
       Loop
       ShowHTML "    </ul>"
    End If
    
  End If
  RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da P�gina Quem Somos
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
  
  'Valida��o dos campos do formul�rio de envio de email
  'ScriptOpen "JavaScript"
  'ValidateOpen "Validacao"
  'Validate "w_nome", "Nome", "1", "1", "3", "60", "1", "1"
  'Validate "w_email", "e-Mail", "1", "1", "3", "60", "1", "1"
  'Validate "w_endereco", "Endere�o", "1", "1", "3", "60" , "1", "1"
  'Validate "w_telefone", "Telefone", "1", "1", "7", "30" , "1", "1"
  'Validate "w_assunto", "Assunto", "1", "1", "5", "80" , "1", "1"
  'Validate "w_mensagem", "Mensagem", "1", "1", "5", "1000" , "1", "1"
  'ShowHTML "  theForm.Botao.disabled=true;"
  'ValidateClose
  'ScriptClose
  
  
  ShowHTML "<tr><td>"

  ShowHTML "<p align=""justify""><li>Informa��es, sugest�es e reclama��es podem ser feitas utilizando os dados abaixo:"

  If Not RS.EOF Then

  ShowHTML "<table border=""0"" cellspacing=""3"">"
  ShowHTML "  <tr>"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td width=""20%"" align=""right""><b>Diretor(a):"
  ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_diretor") 
  ShowHTML "  </tr>"
  'ShowHTML "  <tr>"
  'ShowHTML "    <td width=""10%"">"
  'ShowHTML "    <td width=""20%"" align=""right""><b>Secret�rio(a):"
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
  ShowHTML "    <td width=""20%"" align=""right""><b>Endere�o:"
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
  ShowHTML "    <td width=""20%"" align=""right""><b>Munic�pio:"
  ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_municipio") & "-" & RS("sg_uf") & "&nbsp;&nbsp;<b>CEP:</b> " & RS("nr_cep")
  ShowHTML "  </tr>"
  ShowHTML "</table>"
  
  'If Nvl(RS("ds_email_internet"),"") > "" Then
  '   ShowHTML "<FORM action=""" & w_dir & sstrSN & "?EW=PreparaMail"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"">"
  '   ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
  '   ShowHTML "<p align=""justify""><li>Se desejar, envie uma mensagem de e-mail usando o formul�rio abaixo:"
  '   ShowHTML "<table border=""1"" bgcolor=""#DAEABD"" style=""border: 1px solid rgb(0,0,0);""><tr><td>"
  '   ShowHTML "  <table border=""0"" cellspacing=""3"">"
  '   ShowHTML "    <tr><td><font size=""1""><b>Nome: </b><br><input type=""text"" name=""w_nome"" class=""STI"" size=""60"" maxlength=""60"" value="""" title=""Informe o nome do remetente"">"
  '   ShowHTML "    <tr><td><font size=""1""><b>e-Mail: </b><br><input type=""text"" name=""w_email"" class=""STI"" size=""60"" maxlength=""60"" value="""" title=""Informe o email do remetente"">"
  '   ShowHTML "    <tr><td><font size=""1""><b>Endere�o: </b><br><input type=""text"" name=""w_endereco"" class=""STI"" size=""60"" maxlength=""60"" value="""" title=""Informe o endereco do remetente"">"
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
REM Final da P�gina Fale Conosco
REM =========================================================================

REM =========================================================================
REM Monta a tela Not�cias
REM -------------------------------------------------------------------------
Public Sub ShowNoticia

  Dim sql, strNome, wCont, wRede, wAno
  
  wAno =  Request.QueryString("wAno")
  
  If wAno = "" Then
     wAno = Year(Date())
  End If

  If sstrIN = "" then
     sql = "SELECT a.ln_externo, a.sq_noticia, a.dt_noticia AS data, a.ds_titulo AS ocorrencia, 'G' AS origem " & VbCrLf & _
           "  FROM escNoticia_Cliente  a" & VbCrLf & _
           " WHERE sq_site_cliente   = " & CL & VbCrLf & _
           "  and a.in_ativo         = 'Sim'" & VbCrLf & _
           "  and a.in_exibe         = 'Sim'" & VbCrLf & _
           "  and YEAR(a.dt_noticia) = " & wAno & VbCrLf & _
           "UNION " & VbCrLf & _
           "SELECT a.ln_externo, a.sq_noticia, dt_noticia AS data, ds_titulo AS ocorrencia, 'R' AS origem " & VbCrLf & _
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
     ShowHTML "    <TD width=""15%""><b>Ocorr�ncia"
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
            If RS("ln_externo") > "" Then
              ShowHTML "  <TD colspan=2><a target=""_blank"" href="""&RS("ln_externo")&"""> " & RS("ocorrencia") & "</a>&nbsp;&nbsp;&nbsp;<a href=""" & w_dir & sstrSN & "?EW=" & conWhatExNoticia & "&CL=" & CL & "&EF=" & sstrEF & "&IN=" & RS("sq_noticia") & """><b>[Ler]</a>"      
            Else
              ShowHTML "  <TD colspan=2>" & RS("ocorrencia") & "&nbsp;&nbsp;&nbsp;<a href=""" & w_dir & sstrSN & "?EW=" & conWhatExNoticia & "&CL=" & CL & "&EF=" & sstrEF & "&IN=" & RS("sq_noticia") & """><b>[Ler]</a>"
            End If
          Else
            ShowHTML "    <TD align=""center"">" & Mid(100+Day(RS("data")),2,2) & "/" & Mid(100+Month(RS("data")),2,2) & "/" &Year(RS("data"))
            ShowHTML "    <TD align=""center"">Regional"
            If RS("ln_externo") > "" Then
              ShowHTML "  <TD colspan=2><a target=""_blank"" href="""&RS("ln_externo")&"""> " & RS("ocorrencia") & "</a>&nbsp;&nbsp;&nbsp;<a href=""" & w_dir & sstrSN & "?EW=" & conWhatExNoticia & "&CL=" & CL & "&EF=" & sstrEF & "&IN=" & RS("sq_noticia") & """><b>[Ler]</a>"      
            Else
              ShowHTML "  <TD colspan=2>" & RS("ocorrencia") & "&nbsp;&nbsp;&nbsp;<a href=""" & w_dir & sstrSN & "?EW=" & conWhatExNoticia & "&CL=" & CL & "&EF=" & sstrEF & "&IN=" & RS("sq_noticia") & """><b>[Ler]</a>"
            End If
          End If
          ShowHTML "  </TR>"
          wCont = wCont + 1
          RS.MoveNext
        Loop
        
     Else
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD colspan=""5""><BR>Nenhuma not�cia foi encontrada para o ano " & wAno & "</TD>"
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
        ShowHTML "  <TR valign=""top""><TD colspan=""5""><br><b>Not�cias de outros anos</b><br>"     
        While Not RS.EOF
           ShowHTML "    <li><a href=""" & w_dir & "Default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & "&wAno=" & RS("ano") & """ id=""link2"">Not�cias de " & RS("ano") & "</a> </li>"
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
REM Final da P�gina Not�cias
REM =========================================================================

REM =========================================================================
REM Monta a tela Calend�rio
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
        
  RS1.Open sql, dbms, adOpenForwardOnly
 
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
     RS.Open sql, dbms, adOpenForwardOnly

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
        ShowHTML "  <td>Sem informa��o"
     End If
     
     RS.Close
     'Feriados
     sql = "SELECT '' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'B' AS origem FROM escCalendario_base a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE YEAR(dt_ocorrencia)=" & wAno & " AND b.sigla <> 'CN' " & VbCrLf & _
           "ORDER BY data, origem desc, ocorrencia" & VbCrLf
     RS.Open sql, dbms, adOpenForwardOnly
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
        ShowHTML "  <td>Sem informa��o"
     End If
     RS.Close
     
     'Recessos
     sql = "SELECT '#99CCFF' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla in ('RE','RA')) WHERE sq_site_cliente = " & CL & "  AND YEAR(dt_ocorrencia)= " & Year(Date()) & " " & VbCrLf & _
           "UNION " & VbCrLf & _
           "SELECT '#FFFF99' as cor, e.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'R' AS origem " & VbCrLf & _
           "FROM escCalendario_Cliente   AS a" & VbCrLf & _
           "     INNER JOIN escCliente   AS b ON (a.sq_site_cliente = b.sq_cliente) " & VbCrLf & _
           "     INNER JOIN escCliente   AS c ON (b.sq_cliente      = c.sq_cliente_pai) " & VbCrLf & _
           "     INNER JOIN escCliente   AS d ON (c.sq_cliente      = d.sq_cliente_pai) " & VbCrLf & _
           "     INNER JOIN escTipo_Data AS e ON (a.sq_tipo_data    = e.sq_tipo_data and e.sigla in ('RE','RA')) " & VbCrLf & _
           "WHERE d.sq_cliente = " & CL & VbCrLf & _
           "  AND YEAR(dt_ocorrencia)= " & wAno & " " & VbCrLf & _
           "ORDER BY data, origem desc, ocorrencia" & VbCrLf
            
     RS.Open sql, dbms, adOpenForwardOnly
  
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
        ShowHTML "  <td>Sem informa��o"
     End If
     RS.Close
     ShowHTML "  <td valign=""top""><TABLE align=""center"" border=0 cellSpacing=1>"
     
      ' Recupera o ano letivo e o per�odo de recesso
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
     RS.Open sql, dbms, adOpenForwardOnly
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
      
        ShowHTML "  <TR valign=""top"" ALIGN=""CENTER""><TD COLSPAN=""4""><b>" & w_cont & "� Semestre</b></td></tr>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD><b>M�S"
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
              RS.Open sql, dbms, adOpenForwardOnly
              w_final = RS("fim")
              RS.Close 
           End If
           
           sql = "SELECT dbo.diasLetivos('" & w_inicial & "', '" & w_final & "'," & CL & ",null) qtd" & VbCrLf 
           RS.Open sql, dbms, adOpenForwardOnly
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
     ShowHTML "    <TD width=""68%""><b>Ocorr�ncia"
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
     ShowHTML "    <TD><b>M�s"
     ShowHTML "    <TD NOWRAP><b>Dias letivos"
     ShowHTML "  </TR>"
     ShowHTML "  <TR>"
     ShowHTML "    <TD COLSPAN=""2"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
     ShowHTML "  </TR>"
     For wCont = 1 to 12
        sql = "SELECT dbo.diasLetivos('01/" & mid(100+wCont,2,2) & "/'+cast(year(getDate()) as varchar),convert(varchar,dbo.lastDay(convert(datetime,'01/" & mid(100+wCont,2,2) & "/'+cast(year(getDate()) as varchar),103)),103)," & CL & ") qtd" & VbCrLf 
        RS1.Open sql, dbms, adOpenForwardOnly
        
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
REM Final da P�gina Calend�rio
REM =========================================================================

REM =========================================================================
REM Monta a tela de escolas
REM -------------------------------------------------------------------------
Public Sub ShowEscolas

  Dim sql2, wCont, sql1, wAtual, wIN
  
   ShowHTML "        <FORM ACTION=""" & w_dir & "default.asp"" id=form1 name=form1 METHOD=""POST"">"
   ShowHTML "        <input type=""Hidden"" name=""EW"" value=""" & sstrEW & """>"
   ShowHTML "        <input type=""Hidden"" name=""IN"" value=""" & sstrIN & """>"
   ShowHTML "        <input type=""Hidden"" name=""CL"" value=""" & CL & """>"
   ShowHTML "        <input type=""Hidden"" name=""htBT"" value=""1"">"
   ShowHTML "        <input type=""Hidden"" name=""str2"" value=""10"">"
   ShowHTML "          <TR><TD colspan=""2"" align=""left"" valign""middle""><p align=""justify"">Selecione as formas de busca desejadas para listar as escolas:</p></td></tr>"
   SQL = "SELECT sq_tipo_cliente, ds_tipo_cliente FROM escTipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
   ConectaBD SQL
   ShowHTML "         <div id=""instituicao"">"
   ShowHTML "         <div class=""topo""> </div>"
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
   ShowHTML "        <div class=""base""> </div>"
   ShowHTML "        </div>"
   DesconectaBD
   wCont = 0
   wIN = 0
   sql1 = ""
   ShowHTML "          <div id=""opcoes"">"
   ShowHTML "          <div class=""topo""> </div>"
   sql = "SELECT DISTINCT a.curso as sq_especialidade, a.curso as ds_especialidade, 1 as nr_ordem, 'M' as tp_especialidade " & VbCrLf & _ 
        " from escTurma_Modalidade                 AS a " & VbCrLf & _ 
        "      INNER JOIN escTurma                 AS c ON (a.serie           = c.ds_serie) " & VbCrLf & _
        "      INNER JOIN escCliente               AS d ON (c.sq_site_cliente = d.sq_cliente and sq_cliente_pai = " & CL & ") " & VbCrLf & _
        "UNION " & VbCrLf & _ 
        "SELECT DISTINCT cast(a.sq_especialidade as varchar) as sq_especialidade, a.ds_especialidade,  " & VbCrLf & _ 
        "       case a.tp_especialidade when 'J' then '1' else a.nr_ordem end as nr_ordem, " & VbCrLf & _ 
        "       case a.tp_especialidade when 'J' then 'M' else a.tp_especialidade end as tp_especialidade" & VbCrLf & _ 
        " from escEspecialidade AS a " & VbCrLf & _ 
        "      INNER JOIN escEspecialidade_cliente AS c ON (a.sq_especialidade = c.sq_codigo_espec) " & VbCrLf & _
        "      INNER JOIN escCliente               AS d ON (c.sq_cliente       = d.sq_cliente and sq_cliente_pai = " & CL & ") " & VbCrLf & _
        " where a.tp_especialidade <> 'M' " & VbCrLf & _
        "ORDER BY a.nr_ordem, a.ds_especialidade " & VbCrLf
   ConectaBD SQL
     
   If Not RS.EOF Then
      wCont = 0
      wAtual = ""

      Do While Not RS.EOF
                
        If wAtual = "" or wAtual <> RS("tp_especialidade") Then
           wAtual = RS("tp_especialidade")
           If wAtual = "M" Then
              ShowHTML "          <dt><b>Etapas / Modalidades de ensino</b>:</dt>"
           ElseIf wAtual = "R" Then
              ShowHTML "          <dt><b>Em Regime de Intercomplementaridade</b>:</dt>"
           End If
        End If
        wCont = wCont + 1
        marcado = "N"
        For i = 1 to Request("H").Count
            If (RS("sq_especialidade")) = (Request("H")(i)) Then marcado = "S" End If
        Next
            
        If marcado = "S" Then
           ShowHTML chr(13) & "           <dd><input type=""checkbox"" name=""H"" value=""" & RS("sq_especialidade") & """ checked> " & RS("ds_especialidade") & "</dd>"
           sql1 = Request("H")
           wIN = 1
        Else
           ShowHTML chr(13) & "           <dd><input type=""checkbox"" name=""H"" value=""" & RS("sq_especialidade") & """> " & RS("ds_especialidade") & "</dd>"
        End If
         RS.MoveNext

        If (wCont Mod 2) = 0 Then 
          wCont = 0
          'ShowHTML "          <TR>"
         End If
      Loop
   End If
   ShowHTML "        <div class=""base""> </div>"
   ShowHTML "        </div>"
   ShowHTML "        <div id=""botao"">"
   ShowHTML "           <input type=""button"" id=""pesquisar"" name=""Botao"" value=""Pesquisar"" class=""botao"" onClick=""javascript: document.form1.Botao.disabled=true; document.form1.submit();"">"
   ShowHTML "        </div>"
   ShowHTML "        </form>"

   If Request("H") > "" Then
      w_h = ""
      For w_cont = 1 To Request.Form("H").Count
          If Nvl(Request("H")(w_cont),"0") <> "0" Then
             w_h = ",'" & Request("H")(w_cont) & "'" & w_h
          End If
      Next
      w_h = mid(w_h,2)
   End If
   if Request("Q") > "" or wIN > 0 or CL > 0 then
     ShowHTML "          <div id=""resultado"">"
     ShowHTML "          <h3>Resultado da pesquisa:</h3>"
     sql = "SELECT DISTINCT 1 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " & VbCrLf & _ 
           "       e.ds_logradouro, e.no_bairro,e.nr_cep, e.nr_fone_contato, d.localizacao, d.publica, " & VbCrLf & _ 
           "       d.ds_username, r.no_regiao " & VbCrLf & _ 
           "  from escCliente                            d " & VbCrLf & _
           "       INNER JOIN escCliente_Dados           e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
           "       INNER   JOIN escRegiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " & VbCrLf & _
           " where d.sq_cliente_pai = " & CL & " and d.ativo <> 'Nao' and d.publica = 'S' " & VbCrLf
               
     If Request("Z") > ""          Then sql = sql + "    and d.localizacao    = " & Request("Z")          & VbCrLf End If
     If Request("p_regional") > "" Then sql = sql + "    and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf End If
     If Request("q") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("q")          & VbCrLf End If

     If sql1 > ""         Then 
        sql = sql & _
              "    and (0 < (select count(*) from escEspecialidade_Cliente where sq_cliente = d.sq_cliente and cast(sq_codigo_espec as varchar) in (" + w_h + ")) or " & VbCrLf & _
              "         0 < (select count(*) from escTurma_Modalidade  w INNER JOIN escTurma x ON (w.serie = x.ds_serie) INNER JOIN escCliente y ON (x.sq_site_cliente = y.sq_cliente) where y.sq_cliente = d.sq_cliente and w.curso in (" + w_h + ")) " & VbCrLf & _
              "        ) " & VbCrLf 
     End If
     sql = sql + "ORDER BY tipo, ds_cliente " & VbCrLf

     ConectaBD SQL
       
     If Not RS.EOF Then

       If Request("str2") > "" Then RS.PageSize = cDbl(Request("str2")) Else RS.PageSize = 10 End If
           
       rs.AbsolutePage = nvl(Request("htBT"),1)

       'ShowHTML "          <TR><TD><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
       ShowHTML "          <UL> "
       'ShowHTML "            <tr><td width=""5%""><td width=""95%"">"

       While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(nvl(Request("htBT"),1))
         ShowHTML "                <font size=""2""><b>"
         If Not IsNull (RS("LN_INTERNET")) Then
            If inStr(lcase(RS("LN_INTERNET")),"http://") > 0 Then
               ShowHTML "                <li><a href=""http://" & replace(RS("LN_INTERNET"),"http://","") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b>"
            Else
              ShowHTML "                 <li><a href=""" & RS("LN_INTERNET") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b>"
            End If
         Else
            ShowHTML "<li>http://" & RS("DS_CLIENTE") & "</b><li>"
         End If
         If RS("tipo") <> "0" Then
           ShowHTML "<ul>"
           ShowHTML "<li>Telefone: " & RS("nr_fone_contato") & "</li>"
           ShowHTML "<li>Endere�o: " & RS("ds_logradouro")
           If Nvl(RS("no_bairro"),"nulo") <> "nulo" Then ShowHTML " - Bairro: " & RS("no_bairro") End If
           ShowHTML " - " & RS("no_regiao")
           sql2 = "SELECT f.ds_especialidade " & VbCrLf & _ 
                  "  from escEspecialidade_cliente AS e " & VbCrLf & _
                  "       INNER JOIN escEspecialidade AS f ON (e.sq_codigo_espec = f.sq_especialidade) " & VbCrLf & _
                  " WHERE e.sq_cliente = " & RS("sq_cliente") & " " & VbCrLf & _
                  "UNION " & VbCrLf & _
                  "SELECT f.ds_especialidade " & VbCrLf & _ 
                  "  from escEspecialidade f " & VbCrLf & _
                  " WHERE upper(ds_especialidade) like '%PROFISSIONAL%' " & VbCrLf & _
                  "   and 0 < (select count(sq_cliente) from escParticular_curso where sq_cliente = " & RS("sq_cliente") & ") " & VbCrLf & _
                  " UNION " & VbCrLf & _
                  "select distinct ltrim(rtrim(w.curso)) as ds_especialidade " & VbCrLf & _ 
                  "  from escTurma_Modalidade     w " & VbCrLf & _
                  "       INNER   JOIN escTurma   x ON (w.serie = x.ds_serie) " & VbCrLf & _
                  "         INNER JOIN escCliente y ON (x.sq_site_cliente = y.sq_cliente) " & VbCrLf & _
                  " where y.sq_cliente = " & RS("sq_cliente") & " " & VbCrLf & _
                  "ORDER BY 1 "
           RS1.Open sql2, dbms, adOpenForwardOnly
           wCont = 0
                  
            ShowHTML "<li>Oferta: "
            While Not RS1.EOF         
              If wCont = 0 Then
                   Response.write lcase(RS1("ds_especialidade"))
                  wCont = 1
                Else
                  Response.write ", " & lcase(RS1("ds_especialidade"))
                End if
              RS1.MoveNext
            Wend
            ShowHTML "</li>"
            RS1.Close
            ShowHTML "</ul></li>"
         End If

         RS.MoveNext

       Wend  
       ShowHTML "            </li>"
       ShowHTML "            </ul>" 
       ShowHTML "            </div>"  

       MontaBarra w_dir & "default.asp?EW=116&CL=" & CL & "&T=" & Request("T") & "&Q=" & Request("Q") & "&U=" & Request("U")& "&Z=" & Request("Z"), cDbl(RS.PageCount), cDbl(nvl(Request("htBT"),1)), cDbl(nvl(Request("str2"),10)), cDbl(RS.RecordCount)

     Else
       ShowHTML "            <p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<b>Nenhuma ocorr�ncia encontrada para as op��es acima."
     End If
   End If
End Sub

REM =========================================================================
REM Rotina de prepara��o para envio de e-mail do fale conosco
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
  w_html = w_html & "      <tr><td><font size=2><b><font color=""#BC3131"">ATEN��O</font>: Esta mensagem foi enviada pelo site da regional de ensino. Favor n�o responder.</b></font><br><br><td></tr>" & VbCrLf
  
  'Cabe�alho do e-Mail
  w_html = w_html & VbCrLf & "      <tr><td align=""left""><font size=""1"">&nbsp;</td>"
  w_html = w_html & VbCrLf & "      <tr bgcolor=""" & conTrBgColor & """><td align=""center""><table width=""99%"" border=""0"">"
  w_html = w_html & VbCrLf & "        <tr><td align=""center"" bgcolor=""#D0D0D0"" style=""border: 2px solid rgb(0,0,0);""><font size=""1""><b>REMETENTE</b></td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left"" bgcolor=""#FDFDFD"" style=""border: 1px solid rgb(0,0,0);""><font size=""1"">Nome: <b>" & w_fc_nome & "</td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left"" bgcolor=""#FDFDFD"" style=""border: 1px solid rgb(0,0,0);""><font size=""1"">e-Mail: <b><a href=""mailto:" & w_fc_email & """>" & w_fc_email & "</td>"
  w_html = w_html & VbCrLf & "        <tr><td align=""left"" bgcolor=""#FDFDFD"" style=""border: 1px solid rgb(0,0,0);""><font size=""1"">Endere�o: <b>" & w_fc_endereco & "</td>"
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
  w_html = w_html & VbCrLf & "      <tr><td align=""center"" bgcolor=""#D0D0D0"" style=""border: 2px solid rgb(0,0,0);""><font size=""1""><b>DADOS DA OCORR�NCIA</b></td>"
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

 ' Prepara os dados necess�rios ao envio
  w_assunto = "Encaminhamento de mensagem pelo site"
  
  RSM.Close
  
  
  If w_destinatario > "" Then
     ' Executa o envio do e-mail
     w_resultado = EnviaMail1(w_assunto, w_html, w_destinatario)
  End If
  
  ' Se ocorreu algum erro, avisa da impossibilidade de envio, sen�o volta para pagina do fale conosco
  If w_resultado > "" Then 
     ScriptOpen "JavaScript"
     ShowHTML "  alert('ATEN��O: n�o foi poss�vel proceder o envio do e-mail.\n" & w_resultado & "');" 
     ShowHTML "  history.back(1);"
     ScriptClose
  Else
     ScriptOpen "JavaScript"
     ShowHTML "  alert('Mensagem enviada com sucesso, uma c�pia foi enviada � voc� a t�tulo de registro!');" 
     ShowHTML "  location.href='" & conVirtualPath & w_dir & "Default.asp?EW=" & conWhatExFale & "&cl=" & cl & "';"
     ScriptClose
  End If

  Set w_html                   = Nothing
  Set w_destinatarios          = Nothing
  Set w_assunto                = Nothing
  Set RSM                      = Nothing
End Sub
REM =========================================================================
REM Fim da rotina da prepara��o para envio de e-mail
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da barra de navega��o de recordsets
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
     ShowHTML "<br>na p�gina " & p_AbsolutePage & " de " & p_PageCount & " p�ginas" & VbCrLf
     If p_AbsolutePage > 1 Then
        ShowHTML "<br>[<A TITLE=""Primeira p�gina"" HREF=""javascript:pagina(1)"">Primeira</A>]&nbsp;" & VbCrLf
        ShowHTML "[<A TITLE=""P�gina anterior"" HREF=""javascript:pagina(" & p_AbsolutePage-1 & ")"">Anterior</A>]&nbsp;" & VbCrLf
     Else
        ShowHTML "<br>[Primeira]&nbsp;" & VbCrLf
        ShowHTML "[Anterior]&nbsp;" & VbCrLf
     End If
     If p_PageCount = p_AbsolutePage Then
        ShowHTML "[Pr�xima]&nbsp;" & VbCrLf
        ShowHTML "[�ltima]" & VbCrLf
     Else
        ShowHTML "[<A TITLE=""P�gina seguinte"" HREF=""javascript:pagina(" & p_AbsolutePage+1 & ")"">Pr�xima</A>]&nbsp;" & VbCrLf
        ShowHTML "[<A TITLE=""�ltima p�gina"" HREF=""javascript:pagina(" & p_PageCount & ")"">�ltima</A>]" & VbCrLf
     End If
  Else
     ShowHTML "<br>na p�gina " & p_AbsolutePage & " de " & p_PageCount & " p�gina" & VbCrLf
  End If
  ShowHTML "</FORM></center>" & VbCrLf

End Sub

REM =========================================================================
REM Monta a tela de Valida��o de Senha
REM -------------------------------------------------------------------------
Public Sub ShowValida

  Dim strLabel, strTexto, strAction, strCampo
  
  If sstrIN > 0 Then
    sql = "SELECT b.ds_diretorio, b.ln_prop_pedagogica " & _
          "From escCliente as a INNER JOIN escCLIENTE_SITE as b on (a.sq_cliente = b.sq_cliente) " & _
          "                     INNER JOIN escModelo as c on (b.sq_modelo = c.sq_modelo) " & _
          "                     INNER JOIN escCliente_Dados as d on (a.sq_cliente = d.sq_cliente) " & _
          "WHERE a." & sstrEF 
      RS.CursorLocation = 3
      RS.Open sql, dbms, adOpenForwardOnly
      ShowHTML "<tr><td>"
      'If Nvl(RS("ln_prop_pedagogica"),"") > "" Then
      If RS.RecordCount > 0 Then
        ShowHTML "    <font size=1><b>"
        If InStr(lCase(RS("ln_prop_pedagogica")),lCase(".pdf")) > 0 Then
           ShowHTML "        A exibi��o do arquivo exige que o Acrobat Reader tenha sido instalado em seu computador."
           ShowHTML "        <br>Se o arquivo n�o for exibido no quadro abaixo, clique <a href=""http://www.adobe.com.br/products/acrobat/readstep2.html"" target=""_blank"">aqui</a> para instalar ou atualizar o Acrobat Reader."
        Else
           ShowHTML "        A exibi��o do arquivo exige o editor de textos Word ou equivalente. "
           ShowHTML "        <br>Se o arquivo n�o for exibido no quadro abaixo, verifique se o Word foi corretamente instalado em seu computador."        
        End If
        ShowHTML "<table align=""center"" width=""100%"" cellspacing=0 style=""border: 1px solid rgb(0,0,0);""><tr><td>"
        ShowHTML "    <iframe src=""" & RS("ds_diretorio") & "/" & RS("ln_prop_pedagogica") & """ width=""100%"" height=""510"">"
        ShowHTML "    </iframe>"
        ShowHTML "</table>"
      Else
        ShowHTML "    <font size=1><b>Composi��o Administrativa n�o informada."
      End If
  Else
     strLabel = "Alunos"
     strTexto = "Informe nos campos abaixo sua Matr�cula e Senha de Acesso, conforme informado pela escola. Se voc� n�o recebeu esses dados, clique no bot�o <i>Fale Conosco</I>, acima, para ver como entrar em contato com a escola e consegui-los."
     strAction = sstrSN & "?EW=" & conWhatSenha & "&EF=" & sstrEF & "&CL=" & CL & "&IN=" & sstrIN
     strCampo = "Matr�cula"

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
     ShowHTML "    alert(""Informe apenas letras e n�meros no campo \""Nome de Usu�rio\""."");" & chr(13)
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
     ShowHTML "    alert(""Informe apenas letras e n�meros no campo \""Senha\""."");" & chr(13)
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
    Case "composicao"       ShowValida
  End Select
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>
