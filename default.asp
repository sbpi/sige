<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="Constants_ADO.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="Funcoes.asp"-->
<!--#INCLUDE FILE="FuncoesGR.asp"-->
<%
REM =========================================================================
REM  /Default.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: Página do portal 
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 13/04/2000 23:10
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Informática
REM -------------------------------------------------------------------------

Private RS
Private RS1
Private sobjConn
Private sstrSN
Private sstrCL
Private strTitulo
Private marcado, i, p_regional, dbms

sstrSN = "default.Asp"

Public sstrEF
Public sstrEW
Public sstrIN
Public sstrDiretorio
Public sstrModelo
  
sstrEF = "sq_cliente=" & Session("CL")
sstrEW = Request("EW")
sstrIN = Request("IN")

Set RS = Server.CreateObject("ADODB.RecordSet")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
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
ShowHTML "   <link href=""css/estilo.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <link href=""css/print.css""  media=""print""  rel=""stylesheet"" type=""text/css"" />"
ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
ShowHTML "</head>"

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
ShowHTML "    <div id=""cabbase"">"
ShowHTML "      <div id=""busca"" valign=""center""><marquee width=""100%"" align=""middle""><font color=""white"" face=""Arial"" size=""2""><b>GDF - As grandes transformações também no ensino público.</b></font></marquee></div> "
ShowHTML "    </div>"
ShowHTML "  </div>"
ShowHTML "  <div id=""corpo"">"
ShowHTML "    <div id=""menuesq"">"
ShowHTML "      <div id=""logomenuesq""><img src=""img/fundo_logomenuesq.gif"" border=0></div>"
ShowHTML "      <ul id=""menugov"">"
ShowHTML "        <li>O ensino moderno oferece amplo apoio tecnológico ao estudante. A Internet modificou, substancialmente, o paradigma existente na área educacional."
ShowHTML "        <li>Modernizamos nossas atividades a fim de garantir a nossos alunos, pais e responsáveis qualidade, eficiência e eficácia na prestação de nossos serviços."
ShowHTML "        <li><a href=""newsletter.asp?ew=i"">Clique aqui para receber informativos da SEDF</a>."
ShowHTML "      </ul>"
ShowHTML "      <div id=""menusep""><hr /></div>"
ShowHTML "      <div id=""menunav""></div>"
ShowHTML "    </div>"
ShowHTML "    <div id=""texto""><!-- Conteúdo -->"
ShowHTML "        <table width=""570"" border=""0"">"

AbreSessao
Main
FechaSessao

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
REM =========================================================================
REM Monta a tela de Pesquisa
REM -------------------------------------------------------------------------
Public Sub ShowPesquisa

  Dim sql, sql2, wCont, sql1, wAtual, wIN
  
  ShowHTML "        <FORM ACTION=""default.asp?EW=121"" id=form1 name=form1 METHOD=""POST"">"
  ShowHTML "        <input type=""Hidden"" name=""htBT"" value=""1"">"
  ShowHTML "        <input type=""Hidden"" name=""str2"" value=""10"">"
  ShowHTML "          <TR><TD colspan=""2"" align=""left"" valign""middle""><p align=""justify"">Selecione as formas de busca desejadas para listar as escolas (pelo menos uma delas deve ser selecionada):</p></td></tr>"
  ShowHTML "          <tr><td colspan=2><b>Regional de ensino:</b><br><SELECT class=""texto"" NAME=""p_regional"">"
  SQL = "SELECT a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " & VbCrLf & _
        "  FROM escCLIENTE a " & VbCrLf & _
        "       inner join escTipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente) " & VbCrLf & _
        " WHERE b.tipo = 2 " & VbCrLf & _
        "ORDER BY a.ds_cliente" & VbCrLf
  ConectaBD SQL
  If RS.RecordCount > 1 Then ShowHTML "          <option value="""">Todas" End If
  While Not RS.EOF
     If cDbl(nvl(RS("sq_cliente"),0)) = cDbl(nvl(Request("p_regional"),0)) Then
        ShowHTML "          <option value=""" & RS("sq_cliente") & """ SELECTED>" & RS("ds_cliente")
     Else
        ShowHTML "          <option value=""" & RS("sq_cliente") & """>" & RS("ds_cliente")
     End If
     RS.MoveNext
  Wend
  ShowHTML "          </select>"
  DesconectaBD
  SQL = "SELECT * FROM escTipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
  ConectaBD SQL
  ShowHTML "          <tr><td colspan=2><b>Tipo de instituição:</b><br><SELECT class=""texto"" NAME=""Q"">"
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
          Else
             ShowHTML "          <TR><TD colspan=2><b>Outras</b>:"
          End If
       End If
       ShowHTML "          <TR>"
       wCont = wCont + 1
       marcado = "N"
       For i = 1 to Request("H").Count
           If cDbl(RS("sq_especialidade")) = cDbl(Request("H")(i)) Then marcado = "S" End If
       Next
       
       If marcado = "S" Then
          ShowHTML chr(13) & "               <td><input type=""checkbox"" name=""H"" value=""" & RS("sq_especialidade") & """ checked> " & RS("ds_especialidade")
          sql1 = Request("H")
          wIN = 1
       Else
          ShowHTML chr(13) & "               <td><input type=""checkbox"" name=""H"" value=""" & RS("sq_especialidade") & """> " & RS("ds_especialidade")
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

  if (Request("Q") > "") or (Request("p_regional") > "") or (wIN > 0) then
    ShowHTML "        <tr><td colspan=2><table width=""570"" border=""0"" align=""left"" cellpadding=0 cellspacing=0>"
    ShowHTML "          <TR><TD><hr>"
    ShowHTML "          <TR><TD><b><a name=""resultado"">Resultado da pesquisa:</a><br><br>"

    sql = "SELECT DISTINCT 0 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf,e.ds_logradouro,e.no_bairro,e.nr_cep " & VbCrLf & _
          "  from escCliente                               d " & VbCrLf & _
          "       INNER      JOIN escTipo_Cliente          b ON (d.sq_tipo_cliente  = b.sq_tipo_cliente and " & VbCrLf & _
          "                                                      b.tipo             = 2 " & VbCrLf & _
          "                                                     ) " & VbCrLf & _
          "       LEFT OUTER JOIN escCliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
          " where d.sq_cliente = " & Nvl(Request("p_regional"),0) & " " & VbCrLf & _
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
          " where g.sq_cliente = " & Nvl(Request("p_regional"),0) & " " & VbCrLf
    If Request("p_regional") > "" Then sql = sql + "    and c.sq_cliente_pai = " & Request("p_regional") & VbCrLf End If
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

    If Request("p_regional") > "" Then sql = sql + "    and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf End If
    If Request("q") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("q")          & VbCrLf End If

    if sql1 > "" then
       sql = sql + "  and a.sq_especialidade in (" + Request("H") + ")" & VbCrLf 
    end if
    sql = sql + "ORDER BY tipo, d.ds_cliente " & VbCrLf

    RS.Close
    RS.Open sql, sobjConn, adOpenStatic
    
    If Not RS.EOF Then

      If Request("str2") > "" Then RS.PageSize = cDbl(Request("str2")) Else RS.PageSize = 10 End If
      
      rs.AbsolutePage = Request("htBT")

      ShowHTML "          <TR><TD><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
      ShowHTML "            <tr><td width=""5%""><td width=""95%"">"

      While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(Request("htBT"))

        ShowHTML "            <TR><TD valign=""top""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center""> "
        ShowHTML "                <td align=""left""><font size=""2""><b>"
        If Not IsNull (RS("LN_INTERNET")) Then
           ShowHTML "                <a href=""http://" & replace(RS("LN_INTERNET"),"http://","") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b>"
        Else
           ShowHTML RS("DS_CLIENTE") & "</b>"
        End If
	    If RS("tipo") <> "0" Then
          ShowHTML "<br>Endereço: " & RS("ds_logradouro")
          If Nvl(RS("no_bairro"),"nulo") <> "nulo" Then ShowHTML " - Bairro: " & RS("no_bairro") End If
          If Nvl(RS("nr_cep"),"nulo") <> "nulo" and len(RS("nr_cep")) > 5 Then ShowHTML " - CEP: " & RS("nr_cep") End If
          ShowHTML "<br>Localização: " & RS("no_municipio") & "-" & RS("sg_uf")
	       sql2 = "SELECT f.ds_especialidade " & _ 
	             "from escEspecialidade_cliente AS e " & _
                 "     INNER JOIN escEspecialidade AS f ON (e.sq_codigo_espec = f.sq_especialidade) " & _
                 "WHERE e.sq_cliente = " & RS("sq_cliente") & " " & _
                 "ORDER BY f.ds_especialidade "

           RS1.Open sql2, sobjConn, adOpenForwardOnly

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
      MontaBarra "default.asp?EW=" & conWhatPrincipal & "&p_regional=" & Request("p_regional") & "&Q=" & Request("Q") & "&U=" & Request("U"), cDbl(RS.PageCount), cDbl(Request("htBT")), cDbl(Request("str2")), cDbl(RS.RecordCount)
      ShowHTML "</tr>"


    Else

      ShowHTML "            <TR><TD colspan=2><p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<b>Nenhuma ocorrência encontrada para as opções acima."

    End If

    ShowHTML "          </table>"
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Pesquisa
REM =========================================================================

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
     ShowHTML "<br>" & (p_RecordCount-((p_PageCount-1)*p_PageSize)) & " linhas apresentadas de " & p_RecordCount & " linhas." & VbCrLf
  Else
     ShowHTML "<br>" & p_PageSize & " linhas apresentadas de " & p_RecordCount & " linhas." & VbCrLf
  End If
  If p_PageSize < p_RecordCount Then
     ShowHTML "<br>na página " & p_AbsolutePage & " de " & p_PageCount & " páginas." & VbCrLf
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
     ShowHTML "<br>na página " & p_AbsolutePage & " de " & p_PageCount & " páginas." & VbCrLf
  End If
  ShowHTML "</FORM></center>" & VbCrLf

End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub Main

  strTitulo = "Pesquisa"
  ShowPesquisa

  Set RS = nothing
  Set RS1 = nothing
  Set sobjConn  = nothing

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

%>