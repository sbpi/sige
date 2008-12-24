<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="Constants_ADO.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="Funcoes.asp"-->
<!--#INCLUDE FILE="FuncoesGR.asp"-->
<%
REM =========================================================================
REM  /default.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papad�polis
REM Descricao: P�gina do portal 
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 13/04/2000 23:10
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Bras�lia - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Inform�tica
REM -------------------------------------------------------------------------

Private RS
Private RS1
Private sobjConn
Private sstrSN
Private sstrCL
Private strTitulo
Private marcado, i, p_regional, dbms

sstrSN = "default.asp"

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
ShowHTML "   <script src=""jquery.js""></script>"
ShowHTML "   <script type=""text/javascript"" language=""javascript"">"
ShowHTML "   $(document).ready(function() {"
ShowHTML "   $(""a#showhide"").click(function() {"
ShowHTML "   $(""#content"").slideToggle('slow');"
ShowHTML "   if ($(""#collexp"").html() == '(exibir)'){"
ShowHTML "   $(""#collexp"").html('(ocultar)');"
ShowHTML "   }else{"
ShowHTML "   $(""#collexp"").html('(exibir)');"
ShowHTML "   }"
ShowHTML "   });"
ShowHTML "   });"
ShowHTML "   </script>"
ShowHTML "</head>"

If (Request("H") > "") or (Request("Q") > "") or (Request("p_regional") > "") or (wIN > 0) Then
   ShowHTML "<body onLoad=""location.href='#resultado'"">"
Else
   ShowHTML "<body>"
End If
 %>
<div id="pagina">
  <div id="topo"> </div>
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
          <li><a class="aluno" href="http://www.se.df.gov.br/300/30002001.asp">Aluno</a></li>
          <li><a class="educador" href="http://www.se.df.gov.br/300/30003001.asp">Educador</a></li>
          <li><a class="comunidade" href="http://www.se.df.gov.br/300/30004001.asp">Comunidade</a></li>
        </ul>
      </div>
    </div>
    <div id="menuMiddle"> </div>
  </div>
  <div class="clear"></div>
<div id="menuBottom">
  <ul>
    <li><a href="#"><span>Inicial</span></a></li>
    <li><a href="newsletter.asp?ew=i"><span>Assine nosso boletim</span></a></li>
  </ul>
  <div class="clear"></div>
</div>
<div id="conteudo"><h2>Solu��o Integrada de Gest�o Educacional - SIGE</h2>
<%
AbreSessao
'Main
'FechaSessao
ShowHTML " <h3>Busca avan�ada</h3>"

    <!-- <form action="" id="formConteudo">-->
    ShowHTML " <FORM ACTION=""default.asp?EW=121"" id=""form1"" name=""form1"" METHOD=""POST"">"
    ShowHTML " <div id=""regional"">"
    ShowHTML " <div class=""topo""> </div>"
   
    Dim sql, sql2, wCont, sql1, wAtual, wIN

    ShowHTML "<SCRIPT LANGUAGE=""JAVASCRIPT""><!--  "
    ShowHTML "function marcaTipo() {  "
    ShowHTML "  document.form1.submit();"
    ShowHTML "}  "
    ShowHTML "--></SCRIPT>"
  
    ShowHTML "        <FORM ACTION=""default.asp?EW=121"" id=""form1"" name=""form1"" METHOD=""POST"">"
    ShowHTML "        <input type=""Hidden"" name=""htBT"" value=""1"">"
    ShowHTML "        <input type=""Hidden"" name=""str2"" value=""10"">"
      'ShowHTML "          <TR><TD colspan=""2"" align=""left"" valign""middle""><p align=""justify"">Selecione formas de busca desejadas para listar escolas (pelo menos uma delas deve ser selecionada):</p></td></tr>"

    ' Verifica se h� alguma escola particular no cadastro. Se n�o, marca para tratar apenas escolas p�blicas
    SQL = "select count(sq_cliente) cadprivada from sbpi.Cliente_Particular"
    ConectaBD SQL
    'If RS("cadprivada") > 0 Then
    '    If Nvl(Request("T"),"S") = "S" Then ShowHtml("<input type=""radio"" name=""T"" value=""S"" checked=""checked"" onclick=""marcaTipo();""> Rede p�blica")  Else ShowHtml("<input type=""radio"" name=""T"" value=""S"" onclick=""marcaTipo();""> Rede p�blica") End If
    '    If Request("T") = "P"          Then ShowHtml("<input type=""radio"" name=""T"" value=""P"" checked=""checked"" onclick=""marcaTipo();""> Rede privada")  Else ShowHtml("<input type=""radio"" name=""T"" value=""P"" onclick=""marcaTipo();""> Rede privada") End If        
    '    ShowHTML "<br/><br/>"
    'Else
       ShowHTML "        <input type=""Hidden"" name=""T"" value=""S"">"
    'End If
    If Request("T") = "P" Then
       ' Tratamento para escolas particulares
       ShowHTML "          <h4>Regi�o administrativa:</h4>"
       SQL = "SELECT distinct b.sq_regiao_adm, b.no_regiao " & VbCrLf & _
             "  from sbpi.CLIENTE a " & VbCrLf & _
             "       inner join sbpi.Regiao_Administrativa b on (a.sq_regiao_adm = b.sq_regiao_adm) " & VbCrLf & _
             "ORDER BY b.no_regiao" & VbCrLf
       ConectaBD SQL
       ShowHTML " <SELECT class=""texto"" NAME=""p_regional"">"
       If RS.RecordCount > 1 Then ShowHTML "          <option value="""">Todas" End If
       While Not RS.EOF
          If cDbl(nvl(RS("sq_regiao_adm"),0)) = cDbl(nvl(Request("p_regional"),0)) Then
             ShowHTML "          <option value=""" & RS("sq_regiao_adm") & """ SELECTED>" & RS("no_regiao")
          Else
             ShowHTML "          <option value=""" & RS("sq_regiao_adm") & """>" & RS("no_regiao")
          End If
          RS.MoveNext
       Wend
       ShowHTML "          </select>"
       RS.Close
    Else
       ' Sele��o de regional
       SQL = "SELECT a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " & VbCrLf & _
             "  from sbpi.CLIENTE a " & VbCrLf & _
             "       inner join sbpi.Tipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente) " & VbCrLf & _
             " WHERE b.tipo = 2 " & VbCrLf & _
             "ORDER BY a.ds_cliente"    
       ShowHTML " <h4>Regional de ensino:</h4>"    
       ConectaBD SQL
       ShowHTML " <SELECT class=""texto"" NAME=""p_regional"">"    
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

       ShowHTML " <div class=""base""></div>"
       ShowHTML " </div>"
       ShowHTML " <div id=""instituicao"">"
       ShowHTML " <div class=""topo""> </div>"
    
       ' Sele��o de tipo de institui��o
       SQL = "SELECT * from sbpi.Tipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
       ConectaBD SQL   
       ShowHTML "         <h4>Tipo de institui��o:</h4><SELECT class=""texto"" NAME=""Q"">"
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

       ShowHTML " <div class=""base""> </div>"
       ShowHTML " </div>"

       ShowHTML " <div id=""urbana"">"
       ShowHTML " <div class=""topo""> </div>"
  
       ShowHTML "         <h4>Localiza��o:</h4><div class=""radio_option"">"
       If Request("Z") = "2" Then ShowHtml("<input type=""radio"" name=""Z"" value=""2"" checked=""checked""> Rural")  Else ShowHtml("<input type=""radio"" name=""Z"" value=""2""> Rural")  End If
       If Request("Z") = "1" Then ShowHtml("&nbsp;<input type=""radio"" name=""Z"" value=""1"" checked=""checked""> Urbana") Else ShowHtml("&nbsp;<input type=""radio"" name=""Z"" value=""1""> Urbana") End If
       If Request("Z") = ""  Then ShowHtml("&nbsp;<input type=""radio"" name=""Z"" value=""""  checked=""checked""> Ambas")  Else ShowHtml("&nbsp;<input type=""radio"" name=""Z"" value="""" > Ambas")  End If

       ShowHTML " </div><div class=""base""> </div>"  
       ShowHTML " </div>"  

       ShowHTML " <div id=""opcoes"">"
       ShowHTML " <div class=""topo""> </div>"
  
       wCont = 0
       wIN = 0
       sql1 = ""

       ' Sele��o de etapas/modalidades
       sql = "SELECT DISTINCT a.curso sq_especialidade, a.curso ds_especialidade, '1' nr_ordem, 'M' tp_especialidade " & VbCrLf & _ 
             " from sbpi.Turma_Modalidade                 a " & VbCrLf & _ 
             "      INNER join sbpi.Turma                 c ON (a.serie           = c.ds_serie) " & VbCrLf & _
             "      INNER join sbpi.Cliente               d ON (c.sq_cliente = d.sq_cliente) " & VbCrLf & _
             "UNION " & VbCrLf & _ 
             "SELECT DISTINCT to_char(a.sq_especialidade) sq_especialidade, a.ds_especialidade,  " & VbCrLf & _ 
             "       case a.tp_especialidade when 'J' then '1' else to_char(a.nr_ordem) end nr_ordem, " & VbCrLf & _ 
             "       case a.tp_especialidade when 'J' then 'M' else a.tp_especialidade end tp_especialidade" & VbCrLf & _ 
             " from sbpi.Especialidade a " & VbCrLf & _ 
             "      INNER join sbpi.Especialidade_cliente c ON (a.sq_especialidade = c.sq_especialidade) " & VbCrLf & _
             "      INNER join sbpi.Cliente               d ON (c.sq_cliente       = d.sq_cliente) " & VbCrLf & _
             " where a.tp_especialidade <> 'M' " & VbCrLf & _
             "ORDER BY nr_ordem, ds_especialidade " & VbCrLf
       ConectaBD sql
      
       If Not RS.EOF Then
          wCont = 0
          wAtual = ""

          Do While Not RS.EOF
                 
             If wAtual = "" or wAtual <> RS("tp_especialidade") Then
                wAtual = RS("tp_especialidade")
                If wAtual = "M" Then
                   If Nvl(Request("T"),"S") = "P" Then
                      ShowHTML "<dl>"
                      ShowHTML "          <dt><b>Modalidades de ensino</b>: </dt>"
                   Else
                      ShowHTML "          <dt><b>Etapas / Modalidades de ensino</b>: </dt>"
                   End If
                ElseIf wAtual = "R" Then
                   ShowHTML "          <dt><b>Em Regime de Intercomplementaridade</b>: </dt>"
                Else
                   ShowHTML "          <dt><b>Outras</b>: </dt>"
                End If
             End If
             wCont = wCont + 1
             marcado = "N"
             For i = 1 to Request("H").Count
                 If RS("sq_especialidade") = Request("H")(i) Then marcado = "S" End If
             Next
             
             If marcado = "S" Then
                ShowHTML chr(13) & "               <dd><input type=""checkbox"" name=""H"" value=""" & RS("sq_especialidade") & """ checked> " & RS("ds_especialidade")
                sql1 = Request("H")
                wIN = 1
             Else
                ShowHTML chr(13) & "               <dd><input type=""checkbox"" name=""H"" value=""" & RS("sq_especialidade") & """> " & RS("ds_especialidade")
             End If
           RS.MoveNext

             If (wCont Mod 2) = 0 Then 
               wCont = 0
           End If
       Loop
    End If    
  End If    
  ShowHTML " </dd>"  
  ShowHTML " <div class=""base""> </div>"  
  ShowHTML " </div>"  
  ShowHTML " <div id=""botao"">"
  ShowHTML " <input id=""pesquisar"" type=""button"" name=""Botao"" value=""Pesquisar"" class=""botao"" onClick=""javascript: document.form1.Botao.disabled=true; document.form1.submit();"">"
  ShowHTML " </form>"
  ShowHTML " </div> "
  ShowHTML " <div id=""resultado"">"
  FechaSessao

  If Request("H") > "" Then
     w_h = ""
     For w_cont = 1 To Request.Form("H").Count
         If Nvl(Request("H")(w_cont),"0") <> "0" Then
            w_h = ",'" & Request("H")(w_cont) & "'" & w_h
         End If
     Next
     w_h = mid(w_h,2)
  End If
  if (Request("Z") > "") or (Request("Q") > "") or (Request("p_regional") > "") or (wIN > 0) then


    ShowHTML " <h3>Resultado da Pesquisa</h3>"

    sql = "SELECT DISTINCT 0 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " & VbCrLf & _
          "       e.ds_logradouro,e.no_bairro,e.nr_cep, e.nr_fone_contato, d.localizacao, 'S' publica, null ds_username, " & VbCrLf & _
          "       r.no_regiao " & VbCrLf & _
          "  from sbpi.Cliente                               d " & VbCrLf & _
          "       INNER      join sbpi.Tipo_Cliente          b ON (d.sq_tipo_cliente  = b.sq_tipo_cliente and " & VbCrLf & _
          "                                                      b.tipo             = 2 " & VbCrLf & _
          "                                                     ) " & VbCrLf & _
          "       LEFT OUTER join sbpi.Cliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
          "       INNER      join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " & VbCrLf & _
          " where d.ativo <> 'Nao' and d.publica = 'S' and 'P' <> '" & REQUEST("T") & "' and d.sq_cliente = " & Nvl(Request("p_regional"),0) & " " & VbCrLf
    If Request("Z") > "" Then sql = sql + "    and 0 < (select count(*) from sbpi.Cliente where sq_cliente_pai = d.sq_cliente and localizacao = " & Request("Z") & ") " & VbCrLf End If
    If Request("q") > "" Then sql = sql + "    and 0 < (select count(*) from sbpi.Cliente where sq_cliente_pai = d.sq_cliente and d.sq_tipo_cliente= " & Request("q") & ") " & VbCrLf End If
    If sql1 > "" Then 
       sql = sql & _
             "    and (0 < (select count(sq_cliente) from sbpi.Especialidade_Cliente where sq_cliente_pai = d.sq_cliente and to_char(sq_especialidade) in (" + w_h + ")) or " & VbCrLf & _
             "         0 < (select count(*) from sbpi.Turma_Modalidade  w INNER join sbpi.Turma x ON (w.serie = x.ds_serie) INNER join sbpi.Cliente y ON (x.sq_cliente = y.sq_cliente) where y.sq_cliente_pai = d.sq_cliente and w.curso in (" + w_h + ")) " & VbCrLf & _
             "        ) " & VbCrLf 
    End If
    sql = sql & _
          "UNION " & VbCrLf & _
          "SELECT DISTINCT 0 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " & VbCrLf & _
          "       e.ds_logradouro,e.no_bairro,e.nr_cep, e.nr_fone_contato, d.localizacao, 'S' publica, null ds_username, " & VbCrLf & _
          "       r.no_regiao " & VbCrLf & _
          "  from sbpi.Cliente                               d " & VbCrLf & _
          "       LEFT OUTER join sbpi.Cliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
          "       INNER      join sbpi.Cliente_Site          a ON (d.sq_cliente       = a.sq_cliente) " & VbCrLf & _
          "       INNER      join sbpi.Tipo_Cliente          b ON (d.sq_tipo_cliente  = b.sq_tipo_cliente and " & VbCrLf & _
          "                                                      b.tipo             = 2 " & VbCrLf & _
          "                                                     ) " & VbCrLf & _
          "       INNER      join sbpi.Cliente               c ON (d.sq_cliente       = c.sq_cliente_pai) " & VbCrLf & _
          "       INNER      join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " & VbCrLf & _
          " where d.ativo <> 'Nao' and 'P' <> '" & REQUEST("T") & "' " & VbCrLf
          
        
    If Request("Z") > ""          Then sql = sql + "    and d.localizacao    = " & Request("Z")          & VbCrLf End If
    If Request("p_regional") > "" Then sql = sql + "    and c.sq_cliente_pai = " & Request("p_regional") & VbCrLf End If
    If Request("q") > ""          Then sql = sql + "    and c.sq_tipo_cliente= " & Request("q")          & VbCrLf End If
    if sql1 > "" then
       sql = sql & _
             "    and (0 < (select count(*) from sbpi.Especialidade_Cliente w inner join sbpi.Cliente x on (w.sq_cliente = x.sq_cliente) where x.sq_cliente_pai = d.sq_cliente and to_char(w.sq_especialidade) in (" + w_h + ")) or " & VbCrLf & _
             "         0 < (select count(*) from sbpi.Turma_Modalidade  w INNER join sbpi.Turma x ON (w.serie = x.ds_serie) INNER join sbpi.Cliente y ON (x.sq_cliente = y.sq_cliente) where y.sq_cliente_pai = d.sq_cliente and w.curso in (" + w_h + ")) " & VbCrLf & _
             "        ) " & VbCrLf 
    end if
    sql = sql & _
          "UNION " & VbCrLf & _
          "SELECT DISTINCT 1 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " & VbCrLf & _ 
          "       f.endereco logradouro, null no_bairro, null nr_cep, null nr_fone_contato, d.localizacao, d.publica, " & VbCrLf & _ 
          "       d.ds_username, r.no_regiao " & VbCrLf & _
          "  from sbpi.Cliente                          d " & VbCrLf & _
          "       INNER join sbpi.Cliente_Particular    f ON (d.sq_cliente       = f.sq_cliente) " & VbCrLf & _
          "       LEFT  join sbpi.Especialidade_cliente g ON (d.sq_cliente       = g.sq_cliente) " & VbCrLf & _
          "       INNER join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " & VbCrLf & _
          " where d.ativo <> 'Nao' and f.situacao = 1 and d.publica = 'N' and 'S' <> '" & REQUEST("T") & "' " & VbCrLf
          
        
    If Request("Z") > ""          Then sql = sql + "    and d.localizacao    = " & Request("Z")          & VbCrLf End If
    If Request("p_regional") > "" Then sql = sql + "    and d.sq_regiao_adm  = " & Request("p_regional") & VbCrLf End If
    If Request("q") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("q")          & VbCrLf End If
    if sql1 > "" then
       sql = sql + "  and (to_char(g.sq_especialidade) in (" + w_h + ")) " & VbCrLf
    end if
    sql = sql & _   
          "UNION " & VbCrLf & _ 
          "SELECT DISTINCT 1 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " & VbCrLf & _ 
          "       e.ds_logradouro, e.no_bairro,e.nr_cep, e.nr_fone_contato, d.localizacao, d.publica, " & VbCrLf & _ 
          "       d.ds_username, r.no_regiao " & VbCrLf & _ 
          "  from sbpi.Cliente                            d " & VbCrLf & _
          "       INNER join sbpi.Cliente_Dados           e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
          "       INNER   join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " & VbCrLf & _
          " where d.ativo <> 'Nao' and d.publica = 'S' and 'P' <> '" & REQUEST("T") & "' " & VbCrLf
          
    If Request("Z") > ""          Then sql = sql + "    and d.localizacao    = " & Request("Z")          & VbCrLf End If
    If Request("p_regional") > "" Then sql = sql + "    and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf End If
    If Request("q") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("q")          & VbCrLf End If

    If sql1 > ""         Then 
       sql = sql & _
             "    and (0 < (select count(*) from sbpi.Especialidade_Cliente where sq_cliente = d.sq_cliente and to_char(sq_especialidade) in (" + w_h + ")) or " & VbCrLf & _
             "         0 < (select count(*) from sbpi.Turma_Modalidade  w INNER join sbpi.Turma x ON (w.serie = x.ds_serie) INNER join sbpi.Cliente y ON (x.sq_cliente = y.sq_cliente) where y.sq_cliente = d.sq_cliente and w.curso in (" + w_h + ")) " & VbCrLf & _
             "        ) " & VbCrLf 
    End If
    sql = sql + "ORDER BY tipo, ds_cliente " & VbCrLf
    
    RS.Open sql, sobjConn, adOpenStatic
   
    If Not RS.EOF Then

      If Request("str2") > "" Then RS.PageSize = cDbl(Request("str2")) Else RS.PageSize = 10 End If
      
      rs.AbsolutePage = Request("htBT")

      'ShowHTML "          <TR><TD><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
      ShowHTML "          <UL> "
      'ShowHTML "            <tr><td width=""5%""><td width=""95%"">"

      While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(Request("htBT"))
        If RS("PUBLICA") = "S" Then
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
        Else
           ShowHTML "                <li><a href=""/modelos/ModPart/default.asp?EW=110&EF=sq_cliente=" & RS("SQ_CLIENTE") & "&CL=" & RS("SQ_CLIENTE") & """ target=""" & RS("DS_USERNAME") & """>" & RS("DS_CLIENTE") & "</a></b>"
        End If
        If RS("tipo") <> "0" Then
          ShowHTML "<ul>"
          ShowHTML "<li>Telefone: " & RS("nr_fone_contato") & "</li>"
          ShowHTML "<li>Endere�o: " & RS("ds_logradouro")
          If Nvl(RS("no_bairro"),"nulo") <> "nulo" Then ShowHTML " - Bairro: " & RS("no_bairro") End If
          'If Nvl(RS("nr_cep"),"nulo") <> "nulo" and len(RS("nr_cep")) > 5 Then ShowHTML " - CEP: " & RS("nr_cep") & " - " & RS("no_regiao") End If
          ShowHTML " - " & RS("no_regiao")
          'ShowHTML "<li>Localiza��o: " & RS("no_regiao") & "</li>"
          sql2 = "SELECT f.ds_especialidade " & VbCrLf & _ 
                 "  from sbpi.Especialidade_cliente e " & VbCrLf & _
                 "       INNER join sbpi.Especialidade f ON (e.sq_especialidade = f.sq_especialidade) " & VbCrLf & _
                 " WHERE e.sq_cliente = " & RS("sq_cliente") & " " & VbCrLf & _
                 "UNION " & VbCrLf & _
                 "SELECT f.ds_especialidade " & VbCrLf & _ 
                 "  from sbpi.Especialidade f " & VbCrLf & _
                 " WHERE upper(ds_especialidade) like '%PROFISSIONAL%' " & VbCrLf & _
                 "   and 0 < (select count(sq_cliente) from sbpi.Particular_curso where sq_cliente = " & RS("sq_cliente") & ") " & VbCrLf & _
                 "UNION " & VbCrLf & _
                 "select distinct ltrim(rtrim(w.curso)) ds_especialidade " & VbCrLf & _ 
                 "  from sbpi.Turma_Modalidade     w " & VbCrLf & _
                 "       INNER   join sbpi.Turma   x ON (w.serie = x.ds_serie) " & VbCrLf & _
                 "         INNER join sbpi.Cliente y ON (x.sq_cliente = y.sq_cliente) " & VbCrLf & _
                 " where y.sq_cliente = " & RS("sq_cliente") & VbCrLf & _
                 "ORDER BY ds_especialidade "

          RS1.Open sql2, sobjConn, adOpenForwardOnly

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

    MontaBarra "default.asp?EW=" & conWhatPrincipal & "&p_regional=" & Request("p_regional") & "&T=" & Request("T") & "&Q=" & Request("Q") & "&U=" & Request("U")& "&Z=" & Request("Z"), cDbl(RS.PageCount), cDbl(Request("htBT")), cDbl(Request("str2")), cDbl(RS.RecordCount)

    Else
      ShowHTML "            <p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<b>Nenhuma ocorr�ncia encontrada para op��es acima."
    End If

 End If
  %>  
    </div>  
    <div class="clear"/>
    </div>

<%
'ShowHTML "    <div id=""cabbase"">"
'ShowHTML "      <div id=""busca"" valign=""center""><marquee width=""100%"" align=""middle""><font color=""white"" face=""Arial"" size=""2""><b>GDF - grandes transforma��es tamb�m no ensino p�blico.</b></font></marquee></div> "
'ShowHTML "    </div>"
'ShowHTML "  </div>"
'ShowHTML "  <div id=""corpo"">"
'ShowHTML "      <ul id=""menugov"">"
'ShowHTML "        <li>O ensino moderno oferece amplo apoio tecnol�gico ao estudante. A Internet modificou, substancialmente, o paradigma existente na �rea educacional."
'ShowHTML "        <li>Modernizamos nossas atividades a fim de garantir a nossos alunos, pais e respons�veis qualidade, efici�ncia e efic�cia na presta��o de nossos servi�os."
'ShowHTML "        <li><a href=""newsletter.asp?ew=i"">Clique aqui para receber informativos da SEDF</a>."
'ShowHTML "      </ul>"
'ShowHTML "      <div id=""menusep""><hr /></div>"
'ShowHTML "      <div id=""menunav""></div>"
'ShowHTML "    </div>"

'ShowHTML "      <div></div>"
'ShowHTML "      <div></div>"
'ShowHTML "      <div><br /></div> "
ShowHTML "  </div>"
ShowHTML "</div>"
      %><div id="rodape">
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

REM =========================================================================
REM Monta a tela de ofertas da escola
REM -------------------------------------------------------------------------
Public Sub ShowOferta

Dim sql2, biblioteca, linguas
Dim oferta(50), w_modalidade, w_serie, w_turno, w_mod_atual, w_ser_atual, w_tur_atual
Dim w_texto, w_numero, w_pos, i, w_cont, w_temp

    w_cont = 0
    
    ' Recupera a oferta a partir das turmas da escola
    sql  = "select distinct a.sq_cliente, b.nome ds_serie, b.curso ds_modalidade, case rtrim(ltrim(a.ds_turno)) when 'M' then 1 when 'V' then 2 else 3 end ds_turno " & VbCrLf & _ 
           "  from sbpi.Turma a inner join sbpi.Turma_Modalidade b on (upper(rtrim(ltrim(a.ds_serie)))=upper(rtrim(ltrim(b.serie))))" & VbCrLf & _
           " where sq_cliente = " & CL & " " & VbCrLf & _
           "order by 3,2 "
    RS.Open sql, sobjConn, adOpenForwardOnly

    If Not RS.EOF Then
        w_cont = 1
        While Not RS.EOF
          oferta(w_cont) = RS("ds_modalidade") & "}" & RS("ds_serie") & "|" & exibeTurno(RS("ds_turno"))
          w_cont = w_cont + 1
          RS.MoveNext
        Wend
        Sort oferta

        ShowHTML("<p><b>Etapas / Modalidades de ensino oferecidas:</b>")
        ShowHTML("<dl>")
        
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
        ShowHTML("</dl>")
    End If
    
    sql2 = "SELECT f.ds_especialidade " & VbCrLf & _ 
           "  from sbpi.Especialidade_cliente e " & VbCrLf & _
           "       INNER join sbpi.Especialidade f ON (e.sq_especialidade = f.sq_especialidade and " & VbCrLf & _
           "                                            'M'               <> f.tp_especialidade " & VbCrLf & _
           "                                           ) " & VbCrLf & _
           " WHERE e.sq_cliente = " & CL & " " & VbCrLf & _
           "ORDER BY ds_especialidade "
    RS1.Open sql2, sobjConn, adOpenForwardOnly
    
    While Not RS1.EOF
      If uCase(RS1("ds_especialidade")) = "BIBLIOTECA" Then
          biblioteca = RS1("ds_especialidade")
      End If
      If uCase(RS1("ds_especialidade")) = "L�NGUAS" Then
          linguas =  RS1("ds_especialidade")
      End If
      RS1.MoveNext
    Wend

    if biblioteca <> "" OR linguas <> "" Then
       w_cont = 1
       ShowHTML("<p><b>Em Regime de Intercomplementaridade: </b></p>")
       If(biblioteca <> "") Then
          ShowHTML("<li>" & biblioteca & "</li>")
       End If
       If(linguas <> "") Then
          ShowHTML("<li>" & linguas & "</li>")
       End If
    End If
    RS1.Close
    If w_cont = 0 Then
       ShowHTML "<tr><td><font size=1><b>Oferta n�o informada."
    End If
    
End Sub

REM =========================================================================
REM Monta a tela de Pesquisa
REM -------------------------------------------------------------------------
Public Sub ShowPesquisa

  Dim sql, sql2, wCont, sql1, wAtual, wIN

  ShowHTML "<SCRIPT LANGUAGE=""JAVASCRIPT""><!--  "
  ShowHTML "function marcaTipo() {  "
  ShowHTML "  document.form1.submit();"
  ShowHTML "}  "
  ShowHTML "--></SCRIPT>"
  
  ShowHTML "        <FORM ACTION=""default.asp?EW=121"" id=form1 name=form1 METHOD=""POST"">"
  ShowHTML "        <input type=""Hidden"" name=""htBT"" value=""1"">"
  ShowHTML "        <input type=""Hidden"" name=""str2"" value=""10"">"
  'ShowHTML "          <TR><TD colspan=""2"" align=""left"" valign""middle""><p align=""justify"">Selecione formas de busca desejadas para listar escolas (pelo menos uma delas deve ser selecionada):</p></td></tr>"

  ' Verifica se existe alguma escola particular
  SQL = "select count(sq_cliente) cadprivada from sbpi.Cliente_Particular"
  ConectaBD SQL

  ShowHtml("<tr bgcolor=""#DFDFDF""><td colspan=2><font size=1><b>Exibir institui��es de ensino da:</font><br><font size=2>")
  
  ' Permite ao usu�rio selecionar entre escolas p�blicas e particulares
  If RS("cadprivada") > 0 Then
    If Nvl(Request("T"),"S") = "S" Then ShowHtml("<input type=""radio"" name=""T"" value=""S"" checked=""checked"" onclick=""marcaTipo();""> Rede p�blica")  Else ShowHtml("<input type=""radio"" name=""T"" value=""S"" onclick=""marcaTipo();""> Rede p�blica") End If
    If Request("T") = "P"          Then ShowHtml("<input type=""radio"" name=""T"" value=""P"" checked=""checked"" onclick=""marcaTipo();""> Rede privada")  Else ShowHtml("<input type=""radio"" name=""T"" value=""P"" onclick=""marcaTipo();""> Rede privada") End If
  Else
     ShowHTML "        <input type=""Hidden"" name=""T"" value=""S"">"
  End If
  
  If Request("T") = "P" Then
    ShowHTML "          <tr><td colspan=2><b>Regi�o administrativa:</b><br><SELECT class=""texto"" NAME=""p_regional"">"
    SQL = "SELECT distinct b.sq_regiao_adm, b.no_regiao " & VbCrLf & _
          "  from sbpi.CLIENTE a " & VbCrLf & _
          "       inner join sbpi.Regiao_Administrativa b on (a.sq_regiao_adm = b.sq_regiao_adm) " & VbCrLf & _
          "ORDER BY b.no_regiao" & VbCrLf
    ConectaBD SQL
    If RS.RecordCount > 1 Then ShowHTML "          <option value="""">Todas" End If
    While Not RS.EOF
       If cDbl(nvl(RS("sq_regiao_adm"),0)) = cDbl(nvl(Request("p_regional"),0)) Then
          ShowHTML "          <option value=""" & RS("sq_regiao_adm") & """ SELECTED>" & RS("no_regiao")
       Else
          ShowHTML "          <option value=""" & RS("sq_regiao_adm") & """>" & RS("no_regiao")
       End If
       RS.MoveNext
    Wend
    ShowHTML "          </select>"
    DesconectaBD
  Else
    ShowHTML "          <tr><td colspan=2><b>Regional de ensino:</b><br><SELECT class=""texto"" NAME=""p_regional"">"
    SQL = "SELECT a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " & VbCrLf & _
          "  from sbpi.CLIENTE a " & VbCrLf & _
          "       inner join sbpi.Tipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente) " & VbCrLf & _
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

    SQL = "SELECT * from sbpi.Tipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
    ConectaBD SQL
    ShowHTML "         <tr><td colspan=2><b>Tipo de institui��o:</b><br><SELECT class=""texto"" NAME=""Q"">"
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
  End If

  ShowHtml("<tr><td colspan=2><b>Localiza��o</b>:<br>")
  If Request("Z") = "2" Then ShowHtml("<input type=""radio"" name=""Z"" value=""2"" checked=""checked""> Rural")  Else ShowHtml("<input type=""radio"" name=""Z"" value=""2""> Rural")  End If
  If Request("Z") = "1" Then ShowHtml("<input type=""radio"" name=""Z"" value=""1"" checked=""checked""> Urbana") Else ShowHtml("<input type=""radio"" name=""Z"" value=""1""> Urbana") End If
  If Request("Z") = ""  Then ShowHtml("<input type=""radio"" name=""Z"" value=""""  checked=""checked""> Ambas")  Else ShowHtml("<input type=""radio"" name=""Z"" value="""" > Ambas")  End If
      
  wCont = 0
  wIN = 0
  sql1 = ""

  sql = "SELECT DISTINCT a.* " & VbCrLf & _ 
        " from sbpi.Especialidade a " & VbCrLf & _ 
        "      INNER join sbpi.Especialidade_cliente c ON (a.sq_especialidade = c.sq_especialidade) " & VbCrLf & _
        "      INNER join sbpi.Cliente               d ON (c.sq_cliente       = d.sq_cliente) " & VbCrLf
  If Nvl(Request("T"),"S") = "P" Then 
     sql = sql & _
        " where a.tp_especialidade = 'M' and a.sq_especialidade not in (4,9,10,13,14) " & VbCrLf
  End If
  sql = sql & _
        "ORDER BY a.nr_ordem, a.ds_especialidade " & VbCrLf

  ConectaBD SQL   
    
  If Not RS.EOF Then
  wCont = 0
  wAtual = ""

  Do While Not RS.EOF
           
       If wAtual = "" or wAtual <> RS("tp_especialidade") Then
          wAtual = RS("tp_especialidade")
          If wAtual = "M" Then
             If Nvl(Request("T"),"S") = "P" Then
                ShowHTML "          <TR><TD colspan=2><b>Modalidades de ensino</b>:"
             Else
                ShowHTML "          <TR><TD colspan=2><b>Etapas / Modalidades de ensino</b>: "
             End If
             ShowHTML "<table id=""content"" style=""display: none"">"
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
  ShowHTML "            </table>"
  ShowHTML "          <TR><TD colspan=2 valign""middle"">"  
  ShowHTML "           <input id=""pesquisar"" type=""button"" name=""Botao"" value=""Pesquisar"" class=""botao"" onClick=""javascript: document.form1.Botao.disabled=true; document.form1.submit();"">"
  ShowHTML "        </tr>"
  ShowHTML "        </form>"
  

  if (Request("Z") > "") or (Request("Q") > "") or (Request("p_regional") > "") or (wIN > 0) then
    ShowHTML "        <tr><td colspan=2><table width=""570"" border=""0"" align=""left"" cellpadding=0 cellspacing=0>"
    ShowHTML "          <TR><TD><hr>"
    ShowHTML "          <TR><TD><b><a name=""resultado"">Resultado da Pesquisa:</a><br><br>"

    sql = "SELECT DISTINCT 0 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " & VbCrLf & _
          "       e.ds_logradouro,e.no_bairro,e.nr_cep, d.localizacao, 'S' publica, null ds_username, " & VbCrLf & _
          "       r.no_regiao " & VbCrLf & _
          "  from sbpi.Cliente                               d " & VbCrLf & _
          "       INNER      join sbpi.Tipo_Cliente          b ON (d.sq_tipo_cliente  = b.sq_tipo_cliente and " & VbCrLf & _
          "                                                      b.tipo             = 2 " & VbCrLf & _
          "                                                     ) " & VbCrLf & _
          "       LEFT OUTER join sbpi.Cliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
          "       INNER      join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " & VbCrLf & _
          " where d.publica = 'S' and 'P' <> '" & REQUEST("T") & "' and d.sq_cliente = " & Nvl(Request("p_regional"),0) & " " & VbCrLf
    If Request("Z") > "" Then sql = sql + "    and 0 < (select count(sq_cliente) from sbpi.Cliente where sq_cliente_pai = d.sq_cliente and localizacao = " & Request("Z") & ") " & VbCrLf End If
    If Request("q") > "" Then sql = sql + "    and 0 < (select count(sq_cliente) from sbpi.Cliente where sq_cliente_pai = d.sq_cliente and d.sq_tipo_cliente= " & Request("q") & ") " & VbCrLf End If
    If sql1 > ""         Then sql = sql + "    and 0 < (select count(sq_cliente) from sbpi.Especialidade_Cliente where sq_cliente_pai = d.sq_cliente and sq_especialidade in (" + Request("H") + ")) " & VbCrLf End If
    sql = sql & _
          "UNION " & VbCrLf & _
          "SELECT DISTINCT 0 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " & VbCrLf & _
          "       e.ds_logradouro,e.no_bairro,e.nr_cep, d.localizacao, 'S' publica, null ds_username, " & VbCrLf & _
          "       r.no_regiao " & VbCrLf & _
          "  from sbpi.Cliente                               d " & VbCrLf & _
          "       LEFT OUTER join sbpi.Cliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
          "       INNER      join sbpi.Cliente_Site          a ON (d.sq_cliente       = a.sq_cliente) " & VbCrLf & _
          "       INNER      join sbpi.Tipo_Cliente          b ON (d.sq_tipo_cliente  = b.sq_tipo_cliente and " & VbCrLf & _
          "                                                      b.tipo             = 2 " & VbCrLf & _
          "                                                     ) " & VbCrLf & _
          "       INNER      join sbpi.Cliente               c ON (d.sq_cliente       = c.sq_cliente_pai) " & VbCrLf & _
          "         INNER    join sbpi.Especialidade_cliente f ON (c.sq_cliente       = f.sq_cliente) " & VbCrLf & _
          "       INNER      join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " & VbCrLf & _
          " where 'P' <> '" & REQUEST("T") & "' " & VbCrLf
          
        
    If Request("Z") > ""          Then sql = sql + "    and d.localizacao    = " & Request("Z")          & VbCrLf End If
    If Request("p_regional") > "" Then sql = sql + "    and c.sq_cliente_pai = " & Request("p_regional") & VbCrLf End If
    If Request("q") > ""          Then sql = sql + "    and c.sq_tipo_cliente= " & Request("q")          & VbCrLf End If
    if sql1 > "" then
       sql = sql + "  and (f.sq_especialidade in (" + Request("H") + ")) " & VbCrLf
    end if
    sql = sql & _
          "UNION " & VbCrLf & _
          "SELECT DISTINCT 1 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " & VbCrLf & _ 
          "       f.endereco logradouro, null no_bairro, null nr_cep, d.localizacao, d.publica, " & VbCrLf & _ 
          "       d.ds_username, r.no_regiao " & VbCrLf & _
          "  from sbpi.Cliente                          d " & VbCrLf & _
          "       INNER join sbpi.Cliente_Particular    f ON (d.sq_cliente       = f.sq_cliente) " & VbCrLf & _
          "       LEFT  join sbpi.Especialidade_cliente g ON (d.sq_cliente       = g.sq_cliente) " & VbCrLf & _
          "       INNER join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " & VbCrLf & _
          " where f.situacao = 1 and d.publica = 'N' and 'S' <> '" & REQUEST("T") & "' " & VbCrLf
          
        
    If Request("Z") > ""          Then sql = sql + "    and d.localizacao    = " & Request("Z")          & VbCrLf End If
    If Request("p_regional") > "" Then sql = sql + "    and d.sq_regiao_adm  = " & Request("p_regional") & VbCrLf End If
    If Request("q") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("q")          & VbCrLf End If
    if sql1 > "" then
       sql = sql + "  and (g.sq_especialidade in (" + Request("H") + ")) " & VbCrLf
    end if
    sql = sql & _      
          "UNION " & VbCrLf & _ 
          "SELECT DISTINCT 1 Tipo, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " & VbCrLf & _ 
          "       e.ds_logradouro, e.no_bairro,e.nr_cep, d.localizacao, d.publica, " & VbCrLf & _ 
          "       d.ds_username, r.no_regiao " & VbCrLf & _ 
          "  from sbpi.Especialidade                      a " & VbCrLf & _ 
          "       INNER   join sbpi.Especialidade_cliente c ON (a.sq_especialidade = c.sq_especialidade) " & VbCrLf & _
          "       INNER   join sbpi.Cliente               d ON (c.sq_cliente       = d.sq_cliente) " & VbCrLf & _
          "         INNER join sbpi.Cliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
          "       INNER   join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " & VbCrLf & _
          " where d.publica = 'S' and 'P' <> '" & REQUEST("T") & "' " & VbCrLf
          
    If Request("Z") > ""          Then sql = sql + "    and d.localizacao    = " & Request("Z")          & VbCrLf End If
    If Request("p_regional") > "" Then sql = sql + "    and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf End If
    If Request("q") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("q")          & VbCrLf End If

    if sql1 > "" then
       sql = sql + "  and a.sq_especialidade in (" + Request("H") + ")" & VbCrLf        
    end if
    sql = sql + "ORDER BY tipo, d.ds_cliente " & VbCrLf
    
   ' Response.Write(sql)

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
          If RS("PUBLICA") = "S" Then
             'ShowHTML "            <TR valign=""top""><td><table border=0 cellpadding=0 cellspacing=0><tr align=""center""><TD bgcolor=""#00DD00"" width=""1%"" nowrap><font size=""1"">&nbsp;<B>REDE P�BLICA</B>&nbsp;</font></td></tr></table>"
             'ShowHTML "                <td align=""left""><font size=""2""><b>"
             If Not IsNull (RS("LN_INTERNET")) Then
                If inStr(lcase(RS("LN_INTERNET")),"http://") > 0 Then
                   ShowHTML "                <a href=""http://" & replace(RS("LN_INTERNET"),"http://","") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b>"
                Else
                  ShowHTML "                <a href=""" & RS("LN_INTERNET") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b>"
                End If
             Else
                ShowHTML RS("DS_CLIENTE") & "</b>"
             End If
          Else
             'ShowHTML "            <TR valign=""top""><td><table border=0 cellpadding=0 cellspacing=0><tr align=""center""><TD bgcolor=""#00CCFF"" width=""1%"" nowrap><font size=""1"">&nbsp;<B>REDE PRIVADA</B>&nbsp;</font></td></tr></table>"
             'ShowHTML "                <td align=""left""><font size=""2""><b>"
             ShowHTML "                <a href=""/modelos/ModPart/default.asp?EW=110&EF=sq_cliente=" & RS("SQ_CLIENTE") & "&CL=" & RS("SQ_CLIENTE") & """ target=""" & RS("DS_USERNAME") & """>" & RS("DS_CLIENTE") & "</a></b>"
          End If
          If RS("tipo") <> "0" Then
             ShowHTML "<br>Endere�o: " & RS("ds_logradouro")
             If Nvl(RS("no_bairro"),"nulo") <> "nulo" Then ShowHTML " - Bairro: " & RS("no_bairro") End If
             If Nvl(RS("nr_cep"),"nulo") <> "nulo" and len(RS("nr_cep")) > 5 Then ShowHTML " - CEP: " & RS("nr_cep") End If
             ShowHTML "<br>Regi�o Administrativa: " & RS("no_regiao")
             sql2 = "SELECT f.ds_especialidade " & VbCrLf & _ 
                  "  from sbpi.Especialidade_cliente e " & VbCrLf & _
                    "       INNER join sbpi.Especialidade f ON (e.sq_especialidade = f.sq_especialidade) " & VbCrLf & _
                    " WHERE e.sq_cliente = " & RS("sq_cliente") & " " & VbCrLf & _
                    "UNION " & VbCrLf & _
                    "SELECT f.ds_especialidade " & VbCrLf & _ 
                  "  from sbpi.Especialidade f " & VbCrLf & _
                    " WHERE upper(ds_especialidade) like '%PROFISSIONAL%' " & VbCrLf & _
                    "   and 0 < (select count(sq_cliente) from sbpi.Particular_curso where sq_cliente = " & RS("sq_cliente") & ") " & VbCrLf & _
                    "ORDER BY ds_especialidade "
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
       'MontaBarra "default.asp?EW=" & conWhatPrincipal & "&p_regional=" & Request("p_regional") & "&T=" & Request("T") & "&Q=" & Request("Q") & "&U=" & Request("U")& "&Z=" & Request("Z"), cDbl(RS.PageCount), cDbl(Request("htBT")), cDbl(Request("str2")), cDbl(RS.RecordCount)
       ShowHTML "</tr>"

    Else
       ShowHTML "            <TR><TD colspan=2><p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<b>Nenhuma ocorr�ncia encontrada para op��es acima."
    End If
    ShowHTML "          </table>"
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da P�gina de Pesquisa
REM =========================================================================

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
  ShowHTML "<left><FORM ACTION=""" & p_link &""" METHOD=""POST"" name=""Barra"">" & VbCrLf
  ShowHTML "<input type=""Hidden"" name=""str2"" value=""" & p_pagesize & """>" & VbCrLf
  ShowHTML "<input type=""Hidden"" name=""htBT"" value="""">" & VbCrLf
  For i = 1 to Request("H").Count
     ShowHTML "<input type=""Hidden"" name=""H"" value=""" & Request("H")(i) & """>" & VbCrLf
  Next
  If p_PageCount = p_AbsolutePage Then
     ShowHTML "<br><b>" & (p_RecordCount-((p_PageCount-1)*p_PageSize)) & "</b> resultados de <b>" & p_RecordCount & "</b>" & VbCrLf
  Else
     ShowHTML "<br><b>" & p_PageSize & "</b> resultados de <b>" & p_RecordCount & "</b>" & VbCrLf
  End If
  If p_PageSize < p_RecordCount Then
     ShowHTML "<br/>P�gina <b>" & p_AbsolutePage & "</b> de <b>" & p_PageCount & "</b>" & VbCrLf
     ShowHTML "<p class=""pag"">"
     If p_AbsolutePage > 1 Then
        ShowHTML "<A TITLE=""Primeira p�gina"" HREF=""javascript:pagina(1)""><span class=""primeira"">Primeira</span></A>"
        ShowHTML "<A TITLE=""P�gina anterior"" HREF=""javascript:pagina(" & p_AbsolutePage-1 & ")""><span class=""anterior"">Anterior</span></A>"
     Else
        ShowHTML "<span class=""primeiraOff"">Primeira</span>"
        ShowHTML "<span class=""anteriorOff"">Anterior</span>"
     End If
     If p_PageCount = p_AbsolutePage Then
        ShowHTML "<span class=""proximaOff"">Pr�xima</span>"
        ShowHTML "<span class=""ultimaOff"">�ltima</span>"
     Else
        ShowHTML "<A TITLE=""P�gina seguinte"" HREF=""javascript:pagina(" & p_AbsolutePage+1 & ")""><span class=""proxima"">Pr�xima</span></A>"
        ShowHTML "<A TITLE=""�ltima p�gina"" HREF=""javascript:pagina(" & p_PageCount & ")""><span class=""ultima"">�ltima</span></A>"
     End If
     ShowHTML "</p>"
  Else
     ShowHTML "<br/>P�gina <b>" & p_AbsolutePage & "</b> de <b>" & p_PageCount & "</b>" & VbCrLf
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
  'ShowPesquisa

  Set RS = nothing
  Set RS1 = nothing
  Set sobjConn  = nothing

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

%>