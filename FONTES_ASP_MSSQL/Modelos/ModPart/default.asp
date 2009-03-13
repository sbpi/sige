<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/Constants_ADO.inc"-->
<!--#INCLUDE VIRTUAL="/modelos/Constants.inc"-->
<!--#INCLUDE VIRTUAL="/esc.inc"-->
<!--#INCLUDE VIRTUAL="/Funcoes.asp"-->
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

Private dbms, sobjConn

Private sstrSN
Private sstrCL

sstrSN = "Default.Asp"
sstrCL = "Aluno.Asp"

Public sstrEF
Public sstrEW
Public sstrIN
Public CL

CL = int(Request.QueryString("CL"))
sstrEF = "sq_cliente=" & CL
sstrEW = Request.QueryString("EW")
sstrIN = int(Request.QueryString("IN"))
w_dir  = "Modelos/ModPart/"

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

sql = "SELECT a.sq_cliente, a.ds_cliente, b.profissional FROM escCliente a inner join escCliente_Particular b on (a.sq_cliente = b.sq_cliente) WHERE a." & sstrEF
RS.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly
ShowHTML "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
ShowHTML "<html xmlns=""http://www.w3.org/1999/xhtml"">"
ShowHTML "<head>"
ShowHTML "   <title>" & RS("DS_CLIENTE") & "</title>"
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
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=110&CL=" & replace(CL,"sq_cliente=","") & """ > Inicial </a> </li>"
'ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=113&CL=" & replace(CL,"sq_cliente=","") & """ id=""link1"">Credenciamento</a> </li>"
'ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=117&CL=" & replace(CL,"sq_cliente=","") & """ >Contatos</a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & """ id=""link2"">Oferta</a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=115&CL=" & replace(CL,"sq_cliente=","") & """ id=""link5"">Calendário</a>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & """ >Arquivos</a></li>"
'ShowHTML "	    </ul>"
'ShowHTML "      <li><a href=""newsletter.asp?ew=i"">Clique aqui para receber informativos da SEDF</a>"
ShowHTML "	    </ul>"
ShowHTML "	  </div>"
ShowHTML"  <div class=""clear""></div>"
ShowHTML "    <div id=""conteudo""><!-- Conteúdo -->"
ShowHTML "      <h2>" & RS("ds_cliente") & "</h2>"
Select Case sstrEW
  Case conWhatPrincipal   ShowHTML "      <h3><b> Inicial </b></h3>"
  'Case conWhatQuem       ShowHTML "      <h3><b>Credenciamento</b></h3>"
  Case conWhatExFale      ShowHTML "      <h3><b>Contato</b></h3>"
  Case conWhatExCalend    ShowHTML "      <h3><b>Calendário</b></h3>"
  'Case conWhatArquivo     ShowHTML "      <h3><b>Cursos Profissionalizantes</b></h3>"
  Case conWhatArquivo     ShowHTML "      <h3>Arquivos</h3>"
  Case conWhatExNoticia   ShowHTML "      <h3><b>Oferta</b></h3>"
  Case conWhatValida
     If sstrIN > 0 Then
        ShowHTML "      <li><b>Projeto</b></li>"
     Else
        ShowHTML "      <li><b>Alunos</b></li>"
     End If
End Select
ShowHTML "        <table width=""100%"" border=""0"">"
RS.Close

Main

ShowHTML "        </table>"
ShowHTML "</div>"
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
REM Monta a tela principal
REM -------------------------------------------------------------------------
Public Sub ShowPrincipal

  Dim sql, strNome
       
  sql = "SELECT b.ds_cliente, a.cnpj_escola, a.mantenedora, a.cnpj_executora, a.codinep, a.diretor, a.secretario, " & VbCrLf & _
        "       coalesce(convert(varchar,a.vencimento,103),'Sem informação') as vencimento, coalesce(convert(varchar,a.primeiro_credenc,103),'Sem informação') as primeiro_credenc, " & VbCrLf & _
        "       a.endereco, b.no_municipio, b.sg_uf, " & VbCrLf & _
        "       telefone_1, " & VbCrLf & _
        "       telefone_2, " & VbCrLf & _
        "       a.email_1, a.email_2, a.cep, " & VbCrLf & _
        "       fax, " & VbCrLf & _ 
        "       c.no_regiao " & VbCrLf & _ 
        "FROM esccliente_particular                 a " & VbCrLf & _
        "     INNER   JOIN escCliente               b on (a.sq_cliente   = b.sq_cliente) " & VbCrLf & _
        "       INNER JOIN escRegiao_Administrativa c on (b.sq_regiao_adm = c.sq_regiao_adm) " & VbCrLf & _
        " where a.sq_cliente = " & CL

  RS.Open sql, sobjConn, adOpenForwardOnly

  ShowHTML "IDENTIFICAÇÃO:<hr height=""1"">"

  If Not RS.EOF Then
    Do While Not RS.EOF

  ShowHTML "<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"">"
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
  'ShowHTML "  <tr valign=""top"">"
  'ShowHTML "    <td width=""10%"">"
  'ShowHTML "    <td align=""right"" width=""20%"" nowrap=""nowrap""><b>CNPJ da Mantenedora:"
  'ShowHTML "    <td align=""left"" width=""70%"">" & Nvl(RS("cnpj_executora"),"Sem informação") 
  'ShowHTML "  </tr>"
  'ShowHTML "  <tr valign=""top"">"
  'ShowHTML "    <td width=""10%"">"
  'ShowHTML "    <td align=""right"" width=""20%""><b>Código INEP:"
  'ShowHTML "    <td width=""70%"">" & Nvl(RS("codinep"),"Sem informação") 
  'ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" width=""30%""><b>Diretor(a):"
  ShowHTML "    <td align=""left"" width=""70%"">" & Nvl(RS("diretor"),"Sem informação") 
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" width=""30%""><b>Secretário(a):"
  ShowHTML "    <td align=""left"" width=""70%"">" & Nvl(RS("secretario"),"Sem informação")
  ShowHTML "  </tr>"
  ShowHTML "  <tr>"
  ShowHTML "    <td align=""right"" nowrap=""nowrap"" width=""30%"" ><b>Data do primeiro credenciamento:</b>"
  ShowHTML "    <td><table width=""1%"" border=0 cellpadding=0 cellspacing=0 bgcolor=""#DFDFDF""><tr><td width=""1%"" nowrap><font size=""2""><b>" & RS("primeiro_credenc") & "</b></td></tr></table>"
  ShowHTML "  </tr>"
  ShowHTML "  <tr>"
  ShowHTML "    <td align=""right"" nowrap=""nowrap"" width=""30%"" ><b>Validade credenciamento:</b>"
  ShowHTML "    <td><table width=""1%"" border=0 cellpadding=0 cellspacing=0 bgcolor=""#DFDFDF""><tr><td width=""1%"" nowrap><font size=""2""><b>" & RS("vencimento") & "</b></td></tr></table>"
  ShowHTML "  </tr>"
  ShowHTML "  </TABLE>"

  ShowHTML "<br>LOCALIZAÇÃO E CONTATOS:<hr height=""1"">"
  ShowHTML "<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"">"
    ShowHTML "  <tr valign=""top"">"
    ShowHTML "    <td width=""30%"" align=""right""><b>Endereço:"
    ShowHTML "    <td width=""70%"" align=""left"">" & RS("endereco")' & " - " & RS("cep")
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

  ShowHTML "  </TABLE>"
  
  RS.MoveNext

  Loop

  End If

  RS.Close
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Página Principal
REM =========================================================================

REM =========================================================================
REM Página Credenciamento
REM -------------------------------------------------------------------------
Public Sub ShowQuem
  Dim sql, strNome

  sql = "SELECT a.pasta, a.portaria, a.parecer_resolucao, a.ordem_servico, a.observacao, a.vencimento " & _ 
        "FROM esccliente_particular a INNER JOIN escCliente b on (a.sq_cliente = b.sq_cliente) " & _ 
        "where a.sq_cliente = " & CL       

  RS.Open sql, sobjConn, adOpenForwardOnly

  ShowHTML "<p align=""justify"">Credenciamento da instituição de ensino junto à Secretaria de Estado de Educação:</p>"

  If Not RS.EOF Then
    Do While Not RS.EOF

  ShowHTML "<table width=""95%"" border=""0"" cellspacing=""5"" cellpadding=""2"">"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" width=""20%""><b>Pasta:"
  ShowHTML "    <td align=""left"" width=""20%"">" & Nvl(RS("pasta"),"Sem informação")
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" width=""20%""><b>Portaria:"
  ShowHTML "    <td align=""left"" width=""20%"">" & Nvl(RS("portaria"),"Sem informação")
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" nowrap=""nowrap"" width=""20%"" ><b>Parecer da Resolução:</b>"
  ShowHTML "    <td align=""left"" width=""20%"">" & Nvl(RS("parecer_resolucao"),"Sem informação")
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" nowrap=""nowrap"" width=""20%"" ><b>Ordem de Serviço:</b>"
  ShowHTML "    <td align=""left"" width=""20%"">" & nvl(RS("ordem_servico"),"Sem informação")
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" nowrap=""nowrap""width=""20%"" ><b>Observação:</b>"
  ShowHTML "    <td align=""left"" width=""60%"" >" & Nvl(RS("observacao"),"Sem informação")
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" nowrap=""nowrap"" width=""20%"" ><b>Vencimento:</b>"
  ShowHTML "    <td align=""left"" width=""20%"" >" & RS("vencimento")
  ShowHTML "  </tr>"
  ShowHTML "</table>"

  RS.MoveNext

  Loop

  End If

  RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Página Credenciamento
REM =========================================================================

REM =========================================================================
REM Final da Página Contatos
REM -------------------------------------------------------------------------
Public Sub ShowExFale

  Dim sql, strNome

  sql = "SELECT a.endereco, b.no_municipio, b.sg_uf, " & VbCrLf & _
        "       '(' + substring(telefone_1,0,3)+')'+substring(telefone_1,3,110) as telefone_1, " & VbCrLf & _
        "       '(' + substring(telefone_2,0,3)+')'+substring(telefone_2,3,110) as telefone_2, " & VbCrLf & _
        "       a.email_1, a.email_2, a.cep, " & VbCrLf & _
        "       '(' + substring(fax,0,3)+')'+substring(fax,3,110) as fax, " & VbCrLf & _ 
        "       c.no_regiao " & VbCrLf & _ 
        "FROM esccliente_particular                 a " & VbCrLf & _
        "     INNER   JOIN escCliente               b on (a.sq_cliente   = b.sq_cliente) " & VbCrLf & _
        "       INNER JOIN escRegiao_Administrativa c on (b.sq_regiao_adm = c.sq_regiao_adm) " & VbCrLf & _
        " where a.sq_cliente = " & CL

  RS.Open sql, sobjConn, adOpenForwardOnly
  ShowHTML "<p align=""justify"">Informações relativas a endereço, telefones e contatos da instituição de ensino:</p>"

  If Not RS.EOF Then
    Do While Not RS.EOF

        ShowHTML "<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""2"">"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""5%"">"
        ShowHTML "    <td width=""25%"" align=""right""><b>Endereço:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("endereco")' & " - " & RS("cep")
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""5%"">"
        ShowHTML "    <td width=""25%"" align=""right""><b>Região Administrativa:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_regiao")
        ShowHTML "  </tr>"
        ShowHTML "  <tr>"
        ShowHTML "    <td width=""5%"">"
        ShowHTML "    <td width=""25%"" align=""right"" valign=""top""><b>Telefones:"
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
        ShowHTML "    <td width=""5%"">"
        ShowHTML "    <td width=""25%"" align=""right""><b>Fax:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("fax")
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""5%"">"
        ShowHTML "    <td width=""25%"" align=""right""><b>E-mails:"
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
        ShowHTML "</table>"

      RS.MoveNext

    Loop

  End If

  RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Página Contatos
REM =========================================================================

Public Sub ShowCalend
  Dim sql, strNome, wCont, wRede, wAno, wDatas(31,12,10), wCores(31,12,10), wImagem(31,12,10)
  Dim w_ini, w_fim, w_dias, w_inicial, w_final, w_mes, w_fim1, w_ini2
  Dim w_cliente, w_nome
  
  wAno = Year(Date())
  w_calendario = Request("w_calendario")
  strAction = "default.asp?EW=115&CL=" & CL & "&IN=" & sstrIN

     Response.Write "<BODY topmargin=0 bgcolor=""#FFFFFF"" BGPROPERTIES=""FIXED"" text=""#000000"" link=""#000000"" vlink=""#000000"" alink=""#FF0000"">"
     
   
        SQL = "select distinct(b.nome), a.sq_particular_calendario, b.ordem " & VbCrLf & _
              "  from escCalendario_Cliente a "  & VbCrLf & _ 
              "       left join escParticular_Calendario b on (a.sq_particular_calendario = b.sq_particular_calendario) " & VbCrLf & _
              " where sq_site_cliente = " & CL & " " & VbCrLf & _
              "   and year(dt_ocorrencia) = " & wAno & " " & VbCrLf & _
              "   and ativo = 'S' "  & VbCrLf & _
              "   and homologado = 'S' "  & VbCrLf & _
              " order by ordem"  & VbCrLf
        RS.CursorLocation = 3
        RS.Open sql, sobjConn, adOpenForwardOnly        
        
        If RS.RecordCount > 0 Then
          w_nome = RS("nome")
          ShowHTML "  <ul>"
          Do While Not RS.EOF
            ShowHTML "<li title=""" & RS("nome") & """><a href=""" & w_dir & strAction & "&w_calendario=" &RS("sq_particular_calendario")& "#calendario"">" & RS("nome")
          RS.MoveNext            
          Loop
          ShowHTML "  </ul>"
          ShowHTML "  <br/>"
          ShowHTML "  <br/>"
        Else
          ShowHTML "<p>A instituição não possui calendário(s) homologado(s).</p>"
        exit sub
        End If
        RS.Close         

     if (nvl(trim(Request("w_calendario")),"0") = "0") then 
         Response.End
     end if
  ShowHTML "<a name=""calendario"">"
  
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
        " where a.sq_cliente = " & CL & " AND d.sq_particular_calendario = " & Request("w_calendario")

  RS.Open sql, sobjConn, adOpenForwardOnly
  wObs = RS("observacao")
  RS.Close

  sql = "SELECT '' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'B' AS origem FROM escCalendario_base a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE YEAR(dt_ocorrencia)=" & wAno & " " & VbCrLf & _
        "UNION " & VbCrLf & _
        "SELECT '#99CCFF' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE sq_site_cliente = " & CL & "  AND YEAR(dt_ocorrencia)= " & wAno & " AND sq_particular_calendario = " & Request("w_calendario") & VbCrLf & _
        "ORDER BY data, origem desc, ocorrencia" & VbCrLf 
  RS.Open sql, sobjConn, adOpenForwardOnly
  
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
        ShowHTML "<tr align=""center""><td><strong>" & w_nome & "</strong><br/></td></tr>"
        ShowHTML "<tr><td><TABLE align=""center"" border=1>"
        ShowHTML "<tr><td align=""center""><TABLE border=0 cellSpacing=0>"
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
     RS.Close
     ' Recupera o ano letivo e o período de recesso
     sql = "  select * from " & VbCrLf & _
           "        (select dt_ocorrencia w_let_ini " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  wAno &  " and b.sigla = 'IA') " & VbCrLf & _
           "          where sq_site_cliente = " & CL & VbCrLf & _
           "            and a.sq_particular_calendario = " & Request("w_calendario") & VbCrLf & _
           "        ) a, " & VbCrLf & _
           "        (select dt_ocorrencia w_let_fim " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  wAno &  " and b.sigla = 'TA') " & VbCrLf & _
           "          where sq_site_cliente = " & CL & VbCrLf & _
           "            and a.sq_particular_calendario = " & Request("w_calendario") & VbCrLf & _
           "        ) b, " & VbCrLf & _
           "        (select dt_ocorrencia w_let2_ini " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  wAno &  " and b.sigla = 'I2') " & VbCrLf & _
           "          where sq_site_cliente = " & CL & VbCrLf & _
           "            and a.sq_particular_calendario = " & Request("w_calendario") & VbCrLf & _
           "        ) c, " & VbCrLf & _
           "        (select dt_ocorrencia w_let1_fim " & VbCrLf & _
           "           from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = " &  wAno &  " and b.sigla = 'T1') " & VbCrLf & _
           "          where sq_site_cliente = " & CL & VbCrLf & _
           "            and a.sq_particular_calendario = " & Request("w_calendario") & VbCrLf & _
           "        ) d " & VbCrLf

     RS.Open sql, sobjConn, adOpenForwardOnly
     
     Do While Not RS.EOF
       w_ini1 = RS("w_let_ini")
       w_fim1 = RS("w_let1_fim")
       w_ini2 = RS("w_let2_ini")
       w_fim2 = RS("w_let_fim")
       RS.MoveNext
     Loop
     RS.Close

     For w_cont = 1 to 2
        If w_cont = 1 Then
           w_ini = Month(w_ini1)
           w_fim = Month(w_fim1)
        Else
           w_ini = Month(w_ini2)
           w_fim = Month(w_fim2)
           If nvl(Request("w_imprime"),"N") = "N" Then
              ShowHTML "  <br><br>"
           Else
              ShowHTML "  <td align=""center"">"
           End If
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
        ' ShowHTML "  <TR>"
        ' ShowHTML "    <TD COLSPAN=""2"" HEIGHT=""1"" BGCOLOR=""#DAEABD"">"
        ' ShowHTML "  </TR>"
        
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
              RS.Open sql, sobjConn, adOpenForwardOnly
              w_final = RS("fim")
              RS.Close
           End If

           sql = "SELECT dbo.diasLetivos('" & w_inicial & "', '" & w_final & "'," & CL & ","& Request("w_calendario") &") qtd" & VbCrLf 
'           Response.Write sql
           RS.Open sql, sobjConn, adOpenForwardOnly
           If RS("qtd") > 0 Then
              ShowHTML "  <TR>"
              ShowHTML "    <TD>" & nomeMes(wCont)
              ShowHTML "    <TD ALIGN=""CENTER"">" & RS("qtd")
              w_mes = RS("qtd")

              RS.Close
              sql = "SELECT count(*) qtd FROM escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla ='SL') WHERE sq_particular_calendario = " &  Request("w_calendario") & " and sq_site_cliente = " & CL & "  AND dt_ocorrencia between convert(datetime, '" & w_inicial & "',103) and convert(datetime, '" & w_final & "',103)" & VbCrLf
              RS.Open sql, sobjConn, adOpenForwardOnly
              ShowHTML "    <TD ALIGN=""CENTER"">" & RS("qtd")
              w_mes = w_mes + RS("qtd")
              RS.Close
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
     
     ShowHTML "  </table>"
     ShowHTML "<tr><td colspan=""2""><TABLE width=""100%"" align=""center"" border=1 cellSpacing=1>"
     ShowHTML "<tr valign=""middle"" align=""center"">"
     If nvl(Request("w_imprime"),"N") = "N" Then
        ShowHTML "  <td><font size=1><b>LEGENDA</b></font>"
        ShowHTML "  <td><font size=1><b>FERIADOS</b></font>"
        ShowHTML "  <td><font size=1><b>RECESSOS</b></font>"
        ShowHTML "  <td><font size=1><b>SÁBADOS/DOMINGOS<br>LETIVOS<br>ESPECIAIS</b></font>"
     Else
        ShowHTML "  <td><font size=1><b>LEGENDA</b></font>"
        ShowHTML "  <td><font size=1><b>FERIADOS</b></font>"
     End If
     ShowHTML "</tr>"
     ShowHTML "<tr valign=""top"">"
     'Legenda
     sql = "SELECT * from escTipo_Data where abrangencia <> 'P' order by nome " & VbCrLf
     RS.Open sql, sobjConn, adOpenForwardOnly
     If Not RS.EOF Then
        If nvl(Request("w_imprime"),"N") = "N" Then
           ShowHTML "  <td><TABLE width=""90%"" align=""center"" border=0 cellSpacing=1>"
        Else
           ShowHTML "  <td rowspan=5><TABLE width=""90%"" align=""center"" border=0 cellSpacing=1>"
        End If
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
     
     'Recesso
     sql = "SELECT '#99CCFF' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla in ('RE','RA')) WHERE sq_site_cliente = " & CL & " AND YEAR(dt_ocorrencia)= " & wAno & " AND sq_particular_calendario = " & Request("w_calendario") &  VbCrLf & _
           "ORDER BY data, origem desc, ocorrencia" & VbCrLf
     RS.Open sql, sobjConn, adOpenForwardOnly
  
     If nvl(Request("w_imprime"),"N") = "S" Then
         ShowHTML "  <tr><td align=""center""><font size=1><b>RECESSOS</b></font><tr>"
      End If
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
     RS.Close
     
     ' Sábados letivos especiais
     wCont = 0
     sql = "SELECT '#99CCFF' as cor, b.imagem, a.dt_ocorrencia AS data, a.ds_ocorrencia AS ocorrencia, 'E' AS origem FROM escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla ='SL') WHERE sq_site_cliente = " & CL & "  AND YEAR(dt_ocorrencia)= " & wAno & " AND sq_particular_calendario = " & Request("w_calendario") &  VbCrLf & _
           "ORDER BY data, origem desc, ocorrencia" & VbCrLf 
     RS.Open sql, sobjConn, adOpenForwardOnly
       
     If nvl(Request("w_imprime"),"N") = "S" Then
         ShowHTML "  <tr><td align=""center""><font size=1><b>SÁBADOS/DOMINGOS<br>LETIVOS<br>ESPECIAIS</b></font><tr>"
      End If
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
     RS.Close
     
     ShowHTML " </TABLE>"

  End If
  ShowHTML "</TABLE>"
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
REM -------------------------------------------------------------------------
REM Final da Página Calendário
REM =========================================================================


Public Sub ShowNoticia

  Dim sql, strNome, wCont, wRede, wAno
  
  wAno =  Request.QueryString("wAno")
  
  If wAno = "" Then
     wAno = Year(Date())
  End If

     sql = "SELECT INFANTIL, FUNDAMENTAL, EJA, MEDIO, DISTANCIA, PROFISSIONAL " &_
           "FROM esccliente_particular a " & _
           "WHERE a.sq_cliente = " & CL
     dim modalidade
     
        

     RS.Open sql, sobjConn, adOpenForwardOnly
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

        ShowHTML "<p align=""justify"">Informações relativas às ofertas da instituição de ensino:</p>"

        ShowHTML "<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""2"">"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD align=""right"" width=""30%""><b>Ensino Infantil: </b></TD>"
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
        ShowHTML "    <TD align=""right""><b>Ensino de Jovens e Adultos: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadeeja & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD align=""right""><b>Ensino à Distância: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadedist & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR valign=""top"">"
        ShowHTML "    <TD align=""right""><b>Educação Profissional: </b></TD>"
        ShowHTML "    <TD align=""left"">"

        If cint(RS("PROFISSIONAL")) > 0 Then
          sql = "select sq_cliente, ds_curso " & VbCrLf & _ 
                "       from escParticular_Curso a " & VbCrLf & _
                " inner join escCurso            b on (a.sq_curso = b.sq_curso)" & VbCrLf & _
                " where sq_cliente = " & CL
         
          'response.write sql
          RS1.Open sql, sobjConn, adOpenForwardOnly
          ShowHTML "  <ul id=""cursos_ul"">"
          Do While Not RS1.EOF
          ShowHTML "    <li id=""cursos"">" & RS1("ds_curso") & "</li>"  
          RS1.MoveNext
          Loop
          ShowHTML "  </ul>"
          RS1.Close
        Else
          response.write modalidadeprof
        End If        
        
        ShowHTML "    </TD>"
        ShowHTML "  </TR>"
        ShowHTML "</TABLE>"
        wCont = wCont + 1
        RS.MoveNext
        Loop
        
     Else
        ShowHTML "    <p><BR>Até o presente momento, a Instituição não oferece cursos profissionalizantes. </p>"
     End If
     'Response.Write sql
     'Response.End()
     RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Página Modalidades de Ensino Ofertadas
REM =========================================================================

REM =========================================================================
REM Monta a tela de Arquivos
REM -------------------------------------------------------------------------
Public Sub ShowArquivo

  Dim sql, strNome, wCont, wAno, idEscola
  
  wAno =  Request.QueryString("wAno")
  
  If wAno = "" Then
     wAno = Year(Date())
  End If
  
  sql = "SELECT idescola from escCliente_Particular where sq_cliente = " &  CL
  RS.Open sql, sobjConn, adOpenForwardOnly
    idEscola = RS("idescola")  
  RS.Close

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

  
  wCont = 0
     
  If Not RS.EOF Then
    ShowHTML "<tr><td><ul>"
    wCont = 1
    Do While Not RS.EOF
      ShowHTML "  <li><a href=""http://www.prodatadf.com.br/recadastramento/uploads/esc_id" & idEscola & "/" & RS("ln_arquivo") & """ target=""_blank"">" & RS("ds_titulo") & "</a><br><div align=""justify""><font size=1>.:. " & RS("ds_arquivo") & "</div><br/>"
      RS.MoveNext
    Loop
    ShowHTML "</ul>"
  Else
     ShowHTML "  <TR><TD ALIGN=CENTER>Não há arquivos disponíveis no momento para o ano de " & wAno & " </TR>"
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
REM Corpo Principal do programa
REM -------------------------------------------------------------------------
Private Sub Main

  If Request.QueryString("EW") = conWhatSenha Then
      ShowSenha
  End If

  Select Case sstrEW
    Case conWhatPrincipal   ShowPrincipal ' Inicial
    'Case conWhatQuem        ShowQuem     ' Credenciamento
    'Case conWhatExFale      ShowExFale   ' Contatos
    Case conWhatExNoticia   ShowNoticia   ' Oferta
    Case conWhatExCalend    ShowCalend    ' Calendário
    Case conWhatArquivo     ShowArquivo   ' Cursos profissionais
  End Select
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>
