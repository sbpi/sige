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
ShowHTML "<div id=""container"">"
ShowHTML "  <div id=""cab"">"
ShowHTML "    <div id=""cabtopo"">"
ShowHTML "      <div id=""logoesq""><img src=""img/fundo_logoesq.gif"" border=0></div>"
ShowHTML "      <div id=""logodir""><a href=""http://www.se.df.gov.br""><img src=""img/fundo_logodir.jpg"" border=0></a></div>"
ShowHTML "    </div>"

ShowHTML "    <div id=""cabbase"">"
ShowHTML "      <div id=""busca"" valign=""center""><marquee width=""100%"" align=""middle""><font color=""white"" face=""Arial"" size=""2""><b>Brasília - Patrimônio Cultural da Humanidade</b></font></marquee></div> "
ShowHTML "    </div>"
ShowHTML "  </div>"
ShowHTML "  <div id=""corpo"">"
ShowHTML "    <div id=""menuesq"">"
ShowHTML "      <div id=""logomenuesq""><img src=""img/fundo_logomenuesq.gif"" border=0></div>"
ShowHTML "      <ul id=""menugov"">"
ShowHTML "      <script language=""JavaScript"" src=""inc/mm_menu.js"" type=""text/JavaScript""></script>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=110&CL=" & replace(CL,"sq_cliente=","") & """ > Inicial </a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=113&CL=" & replace(CL,"sq_cliente=","") & """ id=""link1"">Credenciamento</a> </li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=117&CL=" & replace(CL,"sq_cliente=","") & """ >Contatos</a></li>"
ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=114&CL=" & replace(CL,"sq_cliente=","") & """ id=""link2"">Oferta</a> </li>"
If RS("profissional") <> "0" Then ShowHTML "      <li><a href=""" & w_dir & "Default.asp?EW=143&CL=" & replace(CL,"sq_cliente=","") & """ >Cursos profissionais</a></li>"
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
  Case conWhatPrincipal   ShowHTML "      <li><b> Inicial </b></li>"
  Case conWhatQuem        ShowHTML "      <li><b>Credenciamento</b></li>"
  Case conWhatExFale      ShowHTML "      <li><b>Contato</b></li>"
  Case conWhatArquivo     ShowHTML "      <li><b>Cursos Profissionalizantes</b></li>"
  Case conWhatExNoticia   ShowHTML "      <li><b>Oferta</b></li>"
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
ShowHTML "        <table width=""100%"" border=""0"">"
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
REM Monta a tela principal
REM -------------------------------------------------------------------------
Public Sub ShowPrincipal

  Dim sql, strNome
       
  sql = "SELECT b.ds_cliente, a.cnpj_escola, a.mantenedora, a.cnpj_executora, a.codinep, a.diretor, a.secretario " & _
        "FROM esccliente_particular a " & _
        "INNER JOIN escCliente b on (a.sq_cliente = b.sq_cliente) where a.sq_cliente = " & CL  

  RS.Open sql, sobjConn, adOpenForwardOnly

  ShowHTML "<p align=""justify"">Informações relativas à identificação da instituição de ensino:</p>"

  If Not RS.EOF Then
    Do While Not RS.EOF

  ShowHTML "<table width=""95%"" border=""0"" cellspacing=""3"">"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td align=""right"" width=""20%"" nowrap=""nowrap""><b>Nome da Instiuição:"
  ShowHTML "    <td align=""left"" width=""70%"" >" & RS("ds_cliente") 
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td align=""right"" width=""20%"" nowrap=""nowrap""><b>CNPJ da Instituição:"
  ShowHTML "    <td align=""left"" width=""70%"" >" & Nvl(RS("cnpj_escola"),"---") 
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td align=""right"" width=""20%"" nowrap=""nowrap""><b>Mantenedora:"
  ShowHTML "    <td align=""left"" width=""70%"" >" & Nvl(RS("mantenedora"),"---") 
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td align=""right"" width=""20%"" nowrap=""nowrap""><b>CNPJ da Mantenedora:"
  ShowHTML "    <td align=""left"" width=""70%"">" & Nvl(RS("cnpj_executora"),"---") 
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td align=""right"" width=""20%""><b>Código INEP:"
  ShowHTML "    <td width=""70%"">" & Nvl(RS("codinep"),"---") 
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td align=""right"" width=""20%""><b>Diretor(a):"
  ShowHTML "    <td align=""left"" width=""70%"">" & Nvl(RS("diretor"),"---") 
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td width=""10%"">"
  ShowHTML "    <td align=""right"" width=""20%""><b>Secretário(a):"
  ShowHTML "    <td align=""left"" width=""70%"">" & Nvl(RS("secretario"),"---")
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
REM Final da Página Credenciamento
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

  ShowHTML "<table width=""95%"" border=""0"" cellspacing=""3"">"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" width=""20%""><b>Pasta:"
  ShowHTML "    <td align=""left"" width=""20%"">" & Nvl(RS("pasta"),"---")
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" width=""20%""><b>Portaria:"
  ShowHTML "    <td align=""left"" width=""20%"">" & Nvl(RS("portaria"),"---")
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" nowrap=""nowrap"" width=""20%"" ><b>Parecer da Resolução:</b>"
  ShowHTML "    <td align=""left"" width=""20%"">" & Nvl(RS("parecer_resolucao"),"---")
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" nowrap=""nowrap"" width=""20%"" ><b>Ordem de Serviço:</b>"
  ShowHTML "    <td align=""left"" width=""20%"">" & nvl(RS("ordem_servico"),"---")
  ShowHTML "  </tr>"
  ShowHTML "  <tr valign=""top"">"
  ShowHTML "    <td align=""right"" nowrap=""nowrap""width=""20%"" ><b>Observação:</b>"
  ShowHTML "    <td align=""left"" width=""60%"" >" & Nvl(RS("observacao"),"---")
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

  sql = "SELECT a.endereco, b.no_municipio, b.sg_uf, '(' + substring(telefone_1,0,3)+')'+substring(telefone_1,3,110) as telefone_1, '(' + substring(telefone_2,0,3)+')'+substring(telefone_2,3,110) as telefone_2, a.email_1, a.email_2, a.cep, '(' + substring(fax,0,3)+')'+substring(fax,3,110) as fax " & _ 
        "FROM esccliente_particular a " & _
        "INNER JOIN escCliente b on (a.sq_cliente = b.sq_cliente) where a.sq_cliente = " & CL

  RS.Open sql, sobjConn, adOpenForwardOnly
  ShowHTML "<p align=""justify"">Informações relativas a endereço, telefones e contatos da instituição de ensino:</p>"

  If Not RS.EOF Then
    Do While Not RS.EOF

        ShowHTML "<table border=""0"" cellspacing=""3"">"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <td width=""20%"" align=""right""><b>Endereço:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("endereco")' & " - " & RS("cep")
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <td width=""20%"" align=""right""><b>Município:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("no_municipio") & "-" & RS("sg_uf")
        ShowHTML "  </tr>"
        ShowHTML "  <tr>"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <td width=""20%"" align=""right"" valign=""top""><b>Telefones:"
        If(Nvl(RS("telefone_1"),"") <> "" and Nvl(RS("telefone_2"),"") <> "") Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("telefone_1") & " / " & RS("telefone_2")
        ElseIf Nvl(RS("telefone_1"),"") <> "" and Nvl(RS("telefone_2"),"")  = "" Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("telefone_1")
        ElseIf(Nvl(RS("telefone_2"),"") <> "" and Nvl(RS("telefone_1"),"")  = "") Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("telefone_2")
        Else
          ShowHTML "    <td width=""70%"" align=""left"">---"
        End If 
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <td width=""20%"" align=""right""><b>Fax:"
        ShowHTML "    <td width=""70%"" align=""left"">" & RS("fax")
        ShowHTML "  </tr>"
        ShowHTML "  <tr valign=""top"">"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <td width=""20%"" align=""right""><b>E-mails:"
        If Nvl(RS("email_1"),"")  <> "" and Nvl(RS("email_2"),"") <> "" Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("email_1") & " / " & RS("email_2")
        ElseIf(Nvl(RS("email_1"),"")  <> "" and Nvl(RS("email_2"),"")  = "") Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("email_1")
        ElseIf(Nvl(RS("email_2"),"")  <> "" and Nvl(RS("email_1"),"")  = "") Then
          ShowHTML "    <td width=""70%"" align=""left""> " & RS("email_2")
        Else
          ShowHTML "    <td width=""70%"" align=""left"">---"
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

REM =========================================================================
REM Final da Página Modalidades de Ensino Ofertadas
REM -------------------------------------------------------------------------
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

        ShowHTML "<table border=""0"" cellspacing=""3"">"
        ShowHTML "  <TR>"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <TD align=""right""><b>Ensino Infantil: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadeinf & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <TD align=""right""><b>Ensino Fundamental: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadefund & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <TD align=""right""><b>Ensino Médio: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidademed & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <TD align=""right""><b>Ensino de Jovens e Adultos: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadeeja & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <TD align=""right""><b>EAD (Ensino a Distância): </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadedist & "</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
        ShowHTML "    <td width=""10%"">"
        ShowHTML "    <TD align=""right""><b>(EP) Educação Profissional: </b></TD>"
        ShowHTML "    <TD align ""left"">" & modalidadeprof & "</TD>"
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
REM Monta a tela de Cursos Profissionais
REM -------------------------------------------------------------------------
Public Sub ShowArquivo

  Dim sql, strNome, wCont, wRede, wAno
  
  wAno =  Request.QueryString("wAno")
  
  If wAno = "" Then
     wAno = Year(Date())
  End If

  If sstrIN = 0 then
           
     sql = "SELECT a.curso, a.parecer, a.portaria, a.observacao FROM " & _
           "escParticular_curso a " & _ 
           "WHERE sq_cliente = " & CL
           
     RS.Open sql, sobjConn, adOpenForwardOnly
     
     wCont = 0
     If Not RS.EOF Then
        ShowHTML "<p align=""justify"">Cursos profissionalizantes oferecidos pela instituição de ensino:</p>"

        ShowHTML "<TABLE width=""100%"" border=0 cellSpacing=3>"
        ShowHTML "  <TR valign=""top"" align=""center"">"
        ShowHTML "    <TD width=""2%"" >&nbsp;</TD>"
        ShowHTML "    <TD width=""16%""><b>Curso</TD>"
        ShowHTML "    <TD width=""12%""><b>Parecer</TD>"
        ShowHTML "    <TD width=""12%""><b>Portaria</TD>"
        ShowHTML "    <TD width=""50%""><b>Observação</TD>"
        ShowHTML "  </TR>"
        ShowHTML "  <TR>"
        ShowHTML "    <TD COLSPAN=""6"" HEIGHT=""1"" BGCOLOR=""#DAEABD""></TD>"
        ShowHTML "  </TR>"
        wCont = 1
        Do While Not RS.EOF
           ShowHTML "  <TR valign=""top"">"
           ShowHTML "    <TD>&nbsp;</TD>"
           ShowHTML "    <TD align=""left"">" & RS("curso") & "</TD>"
           ShowHTML "    <TD align=""left"">" & RS("parecer") & "</TD>"
           ShowHTML "    <TD align=""left"">" & RS("portaria") & "</TD>"
           ShowHTML "    <TD align=""left"">" & RS("observacao") & "</TD>"
           ShowHTML "  </TR>"
           wCont = wCont + 1
           RS.MoveNext
        Loop
        
     Else
        ShowHTML "<p align=""justify"">Nenhum curso profissionalizante registrado para esta instituição de ensino:</p>"
     End If
     RS.Close    
     ShowHTML "    </TABLE>"
 
  Else
     sql = "SELECT * " & _
           "From escNoticia_Cliente " & _
           "WHERE sq_noticia = " & sstrIN & " "
     RS.Open sql, sobjConn, adOpenForwardOnly

     ShowHTML "<tr><td><p align=""center""><font size=""4""><b>" & RS("ds_titulo") & "</b></font><p align=""left"">"
     ShowHTML replace(server.HTMLEncode(RS("ds_noticia")),chr(13)&chr(10), "<br>")
     ShowHTML "<p><center><img src=""Img/bt_voltar.gif"" border=0 valign=""center"" onClick=""history.go(-1)"" alt=""Volta"">"
     RS.Close
  End If
End Sub
REM -------------------------------------------------------------------------
REM Monta a tela de Cursos Profissionais
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
    Case conWhatQuem        ShowQuem      ' Credenciamento
    Case conWhatExFale      ShowExFale    ' Contatos
    Case conWhatExNoticia   ShowNoticia   ' Oferta
    Case conWhatArquivo     ShowArquivo   ' Cursos profissionais
  End Select
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>
