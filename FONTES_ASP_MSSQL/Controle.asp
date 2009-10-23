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
REM  /Controle.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: Controle geral do site
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 21/01/2004 12:10
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Copyright: 2004 by SBPI Consultoria Ltda
REM -------------------------------------------------------------------------


  Private w_EA
  Private w_IN
  Private w_EF
  Private CL, DBMS, SQL, RS, RS1, RS2
  Private w_Data, w_pagina, w_ds_arquivo
  Private p_tipo, p_campos, p_layout

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
     p_layout = Nvl(Request("p_layout"),"PORTRAIT")
     p_campos = Request("p_campos")
     p_escola_particular = Request("p_escola_particular")
     p_calendario        = Request("p_calendario")
  End If

  w_Data = Mid(100+Day(Date()),2,2) & "/" & Mid(100+Month(Date()),2,2) & "/" &Year(Date())
  w_Pagina = ExtractFileName(Request.ServerVariables("SCRIPT_NAME")) & "?w_ew="
  AbreSessaoManut
  If Session("Username") > "" and InStr("LOGON,VALIDA",w_ew) = 0 Then
     Main
  ElseIf w_ew = "LOGON" Then
     Logon
  ElseIf w_ew = "VALIDA" Then
     Valida
  ElseIf w_ew = "ENCERRAR" Then
     Session.Abandon()
     ShowHTML "<HTML>" & chr(13) & chr(10)
     ShowHTML "<BODY>" & chr(13) & chr(10)
     ShowHTML "<SCRIPT LANGUAGE='JAVASCRIPT'>" & chr(13) & chr(10)
     ShowHTML "  top.location.href='Controle.asp?w_ew=logon';" & chr(13) & chr(10)
     ShowHTML "</SCRIPT>" & chr(13) & chr(10)
     ShowHTML "</BODY>" & chr(13) & chr(10)
     ShowHTML "</HTML>" & chr(13) & chr(10)
  Else
     ShowHTML "<HTML>" & chr(13) & chr(10)
     ShowHTML "<BODY>" & chr(13) & chr(10)
     ShowHTML "<SCRIPT LANGUAGE='JAVASCRIPT'>" & chr(13) & chr(10)
     ShowHTML "  top.location.href='Controle.asp?w_ew=logon';" & chr(13) & chr(10)
     ShowHTML "</SCRIPT>" & chr(13) & chr(10)
     ShowHTML "</BODY>" & chr(13) & chr(10)
     ShowHTML "</HTML>" & chr(13) & chr(10)
  End If
  FechaSessao

  Set p_tipo         = Nothing
  Set w_ds_diretorio = Nothing
  Set w_EA           = Nothing
  Set w_IN           = Nothing
  Set w_EF           = Nothing
  Set CL             = Nothing
  Set SQL            = Nothing
  Set RS             = Nothing
  Set RS1            = Nothing
  Set RS2            = Nothing
  Set DBMS           = Nothing
  Set w_Data         = Nothing
  Set w_Pagina       = Nothing

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
  ShowHTML "    <frame name=""Menu"" src=""Controle.asp?CL=" & CL & "&w_ew=SHOWMENU"" scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
  ShowHTML "    <frame name=""Body"" src=""Controle.asp?CL=" & CL & "&w_ee=1&w_ea=A&w_ew=" & conRelEscolas & """ scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
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
   ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
   ShowHTML "<TITLE>" & "Controle Central" & "</TITLE>"
   ShowHTML "<style>"
   ShowHTML "<// a { color: ""#000000""; text-decoration: ""none""; } "
   ShowHTML "    a:hover { color:""#000000""; text-decoration: ""underline""; }"
   ShowHTML "    .SS{text-decoration:none;} "
   ShowHTML "    .SS:HOVER{text-decoration: ""underline"";} "
   ShowHTML "//></style>"
   ShowHTML "</HEAD>"
   ShowHTML "<BASEFONT FACE=""Verdana, Helvetica, Sans-Serif"" SIZE=""2"">"
   ShowHTML "<BODY topmargin=0 bgcolor=""#FFFFFF"" BACKGROUND=""img/background.gif"" BGPROPERTIES=""FIXED"" text=""#000000"" link=""#000000"" vlink=""#000000"" alink=""#FF0000""> "
   ShowHTML "  <table style=""background-image:url(img/background.gif)"" border=0 cellpadding=0 height=""80"" width=""100%""><tr><td nowrap><font size=1><b>"
   ShowHTML "  <TR><TD align=""center""><font size=2><b>Atualização</TD></TR>"
   ShowHTML "  <TR><TD align=""center""><br><font size=1>Usuário:<b>" & Session("username") & "</TD></TR>"
   ShowHTML "  <TR><TD><font size=1><br>"

   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""_blank"" CLASS=""SS"" HREF=""manuais/operacao/"" Title=""Exibe manual de operação do SIGE-WEB"">Manual SIGE-WEB</A><BR>"
   If Session("username") = "ADMINISTRATIVO" or Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=A&w_ew=" & conWhatAdmin & "&w_ee=1"" Title=""Preenche dados administrativos da unidade de ensino!"">Administrativo</A><BR>"
   End If

   If Session("username") = "SEDF" or Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatDocumento & "&w_ee=1"" Title=""Cadastra arquivos para download!"">Arquivos (<i>download</i>)</A><br> "
   End If

   If Session("username") = "GESTORVERSAO" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatSGE & "&w_ee=1"" Title=""Tela para gestão de versões!"">Sistemas</A><br> "
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=COMPONENTE&w_ee=1"" Title=""Tela para gestão de versões!"">Componentes</A><br> "
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=VERSAO&w_ee=1"" Title=""Tela para gestão de versões!"">Versões</A><br> "
   End If

   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=base&w_ee=1"" Title=""Cadastra calendário oficial!"">Calendário Base</A><br>"
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatCalendario & "&w_ee=1"" Title=""Cadastra datas especiais da rede de ensino!"">Calendário da rede</A><br>"
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""escRegional.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=L&w_ew=ligacao&w_ee=1"" Title=""Permite alterar a regional e o tipo da escola!"">Dados da escola</A><BR>"
   End If

   If Session("username") = "IMPRENSA" or Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Mensagem.asp?CL=" & CL & "&O=L&par=inicial&TP=e-Mail automático"" Title=""Permite o envio de e-mails através do sistema!"">Envio de e-Mail</A><br>"
   End If

   If Session("username") <> "ADMINISTRATIVO" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conRelEscolas & "&w_ee=1"" Title=""Pesquisa unidades de ensino!"">Escolas</A><br> "
   End If
   
   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=escPart&&w_ee=1"" Title=""Pesquisa unidades de ensino da rede privada"">Escolas Particulares</A><br>"
   End If
   
   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=escPartHomolog&&w_ee=1"" Title=""Pesquisa unidades de ensino da rede privada"">Homologação de Calendário</A><br>"
   End If

   If Session("username") = "IMPRENSA" or Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=newsletter&w_ee=1"" Title=""Acessa a lista de distribuição da newsletter!"">Informativo</A><br>"
   End If

   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatEspecialidade & "&w_ee=1"" Title=""Cadastra modalidades de ensino oferecidas pela rede de ensino!"">Modalidades de ensino</A><br>"
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatMensagem & "&w_ee=1"" Title=""Cadastra mensagens dirigidas seus alunos!"">Mensagens a alunos</A><br>"
   End If

   If Session("username") = "SEDF" OR Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & conWhatNotCliente & "&w_ee=1"" Title=""Cadastra notícias da rede de ensino!"">Notícias</A><br>"
   End If

   If Session("username") = "SEDF" or Session("username") = "SBPI" then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""controle.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=L&w_ew=REDEPART&w_ee=1"" Title=""Atualiza base de dados da rede particular!"">Rede particular</A><BR>"
   End If

   If Session("username") = "SEDF" or Session("username") = "SBPI" then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""controle.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=L&w_ew=TIPOCLIENTE&w_ee=1"" Title=""Tipos de Cliente!"">Tipo de Instituição</A><BR>"
   End If

   ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=A&w_ew=" & conWhatCliente & "&w_ee=1"" Title=""Altera a senha de acesso!"">Senha</A><BR>"

   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=senhaesp&w_ee=1"" Title=""Exibe a lista de senhas de regionais e outros usuários (menos escolas)!"">Senhas especiais</A><br>"
   End If

   If Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=VerifArq&w_ee=1"" Title=""Verificação de Arquivos"">Verif. Arquivos</A><BR>"
   End If
   
   If Session("username") = "SEDF" or Session("username") = "SBPI" Then
      ShowHTML "    <img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""GR_Log.asp?CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_ea=A&w_ew=Gerencial&w_ee=1"" Title=""Consulta ao Log!"">Log</A><BR>"
   End If
      
   ShowHTML "    <br><img src=""" & w_imagem & """ border=0 align=""center""> <A TARGET=""Body"" CLASS=""SS"" HREF=""Controle.asp?CL=" & CL & "&w_ea=L&w_ew=" & encerrar & "&w_ee=1"" Title=""Encerra a utilização da aplicação!"">Encerrar</A> "
   ShowHTML "</TABLE>"
   ShowHTML "</BODY>"
   ShowHTML "</HTML>"
   
   Set w_imagem = Nothing
   
End Sub

REM =========================================================================
REM Cadastro de arquivos
REM -------------------------------------------------------------------------
Sub GetDocumento
  Dim w_chave, w_ds_titulo, w_in_ativo, w_ds_arquivo, w_ln_arquivo
  Dim w_dt_arquivo, w_in_destinatario, w_nr_ordem, w_pasta
  
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
     w_pasta           = Request("w_pasta")    
     w_nr_ordem        = Request("w_nr_ordem")    
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then
        SQL = "select * from escCliente_Arquivo where sq_site_cliente = 0 order by nr_ordem, ltrim(upper(ds_titulo))"
     Else
        SQL = "select * from escCliente_Arquivo where " & replace(CL,"sq_cliente","sq_site_cliente") & " order by in_ativo, nr_ordem"
     End If
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select b.ds_diretorio, a.* from escCliente_Arquivo a inner join escCliente_Site b on (a.sq_site_cliente = b.sq_site_cliente) where a.sq_arquivo = " & w_chave
     ConectaBD SQL
     w_dt_arquivo      = FormataDataEdicao(RS("dt_arquivo"))
     w_ds_titulo       = RS("ds_titulo")
     w_in_ativo        = RS("in_ativo")
     w_ds_arquivo      = RS("ds_arquivo")
     w_ln_arquivo      = RS("ln_arquivo")
     w_in_destinatario = RS("in_destinatario")
     w_pasta           = RS("pasta")
     w_nr_ordem        = RS("nr_ordem")
     w_ds_diretorio    = RS("ds_diretorio")
     DesconectaBD
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  If InStr("IAEP",w_ea) > 0 Then
    ScriptOpen "JavaScript"
    FormataDATA
    ValidateOpen "Validacao"
    Validate "w_ds_titulo" , "Título"      , "", "1", "2", "50"  , "1", "1"
    Validate "w_ds_arquivo", "Descrição"   , "", "1", "2", "200" , "1", "1"
    Validate "w_pasta" , "Pasta"           , "SELECT" , "1" , "1" , "10"   , "1" , "1"
    If w_ea = "I" Then
      Validate "w_ln_arquivo", "Link"        , "", "1",  "2", "200" , "1", "1"
    End If
    Validate "w_nr_ordem"  , "Nr. de ordem", "", "1", "1", "4"   , "1", "0123546789"
    ShowHTML " if (theForm.w_ln_arquivo.value > """"){"
    ShowHTML "    if((theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.DLL')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.SH')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.BAT')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.EXE')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.ASP')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf('.PHP')!=-1)) {"
    ShowHTML "       alert('Tipo de arquivo não permitido!');"
    ShowHTML "       theForm.w_ln_arquivo.focus(); "
    ShowHTML "       return false;"
    ShowHTML "    }"
    ShowHTML "  }"           
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
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de arquivos da rede de ensino</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Ordem</font></td>"
    ShowHTML "          <td><font size=""1""><b>Arquivo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Descrição</font></td>"
    ShowHTML "          <td><font size=""1""><b>Ativo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
       ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=5 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
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
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_chave=" & RS("sq_arquivo") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
       ShowHTML "          <td valign=""top""><font size=""1""><b>Cadastramento:</b><br><input type=""text"" name=""w_dt_arquivo"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(FormatDateTime(date(),2)) & """  onKeyDown=""FormataData(this,event);"" ONMOUSEOVER=""popup('Data de inclusão do arquivo.','white')""; ONMOUSEOUT=""kill()""></td>"
    Else
       ShowHTML "          <td valign=""top""><font size=""1""><b>Última alteração:</b><br><input type=""text"" name=""w_dt_arquivo"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(FormatDateTime(w_dt_arquivo,2)) & """ onKeyDown=""FormataData(this,event);"" ONMOUSEOVER=""popup('Data da última alteração deste arquivo.','white')""; ONMOUSEOUT=""kill()""></td>"
    End If
    ShowHTML "        </table>"
    ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>D</u>escrição:</b><br><textarea " & w_Disabled & " accesskey=""D"" name=""w_ds_arquivo"" class=""STI"" ROWS=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva a finalidade do arquivo.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_arquivo & "</TEXTAREA></td>"
    ShowHTML "      </tr>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td><font size=""1""><b>Exibir na <u>p</u>asta:</b><br><select " & w_Disabled & " accesskey=""P"" name=""w_pasta"" class=""STS"" SIZE=""1"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a pasta à qual o arquivo destina-se.','white')""; ONMOUSEOUT=""kill()"">"
    If w_pasta = "1" Then
       ShowHTML "            <OPTION VALUE="""" >Selecione uma opção...<OPTION SELECTED VALUE=""1"">Meses<OPTION VALUE=""2"">Formulários<OPTION VALUE=""3"">Diversos"
    ElseIf w_pasta = "2" Then
       ShowHTML "            <OPTION VALUE="""" >Selecione uma opção...<OPTION VALUE=""1"">Meses<OPTION SELECTED VALUE=""2"">Formulários<OPTION VALUE=""3"">Diversos"
    ElseIf w_pasta = "3" Then
       ShowHTML "            <OPTION VALUE="""" >Selecione uma opção...<OPTION VALUE=""1"">Meses<OPTION VALUE=""2"">Formulários<OPTION SELECTED VALUE=""3"">Diversos"
    Else
       ShowHTML "            <OPTION SELECTED VALUE="""" >Selecione uma opção...<OPTION VALUE=""1"">Meses<OPTION VALUE=""2"">Formulários<OPTION VALUE=""3"">Diversos"
    End If
    ShowHTML "            </SELECTED></TD>"
    ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>L</u>ink:</b><br><input " & w_Disabled & " accesskey=""L"" type=""file"" name=""w_ln_arquivo"" class=""STI"" SIZE=""80"" MAXLENGTH=""100"" VALUE="""" ONMOUSEOVER=""popup('OBRIGATÓRIO. Clique no botão ao lado para localizar o arquivo. Ele será transferido automaticamente para o servidor.','white')""; ONMOUSEOUT=""kill()"">"
    If w_ln_arquivo > "" Then
       ShowHTML "              <b><a class=""SS"" href=""" & w_ds_diretorio & "/" & w_ln_arquivo & """ target=""_blank"" title=""Clique para exibir o arquivo atual."">Exibir</a></b>"
    End If
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td><font size=""1""><b><u>N</u>r. de ordem:</b><br><input " & w_Disabled & " accesskey=""N"" type=""text"" name=""w_nr_ordem"" class=""STI"" SIZE=""4"" MAXLENGTH=""4"" VALUE=""" & w_nr_ordem & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a posição em que este arquivo deve aparecer na lista de arquivos disponíveis. Ex: 1, 2, 3 etc.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "          <td><font size=""1""><b><u>D</u>estinatários:</b><br><select " & w_Disabled & " accesskey=""D"" name=""w_in_destinatario"" class=""STS"" SIZE=""1"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o público ao qual o arquivo destina-se.','white')""; ONMOUSEOUT=""kill()"">"
    If w_in_destinatario = "A" Then
       ShowHTML "            <OPTION VALUE=""A"" SELECTED>Apenas alunos <OPTION VALUE=""P"">Apenas professores <OPTION VALUE=""T"">Professores e alunos <OPTION VALUE=""E"">Escola"
    ElseIf w_in_destinatario = "E" Then
       ShowHTML "            <OPTION VALUE=""A"">Apenas alunos <OPTION VALUE=""P"">Apenas professores <OPTION VALUE=""T"">Professores e alunos <OPTION VALUE=""E"" SELECTED>Escola"
    ElseIf w_in_destinatario = "P" Then
       ShowHTML "            <OPTION VALUE=""A"">Apenas alunos <OPTION VALUE=""P"" SELECTED>Apenas professores <OPTION VALUE=""T"">Professores e alunos <OPTION VALUE=""E"">Escola"
    Else
       ShowHTML "            <OPTION VALUE=""A"">Apenas alunos <OPTION VALUE=""P"">Apenas professores <OPTION VALUE=""T"" SELECTED>Professores e alunos <OPTION VALUE=""E"">Escola"
    End If
    ShowHTML "            </SELECTED></TD>"
    MontaRadioSN "<b>Exibir no site?</b>", w_in_ativo, "w_in_ativo"
    ShowHTML "        </table>"
    ShowHTML "      <tr>"
    ShowHTML "      <tr><td align=""center"" colspan=4><hr>"
    If w_ea = "E" Then
       ShowHTML "   <input class=""STB"" type=""submit"" name=""Botao"" value=""Excluir"" onClick=""return confirm('Confirma a exclusão do registro?');"">"    
       'ShowHTML "   <input class=""STB"" type=""submit"" name=""Botao"" value=""Excluir"">"
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

  Dim w_chave, w_cont
  
  If w_ea = "L" Then
     SQL = "select a.sq_especialidade, a.ds_especialidade, b.sq_codigo_cli " & VbCrLf & _
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
  ShowHTML "      <tr><td><font size=1>Selecione as áreas de atuação oferecidas pela rede de ensino.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  While Not RS.EOF
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

  Set w_chave = Nothing
  Set w_cont  = Nothing


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
     w_sq_cliente       = RS("sq_cliente")
     w_nr_cnpj          = RS("nr_cnpj")
     w_tp_registro      = RS("tp_registro")
     w_ds_ato           = RS("ds_ato")
     w_nr_ato           = RS("nr_ato")
     w_dt_ato           = FormataDataEdicao(RS("dt_ato"))
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
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "Javascript"
  CheckBranco
  FormataDATA
  FormataCEP
  ValidateOpen "Validacao"
  If w_ea = "A" Then
     Validate "w_ds_logradouro"    , "Logradouro"              , "1"    , "1" , "4"  , "60" , "1" , "1"
     Validate "w_no_bairro"        , "Bairro"                  , "1"    , ""  , "2"  , "30" , "1" , "1"
     Validate "w_nr_cep"           , "CEP"                     , "1"    , "1" , "9"  , "9"  , ""  , "0123456789-"
     Validate "w_no_diretor"       , "Nome do(a) Diretor(a)"   , "1"    , ""  , "2"  , "40" , "1" , "1"
     Validate "w_no_secretario"    , "Nome do(a) Secretário(a)", "1"    , ""  , "2"  , "40" , "1" , "1"
     Validate "w_nr_cnpj"          , "CNPJ"                    , "1"    , ""  , "14" , "14" , ""  , "012356789"
     Validate "w_tp_registro"      , "Tipo do registro"        , "1"    , ""  , "6"  , "60" , "1" , "1"
     Validate "w_ds_ato"           , "Ato"                     , "1"    , ""  , "1"  , "30" , "1" , "1"
     Validate "w_nr_ato"           , "Número"                  , "1"    , ""  , "1"  , "15" , "1" , "1"
     Validate "w_dt_ato"           , "Data"                    , "DATA" , ""  , "10" , "10" , "1" , "1"
     Validate "w_ds_orgao"         , "Órgão Emissor"           , "1"    , ""  , "1"  , "30" , "1" , "1"
     Validate "w_no_contato"       , "Nome"                    , "1"    , "1" , "2"  , "35" , "1" , "1"
     Validate "w_ds_email_contato" , "e-Mail"                  , "1"    , ""  , "6"  , "60" , "1" , "1"
     Validate "w_nr_fone_contato"  , "Telefone"                , "1"    , "1" , "6"  , "20" , "1" , "1"
     Validate "w_nr_fax_contato"   , "Fax"                     , "1"    , ""  , "6"  , "20" , "1" , "1"
  End If
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.Form.w_ds_logradouro.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Atualização de dados adicionais da rede de ensino</FONT></B>"
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
  ShowHTML "      <tr><td><font size=1>Informe o endereço da rede de ensino, a ser exibido na seção ""Quem Somos"" do site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>L</u>ogradouro:</b><br><INPUT ACCESSKEY=""L"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_logradouro"" size=""60"" maxlength=""60"" value=""" & w_ds_logradouro & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o logradouro de funcionamento da rede de ensino.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b><u>B</u>airro:</b><br><INPUT ACCESSKEY=""B"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_no_bairro"" size=""30"" maxlength=""30"" value=""" & w_no_bairro & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o bairro de funcionamento da rede de ensino.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b>C<u>E</u>P:</b><br><INPUT ACCESSKEY=""C"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_cep"" size=""9"" maxlength=""9"" value=""" & w_nr_cep & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o CEP da rede de ensino.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        </table>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Dirigentes</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Os dados deste bloco são exibidos na seção ""Quem Somos"" do site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b>Nome do(a) <u>D</u>iretor:</b><br><INPUT ACCESSKEY=""D"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_no_diretor"" size=""40"" maxlength=""40"" value=""" & w_no_diretor & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o nome do(a) diretor(a) da rede de ensino.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b>Nome do(a)<u>S</u>ecretário(a):</b><br><INPUT ACCESSKEY=""S"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_no_secretario"" size=""40"" maxlength=""40"" value=""" & w_no_secretario & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o nome do(a) secretário(a) da rede de ensino.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        </table>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Registro</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Os dados deste bloco não são exibidos no site, servindo apenas para memória das informações relativas ao registro da rede de ensino.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b><u>C</u>NPJ:</b><br><INPUT ACCESSKEY=""C"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_cnpj"" size=""14"" maxlength=""14"" value=""" & w_nr_cnpj & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o CNPJ da rede de ensino.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b><u>T</u>ipo do registro:</b><br><INPUT ACCESSKEY=""T"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_tp_registro"" size=""10"" maxlength=""10"" value=""" & w_tp_registro & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o tipo do registro da unidade. Ex: autorizado, reconhecido etc.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b><u>A</u>to:</b><br><INPUT ACCESSKEY=""A"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_ato"" size=""30"" maxlength=""30"" value=""" & w_ds_ato & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o ato de criação da unidade. Ex: Portaria MEC.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b><u>N</u>úmero:</b><br><INPUT ACCESSKEY=""N"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_ato"" size=""15"" maxlength=""15"" value=""" & w_nr_ato & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o número do ato de criação da unidade. Ex: 3029/1965','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b><u>D</u>ata:</b><br><INPUT ACCESSKEY=""D"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_dt_ato"" size=""10"" maxlength=""10"" value=""" & w_dt_ato & """ ONMOUSEOVER=""popup('OPCIONAL. Informe a data do ato de criação da unidade. O sistema coloca automaticamente as barras separadoras.','white')""; ONMOUSEOUT=""kill()"" onKeyDown=""FormataData(this,event);""></td>"
  ShowHTML "          <td><font size=""1""><b><u>O</u>rgão emissor:</b><br><INPUT ACCESSKEY=""O"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_orgao"" size=""30"" maxlength=""30"" value=""" & w_ds_orgao & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o órgão emissor do registro da unidade.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        </table>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Contato técnico da rede de ensino</td></td></tr>"
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

  Set w_sq_cliente       = Nothing
  Set w_nr_cnpj          = Nothing 
  Set w_tp_registro      = Nothing 
  Set w_ds_ato           = Nothing 
  Set w_nr_ato           = Nothing 
  Set w_dt_ato           = Nothing 
  Set w_ds_orgao         = Nothing 
  Set w_ds_logradouro    = Nothing 
  Set w_no_bairro        = Nothing 
  Set w_nr_cep           = Nothing 
  Set w_no_contato       = Nothing 
  Set w_ds_email_contato = Nothing 
  Set w_nr_fone_contato  = Nothing 
  Set w_nr_fax_contato   = Nothing 
  Set w_no_diretor       = Nothing 
  Set w_no_secretario    = Nothing
End Sub
REM =========================================================================
REM Fim da tela de dados adicionais
REM -------------------------------------------------------------------------

REM =========================================================================
REM Exportação dos dados administrativos
REM -------------------------------------------------------------------------
Sub Administrativo

  Dim w_sq_cliente, w_ds_senha_acesso

  If w_ea = "A" Then
     SQL = "select * from escCliente where " & CL
     ConectaBD SQL
     w_sq_cliente       = RS("sq_cliente")
     w_ds_senha_acesso  = RS("ds_senha_acesso")
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "Javascript"
  ValidateOpen "Validacao"
  ShowHTML "  if (theForm.w_arquivo[0].checked == false && theForm.w_arquivo[1].checked == false) {"
  ShowHTML "     alert('Você deve escolher uma das opções apresentadas antes de gerar o arquivo!');"
  ShowHTML "     return false;"
  ShowHTML "  }"
  ShowHTML "  return(confirm('Confirma a geração do arquivo com os dados indicados?'));"
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Exportação dos dados administrativos</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
  ShowHTML MontaFiltro("POST")
  ShowHTML "<input type=""hidden"" name=""R"" value=""" & w_ew & """>"

  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Exportação dos dados administrativos</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1><ul><li>Esta tela permite a exportação, para um arquivo que pode ser aberto no Excel, dos dados administrativos preenchidos pelas unidades de ensino.<li>Permite também exportar as tabelas de apoio utilizadas pelo formulário.<li>Selecione uma das opções exibidas abaixo e clique no botão ""Gerar arquivo"" para que os dados sejam convertidos para um arquivo.</ul></font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b>O arquivo a ser gerado deve conter dados:</b>"
  ShowHTML "          <br><INPUT " & w_Disabled & " class=""BTM"" type=""radio"" name=""w_arquivo"" value=""Escola""> das unidades de ensino"
  ShowHTML "          <br><INPUT " & w_Disabled & " class=""BTM"" type=""radio"" name=""w_arquivo"" value=""Tipo""> da tabela de equipamentos"
  ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000""></TD></TR>"

  ' Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
  ShowHTML "      <tr><td align=""center""><input class=""STB"" type=""submit"" name=""Botao"" value=""Gerar arquivo""></td>"
  ShowHTML "      </tr>"
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  ShowHTML "</FORM>"
  ShowHTML "</table>"
  Rodape

  Set w_sq_cliente      = Nothing
  Set w_ds_senha_acesso = Nothing
End Sub
REM =========================================================================
REM Fim da tela de dados básicos
REM -------------------------------------------------------------------------

REM =========================================================================
REM Tela de dados básicos
REM -------------------------------------------------------------------------
Sub GetCliente

  Dim w_sq_cliente, w_ds_senha_acesso

  If w_ea = "A" Then
     SQL = "select * from escCliente where " & CL
     ConectaBD SQL
     w_sq_cliente      = RS("sq_cliente")
     w_ds_senha_acesso = RS("ds_senha_acesso")
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "Javascript"
  ValidateOpen "Validacao"
  If w_ea = "A" Then
     Validate "w_ds_senha_acesso", "Senha de acesso", "1", "1", "4", "14", "1", "1"
  End If
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.Form.w_ds_senha_acesso.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Atualização da senha de acesso</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
  ShowHTML MontaFiltro("POST")
  ShowHTML "<input type=""hidden"" name=""R"" value=""" & w_ew & """>"

  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Senha de acesso</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Esta tela permite alterar a senha de acesso à tela de atualização dos dados da rede de ensino. Assim que a nova senha for gravada, ela já deverá ser utilizada para acessar as telas desta aplicação.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>S</u>enha de acesso:</b><br><INPUT ACCESSKEY=""S"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_senha_acesso"" size=""14"" maxlength=""14"" value=""" & w_ds_senha_acesso & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a senha desejada para acessar a tela de atualização dos dados do site.','white')""; ONMOUSEOUT=""kill()""></td>"
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
End Sub
REM =========================================================================
REM Fim da tela de dados básicos
REM -------------------------------------------------------------------------

REM =========================================================================
REM Tela de dados do site
REM -------------------------------------------------------------------------
Sub GetSite

  Dim w_sq_cliente, w_im_foto_abertura1, w_im_foto_abertura2, w_no_contato_internet, w_ds_email_internet
  Dim w_nr_fone_internet, w_nr_fax_internet, w_ds_texto_abertura, w_ds_institucional, w_ds_mensagem

  If w_ea = "A" Then
     SQL = "select * from escCliente_Site where " & CL
     ConectaBD SQL
     w_sq_cliente          = RS("sq_cliente")
     w_no_contato_internet = RS("no_contato_internet")
     w_ds_email_internet   = RS("ds_email_internet")
     w_nr_fone_internet    = RS("nr_fone_internet")
     w_nr_fax_internet     = RS("nr_fax_internet")
     w_im_foto_abertura1   = RS("im_foto_abertura1")
     w_im_foto_abertura2   = RS("im_foto_abertura2")
     w_ds_texto_abertura   = RS("ds_texto_abertura")
     w_ds_institucional    = RS("ds_institucional")
     w_ds_mensagem         = RS("ds_mensagem")
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "Javascript"
  ValidateOpen "Validacao"
  If w_ea = "A" Then
     Validate "w_no_contato_internet" , "Nome"                                , "1" , "1" , "2" , "35"   , "1" , "1"
     Validate "w_ds_email_internet"   , "e-Mail"                              , "1" , "1" , "6" , "60"   , "1" , "1"
     Validate "w_nr_fone_internet"    , "Telefone"                            , "1" , "1" , "6" , "20"   , "1" , "1"
     Validate "w_nr_fax_internet"     , "Fax"                                 , "1" , ""  , "6" , "20"   , "1" , "1"
     Validate "w_im_foto_abertura1"   , "Foto 1 da escola"                    , "1" , ""  , "5" , "60"   , "1" , "1"
     Validate "w_im_foto_abertura2"   , "Foto 2 da escola"                    , "1" , ""  , "5" , "60"   , "1" , "1"
     Validate "w_ds_texto_abertura"   , "Texto de abertura"                   , "1" , "1" , "4" , "3200" , "1" , "1"
     Validate "w_ds_institucional"    , "Texto da seção \""Quem somos\"""     , "1" , "1" , "4" , "3200" , "1" , "1"
     Validate "w_ds_mensagem"         , "\""Texto da mensagem em destaque\""" , "1" , ""  , "4" , "80"   , "1" , "1"
  End If
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.Form.w_no_contato_internet.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">Atualização de dados do site da rede de ensino</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
  ShowHTML MontaFiltro("POST")
  ShowHTML "<input type=""hidden"" name=""R"" value=""" & w_ew & """>"
  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Contato da rede de ensino para divulgação no site</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Informe os dados da pessoa a ser exibida como contato da rede de ensino para divulgação no site.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b>No<u>m</u>e:</b><br><INPUT ACCESSKEY=""M"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_no_contato_internet"" size=""35"" maxlength=""35"" value=""" & w_no_contato_internet & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o nome da pessoa a ser exibida na seção \'Quem somos\' do site da rede de ensino.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b><u>e</u>-Mail:</b><br><INPUT ACCESSKEY=""E"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_email_internet"" size=""40"" maxlength=""60"" value=""" & w_ds_email_internet & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o e-mail da pessoa.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1""><b>Tele<u>f</u>one:</b><br><INPUT ACCESSKEY=""F"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_fone_internet"" size=""20"" maxlength=""20"" value=""" & w_nr_fone_internet & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o telefone da pessoa.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td><font size=""1""><b>Fa<u>x</u>:</b><br><INPUT ACCESSKEY=""X"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_nr_fax_internet"" size=""20"" maxlength=""20"" value=""" & w_nr_fax_internet & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o fax da pessoa.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "        </table>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Imagens da página de abertura do site</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Informe até duas imagens a serem colocadas na página de abertura do site. O sistema tratará adequadamente caso sejam informadas duas, uma ou nenhuma imagem.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=""1""><b>Im<u>a</u>gem 1:</b><br><INPUT ACCESSKEY=""A"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_im_foto_abertura1"" size=""60"" maxlength=""60"" value=""" & w_im_foto_abertura1 & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o nome do arquivo que contém a imagem. Este arquivo deve ser enviado por e-mail para o responsável pelo site.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td><font size=""1""><b>Ima<u>g</u>em 2:</b><br><INPUT ACCESSKEY=""G"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_im_foto_abertura2"" size=""60"" maxlength=""60"" value=""" & w_im_foto_abertura2 & """ ONMOUSEOVER=""popup('OPCIONAL. Informe o nome do arquivo que contém a imagem. Este arquivo deve ser enviado por e-mail para o responsável pelo site.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Textos do site</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1>Os textos abaixo serão exibidos na página de abertura, na seção ""Quem somos"" e na barra rolante localizada na parte inferior da página.</font></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "      <tr><td><font size=""1""><b>Texto da <u>p</u>ágina de abertura:</b><br><TEXTAREA ACCESSKEY=""P"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_texto_abertura"" rows=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o texto a ser exibido na página de abertura do site.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_texto_abertura & "</TEXTAREA></td>"
  ShowHTML "      <tr><td><font size=""1""><b>Texto da seção ""<u>Q</u>uem somos"":</b><br><TEXTAREA ACCESSKEY=""Q"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_institucional"" rows=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o texto a ser exibido na seção \'Quem somos\' do site.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_institucional & "</TEXTAREA></td>"
  ShowHTML "      <tr><td><font size=""1""><b>Texto da men<u>s</u>agem em destaque:</b><br><INPUT ACCESSKEY=""S"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_mensagem"" size=""80"" maxlength=""80"" value=""" & w_ds_mensagem & """ ONMOUSEOVER=""popup('OPCIONAL. Informe um texto que será exibido na parte inferior do site, numa barra rolante.','white')""; ONMOUSEOUT=""kill()""></td>"
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

  Set w_sq_cliente          = Nothing
  Set w_no_contato_internet = Nothing 
  Set w_ds_email_internet   = Nothing 
  Set w_nr_fone_internet    = Nothing 
  Set w_nr_fax_internet     = Nothing 
  Set w_im_foto_abertura1   = Nothing 
  Set w_im_foto_abertura2   = Nothing 
  Set w_ds_texto_abertura   = Nothing 
  Set w_ds_institucional    = Nothing 
  Set w_ds_mensagem         = Nothing
End Sub

REM =========================================================================
REM Cadastro de notícias
REM -------------------------------------------------------------------------
Sub GetNoticiaCliente
  Dim w_chave, w_ds_titulo, w_in_ativo, w_ds_noticia, w_ln_externo
  Dim w_dt_noticia, w_in_exibe
  
  Dim w_troca, i, w_texto
  
  w_Chave         = Request("w_Chave")
  w_troca         = Request("w_troca")
  
  If w_troca > "" Then ' Se for recarga da página
     w_dt_noticia = Request("w_dt_noticia")    
     w_ds_titulo  = replace(Request("w_ds_titulo"),"""","&quot;")
     w_ln_externo = replace(Request("w_ln_externo"),"""","&quot;")
     w_ds_noticia = Request("w_ds_noticia")    
     w_ln_noticia = Request("w_ln_noticia")
     w_in_ativo   = Request("w_in_ativo")    
     w_in_exibe   = Request("w_in_exibe")    
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then
        SQL = "select * from escNoticia_Cliente where sq_site_cliente = 0 order by dt_noticia desc"
      Else
        SQL = "select * from escNoticia_Cliente where " & replace(CL,"sq_cliente","sq_site_cliente") & " order by dt_noticia desc"
      End If
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escNoticia_Cliente where sq_noticia = " & w_chave
     ConectaBD SQL
     w_dt_noticia = FormataDataEdicao(RS("dt_noticia"))
     w_ds_titulo  = RS("ds_titulo")
     w_ds_titulo  = replace(w_ds_titulo,"""","&quot;")
     w_ln_externo = RS("ln_externo")
     if(w_ln_externo > "") Then
        w_ln_externo = replace(w_ln_externo,"""","&quot;")
     End If     
     w_ds_noticia = RS("ds_noticia")
     w_ds_noticia  = replace(w_ds_noticia,"""","&quot;")
     w_in_ativo   = RS("in_ativo")
     w_in_exibe   = RS("in_exibe")
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
        Validate "w_dt_noticia" , "Data"      , "DATA" , "1" , "10" , "10"   , "1" , "1"
        Validate "w_ds_titulo"  , "Título"    , ""     , "1" , "2"  , "50"   , "1" , "1"
        Validate "w_ds_noticia" , "Descrição" , ""     , "1" , "2"  , "4000" , "1" , "1"
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
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de notícias da rede de ensino</FONT></B>"
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
        ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("dt_noticia"),2)) & "</td>"
        if(RS("ln_externo") > "") Then
           ShowHTML "        <td><font size=""1""><a href="""& RS("ln_externo") & """ target=""_blank"">" & RS("ds_titulo") & "</td>"
        Else
           ShowHTML "        <td><font size=""1"">" & RS("ds_titulo") & "</td>"        
        End If
        ShowHTML "        <td><font size=""1"">" & RS("ds_noticia") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("in_ativo") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_noticia") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_chave=" & RS("sq_noticia") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
    ShowHTML "      </tr>"
    ShowHTML "      <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>T</u>ítulo:</b><br><input " & w_Disabled & " accesskey=""T"" type=""text"" name=""w_ds_titulo"" class=""STI"" SIZE=""90"" MAXLENGTH=""100"" VALUE=""" & w_ds_titulo & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe um título para o noticia.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "      </tr>"
    ShowHTML "      <tr valign=""top"">"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>L</u>ink externo:</b><br><input " & w_Disabled & " accesskey=""L"" type=""text"" name=""w_ln_externo"" class=""STI"" SIZE=""90"" MAXLENGTH=""255"" VALUE=""" & w_ln_externo & """ ONMOUSEOVER=""popup('OPCIONAL. Informe um link externo para o noticia.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "      </tr>"
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

  Set w_ds_titulo  = Nothing 
  Set w_ln_externo = Nothing
  Set w_in_ativo   = Nothing 
  Set w_ds_noticia = Nothing 
  Set w_ln_noticia = Nothing
  Set w_chave      = Nothing 
  Set w_troca      = Nothing 
  Set i            = Nothing 
  Set w_texto      = Nothing
End Sub

REM =========================================================================
REM Cadastro de sistemas
REM -------------------------------------------------------------------------
Sub GetSistema
  Dim w_chave, w_no_sistema, w_ativo, w_ds_sistema
  Dim w_sg_sistema
  
  Dim w_troca, i, w_texto
  
  w_Chave         = Request("w_Chave")
  w_troca         = Request("w_troca")
  
  If w_troca > "" Then ' Se for recarga da página
     w_sg_sistema = Request("w_sg_sistema")    
     w_no_sistema = Request("w_no_sistema")
     w_ds_sistema = Request("w_ds_sistema")    
     w_ativo      = Request("w_ativo")    
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select * from escSistema order by no_sistema"
     ConectaBD SQL
  ElseIf InStr("AE",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escSistema where sq_sistema = " & w_chave
     ConectaBD SQL
     w_sg_sistema = RS("sg_sistema")
     w_no_sistema = RS("no_sistema")
     w_ds_sistema = RS("ds_sistema")
     w_ativo      = RS("ativo")
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
        Validate "w_sg_sistema" , "Sigla"     , "" , "1" , "3" , "10"   , "1" , "1"
        Validate "w_no_sistema"  , "Título"   , "" , "1" , "3"  , "30"  , "1" , "1"
        Validate "w_ds_sistema" , "Descrição" , "" , "1" , "3"  , "300" , "1" , "1"
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
     BodyOpen "onLoad='document.Form.w_sg_sistema.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de sistemas</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Sigla</font></td>"
    ShowHTML "          <td><font size=""1""><b>Nome</font></td>"
    ShowHTML "          <td><font size=""1""><b>Ativo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
       ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("sg_sistema") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("no_sistema") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("ativo") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&w_chave=" & RS("sq_sistema") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&w_chave=" & RS("sq_sistema") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
  ElseIf Instr("IAE",O) > 0 Then
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
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>S</u>igla:</b><br><input accesskey=""S"" type=""text"" name=""w_sg_sistema"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & w_sg_sistema & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a sigla do sistema.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>N</u>ome:</b><br><input " & w_Disabled & " accesskey=""N"" type=""text"" name=""w_no_sistema"" class=""STI"" SIZE=""30"" MAXLENGTH=""30"" VALUE=""" & w_no_sistema & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o nome do sistema.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "        </table>"
    ShowHTML "      <tr><td valign=""top""><font size=""1""><b>D<u>e</u>scrição:</b><br><textarea " & w_Disabled & " accesskey=""E"" name=""w_ds_sistema"" class=""STI"" ROWS=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva sucDblamente o objetivo do sistema e seu ambiente de execução.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_sistema & "</TEXTAREA></td>"
    ShowHTML "      <tr>"
    ShowHTML "      </tr>"
    ShowHTML "      <tr>"
    MontaRadioSN "<b>Ativo?</b>", w_ativo, "w_ativo"
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

  Set w_no_sistema  = Nothing 
  Set w_ativo       = Nothing 
  Set w_ds_sistema  = Nothing 
  Set w_chave       = Nothing 
  Set w_troca       = Nothing 
  Set i             = Nothing 
  Set w_texto       = Nothing
End Sub

REM =========================================================================
REM Cadastro de componentes
REM -------------------------------------------------------------------------
Sub GetComponente
  Dim w_chave, w_no_fisico, w_ativo, w_ds_componente
  Dim w_sq_sistema
  
  Dim w_troca, i, w_texto
  
  w_Chave         = Request("w_Chave")
  w_troca         = Request("w_troca")
  
  If w_troca > "" Then ' Se for recarga da página
     w_sq_sistema    = Request("w_sq_sistema")    
     w_no_fisico     = Request("w_no_fisico")
     w_ds_componente = Request("w_ds_componente")    
     w_ativo         = Request("w_ativo")    
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select a.*, b.sg_sistema " & VbCrLf & _
           "  from escComponente a " & VbCrLf & _
           "       inner join escSistema b on (a.sq_sistema = b.sq_sistema) " & VbCrLf & _
           "order by a.no_fisico "
     ConectaBD SQL
  ElseIf InStr("AE",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escComponente where sq_componente = " & w_chave
     ConectaBD SQL
     w_sq_sistema    = RS("sq_sistema")
     w_no_fisico     = RS("no_fisico")
     w_ds_componente = RS("ds_componente")
     w_ativo         = RS("ativo")
     DesconectaBD
  ElseIf w_ea = "V" Then
     ' Recupera todos os registros para a listagem
     Set RS1 = Server.CreateObject("ADODB.RecordSet")
     SQL = "select a.*, b.sg_sistema, b.no_sistema " & VbCrLf & _
           "  from escComponente a " & VbCrLf & _
           "       inner join escSistema b on (a.sq_sistema = b.sq_sistema) " & VbCrLf & _
           "order by a.no_fisico "
     RS1.Open SQL, dbms, adOpenStatic

     ' Recupera as versões do componente informado
     SQL = "select a.*, c.sg_sistema, c.no_sistema, b.no_fisico, b.ds_componente, b.ativo cm_ativo" & VbCrLf & _
           "  from escComponente_Versao     a " & VbCrLf & _
           "       inner join escComponente b on (a.sq_componente = b.sq_componente) " & VbCrLf & _
           "       inner join escSistema    c on (b.sq_sistema    = c.sq_sistema) " & VbCrLf & _
           " where b.sq_componente = " & w_chave & VbCrLf & _
           "order by c.sg_sistema, b.no_fisico, replace(a.cd_versao,'.','') desc "
     ConectaBD SQL
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  If InStr("IAEP",O) > 0 Then
     ScriptOpen "JavaScript"
     CheckBranco
     FormataData
     ValidateOpen "Validacao"
     If InStr("IA",O) > 0 Then
        Validate "w_sq_sistema" , "Sistema"        , "SELECT" , "1" , "1" , "10"   , "1" , "1"
        Validate "w_no_fisico"  , "Nome do arquivo", ""       , "1" , "3"  , "30"  , "1" , "1"
        Validate "w_ds_componente" , "Descrição"   , ""       , "1" , "3"  , "300" , "1" , "1"
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
     BodyOpen "onLoad='document.Form.w_sq_sistema.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  If w_ea = "V" Then
     ShowHTML "<B><FONT COLOR=""#000000"">Histórico de versões de componente</FONT></B>"
  Else
     ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de Componentes</FONT></B>"
  End If
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
    ShowHTML "          <td><font size=""1""><b>Sistema</font></td>"
    ShowHTML "          <td><font size=""1""><b>Componente</font></td>"
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
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("sg_sistema") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("no_fisico") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("ds_componente") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("ativo") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&w_chave=" & RS("sq_componente") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&w_chave=" & RS("sq_componente") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=V&w_chave=" & RS("sq_componente") & """ target=""Versao"">Versões</A>&nbsp"
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
  ElseIf Instr("IAE",w_ea) > 0 Then
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
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>S</u>istema:</b><br><select accesskey=""S"" name=""w_sq_sistema"" class=""STS"" SIZE=""1"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Selecione o sistema ao qual o componente pertence.','white')""; ONMOUSEOUT=""kill()"">"
    ShowHTML "          <OPTION VALUE="""">---"
    SQL = "select sq_sistema, sg_sistema from escSistema order by sg_sistema" & VbCrLf
    ConectaBD SQL
    While Not RS.EOF
       If cDbl(Nvl(w_sq_sistema,0)) = cDbl(RS("sq_sistema")) or RS.RecordCount = 1 Then
          ShowHTML "          <OPTION VALUE=""" & RS("sq_sistema") & """ SELECTED>" & RS("sg_sistema")
       Else
          ShowHTML "          <OPTION VALUE=""" & RS("sq_sistema") & """>" & RS("sg_sistema")
       End If
       If RS.RecordCount = 1 Then
          w_sq_sistema = RS("sq_sistema")
       End If
       RS.MoveNext
    Wend
    DesconectaBD
    ShowHTML "          </select>"
    ShowHTML "          <td valign=""top""><font size=""1""><b><u>N</u>ome do arquivo físico:</b><br><input " & w_Disabled & " accesskey=""N"" type=""text"" name=""w_no_fisico"" class=""STI"" SIZE=""30"" MAXLENGTH=""30"" VALUE=""" & w_no_fisico & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o nome do arquivo físico do componente.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "        </table>"
    ShowHTML "      <tr><td valign=""top""><font size=""1""><b>D<u>e</u>scrição:</b><br><textarea " & w_Disabled & " accesskey=""E"" name=""w_ds_componente"" class=""STI"" ROWS=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva sucDblamente o componente.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_componente & "</TEXTAREA></td>"
    ShowHTML "      <tr>"
    ShowHTML "      </tr>"
    ShowHTML "      <tr>"
    MontaRadioSN "<b>Ativo?</b>", w_ativo, "w_ativo"
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
  ElseIf w_ea = "V" Then
    ShowHTML "<tr><td colspan=2 align=""center"" bgcolor=""#FAEBD7""><table border=1 width=""100%""><tr><td>"
    ShowHTML "    <TABLE WIDTH=""100%"" CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "      <tr valign=""top"">"
    ShowHTML "          <td><font size=""1"">Sistema:<b><br><font size=1>" & RS1("sg_sistema") & " - " & RS1("no_sistema") & "</font></font></td>"
    ShowHTML "          <td><font size=""1"">Componente:<b><br><font size=1>" & RS1("no_fisico") & "</font></font></td>"
    ShowHTML "      <tr valign=""top"">"
    ShowHTML "          <td><font size=""1"">Descrição:<b><br>" & RS1("ds_componente") & "</td>"
    ShowHTML "          <td><font size=""1"">Ativo:<b><br><font size=1>" & RS1("ativo") & "</font></font></td>"
    ShowHTML "      </tr>"
    ShowHTML "    </TABLE>"
    ShowHTML "</table>"

    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML "<tr><td><font size=""2"">&nbsp;"
    ShowHTML "<tr><td><font size=""2"">"
    ShowHTML "    <td align=""right""><font size=""1""><b>Versões existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Versão</font></td>"
    ShowHTML "          <td><font size=""1""><b>Data</font></td>"
    ShowHTML "          <td><font size=""1""><b>Arquivo original</font></td>"
    ShowHTML "          <td><font size=""1""><b>Arquivo físico</font></td>"
    ShowHTML "          <td><font size=""1""><b>Descrição</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
       ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("cd_versao") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("dt_versao"),2)) & "</td>"
        ShowHTML "        <td align=""center""><font size=""1""><a class=""hl"" target=""_blank"" href=""" & RS("ds_caminho") & RS("sq_componente_versao") & Mid(RS("no_arquivo"),Instr(RS("no_arquivo"),"."),50) & """>" & RS("no_arquivo") & "</a></td>"
        ShowHTML "        <td align=""center""><font size=""1""><a class=""hl"" target=""_blank"" href=""" & RS("ds_caminho") & RS("sq_componente_versao") & Mid(RS("no_arquivo"),Instr(RS("no_arquivo"),"."),50) & """>" & RS("sq_componente_versao") & Mid(RS("no_arquivo"),Instr(RS("no_arquivo"),"."),50) & "</a></td>"
        ShowHTML "        <td><font size=""1"">" & RS("ds_versao") & "</td>"
        ShowHTML "        </td>"
        ShowHTML "      </tr>"
        RS.MoveNext
      wend
    End If
    ShowHTML "      </center>"
    ShowHTML "    </table>"
    If RS1("ativo") = "Sim" Then
       ShowHTML "<tr><td colspan=3><font size=1><b>Observação: se existir mais de uma versão para o componente, apenas a que estiver na primeira linha da listagem acima será incluída no arquivo de atualização."
    Else
       ShowHTML "<tr><td colspan=3><font size=1><b>Observação: nenhuma versão deste componente será incluída no arquivo de atualização, pois o componente não está ativo."
    End If
    ShowHTML "  </td>"
    ShowHTML "</tr>"
    DesconectaBD
    RS1.Close
  Else
    ScriptOpen "JavaScript"
    ShowHTML " alert('Opção não disponível');"
    'ShowHTML " history.back(1);"
    ScriptClose
  End If
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape

  Set w_no_fisico       = Nothing 
  Set w_ativo           = Nothing 
  Set w_ds_componente   = Nothing 
  Set w_chave           = Nothing 
  Set w_troca           = Nothing 
  Set i                 = Nothing 
  Set w_texto           = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de componentes
REM -------------------------------------------------------------------------

REM =========================================================================
REM Cadastro de versões
REM -------------------------------------------------------------------------
Sub GetVersao
  Dim w_chave, w_cd_versao, w_dt_versao, w_ds_versao, w_no_fisico, w_no_arquivo
  Dim w_sq_componente
  
  Dim w_troca, i, ds_caminho
  
  If InStr(uCase(Request.ServerVariables("http_content_type")),"MULTIPART/FORM-DATA") > 0 Then
     w_Chave         = ul.Form("w_Chave")
     w_troca         = ul.Form("w_troca")
  Else
     w_Chave         = Request("w_Chave")
     w_troca         = Request("w_troca")
  End If
  
  If w_troca > "" Then ' Se for recarga da página
     w_sq_componente    = ul.Form("w_sq_componente")    
     w_cd_versao        = ul.Form("w_cd_versao")
     w_ds_versao        = ul.Form("w_ds_versao")
     w_dt_versao        = ul.Form("w_dt_versao")
     w_ds_caminho       = ul.Form("w_ds_caminho")
     w_no_arquivo       = ul.Form("w_no_arquivo")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select a.*, c.sg_sistema, b.no_fisico " & VbCrLf & _
           "  from escComponente_Versao     a " & VbCrLf & _
           "       inner join escComponente b on (a.sq_componente = b.sq_componente and " & VbCrLf & _
           "                                      b.ativo         = 'Sim' " & VbCrLf & _
           "                                     ) " & VbCrLf & _
           "       inner join escSistema    c on (b.sq_sistema    = c.sq_sistema and " & VbCrLf & _
           "                                      c.ativo         = 'Sim' " & VbCrLf & _
           "                                     ) " & VbCrLf & _
           "       inner join (select sq_componente, max(replace(cd_versao,'.','')) versao " & VbCrLf & _
           "                     from escComponente_Versao " & VbCrLf & _
           "                   group by sq_componente " & VbCrLf & _
           "                  )             d on (a.sq_componente = d.sq_componente and " & VbCrLf & _
           "                                      replace(a.cd_versao,'.','') = d.versao " & VbCrLf & _
           "                                     ) " & VbCrLf & _
           "order by c.sg_sistema, b.no_fisico, ds_caminho,a.dt_versao desc, replace(a.cd_versao,'.','') desc " & VbCrLf
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escComponente_Versao where sq_componente_versao = " & w_chave
     ConectaBD SQL
     w_sq_componente    = RS("sq_componente")
     w_cd_versao        = RS("cd_versao")
     w_ds_versao        = RS("ds_versao")
     w_dt_versao        = FormataDataEdicao(RS("dt_versao"))
     w_ds_caminho       = RS("ds_caminho")
     w_no_arquivo       = RS("no_arquivo")
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
        Validate "w_sq_componente" , "Sistema"              , "SELECT" , "1" , "1" , "10"   , "1" , "1"
        Validate "w_cd_versao"     , "Código da versão"     , ""       , "1" , "3"  , "20"  ,  "1" , "1"
        Validate "w_dt_versao"     , "Data da versão"       , "DATA"   , "1" , "10"  , "10" ,  "" , "0123456789/"
        CompData "w_dt_versao"     , "Data da versão"       , "<="   , w_Data, "Data atual"
        Validate "w_ds_caminho"    , "URL para download"    , ""       , "1" , "10"  , "100" , "1" , "1"
        If O = "I" Then
           Validate "w_no_arquivo"    , "Arquivo"              , ""       , "1" , "10"  , "100" , "1" , "1"
        Else
           Validate "w_no_arquivo"    , "Arquivo"              , ""       , ""  , "10"  , "100" , "1" , "1"
        End If
        ShowHTML "  if (theForm.w_no_arquivo.value.toUpperCase().indexOf('.EXE') < 0 && theForm.w_no_arquivo.value.toUpperCase().indexOf('.ZIP') < 0) {"
        ShowHTML "     alert('São aceitos apenas arquivos com extenção EXE e ZIP!');"
        ShowHTML "     theForm.w_no_arquivo.focus();"
        ShowHTML "     return false;"
        ShowHTML "  }"
        Validate "w_ds_versao"     , "Descrição"            , ""       , "1" , "3"  , "300" , "1" , "1"
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
     BodyOpen "onLoad='document.Form.w_sq_componente.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de Versões</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  If w_ea = "L" Then
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML "<tr><td><font size=""2"">"
    ShowHTML "        <a accesskey=""I"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=I&CL=" & CL & """><u>I</u>ncluir</a>&nbsp;"
    ShowHTML "        <a accesskey=""G"" class=""SS"" href=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=G&CL=" & CL & """ onClick=""return(confirm('Confirma a geração do arquivo de atualização?\nO arquivo atual, se existir, será sobreposto.'));""><u>G</u>erar arquivo</a>&nbsp;"
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Sistema</font></td>"
    ShowHTML "          <td><font size=""1""><b>Componente</font></td>"
    ShowHTML "          <td><font size=""1""><b>Versão</font></td>"
    ShowHTML "          <td><font size=""1""><b>Data</font></td>"
    ShowHTML "          <td><font size=""1""><b>Arquivo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Descrição</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
       ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=7 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("sg_sistema") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("no_fisico") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("cd_versao") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("dt_versao"),2)) & "</td>"
        ShowHTML "        <td align=""center""><font size=""1""><a class=""hl"" target=""_blank"" href=""" & RS("ds_caminho") & RS("sq_componente_versao") & Mid(RS("no_arquivo"),Instr(RS("no_arquivo"),"."),50) & """>" & RS("no_arquivo") & "</a></td>"
        ShowHTML "        <td><font size=""1"">" & RS("ds_versao") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&w_chave=" & RS("sq_componente_versao") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&w_chave=" & RS("sq_componente_versao") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
    ShowHTML "<INPUT type=""hidden"" name=""w_troca"" value="""">"
    ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & w_ew & """>"
    ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_chave"" value=""" & w_chave & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente"" value=""" & replace(CL,"sq_cliente=","") & """>"
    ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"

    ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
    ShowHTML "    <table width=""95%"" border=""0"">"
    ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top"">"
    ShowHTML "          <td><font size=""1""><b><u>C</u>omponente:</b><br><select accesskey=""C"" name=""w_sq_componente"" class=""STS"" SIZE=""1"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Selecione o componente ao qual a versão pertence.','white')""; ONMOUSEOUT=""kill()"">"
    ShowHTML "          <OPTION VALUE="""">---"
    SQL = "select a.sq_componente, b.sg_sistema+' - '+a.no_fisico nm_componente from escComponente a inner join escSistema b on (a.sq_sistema = b.sq_sistema) order by b.sg_sistema, a.no_fisico " & VbCrLf
    ConectaBD SQL
    While Not RS.EOF
       If cDbl(Nvl(w_sq_componente,0)) = cDbl(RS("sq_componente")) or RS.RecordCount = 1 Then
          ShowHTML "          <OPTION VALUE=""" & RS("sq_componente") & """ SELECTED>" & RS("nm_componente")
       Else
          ShowHTML "          <OPTION VALUE=""" & RS("sq_componente") & """>" & RS("nm_componente")
       End If
       If RS.RecordCount = 1 Then
          w_sq_componente = RS("sq_componente")
       End If
       RS.MoveNext
    Wend
    DesconectaBD
    ShowHTML "          </select>"
    ShowHTML "          <td><font size=""1""><b>Código da <u>v</u>ersão:</b><br><input " & w_Disabled & " accesskey=""V"" type=""text"" name=""w_cd_versao"" class=""STI"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & w_cd_versao & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o código da versão.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "          <td><font size=""1""><b><u>D</u>ata da versão:</b><br><INPUT ACCESSKEY=""D"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_dt_versao"" size=""10"" maxlength=""10"" value=""" & Nvl(w_dt_versao, w_data) & """ ONMOUSEOVER=""popup('OPCIONAL. Informe a data de construção da versão. O sistema coloca automaticamente as barras separadoras.','white')""; ONMOUSEOUT=""kill()"" onKeyDown=""FormataData(this,event);""></td>"
    ShowHTML "        </table>"
    ShowHTML "        <tr><td><font size=""1""><b><u>U</u>RL para download:</b><br><input " & w_Disabled & " accesskey=""U"" type=""text"" name=""w_ds_caminho"" class=""STI"" SIZE=""75"" MAXLENGTH=""100"" VALUE=""" & Nvl(w_ds_caminho, "sge/") & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a URL onde a versão pode ser obtida.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "        <tr><td><font size=""1""><b><u>A</u>rquivo:</b><br><input " & w_Disabled & " accesskey=""A"" type=""file"" name=""w_no_arquivo"" class=""STI"" SIZE=""80"" MAXLENGTH=""100"" VALUE="""" ONMOUSEOVER=""popup('OBRIGATÓRIO. Clique no botão ao lado para localizar o arquivo que contém a versão do componente. Ele será transferido automaticamente para o servidor.','white')""; ONMOUSEOUT=""kill()"">"
    If w_no_arquivo > "" Then
       ShowHTML "              <b><a class=""SS"" href=""" & w_ds_caminho & w_chave & Mid(w_no_arquivo, Instr(w_no_arquivo,"."), 50) & """ target=""_blank"" title=""Clique para recuperar o arquivo atual."">Exibir</a></b>"
    End If
    ShowHTML "      <tr><td><font size=""1""><b>D<u>e</u>scrição:</b><br><textarea " & w_Disabled & " accesskey=""E"" name=""w_ds_versao"" class=""STI"" ROWS=5 cols=65 ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva sucDblamente o componente.','white')""; ONMOUSEOUT=""kill()"">" & w_ds_versao & "</TEXTAREA></td>"
    ShowHTML "      <tr>"
    ShowHTML "      </tr>"
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

  Set w_cd_versao  = Nothing 
  Set w_dt_versao  = Nothing 
  Set w_ds_versao  = Nothing 
  Set w_chave      = Nothing 
  Set w_troca      = Nothing 
  Set i            = Nothing 
  Set ds_caminho   = Nothing
  Set w_no_fisico  = Nothing 
End Sub
REM =========================================================================
REM Fim do cadastro de versões
REM -------------------------------------------------------------------------

REM =========================================================================
REM Cadastro da rede particular
REM -------------------------------------------------------------------------
Sub GetRedeParticular
  Cabecalho
  ShowHTML "<HEAD>"
  ScriptOpen "JavaScript"
  ValidateOpen "Validacao"
  Validate "w_no_arquivo"    , "Arquivo"              , ""       , "1" , "10"  , "100" , "1" , "1"
  ShowHTML "  theForm.Botao.disabled=true;"
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.focus()';"
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro da Rede Particular</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  ShowHTML "<FORM action=""" & w_pagina & "Grava"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"" enctype=""multipart/form-data"">"
  ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & w_ew & """>"
  ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
  ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente"" value=""" & replace(CL,"sq_cliente=","") & """>"
  ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"

  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""95%"" border=""0"">"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><table border=0 width=""100%"" cellspacing=0>"
  ShowHTML "        <tr><td><font size=""1""><b><u>A</u>rquivo:</b><br><input " & w_Disabled & " accesskey=""A"" type=""file"" name=""w_no_arquivo"" class=""STI"" SIZE=""80"" MAXLENGTH=""100"" VALUE="""" ONMOUSEOVER=""popup('OBRIGATÓRIO. Clique no botão ao lado para localizar o arquivo que contém a versão do componente. Ele será transferido automaticamente para o servidor.','white')""; ONMOUSEOUT=""kill()"">"
  ShowHTML "      <tr><td align=""center"" colspan=4><hr>"
  ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Enviar"">"
  ShowHTML "          </td>"
  ShowHTML "      </tr>"
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  ShowHTML "</FORM>"
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape
End Sub
REM =========================================================================
REM Cadastro da rede particular
REM -------------------------------------------------------------------------

REM =========================================================================
REM Cadastro de newsletter
REM -------------------------------------------------------------------------
Sub GetNewsletter
  Dim w_chave, w_nome, w_tipo, w_email, w_envia_mail
  Dim w_troca, i, w_texto
  
  w_Chave           = Request("w_Chave")
  w_troca           = Request("w_troca")
  
  If w_troca > "" Then ' Se for recarga da página
     w_nome       = replace(Request("w_nome"),"""","&quot;")
     w_email      = replace(Request("w_email"),"""","&quot;")
     w_tipo       = Request("w_tipo")
     w_envia_mail = Request("w_envia_mail")
  ElseIf w_ea = "L" or w_ea = "G" Then
     If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then
        SQL = "select sq_newsletter, nome, email, tipo, envia_mail, data_inclusao, data_alteracao, " & VbCrLf & _
              "       case tipo when 1 then 'Responsável' " & VbCrLf & _
              "                 when 2 then 'Aluno' " & VbCrLf & _
              "                 when 3 then 'Outro' " & VbCrLf & _
              "       end nm_tipo, " & VbCrLf & _
              "       case envia_mail when 'S' then 'Sim' else 'Não' end nm_envia " & VbCrLf & _
              "  from escNewsletter " & VbCrLf & _
              " where sq_cliente = 0 " & VbCrLf
     Else
        ' Recupera todos os registros para a listagem
        SQL = "select sq_newsletter, nome, email, tipo, envia_mail, data_inclusao, data_alteracao, " & VbCrLf & _
              "       case tipo when 1 then 'Responsável' " & VbCrLf & _
              "                 when 2 then 'Aluno' " & VbCrLf & _
              "                 when 3 then 'Outro' " & VbCrLf & _
              "       end nm_tipo, " & VbCrLf & _
              "       case envia_mail when 'S' then 'Sim' else 'Não' end nm_envia " & VbCrLf & _
              "  from escNewsletter " & VbCrLf & _
              " where " & CL & " " & VbCrLf
     End If
     If w_ea = "G" Then
        SQL = SQL & "   and envia_mail = 'S' " & VbCrLf
     End If
     SQL = SQL & "order by nome"
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escNewsletter where sq_newsletter = " & w_chave
     ConectaBD SQL
     w_nome       = replace(RS("nome"),"""","&quot;")
     w_email      = replace(RS("email"),"""","&quot;")
     w_tipo       = RS("tipo")
     w_envia_mail = RS("envia_mail")
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
        Validate "w_nome"  , "Nome"   , "" , "1" , "3" , "60" , "1" , "1"
        Validate "w_email" , "e-Mail" , "" , "1" , "4" , "60" , "1" , "1"
        ShowHTML "  if (theForm.w_tipo[0].checked==false && theForm.w_tipo[1].checked==false && theForm.w_tipo[2].checked==false) {"
        ShowHTML "     alert('Você deve selecionar uma das opções apresentadas no formulário!');"
        ShowHTML "     return false;"
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
     BodyOpen "onLoad='document.Form.w_nome.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Lista de distribuição de informativos</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  If w_ea = "L" Then
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML "<tr><td><font size=""2"">"
    ShowHTML "       <a accesskey=""I"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & par & "&w_ea=I&CL=" & CL & """><u>I</u>ncluir</a>&nbsp;"
    ShowHTML "       <a accesskey=""L"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & par & "&w_ea=G&CL=" & CL & """><u>L</u>istar e-mails </a>&nbsp;"
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Nome</font></td>"
    ShowHTML "          <td><font size=""1""><b>Tipo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Envia</font></td>"
    ShowHTML "          <td><font size=""1""><b>Cadastro</font></td>"
    ShowHTML "          <td><font size=""1""><b>Alteração</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
       ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=6 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td><font size=""1""><a class=""HL"" href=""mailto:" & RS("email") & """ title=""" & RS("email") & """>" & RS("nome") & "</a></td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("nm_Tipo") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("nm_envia") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("data_inclusao"),2)) & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">"
        If RS("data_alteracao") > "" Then ShowHTML FormataDataEdicao(FormatDateTime(RS("data_alteracao"),2)) Else ShowHTML "---" End If
        ShowHTML "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_newsletter") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_chave=" & RS("sq_newsletter") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
  ElseIf w_ea = "G" Then
    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML "<tr><td><font size=""2"">"
    ShowHTML "       <a accesskey=""V"" class=""SS"" href=""" & w_Pagina & w_ew & "&R=" & w_Pagina & par & "&w_ea=L&CL=" & CL & """><u>V</u>oltar</a>&nbsp;"
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros existentes: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & "#EFEFEF" & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Lista (cada linha com 20 e-mails)</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
       ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=2 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      i = 0
      j = 0
      While Not RS.EOF
        If i = 0 or j >= 20 Then
           i = 1
           ShowHTML "      <tr valign=""top""><td><font size=""1"">"
           j = 0
        End If
        ShowHTML "        " & RS("email") & "; "
        j = j + 1
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
    ShowHTML "      <TR><TD><font size=""2"" CLASS=""BTM""><b>Nome completo:</b><br><input type=""text"" size=""60"" maxlength=""60"" name=""w_nome"" value=""" & w_nome & """ CLASS=""STI"">"
    ShowHTML "      <TR><TD><font size=""2"" CLASS=""BTM""><b>e-Mail:</b><br><input type=""text"" size=""60"" maxlength=""60"" name=""w_email"" value=""" & w_email & """ CLASS=""STI"">"
    ShowHTML "      <TR><TD><font size=""2"" CLASS=""BTM""><b>Tipo da pessoa:</b> "
    If w_tipo = 1 Then
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""1"" checked> Pai, mãe ou responsável por aluno da rede de ensino "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""2""> Aluno da rede de ensino "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""3""> Outro "
    ElseIf w_tipo = 2 Then
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""1""> Pai, mãe ou responsável por aluno da rede de ensino "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""2"" checked> Aluno da rede de ensino "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""3""> Outro "
    ElseIf w_tipo = 3 Then
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""1""> Pai, mãe ou responsável por aluno da rede de ensino "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""2""> Aluno da rede de ensino "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""3"" checked> Outro "
    Else
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""1""> Pai, mãe ou responsável por aluno da rede de ensino "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""2""> Aluno da rede de ensino "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""3""> Outro "
    End If
    ShowHTML "      <TR> "
    MontaRadioSN "<b>Envia newsletter para este e-mail?</b>", w_envia_mail, "w_envia_mail"
    ShowHTML "      </TR> "
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

  Set w_nome  = Nothing 
  Set w_tipo  = Nothing 
  Set w_email = Nothing 
  Set w_chave = Nothing 
  Set w_troca = Nothing 
  Set i       = Nothing 
  Set w_texto = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de newsletter
REM -------------------------------------------------------------------------

REM =========================================================================
REM Cadastro de calendario base
REM -------------------------------------------------------------------------
Sub GetCalendarioBase
  Dim w_chave, w_dt_ocorrencia, w_ds_ocorrencia, w_ano
  
  Dim w_troca, i, w_texto
  
  w_Chave            = Request("w_Chave")
  w_troca            = Request("w_troca")
  
  If w_troca > "" Then ' Se for recarga da página
     w_dt_ocorrencia = Request("w_dt_ocorrencia")    
     w_ds_ocorrencia = Request("w_ds_ocorrencia")
     w_tipo          = Request("w_tipo")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select a.*, b.nome from escCalendario_Base a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) order by year(dt_ocorrencia) desc, dt_ocorrencia"
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escCalendario_Base where sq_ocorrencia = " & w_chave
     ConectaBD SQL
     w_dt_ocorrencia = FormataDataEdicao(RS("dt_ocorrencia"))
     w_ds_ocorrencia = RS("ds_ocorrencia")
     w_tipo          = RS("sq_tipo_data")
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
        Validate "w_dt_ocorrencia" , "Data"      , "DATA" , "1" , "10" , "10" , "1" , "1"
        Validate "w_ds_ocorrencia" , "Descrição" , ""     , "1" , "2"  , "60" , "1" , "1"
        Validate "w_tipo" , "Tipo" , "SELECT"     , "1" , "1"  , "4" , "" , "1"
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
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro do calendário oficial</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Tipo</font></td>"
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
           ShowHTML "      <tr bgcolor=""#C0C0C0"" valign=""top""><TD colspan=4 align=""center""><font size=2><B>" & year(RS("dt_ocorrencia")) & "</b></font></td></tr>"
           wAno = year(RS("dt_ocorrencia"))
        End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & Mid(FormataDataEdicao(FormatDateTime(RS("dt_ocorrencia"),2)),1,5) & "</td>"
        ShowHTML "        <td><font size=""1"">" & nvl(RS("nome"),"---") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("ds_ocorrencia") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_ocorrencia") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_chave=" & RS("sq_ocorrencia") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
    SQL = "SELECT * FROM escTipo_Data a WHERE a.abrangencia <> 'U' ORDER BY a.nome" & VbCrLf
    ConectaBD SQL
    ShowHTML "          <td><font size=""1""><b>Tipo da ocorrência:</b><br><SELECT CLASS=""STI"" NAME=""w_tipo"">"
    ShowHTML "          <option value=""""> ---"
    While Not RS.EOF
       If cDbl(nvl(RS("sq_tipo_data"),0)) = cDbl(nvl(w_tipo,0)) Then
          ShowHTML "          <option value=""" & RS("sq_tipo_data") & """ SELECTED>" & RS("nome")
       Else
          ShowHTML "          <option value=""" & RS("sq_tipo_data") & """>" & RS("nome")
       End If
       RS.MoveNext
    Wend
    ShowHTML "          </select>"
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

  Set w_ano           = Nothing 
  Set w_ds_ocorrencia = Nothing 
  Set w_dt_ocorrencia = Nothing 
  Set w_chave         = Nothing 
  Set w_troca         = Nothing 
  Set i               = Nothing 
  Set w_texto         = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de calendário base
REM -------------------------------------------------------------------------

REM =========================================================================
REM Cadastro de calendario
REM -------------------------------------------------------------------------
Sub GetCalendario
  Dim w_chave, w_dt_ocorrencia, w_ds_ocorrencia, w_ano
  
  Dim w_troca, i, w_texto
  
  w_Chave            = Request("w_Chave")
  w_troca            = Request("w_troca")
  
  If w_troca > "" Then ' Se for recarga da página
     w_dt_ocorrencia = Request("w_dt_ocorrencia")    
     w_ds_ocorrencia = Request("w_ds_ocorrencia")
     w_tipo          = Request("w_tipo")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select a.*, b.nome from escCalendario_Cliente a left join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) where " & replace(CL,"sq_cliente","sq_site_cliente") & " order by year(dt_ocorrencia) desc, dt_ocorrencia"
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escCalendario_Cliente where sq_ocorrencia = " & w_chave
     ConectaBD SQL
     w_dt_ocorrencia = FormataDataEdicao(RS("dt_ocorrencia"))
     w_ds_ocorrencia = RS("ds_ocorrencia")
     w_tipo          = RS("sq_tipo_data")
     DesconectaBD
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  If InStr("IAEP",O) > 0 Then
     ScriptOpen "JavaScript"
     CheckBranco
     FormataData
     ValidateOpen "Validacao"
     'If InStr("IA",O) > 0 Then
     '   Validate "w_dt_ocorrencia" , "Data"      , "DATA" , "1" , "10" , "10" , "1" , "1"
     '   Validate "w_ds_ocorrencia" , "Descrição" , ""     , "1" , "2"  , "60" , "1" , "1"
     '   Validate "w_tipo" , "Tipo da ocorrência" , "SELECT", "1" , "1"  , "4" , "" , "1"
     'End If
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
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro do calendário da rede de ensino</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Tipo</font></td>"
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
           ShowHTML "      <tr bgcolor=""#C0C0C0"" valign=""top""><TD colspan=4 align=""center""><font size=2><B>" & year(RS("dt_ocorrencia")) & "</b></font></td></tr>"
           wAno = year(RS("dt_ocorrencia"))
        End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td align=""center""><font size=""1"">" & mid(FormataDataEdicao(FormatDateTime(RS("dt_ocorrencia"),2)),1,5) & "</td>"
        ShowHTML "        <td><font size=""1"">" & nvl(RS("nome"),"---") & "</td>"
        ShowHTML "        <td><font size=""1"">" & RS("ds_ocorrencia") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_ocorrencia") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_chave=" & RS("sq_ocorrencia") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
    ShowHTML "          <td><font size=""1""><b><u>D</u>ata:</b><br><input accesskey=""D"" type=""text"" name=""w_dt_ocorrencia"" class=""STI"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormataDataEdicao(FormatDateTime(Nvl(w_dt_ocorrencia,Date()),2)) & """ onKeyDown=""FormataData(this,event);"" ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe a data de ocorrência. O sistema colocará as barras automaticamente.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "          <td><font size=""1""><b>D<u>e</u>scrição:</b><br><input " & w_Disabled & " accesskey=""E"" type=""text"" name=""w_ds_ocorrencia"" class=""STI"" SIZE=""60"" MAXLENGTH=""60"" VALUE=""" & w_ds_ocorrencia & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Descreva a ocorrência.','white')""; ONMOUSEOUT=""kill()""></td>"
    SQL = "SELECT * FROM escTipo_Data a WHERE a.abrangencia <> 'U' ORDER BY a.nome" & VbCrLf
    ConectaBD SQL
    ShowHTML "          <td><font size=""1""><b>Tipo da ocorrência:</b><br><SELECT CLASS=""STI"" NAME=""w_tipo"">"
    ShowHTML "          <option value=""""> ---"
    While Not RS.EOF
       If cDbl(nvl(RS("sq_tipo_data"),0)) = cDbl(nvl(w_tipo,0)) Then
          ShowHTML "          <option value=""" & RS("sq_tipo_data") & """ SELECTED>" & RS("nome")
       Else
          ShowHTML "          <option value=""" & RS("sq_tipo_data") & """>" & RS("nome")
       End If
       RS.MoveNext
    Wend
    ShowHTML "          </select>"
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

  Set w_ano           = Nothing 
  Set w_ds_ocorrencia = Nothing 
  Set w_dt_ocorrencia = Nothing 
  Set w_chave         = Nothing 
  Set w_troca         = Nothing 
  Set i               = Nothing 
  Set w_texto         = Nothing
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
     SQL = "select a.*, b.no_aluno, b.nr_matricula, c.ds_cliente " & VbCrLf & _
           "  from escMensagem_Aluno     a " & VbCrLf & _
           "       inner join escAluno   b on (a.sq_aluno        = b.sq_aluno) " & VbCrLf & _
           "       inner join escCliente c on (b.sq_site_cliente = c.sq_cliente) " & VbCrLf & _
           "order by c.ds_cliente, a.dt_mensagem desc, b.no_aluno, a.in_lida " & VbCrLf
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escMensagem_Aluno where sq_mensagem = " & w_chave
     ConectaBD SQL
     w_dt_mensagem  = FormataDataEdicao(RS("dt_mensagem"))
     w_ds_mensagem  = RS("ds_mensagem")
     w_sq_aluno     = RS("sq_aluno")
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
        Validate "w_dt_mensagem" , "Data"     , "DATA"  , "1" , "10" , "10"   , "1" , "1"
        Validate "w_ds_mensagem" , "Mensagem" , ""      , "1" , "2"  , "4000" , "1" , "1"
        If w_ea = "I" Then
           Validate "w_sq_aluno" , "Aluno"    , "SELECT", "1" , "1"  , "10"   , ""  , "1"
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
  ShowHTML "<B><FONT COLOR=""#000000"">Mensagens a alunos da rede de ensino</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Unidade</font></td>"
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
        ShowHTML "        <td><font size=""1"">" & lcase(RS("ds_cliente")) & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & FormataDataEdicao(FormatDateTime(RS("dt_mensagem"),2)) & "</td>"
        ShowHTML "        <td align=""center"" nowrap><font size=""1"">" & RS("nr_matricula") & "</td>"
        ShowHTML "        <td><font size=""1"">" & lcase(RS("no_aluno")) & "</td>"
        ShowHTML "        <td align=""center""><font size=""1"">" & RS("in_lida") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_mensagem") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_chave=" & RS("sq_mensagem") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
REM Cadastro de modalidades de ensino
REM -------------------------------------------------------------------------
Sub GetEspecialidade
  Dim w_chave, w_ds_especialidade, w_ano, w_nr_ordem
  
  Dim w_troca, i, w_texto
  
  w_Chave = Request("w_Chave")
  w_troca = Request("w_troca")
  
  SQL = "Select * from escEspecialidade order by nr_ordem, ds_especialidade"
  ConectaBD SQL
  If Not RS.EOF Then
     w_texto = "<b>Nºs de ordem em uso para esta subordinação:</b>:<br>" & _
               "<table border=1 width=100% cellpadding=0 cellspacing=0>" & _
               "<tr><td align=center><b><font size=1>Ordem" & _
               "    <td><b><font size=1>Descrição"
     
     While Not RS.EOF
        w_texto = w_texto & "<tr><td valign=top align=center><font size=1>" & RS("nr_ordem") & "<td valign=top><font size=1>" & RS("ds_especialidade")
        RS.MoveNext
     Wend
     w_texto = w_texto & "</table>"
  Else
     w_texto = "Não há outros números de ordem vinculados à subordinação desta opção"
  End If
  DesconectaBD
  
  If w_troca > "" Then ' Se for recarga da página
     w_ds_especialidade = Request("w_ds_especialidade")
     w_nr_ordem         = Request("w_nr_ordem")    
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select * from escEspecialidade order by nr_ordem, ds_especialidade"
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escEspecialidade where sq_especialidade = " & w_chave
     ConectaBD SQL
     w_ds_especialidade = RS("ds_especialidade")
     w_nr_ordem         = RS("nr_ordem")
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
        Validate "w_ds_especialidade" , "Descrição"    , "" , "1" , "2" , "70" , "1" , "1"
        Validate "w_nr_ordem"         , "Nr. de ordem" , "" , "1" , "1" , "4"  , ""  , "0123546789"
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
     BodyOpen "onLoad='document.Form.w_ds_especialidade.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de modalidades de ensino</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Ordem</font></td>"
    ShowHTML "          <td><font size=""1""><b>Modalidade</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
        ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=3 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      w_ano  = ""
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        if Nvl(RS("nr_ordem"),"nulo") <> "nulo" then
           ShowHTML "        <td align=""CENTER""><font size=""1"">" & RS("nr_ordem") & "</td>"
        else
           ShowHTML "        <td align=""center""><font size=""1"">---</td>"
        End If
        ShowHTML "        <td><font size=""1"">" & RS("ds_especialidade") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_especialidade") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_chave=" & RS("sq_especialidade") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
    ShowHTML "        <tr valign=""top""><td valign=""top""><font size=""1""><b>D<u>e</u>scrição:</b><br><input " & w_Disabled & " accesskey=""E"" type=""text"" name=""w_ds_especialidade"" class=""STI"" SIZE=""70"" MAXLENGTH=""70"" VALUE=""" & w_ds_especialidade & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o nome da modalidade de ensino.','white')""; ONMOUSEOUT=""kill()""></td>"
    ShowHTML "              <td valign=""top"" align=""left""><font size=""1""><b><u>O</u>rdem:<br><INPUT ACCESSKEY=""O"" TYPE=""TEXT"" CLASS=""STI"" NAME=""w_nr_ordem"" SIZE=4 MAXLENGTH=4 VALUE=""" & w_nr_ordem & """ " & w_Disabled & " ONFOCUS=""popup1('" & Replace(w_texto,CHR(13)&CHR(10),"<BR>") & "','white')""; ONBLUR=""kill()""></td>"
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

  Set w_ano              = Nothing 
  Set w_ds_especialidade = Nothing 
  Set w_nr_odem          = Nothing 
  Set w_chave            = Nothing 
  Set w_troca            = Nothing 
  Set i                  = Nothing 
  Set w_texto            = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de modalidades de ensino
REM -------------------------------------------------------------------------

REM =========================================================================
REM Cadastro de modalidades de ensino
REM -------------------------------------------------------------------------
Sub GetTipoCliente
  
  Dim w_chave, w_tipo, w_nome
  Dim w_troca, i
  
  w_Chave           = Request("w_Chave")
  w_troca           = Request("w_troca")
  
  If w_troca > "" Then ' Se for recarga da página
     w_tipo         = Request("w_tipo")
     w_nome         = Request("w_nome")
  ElseIf w_ea = "L" Then
     ' Recupera todos os registros para a listagem
     SQL = "select sq_tipo_cliente, ds_tipo_cliente, ds_registro, tipo, " & VbCrLf & _
           "       case tipo when '1' then 'Secretaria' " & VbCrLf & _
           "                 when '2' then 'Regional'   " & VbCrLf & _
           "                 when '3' then 'Escola'     " & VbCrLf & _
           "       end nm_tipo                        " & VbCrLf & _
           "  from escTipo_Cliente                    " & VbCrLf
     ConectaBD SQL
  ElseIf InStr("AEV",w_ea) > 0 and w_Troca = "" Then
     ' Recupera os dados do endereço informado
     SQL = "select * from escTipo_cliente where sq_tipo_Cliente = " & w_chave
     ConectaBD SQL
     w_tipo         = RS("tipo")
     w_nome         = RS("ds_tipo_cliente")
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
        Validate "w_nome" , "Descrição" , "" , "1" , "6" , "60" , "1" , "1"
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
     BodyOpen "onLoad='document.Form.w_nome.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de tipos de Instituições</FONT></B>"
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
    ShowHTML "          <td><font size=""1""><b>Descrição</font></td>"
    ShowHTML "          <td><font size=""1""><b>Tipo</font></td>"
    ShowHTML "          <td><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"

    If RS.EOF Then ' Se não foram selecionados registros, exibe mensagem
       ShowHTML "      <tr bgcolor=""" & "#EFEFEF" & """><td colspan=3 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      w_ano  = ""
      While Not RS.EOF
        If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
        ShowHTML "      <tr bgcolor=""" & w_cor & """ valign=""top"">"
        ShowHTML "        <td><font size=""1"">" & RS("ds_tipo_cliente") & "</td>"
        Select Case RS("tipo")
           case 1
              ShowHTML "          <td align=""center""><font size=""1"">Secretaria</font></td>"
           case 2
              ShowHTML "          <td align=""center""><font size=""1"">Regional</font></td>"
           case 3
              ShowHTML "          <td align=""center""><font size=""1"">Escola</font></td>"
           case else
              ShowHTML "        <td><font size=""1"">---</td>"
        End Select
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & w_ew & "&R=" & w_Pagina & w_ew & "&w_ea=A&CL=" & CL & "&w_chave=" & RS("sq_tipo_cliente") & """>Alterar</A>&nbsp"
        ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_ew & "&w_ea=E&CL=" & CL & "&w_sq_cliente=" & replace(CL,"sq_cliente=","") & "&w_chave=" & RS("sq_tipo_cliente") & """ onClick=""return confirm('Confirma a exclusão do registro?');"">Excluir</A>&nbsp"
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
    ShowHTML "      <tr><td valign=""top"" colspan=""3""><table border=0 width=""100%"" cellspacing=0>"
    ShowHTML "        <tr valign=""top""><td valign=""top""><font size=""1""><b>D<u>e</u>scrição:</b><br><input " & w_Disabled & " accesskey=""E"" type=""text"" name=""w_nome"" class=""STI"" SIZE=""60"" MAXLENGTH=""60"" VALUE=""" & w_nome & """ ONMOUSEOVER=""popup('OBRIGATÓRIO. Informe o nome do tipo da escola.','white')""; ONMOUSEOUT=""kill()""></td></tr>"
    ShowHTML "      <TD><font size=""2"" CLASS=""BTM""><b>Tipo:</b> "
    If     w_tipo = 1 Then
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""1"" checked> Secretaria "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""2""> Regional "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""3""> Escola "
    ElseIf w_tipo = 2 Then
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""1""> Secretaria "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""2"" checked> Regional "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""3""> Escola "
    ElseIf w_tipo = 3 Then
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""1""> Secretaria "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""2""> Regional "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""3"" checked> Escola "
    Else
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""1""> Secretaria de Ensino do DF "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""2""> Rede de Ensino "
       ShowHTML "            <br><input type=""Radio"" name=""w_tipo"" value=""3"" checked> Escola "
    End If
    ShowHTML "        </table>"
    ShowHTML "      <tr>"
    ShowHTML "      <tr><td align=""center"" colspan=3><hr>"
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

  Set w_nome  = Nothing 
  Set w_tipo  = Nothing 
  Set w_chave = Nothing 
  Set w_troca = Nothing 
  Set i       = Nothing 
  Set w_texto = Nothing
End Sub
REM =========================================================================
REM Fim do cadastro de modalidades de ensino
REM -------------------------------------------------------------------------


REM =========================================================================
REM Rotina de criação da tela de logon
REM -------------------------------------------------------------------------
Sub LogOn
  ShowHTML "<HTML>"
  ShowHTML "<HEAD>"
  ShowHTML "<TITLE>SigeWeb - Autenticação</TITLE>"
  ScriptOpen "JavaScript"
  ValidateOpen "Validacao"
  Validate "Login1"    , "Nome de usuário" , ""  , "1" , "4" , "14" , "1" , "1"
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
  ShowHTML "</style> "
  ShowHTML "</HEAD>"
  If wNoUsuario = "" Then
     ShowHTML "<body topmargin=0 leftmargin=10 onLoad=""document.Form.Login1.focus();"">"
  Else
     ShowHTML "<body topmargin=0 leftmargin=10 onLoad=""document.focus();"">"
  End If
  ShowHTML "<CENTER>"
  ShowHTML "<form method=""post"" action=""controle.asp"" onsubmit=""return(Validacao(this));"" name=""Form""> "
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
  ShowHTML "          <tr><td align=""center"" colspan=2><font size=""1"" color=""#990000""><b>Esta aplicação é de uso interno da Secretaria de Estado de Educação.<br>As informações contidas nesta aplicação são restritas e de uso exclusivo.<br>O uso indevido acarretará ao infrator penalidades de acordo com a legislação em vigor.<br><br>Informe seu nome de usuário, senha de acesso e clique no botão <i>OK</i> para ser autenticado pela aplicação.</b></font>"
  ShowHTML "          <tr><td align=""right"" width=""43%""><font size=""2""><b>Nome de usuário:<td><input class=""sti"" name=""Login1"" size=""14"" maxlength=""14"">"
  ShowHTML "          <tr><td align=""right""><font size=""2""><b>Senha:<td><input class=""sti"" type=""Password"" name=""Password1"" size=""19"">"
  ShowHTML "          <tr><td align=""right""><td><font size=""2""><b><input class=""stb"" type=""submit"" value=""OK"" name=""Botao""> "
  ShowHTML "          </font></td> </tr> "
  ShowHTML "          <TR><TD colspan=2 align=""center""><br><table border=0 cellpadding=0 cellspacing=0><tr><td>"
  ShowHTML "              <P><IMG height=37 src=""img/ajuda.jpg"" width=629><br>"
  ShowHTML "              <font face=""Arial"" size=1><b>PARA ACESSAR A PÁGINA DE ATUALIZAÇÃO</b></font>"
  ShowHTML "              <FONT face=""Verdana, Arial, Helvetica, sans-serif"" size=1>"
  ShowHTML "              <li>Nome de usuário - Informe seu nome de usuário"
  ShowHTML "              <li>Senha - Informe sua senha de acesso"
  ShowHTML "              <li>Se esqueceu ou não foi informado dos dados acima, favor entrar em contato com a SEDF / SUBIP / Diretoria de Sistemas de Informação Educacional - DSIE"
  ShowHTML "              </FONT></P>"
  ShowHTML "              <P><font face=""Arial"" size=1><b>DOCUMENTAÇÃO - LEIA COM ATENÇÃO</b></font><br>"
  ShowHTML "              <FONT face=""Verdana"" size=1>"
  ShowHTML "              . <a class=""SS"" href=""sedf/Orientacoes_Acesso.pdf"" target=""_blank"" title=""Abre arquivo que descreve as novas características e funcionalidades do SIGE-WEB."">Apresentação da nova versão do SIGE-WEB (PDF - 130KB - 4 páginas)</a><BR>"
  ShowHTML "              . <a class=""SS"" href=""manuais/operacao/"" target=""_blank"" title=""Exibe manual de operação do SIGE-WEB"">Manual SIGE-WEB (HTML)</A><BR>"
  ShowHTML "              <br></FONT></P>"
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
     SQL = "select count(*) existe " & VbCrLf & _
           "  from escCliente           a " & VbCrLf & _
           "       join escTipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente and " & VbCrLf & _
           "                                  b.tipo            = 1 " & VbCrLf & _
           "                                 ) " & VbCrLf & _
           " where upper(ds_username) = upper('" & w_uid & "')"
     ConectaBD SQL
     If RS("existe") = 0 Then
        w_Erro = 3
     Else
        SQL = "select count(*) existe from escCliente where ds_username = upper('" & w_uid & "') and ds_senha_acesso = '" & w_pwd & "'"
        ConectaBD SQL
        If RS("existe") = 0 Then
           w_Erro = 2
        End If
     End If
  End If
  DesconectaBD
  ScriptOpen "JavaScript"
  If w_erro > 0 Then
     If w_Erro = 1 Then
        ShowHTML "  alert('Usuário inexistente!');"
     ElseIf w_Erro = 2 Then
        ShowHTML "  alert('Senha inválida!');"
     Else
        ShowHTML "  alert('Usuário sem permissão para acessar esta página!');"
     End If
     ShowHTML "  history.back(1);"
  Else
     ' Recupera informações a serem usadas na montagem das telas para o usuário
     SQL = "select * from escCliente where ds_username = '" & w_uid & "' and ds_senha_acesso = '" & w_pwd & "'"
     ConectaBD SQL
     Session("username") = uCase(RS("ds_username"))
     w_cliente           = RS("sq_cliente")
     DesconectaBD
     
     If Session("username") <> "SBPI" Then
        ' Grava o acesso na tabela de log
        SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql) " & VbCrLf & _
              "values ( " & VbCrLf & _
              "         " & w_cliente & ", " & VbCrLf & _
              "         getdate(), " & VbCrLf & _
              "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
              "         0, " & VbCrLf & _
              "         'Usuário """ & w_uid & """ - acesso à tela de manutenção dos dados da rede de ensino.', " & VbCrLf & _
              "         null " & VbCrLf & _
              "       ) " & VbCrLf
        ExecutaSQL(SQL)
     End If

     ShowHTML "  location.href='Controle.asp?w_ew=frames&w_in=1&cl=sq_cliente=" & w_cliente & "';"
  End If
  ScriptClose

  Set w_Erro    = Nothing
  Set w_Cliente = Nothing

End Sub
REM =========================================================================
REM Fim da rotina de autenticação de usuários
REM -------------------------------------------------------------------------

REM =========================================================================
REM Tela de dados administrativos
REM -------------------------------------------------------------------------
Sub ShowAdmin

  Dim w_sq_cliente, w_limpeza, w_merenda, w_banheiro, w_equipamento, w_quantidade
  
  w_disabled = " disabled "

  SQL = "select a.ds_cliente, b.* from escCliente a inner join escCliente_Admin b on (a.sq_cliente = b.sq_cliente) where a." & CL
  ConectaBD SQL
  If RS.EOF Then
     w_sq_cliente = replace(CL,"sq_cliente=","")
  Else
     w_sq_cliente = RS("sq_cliente")
     w_ds_cliente = RS("ds_cliente")
     w_limpeza    = RS("limpeza_terceirizada")
     w_merenda    = RS("merenda_terceirizada")
     w_banheiro   = RS("banheiro")
  End If
  DesconectaBD

  Cabecalho
  ShowHTML "<HEAD>"
  ShowHTML "<TITLE>Ficha de dados administrativos</TITLE>"
  ScriptOpen "Javascript"
  CheckBranco
  FormataDATA
  FormataCEP
  ValidateOpen "Validacao"
  Validate "w_banheiro" , "Quantidade de quadros magnéticos" , "1" , "1" , "1" , "2" , "" , "0123456789"
  ShowHTML "  for (i = 0; i < theForm.w_equipamento.length; i++) {"
  ShowHTML "      if (isNaN(parseInt(theForm.w_equipamento[i].value))) {"
  ShowHTML "         alert('Informe apenas números na quantidade dos equipamentos!');"
  ShowHTML "         theForm.w_equipamento[i].focus();"
  ShowHTML "         return false;"
  ShowHTML "      }"
  ShowHTML "  }"
  ShowHTML "  for (i = 0; i < theForm.w_equipamento.length; i++) {"
  ShowHTML "      if (theForm.w_equipamento[i].value > 0) break;"
  ShowHTML "      if (i == theForm.w_equipamento.length-1) {"
  ShowHTML "         alert('Você deve informar a quantidade de pelo menos um equipamento!');"
  ShowHTML "         return false;"
  ShowHTML "      }"
  ShowHTML "  }"
  ShowHTML "  theForm.Botao.disabled=true;"
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  BodyOpen "onLoad='document.focus();'"
  ShowHTML "<B><FONT COLOR=""#000000"">" & w_ds_cliente & " - Dados administrativos</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""97%"" border=""0"">"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Serviços</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td><font size=1><b><ul>"
  If w_limpeza  = "S"  Then ShowHTML "          <li>O serviço de limpeza da escola é terceirizado."            Else ShowHTML "          <li>O serviço de limpeza da escola não é terceirizado."   End If
  If w_merenda  = "S"  Then ShowHTML "          <li>A escola oferece merenda."                                 Else ShowHTML "          <li>A escola não oferece merenda." End If
  If w_banheiro = "1"  Then ShowHTML "          <li>Na escola existe 1 quadro magnético."                              Else ShowHTML "          <li>Na escola existem " & w_banheiro & " quadros magnéticos."      End If
  ShowHTML "      </ul></u></font></td></tr>"
  ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000""></TD></TR>"
  ShowHTML "      <tr><td align=""center"" height=""2"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  ShowHTML "      <tr><td valign=""top"" align=""center"" bgcolor=""#D0D0D0""><font size=""1""><b>Equipamentos</td></td></tr>"
  ShowHTML "      <tr><td align=""center"" height=""1"" bgcolor=""#000000""></td></tr>"
  SQL = "select a.nome, b.sq_equipamento, b.nome nm_equipamento, IsNull(c.quantidade, 0) quantidade " & VbCrLf & _
        "  from escTipo_Equipamento               a " & VbCrLf & _
        "       inner join escEquipamento         b on (a.sq_tipo_equipamento = b.sq_tipo_equipamento) " & VbCrLf & _
        "       inner join escCliente_Equipamento c on (b.sq_equipamento      = c.sq_equipamento and " & VbCrLf & _
        "                                               c.sq_cliente          = " & replace(CL,"sq_cliente=","") & " " & VbCrLf & _
        "                                              ) " & VbCrLf & _
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
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  ShowHTML "</table>"
  Rodape

  Set w_sq_cliente  = Nothing
  Set w_limpeza     = Nothing 
  Set w_merenda     = Nothing 
  Set w_banheiro    = Nothing 
  Set w_equipamento = Nothing 
  Set w_quantidade  = Nothing 
End Sub
REM =========================================================================
REM Fim da tela de dados administrativos
REM -------------------------------------------------------------------------

REM =========================================================================
REM Monta a tela de Pesquisa
REM -------------------------------------------------------------------------
Public Sub GetVerifArquivo

  Dim RS1, p_regional

  Dim sql, sql2, wCont, sql1, wAtual, wIN, w_especialidade
  
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  
  p_regional = Request("p_regional")

  If p_tipo = "W" Then
      Response.ContentType = "application/msword"
      HeaderWord p_layout
      ShowHTML "<TABLE WIDTH=""100%"" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR=""#000000"">SIGE-WEB<TD ALIGN=""RIGHT""><B><FONT SIZE=4 COLOR=""#000000"">"
      ShowHTML "Consulta a escolas"
      ShowHTML "</FONT><TR><TD ALIGN=""RIGHT""><B><FONT SIZE=2 COLOR=""#000000"">" & DataHora() & "</B></TD></TR>"
      ShowHTML "</FONT></B></TD></TR></TABLE>"
      ShowHTML "<HR>"
  Else
     Cabecalho
     ShowHTML "<HEAD>"
     ScriptOpen("JavaScript")
     ShowHTML " function marca() {"
     ShowHTML "   if (document.Form2.arquivo.length==undefined) {"
     ShowHTML "      if (document.Form2.dummy.checked) {"
     ShowHTML "         document.Form2.arquivo.checked=true;"
     ShowHTML "      } else {"
     ShowHTML "         document.Form2.arquivo.checked=false;"
     ShowHTML "      }"
     ShowHTML "   } else {"
     ShowHTML "      for (var i = 0; i < document.Form2.arquivo.length; i++) {"
     ShowHTML "         if (document.Form2.dummy.checked) {"
     ShowHTML "            document.Form2.arquivo[i].checked=true;"
     ShowHTML "         } else {"
     ShowHTML "            document.Form2.arquivo[i].checked=false;"
     ShowHTML "         }"
     ShowHTML "      }"
     ShowHTML "   }"
     ShowHTML " }"
     ValidateOpen "Validacao2"
     ShowHTML "   if (document.Form2.arquivo.length==undefined) {"
     ShowHTML "     if (!document.Form2.arquivo.checked) {"
     ShowHTML "       alert('Indique pelo menos um arquivo a ser excluído!');"
     ShowHTML "       return false;"
     ShowHTML "     }"
     ShowHTML "   } else {"
     ShowHTML "      var w_erro = true; "
     ShowHTML "      for (var i = 0; i < theForm.arquivo.length; i++) {"
     ShowHTML "         if (theForm.arquivo[i].checked) {"
     ShowHTML "            w_erro = false; "
     ShowHTML "            break;"
     ShowHTML "         }"
     ShowHTML "      }"
     ShowHTML "     if (w_erro) {"
     ShowHTML "       alert('Indique pelo menos um arquivo a ser excluído!');"
     ShowHTML "       return false;"
     ShowHTML "     }"
     ShowHTML "  }"
     ValidateClose
     ScriptClose
     ShowHTML "</HEAD>"
     If Request("pesquisa") > "" Then
        BodyOpen " onLoad=""location.href='#lista'"""
     Else
        BodyOpen "onLoad='document.Form.p_regional.focus()';"
     End If
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
  ShowHTML "    <table width=""95%"" border=""0"">"
  If p_tipo = "H" Then
     Showhtml "<FORM ACTION=""controle.asp"" name=""Form"" METHOD=""POST"">"
     ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""w_ew"" VALUE=""" & w_ew &  """>"
     ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""CL"" VALUE=""" & CL &  """>"
     ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""pesquisa"" VALUE=""X"">"
     ShowHTML "<input type=""Hidden"" name=""P3"" value=""1"">"
     ShowHTML "<input type=""Hidden"" name=""P4"" value=""15"">"
     ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
     ShowHTML "    <table width=""100%"" border=""0"">"
     ShowHTML "          <TR><TD valign=""top""><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
  Else
     ShowHTML "<tr><td><div align=""justify""><font size=1><b>Filtro:</b><ul>"
  End If
  If p_tipo = "H" Then
     ShowHTML "          <tr valign=""top""><td>"
     SelecaoRegional "<u>S</u>ubordinação:", "S", "Indique a subordinação da escola.", p_regional, null, "p_regional", null, null
  ElseIf Nvl(p_regional,0) > 0 Then
     SQL = "SELECT  a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " & VbCrLf & _
           "  FROM  escCLIENTE a " & VbCrLf & _
           " WHERE  a.sq_cliente = " & p_regional & VbCrLf & _
           "ORDER BY a.ds_cliente "
     ConectaBD SQL
     Response.Write "          <li><font size=""1""><b>Escolas da " & RS("ds_cliente") & "</b>"
     DesconectaBD
  End If
  SQL = "SELECT * FROM escTipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
  ConectaBD SQL
  If p_tipo = "H" Then
     ShowHTML "          <tr valign=""top""><td><td><font size=""1""><br><b>Tipo de instituição:</b><br><SELECT CLASS=""STI"" NAME=""p_tipo_cliente"">"
     If RS.RecordCount > 1 Then ShowHTML "          <option value="""">Todos" End If
     While Not RS.EOF
        If cDbl(nvl(RS("sq_tipo_cliente"),0)) = cDbl(nvl(Request("p_tipo_cliente"),0)) Then
           ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """ SELECTED>" & RS("ds_tipo_cliente")
        Else
           ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """>" & RS("ds_tipo_cliente")
        End If
        RS.MoveNext
     Wend
     ShowHTML "          </select>"
  ElseIf nvl(Request("p_tipo_cliente"),0) > 0 Then
     ShowHTML "          <li><font size=""1""><b>Tipo de instituição: "
     While Not RS.EOF
        If cDbl(nvl(RS("sq_tipo_cliente"),0)) = cDbl(nvl(Request("p_tipo_cliente"),0)) Then ShowHTML RS("ds_tipo_cliente") End If
        RS.MoveNext
     Wend
     ShowHTML "</b>"
  End If
  If p_tipo = "H" Then
     ShowHTML "  <TR><TD><TD><font size=""1""><br><b>Lay-out do arquivo Word:</b><br>"
     If nvl(Request("p_layout"),"LANDSCAPE") = "LANDSCAPE" Then ShowHTML "          <input type=""radio"" name=""p_layout"" value=""LANDSCAPE"" CLASS=""BTM"" checked> Paisagem<br>"  Else ShowHTML "          <input type=""radio"" name=""p_layout"" value=""LANDSCAPE"" CLASS=""BTM""> Paisagem<br>"  End If
     If Request("p_layout")                  = "PORTRAIT"  Then ShowHTML "          <input type=""radio"" name=""p_layout"" value=""PORTRAIT"" CLASS=""BTM"" checked> Retrato<br>"    Else ShowHTML "          <input type=""radio"" name=""p_layout"" value=""PORTRAIT"" CLASS=""BTM""> Retrato<br> "   End If
  End If
  If p_tipo = "H" Then
     ShowHTML "  <TR><TD><TD><font size=""1""><br><b>Verificar arquivos que:</b><br>"
     If nvl(Request("E"),"S") = "S" Then ShowHTML "          <input type=""radio"" name=""E"" value=""S"" CLASS=""BTM"" checked> existem somente no banco<br>"     Else ShowHTML "          <input type=""radio"" name=""E"" value=""S"" CLASS=""BTM""> existem somente no banco<br>"      End If
     If Request("E")          = "N" Then ShowHTML "          <input type=""radio"" name=""E"" value=""N"" CLASS=""BTM"" checked> existem somente no file system<br>" Else ShowHTML "          <input type=""radio"" name=""E"" value=""N"" CLASS=""BTM""> existem somente no file system<br> " End If
  ElseIf Request("E") > "" Then
     ShowHTML "  <li><font size=""1""><b>"
     If Request("E") = "S" Then ShowHTML "          Existe somente no banco</b>"        End If
     If Request("E") = "N" Then ShowHTML "          Existe somente no file system</b> " End If
  End If
  If p_tipo <> "W" Then
     ShowHTML "          </table>"
  End If
  DesconectaBD
  wCont = 0
  sql1 = ""

  ' Seleção de etapas/modalidades
  sql = "SELECT DISTINCT a.curso as sq_especialidade, a.curso as ds_especialidade, 1 as nr_ordem, 'M' as tp_especialidade " & VbCrLf & _ 
        " from escTurma_Modalidade                 AS a " & VbCrLf & _ 
        "      INNER JOIN escTurma                 AS c ON (a.serie           = c.ds_serie) " & VbCrLf & _
        "      INNER JOIN escCliente               AS d ON (c.sq_site_cliente = d.sq_cliente) " & VbCrLf & _
        "UNION " & VbCrLf & _ 
        "SELECT DISTINCT cast(a.sq_especialidade as varchar) as sq_especialidade, a.ds_especialidade,  " & VbCrLf & _ 
        "       case a.tp_especialidade when 'J' then '1' else a.nr_ordem end as nr_ordem, " & VbCrLf & _ 
        "       case a.tp_especialidade when 'J' then 'M' else a.tp_especialidade end as tp_especialidade" & VbCrLf & _ 
        " from escEspecialidade AS a " & VbCrLf & _ 
        "      INNER JOIN escEspecialidade_cliente AS c ON (a.sq_especialidade = c.sq_codigo_espec) " & VbCrLf & _
        "      INNER JOIN escCliente               AS d ON (c.sq_cliente       = d.sq_cliente) " & VbCrLf & _
        " where a.tp_especialidade <> 'M' " & VbCrLf & _
        "ORDER BY a.nr_ordem, a.ds_especialidade " & VbCrLf
  ConectaBD sql
   
  If p_tipo = "H" Then
     If Not RS.EOF Then
        wCont = 0
        wAtual = ""
 
        ShowHTML "          <TD valign=""top""><table border=""0"" align=""left"" cellpadding=0 cellspacing=0>"
        Do While Not RS.EOF
           If wAtual = "" or wAtual <> RS("tp_especialidade") Then
              wAtual = RS("tp_especialidade")
              If wAtual = "M" Then
                 ShowHTML "            <TR><TD colspan=2><font size=""1""><b>Etapas/Modalidades de ensino:</b>"
              ElseIf wAtual = "R" Then
                 ShowHTML "            <TR><TD colspan=2><font size=""1""><b>Em Regime de Intercomplementaridade:</b>"
              Else
                 ShowHTML "            <TR><TD colspan=2><font size=""1""><b>Outras:</b>"
              End If
           End If
           
           wCont = wCont + 1
           marcado = "N"
           For i = 1 to Request("p_modalidade").Count
               If RS("sq_especialidade") = Request("p_modalidade")(i) Then 
                  marcado = "S" 
                  sql1 = ",'" & Request("p_modalidade")(i) & "'" & sql1
               End If
           Next
              
           If marcado = "S" Then
              ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""p_modalidade"" value=""" & RS("sq_especialidade") & """ checked><td><font size=1>" & RS("ds_especialidade")
              wIN = 1
           Else
              ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""p_modalidade"" value=""" & RS("sq_especialidade") & """><td><font size=1>" & RS("ds_especialidade")
           End If
           RS.MoveNext
 
           If (wCont Mod 2) = 0 Then 
              wCont = 0
           End If
 
        Loop
        DesconectaBD
        sql1 = Mid(sql1,2)
     End If
  ElseIf Nvl(Request("p_modalidade"), "") > "" Then
     If Not RS.EOF Then
        wCont = 0
 
        ShowHTML "          <li><b>Modalidades de ensino:</b><ul>"
        Do While Not RS.EOF
                    
           ShowHTML chr(13) & "           <li><font size=1>" & RS("ds_especialidade")
           If RS("sq_especialidade") = Request("p_modalidade")(i) Then 
              sql1 = ",'" & Request("p_modalidade")(i) & "'" & sql1
           End If
           wIN = 1
           RS.MoveNext
        Loop
        DesconectaBD
        sql1 = Mid(sql1,2)
     End If
  End If
  ShowHTML "          </tr>"
  ShowHTML "          </table>"
  if p_tipo = "H" Then 
     ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000"">"
     ShowHTML "      <tr><td align=""center"" colspan=""3"">"
     ShowHTML "            <input class=""BTM"" type=""submit"" name=""Botao"" value=""Aplicar filtro"">"
     If Session("username") = "SBPI" Then
        ShowHTML "            <input class=""BTM"" type=""button"" name=""Botao"" onClick=""location.href='" & w_Pagina & "CadastroEscola" & "&CL=" & CL & MontaFiltro("GET") & "&w_ea=I';"" value=""Nova escola"">"
     End If
     ShowHTML "          </td>"
     ShowHTML "      </tr>"
  End If
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  if p_tipo = "H" Then ShowHTML "</form>" End If
  
  ' INÍCIO DA VERIFICAÇÃO
  
  If Request("pesquisa") > "" Then
     If Request("E") = "S" Then
        sql = "SELECT DISTINCT 'ARQUIVO' as tipo, d.ds_username, d.sq_cliente, d.ds_cliente, d.ds_apelido, d.ln_internet, f.ds_diretorio," & VbCrLf & _
              "       h.sq_arquivo as chave, h.ds_titulo, h.nr_ordem, h.ln_arquivo " & VbCrLf & _
              "  from escCliente                                 d " & VbCrLf & _ 
              "       INNER      JOIN escCliente                 b ON (b.sq_cliente       = d.sq_cliente_pai) " & VbCrLf & _
              "       INNER      JOIN escCliente_Site            f ON (d.sq_cliente       = f.sq_cliente) " & VbCrLf & _
              "       INNER      JOIN escTipo_Cliente            g ON (d.sq_tipo_cliente  = g.sq_tipo_cliente) " & VbCrLf & _
              "       INNER      JOIN escCliente_Arquivo         h ON (d.sq_cliente       = h.sq_site_cliente) " & VbCrLf & _
              " where 1 = 1 " & VbCrLf
        If Mid(Session("username"),1,2) = "RE" Then
           SQL = SQL & "   and b.ds_username = '" & Session("USERNAME") & "' " & VbCrLf
        End If
        If Request("D") > "" Then SQL = SQL & "   and d.dt_alteracao     is not null " & VbCrLf End If

        If Request("p_regional") > "" Then 
           sql = sql + "    and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf
        Else
           sql = sql + "    and g.tipo = 3" & VbCrLf
        End If
             
        If Request("p_tipo_cliente") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("p_tipo_cliente")          & VbCrLf End If

        if sql1 > "" then
          sql = sql & _
                "    and (0 < (select count(*) from escEspecialidade_Cliente where sq_cliente = d.sq_cliente and cast(sq_codigo_espec as varchar) in (" + sql1 + ")) or " & VbCrLf & _
                "         0 < (select count(*) from escTurma_Modalidade  w INNER JOIN escTurma x ON (w.serie = x.ds_serie) where x.sq_site_cliente = d.sq_cliente and w.curso in (" + sql1 + ")) " & VbCrLf & _
                "        ) " & VbCrLf 
        end if
        sql = sql & _
              "UNION SELECT DISTINCT 'FOTO' as tipo, d.ds_username, d.sq_cliente, d.ds_cliente, d.ds_apelido, d.ln_internet, f.ds_diretorio," & VbCrLf & _
              "       h.sq_cliente_foto as chave, h.ds_foto as ds_titulo, h.nr_ordem, h.ln_foto as ln_arquivo " & VbCrLf & _
              "  from escCliente                                 d " & VbCrLf & _ 
              "       INNER      JOIN escCliente                 b ON (b.sq_cliente       = d.sq_cliente_pai) " & VbCrLf & _
              "       INNER      JOIN escCliente_Site            f ON (d.sq_cliente       = f.sq_cliente) " & VbCrLf & _
              "       INNER      JOIN escTipo_Cliente            g ON (d.sq_tipo_cliente  = g.sq_tipo_cliente) " & VbCrLf & _
              "       INNER      JOIN escCliente_Foto            h ON (d.sq_cliente       = h.sq_cliente) " & VbCrLf & _
              " where 1 = 1 " & VbCrLf
        If Mid(Session("username"),1,2) = "RE" Then
           SQL = SQL & "   and b.ds_username = '" & Session("USERNAME") & "' " & VbCrLf
        End If
        If Request("D") > "" Then SQL = SQL & "   and d.dt_alteracao     is not null " & VbCrLf End If

        If Request("p_regional") > "" Then 
           sql = sql + "    and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf
        Else
           sql = sql + "    and g.tipo = 3" & VbCrLf
        End If
             
        If Request("p_tipo_cliente") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("p_tipo_cliente")          & VbCrLf End If

        if sql1 > "" then
          sql = sql & _
                "    and (0 < (select count(*) from escEspecialidade_Cliente where sq_cliente = d.sq_cliente and cast(sq_codigo_espec as varchar) in (" + sql1 + ")) or " & VbCrLf & _
                "         0 < (select count(*) from escTurma_Modalidade  w INNER JOIN escTurma x ON (w.serie = x.ds_serie) where x.sq_site_cliente = d.sq_cliente and w.curso in (" + sql1 + ")) " & VbCrLf & _
                "        ) " & VbCrLf 
        end if
        sql = sql + "ORDER BY d.ds_cliente, tipo desc, h.nr_ordem, h.ds_titulo " & VbCrLf
        ConectaBD SQL
        
        ShowHTML "<TR><TD valign=""top""><br><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
        If Not RS.EOF Then

           If p_tipo = "H" Then 
              If Request("P4") > "" Then RS.PageSize = cDbl(Request("P4")) Else RS.PageSize = 15 End If
              rs.AbsolutePage = Nvl(Request("P3"),1)
           Else
              RS.PageSize = RS.RecordCount + 1
              rs.AbsolutePage = 1
           End If
         

           ShowHTML "<tr><td><td align=""right""><b><font face=Verdana size=1></font></b>"
           If p_Tipo = "H" Then ShowHTML "     &nbsp;&nbsp;<A TITLE=""Clique aqui para gerar arquivo Word com a listagem abaixo"" class=""SS"" href=""#""  onClick=""window.open('controle.asp?p_tipo=W&w_ew=" & w_ew & "&Q=" & Request("Q") & "&C=" & Request("C") & "&D=" & Request("D") & "&U=" & Request("U") & w_especialidade & MontaFiltro("GET") & "','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no');"">Gerar Word<IMG ALIGN=""CENTER"" border=0 SRC=""img/word.gif""></A>" End If
           ShowHTML "<tr><td><td>"
           ShowHTML "<table border=""1"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
           AbreForm "Form2", w_Pagina & "Grava", "POST", "return(Validacao2(this));", null
           ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""R"" VALUE=""VERIFBANCO"">"
           ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""CL"" VALUE=""" & CL &  """>"
           ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""pesquisa"" VALUE=""X"">"
           ShowHTML "<input type=""Hidden"" name=""P3"" value=""1"">"
           ShowHTML "<input type=""Hidden"" name=""P4"" value=""15"">"
           
           ShowHTML "<tr align=""center"" valign=""top"">"
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Escola</b></td>"
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Tipo</b></td>"
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Ordem</b></td>"
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Título</b></td>"
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Link</b></td>"
           ShowHTML "    <td id=""checkbox""><font face=""Verdana"" size=""1""><input type=""CHECKBOX"" name=""dummy"" value=""none"" onClick=""marca()""></b></td>"
           ShowHTML "    <script>document.getElementById('checkbox').style.display='none';</script>"
  
           w_cor   = "#FDFDFD"
           w_atual = ""
           checkbox = "unok"

           Set FS = CreateObject("Scripting.FileSystemObject")

           While Not RS.EOF
             strFile = replace(conFilePhysical & "\sedf\" & RS("ds_username") & "\" & RS("ln_arquivo"),"\\","\")

             ' Remove o arquivo, caso ele já exista
             If Not FS.FileExists (strFile) then
                If checkbox <> "ok" Then
                   ShowHTML "    <script>document.getElementById('checkbox').style.display='block';</script>"
                   checkbox = "ok"
                End If
                If w_atual = "" or w_atual <> RS("DS_CLIENTE") Then
                   If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
                   ShowHTML "<tr valign=""top"" bgcolor=""" & w_cor & """>"
                   If Not IsNull (RS("LN_INTERNET")) Then
                      If inStr(lcase(RS("LN_INTERNET")),"http://") > 0 Then
                         ShowHTML "                <a href=""http://" & replace(RS("LN_INTERNET"),"http://","") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b>"
                      Else
                         ShowHTML "                <a href=""" & RS("LN_INTERNET") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b>"
                      End If
                   Else
                      ShowHTML "    <td><font face=""Verdana"" size=""1"">" & RS("DS_CLIENTE") & "</font></td>"
                   End If
                   w_atual = RS("DS_CLIENTE")
                Else
                   ShowHTML "<tr valign=""top"" bgcolor=""" & w_cor & """>"
                   ShowHTML "    <td><font face=""Verdana"" size=""1"">&nbsp;</font></td>"
                End If
                ShowHTML "    <TD align=""center""><font face=""Verdana"" size=""1"">" & RS("tipo") 
                ShowHTML "    <TD align=""center""><font face=""Verdana"" size=""1"">" & RS("nr_ordem") 
                ShowHTML "    <TD><font face=""Verdana"" size=""1"">" & RS("ds_titulo")
                ShowHTML "    <TD><font face=""Verdana"" size=""1""><a href=""" & RS("ln_internet") & "/" & RS("ln_arquivo") & """ target=""_blank"">" & RS("ln_arquivo") & "</a>"
                ShowHTML "<td align=""center"" width=""1%"" nowrap><input type=""checkbox"" name=""arquivo"" value=""" & RS("tipo") & "=|=" & RS("chave") & """></td>"
             End If

             RS.MoveNext
           Wend
    
           ShowHTML "</table>"
           ShowHTML "<tr><td><td colspan=""5"" align=""center""><input type=""SUBMIT"" name=""Botao"" value=""Remover todos os registros indicados"">"
           ShowHTML "</FORM>"
           ShowHTML "<tr><td><td colspan=""5"" align=""center""><hr>"

        Else

           ShowHTML "<TR><TD><TD colspan=""3""><p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<font size=""2""><b>Nenhuma ocorrência encontrada para as opções acima."

        End If
     Else
        sql = "SELECT d.ds_username, d.sq_cliente, d.ds_cliente, d.ds_apelido, d.ln_internet, f.ds_diretorio " & VbCrLf & _
              "  from escCliente                                 d " & VbCrLf & _ 
              "       INNER      JOIN escCliente                 b ON (b.sq_cliente       = d.sq_cliente_pai) " & VbCrLf & _
              "       INNER      JOIN escCliente_Site            f ON (d.sq_cliente       = f.sq_cliente) " & VbCrLf & _
              "       INNER      JOIN escTipo_Cliente            g ON (d.sq_tipo_cliente  = g.sq_tipo_cliente) " & VbCrLf & _
              " where 1 = 1 " & VbCrLf
        If Mid(Session("username"),1,2) = "RE" Then
           SQL = SQL & "   and b.ds_username = '" & Session("USERNAME") & "' " & VbCrLf
        End If
        If Request("D") > "" Then SQL = SQL & "   and d.dt_alteracao     is not null " & VbCrLf End If

        If Request("p_regional") > "" Then 
           sql = sql + "    and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf
        Else
           sql = sql + "    and g.tipo = 3" & VbCrLf
        End If
             
        If Request("p_tipo_cliente") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("p_tipo_cliente")          & VbCrLf End If

        if sql1 > "" then
          sql = sql & _
                "    and (0 < (select count(*) from escEspecialidade_Cliente where sq_cliente = d.sq_cliente and cast(sq_codigo_espec as varchar) in (" + sql1 + ")) or " & VbCrLf & _
                "         0 < (select count(*) from escTurma_Modalidade  w INNER JOIN escTurma x ON (w.serie = x.ds_serie) where x.sq_site_cliente = d.sq_cliente and w.curso in (" + sql1 + ")) " & VbCrLf & _
                "        ) " & VbCrLf 
        end if
        sql = sql + "ORDER BY d.ds_cliente " & VbCrLf
        ConectaBD SQL

        ShowHTML "<TR><TD valign=""top""><br><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
        If Not RS.EOF Then

           If p_tipo = "H" Then 
              If Request("P4") > "" Then RS.PageSize = cDbl(Request("P4")) Else RS.PageSize = 15 End If
              rs.AbsolutePage = Nvl(Request("P3"),1)
           Else
              RS.PageSize = RS.RecordCount + 1
              rs.AbsolutePage = 1
           End If
         

           ShowHTML "<tr><td><td align=""right""><b><font face=Verdana size=1></font></b>"
           If p_Tipo = "H" Then ShowHTML "     &nbsp;&nbsp;<A TITLE=""Clique aqui para gerar arquivo Word com a listagem abaixo"" class=""SS"" href=""#""  onClick=""window.open('controle.asp?p_tipo=W&w_ew=" & w_ew & "&Q=" & Request("Q") & "&C=" & Request("C") & "&D=" & Request("D") & "&U=" & Request("U") & w_especialidade & MontaFiltro("GET") & "','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no');"">Gerar Word<IMG ALIGN=""CENTER"" border=0 SRC=""img/word.gif""></A>" End If
           ShowHTML "<tr><td><td>"
           ShowHTML "<table border=""1"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
           ShowHTML "<tr align=""center"" valign=""top"">"
           AbreForm "Form2", w_Pagina & "Grava", "POST", "return(Validacao2(this));", null
           ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""R"" VALUE=""" & w_ew &  """>"
           ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""CL"" VALUE=""" & CL &  """>"
           ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""pesquisa"" VALUE=""X"">"
           ShowHTML "<input type=""Hidden"" name=""P3"" value=""1"">"
           ShowHTML "<input type=""Hidden"" name=""P4"" value=""15"">"

           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Escola</b></td>"
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Arquivo</b></td>"
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>KB</b></td>"
           ShowHTML "    <td id=""checkbox""><font face=""Verdana"" size=""1""><input type=""CHECKBOX"" name=""dummy"" value=""none"" onClick=""marca()""></b></td>"
           ShowHTML "    <script>document.getElementById('checkbox').style.display='none';</script>"
           
           w_cor   = "#FDFDFD"
           w_atual = ""
           dim checkbox
           checkbox = "unok"

           Set FS = CreateObject("Scripting.FileSystemObject")
           While Not RS.EOF
             strDir = replace(conFilePhysical & "\sedf\" & RS("ds_username") & "\","\\","\")             
             
             dim FileSystem
             dim File
             dim Folder
             dim w_total

             
             set FileSystem = Server.CreateObject("Scripting.FileSystemObject")
             set Folder = FileSystem.GetFolder(strDir)            

             For each File in Folder.files
                If File.name <> "default.asp" Then                        
                   sqlstr = "select coalesce(a.qtd,0) + coalesce(b.qtd,0) as quantidade " & VbCrLf & _
                            "from (select count(*) as qtd from escCliente_Foto where sq_cliente = " & RS("SQ_CLIENTE") & " and ln_foto = '" & File.name & "') as a, " & VbCrLf & _
                            "     (select count(*) as qtd from escCliente_Arquivo where sq_site_cliente = " & RS("SQ_CLIENTE") & " and ln_arquivo = '" & File.name & "') as b " & VbCrLf

                   Set RsQtd = dbms.Execute(sqlstr)
                   If RsQtd("quantidade") = 0 Then
                      If checkbox <> "ok" Then
                          ShowHTML "    <script>document.getElementById('checkbox').style.display='block';</script>"
                          checkbox = "ok"
                      End If
                      If w_atual = "" or w_atual <> RS("DS_CLIENTE") Then
                          If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
                          ShowHTML "<tr valign=""top"" bgcolor=""" & w_cor & """>"
                          ShowHTML "<td width=""1%"" nowrap><font face=""Verdana"" size=""1"">" & RS("DS_CLIENTE") & "<br>" & strDir & "</td>"
                          w_atual = RS("DS_CLIENTE")
                      Else                                            
                          ShowHTML "<tr valign=""top"" bgcolor=""" & w_cor & """>"
                          ShowHTML "<td></td>"
                      End If
                                         
                      ShowHTML "<td><font face=""Verdana"" size=""1"">" & (File.name)
                      ShowHTML "<td align=""right"" width=""1%"" nowrap><font face=""Verdana"" size=""1"">" & formatNumber(File.size/(1024),0) & "&nbsp;&nbsp;"
                      ShowHTML "<td align=""center"" width=""1%"" nowrap><input type=""checkbox"" name=""arquivo"" value=""" & RS("DS_USERNAME") & "=|=" & File.name & """></td>"
                      ShowHTML "</td>"
                      w_total = w_total + File.size
                   End If
                   checkbox = "unok"
                End If
             Next

             RS.MoveNext
           Wend
           ShowHTML "<tr><td colspan=2 align=""right""><font face=""Verdana"" size=""1"">Total<td align=""right"" width=""1%"" nowrap><font face=""Verdana"" size=""1"">" & formatnumber(w_total/1024,0) & "&nbsp;&nbsp;</td>"
           ShowHTML "</table>"
           ShowHTML "<tr><td><td colspan=""5"" align=""center""><input type=""SUBMIT"" name=""Botao"" value=""Remover todos os arquivos indicados"">"
           ShowHTML "</FORM>"
           ShowHTML "<tr><td><td colspan=""5"" align=""center""><hr>"
        Else

           ShowHTML "<TR><TD><TD colspan=""3""><p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<font size=""2""><b>Nenhuma ocorrência encontrada para as opções acima."

        End If
     End If

  End If
  ShowHTML "</TABLE>"
  Set p_regional = Nothing
End Sub

REM =========================================================================
REM Procedimento que executa as operações de BD
REM -------------------------------------------------------------------------
Public Sub Grava
  Dim w_chave, w_sql, w_funcionalidade, w_diretorio, w_imagem, w_arquivo, w_cont, w_extensao, w_atual
  Dim i
  Dim FS, F2, w_linha, w_chave_nova, w_tipo
  Dim w_registros, w_importados, w_rejeitados, w_situacao, w_erro, w_result, field
    
  Cabecalho
  ShowHTML "</HEAD>"
  BodyOpen "onLoad=document.focus();"
  
  ' Recupera o código a ser gravado na tabela de log
  If Instr("REDEPART,VERIFBANCO,VERIFARQ,ESCPARTHOMOLOG,BASE,NEWSLETTER,TIPOCLIENTE,COMPONENTE, VERSAO,CADASTROESCOLA",w_R) > 0 or w_R = conWhatAdmin or w_R = conWhatSGE Then
     w_funcionalidade = "null"
  Else
     SQL = "select sq_funcionalidade from escFuncionalidade where tipo = 1 and codigo = '" & w_R & "'"
     ConectaBD SQL
     w_funcionalidade = RS("sq_funcionalidade")
     DesconectaBD
  End If
  
  Select Case w_R
    Case conWhatCliente
       dbms.BeginTrans()
       SQL = "update escCliente set " & VbCrLf & _
             "   ds_senha_acesso = '" & trim(Request("w_ds_senha_acesso")) & "' " & VbCrLf
       SQL = SQL & _
             "where " & CL & VbCrLf
       ExecutaSQL(SQL)

       If Session("username") <> "SBPI" Then
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         2, " & VbCrLf & _
                "         'Usuário """ & uCase(Session("username")) & """ - atualização da senha de acesso.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       End If
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_sq_cliente=" & Request("w_sq_cliente") & "&w_ea=A&cl=" & cl & "';"
       ScriptClose
  
    Case conWhatDadosAdicionais
       dbms.BeginTrans()
       SQL = "update escCliente_Dados set " & VbCrLf & _
             "   ds_logradouro  = '" & trim(Request("w_ds_logradouro")) & "', " & VbCrLf & _
             "   nr_cep         = '" & Request("w_nr_cep") & "', " & VbCrLf & _
             "   no_contato     = '" & trim(Request("w_no_contato")) & "', " & VbCrLf & _
             "   nr_fone_contato= '" & trim(Request("w_nr_fone_contato")) & "', " & VbCrLf
       If Request("w_no_bairro") > ""        Then SQL = SQL & "   no_bairro        = '" & trim(Request("w_no_bairro")) & "', "         & VbCrLf Else SQL = SQL & "   no_bairro        = null, " & VbCrLf End If
       If Request("w_no_diretor") > ""       Then SQL = SQL & "   no_diretor       = '" & trim(Request("w_no_diretor")) & "', "        & VbCrLf Else SQL = SQL & "   no_diretor       = null, " & VbCrLf End If
       If Request("w_no_secretario") > ""    Then SQL = SQL & "   no_secretario    = '" & trim(Request("w_no_secretario")) & "', "     & VbCrLf Else SQL = SQL & "   no_secretario    = null, " & VbCrLf End If
       If Request("w_nr_cnpj") > ""          Then SQL = SQL & "   nr_cnpj          = '" & Request("w_nr_cnpj") & "', "                 & VbCrLf Else SQL = SQL & "   nr_cnpj          = null, " & VbCrLf End If
       If Request("w_tp_registro") > ""      Then SQL = SQL & "   tp_registro      = '" & Request("w_tp_registro") & "', "             & VbCrLf Else SQL = SQL & "   tp_registro      = null, " & VbCrLf End If
       If Request("w_ds_ato") > ""           Then SQL = SQL & "   ds_ato           = '" & trim(Request("w_ds_ato")) & "', "            & VbCrLf Else SQL = SQL & "   ds_ato           = null, " & VbCrLf End If
       If Request("w_nr_ato") > ""           Then SQL = SQL & "   nr_ato           = '" & trim(Request("w_nr_ato")) & "', "            & VbCrLf Else SQL = SQL & "   nr_ato           = null, " & VbCrLf End If
       If Request("w_dt_ato") > ""           Then SQL = SQL & "   dt_ato           = '" & cDate(trim(Request("w_dt_ato"))) & "', "     & VbCrLf Else SQL = SQL & "   dt_ato           = null, " & VbCrLf End If
       If Request("w_ds_orgao") > ""         Then SQL = SQL & "   ds_orgao         = '" & trim(Request("w_ds_orgao")) & "', "          & VbCrLf Else SQL = SQL & "   ds_orgao         = null, " & VbCrLf End If
       If Request("w_ds_email_contato") > "" Then SQL = SQL & "   ds_email_contato = '" & trim(Request("w_ds_email_contato")) & "', "  & VbCrLf Else SQL = SQL & "   ds_email_contato = null, " & VbCrLf End If
       If Request("w_nr_fax_contato") > ""   Then SQL = SQL & "   nr_fax_contato   = '" & trim(Request("w_nr_fax_contato")) & "' "     & VbCrLf Else SQL = SQL & "   nr_fax_contato   = null "  & VbCrLf End If
       SQL = SQL & _
             "where " & CL & VbCrLf
       ExecutaSQL(SQL)

       If Session("username") <> "SBPI" Then
          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         2, " & VbCrLf & _
                "         'Usuário """ & uCase(Session("username")) & """ - atualização da tela de dados adicionais.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       End If
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=" & w_ea & "&cl=" & cl & "';"
       ScriptClose
  
    Case conWhatSite
       dbms.BeginTrans()
       SQL = "update escCliente_Site set " & VbCrLf & _
             "   no_contato_internet    = '" & trim(Request("w_no_contato_internet")) & "', " & VbCrLf & _
             "   ds_email_internet      = '" & trim(Request("w_ds_email_internet")) & "', " & VbCrLf & _
             "   nr_fone_internet       = '" & trim(Request("w_nr_fone_internet")) & "', " & VbCrLf & _
             "   ds_texto_abertura      = '" & trim(Request("w_ds_texto_abertura")) & "', " & VbCrLf & _
             "   ds_institucional       = '" & trim(Request("w_ds_institucional")) & "', " & VbCrLf
       If Request("w_nr_fax_internet") > ""   Then SQL = SQL & "   nr_fax_internet        = '" & trim(Request("w_nr_fax_internet")) & "', "       & VbCrLf Else SQL = SQL & "   nr_fax_internet        = null, "  & VbCrLf End If
       If Request("w_im_foto_abertura1") > "" Then SQL = SQL & "   im_foto_abertura1      = '" & trim(Request("w_im_foto_abertura1")) & "', "     & VbCrLf Else SQL = SQL & "   im_foto_abertura1      = null, "  & VbCrLf End If
       If Request("w_im_foto_abertura2") > "" Then SQL = SQL & "   im_foto_abertura2      = '" & trim(Request("w_im_foto_abertura2")) & "', "     & VbCrLf Else SQL = SQL & "   im_foto_abertura2      = null, "  & VbCrLf End If
       If Request("w_ds_mensagem") > ""       Then SQL = SQL & "   ds_mensagem            = '" & trim(Request("w_ds_mensagem")) & "' "            & VbCrLf Else SQL = SQL & "   ds_mensagem            = null "   & VbCrLf End If
       SQL = SQL & _
             "where " & CL & VbCrLf
       ExecutaSQL(SQL)

       If Session("username") <> "SBPI" Then
          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         2, " & VbCrLf & _
                "         'Usuário """ & uCase(Session("username")) & """ - atualização da tela de dados do site.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       End If
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=" & w_ea & "&cl=" & cl & "';"
       ScriptClose
  
    Case conWhatEspecialidadeCliente
       dbms.BeginTrans()
       ' Apaga as especialidades existentes para a rede de ensino
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

       If Session("username") <> "SBPI" Then
          ' Grava o acesso na tabela de log
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         2, " & VbCrLf & _
                "         'Usuário """ & uCase(Session("username")) & """ - tela de modalidades de ensino.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       End If
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=" & w_ea & "&cl=" & cl & "';"
       ScriptClose

    Case conWhatDocumento
       If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then
          SQL = "select a.ds_username from escCliente a where a.sq_cliente = 0"
          ConectaBD SQL
          w_diretorio = replace(conFilePhysical & "\" & RS("ds_username") &  "\" & RS("ds_username") & "\","\\","\")
          DesconectaBD
       Elseif Mid(Session("username"),1,2) = "RE" Then
          SQL = "select a.ds_username from escCliente a where a." & CL
          ConectaBD SQL
          w_diretorio = replace(conFilePhysical&"\sedf\" & RS("ds_username") & "\","\\","\")
          DesconectaBD
       Else
          SQL = "select a.ds_username from escCliente a where a." & CL
          ConectaBD SQL
          w_diretorio = replace(conFilePhysical & "\" & RS("ds_username") &  "\" & RS("ds_username") & "\","\\","\")
          DesconectaBD
       End If       
       
       dbms.BeginTrans()
       If w_ea = "I" Then

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
                "     in_ativo,   in_destinatario, nr_ordem,   ds_titulo, pasta) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf
          If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then SQL = SQL & "0," & VbCrLf Else SQL = SQL & "     " & ul.Form("w_sq_cliente") & ", " & VbCrLf End If
          SQL = SQL & _
                "     convert(datetime, '" & ul.Form("w_dt_arquivo") & "',103), " & VbCrLf & _
                "     '" & ul.Form("w_ds_arquivo") & "', " & VbCrLf & _
                "     '" & w_imagem & "', " & VbCrLf & _
                "     '" & ul.Form("w_in_ativo") & "', " & VbCrLf & _
                "     '" & ul.Form("w_in_destinatario") & "', " & VbCrLf & _
                "     '" & ul.Form("w_nr_ordem") & "', " & VbCrLf & _
                "     '" & ul.Form("w_ds_titulo") & "', " & VbCrLf & _
                "     '" & ul.Form("w_pasta") & "' " & VbCrLf & _                
                " )" & VbCrLf
          ExecutaSQL(SQL)

          ' Grava o acesso na tabela de log
          w_sql = SQL
          SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                " values ( " & VbCrLf
          If Session("username") = "IMPRENSA" OR Session("username") = "SBPI" Then SQL = SQL & "0," & VbCrLf Else SQL = SQL & "     " & ul.Form("w_sq_cliente") & ", " & VbCrLf End If
          SQL = SQL & _
                "         getdate(), " & VbCrLf & _
                "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                "         1, " & VbCrLf & _
                "         'Usuário """ & uCase(Session("username")) & """ - atualização de inclusão de arquivos.', " & VbCrLf & _
                "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                "         " & w_funcionalidade & " " & VbCrLf & _
                "       ) " & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "A" Then
          ' Remove o arquivo físico
          SQL = "select ln_arquivo arquivo from escCliente_Arquivo where sq_arquivo = " & ul.Form("w_chave")
          ConectaBD SQL
          w_arquivo = RS("arquivo")
          DesconectaBD

          SQL = " update escCliente_Arquivo set " & VbCrLf & _
                "     ds_titulo       = '" & ul.Form("w_ds_titulo") & "', " & VbCrLf & _
                "     dt_arquivo      = convert(datetime, '" & ul.Form("w_dt_arquivo") & "',103), " & VbCrLf & _
                "     ds_arquivo      = '" & ul.Form("w_ds_arquivo") & "', " & VbCrLf & _
                "     in_ativo        = '" & ul.Form("w_in_ativo") & "', " & VbCrLf & _
                "     in_destinatario = '" & ul.Form("w_in_destinatario") & "', " & VbCrLf & _
                "     nr_ordem        = '" & ul.Form("w_nr_ordem") & "', " & VbCrLf & _
                "     pasta           = '" & ul.Form("w_pasta") & "' " & VbCrLf & _                
                "where sq_arquivo = " & ul.Form("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   " values (" & VbCrLf
             If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then SQL = SQL & "0," & VbCrLf Else SQL = SQL & "     " & ul.Form("w_sq_cliente") & ", " & VbCrLf End If
             SQL = SQL & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         2, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - alteração de arquivo.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If

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
          SQL = "select ln_arquivo from escCliente_Arquivo where sq_arquivo = " & Request("w_chave")
          ConectaBD SQL
          DeleteAFile w_diretorio & RS("ln_arquivo")
          DesconectaBD
          

          SQL = " delete escCliente_Arquivo where sq_arquivo = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   " values (" & VbCrLf
             If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then SQL = SQL & "0," & VbCrLf Else SQL = SQL & "     " & Request("w_sq_cliente") & ", " & VbCrLf End If
             SQL = SQL & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         3, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - exclusão de arquivo.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       End If
       
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & w_R & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
  
    Case "BASE"
       dbms.BeginTrans()
       
       If w_ea = "I" Then
          ' Recupera o valor da próxima chave primária
          SQL = "select IsNull(max(sq_ocorrencia),0)+1 chave from escCalendario_Base" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD
                    
          ' Insere o arquivo
          SQL = " insert into escCalendario_Base " & VbCrLf & _
                "    (sq_ocorrencia, in_abrangencia, dt_ocorrencia, ds_ocorrencia, sq_tipo_data) " & VbCrLf & _
                " values ( " & w_chave & ", 'T', " & VbCrLf & _
                "     convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_ocorrencia"),2)) & "',103), " & VbCrLf & _
                "     '" & Request("w_ds_ocorrencia") & "', " & VbCrLf & _
                "     '" & Request("w_tipo") & "' " & VbCrLf & _
                " )" & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         1, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - inclusão de data no calendário base.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "A" Then
          SQL = " update escCalendario_Base set " & VbCrLf & _
                "     dt_ocorrencia  = convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_ocorrencia"),2)) & "',103), " & VbCrLf & _
                "     ds_ocorrencia  = '" & Request("w_ds_ocorrencia") & "', " & VbCrLf & _
                "     sq_tipo_data   = '" & Request("w_tipo") & "' " & VbCrLf & _
                "where sq_ocorrencia = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         2, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - alteração de data no calendário base.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "E" Then
          SQL = " delete escCalendario_Base where sq_ocorrencia = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         3, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - exclusão de data no calendário base.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       End If
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
  
    Case conWhatAdmin
       If Request("w_arquivo") = "Escola" Then
          SQL = "select count(*) from escCliente_Admin"
       Else
          SQL = "select count(*) from escEquipamento"
       End If
       ConectaBD SQL
       If RS.EOF Then
          ScriptOpen "JavaScript"
          ShowHTML "  alert('Não há informações a serem exportadas para a opção selecionada!);"
          ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
          ScriptClose
       Else
          DesconectaBD
          CONST ForReading = 1, ForWriting = 2, ForAppend = 8
          CONST TristateUsedefault = -2 'Abre o arquivo usando o sistema default
          CONST TristateTrue = -1 'Abre o arquivo como Unicode
          CONST TristateFalse = 0 'Abre o arquivo como ASCII
          Dim f1, w_Line, w_File, w_Arq
          w_Dir = "sedf\sedf"
          If Request("w_arquivo") = "Escola" Then
             w_Arq = "Escola_" & replace(Date(),"/","") & "_" & replace(Time(),":","") & ".csv"
          Else
             w_Arq = "Equipamento_" & replace(Date(),"/","") & "_" & replace(Time(),":","") & ".csv"
          End If
          w_File = replace(conFilePhysical & "\" & w_Dir & "\" & w_Arq,"\\","\")
          Set FS = CreateObject("Scripting.FileSystemObject")
          Set F1 = FS.CreateTextFile(w_File)
          If Request("w_arquivo") = "Escola" Then
             f1.WriteLine "Estrutura de dados:"
             f1.WriteLine "cd_escola; Código da unidade de ensino"
             f1.WriteLine "nm_escola; Nome da unidade de ensino"
             f1.WriteLine "st_limpeza_terceirizada; S - Limpeza com pessoal terceirizado, N - Limpeza com pessoal próprio"
             f1.WriteLine "st_oferece_merenda; S - A escola oferece merenda, N - A escola não oferece merenda"
             f1.WriteLine "qt_banheiro; Quantidade de quadros magnéticos informada pela unidade de ensino"
             f1.WriteLine "cd_tipo_equipamento; Código do tipo de equipamento"
             f1.WriteLine "nm_tipo_equipamento; Nome do tipo de equipamento"
             f1.WriteLine "st_tipo_equipamento; Situação do tipo de equipamento: S-Ativo, N-Inativo"
             f1.WriteLine "cd_equipamento; código do equipamento"
             f1.WriteLine "nm_equipamento; Nome do equipamento"
             f1.WriteLine "st_equipamento; Situação do equipamento: S-Ativo, N-Inativo"
             f1.WriteLine "qt_equipamento; Quantidade existente do equipamento"
             f1.WriteLine ""
             f1.WriteLine "Observações:"
             f1.WriteLine "1 - A lista está ordenada pelo nome da escola (1), pelo nome do tipo de equipamento (2) e pelo nome do equipamento (3);"
             f1.WriteLine "2 - As 5 primeiras colunas são duplicadas para cada um dos equipamentos da unidade de ensino;"
             f1.WriteLine "3 - Equipamentos ou tipos de equipamentos inativos não aparecem para a unidade de ensino informar sua quantidade."
             f1.WriteLine ""
             f1.WriteLine ""

             ' Recupera os dados a serem exportados
             SQL = "select f.sq_siscol cd_escola, e.ds_cliente nm_escola, " & VbCrLf & _
                   "       d.limpeza_terceirizada st_limpeza_terceirizada, d.merenda_terceirizada st_oferece_merenda,  " & VbCrLf & _
                   "       d.banheiro qt_banheiro, " & VbCrLf & _
                   "       a.codigo cd_tipo_equipamento, a.nome nm_tipo_equipamento, a.ativo st_tipo_equipamento, " & VbCrLf & _
                   "       b.codigo cd_equipamento,      b.nome nm_equipamento,      b.ativo st_equipamento,  " & VbCrLf & _
                   "       c.quantidade qt_equipamento " & VbCrLf & _
                   "  from escCliente                                   e " & VbCrLf & _
                   "       inner            join escCliente_Site        f on (e.sq_cliente          = f.sq_cliente) " & VbCrLf & _
                   "       inner            join escCliente_Admin       d on (e.sq_cliente          = d.sq_cliente) " & VbCrLf & _
                   "         left outer     join escCliente_Equipamento c on (d.sq_cliente          = c.sq_cliente) " & VbCrLf & _
                   "           left outer   join escEquipamento         b on (c.sq_equipamento      = b.sq_equipamento) " & VbCrLf & _
                   "             left outer join escTipo_Equipamento    a on (b.sq_tipo_equipamento = a.sq_tipo_equipamento) " & VbCrLf & _
                   "order by e.ds_cliente, a.nome, b.nome " & VbCrLf
          Else
             f1.WriteLine "Estrutura de dados:"
             f1.WriteLine "cd_tipo_equipamento; Código do tipo de equipamento"
             f1.WriteLine "nm_tipo_equipamento; Nome do tipo de equipamento"
             f1.WriteLine "st_tipo_equipamento; Situação do tipo de equipamento: S-Ativo, N-Inativo"
             f1.WriteLine "cd_equipamento; código do equipamento"
             f1.WriteLine "nm_equipamento; Nome do equipamento"
             f1.WriteLine "st_equipamento; Situação do equipamento: S-Ativo, N-Inativo"
             f1.WriteLine ""
             f1.WriteLine "Observações:"
             f1.WriteLine "1 - A lista está ordenada pelo pelo nome do tipo de equipamento (1) e pelo nome do equipamento (2);"
             f1.WriteLine "2 - Equipamentos ou tipos de equipamentos inativos não aparecem para a unidade de ensino informar sua quantidade."
             f1.WriteLine ""
             f1.WriteLine ""
             
             ' Recupera os dados a serem exportados
             SQL = "select a.codigo cd_tipo_equipamento, a.nome nm_tipo_equipamento, a.ativo st_tipo_equipamento, " & VbCrLf & _
                   "       b.codigo cd_equipamento,      b.nome nm_equipamento,      b.ativo st_equipamento " & VbCrLf & _
                   "  from escTipo_Equipamento a " & VbCrLf & _
                   "       inner join escEquipamento b on (a.sq_tipo_equipamento = b.sq_tipo_equipamento) " & VbCrLf & _
                   "order by a.nome, b.nome " & VbCrLf
          End If
          ConectaBD SQL

          ' Grava uma linha com o nome das colunas
          strLine = ""
          For EACH colName IN RS.fields
             strLine = strLine & colName.name & ";"
          Next
          f1.WriteLine strLine

          ' Grava uma linha para cada registro encontrado
          RS.MoveFirst
          While Not RS.EOF
             strLine = ""
             For EACH colName IN rs.Fields
                strLine = strLine & Nvl(colName.Value,"") & ";"
             Next
             f1.WriteLine strLine
             RS.MoveNext
          Wend
          DesconectaBD
          f1.Close
    
          ScriptOpen "JavaScript"
          ShowHTML "  window.open('" & conVirtualPath & "sedf/sedf/" & w_arq & "','Download','');"
          ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
          ScriptClose
          
          Set FS     = Nothing
          Set F1     = Nothing
          Set w_Line = Nothing
          Set w_File = Nothing
          Set w_Arq  = Nothing
       End If
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
                "    (sq_ocorrencia, sq_site_cliente, dt_ocorrencia, ds_ocorrencia, sq_tipo_data) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf & _
                "     " & Request("w_sq_cliente") & ", " & VbCrLf & _
                "     convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_ocorrencia"),2)) & "',103), " & VbCrLf & _
                "     '" & Request("w_ds_ocorrencia") & "', '" & Request("w_tipo") & "' " & VbCrLf & _
                " )" & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         1, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - inclusão de data no calendário da rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "A" Then
          SQL = " update escCalendario_Cliente set " & VbCrLf & _
                "     dt_ocorrencia  = convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_ocorrencia"),2)) & "',103), " & VbCrLf & _
                "     ds_ocorrencia  = '" & Request("w_ds_ocorrencia") & "', " & VbCrLf & _
                "     sq_tipo_data  = '" & Request("w_tipo") & "' " & VbCrLf & _
                "where sq_ocorrencia = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         2, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - alteração de data no calendário da rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "E" Then
          SQL = " delete escCalendario_Cliente where sq_ocorrencia = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         3, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - exclusão de data no calendário da rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       End If
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
  
    Case "NEWSLETTER"
       dbms.BeginTrans()
       
       If w_ea = "I" Then
          ' Insere o registro
          SQL = "insert into escNewsletter (sq_cliente, nome, email, tipo, envia_mail, data_inclusao, data_alteracao) " & VbCrLf & _
                "  values ( 0," & VbCrLf & _
                "          '" & trim(Request("w_nome")) & "', " & VbCrLf & _
                "          '" & trim(Request("w_email")) & "', " & VbCrLf & _
                "          '" & Request("w_tipo") & "', " & VbCrLf & _
                "          '" & Mid(Request("w_envia_mail"),1,1) & "', " & VbCrLf & _
                "          getdate(), " & VbCrLf & _
                "          null " & VbCrLf & _
                "         )" & VbCrLf
       ElseIf w_ea = "A" Then
          SQL = "update escNewsletter set " & VbCrLf & _
                "   envia_mail     = '" & Mid(Request("w_envia_mail"),1,1) & "', " & VbCrLf & _
                "   nome           = '" & trim(Request("w_nome")) & "', " & VbCrLf & _
                "   email          = '" & trim(Request("w_email")) & "', " & VbCrLf & _
                "   tipo           = '" & Request("w_tipo") & "', " & VbCrLf & _
                "   data_alteracao = getdate() " & VbCrLf & _
                "where sq_newsletter = " & Request("w_chave") & VbCrLf
       ElseIf w_ea = "E" Then
          SQL = " delete escNewsletter where sq_newsletter = " & Request("w_chave") & VbCrLf
       End If
       ExecutaSQL(SQL)
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
    
    Case "TIPOCLIENTE"
       dbms.BeginTrans()
       
       If w_ea = "I" Then
          SQL = "select IsNull(max(sq_tipo_cliente),0)+1 chave from escTipo_Cliente" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD
          ' Insere o registro
          SQL = "insert into escTipo_Cliente (sq_tipo_cliente, ds_tipo_cliente, ds_registro, tipo) " & VbCrLf & _
                " values ( " & w_chave           & " ," & VbCrLf & _
                "         '" & Request("w_nome") & "'," & VbCrLf & _
                "              null                  ," & VbCrLf & _
                "          " & Request("w_tipo") & "  " & VbCrLf & _
                "         )" & VbCrLf
       ElseIf w_ea = "A" Then
          SQL = "update escTipo_Cliente set " & VbCrLf & _
                "   ds_tipo_Cliente    = '" & (Request("w_nome")) & "'," & VbCrLf & _
                "   ds_registro        =     null                     ," & VbCrLf & _
                "   tipo               = " & Request("w_tipo")           & VbCrLf & _
                "where sq_tipo_cliente = " & Request("w_chave")          & VbCrLf
       ElseIf w_ea = "E" Then
          SQL = " delete escTipo_Cliente where sq_tipo_cliente = " & Request("w_chave") & VbCrLf
       End If
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
                "    (sq_noticia, sq_site_cliente, dt_noticia, ds_titulo, ds_noticia, in_ativo, ln_externo) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf
          If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then SQL = SQL & "0," & VbCrLf Else SQL = SQL & "     " & Request("w_sq_cliente") & ", " & VbCrLf End If
          SQL = SQL & _
                "     convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_noticia"),2)) & "',103), " & VbCrLf & _
                "     '" & Request("w_ds_titulo") & "', " & VbCrLf & _
                "     '" & Request("w_ds_noticia") & "', " & VbCrLf & _
                "     '" & Request("w_in_ativo") & "', " & VbCrLf & _
                "     '" & Request("w_ln_externo") & "' " & VbCrLf & _                
                " )" & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf
             If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then SQL = SQL & "0," & VbCrLf Else SQL = SQL & "     " & Request("w_sq_cliente") & ", " & VbCrLf End If
             SQL = SQL & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         1, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - inclusão de notícia da rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "A" Then
          SQL = " update escNoticia_Cliente set " & VbCrLf & _
                "     dt_noticia  = convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_noticia"),2)) & "',103), " & VbCrLf & _
                "     ds_titulo   = '" & Request("w_ds_titulo") & "', " & VbCrLf & _
                "     ds_noticia  = '" & Request("w_ds_noticia") & "', " & VbCrLf & _
                "     in_ativo    = '" & Request("w_in_ativo") & "', " & VbCrLf & _
                "     ln_externo  = '" & Request("w_ln_externo") & "' " & VbCrLf & _
                "where sq_noticia = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf
             If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then SQL = SQL & "0," & VbCrLf Else SQL = SQL & "     " & Request("w_sq_cliente") & ", " & VbCrLf End If
             SQL = SQL & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         2, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - alteração de notícia da rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "E" Then
          SQL = " delete escNoticia_Cliente where sq_noticia = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf
             If Session("username") = "IMPRENSA" or Session("username") = "SBPI" Then SQL = SQL & "0," & VbCrLf Else SQL = SQL & "     " & Request("w_sq_cliente") & ", " & VbCrLf End If
             SQL = SQL & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         3, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - exclusão de notícia da rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       End If
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
  
    Case conWhatSGE
       dbms.BeginTrans()
       
       If w_ea = "I" Then
          ' Insere o arquivo
          SQL = " insert into escSistema " & VbCrLf & _
                "    (sq_cliente, sg_sistema, no_sistema, ds_sistema, ativo) " & VbCrLf & _
                " values (0, " & VbCrLf & _
                "     '" & Request("w_sg_sistema") & "', " & VbCrLf & _
                "     '" & Request("w_no_sistema") & "', " & VbCrLf & _
                "     '" & Request("w_ds_sistema") & "', " & VbCrLf & _
                "     '" & Request("w_ativo") & "' " & VbCrLf & _
                " )" & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "A" Then
          SQL = " update escSistema set " & VbCrLf & _
                "     sg_sistema  = '" & Request("w_sg_sistema") & "', " & VbCrLf & _
                "     no_sistema   = '" & Request("w_no_sistema") & "', " & VbCrLf & _
                "     ds_sistema  = '" & Request("w_ds_sistema") & "', " & VbCrLf & _
                "     ativo    = '" & Request("w_ativo") & "' " & VbCrLf & _
                "where sq_sistema = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "E" Then
          SQL = " delete escSistema where sq_sistema = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)
       End If
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L';"
       ScriptClose
  
    Case "COMPONENTE"
       dbms.BeginTrans()
       
       If w_ea = "I" Then
          ' Insere o arquivo
          SQL = " insert into escComponente " & VbCrLf & _
                "    (sq_sistema, no_fisico, ds_componente, ativo) " & VbCrLf & _
                " values ( " & VbCrLf & _
                "     '" & Request("w_sq_sistema") & "', " & VbCrLf & _
                "     '" & Request("w_no_fisico") & "', " & VbCrLf & _
                "     '" & Request("w_ds_componente") & "', " & VbCrLf & _
                "     '" & Request("w_ativo") & "' " & VbCrLf & _
                " )" & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "A" Then
          SQL = " update escComponente set " & VbCrLf & _
                "     sq_sistema  = '" & Request("w_sq_sistema") & "', " & VbCrLf & _
                "     no_fisico   = '" & Request("w_no_fisico") & "', " & VbCrLf & _
                "     ds_componente  = '" & Request("w_ds_componente") & "', " & VbCrLf & _
                "     ativo    = '" & Request("w_ativo") & "' " & VbCrLf & _
                "where sq_componente = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)
       ElseIf w_ea = "E" Then
          SQL = " delete escComponente where sq_componente = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)
       End If
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L';"
       ScriptClose
  
    Case "VERSAO"
       dbms.BeginTrans()
       
       w_diretorio = replace(conFilePhysical & "\sge\","\\","\")
       
       If w_ea = "I" Then
          w_arquivo = extractFileName(ul.Files("w_no_arquivo").OriginalPath)
          w_extensao = Mid(w_arquivo,InStr(w_arquivo,"."),50)

          ' Insere o arquivo
          SQL = "set nocount on " & VbCrLf & _
                "insert into escComponente_Versao " & VbCrLf & _
                "    (sq_componente, cd_versao, ds_versao, dt_versao, ds_caminho, no_arquivo) " & VbCrLf & _
                "values ( " & VbCrLf & _
                "     '" & ul.Form("w_sq_componente") & "', " & VbCrLf & _
                "     '" & ul.Form("w_cd_versao") & "', " & VbCrLf & _
                "     '" & ul.Form("w_ds_versao") & "', " & VbCrLf & _
                "     convert(datetime, '" & FormataDataEdicao(FormatDateTime(ul.Form("w_dt_versao"),2)) & "',103), " & VbCrLf & _
                "     '" & ul.Form("w_ds_caminho") & "', " & VbCrLf & _
                "     '" & w_arquivo & "' " & VbCrLf & _
                ") " & VbCrLf & _
                "select @@identity as chave " & VbCrLf & _
                "set nocount off " & VbCrLf
          Set RS = dbms.execute(SQL)

          ul.Files("w_no_arquivo").SaveAs(w_diretorio & RS("chave") & w_extensao)
          DesconectaBD
       ElseIf w_ea = "A" Then
          ' Recupera o nome do arquivo físico, caso precise alterá-lo
          SQL = "select no_arquivo arquivo from escComponente_Versao where sq_componente_versao = " & ul.Form("w_chave")
          ConectaBD SQL
          w_arquivo = ul.Form("w_chave") & Mid(RS("arquivo"), Instr(RS("arquivo"),"."), 50)
          DesconectaBD

          SQL = " update escComponente_Versao set " & VbCrLf & _
                "     sq_componente = '" & ul.Form("w_sq_componente") & "', " & VbCrLf & _
                "     cd_versao     = '" & ul.Form("w_cd_versao") & "', " & VbCrLf & _
                "     ds_versao     = '" & ul.Form("w_ds_versao") & "', " & VbCrLf & _
                "     dt_versao     = convert(datetime, '" & FormataDataEdicao(FormatDateTime(ul.Form("w_dt_versao"),2)) & "',103), " & VbCrLf & _
                "     ds_caminho    = '" & ul.Form("w_ds_caminho") & "' " & VbCrLf & _
                "where sq_componente_versao = " & ul.Form("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If ul.Files("w_no_arquivo").Size > 0 Then 
             ' Remove o arquivo físico
             DeleteAFile w_diretorio & w_arquivo
          
             w_arquivo = extractFileName(ul.Files("w_no_arquivo").OriginalPath)
             w_extensao = Mid(w_arquivo,InStr(w_arquivo,"."),50)
             ul.Files("w_no_arquivo").SaveAs(w_diretorio & ul.Form("w_chave") & w_extensao)
             
             SQL = " update escComponente_Versao set " & VbCrLf & _
                   "     no_arquivo = '" & w_arquivo & "' " & VbCrLf & _
                   "where sq_componente_versao = " & ul.Form("w_chave") & VbCrLf
             ExecutaSQL(SQL)
          End If

       ElseIf w_ea = "E" Then
          ' Remove o arquivo físico
          SQL = "select no_arquivo arquivo from escComponente_Versao where sq_componente_versao = " & Request("w_chave")
          ConectaBD SQL
          DeleteAFile w_diretorio & Request("w_chave") & Mid(RS("arquivo"), Instr(RS("arquivo"),"."), 50)
          DesconectaBD
          

          ' Remove o registro
          SQL = " delete escComponente_Versao where sq_componente_versao = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)
          dbms.CommitTrans()

          ScriptOpen "JavaScript"
          ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L';"
          ScriptClose
          Response.End()
          Exit Sub

       ElseIf w_ea = "G" Then
          strDir = "sge\"
          strArq = "lista.xml"
          strFile = replace(conFilePhysical & "\" & strDir & "\" & strArq,"\\","\")
          strDir = replace(conFilePhysical & "\" & strDir,"\\","\")
          Set FS = CreateObject("Scripting.FileSystemObject")

          ' Garante a existência do diretório
          If Not (FS.FolderExists (strDir)) then
             Set f1 = FS.CreateFolder(strDir)
          End If

          ' Remove o arquivo, caso ele já exista
          If FS.FileExists (strFile) then
             FS.DeleteFile (strFile)
          End If
                      
          Set f1 = FS.CreateTextFile(strFile)
          f1.WriteLine "<?xml version=""1.0"" encoding=""iso-8859-1""?>"
          f1.WriteLine " <!--" 
          f1.WriteLine "     Após a instalação de cada um dos componentes abaixo, URL_confirmacao deve "
          f1.WriteLine "     ser executada, substituindo 9999 pelo código da escola."
          f1.WriteLine "-->"
          w_atual = ""
          ScriptOpen "JavaScript"
          ' Recupera as versões mais atuais para a geração do arquivo
          SQL = "select a.sq_componente_versao, c.sg_sistema, b.no_fisico, a.ds_caminho, a.no_arquivo, a.cd_versao, a.dt_versao " & VbCrLf & _
                "  from escComponente_Versao     a " & VbCrLf & _
                "       inner join escComponente b on (a.sq_componente = b.sq_componente and " & VbCrLf & _
                "                                      b.ativo         = 'Sim' " & VbCrLf & _
                "                                     ) " & VbCrLf & _
                "       inner join escSistema    c on (b.sq_sistema    = c.sq_sistema and " & VbCrLf & _
                "                                      c.ativo         = 'Sim' " & VbCrLf & _
                "                                     ) " & VbCrLf & _
                "       inner join (select sq_componente, max(replace(cd_versao,'.','')) versao " & VbCrLf & _
                "                     from escComponente_Versao " & VbCrLf & _
                "                   group by sq_componente " & VbCrLf & _
                "                  )             d on (a.sq_componente = d.sq_componente and " & VbCrLf & _
                "                                      replace(a.cd_versao,'.','') = d.versao " & VbCrLf & _
                "                                     ) " & VbCrLf & _
                "order by c.sg_sistema, b.no_fisico, ds_caminho,a.dt_versao desc, replace(a.cd_versao,'.','') desc " & VbCrLf
          ConectaBD SQL
          If Not RS.EOF Then
             w_cont = RS.RecordCount
             w_atual = RS("sg_sistema")
             f1.WriteLine "<" & w_atual & ">"
             While Not RS.EOF
                If w_atual <> RS("sg_sistema") Then
                   f1.WriteLine " </" & w_atual & ">"
                   w_atual = RS("sg_sistema")
                   f1.WriteLine " <" & w_atual & ">"
                End If
                f1.WriteLine "  <componente>"
                f1.WriteLine "   <chave_versao>" & RS("sq_componente_versao") & "</chave_versao>"
                f1.WriteLine "   <nome_componente>" & RS("no_fisico") & "</nome_componente>"
                f1.WriteLine "   <versao>" & RS("cd_versao") & "</versao>"
                f1.WriteLine "   <data_construcao>" & FormataDataEdicao(FormatDateTime(RS("dt_versao"),2)) & "</data_construcao>"
                f1.WriteLine "   <URL_arquivo>" & RS("ds_caminho") & RS("sq_componente_versao") & Mid(RS("no_arquivo"), Instr(RS("no_arquivo"),"."), 50) & "</URL_arquivo>"
                f1.WriteLine "   <URL_confirmacao>" & RS("ds_caminho") & "versao.asp?versao=" & RS("sq_componente_versao") & "&amp;escola=9999</URL_confirmacao>"
                f1.WriteLine "  </componente>"
                RS.MoveNext
             Wend
             f1.WriteLine "</" & w_atual & ">"
          End If
          DesconectaBD
          f1.Close
          ShowHTML "  alert('Nova versão do arquivo " & strArq & " concluída com sucesso.\nComponente(s) relacionados: " & Nvl(w_cont,0) & ".');"

          dbms.CommitTrans()

          ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L';"
          ScriptClose
          Response.End()
          Exit Sub
       End If
       dbms.CommitTrans()

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & ul.Form("R") & "&w_ea=L';"
       ScriptClose

    Case "REDEPART"
       w_diretorio = replace(conFilePhysical & "\sge\","\\","\")
       
       If ul.Files("w_no_arquivo").Size > 0 Then 
          ' Remove o arquivo físico
          DeleteAFile w_diretorio & w_arquivo
          
          w_arquivo = extractFileName(ul.Files("w_no_arquivo").OriginalPath)
          ul.Files("w_no_arquivo").SaveAs(w_diretorio & w_arquivo)


          ' Gera o arquivo registro da importação
          Set FS = CreateObject("Scripting.FileSystemObject")
          
          'Abre o arquivo recebido para gerar o arquivo registro
          Set F2 = FS.OpenTextFile(w_diretorio & w_arquivo)
          
          ' Varre o arquivo recebido, linha a linha
          w_registros  = 0
          w_importados = 0
          w_rejeitados = 0
          w_cont       = 0

          Dim sq_cliente, idCursos, Curso, idProfissionais, idEscola, Pasta, Parecer, Portaria, Observacao, delimitador
          Dim idInstituicao, idDiretor, idMantenedora, idEndereco, idTelefone_1, idFax, idVencimento
          Dim idCodinep, idTelefone_2, idEmail_1, idEmail_2, idLocalizacao, idCnpjExecutora, idCnpjEscola, idOrdemServico
          Dim idSecretario, idCep, idAutHabSecretario, idEinf, idEf, idEja, idEm, idEDA, idEprof, idRegiao
          Dim idEndMantenedora, SiteEscola, Situacao, idCategoria
          Dim NomeArquivo, idArquivos, Descricao, idOS, idPortaria, Numero, Data, Dodf, PagDodf, DataDodf
          Dim login, senha
          Dim w_campo(50)
          
          delimitador = """"

          'Abre transação para garantir que os dados da escola estarão integros
          dbms.BeginTrans()
          
        'Remove os registros vinculados à escola
        SQL = "DELETE FROM escParticular_Portaria"
        ExecutaSQL(SQL)

        SQL = "DELETE FROM escParticular_OS"
        ExecutaSQL(SQL)

        SQL = "DELETE FROM escParticular_Curso"
        ExecutaSQL(SQL)

        SQL = "DELETE FROM escCurso"
        ExecutaSQL(SQL)

          Do While Not F2.AtEndOfStream
             w_linha = replace(trim(F2.ReadLine),"\""","`")

             if len(w_linha) > 0 then
                 w_cont  = w_cont + 1
                 If w_cont = 1 Then
                    ShowHTML "<B><FONT COLOR=""#000000"">Atualização da rede particular de ensino</FONT></B>"
                    ShowHTML "<HR>"
                    ShowHTML "Processando..."
                    Response.Flush
                 End If
                 if mid(w_linha,1,1)="[" then
                    w_tipo = replace(replace(w_linha,"[",""),"]","")
                    ShowHTML "<br>Carregando " & w_tipo & "..."
                    w_linha = F2.ReadLine
                    Response.Flush
                 ElseIf w_tipo = "REDE" Then
                    For i = 1 to 50
                       If mid(w_linha,len(w_linha)) <> """" Then
                          w_linha = w_linha & replace(trim(F2.ReadLine),"\""","`")
                       Else
                          i = 50
                       End If
                    Next
                       
                    ' Carrega os dados em array
                    For i = 1 to 33
                        w_campo(i) = "'" & replace(trim(Piece(w_linha,delimitador,",",i)),"'","''") & "'"
                    Next
                    
                    ' Trata valores nulos
                    For i = 1 to 33
                        If w_campo(i) = "'NULL'" Then
                           w_campo(i) = "NULL"
                        End If
                    Next
                    
                    idPasta             = w_campo(1)
                    idInstituicao       = w_campo(2)
                    idDiretor           = w_campo(3)
                    idMantenedora       = w_campo(4)
                    idEndereco          = w_campo(5)
                    idTelefone_1        = w_campo(6)
                    idFax               = w_campo(7)
                    idVencimento        = w_campo(8)
                    idObservacao        = w_campo(9)
                    idEscola            = w_campo(10)
                    idCodinep           = w_campo(11)
                    idTelefone_2        = w_campo(12)
                    idEmail_1           = w_campo(13)
                    idEmail_2           = w_campo(14)
                    idLocalizacao       = w_campo(15)
                    idCnpjExecutora     = w_campo(16)
                    idCnpjEscola        = w_campo(17)
                    idSecretario        = w_campo(18)
                    idCep               = w_campo(19)
                    idAutHabSecretario  = w_campo(20)
                    idEinf              = w_campo(21)
                    idEf                = w_campo(22)
                    idEja               = w_campo(23)
                    idEm                = w_campo(24)
                    idEDA               = w_campo(25)
                    idEprof             = w_campo(26)
                    idRegiao            = w_campo(27)
                    idEndMantenedora    = w_campo(28)
                    SiteEscola          = w_campo(29)
                    Situacao            = w_campo(30)
                    Credenc             = w_campo(31)
                    Login               = w_campo(32)
                    Senha               = w_campo(33)
                    
                    If idMantenedora = "NULL"       Then : idMantenedora = "'Sem informação'" : End If
                    If idPasta = "NULL"             Then : idPasta = 0 : End If
                    If idParecereResolucao = "NULL" Then : idParecereResolucao = "'Sem informação'" : End If
                    If idTelefone_1 = "NULL"        Then : idTelefone_1 = "'Sem informação'" : End If
                    If idEndereco = "NULL"          Then : idEndereco = "'Sem informação'" : End If
                    If idCep = "NULL"               Then : idCep = "'Sem informação'" : End If
                    If idAutHabSecretario = "NULL"  Then : idAutHabSecretario = "0" : End If
                    If idCodinep = "NULL"           Then : idCodinep = "0" : End If
                    If idRegiao = "'0'"             Then : idRegiao = "1" : End If
                    If idVencimento = "'0000-00-00'" Then : idVencimento = "NULL" : End If
                    'ShowHTML "<br>&nbsp;&nbsp;&nbsp;Linha " & w_cont & ": " & replace(idInstituicao,"'","")
                    Response.Flush
                    
                    
                    SQL = "SELECT count(idescola) as Registros FROM escCliente_Particular WHERE idEscola = " & idEscola 
                    ConectaBD SQL

                    If cDbl(RS("Registros")) = 0 Then
                       SQL = "SELECT (MAX(sq_cliente) + 1) as MaxValue from escCliente WHERE sq_cliente < 99000"
                       ConectaBD SQL
                       sq_cliente = cDbl(RS("MaxValue"))
          
                       SQL = "INSERT INTO escCliente (sq_cliente,ativo,sq_tipo_cliente, ds_cliente, no_municipio, sg_uf, ds_username, ds_senha_acesso, localizacao, publica, sq_regiao_adm) " & VbCrLf & _
                             "(SELECT " & sq_cliente & ", 'Sim', a.sq_tipo_cliente, " & idInstituicao & ", " & VbCrLf & _
                             "        substring(b.no_regiao, charIndex(' ',b.no_regiao)+1,500), 'DF', " & VbCrLf & _
                             "        " & Login & ", " & Senha & ", " & idLocalizacao & ", 'N', " & idRegiao & VbCrLf & _
                             "   FROM escTipo_Cliente a, " & VbCrLf & _
                             "        escRegiao_Administrativa b " & VbCrLf & _
                             "  WHERE a.tipo = 4 " & VbCrLf & _
                             "    and b.sq_regiao_adm = " & idRegiao & VbCrLf & _
                             ")"
                       ExecutaSQL(SQL)
                       

                       SQL = "INSERT INTO escCliente_Particular (sq_cliente, Pasta, Diretor, Mantenedora, Endereco, Telefone_1, Fax, Vencimento, Observacao, idEscola, Codinep, Telefone_2, Email_1, Email_2, Cnpj_Executora, Cnpj_Escola, Secretario, Cep, Aut_Hab_Secretario, Infantil, Fundamental, EJA, Medio, Distancia, Profissional, Regiao, mantenedora_endereco, url, situacao, primeiro_credenc) " & VbCrLf & _
                             "VALUES(" & sq_cliente & ", " & idPasta & ", " & idDiretor & ", " & idMantenedora & ", " & idEndereco & ", " & idTelefone_1 & ", " & idFax & ", " & idVencimento & ", " & idObservacao & ", " & idEscola & ", " & idCodinep & ", " & idTelefone_2 & ", " & idEmail_1 & ", " & idEmail_2 & ", " & idCnpjExecutora & ", " & idCnpjEscola & ", " & idSecretario & ", " & idCep & ", " & idAutHabSecretario & ", " & idEinf & ", " & idEf & ", " & idEja & ", " & idEm & ", " & idEDA & ", " & idEprof & ", " & idRegiao & ", " & idEndMantenedora & ", " & SiteEscola & ", " & situacao & ", " & Credenc & ")"
                       ExecutaSQL(SQL)
                    Else
                       SQL = "SELECT sq_cliente FROM escCliente_Particular WHERE idescola = " & idEscola 
                       ConectaBD SQL
                       sq_cliente = RS("sq_cliente")

                       SQL = "UPDATE escCliente SET " & VbCrLf & _
                             "       ds_cliente      = " & idInstituicao & ", " & VbCrLf & _
                             "       ds_username     = " & Login & ", " & VbCrLf & _
                             "       ds_senha_acesso = " & Senha & ", " & VbCrLf & _
                             "       no_municipio    = (select substring(no_regiao, charIndex(' ',no_regiao)+1,500) from escRegiao_Administrativa where sq_regiao_adm = " & idRegiao & "), " & VbCrLf & _
                             "       localizacao     = " & idLocalizacao & ", " & VbCrLf & _
                             "       sq_regiao_adm   = " & idRegiao & VbCrLf & _
                             "  WHERE sq_cliente = " & sq_cliente
                       ExecutaSQL(SQL)

                       SQL = "UPDATE escCliente_Particular SET " & VbCrLf & _
                             "       Pasta                = " & idPasta & ", " & VbCrLf & _
                             "       Diretor              = " & idDiretor & ", " & VbCrLf & _ 
                             "       Mantenedora          = " & idMantenedora & ", " & VbCrLf & _ 
                             "       Endereco             = " & idEndereco & ", " & VbCrLf & _ 
                             "       Telefone_1           = " & idTelefone_1 & ", " & VbCrLf & _ 
                             "       Fax                  = " & idFax& ", " & VbCrLf & _ 
                             "       Vencimento           = " & idVencimento & ", " & VbCrLf & _ 
                             "       Observacao           = " & idObservacao & ", " & VbCrLf & _ 
                             "       Codinep              = " & idCodinep & ", " & VbCrLf & _ 
                             "       Telefone_2           = " & idTelefone_2 & ", " & VbCrLf & _ 
                             "       Email_1              = " & idEmail_1 & ", " & VbCrLf & _ 
                             "       Email_2              = " & idEmail_2 & ", " & VbCrLf & _ 
                             "       Cnpj_Executora       = " & idCnpjExecutora & ", " & VbCrLf & _ 
                             "       Cnpj_Escola          = " & idCnpjEscola & ", " & VbCrLf & _ 
                             "       Secretario           = " & idSecretario & ", " & VbCrLf & _ 
                             "       Cep                  = " & idCep & ", " & VbCrLf & _ 
                             "       Aut_Hab_Secretario   = " & idAutHabSecretario & ", " & VbCrLf & _ 
                             "       Infantil             = " & idEinf & ", " & VbCrLf & _ 
                             "       Fundamental          = " & idEf & ", " & VbCrLf & _ 
                             "       EJA                  = " & idEja & ", " & VbCrLf & _ 
                             "       Medio                = " & idEm & ", " & VbCrLf & _ 
                             "       Distancia            = " & idEDA & ", " & VbCrLf & _ 
                             "       Profissional         = " & idEprof & ", " & VbCrLf & _ 
                             "       Regiao               = " & idRegiao & ", " & VbCrLf & _ 
                             "       Mantenedora_Endereco = " & idEndMantenedora & ", " & VbCrLf & _ 
                             "       URL                  = " & SiteEscola & ", " & VbCrLf & _ 
                             "       Situacao             = " & Situacao & ", " & VbCrLf & _ 
                             "       primeiro_credenc     = " & Credenc  & "  " & VbCrLf & _ 
                             "  WHERE sq_cliente = " & sq_cliente
                       ExecutaSQL(SQL)
                       
                       'Remove os arquivos vinculados à escola
                       SQL = "DELETE FROM escCliente_Arquivo WHERE SQ_SITE_CLIENTE = " & sq_cliente
                       ExecutaSQL(SQL)

                    End If

                    SQL = "DELETE FROM escEspecialidade_Cliente WHERE SQ_CLIENTE = " & sq_cliente
                    ExecutaSQL(SQL)

                    SQL = "SELECT MAX(sq_codigo_cli) as MaxValue from escEspecialidade_Cliente"
                    ConectaBD SQL
                    chave = cDbl(RS("MaxValue"))
                    
                    SQL = "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " & VbCrLf & _
                          "  from escCliente                       a " & VbCrLf & _
                          "       inner join escCliente_Particular b on (a.sq_cliente = b.sq_cliente), " & VbCrLf & _
                          "       escEspecialidade                 c " & VbCrLf & _
                          " where a.sq_cliente = " & sq_cliente & VbCrLf & _
                          "   and (b.infantil  = c.sq_especialidade or  " & VbCrLf & _
                          "        (b.infantil = 3 and c.sq_especialidade in (1,2)) " & VbCrLf & _
                          "       ) " & VbCrLf & _
                          "UNION " & VbCrLf & _
                          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " & VbCrLf & _
                          "  from escCliente                       a " & VbCrLf & _
                          "       inner join escCliente_Particular b on (a.sq_cliente = b.sq_cliente), " & VbCrLf & _
                          "       escEspecialidade                 c " & VbCrLf & _
                          " where a.sq_cliente = " & sq_cliente & VbCrLf & _
                          "   and ((b.fundamental <> 0 and c.sq_especialidade = 5) or  " & VbCrLf & _
                          "        (b.fundamental in (1,2,3,7,11) and c.sq_especialidade =6) " & VbCrLf & _
                          "       ) " & VbCrLf & _
                          "UNION " & VbCrLf & _
                          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " & VbCrLf & _
                          "  from escCliente                       a " & VbCrLf & _
                          "       inner join escCliente_Particular b on (a.sq_cliente = b.sq_cliente), " & VbCrLf & _
                          "       escEspecialidade                 c " & VbCrLf & _
                          " where a.sq_cliente       = " & sq_cliente & VbCrLf & _
                          "   and b.medio            = 1  " & VbCrLf & _
                          "   and c.sq_especialidade = 21 " & VbCrLf & _
                          "UNION " & VbCrLf & _
                          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " & VbCrLf & _
                          "  from escCliente                       a " & VbCrLf & _
                          "       inner join escCliente_Particular b on (a.sq_cliente = b.sq_cliente), " & VbCrLf & _
                          "       escEspecialidade                 c " & VbCrLf & _
                          " where a.sq_cliente       = " & sq_cliente & VbCrLf & _
                          "   and b.eja              <> 0  " & VbCrLf & _
                          "   and c.sq_especialidade = 22 " & VbCrLf & _
                          "UNION " & VbCrLf & _
                          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " & VbCrLf & _
                          "  from escCliente                       a " & VbCrLf & _
                          "       inner join escCliente_Particular b on (a.sq_cliente = b.sq_cliente), " & VbCrLf & _
                          "       escEspecialidade                 c " & VbCrLf & _
                          " where a.sq_cliente       = " & sq_cliente & VbCrLf & _
                          "   and b.profissional     = 1  " & VbCrLf & _
                          "   and c.sq_especialidade = 8 " & VbCrLf & _
                          "UNION " & VbCrLf & _
                          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " & VbCrLf & _
                          "  from escCliente                       a " & VbCrLf & _
                          "       inner join escCliente_Particular b on (a.sq_cliente = b.sq_cliente), " & VbCrLf & _
                          "       escEspecialidade                 c " & VbCrLf & _
                          " where a.sq_cliente       = " & sq_cliente & VbCrLf & _
                          "   and b.distancia        = 1  " & VbCrLf & _
                          "   and c.sq_especialidade = 24" & VbCrLf
                    ConectaBD SQL

                    While not RS.EOF
                      chave = chave + 1
                      SQL = "INSERT INTO escEspecialidade_Cliente (sq_codigo_cli, sq_cliente, sq_codigo_espec) " & VbCrLf & _
                            "VALUES (" & chave & ", " & RS("sq_cliente") & ", " & RS("sq_especialidade") & ") "
                      ExecutaSQL(SQL)
                      RS.MoveNext
                    Wend
                ElseIf w_tipo = "CURSOS" Then
                    ' Carrega os dados em array
                    For i = 1 to 2
                        w_campo(i) = "'" & trim(Piece(w_linha,delimitador,",",i)) & "'"
                    Next
                    
                    ' Trata valores nulos
                    For i = 1 to 2
                        If w_campo(i) = "'NULL'" Then
                           w_campo(i) = "NULL"
                        End If
                    Next
                    
                    Curso    = w_campo(1)
                    idCursos = w_campo(2)
                    'ShowHTML "<br>&nbsp;&nbsp;&nbsp;Linha " & w_cont & ": " & replace(Curso,"'","")
                    Response.Flush

                    SQL = "INSERT INTO escCurso (sq_curso, ds_curso, idcurso) VALUES (" & idCursos & ", " &  Curso & ", " & idCursos &  ");"
                    ExecutaSQL(SQL)
                ElseIf w_tipo = "ARQUIVOS" Then
                    ' Carrega os dados em array
                    For i = 1 to 4
                        w_campo(i) = "'" & trim(Piece(w_linha,delimitador,",",i)) & "'"
                    Next
                    
                    ' Trata valores nulos
                    For i = 1 to 4
                        If w_campo(i) = "'NULL'" Then
                           w_campo(i) = "NULL"
                        End If
                    Next
                    
                    idEscola        = w_campo(1)
                    idNomeArquivo   = w_campo(2)
                    idarquivos      = w_campo(3)
                    descricao       = w_campo(4)
                    'ShowHTML "<br>&nbsp;&nbsp;&nbsp;Linha " & w_cont & ": id " & replace(idArquivos,"'","")
                    Response.Flush

                    SQL = "SELECT MAX(sq_arquivo)+1 as MaxValue from escCliente_Arquivo"
                    ConectaBD SQL
                    chave = cDbl(RS("MaxValue"))

                    SQL =  "INSERT INTO escCLIENTE_ARQUIVO (SQ_ARQUIVO,SQ_SITE_CLIENTE,DT_ARQUIVO,DS_TITULO,DS_ARQUIVO,LN_ARQUIVO,IN_ATIVO,IN_DESTINATARIO,NR_ORDEM) " & VbCrLf & _
                           "(SELECT " & chave & "," & VbCrLf & _
                           "        a.sq_cliente," & VbCrLf & _
                           "        getDate()," & VbCrLf & _
                           "        " & mid(descricao,1,100) & "," & VbCrLf & _
                           "        " & mid(descricao,1,200) & "," & VbCrLf & _
                           "        " & idNomeArquivo & "," & VbCrLf & _
                           "        'Sim'," & VbCrLf & _
                           "        'A'," & VbCrLf & _
                           "        1" & VbCrLf & _
                           "   FROM escCliente_Particular a " & VbCrLf & _
                           "  WHERE a.idEscola = " & idEscola & VbCrLf & _
                           ")"
                    ExecutaSQL(SQL)
                ElseIf w_tipo = "PROFISSIONAIS" Then
                
                    ' Carrega os dados em array
                    For i = 1 to 7
                        w_campo(i) = "'" & trim(Piece(w_linha,delimitador,",",i)) & "'"
                    Next
                    
                    ' Trata valores nulos
                    For i = 1 to 7
                        If w_campo(i) = "'NULL'" Then
                           w_campo(i) = "NULL"
                        End If
                    Next
                    '"id","esc_id","curso_id","tipo","eixo","nic","minCarga"
                    idProfissionais = w_campo(1)
                    idEscola        = w_campo(2)
                    Pasta           = "0" 'w_campo(3)
                    Parecer         = "'Não Inf'" 'w_campo(4)
                    Portaria        = "'Não Inf'" 'w_campo(5)
                    Observacao      = "'Não Inf'" 'w_campo(6)
                    idCursos        = w_campo(3)
                    ShowHTML "<br>&nbsp;&nbsp;&nbsp;Linha " & w_cont & ": id " & replace(idProfissionais,"'","")
                    Response.Flush

                    SQL =  "INSERT INTO escParticular_Curso (sq_particular_curso, sq_cliente, sq_curso, pasta, parecer, portaria, observacao, idProfissional) " & VbCrLf & _ 
                           "(SELECT " & idProfissionais & ", a.sq_cliente, b.sq_curso, " & Pasta & ", " & Parecer & ", " & Portaria & ", " & Observacao & ", " & idProfissionais & VbCrLf & _
                           "   FROM escCliente_Particular a, " & VbCrLf & _
                           "        escCurso b " & VbCrLf & _
                           "  WHERE a.idEscola = " & idEscola & VbCrLf & _
                           "    and b.idCurso  = " & idCursos & VbCrLf & _
                           ")"
                    ExecutaSQL(SQL)
                ElseIf w_tipo = "OS" Then
                    For i = 1 to 50
                       If mid(w_linha,len(w_linha)) <> """" Then
                          w_linha = w_linha & replace(trim(F2.ReadLine),"\""","`")
                       Else
                          i = 50
                       End If
                    Next
                       
                    ' Carrega os dados em array
                    For i = 1 to 8
                        w_campo(i) = "'" & trim(Piece(w_linha,delimitador,",",i)) & "'"
                    Next
                    
                    ' Trata valores nulos
                    For i = 1 to 8
                        If w_campo(i) = "'NULL'" Then
                           w_campo(i) = "NULL"
                        End If
                    Next
                    
                    idOS            = w_campo(1)
                    idEscola        = w_campo(2)
                    Numero          = w_campo(3)
                    Data            = replace(w_campo(4),"-00-00","-01-01")
                    Dodf            = w_campo(5)
                    PagDodf         = w_campo(6)
                    Observacao      = w_campo(7)
                    DataDodf        = replace(w_campo(8),"-00-00","-01-01")
                    If Data = "NULL"     or Data = "'0000-01-01'"      Then : Data = "NULL"     : End If
                    If DataDodf = "NULL" or DataDodf = "'0000-01-01'"  Then : DataDodf = "NULL" : End If
                    ShowHTML "<br>&nbsp;&nbsp;&nbsp;Linha " & w_cont & ": id " & replace(idOS,"'","")
                    Response.Flush

                    SQL =  "INSERT INTO escParticular_OS (sq_particular_os, sq_cliente, numero, data, dodf, dodf_pagina, dodf_data, observacao) " & VbCrLf & _ 
                           "(SELECT " & idOS & ", a.sq_cliente, " & Numero & ", " & Data & ", " & Dodf & ", " & PagDodf & ", " & DataDodf & ", " & Observacao & VbCrLf & _
                           "   FROM escCliente_Particular a " & VbCrLf & _
                           "  WHERE a.idEscola = " & idEscola & VbCrLf & _
                           ")"
                    ExecutaSQL(SQL)
                ElseIf w_tipo = "PORTARIAS" Then
                    For i = 1 to 50
                       If mid(w_linha,len(w_linha)) <> """" Then
                          w_linha = w_linha & " " & replace(trim(F2.ReadLine),"\""","`")
                       Else
                          i = 50
                       End If
                    Next
                       
                    ' Carrega os dados em array
                    For i = 1 to 8
                        w_campo(i) = "'" & trim(Piece(w_linha,delimitador,",",i)) & "'"
                    Next
                    
                    ' Trata valores nulos
                    For i = 1 to 8
                        If w_campo(i) = "'NULL'" Then
                           w_campo(i) = "NULL"
                        End If
                    Next
                    
                    idPortaria      = w_campo(1)
                    idEscola        = w_campo(2)
                    Numero          = w_campo(3)
                    Data            = replace(w_campo(4),"-00-00","-01-01")
                    Dodf            = w_campo(5)
                    PagDodf         = w_campo(6)
                    Observacao      = w_campo(7)
                    DataDodf        = replace(w_campo(8),"-00-00","-01-01")
                    If Data = "NULL"     or Data = "'0000-01-01'"      Then : Data = "NULL"     : End If
                    If DataDodf = "NULL" or DataDodf = "'0000-01-01'"  Then : DataDodf = "NULL" : End If
                    ShowHTML "<br>&nbsp;&nbsp;&nbsp;Linha " & w_cont & ": id " & replace(idPortaria,"'","")
                    Response.Flush

                    SQL =  "INSERT INTO escParticular_Portaria (sq_particular_portaria, sq_cliente, numero, data, dodf, dodf_pagina, dodf_data, observacao) " & VbCrLf & _ 
                           "(SELECT " & idPortaria & ", a.sq_cliente, " & Numero & ", " & Data & ", " & Dodf & ", " & PagDodf & ", " & DataDodf & ", " & Observacao & VbCrLf & _
                           "   FROM escCliente_Particular a " & VbCrLf & _
                           "  WHERE a.idEscola = " & idEscola & VbCrLf & _
                           ")"
                    ExecutaSQL(SQL)
                 End If
             end if
          Loop
          F2.Close

          'Encerra a transação
          dbms.CommitTrans()
          ScriptOpen "JavaScript"
          'ShowHTML "  confirm('Atualização da rede particular executada com sucesso.');"
          ShowHTML "  alert('Atualização da rede particular executada com sucesso.');"
          ShowHTML "  location.href='" & w_pagina & ul.Form("R") & "&w_ea=L';"
          ScriptClose
          Rodape
          Response.End()
          exit sub
       End If

       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & ul.Form("R") & "&w_ea=L';"
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

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         1, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - inclusão de mensagem para aluno da rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "A" Then
          SQL = " update escMensagem_Aluno set " & VbCrLf & _
                "     dt_mensagem  = convert(datetime, '" & FormataDataEdicao(FormatDateTime(Request("w_dt_mensagem"),2)) & "',103), " & VbCrLf & _
                "     ds_mensagem  = '" & Request("w_ds_mensagem") & "' " & VbCrLf & _
                "where sq_mensagem = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         2, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - alteração de mensagem para aluno da rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "E" Then
          SQL = " delete escMensagem_Aluno where sq_mensagem = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         3, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - exclusão de mensagem para aluno da rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       End If
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
  
    Case conWhatEspecialidade
       dbms.BeginTrans()
       
       If w_ea = "I" Then
          ' Recupera o valor da próxima chave primária
          SQL = "select IsNull(max(sq_especialidade),0)+1 chave from escEspecialidade" & VbCrLf
          ConectaBD SQL
          w_chave = RS("chave")
          DesconectaBD
          
          ' Insere o arquivo
          SQL = " insert into escEspecialidade " & VbCrLf & _
                "    (sq_especialidade, ds_especialidade, nr_ordem) " & VbCrLf & _
                " values ( " & w_chave & ", " & VbCrLf & _
                "     '" & Request("w_ds_especialidade") & "', " & VbCrLf & _
                "     '" & Request("w_nr_ordem") & "' " & VbCrLf & _                
                " )" & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         1, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - inclusão de modalidade de ensino na rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "A" Then
          SQL = " update escEspecialidade set " & VbCrLf & _
                "     ds_especialidade  = '" & Request("w_ds_especialidade") & "', " & VbCrLf & _
                "             nr_ordem  = '" & Request("w_nr_ordem") & "' " & VbCrLf & _
                "where sq_especialidade = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         2, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - alteração de modalidade de ensino na rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       ElseIf w_ea = "E" Then
          SQL = " delete escEspecialidade_Cliente where sq_codigo_espec = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         3, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - exclusão de modalidade de ensino na rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If

          SQL = " delete escEspecialidade where sq_especialidade = " & Request("w_chave") & VbCrLf
          ExecutaSQL(SQL)

          If Session("username") <> "SBPI" Then
             ' Grava o acesso na tabela de log
             w_sql = SQL
             SQL = "insert into escCliente_Log (sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " & VbCrLf & _
                   "values ( " & VbCrLf & _
                   "         " & Request("w_sq_cliente") & ", " & VbCrLf & _
                   "         getdate(), " & VbCrLf & _
                   "         '" & Request.ServerVariables("REMOTE_HOST") & "', " & VbCrLf & _
                   "         3, " & VbCrLf & _
                   "         'Usuário """ & uCase(Session("username")) & """ - exclusão de modalidade de ensino na rede.', " & VbCrLf & _
                   "         '" & replace(w_sql,"'", "''") & "', " & VbCrLf & _
                   "         " & w_funcionalidade & " " & VbCrLf & _
                   "       ) " & VbCrLf
             ExecutaSQL(SQL)
          End If
       End If
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
       
    Case "CADASTROESCOLA"
       If Nvl(w_ea,"") = "I" Then
          SQL = "SELECT ds_username from escCliente where ds_username = '" & Request("w_ds_username") & "'" & VbCrLf
          ConectaBD SQL
          If RS.RecordCount > 0 Then
             ScriptOpen "JavaScript"
             ShowHTML "alert('Username já existente!');"
             ShowHTML "history.back(1);"
             ScriptClose
          End If
          DesconectaBD
          SQL = "SELECT max(sq_cliente) sq_cliente from EscCliente where sq_cliente < 99000 " & VbCrLf
          ConectaBD SQL
          w_chave = cDbl(RS("sq_cliente")) + 1
          DesconectaBD    
       ElseIf Nvl(w_ea,"") = "A" Then
          If Request("w_ds_username") <> Request("w_username_atual") Then
             SQL = "SELECT ds_username from escCliente where ds_username = '" & Request("w_ds_username") & "'" & VbCrLf
             ConectaBD SQL
             If RS.RecordCount > 0 Then
                ScriptOpen "JavaScript"
                ShowHTML "alert('Username já existente!');"
                ShowHTML "history.back(1);"
                ScriptClose
             End If
             DesconectaBD
          End If
          w_chave = Request("w_chave")
          
       End If
       strDirA = "sedf\" & Nvl(Request("w_username_atual"),Request("w_ds_username"))
       strDirN = "sedf\" & Request("w_ds_username")
       strArq = "default.asp"
       strMod = Request("w_sq_modelo")
       sqCliente = w_chave
       strFile = replace(conFilePhysical & "\" & strDirN & "\" & strArq,"\\","\")
       strDirA = replace(conFilePhysical & "\" & strDirA,"\\","\")
       strDirN = replace(conFilePhysical & "\" & strDirN,"\\","\")
       strMod = replace(conFilePhysical & "\" & "Modelos\Mod" & strMod & "\Default_cliente.asp","\\","\")
       Set FS = CreateObject("Scripting.FileSystemObject")

       If (FS.FolderExists (strDirA)) then
          If Nvl(Request("w_username_atual"),Request("w_ds_username")) <> Request("w_ds_username") Then
             FS.MoveFolder strDirA,strDirN
          End If
       Else
         Set f1 = FS.CreateFolder(strDirN)       
       End If

          
       Set f1 = FS.CreateTextFile(strFile)
          
       Set f2 = FS.OpenTextFile(strMod, ForReading)
          
       Do While Not f2.AtEndOfStream 
          strLine = f2.ReadLine  
          strLine = replace(strLine,"= *%*","= " & sqCliente)
          f1.WriteLine strLine
       Loop
              
       f2.Close
       f1.Close

       dbms.BeginTrans()
       If Nvl(w_ea,"") = "I" Then
          
          'Criação das escolas na tabela principal
          SQL = "insert into escCliente (sq_cliente,              sq_cliente_pai,              sq_tipo_cliente,                            ds_cliente,  " & VbCrLf &_
                "                        ds_apelido,              no_municipio,                sg_uf,                                      ds_username, " & VbCrLf &_ 
                "                        ds_senha_acesso,         ln_internet,                 ds_email,                                   dt_cadastro,  " & VbCrLf &_ 
                "                        ativo,                   sq_regiao_adm,               localizacao  " & VbCrLf &_ 
                " )values( " & VbCrLf &_
                "                        " & w_chave & "," & Request("w_sq_cliente_pai") & "," &  Request("w_sq_tipo_cliente") & ",'" & Request("w_ds_cliente") & "'," & VbCrLf &_ 
                "                        '" & Request("w_ds_apelido") & "','" & Request("w_no_municipio")& "','" & Request("w_sg_uf") & "','" & Request("w_ds_username") & "'," & VbCrLf &_ 
                "                        '9"&w_chave & "','" & Request("w_ln_internet") & "','" & Nvl(Request("w_ds_email"),"Não informado") & "', getDate(),'" & Request("w_ativo") & "', " & VbCrLf &_ 
                "                        '" & Request("w_regiao") & "','" & Request("w_local")& "'); " & VbCrLf
          
          ExecutaSQL(SQL)
                
          'Criação das escolas na tabela que contém dados complementares'
          SQL = "insert into escCliente_Dados (sq_cliente_dados,            sq_cliente,          nr_cnpj,                 tp_registro,  " & VbCrLf &_ 
                "                              ds_ato,                      nr_ato,              dt_ato,                  ds_orgao,         ds_logradouro, " & VbCrLf &_ 
                "                              no_bairro,                   nr_cep,              no_contato,              ds_email_contato, nr_fone_contato, " & VbCrLf &_ 
                "                              nr_fax_contato,              no_diretor,          no_secretario " & VbCrLf &_ 
                " ) values ( " & VbCrLf &_   
                "                              " & w_chave & "," & w_chave & "," & w_chave & ",null, " & VbCrLf &_        
                "                              null,                        null,                null,                    null,'" & Request("w_ds_logradouro") & "'," & VbCrLf &_ 
                "                              '" & Request("w_no_bairro") & "','" & trim(Request("w_nr_cep")) & "','" & Request("w_no_contato") & "', '" & Request("w_email_contato") & "','" & Request("w_nr_fone_contato") & "'," &  VbCrLf &_ 
                "                              '" & Request("w_nr_fax_contato") & "','" & Nvl(Request("w_no_diretor"),"Não informado") & "','" & Nvl(Request("w_no_secretario"),"Não informado") & "'); " & VbCrLf
          
          ExecutaSQL(SQL)
          
          'Criação das escolas na tabela que contém informações sobre o site web
          SQL = "insert into escCliente_Site (sq_site_cliente,               sq_cliente,                   sq_modelo,                      no_contato_internet,              " & VbCrLf &_ 
                "                             nr_fone_internet,              nr_fax_internet,              ds_email_internet,              ds_diretorio,                     " & VbCrLf &_
                "                             sq_siscol,                     ds_mensagem,                  ds_institucional,               ds_texto_abertura                 " & VbCrLf &_ 
                " ) values ( " & VbCrLf &_
                "                             " & w_chave & "," & w_chave & "," & Request("w_sq_modelo") & ",'" & Request("w_no_contato_internet") & "', " & VbCrLf &_  
                "                             '" & Request("w_nr_fone_internet") & "','" & Request("w_nr_fax_internet") & "','" & Request("w_ds_email_internet") & "','" & Request("w_ds_diretorio") & "'," &  VbCrLf &_  
                "                             '" & Request("w_sq_siscol") & "','" & Nvl(Request("w_ds_mensagem"),"A ser inserido") & "','" & Nvl(Request("w_ds_institucional"),"A ser inserido") & "','" & Nvl(Request("w_ds_texto_abertura"),"A ser inserido") & "');" & VbCrLf
          
          ExecutaSQL(SQL)
          
          'Criação das escolas na tabela que contém as modalidades de ensino oferecidas pela escola
          For i = 1 to Request("w_sq_codigo_espec").Count 
             SQL = "select max(sq_codigo_cli) sq_codigo_cli from escespecialidade_cliente " & VbCrLf
             ConectaBD SQL
             w_sq_codigo_cli = cDbl(RS("sq_codigo_cli")) + 1
             DesconectaBD    
             
             SQL = "insert into escEspecialidade_Cliente (sq_codigo_cli, sq_cliente, sq_codigo_espec " & VbCrLf &_
                   " ) values ( " & VbCrLf &_
                   "                                 " & w_sq_codigo_cli & "," & w_chave & "," & Request("w_sq_codigo_espec") & ");" & VbCrLf
             ExecutaSQL(SQL)
          Next
       ElseIf Nvl(w_ea,"") = "A" Then
       
          'Altera os dados das escolas na tabela principal
          SQL = "update escCliente set " & VbCrLf & VbCrLf &_
                "   sq_cliente_pai  = " & Request("w_sq_cliente_pai") & "," & VbCrLf &_
                "   sq_regiao_adm   = " & Request("w_regiao") & "," & VbCrLf &_
                "   localizacao     = '" & Request("w_local") & "'," & VbCrLf &_
                "   sq_tipo_cliente = " & Request("w_sq_tipo_cliente") & "," & VbCrLf &_
                "   ds_cliente      = '" & Request("w_ds_cliente") & "'," & VbCrLf &_
                "   ds_apelido      = '" & Request("w_ds_apelido") & "'," & VbCrLf &_
                "   no_municipio    = '" & Request("w_no_municipio") & "'," & VbCrLf &_
                "   sg_uf           = '" & Request("w_sg_uf") & "'," & VbCrLf &_
                "   ds_username     = '" & Request("w_ds_username") & "'," & VbCrLf &_
                "   ln_internet     = '" & Request("w_ln_internet") & "'," & VbCrLf &_
                "   ds_email        = '" & Request("w_ds_email") & "'," & VbCrLf &_
                "   ativo        = '" & Request("w_ativo") & "'" & VbCrLf &_
                " where sq_cliente = " & w_chave & VbCrLf
          ExecutaSQL(SQL)
          'Altera as escolas na tabela que contém dados complementares
          SQL = "update escCliente_Dados set " & VbCrLf & VbCrLf &_
                "   ds_logradouro   = '" & Request("w_ds_logradouro") & "'," & VbCrLf &_
                "   no_bairro       = '" & Request("w_no_bairro") & "'," & VbCrLf &_
                "   nr_cep          = '" & trim(Request("w_nr_cep")) & "'," & VbCrLf &_
                "   no_contato      = '" & Request("w_no_contato") & "'," & VbCrLf &_
                "   ds_email_contato= '" & Request("w_email_contato") & "'," & VbCrLf &_
                "   nr_fone_contato = '" & Request("w_nr_fone_contato") & "'," & VbCrLf &_
                "   nr_fax_contato  = '" & Request("w_nr_fax_contato") & "'," & VbCrLf &_
                "   no_diretor      = '" & Nvl(Request("w_no_diretor"),"Não informado") & "'," & VbCrLf &_
                "   no_secretario   = '" & Nvl(Request("w_no_secretario"),"Não informado")& "'" & VbCrLf &_
                " where sq_cliente = " & w_chave & VbCrLf
          ExecutaSQL(SQL)
          
          'Altera as escolas na tabela que contém informações sobre o site web
          SQL = "update escCliente_Site set " & VbCrLf & VbCrLf &_
                "   sq_modelo             = '" & Request("w_sq_modelo") & "'," & VbCrLf &_
                "   no_contato_internet   = '" & Request("w_no_contato_internet") & "'," & VbCrLf &_
                "   nr_fone_internet      = '" & Request("w_nr_fone_internet") & "'," & VbCrLf &_
                "   nr_fax_internet       = '" & Request("w_nr_fax_internet") & "'," & VbCrLf &_
                "   ds_email_internet     = '" & Request("w_ds_email_internet") & "'," & VbCrLf &_
                "   ds_diretorio          = '" & Request("w_ds_diretorio") & "'," & VbCrLf &_
                "   sq_siscol             = '" & Request("w_sq_siscol") & "'," & VbCrLf &_
                "   ds_mensagem           = '" & Nvl(Request("w_ds_mensagem"),"A ser inserido") & "'," & VbCrLf &_
                "   ds_institucional      = '" & Nvl(Request("w_ds_institucional"),"A ser inserido") & "'," & VbCrLf &_
                "   ds_texto_abertura     = '" & Nvl(Request("w_ds_texto_abertura"),"A ser inserido") & "'" & VbCrLf &_
                " where sq_cliente = " & w_chave & VbCrLf   
          ExecutaSQL(SQL)
          
          'Apaga as escolas na tabela que contém as modalidades de ensino oferecidas pela escola
          SQL = "Delete escEspecialidade_Cliente where sq_cliente = " & w_chave
          ExecutaSQL(SQL)
          
          'Criação das escolas na tabela que contém as modalidades de ensino oferecidas pela escola
          For i = 1 to Request("w_sq_codigo_espec").Count 
             SQL = "select max(sq_codigo_cli) sq_codigo_cli from escespecialidade_cliente " & VbCrLf
             ConectaBD SQL
             w_sq_codigo_cli = cDbl(RS("sq_codigo_cli")) + 1
             DesconectaBD    
             SQL = "insert into escEspecialidade_Cliente (sq_codigo_cli, sq_cliente, sq_codigo_espec " & VbCrLf &_
                   " ) values ( " & VbCrLf &_
                   "                                 " & w_sq_codigo_cli & "," & w_chave & "," & Request("w_sq_codigo_espec")(i) & ");" & VbCrLf
                     
             ExecutaSQL(SQL)
          Next
       End If
       
       dbms.CommitTrans()
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_Pagina & conRelEscolas & "&pesquisa=Sim" & MontaFiltro("GETX") & "&w_ea=L';"
       ScriptClose
    
    Case "VERIFARQ"
       ' Em seguida, cria apenas as especialidades indicadas pelo cliente
       Set FS = CreateObject("Scripting.FileSystemObject")
       For w_cont = 1 To Request.Form("arquivo").Count
          
            strDir  = mid(Request.Form("arquivo")(w_cont),1,instr(Request.Form("arquivo")(w_cont),"=|=")-1)
            strFile = mid(Request.Form("arquivo")(w_cont),instr(Request.Form("arquivo")(w_cont),"=|=")+3)
            strFile = replace(conFilePhysical & "\sedf\" & strDir & "\" & strFile,"\\","\")
            ' Remove o arquivo físico
            DeleteAFile strFile
       Next
       Set FS = Nothing
          
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & Request("R") & "&w_ea=L&cl=" & cl & "';"
       ScriptClose
       
    Case "VERIFBANCO"
       dbms.BeginTrans()
       For w_cont = 1 To Request.Form("arquivo").Count
        strDir  = mid(Request.Form("arquivo")(w_cont),1,instr(Request.Form("arquivo")(w_cont),"=|=")-1) 'tipo
        strFile = mid(Request.Form("arquivo")(w_cont),instr(Request.Form("arquivo")(w_cont),"=|=")+3) ' chave a ser removida
        
        If strDir = "FOTO" Then
          SQL = "DELETE escCliente_Foto WHERE SQ_CLIENTE_FOTO = " & strFile
        Else
          SQL = "DELETE escCliente_Arquivo WHERE SQ_Arquivo = " & strFile
        End If        
        ExecutaSQL(SQL)
        
       Next
       dbms.CommitTrans()
        
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & "VERIFARQ&w_ea=L&cl=" & cl & "';"
       ScriptClose       
       
    Case "ESCPARTHOMOLOG"
       dbms.BeginTrans()
       For w_cont = 1 To Request.Form("w_chave").Count
          SQL = "UPDATE escParticular_Calendario SET homologado = '" & Request.Form("w_homologado")(w_cont) & "' , ultima_homologacao = getdate() WHERE sq_particular_calendario = " & Request.Form("w_chave")(w_cont)
          ExecutaSQL(SQL)       
       Next
       dbms.CommitTrans()
        
       ScriptOpen "JavaScript"
       ShowHTML "  location.href='" & w_pagina & "ESCPARTHOMOLOG&w_ea=L&cl=" & cl & "';"
       ScriptClose         

    Case Else
    'Response.Write w_R
    'Response.End 
       ScriptOpen "JavaScript"
       ShowHTML "  alert('Bloco de dados não encontrado: " & SG & "');"
       ShowHTML "  history.back(1);"
       ScriptClose
  End Select
  
  Set w_chave          = Nothing
  Set w_funcionalidade = Nothing
End Sub

REM =========================================================================
REM Monta a tela de Pesquisa
REM -------------------------------------------------------------------------
Public Sub ShowEscolas

  Dim RS1, p_regional

  Dim sql, sql2, wCont, sql1, wAtual, wIN, w_especialidade
  
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  
  p_regional = Request("p_regional")

  If p_tipo = "W" Then
      Response.ContentType = "application/msword"
      HeaderWord p_layout
      ShowHTML "<TABLE WIDTH=""100%"" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR=""#000000"">SIGE-WEB<TD ALIGN=""RIGHT""><B><FONT SIZE=4 COLOR=""#000000"">"
      ShowHTML "Consulta a escolas"
      ShowHTML "</FONT><TR><TD ALIGN=""RIGHT""><B><FONT SIZE=2 COLOR=""#000000"">" & DataHora() & "</B></TD></TR>"
      ShowHTML "</FONT></B></TD></TR></TABLE>"
      ShowHTML "<HR>"
  Else
     Cabecalho
     ShowHTML "<HEAD>"
     ShowHTML "</HEAD>"
     If Request("pesquisa") > "" Then
        BodyOpen " onLoad=""location.href='#lista'"""
     Else
        BodyOpen "onLoad='document.Form.p_regional.focus()';"
     End If
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
  ShowHTML "    <table width=""95%"" border=""0"">"
  If p_tipo = "H" Then
     Showhtml "<FORM ACTION=""Controle.asp"" name=""Form"" METHOD=""POST"">"
     ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""w_ew"" VALUE=""" & conRelEscolas &  """>"
     ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""CL"" VALUE=""" & CL &  """>"
     ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""pesquisa"" VALUE=""X"">"
     ShowHTML "<input type=""Hidden"" name=""P3"" value=""1"">"
     ShowHTML "<input type=""Hidden"" name=""P4"" value=""15"">"
     ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
     ShowHTML "    <table width=""100%"" border=""0"">"
     ShowHTML "          <TR><TD valign=""top""><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
  Else
     ShowHTML "<tr><td><div align=""justify""><font size=1><b>Filtro:</b><ul>"
  End If
  If p_tipo = "H" Then
     ShowHTML "          <tr valign=""top""><td>"
     SelecaoRegional "<u>S</u>ubordinação:", "S", "Indique a subordinação da escola.", p_regional, null, "p_regional", null, null
  ElseIf Nvl(p_regional,0) > 0 Then
     SQL = "SELECT  a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " & VbCrLf & _
           "  FROM  escCLIENTE a " & VbCrLf & _
           " WHERE  a.sq_cliente = " & p_regional & VbCrLf & _
           "ORDER BY a.ds_cliente "
     ConectaBD SQL
     Response.Write "          <li><font size=""1""><b>Escolas da " & RS("ds_cliente") & "</b>"
     DesconectaBD
  End If
  SQL = "SELECT * FROM escTipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
  ConectaBD SQL
  If p_tipo = "H" Then
     ShowHTML "          <tr valign=""top""><td><td><font size=""1""><br><b>Tipo de instituição:</b><br><SELECT CLASS=""STI"" NAME=""p_tipo_cliente"">"
     If RS.RecordCount > 1 Then ShowHTML "          <option value="""">Todos" End If
     While Not RS.EOF
        If cDbl(nvl(RS("sq_tipo_cliente"),0)) = cDbl(nvl(Request("p_tipo_cliente"),0)) Then
           ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """ SELECTED>" & RS("ds_tipo_cliente")
        Else
           ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """>" & RS("ds_tipo_cliente")
        End If
        RS.MoveNext
     Wend
     ShowHTML "          </select>"
  ElseIf nvl(Request("p_tipo_cliente"),0) > 0 Then
     ShowHTML "          <li><font size=""1""><b>Tipo de instituição: "
     While Not RS.EOF
        If cDbl(nvl(RS("sq_tipo_cliente"),0)) = cDbl(nvl(Request("p_tipo_cliente"),0)) Then ShowHTML RS("ds_tipo_cliente") End If
        RS.MoveNext
     Wend
     ShowHTML "</b>"
  End If
  If p_tipo = "H" Then
     ShowHTML "  <TR><TD><TD><font size=""1""><br>"
     If Request("C") > "" Then
        ShowHTML "          <input type=""checkbox"" name=""C"" value=""X"" CLASS=""BTM"" checked> Exibir apenas escolas com alunos carregados "
     Else
        ShowHTML "          <input type=""checkbox"" name=""C"" value=""X"" CLASS=""BTM""> Exibir apenas escolas com alunos carregados "
     End If
  ElseIf Request("C") > "" Then
     ShowHTML "  <li><font size=""1""><b>Apenas escolas com alunos carregados </b>"
  End If
  If p_tipo = "H" Then
     ShowHTML "  <TR><TD><TD><font size=""1"">"
     If Request("D") > "" Then
        ShowHTML "          <input type=""checkbox"" name=""D"" value=""X"" CLASS=""BTM"" checked> Exibir apenas escolas que tenham alterado seus dados "
     Else
        ShowHTML "          <input type=""checkbox"" name=""D"" value=""X"" CLASS=""BTM""> Exibir apenas escolas que tenham alterado seus dados "
     End If
  ElseIf Request("D") > "" Then
     ShowHTML "  <li><font size=""1""><b>Apenas escolas que tenham alterado seus dados </b>"
  End If
  If p_tipo = "H" Then
     ShowHTML "  <TR><TD><TD><font size=""1""><br><b>Quanto às informações administrativas?</b><br>"
     If Request("E") = "S" Then ShowHTML "          <input type=""radio"" name=""E"" value=""S"" CLASS=""BTM"" checked> Listar apenas as escolas que informaram<br>"     Else ShowHTML "          <input type=""radio"" name=""E"" value=""S"" CLASS=""BTM""> Listar apenas as escolas que informaram<br>"      End If
     If Request("E") = "N" Then ShowHTML "          <input type=""radio"" name=""E"" value=""N"" CLASS=""BTM"" checked> Listar apenas as escolas que não informaram<br>" Else ShowHTML "          <input type=""radio"" name=""E"" value=""N"" CLASS=""BTM""> Listar apenas as escolas que não informaram<br> " End If
     If Request("E") = ""  Then ShowHTML "          <input type=""radio"" name=""E"" value="""" CLASS=""BTM"" checked> Tanto faz"                                        Else ShowHTML "          <input type=""radio"" name=""E"" value="""" CLASS=""BTM""> Tanto faz "                                        End If
  ElseIf Request("E") > "" Then
     ShowHTML "  <li><font size=""1""><b>"
     If Request("E") = "S" Then ShowHTML "          Listar apenas as escolas que informaram dados administrativos</b>"      End If
     If Request("E") = "N" Then ShowHTML "          Listar apenas as escolas que não informaram dados administrativos</b> " End If
  End If
  If p_tipo <> "W" Then
     ShowHTML "          </table>"
  End If
  DesconectaBD
  wCont = 0
  sql1 = ""

  ' Seleção de etapas/modalidades
  sql = "SELECT DISTINCT a.curso as sq_especialidade, a.curso as ds_especialidade, 1 as nr_ordem, 'M' as tp_especialidade " & VbCrLf & _ 
        " from escTurma_Modalidade                 AS a " & VbCrLf & _ 
        "      INNER JOIN escTurma                 AS c ON (a.serie           = c.ds_serie) " & VbCrLf & _
        "      INNER JOIN escCliente               AS d ON (c.sq_site_cliente = d.sq_cliente) " & VbCrLf & _
        "UNION " & VbCrLf & _ 
        "SELECT DISTINCT cast(a.sq_especialidade as varchar) as sq_especialidade, a.ds_especialidade,  " & VbCrLf & _ 
        "       case a.tp_especialidade when 'J' then '1' else a.nr_ordem end as nr_ordem, " & VbCrLf & _ 
        "       case a.tp_especialidade when 'J' then 'M' else a.tp_especialidade end as tp_especialidade" & VbCrLf & _ 
        " from escEspecialidade AS a " & VbCrLf & _ 
        "      INNER JOIN escEspecialidade_cliente AS c ON (a.sq_especialidade = c.sq_codigo_espec) " & VbCrLf & _
        "      INNER JOIN escCliente               AS d ON (c.sq_cliente       = d.sq_cliente) " & VbCrLf & _
        " where a.tp_especialidade <> 'M' " & VbCrLf & _
        "ORDER BY a.nr_ordem, a.ds_especialidade " & VbCrLf
  ConectaBD sql
   
  If p_tipo = "H" Then
     If Not RS.EOF Then
        wCont = 0
        wAtual = ""
 
        ShowHTML "          <TD valign=""top""><table border=""0"" align=""left"" cellpadding=0 cellspacing=0>"
        Do While Not RS.EOF
           If wAtual = "" or wAtual <> RS("tp_especialidade") Then
              wAtual = RS("tp_especialidade")
              If wAtual = "M" Then
                 ShowHTML "            <TR><TD colspan=2><font size=""1""><b>Etapas/Modalidades de ensino:</b>"
              ElseIf wAtual = "R" Then
                 ShowHTML "            <TR><TD colspan=2><font size=""1""><b>Em Regime de Intercomplementaridade:</b>"
              Else
                 ShowHTML "            <TR><TD colspan=2><font size=""1""><b>Outras:</b>"
              End If
           End If
           
           wCont = wCont + 1
           marcado = "N"
           For i = 1 to Request("p_modalidade").Count
               If RS("sq_especialidade") = Request("p_modalidade")(i) Then 
                  marcado = "S" 
                  sql1 = ",'" & Request("p_modalidade")(i) & "'" & sql1
               End If
           Next
              
           If marcado = "S" Then
              ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""p_modalidade"" value=""" & RS("sq_especialidade") & """ checked><td><font size=1>" & RS("ds_especialidade")
              wIN = 1
           Else
              ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""p_modalidade"" value=""" & RS("sq_especialidade") & """><td><font size=1>" & RS("ds_especialidade")
           End If
           RS.MoveNext
 
           If (wCont Mod 2) = 0 Then 
              wCont = 0
           End If
 
        Loop
        DesconectaBD
        sql1 = Mid(sql1,2)
     End If
  ElseIf Nvl(Request("p_modalidade"), "") > "" Then
     If Not RS.EOF Then
        wCont = 0
 
        ShowHTML "          <li><b>Modalidades de ensino:</b><ul>"
        sql1 = "'" & replace(Request("p_modalidade"),", ","','") & "'"
        Do While Not RS.EOF
           For i = 1 to Request("p_modalidade").Count
               If InStr(sql1,"'"&RS("sq_especialidade")&"'")>0 Then 
                  ShowHTML chr(13) & "           <li><font size=1>" & RS("ds_especialidade")
               End If
           Next
           wIN = 1
           RS.MoveNext
        Loop
        DesconectaBD
     End If
  End If
  ShowHTML "          </tr>"
  ShowHTML "          </table>"
  if p_tipo = "H" Then 
     ShowHTML "  <TR><TD colspan=2><font size=""1""><b>Campos a serem exibidos"
     If p_layout = "PORTRAIT" Then 
        ShowHTML "          (<input type=""radio"" name=""p_layout"" value=""PORTRAIT"" CLASS=""BTM"" checked> Retrato <input type=""radio"" name=""p_layout"" value=""LANDSCAPE"" CLASS=""BTM""> Paisagem)"
     Else
        ShowHTML "          (<input type=""radio"" name=""p_layout"" value=""PORTRAIT"" CLASS=""BTM""> Retrato <input type=""radio"" name=""p_layout"" value=""LANDSCAPE"" CLASS=""BTM"" checked> Paisagem)"
     End If
     ShowHTML "     <table width=""100%"" border=0>"
     ShowHTML "       <tr valign=""top"">"
     If Session("username") = "SEDF" or Session("username") = "SBPI" or Session("username") = "CTIS" or Mid(Session("username"),1,2) = "RE" Then
        If Instr(p_campos,"username")    > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""username"" CLASS=""BTM"" checked>Username"           Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""username"" CLASS=""BTM"">Username"          End If
        If Instr(p_campos,"senha")       > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""senha"" CLASS=""BTM"" checked>Senha"                 Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""senha"" CLASS=""BTM"">Senha"                End If
     End If
     If Instr(p_campos,"alteracao")   > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""alteracao"" CLASS=""BTM"" checked>Última alteração"  Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""alteracao"" CLASS=""BTM"">Última alteração" End If
     If Instr(p_campos,"diretor")     > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""diretor"" CLASS=""BTM"" checked>Diretor"             Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""diretor"" CLASS=""BTM"">Diretor"            End If
     If Instr(p_campos,"secretario")  > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""secretario"" CLASS=""BTM"" checked>Secretário"       Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""secretario"" CLASS=""BTM"">Secretário"      End If
     If Instr(p_campos,"contato")     > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""contato"" CLASS=""BTM"" checked>Contato"             Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""contato"" CLASS=""BTM"">Contato"            End If
     ShowHTML "       <tr valign=""top"">"
     If Instr(p_campos,"mail")        > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""mail"" CLASS=""BTM"" checked>e-Mail"                 Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""mail"" CLASS=""BTM"">e-Mail"                End If
     If Instr(p_campos,"telefone")    > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""telefone"" CLASS=""BTM"" checked>Telefone"           Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""telefone"" CLASS=""BTM"">Telefone"          End If
     If Instr(p_campos,"endereco")    > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""endereco"" CLASS=""BTM"" checked>Endereço"           Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""endereco"" CLASS=""BTM"">Endereço"          End If
     If Instr(p_campos,"bairro")      > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""bairro"" CLASS=""BTM"" checked>Bairro"               Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""bairro"" CLASS=""BTM"">Bairro"              End If
     If Instr(p_campos,"localizacao") > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""localizacao"" CLASS=""BTM"" checked>Localização"     Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""localizacao"" CLASS=""BTM"">Localização"    End If
     If Instr(p_campos,"cep")         > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""cep"" CLASS=""BTM"" checked>CEP"                     Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""cep"" CLASS=""BTM"">CEP"                    End If
     ShowHTML "     </table>"
     ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000"">"
     ShowHTML "      <tr><td align=""center"" colspan=""3"">"
     ShowHTML "            <input class=""BTM"" type=""submit"" name=""Botao"" value=""Aplicar filtro"">"
     If Session("username") = "SBPI" Then
        ShowHTML "            <input class=""BTM"" type=""button"" name=""Botao"" onClick=""location.href='" & w_Pagina & "CadastroEscola" & "&CL=" & CL & MontaFiltro("GET") & "&w_ea=I';"" value=""Nova escola"">"
     End If
     ShowHTML "          </td>"
     ShowHTML "      </tr>"
  End If
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  if p_tipo = "H" Then ShowHTML "</form>" End If
  If Request("pesquisa") > "" Then
     sql = "SELECT DISTINCT b.ds_username, d.sq_cliente, d.ds_cliente, d.ds_apelido, d.ln_internet, " & VbCrLf & _
           "       d.ds_username, d.ds_senha_acesso, d.no_municipio, d.sg_uf, d.dt_alteracao, " & VbCrLf & _ 
           "       e.no_diretor, e.no_secretario, e.no_bairro, e.nr_cep, e.ds_logradouro, " & VbCrLf & _ 
           "       f.ds_email_internet, f.no_contato_internet, nr_fone_internet, nr_fax_internet, " & VbCrLf & _ 
           "       IsNull(h.existe,0) alunos, i.sq_cliente adm, d.ativo " & VbCrLf & _ 
           "  from escCliente                                 d " & VbCrLf & _ 
           "       INNER      JOIN escCliente                 b ON (b.sq_cliente       = d.sq_cliente_pai) " & VbCrLf & _
           "       LEFT OUTER JOIN escCliente_Dados           e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
           "       INNER      JOIN escCliente_Site            f ON (d.sq_cliente       = f.sq_cliente) " & VbCrLf & _
           "       INNER      JOIN escTipo_Cliente            g ON (d.sq_tipo_cliente  = g.sq_tipo_cliente) " & VbCrLf & _
           "       LEFT OUTER JOIN (select sq_site_cliente, count(*) existe " & VbCrLf & _
           "                          from escAluno " & VbCrLf & _
           "                         group by sq_site_cliente " & VbCrLf & _
           "                       )                          h on (f.sq_site_cliente  = h.sq_site_cliente) " & VbCrLf & _
           "       LEFT OUTER JOIN escCliente_Admin           i on (d.sq_cliente       = i.sq_cliente) " & VbCrLf & _
           " where 1 = 1 " & VbCrLf
     If Mid(Session("username"),1,2) = "RE" Then
        SQL = SQL & "   and b.ds_username = '" & Session("USERNAME") & "' " & VbCrLf
     End If
     If Request("C") > "" Then SQL = SQL & "   and IsNull(h.existe,0) > 0 " & VbCrLf End If
     If Request("D") > "" Then SQL = SQL & "   and d.dt_alteracao     is not null " & VbCrLf End If
     If Request("E") > "" Then 
        If Request("E") = "S" Then 
           SQL = SQL & "   and i.sq_cliente       is not null " & VbCrLf
        Else
           SQL = SQL & "   and i.sq_cliente       is null " & VbCrLf
        End If
     End If

     If Request("p_regional") > "" Then 
        sql = sql + "    and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf
     Else
        sql = sql + "    and g.tipo = 3" & VbCrLf
     End If
          
     If Request("p_tipo_cliente") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("p_tipo_cliente")          & VbCrLf End If

     if sql1 > "" then
       sql = sql & _
             "    and (0 < (select count(*) from escEspecialidade_Cliente where sq_cliente = d.sq_cliente and cast(sq_codigo_espec as varchar) in (" + sql1 + ")) or " & VbCrLf & _
             "         0 < (select count(*) from escTurma_Modalidade  w INNER JOIN escTurma x ON (w.serie = x.ds_serie) where x.sq_site_cliente = d.sq_cliente and w.curso in (" + sql1 + ")) " & VbCrLf & _
             "        ) " & VbCrLf 
     end if
     sql = sql + "ORDER BY d.ds_cliente " & VbCrLf
     ConectaBD SQL

     ShowHTML "<TR><TD valign=""top""><br><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
     If Not RS.EOF Then

        If p_tipo = "H" Then 
           If Request("P4") > "" Then RS.PageSize = cDbl(Request("P4")) Else RS.PageSize = 15 End If
           rs.AbsolutePage = Nvl(Request("P3"),1)
        Else
           RS.PageSize = RS.RecordCount + 1
           rs.AbsolutePage = 1
        End If
      

        ShowHTML "<tr><td><td align=""right""><b><font face=Verdana size=1>Registros encontrados: " & RS.RecordCount & "</font></b>"
        If p_Tipo = "H" Then ShowHTML "     &nbsp;&nbsp;<A TITLE=""Clique aqui para gerar arquivo Word com a listagem abaixo"" class=""SS"" href=""#""  onClick=""window.open('Controle.asp?p_tipo=W&w_ew=" & w_ew & "&Q=" & Request("Q") & "&C=" & Request("C") & "&D=" & Request("D") & "&U=" & Request("U") & w_especialidade & MontaFiltro("GET") & "','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no');"">Gerar Word<IMG ALIGN=""CENTER"" border=0 SRC=""img/word.gif""></A>" End If
        ShowHTML "<tr><td><td>"
        ShowHTML "<table border=""1"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
        ShowHTML "<tr align=""center"" valign=""top"">"
        ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Escola</b></td>"
        If Session("username") = "SEDF" or Session("username") = "CTIS" or Mid(Session("username"),1,2) = "RE" or Session("username") = "SBPI" Then
           If Instr(p_campos,"username") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Username</b></td>" End If
           If Instr(p_campos,"senha") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Senha</b></td>" End If
        End If
        'ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Alunos</b></td>"
        If Instr(p_campos,"alteracao")   > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Última alteração</b></td>" End If
        If Instr(p_campos,"diretor")     > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Diretor</b></td>" End If
        If Instr(p_campos,"secretario")  > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Secretario</b></td>" End If
        If Instr(p_campos,"contato")     > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Contato</b></td>" End If
        If Instr(p_campos,"mail")        > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>e-Mail</b></td>" End If
        If Instr(p_campos,"telefone")    > 0 Then 
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Telefone</b></td>" 
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Fax</b></td>" 
        End If
        If Instr(p_campos,"endereco")    > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Endereço</b></td>" End If
        If Instr(p_campos,"bairro")      > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Bairro</b></td>" End If
        If Instr(p_campos,"localizacao") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Localização</b></td>" End If
        If Instr(p_campos,"cep")         > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>CEP</b></td>" End If
        If p_tipo = "H" Then
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Outros</b></td>"
        End If
  
        w_cor = "#FDFDFD"
        While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(Nvl(Request("P3"),RS.AbsolutePage))
          If RS("ativo")="Nao" Then 
             w_cor = "#31BCBC" 
          ElseIf w_cor = "#EFEFEF" or w_cor = "" Then 
             w_cor = "#FDFDFD" 
          Else 
             w_cor = "#EFEFEF" 
          End If
          ShowHTML "<tr valign=""top"" bgcolor=""" & w_cor & """>"
          ShowHTML "    <td><font face=""Verdana"" size=""1"">"
          If RS("ativo")="Nao" Then 
             ShowHTML "      (DESATIVADA) "
          End If
          If p_tipo = "H" Then 
             ShowHTML "      <a href=""" & RS("LN_INTERNET") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b></font></td>"
          Else
             ShowHTML "      " & RS("DS_CLIENTE") & "</font></td>"
          End If
          If Session("username") = "SEDF" or Session("username") = "CTIS" or Mid(Session("username"),1,2) = "RE" or Session("username") = "SBPI" Then
             If Instr(p_campos,"username") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & RS("DS_USERNAME") & "</font></td>" End If
             If Instr(p_campos,"senha") > 0 Then ShowHTML "    <td align=""center""><font face=""Verdana"" size=""1"">" & RS("DS_SENHA_ACESSO") & "</font></td>" End If
          End If
          'ShowHTML "    <td align=""right""><font face=""Verdana"" size=""1"">" & RS("alunos") & "</font></td>"
          If Instr(p_campos,"alteracao")   > 0 Then ShowHTML "    <td align=""center""><font face=""Verdana"" size=""1"">" & Nvl(RS("dt_alteracao"),"---") & "</font></td>" End If
          If Instr(p_campos,"diretor")     > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("no_diretor"),"---") & "</td>" End If
          If Instr(p_campos,"secretario")  > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("no_secretario"),"---") & "</td>" End If
          If Instr(p_campos,"contato")     > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("no_contato_internet"),"---") & "</td>" End If
          If Instr(p_campos,"mail")        > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("ds_email_internet"),"---") & "</td>" End If
          If Instr(p_campos,"telefone")    > 0 Then 
             ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("nr_fone_internet"),"---") & "</td>" 
             ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("nr_fax_internet"),"---") & "</td>" 
          End If
          If Instr(p_campos,"endereco")    > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("ds_logradouro"),"---") & "</td>" End If
          If Instr(p_campos,"bairro")      > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("no_bairro"),"---") & "</td>" End If
          If Instr(p_campos,"localizacao") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & RS("no_municipio") & "-" & RS("sg_uf") & "</font></td>" End If
          If Instr(p_campos,"cep")         > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("nr_cep"),"---") & "</td>" End If
          If p_tipo = "H" Then
             ShowHTML "    <td><font face=""Verdana"" size=""1"">"
             If Session("username") = "SBPI" Then
                ShowHTML "       <A CLASS=""SS"" HREF=""" & w_Pagina & "CadastroEscola" & "&CL=sq_cliente=" & RS("sq_cliente") & "&w_chave=" & RS("sq_cliente") & "&w_ea=A" & MontaFiltro("GETX") & """ Title=""Alteração dos dados da escola!"">Alt</A>"
             End If             
             ShowHTML "       <A CLASS=""SS"" HREF=""Manut.asp?CL=sq_cliente=" & RS("sq_cliente") & "&w_ea=L&w_ew=Log&w_ee=1&P3=1&P4=30"" Title=""Exibe o registro de ocorrências da escola!"" target=""_blank"">Log</A>"
             If nvl(RS("adm"),"nulo") <> "nulo" Then
                ShowHTML "       <A CLASS=""SS"" HREF=""Controle.asp?CL=sq_cliente=" & RS("sq_cliente") & "&w_ea=L&w_ew=Adm&w_ee=1&P3=1&P4=30"" Title=""Exibe o formulário de dados administrativos preenchido pela escola!"" target=""_blank"">Adm</A>"
             End If
          End If
          RS.MoveNext
          Wend
    
        ShowHTML "</table>"
        ShowHTML "<tr><td><td colspan=""5"" align=""center""><hr>"

        If p_tipo = "H" Then
           ShowHTML "<tr><td align=""center"" colspan=5>"
           MontaBarra "Controle.asp?CL=" & CL & "&w_ew=" & w_ew & "&Q=" & Request("Q") & "&C=" & Request("C") & "&D=" & Request("D") & "&U=" & Request("U") & w_especialidade, cDbl(RS.PageCount), cDbl(Request("P3")), cDbl(Request("P4")), cDbl(RS.RecordCount)
           ShowHTML "</tr>"
        End If

     Else

        ShowHTML "<TR><TD><TD colspan=""3""><p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<font size=""2""><b>Nenhuma ocorrência encontrada para as opções acima."

     End If

  End If
  ShowHTML "</TABLE>"
  
  Set p_regional = Nothing
End Sub


REM =========================================================================
REM Monta a tela de Pesquisa
REM -------------------------------------------------------------------------
Public Sub ShowEscolaParticular
  Dim RS1, p_regiao

  Dim sql, sql2, wCont, sql1, wAtual, wIN, w_especialidade
  
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  
  p_regiao = Request("p_regiao")

  If p_tipo = "W" Then
      Response.ContentType = "application/msword"
      HeaderWord p_layout
      ShowHTML "<TABLE WIDTH=""100%"" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR=""#000000"">SIGE-WEB<TD ALIGN=""RIGHT""><B><FONT SIZE=4 COLOR=""#000000"">"
      ShowHTML "Consulta a escolas"
      ShowHTML "</FONT><TR><TD ALIGN=""RIGHT""><B><FONT SIZE=2 COLOR=""#000000"">" & DataHora() & "</B></TD></TR>"
      ShowHTML "</FONT></B></TD></TR></TABLE>"
      ShowHTML "<HR>"
  Else
     Cabecalho
     ShowHTML "<HEAD>"
     
     ShowHTML "</HEAD>"
     If Request("pesquisa") > "" Then
        BodyOpen " onLoad=""location.href='#lista'"""
     Else
        BodyOpen "onLoad='document.focus()';"
     End If
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
  ShowHTML "    <table width=""95%"" border=""0"">"
  If p_tipo = "H" Then
     Showhtml "<FORM ACTION=""controle.asp"" name=""Form"" METHOD=""POST"">"
     ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""w_ew"" VALUE=""ESCPART"">"
     ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""CL"" VALUE=""" & CL &  """>"
     ShowHTML "<INPUT TYPE=""HIDDEN"" NAME=""pesquisa"" VALUE=""X"">"
     ShowHTML "<input type=""Hidden"" name=""P3"" value=""1"">"
     ShowHTML "<input type=""Hidden"" name=""P4"" value=""15"">"
     ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
     ShowHTML "    <table width=""100%"" border=""0"">"
     ShowHTML "          <TR><TD valign=""top""><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
  Else
     ShowHTML "<tr><td><div align=""justify""><font size=1><b>Filtro:</b><ul>"
  End If
  If p_tipo = "H" Then
     ShowHTML "          <tr valign=""top""><td>"
     SelecaoRegiaoAdm "Região a<u>d</u>ministrativa:", "D", "Indique a região administrativa.", p_regiao, "p_regiao", "p_regiao", null, null
  ElseIf Nvl(p_regiao,0) > 0 Then
     SQL = "SELECT  a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " & VbCrLf & _
           "  FROM  escCLIENTE a " & VbCrLf & _
           " WHERE  a.sq_cliente = " & p_regiao & VbCrLf & _
           "ORDER BY a.ds_cliente "
     ConectaBD SQL
     Response.Write "          <li><font size=""1""><b>Escolas da " & RS("ds_cliente") & "</b>"
     DesconectaBD
  End If
  'SQL = "SELECT * FROM escTipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
  'ConectaBD SQL
  'If p_tipo = "H" Then
  '   ShowHTML "          <tr valign=""top""><td><td><font size=""1""><br><b>Tipo de instituição:</b><br><SELECT CLASS=""STI"" NAME=""p_tipo_cliente"">"
  '   If RS.RecordCount > 1 Then ShowHTML "          <option value="""">Todos" End If
  '   While Not RS.EOF
  '      If cDbl(nvl(RS("sq_tipo_cliente"),0)) = cDbl(nvl(Request("p_tipo_cliente"),0)) Then
  '         ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """ SELECTED>" & RS("ds_tipo_cliente")
  '      Else
  '         ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """>" & RS("ds_tipo_cliente")
  '      End If
  '      RS.MoveNext
  '   Wend
  '   ShowHTML "          </select>"
  'ElseIf nvl(Request("p_tipo_cliente"),0) > 0 Then
  '   ShowHTML "          <li><font size=""1""><b>Tipo de instituição: "
  '   While Not RS.EOF
  '      If cDbl(nvl(RS("sq_tipo_cliente"),0)) = cDbl(nvl(Request("p_tipo_cliente"),0)) Then ShowHTML RS("ds_tipo_cliente") End If
  '      RS.MoveNext
  '   Wend
  '   ShowHTML "</b>"
  'End If
  If p_tipo = "H" Then
     ShowHTML "  <TR><TD><TD><font size=""1""><br>"
     If Request("C") > "" Then
        ShowHTML "          <input type=""checkbox"" name=""C"" value=""X"" CLASS=""BTM"" checked> Exibir apenas escolas com calendário(s) cadastrado(s) "
     Else
        ShowHTML "          <input type=""checkbox"" name=""C"" value=""X"" CLASS=""BTM""> Exibir apenas escolas com calendário(s) cadastrado(s)  "
     End If
  ElseIf Request("C") > "" Then
     ShowHTML "  <li><font size=""1""><b>Apenas escolas com alunos carregados </b>"
  End If
  
  'If p_tipo = "H" Then
  '   ShowHTML "  <TR><TD><TD><font size=""1"">"
  '   If Request("D") > "" Then
  '      ShowHTML "          <input type=""checkbox"" name=""D"" value=""X"" CLASS=""BTM"" checked> Exibir apenas escolas que tenham alterado seus dados "
  '   Else
  '      ShowHTML "          <input type=""checkbox"" name=""D"" value=""X"" CLASS=""BTM""> Exibir apenas escolas que tenham alterado seus dados "
  '   End If
  'ElseIf Request("D") > "" Then
  '   ShowHTML "  <li><font size=""1""><b>Apenas escolas que tenham alterado seus dados </b>"
  'End If
  'If p_tipo = "H" Then
  '   ShowHTML "  <TR><TD><TD><font size=""1""><br><b>Quanto às informações administrativas?</b><br>"
  '   If Request("E") = "S" Then ShowHTML "          <input type=""radio"" name=""E"" value=""S"" CLASS=""BTM"" checked> Listar apenas as escolas que informaram<br>"     Else ShowHTML "          <input type=""radio"" name=""E"" value=""S"" CLASS=""BTM""> Listar apenas as escolas que informaram<br>"      End If
  '   If Request("E") = "N" Then ShowHTML "          <input type=""radio"" name=""E"" value=""N"" CLASS=""BTM"" checked> Listar apenas as escolas que não informaram<br>" Else ShowHTML "          <input type=""radio"" name=""E"" value=""N"" CLASS=""BTM""> Listar apenas as escolas que não informaram<br> " End If
  '   If Request("E") = ""  Then ShowHTML "          <input type=""radio"" name=""E"" value="""" CLASS=""BTM"" checked> Tanto faz"                                        Else ShowHTML "          <input type=""radio"" name=""E"" value="""" CLASS=""BTM""> Tanto faz "                                        End If
  'ElseIf Request("E") > "" Then
  '   ShowHTML "  <li><font size=""1""><b>"
  '   If Request("E") = "S" Then ShowHTML "          Listar apenas as escolas que informaram dados administrativos</b>"      End If
  '   If Request("E") = "N" Then ShowHTML "          Listar apenas as escolas que não informaram dados administrativos</b> " End If
  'End If
  'If p_tipo <> "W" Then
  '   ShowHTML "          </table>"
  'End If
  'DesconectaBD
  'wCont = 0
  'sql1 = ""

'  If p_tipo = "H" Then
'     sql = "SELECT DISTINCT a.* " & _ 
'           "from escEspecialidade AS a " & _ 
'           "     INNER JOIN escEspecialidade_cliente AS c ON (a.sq_especialidade = c.sq_codigo_espec) " & _
'           "     INNER JOIN escCliente AS d ON (c.sq_cliente = d.sq_cliente) " & _
'           "ORDER BY a.nr_ordem, a.ds_especialidade "
' 
'     ConectaBD SQL
'   
'     If Not RS.EOF Then
'        wCont = 0
'        wAtual = ""
' 
'        ShowHTML "          <TD valign=""top""><table border=""0"" align=""left"" cellpadding=0 cellspacing=0>"
'        Do While Not RS.EOF
'           If wAtual = "" or wAtual <> RS("tp_especialidade") Then
'              wAtual = RS("tp_especialidade")
'              If wAtual = "M" Then
'                 ShowHTML "            <TR><TD colspan=2><font size=""1""><b>Etapas/Modalidades de ensino:</b>"
'              ElseIf wAtual = "R" Then
'                 ShowHTML "            <TR><TD colspan=2><font size=""1""><b>Em Regime de Intercomplementaridade:</b>"
'              Else
'                 ShowHTML "            <TR><TD colspan=2><font size=""1""><b>Outras:</b>"
'              End If
'           End If
'           
'           wCont = wCont + 1
'           marcado = "N"
'           For i = 1 to Request("p_modalidade").Count
'               If cDbl(RS("sq_especialidade")) = cDbl(Nvl(Request("p_modalidade")(i),0)) Then marcado = "S" End If
'           Next
'              
'           If marcado = "S" Then
'              ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""p_modalidade"" value=""" & RS("sq_especialidade") & """ checked><td><font size=1>" & RS("ds_especialidade")
'              sql1 = Request("p_modalidade")
'              wIN = 1
'           Else
'              ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""p_modalidade"" value=""" & RS("sq_especialidade") & """><td><font size=1>" & RS("ds_especialidade")
'           End If
'           RS.MoveNext
' 
'           If (wCont Mod 2) = 0 Then 
'              wCont = 0
'           End If
' 
'       Loop
'        DesconectaBD
'     End If
'  ElseIf Nvl(Request("p_modalidade"), "") > "" Then
'     sql = "SELECT DISTINCT a.* " & _ 
'           "from escEspecialidade AS a " & _ 
'           "     INNER JOIN escEspecialidade_cliente AS c ON (a.sq_especialidade = c.sq_codigo_espec) " & _
'           "     INNER JOIN escCliente AS d ON (c.sq_cliente = d.sq_cliente) " & _
'           "where a.sq_especialidade in (" & Request("p_modalidade") & ") " & _
'           "ORDER BY a.nr_ordem, a.ds_especialidade "
' 
'     ConectaBD SQL
'   
'     If Not RS.EOF Then
'        wCont = 0
' 
'        ShowHTML "          <li><b>Modalidades de ensino:</b><ul>"
'        Do While Not RS.EOF
'                    
'           ShowHTML chr(13) & "           <li><font size=1>" & RS("ds_especialidade")
'           sql1 = Request("p_modalidade")
'           wIN = 1
'           RS.MoveNext
' 
'        Loop
'        DesconectaBD
'     End If
'  End If
  ShowHTML "          </tr>"
  ShowHTML "          </table>"
  if p_tipo = "H" Then 
     ShowHTML "  <TR><TD colspan=2><font size=""1""><b>Campos a serem exibidos"
     If p_layout = "PORTRAIT" Then 
        ShowHTML "          (<input type=""radio"" name=""p_layout"" value=""PORTRAIT"" CLASS=""BTM"" checked> Retrato <input type=""radio"" name=""p_layout"" value=""LANDSCAPE"" CLASS=""BTM""> Paisagem)"
     Else
        ShowHTML "          (<input type=""radio"" name=""p_layout"" value=""PORTRAIT"" CLASS=""BTM""> Retrato <input type=""radio"" name=""p_layout"" value=""LANDSCAPE"" CLASS=""BTM"" checked> Paisagem)"
     End If
     ShowHTML "     <table width=""100%"" border=0>"
     ShowHTML "       <tr valign=""top"">"
     If Session("username") = "SEDF" or Session("username") = "SBPI" or Session("username") = "CTIS" or Mid(Session("username"),1,2) = "RE" Then
        If Instr(p_campos,"username")    > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""username"" CLASS=""BTM"" checked>Username"           Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""username"" CLASS=""BTM"">Username"          End If
        If Instr(p_campos,"senha")       > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""senha"" CLASS=""BTM"" checked>Senha"                 Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""senha"" CLASS=""BTM"">Senha"                End If
     End If
     If Instr(p_campos,"alteracao")   > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""alteracao"" CLASS=""BTM"" checked>Última alteração"  Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""alteracao"" CLASS=""BTM"">Última alteração" End If
     If Instr(p_campos,"diretor")     > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""diretor"" CLASS=""BTM"" checked>Diretor"             Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""diretor"" CLASS=""BTM"">Diretor"            End If
     If Instr(p_campos,"secretario")  > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""secretario"" CLASS=""BTM"" checked>Secretário"       Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""secretario"" CLASS=""BTM"">Secretário"      End If
     If Instr(p_campos,"contato")     > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""contato"" CLASS=""BTM"" checked>Contato"             Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""contato"" CLASS=""BTM"">Contato"            End If
     ShowHTML "       <tr valign=""top"">"
     If Instr(p_campos,"mail")        > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""mail"" CLASS=""BTM"" checked>e-Mail"                 Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""mail"" CLASS=""BTM"">e-Mail"                End If
     If Instr(p_campos,"telefone")    > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""telefone"" CLASS=""BTM"" checked>Telefone"           Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""telefone"" CLASS=""BTM"">Telefone"          End If
     If Instr(p_campos,"endereco")    > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""endereco"" CLASS=""BTM"" checked>Endereço"           Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""endereco"" CLASS=""BTM"">Endereço"          End If
     If Instr(p_campos,"localizacao") > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""localizacao"" CLASS=""BTM"" checked>Localização"     Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""localizacao"" CLASS=""BTM"">Localização"    End If
     If Instr(p_campos,"cep")         > 0 Then ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""cep"" CLASS=""BTM"" checked>CEP"                     Else ShowHTML "          <td><font size=1><input type=""checkbox"" name=""p_campos"" value=""cep"" CLASS=""BTM"">CEP"                    End If
     ShowHTML "     </table>"
     ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000"">"
     ShowHTML "      <tr><td align=""center"" colspan=""3"">"
     ShowHTML "            <input class=""BTM"" type=""submit"" name=""Botao"" value=""Aplicar filtro"">"
     If Session("username") = "SBPI" Then
        ShowHTML "            <input class=""BTM"" type=""button"" name=""Botao"" onClick=""location.href='" & w_Pagina & "CadastroEscola" & "&CL=" & CL & MontaFiltro("GET") & "&w_ea=I';"" value=""Nova escola"">"
     End If
     ShowHTML "          </td>"
     ShowHTML "      </tr>"
  End If
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  if p_tipo = "H" Then ShowHTML "</form>" End If
  If Request("pesquisa") > "" Then
     'sql = " SELECT DISTINCT a.sq_cliente, a.ds_cliente, a.ds_apelido, a.ln_internet, a.ds_username, a.ds_senha_acesso, a.no_municipio, a.sg_uf, a.dt_alteracao, d.diretor, d.secretario, d.telefone_1, d.fax, d.cep, d.endereco, d.email_1 " & VbCrLf & _ 
     '      "   from escCliente a " & VbCrLf & _ 
     '      "        INNER JOIN escTipo_Cliente b ON (a.sq_tipo_cliente = b.sq_tipo_cliente) " & VbCrLf & _ 
     '      "--        INNER JOIN escCalendario_cliente c ON (a.sq_cliente = c.sq_site_cliente) " & VbCrLf & _ 
     '      "        INNER JOIN escCliente_Particular d ON (a.sq_cliente = d.sq_cliente) " & VbCrLf & _ 
     '      "        where a.sq_tipo_cliente = 14 "
           
           
     sql = " SELECT DISTINCT a.sq_cliente, a.ds_cliente, a.ds_apelido, a.ln_internet, a.ds_username, a.ds_senha_acesso, a.no_municipio, a.sg_uf, a.dt_alteracao, d.diretor, d.secretario, d.telefone_1, d.fax, d.cep, d.endereco, d.email_1 " & VbCrLf & _ 
           "   from escCliente a " & VbCrLf & _ 
           "        INNER JOIN escTipo_Cliente b ON (a.sq_tipo_cliente = b.sq_tipo_cliente and b.tipo=4) "
           if (Request("c") > "") Then 
              sql = sql & "      INNER JOIN escCalendario_cliente c ON (a.sq_cliente = c.sq_site_cliente) " & VbCrLf & _ 
              "        INNER JOIN escCliente_Particular d ON (a.sq_cliente = d.sq_cliente) " & VbCrLf
           else
              sql = sql & VbCrLf & _ 
              "        INNER JOIN escCliente_Particular d ON (a.sq_cliente = d.sq_cliente) " & VbCrLf
           end if
           

           If (p_regiao > "") Then
              sql = sql & "and a.sq_regiao_adm = " & p_regiao & VbCrLf & _ 
              "        order by a.ds_cliente "
           else
              sql = sql & VbCrLf & _ 
              "        order by a.ds_cliente "
           End If           
        
     ConectaBD SQL
     
     ShowHTML "<TR><TD valign=""top""><br><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
     If Not RS.EOF Then

        If p_tipo = "H" Then 
           If Request("P4") > "" Then RS.PageSize = cDbl(Request("P4")) Else RS.PageSize = 15 End If
           rs.AbsolutePage = Nvl(Request("P3"),1)
        Else
           RS.PageSize = RS.RecordCount + 1
           rs.AbsolutePage = 1
        End If
      

        ShowHTML "<tr><td><td align=""right""><b><font face=Verdana size=1>Registros encontrados: " & RS.RecordCount & "</font></b>"
        If p_Tipo = "H" Then ShowHTML "     &nbsp;&nbsp;<A TITLE=""Clique aqui para gerar arquivo Word com a listagem abaixo"" class=""SS"" href=""#""  onClick=""window.open('Controle.asp?p_tipo=W&w_ew=" & w_ew & "&Q=" & Request("Q") & "&C=" & Request("C") & "&D=" & Request("D") & "&U=" & Request("U") & w_especialidade & MontaFiltro("GET") & "','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no');"">Gerar Word<IMG ALIGN=""CENTER"" border=0 SRC=""img/word.gif""></A>" End If
        ShowHTML "<tr><td><td>"
        ShowHTML "<table border=""1"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
        ShowHTML "<tr align=""center"" valign=""top"">"
        ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Escola</b></td>"
        If Session("username") = "SEDF" or Session("username") = "CTIS" or Mid(Session("username"),1,2) = "RE" or Session("username") = "SBPI" Then
           If Instr(p_campos,"username") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Username</b></td>" End If
           If Instr(p_campos,"senha") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Senha</b></td>" End If
        End If
        'ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Alunos</b></td>"
        If Instr(p_campos,"alteracao")   > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Última alteração</b></td>" End If
        If Instr(p_campos,"diretor")     > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Diretor</b></td>" End If
        If Instr(p_campos,"secretario")  > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Secretario</b></td>" End If
        If Instr(p_campos,"mail")        > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>e-Mail</b></td>" End If
        If Instr(p_campos,"telefone")    > 0 Then 
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Telefone</b></td>" 
           ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Fax</b></td>" 
        End If
        If Instr(p_campos,"endereco")    > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Endereço</b></td>" End If
        If Instr(p_campos,"localizacao") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Localização</b></td>" End If
        If Instr(p_campos,"cep")         > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1""><b>CEP</b></td>" End If
        w_cor = "#FDFDFD"
        While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(Nvl(Request("P3"),RS.AbsolutePage))
          If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
          ShowHTML "<tr valign=""top"" bgcolor=""" & w_cor & """>"
          If p_tipo = "H" Then
          
             If(RS("LN_INTERNET") > "") Then
                ShowHTML "    <td><font face=""Verdana"" size=""1""><a href=""" & RS("LN_INTERNET") & """ target=""_blank"">" & RS("DS_CLIENTE") & "</a></b></font></td>"
             Else
                ShowHTML "    <td><font face=""Verdana"" size=""1""><b>" & RS("DS_CLIENTE") & "</b></font></td>"
             End if                
          Else
             ShowHTML "    <td><font face=""Verdana"" size=""1"">" & RS("DS_CLIENTE") & "</font></td>"
          End If
          If Session("username") = "SEDF" or Session("username") = "CTIS" or Mid(Session("username"),1,2) = "RE" or Session("username") = "SBPI" Then
             If Instr(p_campos,"username") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & RS("DS_USERNAME") & "</font></td>" End If
             If Instr(p_campos,"senha") > 0 Then ShowHTML "    <td align=""center""><font face=""Verdana"" size=""1"">" & RS("DS_SENHA_ACESSO") & "</font></td>" End If
          End If
          'ShowHTML "    <td align=""right""><font face=""Verdana"" size=""1"">" & RS("alunos") & "</font></td>"
          If Instr(p_campos,"alteracao")   > 0 Then ShowHTML "    <td align=""center""><font face=""Verdana"" size=""1"">" & Nvl(RS("dt_alteracao"),"---") & "</font></td>" End If
          If Instr(p_campos,"diretor")     > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("diretor"),"---") & "</td>" End If
          If Instr(p_campos,"secretario")  > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("secretario"),"---") & "</td>" End If
          If Instr(p_campos,"mail")        > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("email_1"),"---") & "</td>" End If
          If Instr(p_campos,"telefone")    > 0 Then 
             ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("telefone_1"),"---") & "</td>" 
             ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("fax"),"---") & "</td>" 
          End If
          If Instr(p_campos,"endereco")    > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("endereco"),"---") & "</td>" End If
          If Instr(p_campos,"localizacao") > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & RS("no_municipio") & "-" & RS("sg_uf") & "</font></td>" End If
          If Instr(p_campos,"cep")         > 0 Then ShowHTML "    <td><font face=""Verdana"" size=""1"">" & Nvl(RS("cep"),"---") & "</td>" End If
          If p_tipo = "H" Then
             ShowHTML "    <td><font face=""Verdana"" size=""1"">"
             If Session("username") = "SBPI" Then
                ShowHTML "       <A CLASS=""SS"" HREF=""" & w_Pagina & "CadastroEscola" & "&w_chave=" & RS("sq_cliente") & "&w_ea=A" & MontaFiltro("GET") & """ Title=""Alteração dos dados da escola!"">Alt</A>"
             End If             
'             If nvl(RS("adm"),"nulo") <> "nulo" Then
'                ShowHTML "       <A CLASS=""SS"" HREF=""controle.asp?CL=sq_cliente=" & RS("sq_cliente") & "&w_ea=L&w_ew=Adm&w_ee=1&P3=1&P4=30"" Title=""Exibe o formulário de dados administrativos preenchido pela escola!"" target=""_blank"">Adm</A>"
'             End If
          End If
          RS.MoveNext
          Wend
    
        ShowHTML "</table>"
        ShowHTML "<tr><td><td colspan=""5"" align=""center""><hr>"

        If p_tipo = "H" Then
           ShowHTML "<tr><td align=""center"" colspan=5>"
           MontaBarra "Controle.asp?CL=" & CL & "&w_ew=" & w_ew & "&Q=" & Request("Q") & "&C=" & Request("C") & "&D=" & Request("D") & "&U=" & Request("U") & w_especialidade, cDbl(RS.PageCount), cDbl(Request("P3")), cDbl(Request("P4")), cDbl(RS.RecordCount)
           ShowHTML "</tr>"
        End If

     Else

        ShowHTML "<TR><TD><TD colspan=""3""><p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<font size=""2""><b>Nenhuma ocorrência encontrada para as opções acima."

     End If

  End If
  ShowHTML "</TABLE>"
  
  Set p_regiao = Nothing
End Sub
REM Fim da Pesquisa de Escolas Particulares


REM =========================================================================
REM Monta a tela de Homologação do Calendário das Escolas Particulares
REM -------------------------------------------------------------------------
Public Sub ShowEscolaParticularHomolog

  Dim RS1, p_regiao

  Dim sql, sql2, wCont, sql1, wAtual, wIN, w_especialidade
  
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  
  p_regiao = Request("p_regiao")
  w_homologado = Request("w_homologado")
  if(w_homologado <> "S") Then
    w_homologado = "Não"
  Else
    w_homologado = "Sim"
  End If

  If p_tipo = "W" Then
      Response.ContentType = "application/msword"
      HeaderWord p_layout
      ShowHTML "<TABLE WIDTH=""100%"" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR=""#000000"">SIGE-WEB<TD ALIGN=""RIGHT""><B><FONT SIZE=4 COLOR=""#000000"">"
      ShowHTML "Consulta a escolas"
      ShowHTML "</FONT><TR><TD ALIGN=""RIGHT""><B><FONT SIZE=2 COLOR=""#000000"">" & DataHora() & "</B></TD></TR>"
      ShowHTML "</FONT></B></TD></TR></TABLE>"
      ShowHTML "<HR>"
  Else
     Cabecalho
     ShowHTML "<HEAD>"
     ShowHTML "   <link href=""/css/particular.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
     ShowHTML "</HEAD>"
     If Request("pesquisa") > "" Then
        BodyOpen " onLoad=""location.href='#lista'"""
     'Else
        'BodyOpen "onLoad='document.Form.p_regional.focus()';"
     End If
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
  ShowHTML "<B><FONT size=""2"" COLOR=""#000000"">Vinculação e tipologia</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">" 
  ShowHTML "      <tr bgcolor=""" & conTrBgColor & """><td align=""center"" valign=""top""><table border=0 width=""90%"" cellspacing=0>"
  AbreForm "Form", w_Pagina & "Grava", "POST", "return(Validacao(this));", null
  ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & w_ew & """>"
  ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"
  ShowHTML "<INPUT type=""hidden"" name=""w_troca"" value="""">"
  ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
  ShowHTML "      <tr><td colspan=2><table border=0 width=""90%"" cellspacing=0>"
  ShowHTML "        <tr valign=""top"">"
  SelecaoEscolaParticular         "Unidad<u>e</u> de ensino:", "E", "Selecione unidade." , p_escola_particular, null, "p_escola_particular", null, "onChange=""document.Form.action='" & w_pagina & w_ew & "'; document.Form.w_ea.value='P'; document.Form.w_troca.value='p_calendario'; document.Form.submit();"""
  'ShowHTML "        <tr valign=""top"">"
  'SelecaoCalendarioParticular         "<u>C</u>alendário:", "E", "Selecione unidade." , p_calendario, p_escola_particular, "p_calendario", null, "onChange=""document.Form.action='" & w_pagina & w_ew & "'; document.Form.w_ea.value='P'; document.Form.w_troca.value='w_homologado'; document.Form.submit();"""
  'ShowHTML "        <tr valign=""top"">"
  'SelecaoRegionalEscola "Regional de en<u>s</u>ino:", "S", "Indique a regional de ensino.", p_regional, p_escola_particular, "p_regional", null, null
  'ShowHTML "        <tr valign=""top"">"
  'SelecaoTipoEscola     "<u>T</u>ipo de Escola:", "T", "Selecione o tipo da Escola.", p_tipo_escola, p_escola_particular, "p_tipo_escola", null, null
  if(p_escola_particular > "") Then
  SQL = "SELECT sq_particular_calendario, sq_cliente as cliente, nome, homologado FROM escParticular_Calendario WHERE sq_cliente = " & p_escola_particular
  ConectaBD SQL
  
  ShowHTML "<font color=""#ff3737""><strong><a href=""javascript:this.status.value;"" onClick=""window.open('calendario.asp?CL=sq_cliente=" & RS("cliente") & "&w_ea=L&w_ew=formcal&w_ee=1','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no');"">Acessar o(s) calendário(s)</a></strong></font>"
    
  ShowHTML "<table id=""tbhomologacao"" border=""1"">"
  ShowHTML "<tr><td>Título do Calendário</td><td>Homologado?</td></tr>"
  
  Dim homologado
  While Not RS.EOF
     homologado = RS("homologado")
     ShowHTML "<INPUT type=""hidden"" name=""w_cliente""    value=""" & RS("cliente") & """>"
     ShowHTML "<tr><td>" &  RS("nome") & "</td>"     
     ShowHTML "<INPUT type=""hidden"" name=""w_chave"" value=""" & RS("sq_particular_calendario") &""">"
     ShowHTML "<td><select name=""w_homologado"">"          
     if(inStr(uCase(homologado),"N"))Then
        ShowHTML "<option value=""N"" SELECTED>Não"
        ShowHTML "<option value=""S"">Sim"
     elseif(inStr(uCase(homologado),"S")) Then
        ShowHTML "<option value=""S"" SELECTED>Sim"
        ShowHTML "<option value=""N"">Não"
     end if     
     ShowHTML "</select></td>"
     'ShowHTML " <td><a href=""http://netuno:8080/calendario.asp?CL=sq_cliente=" & RS("cliente") & "&w_ea=L&w_ew=formcal&w_ee=1"">Acessar</a></td>"
     'if(inStr(uCase(homologado),"N"))Then
     '   ShowHTML "<td><a href=""homolog.asp?homologado="S"&calendario="& RS("sq_particular_calendario") & """>Sim</a>"
     '   ShowHTML "<span>Não</span></td>"          
     'elseif(inStr(uCase(homologado),"S")) Then
     '   ShowHTML "<td><span>Sim</span>"
     '   ShowHTML "<a href=""homolog.asp?homologado=S&calendario="& RS("sq_particular_calendario") & """>Não</a></td>"
     'end if     

     ShowHTML "<td align=""center"" colspan=""2"">"
     'ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Gravar"">"

     ShowHTML "          </td>"
     ShowHTML "</tr>" 
     RS.MoveNext
  Wend  
  ShowHTML "</table>"
  DesconectaBD
  
  End If
  ShowHTML "         <tr valign=""top"">"

  ShowHTML "          </table>"
  ShowHTML "      <tr valign=""top"">"
  ShowHTML "      <tr>"
  ShowHTML "      <tr><td align=""center"" colspan=""2"" height=""1"" bgcolor=""#000000"">"
  ShowHTML "      <tr><td align=""center"" colspan=""2"">"
  ShowHTML "            <input class=""STB"" type=""submit"" name=""Botao"" value=""Gravar"">"
  ShowHTML "          </td>"
  ShowHTML "      </tr>"
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  ShowHTML "</FORM>"
  ShowHTML "</table>"
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape
End Sub
REM Fim da Pesquisa de Escolas Particulares

REM =========================================================================
REM Cadastro de escolas
REM -------------------------------------------------------------------------
Sub CadastroEscolas
  Dim w_chave, w_sq_cliente_pai, w_sq_tipo_cliente, w_ds_cliente, w_ds_apelido, w_no_municipio
  Dim w_sg_uf, w_ds_username, w_ln_internet, w_ds_email, w_regiao, w_local
  Dim w_ds_logradouo, w_no_bairro, w_nr_cep
  Dim w_no_contato, w_email_contato, w_nr_fone_contato, w_nr_fax_contato, w_no_diretor, w_no_secretario
  Dim w_sq_modelo, w_no_contato_internet, w_nr_fone_internet, w_nr_fax_internet, w_ds_email_internet
  Dim w_ds_diretorio, w_sq_siscol, w_ds_mensagem, w_ds_institucional
  Dim w_ds_texto_abertura, w_sq_codigo_espec
  Dim w_diretorio, w_username_atual
  Dim w_ativo
  
  Dim w_troca, i, w_texto
  
  Set RS2 = Server.CreateObject("ADODB.RecordSet")
  
  w_Chave           = Request("w_Chave")
  w_troca           = Request("w_troca")
  
  If InStr("A",w_ea) > 0 and w_Troca = "" Then
     
     ' Recupera os dados do cliente
     SQL = " SELECT a.sq_cliente, sq_cliente_pai, sq_tipo_cliente, ds_cliente, ds_apelido, a.sq_regiao_adm, a.localizacao, " & VbCrLf &_
           "        no_municipio, sg_uf, ds_username, ln_internet, ds_email, nr_cep, no_contato, " & VbCrLf &_
           "        ds_email_contato, nr_fone_contato, nr_fax_contato, no_diretor, no_secretario, " & VbCrLf &_
           "        ds_logradouro, no_bairro, sq_modelo, " & VbCrLf &_
           "        no_contato_internet, nr_fone_internet, nr_fax_internet, ds_email_internet, " & VbCrLf &_
           "        ds_diretorio, sq_siscol, ds_mensagem, " & VbCrLf &_ 
           "        ds_institucional, ds_texto_abertura, a.ativo " & VbCrLf &_
           "   FROM escCliente a " & VbCrLf &_
           "        LEFT OUTER JOIN escCliente_Dados b on (a.sq_cliente   = b.sq_cliente)  " & VbCrLf &_
           "        LEFT OUTER JOIN escCliente_Site  c on (a.sq_cliente   = c.sq_cliente)  " & VbCrLf &_
           "  WHERE a.sq_cliente = " & w_chave  & VbCrLf
     ConectaBD SQL
     w_chave               = RS("sq_cliente")
     w_regiao              = RS("sq_regiao_adm")
     w_local               = RS("localizacao")
     w_sq_cliente_pai      = RS("sq_cliente_pai")
     w_sq_tipo_cliente     = RS("sq_tipo_cliente") 
     w_ds_cliente          = RS("ds_cliente") 
     w_ds_apelido          = RS("ds_apelido") 
     w_no_municipio        = RS("no_municipio")
     w_sg_uf               = RS("sg_uf") 
     w_ds_username         = RS("ds_username") 
     w_ln_internet         = RS("ln_internet")
     w_ds_email            = RS("ds_email")
     w_ds_logradouro       = RS("ds_logradouro")  
     w_no_bairro           = RS("no_bairro")  
     w_nr_cep              = trim(RS("nr_cep"))
     w_no_contato          = RS("no_contato")  
     w_email_contato       = RS("ds_email_contato")  
     w_nr_fone_contato     = RS("nr_fone_contato")   
     w_nr_fax_contato      = RS("nr_fax_contato")   
     w_no_diretor          = RS("no_diretor")   
     w_no_secretario       = RS("no_secretario")  
     w_sq_modelo           = RS("sq_modelo")   
     w_no_contato_internet = RS("no_contato_internet")   
     w_nr_fone_internet    = RS("nr_fone_internet")   
     w_nr_fax_internet     = RS("nr_fax_internet")   
     w_ds_email_internet   = RS("ds_email_internet")  
     w_ds_diretorio        = RS("ds_diretorio")   
     w_sq_siscol           = RS("sq_siscol")   
     w_ds_mensagem         = RS("ds_mensagem")   
     w_ds_institucional    = RS("ds_institucional")  
     w_ds_texto_abertura   = RS("ds_texto_abertura")
     w_username_atual      = RS("ds_username")
     w_ativo               = RS("ativo")
     DesconectaBD
  ElseIf InStr("I",w_ea) > 0 Then
  
  End If
  SQL = "select ds_username " & VbCrLf & _
        "  from escCliente  a " & VbCrLf & _
        " where sq_cliente = 0 "
  ConectaBD SQL
  w_diretorio = RS("ds_username") & "/"
  DesconectaBD

  Cabecalho  
  ShowHTML "<HEAD>"
  ScriptOpen "JavaScript"
  modulo
  checkbranco
  FormataCEP  
  ShowHTML "function montaLink() {"
  ShowHTML "  var link = '" & conSite & conVirtualPath & w_diretorio & "';"
  ShowHTML "  document.Form.w_ln_internet.value = link + document.Form.w_ds_username.value;"
  ShowHTML "  document.Form.w_ds_diretorio.value = link + document.Form.w_ds_username.value;"
  ShowHTML "}"
  ValidateOpen "Validacao"
  Validate "w_sq_cliente_pai"       , "D.R.E."      , "SELECT" , "1" , "1"  , "18" , "1" , "1"
  Validate "w_regiao"               , "Região Administrativa" , "SELECT" , "1" , "1"  , "18" , "1" , "1"
  Validate "w_sq_tipo_cliente"      , "Tipo de instituição" , "SELECT" , "1" , "1"  , "18" , "1" , "1"
  Validate "w_ds_cliente"           , "Nome da escola"          , "1"      , "1" , "3"  , "60" , "1" , "1"
  Validate "w_ds_apelido"           , "Apelido"                 , "1"      ,  "" , "3"  , "30" , "1" , "1"
  Validate "w_ds_username"          , "Username"                , "1"      , "1" , "3"  , "14" , "1" , "1"
  Validate "w_ln_internet"          , "Link para acesso"        , "1"      , "1" , "10" , "60" , "1" , "1"  
  Validate "w_ds_email"             , "e_Mail"                  , "1"      , ""  , "4"  , "60" , "1" , "1"  
  Validate "w_ds_logradouro"        , "Logradouro"              , "1"      , "1" , "4"  , "60" , "1" , "1"
  Validate "w_no_bairro"            , "Bairro"                  , "1"      , ""  , "2"  , "30" , "1" , "1"
  Validate "w_no_municipio"         , "Cidade"                  , "1"      , "1" , "2"  , "30" , "1" , "1"
  Validate "w_sg_uf"                , "UF"                      , "SELECT" , "1" , "1"  , "2"  , "1" , "1"    
  Validate "w_nr_cep"               , "CEP"                     , "1"      , "1" , "9"  , "9"  , ""  , "0123456789-"
  Validate "w_no_contato"           , "Contato técnico"         , "1"      , "1" , "2"  , "35" , "1" , "1"
  Validate "w_nr_fone_contato"      , "Telefone contato tencnico", "1"     , "1" , "6"  , "20" , "1" , "1"
  Validate "w_nr_fax_contato"       , "Fax contato técnico"     , "1"      , ""  , "6"  , "20" , "1" , "1"
  Validate "w_email_contato"        , "e-Mail contato técnico"  , "1"      , ""  , "6"  , "60" , "1" , "1"
  Validate "w_no_diretor"           , "Nome do(a) Diretor(a)"   , "1"      , ""  , "2"  , "40" , "1" , "1"
  Validate "w_no_secretario"        , "Nome do(a) Secretário(a)", "1"      , ""  , "2"  , "40" , "1" , "1"
  Validate "w_no_contato_internet"  , "Contato internet"        , "1"      , "1" , "2"  , "35" , "1" , "1"
  Validate "w_nr_fone_internet"     , "Telefone contato internet", "1"     , "1" , "6"  , "20" , "1" , "1"
  Validate "w_nr_fax_internet"      , "Fax contato internet"    , "1"      , ""  , "6"  , "20" , "1" , "1"
  Validate "w_ds_email_internet"    , "e-Mail contato internet" , "1"      , "1" , "6"  , "60" , "1" , "1"
  Validate "w_sq_modelo"            , "Modelo de site"          , "SELECT" , "1" , "1"  , "18" , "1" , "1" 
  Validate "w_ds_diretorio"         , "Diretório"               , "1"      , "1" , "4"  , "60" , "1" , "1"
  Validate "w_sq_siscol"            , "SISCOL"                  , "1"      ,  "" , "5"  , "6"  , "1" , "1"
  Validate "w_ds_mensagem"          , "Mensagem"                , "1"      ,  "" , "4"  , "80" , "1" , "1"
  Validate "w_ds_institucional"     , "Institucional"           , "1"      ,  "" , "4"  ,"7000", "1" , "1"
  Validate "w_ds_texto_abertura"    , "Texto de abertura"       , "1"      ,  "" , "4"  ,"7000", "1" , "1"
  ShowHTML "  var i; "
  ShowHTML "  var w_erro=true; "
  ShowHTML "  if (theForm.w_sq_codigo_espec.value==undefined) {"
  ShowHTML "     for (i=0; i < theForm.w_sq_codigo_espec.length; i++) {"
  ShowHTML "       if (theForm.w_sq_codigo_espec[i].checked) w_erro=false;"
  ShowHTML "     }"
  ShowHTML "  }"
  ShowHTML "  else {"
  ShowHTML "     if (theForm.w_sq_codigo_espec.checked) w_erro=false;"
  ShowHTML "  }"
  ShowHTML "  if (w_erro) {"
  ShowHTML "    alert('Você deve informar pelo menos uma etapa/modalidade!'); "
  ShowHTML "    return false;"
  ShowHTML "  }"
  
  ShowHTML "  theForm.Botao[0].disabled=true;"
  ShowHTML "  theForm.Botao[1].disabled=true;"
  ValidateClose
  ScriptClose
  ShowHTML "</HEAD>"
  If w_troca > "" Then
     BodyOpen "onLoad='document.Form." & w_troca & ".focus()';"
  ElseIf w_ea = "I" or w_ea = "A" Then
     BodyOpen "onLoad='document.Form.w_sq_cliente_pai.focus()';"
  Else
     BodyOpen "onLoad='document.focus()';"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">Cadastro de escolas da rede de ensino</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  If InStr("EV",w_ea) Then
     w_Disabled = " DISABLED "
  End If
  ShowHTML "<FORM action=""" & w_pagina & "Grava"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"">"
  ShowHTML MontaFiltro("POSTX")
  ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & w_ew & """>"
  ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
  ShowHTML "<INPUT type=""hidden"" name=""w_chave"" value=""" & w_chave & """>"
  ShowHTML "<INPUT type=""hidden"" name=""w_ea"" value=""" & w_ea & """>"
  If Nvl(w_ea,"") = "A" Then
    ShowHTML "<INPUT type=""hidden"" name=""w_username_atual"" value=""" & w_ds_username & """>"
  End If
  ShowHTML "<tr bgcolor=""" & "#EFEFEF" & """><td align=""center"">"
  ShowHTML "    <table width=""100%"" border=""0"">"
  ShowHTML "        <tr valign=""top"">"
  SelecaoRegional "Regional de En<u>s</u>ino:", "S", "Indique a regional de ensino da escola.", w_sq_cliente_pai, null, "w_sq_cliente_pai", "CADASTRO", null
  SelecaoRegiaoAdm "Região a<u>d</u>ministrativa:", "D", "Indique a região administrativa.", w_regiao, w_chave, "w_regiao", null, null
  SQL = "SELECT * FROM escTipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
  ConectaBD SQL  
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "              <td><font size=""1""><b>Tipo de in<u>s</u>tituição:</b><br><SELECT accesskey=""S"" class=""STI"" NAME=""w_sq_tipo_cliente"">"
  If RS.RecordCount > 1 Then ShowHTML "          <option value="""">---" End If
  While Not RS.EOF
     If cDbl(nvl(RS("sq_tipo_cliente"),0)) = cDbl(nvl(w_sq_tipo_cliente,0)) Then
        ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """ SELECTED>" & RS("ds_tipo_cliente")
     Else
        ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """>" & RS("ds_tipo_cliente")
     End If
     RS.MoveNext
  Wend
  ShowHTML "          </select>"
  DesconectaBD  
  ShowHTML "          <td><font size=""1""><b>Localização</b><br>"
  If w_local = "1" Then
     ShowHTML "              <input  type=""radio"" name=""w_local"" value=""0""> Não informada <input  type=""radio"" name=""w_local"" value=""1"" checked> Urbana <input  type=""radio"" name=""w_local"" value=""2""> Rural "
  ElseIf w_local = "2" Then
     ShowHTML "              <input  type=""radio"" name=""w_local"" value=""0""> Não informada <input  type=""radio"" name=""w_local"" value=""1""> Urbana <input  type=""radio"" name=""w_local"" value=""2"" checked> Rural "
  Else
     ShowHTML "              <input  type=""radio"" name=""w_local"" value=""0"" checked> Não informada <input  type=""radio"" name=""w_local"" value=""1""> Urbana <input  type=""radio"" name=""w_local"" value=""2""> Rural "
  End If
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b>Nome <u>d</u>a escola:</b><br><input " & w_Disabled & " accesskey=""D"" type=""text"" name=""w_ds_cliente"" class=""STI"" SIZE=""60"" MAXLENGTH=""60"" VALUE=""" & w_ds_cliente & """ ONMOUSEOVER=""popup('Informe uma descrição para a escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td valign=""top""><font size=""1""><b><u>A</u>pelido:</b><br><input " & w_Disabled & " accesskey=""A"" type=""text"" name=""w_ds_apelido"" class=""STI"" SIZE=""30"" MAXLENGTH=""30"" VALUE=""" & w_ds_apelido & """ ONMOUSEOVER=""popup('Informe o apelido da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>U</u>sername:</b><br><input " & w_Disabled & " accesskey=""U"" type=""text"" name=""w_ds_username"" class=""STI"" SIZE=""14"" MAXLENGTH=""14"" VALUE=""" & w_ds_username & """ ONMOUSEOVER=""popup('Informe o usernanme para acesso.','white')""; ONMOUSEOUT=""kill()"" ONBLUR=""javascript:montaLink();""></td>"
  MontaRadioSN "<b>Ativo?</b>", w_ativo, "w_ativo"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>L</u>ink para acesso:</b><br><input " & w_Disabled & " accesskey=""L"" type=""text"" name=""w_ln_internet"" class=""STI"" SIZE=""60"" MAXLENGTH=""60"" VALUE=""" & w_ln_internet & """ ONMOUSEOVER=""popup('Informe link de acesso para o site da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td valign=""top""><font size=""1""><b><u>e</u>-Mail:</b><br><input " & w_Disabled & " accesskey=""E"" type=""text"" name=""w_ds_email"" class=""STI"" SIZE=""50"" MAXLENGTH=""60"" VALUE=""" & w_ds_email & """ ONMOUSEOVER=""popup('Informe e-Mail da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b>Lo<u>g</u>radouro:</b><br><input " & w_Disabled & " accesskey=""G"" type=""text"" name=""w_ds_logradouro"" class=""STI"" SIZE=""50"" MAXLENGTH=""60"" VALUE=""" & w_ds_logradouro & """ ONMOUSEOVER=""popup('Informe o logradouro da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td valign=""top""><font size=""1""><b><u>B</u>airro:</b><br><input " & w_Disabled & " accesskey=""B"" type=""text"" name=""w_no_bairro"" class=""STI"" SIZE=""30"" MAXLENGTH=""60"" VALUE=""" & w_no_bairro & """ ONMOUSEOVER=""popup('Informe o bairro da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr valign=""top""><td colspan=""2""><font size=""1"">"
  ShowHTML "        <table width=""100%"" border=""0"">"
  ShowHTML "          <tr><td valign=""top""><font size=""1""><b><u>C</u>idade:</b><br><input " & w_Disabled & " accesskey=""C"" type=""text"" name=""w_no_municipio"" class=""STI"" SIZE=""30"" MAXLENGTH=""30"" VALUE=""" & w_no_municipio & """ ONMOUSEOVER=""popup('Informe a cidade.','white')""; ONMOUSEOUT=""kill()""></td>"
  Dim AC,AM,AP,AL,BA,CE,DF,ES,GO,MA,MT,MS,MG,PA,PB,PR,PE,PI,RJ,RN,RO,RR,RGS,SC,SP,SE,TOC
  Select Case w_sg_uf
      Case "AC" AC = "selected"
     Case "AM" AM = "selected"
     Case "AP" AP = "selected"
     Case "AL" AL = "selected"
     Case "BA" BA = "selected"
     Case "CE" CE = "selected"
     Case "DF" DF = "selected"
     Case "ES" ES = "selected"
     Case "GO" GO = "selected"
     Case "MA" MA = "selected"
     Case "MT" MT = "selected"
     Case "MS" MS = "selected"
     Case "MG" MG = "selected"
     Case "PA" PA = "selected"
     Case "PB" PB = "selected"
     Case "PR" PR = "selected"
     Case "PE" PE = "selected"
     Case "PI" PI = "selected"
     Case "RJ" RJ = "selected"
     Case "RN" RN = "selected"
     Case "RO" RO = "selected"
     Case "RR" RR = "selected"
     Case "RS" RGS = "selected"
     Case "SC" SC = "selected"
     Case "SP" SP = "selected"
     Case "SE" SE = "selected"
     Case "TO" TOC = "selected"
  End Select
  ShowHTML "         <td valign=""top""><font  size=""1""><b>U<U>F</U>:<br></b> "
  ShowHTML "             <select ACCESSKEY=""F"" name=""w_sg_uf"" class=""STI"">"
  ShowHTML "                <option value="""">---</option>"
  ShowHTML "                <option value=""AC"""&AC&">AC</option>"
  ShowHTML "                <option value=""AM"""&AM&">AM</option>"
  ShowHTML "                <option value=""AP"""&AP&">AP</option>"
  ShowHTML "                <option value=""AL"""&AL&">AL</option>"
  ShowHTML "                <option value=""BA"""&BA&">BA</option>"
  ShowHTML "                <option value=""CE"""&CE&">CE</option>"
  ShowHTML "                <option value=""DF"""&DF&">DF</option>"
  ShowHTML "                <option value=""ES"""&ES&">ES</option>"
  ShowHTML "                <option value=""GO"""&GO&">GO</option>"
  ShowHTML "                <option value=""MA"""&MA&">MA</option>"
  ShowHTML "                <option value=""MT"""&MT&">MT</option>"
  ShowHTML "                <option value=""MS"""&MS&">MS</option>"
  ShowHTML "                <option value=""MG"""&MG&">MG</option>"
  ShowHTML "                <option value=""PA"""&PA&">PA</option>"
  ShowHTML "                <option value=""PB"""&PB&">PB</option>"
  ShowHTML "                <option value=""PR"""&PR&">PR</option>"
  ShowHTML "                <option value=""PE"""&PE&">PE</option>"
  ShowHTML "                <option value=""PI"""&PI&">PI</option>"
  ShowHTML "                <option value=""RJ"""&RJ&">RJ</option>"
  ShowHTML "                <option value=""RN"""&RN&">RN</option>"
  ShowHTML "                <option value=""RO"""&RO&">RO</option>"
  ShowHTML "                <option value=""RR"""&RR&">RR</option>"
  ShowHTML "                <option value=""RS"""&RGS&">RS</option>"
  ShowHTML "                <option value=""SC"""&SC&">SC</option>"
  ShowHTML "                <option value=""SP"""&SP&">SP</option>"
  ShowHTML "                <option value=""SE"""&SE&">SE</option>"
  ShowHTML "                <option value=""TO"""&TOC&">TO</option>"
  ShowHTML "             </select>"  
  
  ShowHTML "          <td valign=""top""><font size=""1""><b>Ce<u>p</u>:</b><br><input " & w_Disabled & " accesskey=""P"" type=""text"" name=""w_nr_cep"" class=""sti"" SIZE=""9"" MAXLENGTH=""9"" VALUE=""" & w_nr_cep & """ onKeyDown=""FormataCEP(this,event)"" ONMOUSEOVER=""popup('Informe o CEP deste endereço.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b>Co<u>n</u>tato técnico:</b><br><input " & w_Disabled & " accesskey=""N"" type=""text"" name=""w_no_contato"" class=""STI"" SIZE=""35"" MAXLENGTH=""35"" VALUE=""" & w_no_contato & """ ONMOUSEOVER=""popup('Informe o nome do contato técnico da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td valign=""top""><font size=""1""><b><u>T</u>elefone contato técnico:</b><br><input " & w_Disabled & " accesskey=""T"" type=""text"" name=""w_nr_fone_contato"" class=""STI"" SIZE=""13"" MAXLENGTH=""20"" VALUE=""" & w_nr_fone_contato & """ ONMOUSEOVER=""popup('Informe o telefone do contato técnico da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td valign=""top""><font size=""1""><b>Fa<u>x</u> contato técnico:</b><br><input " & w_Disabled & " accesskey=""X"" type=""text"" name=""w_nr_fax_contato"" class=""STI"" SIZE=""13"" MAXLENGTH=""20"" VALUE=""" & w_nr_fax_contato & """ ONMOUSEOVER=""popup('Informe o fax do contato técnico da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b>e-Mail contato técnic<u>o</u>:</b><br><input " & w_Disabled & " accesskey=""O"" type=""text"" name=""w_email_contato"" class=""STI"" SIZE=""35"" MAXLENGTH=""60"" VALUE=""" & w_email_contato & """ ONMOUSEOVER=""popup('Informe o e-Mail do contato técnico da escola.','white')""; ONMOUSEOUT=""kill()""></td>"  
  ShowHTML "          <td valign=""top""><font size=""1""><b>No<u>m</u>e do(a) Diretor(a):</b><br><input " & w_Disabled & " accesskey=""M"" type=""text"" name=""w_no_diretor"" class=""STI"" SIZE=""40"" MAXLENGTH=""40"" VALUE=""" & w_no_diretor & """ ONMOUSEOVER=""popup('Informe o nome do diretor da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td valign=""top""><font size=""1""><b>No<u>m</u>e do(a) <u>S</u>ecretário(a):</b><br><input " & w_Disabled & " accesskey=""M"" type=""text"" name=""w_no_secretario"" class=""STI"" SIZE=""40"" MAXLENGTH=""40"" VALUE=""" & w_no_secretario & """ ONMOUSEOVER=""popup('Informe o telefone do contato técnico da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b>Co<u>n</u>tato internet:</b><br><input " & w_Disabled & " accesskey=""N"" type=""text"" name=""w_no_contato_internet"" class=""STI"" SIZE=""35"" MAXLENGTH=""35"" VALUE=""" & w_no_contato_internet & """ ONMOUSEOVER=""popup('Informe o nome do contato na internet da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "          <td valign=""top""><font size=""1""><b><u>T</u>elefone contato internet:</b><br><input " & w_Disabled & " accesskey=""Y"" type=""text"" name=""w_nr_fone_internet"" class=""STI"" SIZE=""13"" MAXLENGTH=""20"" VALUE=""" & w_nr_fone_internet & """ ONMOUSEOVER=""popup('Informe o telefone do contato na internet da escola.','white')""; ONMOUSEOUT=""kill()""></td>"  
  ShowHTML "          <td valign=""top""><font size=""1""><b>Fa<u>x</u> contato internet:</b><br><input " & w_Disabled & " accesskey=""X"" type=""text"" name=""w_nr_fax_internet"" class=""STI"" SIZE=""13"" MAXLENGTH=""20"" VALUE=""" & w_nr_fax_internet & """ ONMOUSEOVER=""popup('Informe o fax do contato na internet da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      </table>"  
  ShowHTML "      <tr valign=""top""><td colspan=""2""><font size=""1"">"
  ShowHTML "        <table width=""100%"" border=""0"">"  
  ShowHTML "          <td valign=""top""><font size=""1""><b><u>e</u>-Mail contato internet:</b><br><input " & w_Disabled & " accesskey=""E"" type=""text"" name=""w_ds_email_internet"" class=""STI"" SIZE=""35"" MAXLENGTH=""60"" VALUE=""" & w_ds_email_internet & """ ONMOUSEOVER=""popup('Informe o e-Mail do contato na internet da escola.','white')""; ONMOUSEOUT=""kill()""></td>"  

  SQL = "SELECT a.sq_modelo, a.ds_modelo FROM escModelo a ORDER BY a.sq_modelo desc" & VbCrLf
  ConectaBD SQL  
  ShowHTML "              <td valign=""top""><font size=""1""><b>Modelo de site:</b><br><SELECT class=""STI"" NAME=""w_sq_modelo"">"
  If RS.RecordCount > 1 Then ShowHTML "          <option value="""">---" End If
  While Not RS.EOF
     If cDbl(nvl(RS("sq_modelo"),0)) = cDbl(nvl(w_sq_modelo,0)) Then
        ShowHTML "          <option value=""" & RS("sq_modelo") & """ SELECTED>" & RS("ds_modelo")
     Else
        ShowHTML "          <option value=""" & RS("sq_modelo") & """>" & RS("ds_modelo")
     End If
     RS.MoveNext
  Wend
  ShowHTML "              </select>"
  DesconectaBD  
  ShowHTML "              <td valign=""top""><font size=""1""><b><u>D</u>iretório:</b><br><input " & w_Disabled & " accesskey=""D"" type=""text"" name=""w_ds_diretorio"" class=""STI"" SIZE=""40"" MAXLENGTH=""60"" VALUE=""" & w_ds_diretorio & """ ONMOUSEOVER=""popup('Informe o diretório físico da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "              <td valign=""top""><font size=""1""><b>SIS<u>C</u>OL:</b><br><input " & w_Disabled & " accesskey=""C"" type=""text"" name=""w_sq_siscol"" class=""STI"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & w_sq_siscol & """ ONMOUSEOVER=""popup('Informe o número siscol da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      </table>"  
  ShowHTML "      <tr><td valign=""top""><font size=""1""><b><u>M</u>ensagem:</b><br><input " & w_Disabled & " accesskey=""M"" type=""text"" name=""w_ds_mensagem"" class=""STI"" SIZE=""80"" MAXLENGTH=""80"" VALUE=""" & Nvl(w_ds_mensagem,"Brasília - Patrimônio Cultural da Humanidade") & """ ONMOUSEOVER=""popup('Informe o mensagem rolante da tela da escola.','white')""; ONMOUSEOUT=""kill()""></td>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><font  size=""1""><b><U>I</U>nstitucional:</b><br><TEXTAREA ACCESSKEY=""I"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_institucional"" ROWS=4 COLS=76>"  & Nvl(w_ds_institucional,"Desenvolvemos essa solução a fim de prestar serviços eficientes e gratuitos à comunidade, criando um espaço de intercâmbio com a escola para, juntos, trabalharmos pela melhora do ensino público.") &"</TEXTAREA ></td>"
  ShowHTML "      <tr><td valign=""top"" colspan=""2""><font  size=""1""><b><U>T</U>exto de abertura:</b><br><TEXTAREA ACCESSKEY=""X"" " & w_Disabled & " class=""STI"" type=""text"" name=""w_ds_texto_abertura"" ROWS=4 COLS=76>"  & Nvl(w_ds_texto_abertura,"O ensino moderno oferece amplo apoio tecnológico ao estudante. A Internet modificou, substancialmente, o paradigma existente na área educacional. Modernizamos nossas atividades a fim de garantir a nossos alunos, pais e responsáveis qualidade, eficiência e rapidez na prestação de nossos serviços.") & "</TEXTAREA ></td>"
  
  SQL = "SELECT a.*, c.sq_codigo_espec " & _ 
        "  from escEspecialidade AS a " & _ 
        "       LEFT JOIN escEspecialidade_cliente AS c ON (a.sq_especialidade = c.sq_codigo_espec and " & _
        "                                     c.sq_cliente  = " & Nvl(w_chave,0) & ") " & _
        "       LEFT JOIN escCliente AS d ON (c.sq_cliente = d.sq_cliente  and " & _
        "                                     d.sq_cliente  = " & Nvl(w_chave,0) & ") " & _
        " where a.tp_especialidade <> 'M' or a.tp_especialidade = 'J'" & VbCrLf & _
        "ORDER BY a.nr_ordem, a.ds_especialidade "
  ConectaBD SQL
   
  If Not RS.EOF Then
     wAtual = ""
     
     ShowHTML "          <tr><TD colspan=2><table border=""0"" align=""left"" cellpadding=0 cellspacing=0>"
     Do While Not RS.EOF
        If wAtual = "" or wAtual <> RS("tp_especialidade") Then
           wAtual = RS("tp_especialidade")
           If wAtual = "M" Then
              ShowHTML "            <TR><TD colspan=2><font size=""1"" CLASS=""BTM""><b>Etapas/Modalidades de ensino:</b>"
           ElseIf wAtual = "R" Then
              ShowHTML "            <TR><TD colspan=2><font size=""1"" CLASS=""BTM""><b>Em Regime de Intercomplementaridade:</b>"
           Else
              ShowHTML "            <TR><TD colspan=2><font size=""1"" CLASS=""BTM""><b>Outras:</b>"
           End If
        End If
           
        If Nvl(RS("sq_codigo_espec"),"") > "" Then
           ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""w_sq_codigo_espec"" value=""" & RS("sq_especialidade") & """ checked><td><font size=1>" & RS("ds_especialidade")
        Else
           ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""w_sq_codigo_espec"" value=""" & RS("sq_especialidade") & """><td><font size=1>" & RS("ds_especialidade")
        End If
        RS.MoveNext
                
     Loop
     DesconectaBD
  End If  
  ShowHTML "      </table>"
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
  ShowHTML "            <input class=""STB"" type=""button"" onClick=""location.href='" & w_Pagina & conRelEscolas & replace(replace(MontaFiltro("GET"),CL,""),"CL=","") & "&w_ea=L&pesquisa=Sim';"" name=""Botao"" value=""Cancelar"">"
  ShowHTML "          </td>"
  ShowHTML "      </tr>"
  ShowHTML "    </table>"
  ShowHTML "    </TD>"
  ShowHTML "</tr>"
  ShowHTML "</FORM>"
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape

  Set w_chave                 = Nothing 
  Set w_regiao                = Nothing 
  Set w_local                 = Nothing 
  Set w_sq_cliente_pai        = Nothing    
  Set w_sq_tipo_cliente       = Nothing 
  Set w_ds_cliente            = Nothing 
  Set w_ds_apelido            = Nothing 
  Set w_no_municipio          = Nothing 
  Set w_sg_uf                 = Nothing 
  Set w_ds_username           = Nothing 
  Set w_ln_internet           = Nothing 
  Set w_ds_email              = Nothing 
  Set w_ds_logradouo          = Nothing 
  Set w_no_bairro             = Nothing 
  Set w_nr_cep                = Nothing 
  Set w_no_contato            = Nothing 
  Set w_email_contato         = Nothing 
  Set w_nr_fone_contato       = Nothing 
  Set w_nr_fax_contato        = Nothing 
  Set w_no_diretor            = Nothing  
  Set w_no_secretario         = Nothing 
  Set w_sq_modelo             = Nothing 
  Set w_no_contato_internet   = Nothing  
  Set w_nr_fone_internet      = Nothing 
  Set w_nr_fax_internet       = Nothing 
  Set w_ds_email_internet     = Nothing 
  Set w_ds_diretorio          = Nothing 
  Set w_sq_siscol             = Nothing 
  Set w_ds_mensagem           = Nothing 
  Set w_ds_institucional      = Nothing 
  Set w_ds_texto_abertura     = Nothing 
  Set w_sq_codigo_espec       = Nothing 
  Set w_username_atual        = Nothing

End Sub
REM =========================================================================
REM Fim do cadastro de escolas
REM -------------------------------------------------------------------------

REM =========================================================================
REM Monta a tela de senhas especiais
REM -------------------------------------------------------------------------
Public Sub ShowSenhaEspecial

  Dim RS1
  Dim RS2

  Dim sql, sql2, wCont, sql1, wAtual, wIN, w_especialidade
  
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  Set RS2 = Server.CreateObject("ADODB.RecordSet")

  Cabecalho
  BodyOpen "onLoad='document.focus()';"
  ShowHTML "<B><FONT COLOR=""#000000"">Senhas especiais - Listagem</FONT></B>"
  ShowHTML "<div align=center><center>"
  sql = "SELECT a.ds_username, a.sq_cliente, a.ds_cliente,  " & VbCrLf & _
        "       a.ds_senha_acesso, a.no_municipio, a.sg_uf, a.dt_alteracao " & VbCrLf & _ 
        "  from escCliente a " & VbCrLf & _ 
        " where a.publica = 'S' and a.sq_cliente_pai is null or a.sq_cliente_pai = 0 and a.ds_username <> 'SBPI'" & VbCrLf & _
        "ORDER BY a.sq_cliente_pai, a.ds_cliente " & VbCrLf
  ConectaBD SQL

  ShowHTML "<TR><TD valign=""top""><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
  If Not RS.EOF Then

     RS.PageSize = RS.RecordCount + 1
     rs.AbsolutePage = 1
      

     ShowHTML "<tr><td><td align=""right""><b><font face=Verdana size=1>Registros encontrados: " & RS.RecordCount & "</font></b>"
     ShowHTML "<tr><td><td>"
     ShowHTML "<table border=""1"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
     ShowHTML "<tr align=""center"" valign=""top"">"
     ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Escola</b></td>"
     ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Username</b></td>"
     ShowHTML "    <td><font face=""Verdana"" size=""1""><b>Senha</b></td>"
  
     w_cor = "#FDFDFD"
     While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(Nvl(Request("P3"),RS.AbsolutePage))
       If w_cor = "#EFEFEF" or w_cor = "" Then w_cor = "#FDFDFD" Else w_cor = "#EFEFEF" End If
       ShowHTML "<tr valign=""top"" bgcolor=""" & w_cor & """>"
       ShowHTML "    <td><font face=""Verdana"" size=""1"">" & RS("DS_CLIENTE") & "</font></td>"
       ShowHTML "    <td><font face=""Verdana"" size=""1"">" & RS("DS_USERNAME") & "</font></td>"
       ShowHTML "    <td align=""center""><font face=""Verdana"" size=""1"">" & RS("DS_SENHA_ACESSO") & "</font></td>"
       RS.MoveNext
     Wend
    
     ShowHTML "</table>"
     ShowHTML "<tr><td><td colspan=""4"" align=""center""><hr>"
  Else

     ShowHTML "<TR><TD><TD colspan=""3""><p align=""justify""><img src=""img/ico_educacao.gif"" width=""16"" height=""16"" border=""0"" align=""center"">&nbsp;<font size=""2""><b>Nenhuma ocorrência encontrada para as opções acima."

  End If

  ShowHTML "</TABLE>"
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Pesquisa
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  If Nvl(w_IN,0) = 1 then
    ShowFrames
  ElseIf CL > "" or _
         w_ew = conWhatSGE or w_R = conWhatSGE or _
         Instr("COMPONENTE,VERSAO", w_ew) > 0 or Instr("COMPONENTE,VERSAO", w_R) > 0 _
  Then
     
    Select Case uCase(w_EW)
      Case "SHOWMENU"                   ShowMenu
      Case "ADM"                        ShowAdmin
      Case "BASE"                       GetCalendarioBase
      Case conWhatAdmin                 Administrativo
      Case conWhatCliente               GetCliente
      Case conWhatDadosAdicionais       GetDadosAdicionais
      Case conWhatSite                  GetSite
      Case conWhatNotCliente            GetNoticiaCliente
      Case conWhatCalendario            GetCalendario
      Case conWhatEspecialidadeCliente  GetEspecialidadeCliente
      Case conWhatDocumento             GetDocumento
      Case conWhatCalendario            GetCalendario
      Case conWhatMensagem              GetMensagem
      Case "SENHAESP"                   ShowSenhaEspecial
      Case "NEWSLETTER"                 GetNewsletter
      Case conWhatEspecialidade         GetEspecialidade
      Case conRelEscolas                ShowEscolas
      Case "ESCPART"                    ShowEscolaParticular
      Case "ESCPARTHOMOLOG"             ShowEscolaParticularHomolog
      Case "GETESCOLAS"                 GetEscolas
      Case "CADASTROESCOLA"             CadastroEscolas
      Case conWhatSGE                   GetSistema
      Case "COMPONENTE"                 GetComponente
      Case "VERSAO"                     GetVersao
      Case "TIPOCLIENTE"                GetTipoCliente
      Case "VERIFARQ"                   GetVerifArquivo
      Case "REDEPART"                   GetRedeParticular
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
REM Fim do Controle.asp
REM =========================================================================
%>