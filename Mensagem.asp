<%@ Language=VBScript %>
<!--#INCLUDE FILE="Constants_ADO.inc"-->
<!-- #INCLUDE FILE="Constants.inc" -->
<!-- #INCLUDE FILE="jScript.asp" -->
<!-- #INCLUDE FILE="Funcoes.asp" -->
<!--#INCLUDE FILE="FuncoesGR.asp"-->
<%
Response.Expires = 0
REM =========================================================================
REM  /Mensagem.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: Gerencia o módulo de envio eletrônico de mensagens
REM Mail     : alex@sbpi.com.br
REM Criacao  : 07/11/2001 9:02
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM -------------------------------------------------------------------------
REM
REM Parâmetros recebidos:
REM    R (referência) = usado na rotina de gravação, com conteúdo igual ao parâmetro T
REM    O (operação)   = I   : Inclusão
REM                   = A   : Alteração
REM                   = C   : Cancelamento
REM                   = E   : Exclusão
REM                   = V   : Envio
REM                   = L   : Listagem
REM                   = P   : Pesquisa
REM                   = D   : Detalhes
REM                   = N   : Nova solicitação de envio

' Verifica se o usuário está autenticado
'If Session("LogOn") <> "Sim" Then
'  ScriptOpen "JavaScript"
'  ShowHTML " alert('Você precisa autenticar-se para utilizar o sistema!'); "
'  ShowHTML " top.location.href='Default.asp'; "
'  ScriptClose
'End If

' Declaração de variáveis
Dim OraDatabase, RS, SQL, dbms
Dim CL, TP, SG
Dim R, O, w_Cont, w_Pagina, w_Disabled, w_TP
Public Upload,File
Private Par

AbreSessao

' Carrega variáveis locais com os dados dos parâmetros recebidos
Par          = ucase(Request("Par"))
If InStr(uCase(Request.ServerVariables("http_content_type")),"MULTIPART/FORM-DATA") > 0 Then
   ' Cria o objeto de upload
   Set ul       = Nothing
   Set ul       = Server.CreateObject("Dundas.Upload.2")
   ul.SaveToMemory
   CL           = ul.Form("CL")
   TP           = ul.Form("TP")
   SG           = ucase(ul.Form("SG"))
   R            = uCase(ul.Form("R"))
   O            = uCase(ul.Form("O"))
Else
   CL           = Request("CL")
   TP           = Request("TP")
   SG           = ucase(Request("SG"))
   R            = uCase(Request("R"))
   O            = uCase(Request("O"))
End If
w_Pagina     = "Mensagem.asp?par="
w_Disabled   = "ENABLED"

If O = "" Then O = "L" End If

Select Case O
  Case "I" 
     w_TP = TP & " - Inclusão"
  Case "A" 
     w_TP = TP & " - Alteração"
  Case "E" 
     w_TP = TP & " - Exclusão"
  Case "V" 
     w_TP = TP & " - Envio"
  Case "P" 
     w_TP = TP & " - Filtragem"
  Case Else
     w_TP = TP & " - Listagem"
End Select
Main

FechaSessao

Set OraDatabase = Nothing
Set RS          = Nothing
Set SQL         = Nothing
Set Par         = Nothing
Set TP          = Nothing
Set SG          = Nothing
Set R           = Nothing
Set O           = Nothing
Set w_Cont      = Nothing
Set w_Pagina    = Nothing
Set w_Disabled  = Nothing
Set w_TP        = Nothing
Set Upload      = Nothing
Set File        = Nothing

REM =========================================================================
REM Rotina de inclusão de novas mensagens
REM -------------------------------------------------------------------------
Sub Inicial

  Dim RS1, RS2
  
  Dim w_cor
  Dim w_sq_cliente_mail
  Dim w_assunto
  Dim w_envia_lista
  Dim w_texto
  Dim w_Cont

  If O = "L" Then
     SQL = "select a.sq_cliente_mail, a.enviada, a.assunto, substring(a.texto,1,200) texto, " & VbCrLf & _
           "       convert(varchar, a.data_inclusao, 103) data_inclusao, a.data_inclusao dt_inclusao," & VbCrLf & _
           "       case when a.data_envio is null then '---' else convert(varchar, a.data_envio, 103) end data_envio, " & VbCrLf & _
           "       a.envia_lista, case a.envia_lista when 'S' then 'Sim' else 'Não' end nm_envia_lista, " & VbCrLf & _
           "       (select count(*) Itens from escMail_Destinatario where sq_cliente_mail = a.sq_cliente_mail) Itens " & VbCrLf & _
           "from escCliente_Mail        a " & VbCrLf & _
           "     inner join escCliente  b on (a.sq_cliente = b.sq_cliente) " & VbCrLf
     If Session("username") = "IMPRENSA" Then
        SQL = SQL & "where a.sq_cliente  = 0 " & VbCrLf
      Else
        SQL = SQL & "where a.sq_cliente  = " & replace(CL,"sq_cliente=","") & " " & VbCrLf
      End If
     
     SQL = SQL & "order by a.dt_inclusao desc"
     ConectaBD SQL
  ElseIf InStr("AEV",O) > 0 Then
     w_sq_cliente_mail = Request("w_sq_cliente_mail")
     SQL = "select a.assunto, a.texto, a.envia_lista, a.data_envio " & VbCrLf & _
           "from escCliente_Mail a " & VbCrLf & _
           "where a.sq_cliente_mail =  " & w_sq_cliente_mail 
     ConectaBD SQL
     w_assunto     = RS("assunto")
     w_envia_lista = RS("envia_lista")
     w_texto       = RS("texto")
     w_data_envio  = RS("data_envio")
     DesconectaBD
     
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  If InStr("IAEVP",O) > 0 Then
     ScriptOpen "JavaScript"
     CheckBranco
     FormataData
     FormataDataHora
     ValidateOpen "Validacao"
     If InStr("IA",O) > 0 Then
        Validate "w_assunto", "Assunto", "1", "1", "1", "200", "1", "1"
        Validate "w_texto", "Mensagem", "1", "1", "1", "4000", "1", "1"
     ElseIf O = "E" Then
        ShowHTML "  if (confirm('Confirma a exclusão deste e-mail?')) "
        ShowHTML "     { return (true); }; "
        ShowHTML "     { return (false); }; "
     ElseIf O = "V" Then
        ShowHTML "  if (confirm('Confirma o envio deste e-mail?')) { "
        ShowHTML "     theForm.Botao[0].disabled=true;"
        ShowHTML "     theForm.Botao[1].disabled=true;"
        ShowHTML "     return (true); "
        ShowHTML "  }; "
        ShowHTML "  else return (false); "
     'ElseIf O="P" Then
     '   Validate "p_cor", "Cor", "1", "", "3", "10", "1", "1"
     '   Validate "p_codigo_siape", "Código SIAPE", "1", "", "2", "2", "", "1"
     End If
     ShowHTML "  theForm.Botao[0].disabled=true;"
     ShowHTML "  theForm.Botao[1].disabled=true;"
     ValidateClose
     ScriptClose
  End If
  ShowHTML "</HEAD>"
  If InStr("IA",O) > 0 Then
     BodyOpen "onLoad='document.Form.w_assunto.focus()';"
  ElseIf InStr("EV",O) > 0 Then
     BodyOpen "onLoad='document.focus()';"
  ElseIf O = "P" Then
     BodyOpen "onLoad='document.Form.p_data_inicio_inclusao.focus()';"
  Else
     BodyOpen "onLoad=document.focus();"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">"
  If O = "L" Then
    ShowHTML "<tr><td><font size=""2""><a accesskey=""I"" class=""SS"" href=""" & w_Pagina & par & "&R=" & w_Pagina & par & "&O=I&CL=" & CL & "&TP=" & TP & "&SG=" & SG & """><u>I</u>ncluir</a>&nbsp;"
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=3>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & conTrBgColor & """ align=""center"">"
    ShowHTML "          <td colspan=2><font size=""1""><b>Data</font></td>"
    ShowHTML "          <td rowspan=2><font size=""1""><b>Nº</font></td>"
    ShowHTML "          <td rowspan=2><font size=""1""><b>Assunto</font></td>"
    ShowHTML "          <td rowspan=2><font size=""1""><b>Envia p/lista</font></td>"
    ShowHTML "          <td rowspan=2><font size=""1""><b>Operações</font></td>"
    ShowHTML "        </tr>"
    ShowHTML "        <tr bgcolor=""" & conTrBgColor & """ align=""center"">"
    ShowHTML "          <td><font size=""1""><b>Inclusão</font></td>"
    ShowHTML "          <td><font size=""1""><b>Envio</font></td>"
    ShowHTML "        </tr>"
    If RS.EOF Then
        ShowHTML "      <tr bgcolor=""" & conTrBgColor & """><td colspan=6 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        If RS("ENVIADA") = "S" Then w_Cor = " color=""#0000FF""" Else w_Cor = "" End If
        If RS("texto") > "" Then
           ShowHTML "      <tr bgcolor=""" & conTrBgColor & """ ONMOUSEOVER=""popup('" & Replace(RS("texto"),CHR(13)&CHR(10),"<BR>") & "','white')""; ONMOUSEOUT=""kill()"" valign=""top"">"
        Else
           ShowHTML "      <tr bgcolor=""" & conTrBgColor & """ valign=""top"">"
        End If
        ShowHTML "        <td align=""center""><font size=""1""" & w_Cor & ">" & RS("data_inclusao") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1""" & w_Cor & ">" & RS("data_envio") & "</td>"
        ShowHTML "        <td align=""right""><font size=""1""" & w_Cor & ">" & RS("sq_cliente_mail") & "</td>"
        ShowHTML "        <td><font size=""1""" & w_Cor & ">" & RS("assunto") & "</td>"
        ShowHTML "        <td align=""center""><font size=""1""" & w_Cor & ">" & RS("nm_envia_lista") & "</td>"
        ShowHTML "        <td align=""top"" nowrap><font size=""1"">"
        If RS("data_envio") <> "---" Then
           ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & par & "&R=" & w_Pagina & par & "&O=V&w_sq_cliente_mail=" & RS("sq_cliente_mail") & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG & """>Exibir</A>&nbsp"
        Else
           ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & par & "&R=" & w_Pagina & par & "&O=A&w_sq_cliente_mail=" & RS("sq_cliente_mail") & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG  & """>Alterar</A>&nbsp"
           ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & par & "&R=" & w_Pagina & par & "&O=E&w_sq_cliente_mail=" & RS("sq_cliente_mail") & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG  & """>Excluir</A>&nbsp"
           ShowHTML "          <u class=""HL"" style=""cursor:hand;"" onclick=""javascript:window.open('" & w_Pagina & "Item&R=" & w_Pagina & "Item&O=L&w_sq_cliente_mail=" & RS("sq_cliente_mail") & "&CL=" & CL & "&TP=" & TP & " - Destinatários&SG=Destino','Bens','toolbars=no,width=720,height=530,top=30,left=50,resizable=yes,scrollbars=yes')"">Destinos</u>&nbsp"
           ShowHTML "          <u class=""HL"" style=""cursor:hand;"" onclick=""javascript:window.open('" & w_Pagina & "Item&R=" & w_Pagina & "Item&O=L&w_sq_cliente_mail=" & RS("sq_cliente_mail") & "&CL=" & CL & "&TP=" & TP & " - Anexos&SG=Anexo','Bens','toolbars=no,width=720,height=430,top=30,left=50,resizable=yes,scrollbars=yes')"">Anexos</u>&nbsp"
           If RS("Itens") = 0 and RS("envia_lista") = "N" Then
              ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & par & "&R=" & w_Pagina & par & "&O=V&w_sq_cliente_mail=" & RS("sq_cliente_mail") & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG & """ onClick=""alert('Você não pode enviar esta solicitação sem informar pelo menos um destinatário!'); return false;"">Enviar</A>&nbsp"
           Else
              ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & par & "&R=" & w_Pagina & par & "&O=V&w_sq_cliente_mail=" & RS("sq_cliente_mail") & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG & """>Enviar</A>&nbsp"
           End If
        End If
        ShowHTML "        </td>"
        ShowHTML "      </tr>"
        RS.MoveNext
      wend
    End If
    ShowHTML "      </center>"
    ShowHTML "    </table>"
    ShowHTML "  </td>"
    ShowHTML "</tr>"
    ShowHTML "<tr><td><font size=1><b>* A cor azul na mensagem indica que ela já foi enviada.</b></td></tr>"
    DesConectaBD     
  ElseIf Instr("IAEV",O) > 0 Then
    If InStr("EV",O) Then
       w_Disabled = "DISABLED"
    End If
    ShowHTML "<FORM action=""" & w_Pagina & "Grava"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"">"
    ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
    ShowHTML "<INPUT type=""hidden"" name=""TP"" value=""" & TP & """>"
    ShowHTML "<INPUT type=""hidden"" name=""SG"" value=""" & SG & """>"
    ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & R & """>"
    ShowHTML "<INPUT type=""hidden"" name=""O"" value=""" & O &""">"
    ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente_mail"" value=""" & w_sq_cliente_mail &""">"

    ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
    ShowHTML "    <table width=""70%"" border=""0"">"
    If O <> "I" Then
       ShowHTML "      <tr><td><font size=""1""><b>Remetente:<br><b>" & Session("NOME") & " (" & Session("USERNAME") & ")</b></td>"
    Else
       ShowHTML "      <tr><td><font size=""1""><b>Remetente:<br><b>" & Session("Nome") & " (" & Session("Username") & ")</b></td>"
    End If
    ShowHTML "      <tr><td><font size=""1""><b>A<U>s</U>sunto:<br><INPUT ACCESSKEY=""S"" " & w_Disabled & " class=""BTM"" type=""text"" name=""w_assunto"" size=""60"" maxlength=""200"" value=""" & w_assunto & """></td>"
    ShowHTML "      <tr><td><font size=""1""><b><U>M</U>ensagem:<br><TEXTAREA ACCESSKEY=""M"" " & w_Disabled & " class=""BTM"" name=""w_texto"" rows=""5"" cols=""90"">" & w_texto & "</TEXTAREA></td>"
    ShowHTML "      <tr>"
    MontaRadioNS "<b>Envia para lista de distribuição de informativos?</b>", w_envia_lista, "w_envia_lista"
    ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000""></TD></TR>"
    ShowHTML "      <tr><td align=""center"" colspan=""3"">"
    If O = "E" Then
       ShowHTML "            <input class=""BTM"" type=""submit"" name=""Botao"" value=""Excluir"">"
       ShowHTML "            <input class=""BTM"" type=""button"" onClick=""location.href='" & w_Pagina & par & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "';"" name=""Botao"" value=""Cancelar"">"
    ElseIf O = "V" Then
       If NVL(w_data_envio,"nulo") <> "nulo" Then
          ShowHTML "            <input class=""BTM"" type=""button"" onClick=""location.href='" & w_Pagina & par & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "';"" name=""Botao"" value=""Voltar"">"
       Else
          ShowHTML "            <input class=""BTM"" type=""submit"" name=""Botao"" value=""Enviar"">"
          ShowHTML "            <input class=""BTM"" type=""button"" onClick=""location.href='" & w_Pagina & par & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "';"" name=""Botao"" value=""Cancelar"">"
       End If
    Else
       ShowHTML "            <input class=""BTM"" type=""submit"" name=""Botao"" value=""Gravar"">"
       ShowHTML "            <input class=""BTM"" type=""button"" onClick=""location.href='" & w_Pagina & par & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "';"" name=""Botao"" value=""Cancelar"">"
    End If
    ShowHTML "          </td>"
    ShowHTML "      </tr>"
    ShowHTML "    </table>"
    ShowHTML "    </TD>"
    ShowHTML "</tr>"
    ShowHTML "</FORM>"
    ' Exibe destinatários e anexos da mensagem
    If O <> "I" Then
       ShowHTML "<tr><td align=""center"" colspan=3 bgcolor=""" & conTableBgColor & """><hr>"
       SQL = "select sq_mail_anexo, arquivo_origem, arquivo_destino " & VbCrLf & _
             "  from escCliente_Mail          a " & VbCrLf & _
             "       inner join escMail_Anexo b on (a.sq_cliente_mail        = b.sq_cliente_mail) " & VbCrLf & _
             " where a.sq_cliente_mail = " & w_sq_cliente_mail & " " & VbCrLf & _
             "order by 2 "
       Set RS1 = dbms.Execute(SQL)
       ShowHTML "<tr><td align=""center"" colspan=3 bgcolor=""" & conTableBgColor & """><div align=""left""><font size=1><b>Anexos:</div>"
       ShowHTML "    <TABLE WIDTH=""100%"" align=""center"" BORDER=0 CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
       If NOT RS1.EOF Then
          While Not RS1.EOF
             ShowHTML "        <tr valign=""top""><td><img src=""img/arrow-closed.gif"" border=0 align=""center""><font face=""Arial"" size=""1""><a href=""sedf/sedf/" & RS1("arquivo_destino") & """ class=""SS"" target=""_blank"">" & RS1("arquivo_origem") & "</a></font></td>"
             RS1.MoveNext
          Wend
       Else
          ShowHTML "        <tr valign=""top""><td><img src=""img/arrow-closed.gif"" border=0 align=""center""><font face=""Arial"" size=""1"">Nenhum anexo foi inserido.</font></td>"
       End If
       ShowHTML "    </table>"

       SQL = "select c.sq_cliente, c.ds_cliente nome, a.envia_lista " & VbCrLf & _
             "  from escCliente_Mail                   a " & VbCrLf & _
             "       left outer join escMail_Destinatario b on (a.sq_cliente_mail  = b.sq_cliente_mail) " & VbCrLf & _
             "         inner join escCliente           c on (b.destinatario     = c.sq_cliente) " & VbCrLf & _
             " where a.sq_cliente_mail = " & w_sq_cliente_mail & " " & VbCrLf & _
             "order by 2 "
       Set RS1 = dbms.Execute(SQL)
       w_Cont = 0
       ShowHTML "<tr><td align=""center"" colspan=3 bgcolor=""" & conTableBgColor & """>&nbsp;"
       ShowHTML "<tr><td align=""center"" colspan=3 bgcolor=""" & conTableBgColor & """><div align=""left""><font size=1><b>Destinatários:</div>"
       If NOT RS1.EOF Then
          ShowHTML "    <TABLE WIDTH=""100%"" align=""center"" BORDER=0 CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
          If RS1("envia_lista") = "S" Then
             w_Cont = w_Cont + 1
             If w_Cont = 1 or (w_Cont Mod 2 = 1 ) Then
                ShowHTML "        <tr valign=""top"">"
             End If
             ShowHTML "          <td><img src=""img/arrow-closed.gif"" border=0 align=""center""><font face=""Arial"" size=""1"">Lista de distribuição de informativos</font></td>"
          End If
          While Not RS1.EOF
             w_Cont = w_Cont + 1
             If w_Cont = 1 or (w_Cont Mod 2 = 1 ) Then
                ShowHTML "        <tr valign=""top"">"
             End If
             ShowHTML "          <td><img src=""img/arrow-closed.gif"" border=0 align=""center""><font face=""Arial"" size=""1"">" & RS1("Nome") & "</font></td>"
             RS1.MoveNext
          Wend
       Else
          SQL = "select a.envia_lista " & VbCrLf & _
                "  from escCliente_Mail                   a " & VbCrLf & _
                " where a.sq_cliente_mail = " & w_sq_cliente_mail & " " & VbCrLf
          Set RS1 = dbms.Execute(SQL)
          ShowHTML "        <tr valign=""top"">"
          If RS1("envia_lista") = "S" Then
             ShowHTML "          <td><img src=""img/arrow-closed.gif"" border=0 align=""center""><font face=""Arial"" size=""1"">Lista de distribuição de informativos</font></td>"
          Else
             ShowHTML "          <td><img src=""img/arrow-closed.gif"" border=0 align=""center""><font face=""Arial"" size=""1"">Nenhum destinatário foi informado.</font></td>"
          End If
       End If
       ShowHTML "    </table>"
    End If
  Else
    ScriptOpen "JavaScript"
    ShowHTML " alert('Opção não disponível');"
    ShowHTML " history.back(1);"
    ScriptClose
  End If
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape

  Set w_Cont                = Nothing
  Set RS1                   = Nothing
  Set RS2                   = Nothing
  Set w_cor                 = Nothing
  Set w_sq_cliente_mail     = Nothing
  Set w_assunto             = Nothing
  Set w_envia_lista         = Nothing
  Set w_texto               = Nothing
End Sub
REM =========================================================================
REM Fim da tela de inclusão de solicitações
REM -------------------------------------------------------------------------

REM =========================================================================
REM Trata a manipulação dos itens
REM -------------------------------------------------------------------------
Sub Item
 
  Dim w_sq_cliente_mail
  Dim w_sq_cliente
  Dim w_sq_unidade
  Dim p_regional
  Dim p_tipo_cliente
  Dim p_modalidade
  Dim RS1
  Dim SQL1
  Dim SQL2
  Dim SQL3
   
  w_sq_cliente_mail = Request("w_sq_cliente_mail")
  p_regional        = Request("p_regional")
  p_tipo_cliente    = Request("p_tipo_cliente")
  p_modalidade      = Request("p_modalidade")
  
  If O = "" Then O="L" end if
  
  If InStr("L",O) Then
     If SG = "DESTINO" Then ' Se for destinatário
        SQL = "select c.sq_cliente, c.ds_cliente nome " & VbCrLf & _
              "  from escCliente_Mail                   a " & VbCrLf & _
              "       inner   join escMail_Destinatario b on (a.sq_cliente_mail = b.sq_cliente_mail) " & VbCrLf & _
              "         inner join escCliente           c on (b.destinatario    = c.sq_cliente)" & VbCrLf & _
              " where a.sq_cliente_mail = " & w_sq_cliente_mail & " " & VbCrLf & _
              "order by 2 "
     ElseIf SG = "ANEXO" Then ' Se for arquivo atachado
        SQL = "select sq_mail_anexo, arquivo_origem, arquivo_destino " & VbCrLf & _
              "  from escCliente_mail          a " & VbCrLf & _
              "       inner join escMail_Anexo b on (a.sq_cliente_mail = b.sq_cliente_mail) " & VbCrLf & _
              " where a.sq_cliente_mail = " & w_sq_cliente_mail & " " & VbCrLf & _
              "order by 2 "
     End If
     ConectaBD SQL
  End If
  
  Cabecalho
  ShowHTML "<HEAD>"
  If InStr("IAE",O) > 0 Then
     ScriptOpen "JavaScript"
     If InStr("IA",O) > 0 Then
        If SG = "DESTINO" and (p_regional & p_modalidade & p_tipo_cliente) > "" Then
           ShowHTML "  function MarcaTodos() {"
           ShowHTML "    if (document.Form1.w_sq_cliente.value==undefined) "
           ShowHTML "       for (i=0; i < document.Form1.w_sq_cliente.length; i++) "
           ShowHTML "         if (!document.Form1.w_sq_cliente[i].disabled) document.Form1.w_sq_cliente[i].checked=true;"
           ShowHTML "    else if (!document.Form1.w_sq_cliente.disabled) document.Form1.w_sq_cliente.checked=true;"
           ShowHTML "  }"
           ShowHTML "  function DesmarcaTodos() {"
           ShowHTML "    if (document.Form1.w_sq_cliente.value==undefined) "
           ShowHTML "       for (i=0; i < document.Form1.w_sq_cliente.length; i++) "
           ShowHTML "         if (!document.Form1.w_sq_cliente[i].disabled) document.Form1.w_sq_cliente[i].checked=false;"
           ShowHTML "    "
           ShowHTML "    else if (!document.Form1.w_sq_cliente.disabled) document.Form1.w_sq_cliente.checked=false;"
           ShowHTML "  }"
        End If
     End If
     ValidateOpen "Validacao"
     If InStr("IA",O) > 0 Then
        If SG = "DESTINO" Then
           ShowHTML "  theForm.Botao[0].disabled=true;"
           ShowHTML "  theForm.Botao[1].disabled=true;"
           ShowHTML "  theForm.Botao[2].disabled=true;"
        ElseIf SG = "ANEXO" Then
           Validate "p_regional", "Arquivo", "1", "1", "4", "100", "1", "1"
           ShowHTML "  theForm.Botao[0].disabled=true;"
           ShowHTML "  theForm.Botao[1].disabled=true;"
        End If
     End If
     ValidateClose
     If SG = "DESTINO" Then
        If Request("pesquisa") > "" Then
           ValidateOpen "Validacao1"
           If InStr("IA",O) > 0 Then
              ShowHTML "  var i; "
              ShowHTML "  var w_erro=true; "
              ShowHTML "  if (theForm.w_sq_cliente.value==undefined) {"
              ShowHTML "     for (i=0; i < theForm.w_sq_cliente.length; i++) {"
              ShowHTML "       if (theForm.w_sq_cliente[i].checked) w_erro=false;"
              ShowHTML "     }"
              ShowHTML "  }"
              ShowHTML "  else {"
              ShowHTML "     if (theForm.w_sq_cliente.checked) w_erro=false;"
              ShowHTML "  }"
              ShowHTML "  if (w_erro) {"
              ShowHTML "    alert('Você deve informar pelo menos um destinatário!'); "
              ShowHTML "    return false;"
              ShowHTML "  }"
           End If
           ShowHTML "  theForm.Botao[0].disabled=true;"
           ShowHTML "  theForm.Botao[1].disabled=true;"
           ShowHTML "  theForm.Botao[2].disabled=true;"
           ValidateClose
        End If
     End If
     ScriptClose
  End If
  ShowHTML "</HEAD>"
  If O = "I" Then
     If Request("pesquisa") > "" Then
        BodyOpen " onLoad=""location.href='#lista'"""
     Else
        BodyOpen "onLoad='document.Form.p_regional.focus()';"
     End If
  Else
     BodyOpen "onLoad=document.focus();"
  End If
  ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
  ShowHTML "<HR>"
  ShowHTML "<div align=center><center>"
  ShowHTML "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
  ShowHTML "<tr><td align=""center"" colspan=3 bgcolor=""#FAEBD7""><table border=1 width=""100%""><tr><td>"
  ShowHTML "    <TABLE WIDTH=""100%"" CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
  ShowHTML "        <tr valign=""top"">"
  ShowHTML "          <td><font size=""1"">Mensagem:<br><b>" & w_sq_cliente_mail & "</font></td>"
  ShowHTML "          <td align=""right""><font size=""1"">Remetente:<br><b>" & Session("USERNAME") & "</font></td>"
  ShowHTML "    </TABLE>"
  ShowHTML "    </TABLE>"
  If O = "L" Then
    ShowHTML "<tr><td><font size=""2"">"    
    ShowHTML "    <a accesskey=""I"" class=""SS"" href=""" & w_Pagina & par & "&R=" & w_Pagina & par & "&O=I&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "&w_sq_cliente_mail=" & w_sq_cliente_mail & """><u>I</u>ncluir</a>&nbsp;"
    ShowHTML "    <font size=""2""><a accesskey=""F"" class=""SS"" href=""javascript:window.close();javascript:opener.location.reload();""><u>F</u>echar</a>&nbsp;"
    ShowHTML "    <td align=""right""><font size=""1""><b>Registros: " & RS.RecordCount
    ShowHTML "<tr><td align=""center"" colspan=6>"
    ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
    ShowHTML "        <tr bgcolor=""" & conTrBgColor & """ align=""center"" valign=""top"">"
    If SG = "DESTINO" Then ' Se for Destinatário
       ShowHTML "          <td><font size=""2""><b>Nome</font></td>"
    ElseIf SG = "ANEXO" Then ' Se for arquivo atachado
       ShowHTML "          <td><font size=""2""><b>Arquivo</font></td>"
    End If
    ShowHTML "          <td><font size=""2""><b>Operações</font></td>"    
    ShowHTML "        </tr>"
    If RS.EOF Then
        ShowHTML "      <tr bgcolor=""" & conTrBgColor & """><td colspan=6 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
    Else
      While Not RS.EOF
        ShowHTML "      <tr bgcolor=""" & conTrBgColor & """ valign=""top"">"
        If SG = "DESTINO" Then ' Se for Destinatário
           ShowHTML "        <td><font size=""1"">" & RS("nome") & "</td>"
           ShowHTML "        <td><font size=""1"">"
           ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_Pagina & par & "&O=E&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "&w_sq_cliente_mail=" & w_sq_cliente_mail & "&w_destinatario=" & RS("sq_cliente") & """ onClick=""return(confirm('Confirma exclusão deste destinatário?'));"">Excluir</A>&nbsp"
        ElseIf SG = "ANEXO" Then ' Se for arquivo atachado
           ShowHTML "        <td><font size=""1"">" & RS("arquivo_origem") & "</td>"
           ShowHTML "        <td><font size=""1"">"
           ShowHTML "          <A class=""HL"" HREF=""" & w_Pagina & "GRAVA&R=" & w_Pagina & par & "&O=E&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "&w_sq_cliente_mail=" & w_sq_cliente_mail & "&w_sq_mail_anexo=" & RS("sq_mail_anexo") & "&w_atual=" & RS("arquivo_destino") & """ onClick=""return(confirm('Confirma exclusão deste arquivo?'));"">Excluir</A>&nbsp"
        End If
        ShowHTML "&nbsp"
        ShowHTML "        </td>"          
        ShowHTML "      </tr>"
        RS.MoveNext
      wend
    End If
    ShowHTML "      </center>"
    ShowHTML "    </table>"
    ShowHTML "  </td>"
    ShowHTML "</tr>"
    DesConectaBD     
  ElseIf Instr("I",O) > 0 Then
    ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
    ShowHTML "    <table width=""95%"" border=""0"">"
    If SG = "DESTINO" Then ' Se for destinatário da mensagem
       ShowHTML "<FORM action=""" & w_Pagina & "Item"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"">"
       ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
       ShowHTML "<INPUT type=""hidden"" name=""TP"" value=""" & TP & """>"
       ShowHTML "<INPUT type=""hidden"" name=""SG"" value=""" & SG & """>"
       ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & R & """>"
       ShowHTML "<INPUT type=""hidden"" name=""O"" value=""" & O &""">"
       ShowHTML "<INPUT type=""hidden"" name=""pesquisa"" value=""S"">"
       ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente_mail"" value=""" & w_sq_cliente_mail & """>"

       ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td><div align=""justify""><font size=1><b><ul>Instruções</b>:<li>Informe os parâmetros desejados para recuperar a lista de possíveis destinatários.<li>Quando a relação for exibida, selecione os destinatários desejados clicando sobre a caixa ao lado do nome.<li>Após informar os parâmetros desejados, clique sobre o botão <i>Aplicar filtro</i>. Clicando sobre o botão <i>Limpar campos</i>, os parâmetros serão apagados.</ul><hr><b>Filtro</b></div>"
       ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
       ShowHTML "    <table width=""100%"" border=""0"">"
       ShowHTML "          <TR><TD valign=""top""><br><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
       ShowHTML "          <tr valign=""top""><td><td valign=""top""><font size=""1""><table border=0 width=""100%"" cellpadding=0 cellspacing=0>"
       SelecaoRegional "<u>S</u>ubordinação:", "S", "Indique a subordinação da escola.", p_regional, null, "p_regional", null, null
       ShowHTML "</table></td></tr>"
       SQL = "SELECT * FROM escTipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" & VbCrLf
       ConectaBD SQL
       ShowHTML "          <tr valign=""top""><td><td valign=""top""><font size=""1""><br><b>Tipo de instituição:</b><br><SELECT CLASS=""STI"" NAME=""p_tipo_cliente"">"
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
       ShowHTML "          </table>"
       DesconectaBD
       wCont = 0
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
                 If cDbl(RS("sq_especialidade")) = cDbl(Nvl(Request("p_modalidade")(i),0)) Then marcado = "S" End If
             Next
                ShowHTML chr(13) & "           <tr><td><input type=""checkbox"" name=""p_modalidade"" value=""" & RS("sq_especialidade") & """><td><font size=1>" & RS("ds_especialidade")
             RS.MoveNext     
             If (wCont Mod 2) = 0 Then 
                wCont = 0
             End If
          Loop
          DesconectaBD
       End If
       ShowHTML "          </tr>"
       ShowHTML "          </table>"
       ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000"">"
       ShowHTML "      <tr><td align=""center"" colspan=""3"">"
       ShowHTML "            <input class=""BTM"" type=""submit"" name=""Botao"" value=""Aplicar filtro"">"
       ShowHTML "            <input class=""BTM"" type=""reset"" name=""Botao"" value=""Limpar campos"">"
       ShowHTML "            <input class=""BTM"" type=""button"" onClick=""location.href='" & w_Pagina & par & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "&w_sq_cliente_mail=" & w_sq_cliente_mail & "&O=L';"" name=""Botao"" value=""Cancelar"">"
       ShowHTML "          </td>"
       ShowHTML "      </tr>"
       ShowHTML "    </table>"
       ShowHTML "    </TD>"
       ShowHTML "</tr>"
       ShowHTML "</form>"
       If Request("pesquisa") > "" Then
          sql = "SELECT DISTINCT d.sq_cliente, d.ds_cliente, d.ds_apelido, " & VbCrLf & _
                "       IsNull(d.ds_email, 'não informado') ds_email, " & VbCrLf & _ 
                "       IsNull(e.ds_email_contato, 'não informado') ds_email_contato, " & VbCrLf & _ 
                "       IsNull(f.ds_email_internet, 'não informado') ds_email_internet " & VbCrLf & _ 
                "  from escCliente                                 d " & VbCrLf & _ 
                "       LEFT OUTER JOIN escEspecialidade_cliente   c ON (c.sq_cliente       = d.sq_cliente) " & VbCrLf & _
                "       LEFT OUTER JOIN escEspecialidade           a ON (a.sq_especialidade = c.sq_codigo_espec) " & VbCrLf & _
                "       LEFT OUTER JOIN escCliente_Dados           e ON (d.sq_cliente       = e.sq_cliente) " & VbCrLf & _
                "       LEFT OUTER JOIN escCliente_Site            f ON (d.sq_cliente       = f.sq_cliente) " & VbCrLf & _
                "       INNER      JOIN escTipo_Cliente            g ON (d.sq_tipo_cliente  = g.sq_tipo_cliente) " & VbCrLf & _
                "       LEFT OUTER JOIN (select destinatario, count(*) existe " & VbCrLf & _
                "                          from escMail_Destinatario " & VbCrLf & _
                "                         where sq_cliente_mail = " & Request("w_sq_cliente_mail") & " " & VbCrLf & _
                "                        group by destinatario " & VbCrLf & _
                "                       )                          h ON (d.sq_cliente       = h.destinatario) " & VbCrLf & _
                " where (CharIndex('@', IsNull(d.ds_email,'---')) > 0) " & VbCrLf & _
                "   and IsNull(h.existe,0) = 0 " & VbCrLf
          If Nvl(Request("p_regional"),-1) = 0  Then
             sql = sql + "    and g.tipo = 2 and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf
          ElseIf Request("p_regional") > "" Then 
             sql = sql + "    and d.sq_cliente_pai = " & Request("p_regional") & VbCrLf
          Else
             sql = sql + "    and g.tipo = 3" & VbCrLf
          End If
          
          If Nvl(Request("p_regional"),-1) <> 0 and Request("p_tipo_cliente") > ""          Then sql = sql + "    and d.sq_tipo_cliente= " & Request("p_tipo_cliente")          & VbCrLf End If

          if Nvl(Request("p_regional"),-1) <> 0 and sql1 > "" then
             sql = sql + "  and a.sq_especialidade in (" + Request("p_modalidade") + ")" & VbCrLf 
          end if
          sql = sql + "ORDER BY d.ds_cliente " & VbCrLf
          ConectaBD SQL
          ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td><font size=2><hr><b><a name=""lista"">Seleção de destinatários</a></b></div>"
          ShowHTML "<FORM action=""" & w_Pagina & "Grava"" method=""POST"" name=""Form1"" onSubmit=""return(Validacao1(this));"">"
          ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
          ShowHTML "<INPUT type=""hidden"" name=""TP"" value=""" & TP & """>"
          ShowHTML "<INPUT type=""hidden"" name=""SG"" value=""" & SG & """>"
          ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & R & """>"
          ShowHTML "<INPUT type=""hidden"" name=""O"" value=""" & O &""">"
          ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente_mail"" value=""" & w_sq_cliente_mail & """>"
          ShowHTML "<INPUT type=""hidden"" name=""p_regional"" value=""" & p_regional & """>"
          ShowHTML "<INPUT type=""hidden"" name=""p_tipo_cliente"" value=""" & p_tipo_cliente & """>"
          ShowHTML "<INPUT type=""hidden"" name=""p_modalidade"" value=""" & p_modalidade & """>"
          ShowHTML "<tr><td align=""right""><font size=""1""><b>Registros: " & RS.RecordCount
          ShowHTML "<tr><td align=""center"" colspan=6>"
          ShowHTML "    <TABLE WIDTH=""100%"" bgcolor=""" & conTableBgColor & """ BORDER=""" & conTableBorder & """ CELLSPACING=""" & conTableCellSpacing & """ CELLPADDING=""" & conTableCellPadding & """ BorderColorDark=""" & conTableBorderColorDark & """ BorderColorLight=""" & conTableBorderColorLight & """>"
          If RS.EOF Then
             ShowHTML "      <tr bgcolor=""" & conTrBgColor & """><td colspan=5 align=""center""><font size=""2""><b>Não foram encontrados registros.</b></td></tr>"
          Else
            ShowHTML "        <tr bgcolor=""" & conTrBgColor & """ align=""center"" valign=""top"">"
            ShowHTML "          <td valign=""center"" rowspan=2 width=""70""NOWRAP><font size=""1""><U ID=""INICIO"" STYLE=""cursor:hand;"" CLASS=""HL"" onClick=""javascript:MarcaTodos();"" TITLE=""Marca todos os itens da relação""><IMG SRC=""img/BookmarkAndPageActivecolor.gif"" BORDER=""1"" width=""15"" height=""15""></U>&nbsp;"
            ShowHTML "                                      <U STYLE=""cursor:hand;"" CLASS=""HL"" onClick=""javascript:DesmarcaTodos();"" TITLE=""Desmarca todos os itens da relação""><IMG SRC=""img/BookmarkAndPageInactive.gif"" BORDER=""1"" width=""15"" height=""15""></U>"
            ShowHTML "          <td valign=""center"" rowspan=2><font size=""1""><b>Nome</font></td>"
            ShowHTML "          <td colspan=3><font size=""1""><b>e-Mails</font></td>"
            ShowHTML "        </tr>"
            ShowHTML "        <tr bgcolor=""" & conTrBgColor & """ align=""center"" valign=""top"">"
            ShowHTML "          <td title=""e-Mail institucional da escola""><font size=""1""><b>Escola</font></td>"
            ShowHTML "          <td title=""e-Mail cadastrado na tela de dados adicionais da escola, usado para contatos técnicos com a escola.""><font size=""1""><b>Técnico</font></td>"
            ShowHTML "          <td title=""e-Mail cadastrado na tela de dados do site, divulgado na tela ""Fale conosco"" do site da escola.""><font size=""1""><b>Site</font></td>"
            ShowHTML "        </tr>"
            While Not RS.EOF
              ShowHTML "      <tr bgcolor=""" & conTrBgColor & """ valign=""top"">"
              If (Instr(RS("ds_email"),"@") > 0 or Instr(RS("ds_email_contato"),"@") > 0 or Instr(RS("ds_email_internet"),"@") > 0) Then 
                 ShowHTML "        <td align=""center""><input type=""checkbox"" name=""w_sq_cliente"" value=""" & RS("sq_cliente") & """>"
              Else
                 ShowHTML "        <td align=""center""><input disabled type=""checkbox"" name=""w_sq_cliente"" value=""" & RS("sq_cliente") & """>"
              End If
              ShowHTML "        <td><font size=""1"">" & RS("ds_cliente") & "</td>"
              ShowHTML "        <td><font size=""1"">" & RS("ds_email") & "</td>"
              ShowHTML "        <td><font size=""1"">" & RS("ds_email_contato") & "</td>"
              ShowHTML "        <td><font size=""1"">" & RS("ds_email_internet") & "</td>"
              ShowHTML "      </tr>"
              RS.MoveNext
            wend
             ShowHTML "      </center>"
             ShowHTML "    </table>"
             ShowHTML "  </td>"
             ShowHTML "</tr>"
             ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000"">"
             ShowHTML "      <tr><td align=""center"" colspan=""3"">"
             ShowHTML "            <input class=""BTM"" type=""submit"" name=""Botao"" value=""Incluir"">"
             ShowHTML "            <input class=""BTM"" type=""button"" onClick=""javascript:window.close();javascript:opener.location.reload();"" name=""Botao"" value=""Fechar"">"
             ShowHTML "            <input class=""BTM"" type=""button"" onClick=""location.href='" & w_Pagina & par & "&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "&w_sq_cliente_mail=" & w_sq_cliente_mail & "&O=L';"" name=""Botao"" value=""Cancelar"">"
             ShowHTML "          </td>"
             ShowHTML "      </tr>"
             ShowHTML "    </table>"
             ShowHTML "    </TD>"
             ShowHTML "</tr>"
             ShowHTML "</FORM>"
          End If
          DesConectaBD     
       End If
    ElseIf SG = "ANEXO" Then ' Se for arquivo atachado
       ShowHTML "<FORM action=""" & w_Pagina & "Grava"" method=""POST"" name=""Form"" onSubmit=""return(Validacao(this));"" ENCTYPE=""multipart/form-data"">"
       ShowHTML "<INPUT type=""hidden"" name=""CL"" value=""" & CL & """>"
       ShowHTML "<INPUT type=""hidden"" name=""TP"" value=""" & TP & """>"
       ShowHTML "<INPUT type=""hidden"" name=""SG"" value=""" & SG & """>"
       ShowHTML "<INPUT type=""hidden"" name=""R"" value=""" & R & """>"
       ShowHTML "<INPUT type=""hidden"" name=""O"" value=""" & O &""">"
       ShowHTML "<INPUT type=""hidden"" name=""w_sq_cliente_mail"" value=""" & w_sq_cliente_mail & """>"

       ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td><div align=""justify""><font size=2><b><ul>Instruções</b>:<li>Clique sobre o botão <i>Procurar</i> para localizar o arquivo desejado.<li>Clique sobre o botão <i>Incluir</i> para adicionar o arquivo à relação.</ul><hr></div>"
       ShowHTML "<tr bgcolor=""" & conTrBgColor & """><td align=""center"">"
       ShowHTML "    <table width=""70%"" border=""0"">"
       ShowHTML "      <tr><td valign=""top""><font  size=""1""><b><U>A</U>rquivo:<br><INPUT ACCESSKEY=""N"" " & w_Disabled & " class=""BTM"" type=""FILE"" name=""p_regional"" size=""50"" maxlength=""80"" value=""" & p_regional & """>"
       ShowHTML "      <tr><td align=""center"" colspan=""3"" height=""1"" bgcolor=""#000000"">"
       ShowHTML "      <tr><td align=""center"" colspan=""3"">"
       ShowHTML "            <input class=""BTM"" type=""submit"" name=""Botao"" value=""Incluir"">"
       ShowHTML "            <input class=""BTM"" type=""button"" name=""Botao"" value=""Cancelar"" onClick=""history.back(1)"">"
       ShowHTML "          </td>"
       ShowHTML "      </tr>"
       ShowHTML "    </table>"
       ShowHTML "    </TD>"
       ShowHTML "</tr>"
       ShowHTML "</form>"
    End If
  Else
    ScriptOpen "JavaScript"
    ShowHTML " alert('Opção não disponível');"
    ShowHTML " history.back(1);"
    ScriptClose
  End If
  ShowHTML "</table>"
  ShowHTML "</center>"
  Rodape

  Set SQL1              = Nothing
  Set SQL2              = Nothing
  Set SQL3              = Nothing
  Set p_regional        = Nothing
  Set p_tipo_cliente    = Nothing
  Set p_modalidade      = Nothing
  Set w_sq_cliente_mail = Nothing
  Set w_sq_cliente      = Nothing
  Set w_sq_unidade      = Nothing
  Set RS1               = Nothing
        
End Sub
REM =========================================================================
REM Fim da tabela de Itens
REM -------------------------------------------------------------------------

REM =========================================================================
REM Procedimento que executa as operações de BD
REM -------------------------------------------------------------------------
Public Sub Grava

  Dim w_Null
  Dim w_mensagem
  Dim w_ordem
  Dim w_servico
  Dim I

  Cabecalho
  ShowHTML "</HEAD>"
  BodyOpen "onLoad=document.focus();"
  
  AbreSessao

  If O = "I" Then
     dbms.BeginTrans()
     If SG = "DESTINO" Then ' Se for destinatário
        For i = 1 To Request.Form("w_sq_cliente").Count
           w_ordem = Request.Form("w_sq_cliente")(i)
           'Verifica se o destinatário já consta da relação
           SQL = "select count(*) Existe " & VbCrLf & _
                 "from escMail_Destinatario " & VbCrLf & _
                 "where sq_cliente_mail = " & Request("w_sq_cliente_mail") & " " & VbCrLf & _
                 "  and destinatario    = " & w_ordem
           ConectaBD SQL
           ' Grava somente se não existir
           If Rs("Existe") = 0 Then
              SQL = "insert into escMail_Destinatario " & VbCrLf & _
                    "  (sq_cliente_mail, remetente, destinatario) " & VbCrLf & _
                    "values ( " & VbCrLf & _
                    "   " & Request("w_sq_cliente_mail") & ", " & VbCrLf & _
                    "   0, " & VbCrLf & _
                    "   " & w_ordem & " " & VbCrLf & _
                    ")"
              ExecutaSQL(SQL)
           End If
           DesconectaBD
        Next
        ' O SQL abaixo apenas desarma o DBExecuteSQL do fim da página, de modo a evitar que
        ' o último registro da lista seja gravado duas vezes
        SQL = "select count(*) Existe " & VbCrLf & _
              "from escMail_Destinatario " & VbCrLf & _
              "where sq_cliente_mail = " & Request("w_sq_cliente_mail")
        R = R & "&w_sq_cliente_mail=" & Request("w_sq_cliente_mail")
     ElseIf SG = "ANEXO" Then ' Se for arquivo atachado
        ' Se foi feito o upload de um arquivo
        If ul.Files("p_regional").OriginalPath > "" Then
           ' Verifica se o tamanho do arquivo está compatível com  o limite de 5MB.
           If ul.Files("p_regional").Size > 5242880 Then
              ScriptOpen("JavaScript")
              ShowHTML "  alert('Atenção: o tamanho máximo do arquivo não pode exceder 5 MBytes!');"
              ShowHTML "  history.back(1);"
              ScriptClose
              Response.End()
              exit sub
           End If

           w_file = ul.GetUniqueName()

           ' Grava o registro
           SQL = "insert into escMail_Anexo " & VbCrLf & _
                 "  (sq_cliente_mail, arquivo_origem, arquivo_destino) " & VbCrLf & _
                 "values ( " & VbCrLf & _
                 "   " & ul.Form("w_sq_cliente_mail") & ", " & VbCrLf & _
                 "   '" & ul.GetFileName(ul.Files("p_regional").OriginalPath) & "', " & VbCrLf & _
                 "   '" & w_file & "' " & VbCrLf & _
                 ")"
           ExecutaSQL(SQL)

           ' Grava o arquivo
           ul.Files("p_regional").SaveAs(Request.ServerVariables("APPL_PHYSICAL_PATH") & "sedf\sedf\" & w_file)

        End If
        dbms.CommitTrans()

        R = R & "&w_sq_cliente_mail=" & ul.Form("w_sq_cliente_mail")
        ScriptOpen "JavaScript"
        ShowHTML "  location.href='" & R & "&O=L&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "';"
        ScriptClose
        Exit Sub
     Else
        ' Recupera o próximo valor da chave
        SQL = "select IsNull(Max(sq_cliente_mail),0) + 1 chave from escCliente_Mail"
        ConectaBD SQL
        w_mensagem = RS("chave")
        DesconectaBD
        
        SQL = "insert into escCliente_Mail " & VbCrLf & _
              " (sq_cliente_mail, sq_cliente, sq_cliente_remetente, assunto, texto, envia_lista, data_inclusao) " & VbCrLf & _
              "values ( " & VbCrLf & _
              " " & w_mensagem & ", " & VbCrLf & _
              " 0, " & VbCrLf & _
              " " & replace(Request("CL"),"sq_cliente=","") & "," & VbCrLf & _
              " '" & trim(Request("w_assunto")) & "'," & VbCrLf & _
              " '" & trim(Request("w_texto")) & "'," & VbCrLf & _
              " '" & Request("w_envia_lista") & "'," & VbCrLf & _
              " getDate()" & VbCrLf & _
              ")" & VbCfLf
     End If
  ElseIf O = "A" Then
     dbms.BeginTrans()
     SQL = "update escCliente_Mail set " & VbCrLf & _
           "assunto     = '" & trim(Request("w_assunto")) & "', " & VbCrLf & _
           "texto       = '" & trim(Request("w_texto")) & "', " & VbCfLf & VbCrLf & _
           "envia_lista = '" & trim(Request("w_envia_lista")) & "' " & VbCrLf & _
           "where sq_cliente_mail = '" & Request("w_sq_cliente_mail") & "' " & VbCfLf
  ElseIf O = "E" Then
     dbms.BeginTrans()
     If SG = "DESTINO" Then ' Se for destinatário da mensagem
        SQL = "delete escMail_Destinatario " & VbCrLf & _
              "where sq_cliente_mail = " & Request("w_sq_cliente_mail") & " " & VbCrLf & _
              "  and destinatario    = " & Request("w_destinatario")
        R = R & "&w_sq_cliente_mail=" & Request("w_sq_cliente_mail")
     ElseIf SG = "ANEXO" Then ' Se for arquivo atachado
        DeleteAFile(Request.ServerVariables("APPL_PHYSICAL_PATH") & "sedf\sedf\" & Request("w_atual"))
        SQL = "delete escMail_Anexo " & VbCrLf & _
              "where sq_mail_anexo = " & Request("w_sq_mail_anexo")
        R = R & "&w_sq_cliente_mail=" & Request("w_sq_cliente_mail")
     Else
        ' Apaga fisicamente os arquivos anexos à mensagem
        SQL = "select arquivo_destino from escMail_Anexo " & VbCrLf & _
              "where sq_cliente_mail = " & Request("w_sq_cliente_mail")
        ConectaBD SQL
        While Not RS.EOF
           DeleteAFile(Request.ServerVariables("APPL_PHYSICAL_PATH") & "files\" & RS("arquivo_destino"))
           RS.MoveNext
        Wend
        DesconectaBD
        ' Apaga os registros dos arquivos anexos à mensagem
        SQL = "delete escMail_Anexo where sq_cliente_mail = " & Request("w_sq_cliente_mail")
        ExecutaSQL(SQL)
        ' Apaga os destinatários da mensagem
        SQL = "delete escMail_Destinatario where sq_cliente_mail = " & Request("w_sq_cliente_mail")
        ExecutaSQL(SQL)
        ' Apaga a mensagem
        SQL = "delete escCliente_Mail where sq_cliente_mail = " & Request("w_sq_cliente_mail")
     End If
  ElseIf O = "V" Then
     ' Define as variáveis necessárias ao envio da mensagem
     Dim Mail
     Dim w_Remetente, w_Destino, w_arquivo, w_lista
     Dim w_Nome
     Dim w_Assunto
     Dim w_Texto
     Dim w_Maximo
     Dim w_Cont
     Dim FS
        
     'Recupera os dados da mensagem
     SQL = "select a.sq_cliente_mail, a.envia_lista, b.ds_cliente nome, a.assunto, a.texto, b.ds_email email, b.ds_username " & VbCrLf & _
           "from escCliente_Mail              a  " & VbCrLf & _
           "     inner   join escCliente      b  on (a.sq_cliente_remetente = b.sq_cliente) " & VbCrLf & _
           "where a.sq_cliente_mail = " & Request("w_sq_cliente_mail")
     ConectaBD SQL
     w_remetente = RS("email")
     w_nome      = RS("nome")
     w_assunto   = RS("assunto")
     w_texto     = RS("texto")
     w_lista     = RS("envia_lista")
     w_diretorio = replace(conFilePhysical & "\sedf\sedf\","\\","\")
     w_maximo    = 20
     w_cont      = 0
     DesconectaBD

     ' Instancia os objetos de e-mail e de tratamento de arquivo
     Set Mail  = Server.CreateObject("Persits.MailSender")
     Mail.Host = conSMTPServer
      
     ' Carrega os valores dos atributos
     Mail.From     = w_remetente
     Mail.FromName = w_nome
     Mail.Subject  = w_assunto
     Mail.Body     = Replace(w_texto,VbCrLf,"<br>")
     Mail.isHTML   = True

     'Recupera os arquivos anexos à mensagem
     SQL = "select a.arquivo_origem, a.arquivo_destino " & VbCrLf & _
           "from escMail_Anexo a  " & VbCrLf & _
           "where sq_cliente_mail = " & Request("w_sq_cliente_mail")
     ConectaBD SQL
     If Not RS.EOF Then
        Set FS = CreateObject("Scripting.FileSystemObject")
        While Not RS.EOF
          FS.CopyFile w_diretorio & RS("arquivo_destino"), w_diretorio & RS("arquivo_origem"), True
          Mail.AddAttachment w_diretorio & RS("arquivo_origem")
          RS.MoveNext
        Wend
     End If
     DesconectaBD

     If w_lista = "S" Then
        'Recupera os integrantes da lista de distribuição
        SQL = "select * from escNewsLetter a where envia_mail = 'S'"
        ConectaBD SQL
        If Not RS.EOF Then
          While Not RS.EOF
             Mail.AddAddress RS("email"), RS("nome")
             RS.MoveNext
          Wend
        End If
        DesconectaBD
     End If
     
     'Recupera os destinatários da mensagem
     SQL = "select b.ds_cliente nome, b.ds_email email " & VbCrLf & _
           "from escMail_Destinatario  a  " & VbCrLf & _
           "     inner join escCliente b on (a.destinatario = b.sq_cliente) " & VbCrLf & _
           "where a.sq_cliente_mail = " & Request("w_sq_cliente_mail")
     ConectaBD SQL
     If Not RS.EOF Then
       While Not RS.EOF
          Mail.AddAddress RS("email"), RS("nome")
          RS.MoveNext
       Wend
     End If
     DesconectaBD
     On Error Resume Next
     Mail.Send
     If Err <> 0 Then 
       ScriptOpen "JavaScript"
       ShowHTML "  alert('Não foi possível enviar a mensagem devido a problemas de rede. Entre em contato com a área de rede.\nErro: " & Err.number & " - " & replace(Err.description,VbCrLf," - ") & "');"
       ShowHTML "  history.back(1);"
       ScriptClose
       Exit Sub
     End If

     'Apaga os arquivos temporários que foram anexados à mensagem
     SQL = "select a.arquivo_origem, a.arquivo_destino " & VbCrLf & _
           "from escMail_Anexo a  " & VbCrLf & _
           "where sq_cliente_mail = " & Request("w_sq_cliente_mail")
     ConectaBD SQL
     If Not RS.EOF Then
        Set FS = CreateObject("Scripting.FileSystemObject")
        While Not RS.EOF
          FS.DeleteFile w_diretorio & RS("arquivo_origem"), True
          RS.MoveNext
        Wend
     End If
     DesconectaBD

     dbms.BeginTrans()
     ' Atualiza a mensagem
     SQL = "update escCliente_Mail set data_envio = getDate() where sq_cliente_mail = " & Request("w_sq_cliente_mail")
        
     Set FS          = Nothing
     Set Mail        = Nothing
     Set w_lista     = Nothing
     Set w_remetente = Nothing
     Set w_destino   = Nothing
     Set w_assunto   = Nothing
     Set w_texto     = Nothing
     Set w_nome      = Nothing
     Set w_maximo    = Nothing
     Set w_cont      = Nothing

  End If
  ExecutaSQL(SQL)
  dbms.CommitTrans()

  ScriptOpen "JavaScript"
  ShowHTML "  location.href='" & R & "&O=L&CL=" & CL & "&TP=" & TP & "&SG=" & SG & "';"
  ScriptClose

  Set I                 = Nothing
  Set w_servico         = Nothing
  Set w_ordem           = Nothing
  Set w_mensagem        = Nothing
  Set w_Null            = Nothing
End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que executa as operações de BD
REM =========================================================================

REM =========================================================================
REM Rotina principal
REM -------------------------------------------------------------------------
Sub Main
  ' Verifica se o usuário tem lotação e localização
  If (len(Session("LOTACAO")&"") = 0 or len(Session("LOCALIZACAO")&"") = 0) and Session("LogOn") = "Sim" Then
    ScriptOpen "JavaScript"
    ShowHTML " alert('Você não tem lotação ou localização definida. Entre em contato com o RH!'); "
    ShowHTML " top.location.href='Default.asp'; "
    ScriptClose
   Exit Sub
  End If

  Select Case Par
    Case "INICIAL"
       Inicial
    Case "ITEM"
       Item
    Case "GRAVA"
       Grava
    Case Else
       Cabecalho
       BodyOpen "onLoad=document.focus();"
       ShowHTML "<B><FONT COLOR=""#000000"">" & w_TP & "</FONT></B>"
       ShowHTML "<HR>"
       ShowHTML "<div align=center><center><br><br><br><br><br><br><br><br><br><br><img src=""img/underc.gif"" align=""center""> <b>Esta opção está sendo desenvolvida.</b><br><br><br><br><br><br><br><br><br><br></center></div>"
       Rodape
  End Select
End Sub
REM =========================================================================
REM Fim da rotina principal
REM -------------------------------------------------------------------------
%>

