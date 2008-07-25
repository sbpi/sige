<%
Session.LCID = 1046

REM =========================================================================
REM Declaração inicial para páginas OLE com Word
REM -------------------------------------------------------------------------
Sub headerWord (p_orientation)
  Response.ContentType = "application/msword"
  ShowHTML "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" "
  ShowHTML "xmlns:w=""urn:schemas-microsoft-com:office:word"" "
  ShowHTML "xmlns=""http://www.w3.org/TR/REC-html40""> "
  ShowHTML "<head> "
  ShowHTML "<meta http-equiv=Content-Type content=""text/html; charset=windows-1252""> "
  ShowHTML "<meta name=ProgId content=Word.Document> "
  ShowHTML "<!--[if gte mso 9]><xml> "
  ShowHTML " <w:WordDocument> "
  ShowHTML "  <w:View>Print</w:View> "
  ShowHTML "  <w:Zoom>BestFit</w:Zoom> "
  ShowHTML "  <w:SpellingState>Clean</w:SpellingState> "
  ShowHTML "  <w:GrammarState>Clean</w:GrammarState> "
  ShowHTML "  <w:HyphenationZone>21</w:HyphenationZone> "
  ShowHTML "  <w:Compatibility> "
  ShowHTML "   <w:BreakWrappedTables/> "
  ShowHTML "   <w:SnapToGridInCell/> "
  ShowHTML "   <w:WrapTextWithPunct/> "
  ShowHTML "   <w:UseAsianBreakRules/> "
  ShowHTML "  </w:Compatibility> "
  ShowHTML "  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel> "
  ShowHTML " </w:WordDocument> "
  ShowHTML "</xml><![endif]--> "
  ShowHTML "<style> "
  ShowHTML "<!-- "
  ShowHTML " /* Style Definitions */ "
  ShowHTML "@page Section1 "
  If uCase(Nvl(p_orientation,"LANDSCAPE")) = "PORTRAIT" Then
     ShowHTML "    {mso-page-orientation:portrait; "
     ShowHTML "    margin:1.0cm 1.0cm 1.28cm 1.0cm; "
     ShowHTML "    mso-header-margin:10; "
     ShowHTML "    mso-footer-margin:10; "
     ShowHTML "    mso-paper-source:0;} "
  Else
     ShowHTML "    {size:11.0in 8.5in; "
     ShowHTML "    mso-page-orientation:landscape; "
     ShowHTML "    margin:1.0cm 1.0cm 1.0cm 1.0cm; "
     ShowHTML "    mso-header-margin:15pt; "
     ShowHTML "    mso-footer-margin:15pt; "
     ShowHTML "    mso-paper-source:0;} "
  End If
  ShowHTML "div.Section1 "
  ShowHTML "    {page:Section1;} "
  ShowHTML "--> "
  ShowHTML "</style> "
  ShowHTML "</head> "
  BodyOpenClean "onLoad=document.focus();"
  ShowHTML "<div class=Section1> "
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da barra de navegação de recordsets
REM -------------------------------------------------------------------------
Sub MontaBarra (p_link, p_PageCount, p_AbsolutePage, p_PageSize, p_RecordCount)

  ShowHTML "<SCRIPT LANGUAGE=""JAVASCRIPT"">"
  ShowHTML "  function pagina (pag) {"
  ShowHTML "    document.Barra.P3.value = pag;"
  ShowHTML "    document.Barra.submit();"
  ShowHTML "  }"
  ShowHTML "</SCRIPT>"
  ShowHTML "<FORM ACTION=""" & p_link &""" METHOD=""POST"" name=""Barra"">"
  ShowHTML "<input type=""Hidden"" name=""P4"" value=""" & p_pagesize & """>"
  ShowHTML "<input type=""Hidden"" name=""P3"" value="""">"
  ShowHTML "<input type=""Hidden"" name=""Q"" value=""" & Q & """>"
  ShowHTML "<input type=""Hidden"" name=""" & w_ew & """ value=""" & w_ew & """>"
  ShowHTML MontaFiltro("POST")
  If p_PageSize < p_RecordCount Then
     If p_PageCount = p_AbsolutePage Then
        ShowHTML "<span class=""BTM""><br>" & (p_RecordCount-((p_PageCount-1)*p_PageSize)) & " linhas apresentadas de " & p_RecordCount & " linhas"
     Else
        ShowHTML "<span class=""BTM""><br>" & p_PageSize & " linhas apresentadas de " & p_RecordCount & " linhas"
     End If
     ShowHTML "<br>na página " & p_AbsolutePage & " de " & p_PageCount & " páginas"
     If p_AbsolutePage > 1 Then
        ShowHTML "<br>[<A class=""SS"" TITLE=""Primeira página"" HREF=""javascript:pagina(1)"" onMouseOver=""window.status='Primeira (1/" & p_PageCount& ")'; return true"" onMouseOut=""window.status=''; return true"">Primeira</A>]&nbsp;"
        ShowHTML "[<A class=""SS"" TITLE=""Página anterior"" HREF=""javascript:pagina(" & p_AbsolutePage-1 & ")"" onMouseOver=""window.status='Anterior (" & p_AbsolutePage-1 & "/" & p_PageCount& ")'; return true"" onMouseOut=""window.status=''; return true"">Anterior</A>]&nbsp;"
     Else
        ShowHTML "<br>[Primeira]&nbsp;"
        ShowHTML "[Anterior]&nbsp;"
     End If
     If p_PageCount = p_AbsolutePage Then
        ShowHTML "[Próxima]&nbsp;"
        ShowHTML "[Última]"
     Else
        ShowHTML "[<A class=""SS"" TITLE=""Página seguinte"" HREF=""javascript:pagina(" & p_AbsolutePage+1 & ")""  onMouseOver=""window.status='Próxima (" & p_AbsolutePage+1 & "/" & p_PageCount& ")'; return true"" onMouseOut=""window.status=''; return true"">Próxima</A>]&nbsp;"
        ShowHTML "[<A class=""SS"" TITLE=""Última página"" HREF=""javascript:pagina(" & p_PageCount & ")""  onMouseOver=""window.status='Última (" & p_PageCount & "/" & p_PageCount& ")'; return true"" onMouseOut=""window.status=''; return true"">Última</A>]"
     End If
     ShowHTML "</span>"
  End If
  ShowHtml "</FORM>"

End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da URL com os parâmetros de filtragem
REM -------------------------------------------------------------------------
Function MontaFiltro (p_method)
  Dim l_string, l_item
  If Instr("POST, GET",p_method) > 0 Then
     l_string = ""
     For Each l_Item IN Request.Form
        If lCase(l_item) <> "w_ew" and lCase(l_item) <> "w_ea" and lCase(l_item) <> "p3" and lCase(l_item) <> "p4" Then
           If Request(l_item) > "" Then
              If uCase(p_method) = "GET" Then
                 l_string = l_string & "&" & l_Item & "=" & Request(l_Item)
              ElseIf uCase(p_method) = "POST" Then
                 l_string = l_string & "<INPUT TYPE=""HIDDEN"" NAME=""" & l_Item & """ VALUE=""" & Request(l_Item) & """>" & VbCrLf
              End If
           End If
        End If
     Next
     For Each l_Item IN Request.QueryString
        If lCase(l_item) <> "w_ew" and lCase(l_item) <> "w_ea" and lCase(l_item) <> "p3" and lCase(l_item) <> "p4" Then
           If Request(l_item) > "" Then
              If uCase(p_method) = "GET" Then
                 l_string = l_string & "&" & l_Item & "=" & Request(l_Item)
              ElseIf uCase(p_method) = "POST" Then
                 l_string = l_string & "<INPUT TYPE=""HIDDEN"" NAME=""" & l_Item & """ VALUE=""" & Request(l_Item) & """>" & VbCrLf
              End If
           End If
        End If
     Next
  End If
  
  MontaFiltro = l_string
  
  Set l_string = Nothing
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------

REM =========================================================================
REM Monta string html para montagem de calendário do mês informado
REM -------------------------------------------------------------------------
Function MontaCalendario(p_mes, p_datas, p_cores)
   Dim l_html, l_mes, l_ano, l_inicio, l_cont, l_cor, l_cor_padrao, l_borda, l_ocorrencia
   Dim l_celulas(42), l_meses(12), l_dias(7), l_qtd(12)
   
   ' Atribui nomes dos meses
   l_meses(1) = "Janeiro"    : l_meses(2) = "Fevereiro" : l_meses(3) = "Março"     : l_meses(4) = "Abril"
   l_meses(5) = "Maio"       : l_meses(6) = "Junho"     : l_meses(7) = "Julho"     : l_meses(8) = "Agosto"
   l_meses(9) = "Setembro"   : l_meses(10) = "Outubro"  : l_meses(11) = "Novembro" : l_meses(12) = "Dezembro"

   ' Atribui quantidade de dias em cada mês
   l_qtd(1) = 31 : l_qtd(2) = 28 : l_qtd(3) = 31 : l_qtd(4) = 30  : l_qtd(5) = 31  : l_qtd(6) = 30
   l_qtd(7) = 31 : l_qtd(8) = 31 : l_qtd(9) = 30 : l_qtd(10) = 31 : l_qtd(11) = 30 : l_qtd(12) = 31

   ' Atribui sigla para cada dia da semana
   l_dias(1)  = "D" : l_dias(2)  = "S" : l_dias(3)  = "T" : l_dias(4)  = "Q"
   l_dias(5)  = "Q" : l_dias(6)  = "S" : l_dias(7)  = "S"

   ' Recupera o mês e o ano desejado para montagem do calendário
   l_mes = Mid(p_mes,1,2)
   l_ano = Mid(p_mes,3,4)
   
   ' Define cor de fundo padrão para as células de sábado e domingo
   l_cor_padrao = " bgcolor=""#DAEABD"""
   
   ' Trata o mês de fevereiro anos bissextos
   If (l_ano Mod 4) = 0 Then l_qtd(2) = 29 End If
   
   l_html = ""
   l_html = l_html & "<table border=0>" & VbCrLf
   l_html = l_html & "  <tr><td colspan=7 align=""center""" & l_cor_padrao & "><b>" & l_meses(l_mes) & "/" & l_ano & "</td></tr>" & VbCrLf
   l_html = l_html & "  <tr align=""center"">" & VbCrLf
   ' Monta a linha com a sigla para os dias das semanas
   For Each l_item IN l_dias
       If l_item > "" Then l_html = l_html & "    <td" & l_cor_padrao & """><b>" & l_item & "</td>" & VbCrLf End If
   Next
   l_html = l_html & "  </tr>" & VbCrLf
   
   ' Define em que dia da semana o mês inicia
   l_inicio = WeekDay(cDate("01/"&l_mes&"/"&l_ano))

   ' Carrega os dias do mês num array que será usado para montagem do calendário, colocando
   ' o dia ou um espaço em branco, dependendo do início e do fim do mês
   For l_cont = 1              To l_inicio-1   : l_celulas(l_cont) = "&nbsp;" : Next
   For l_cont = l_qtd(l_mes)+1 To 42           : l_celulas(l_cont) = "&nbsp;" : Next

   For l_cont = 1 To l_qtd(l_mes)
      l_celulas(l_cont + l_inicio - 1) = l_cont
   Next
   
   ' Monta o calendário, usando o array l_celulas
   l_cont = 1
   l_html = l_html & "  <tr align=""center"">" & VbCrLf 
   For Each l_item IN l_celulas
       ' Desconsidera itens do array sem conteúdo
       If l_item > "" Then
          ' Trata a borda da célula para datas especiais
          l_borda = ""
          l_ocorrencia = ""
          If l_item <> "&nbsp;" Then
             If p_datas(l_item, l_mes, Mid(l_ano,3,2)) > "" Then
                l_borda = " style=""border: 1px solid rgb(0,0,0);"""
                l_ocorrencia = " title=""" & p_datas(l_item, l_mes, Mid(l_ano,3,2)) & """"
             End If
          End If
          
          ' Trata a cor de fundo da célula
          l_cor = "" 
          If l_item <> "&nbsp;" and (l_cont = 1 or (l_cont Mod 7) = 0 or ((l_cont-1) Mod 7) = 0) Then 
             l_cor = l_cor_padrao
          ElseIf l_item <> "&nbsp;" Then 
             If p_cores(l_item, l_mes, Mid(l_ano,3,2)) > "" Then
                l_cor = " bgcolor=""" & p_cores(l_item, l_mes, Mid(l_ano,3,2)) & """"
             End If
          End If
          
          ' Coloca uma célula do calendário
          l_html = l_html & "    <td" & l_cor & l_borda & l_ocorrencia & ">" & l_item & "</td>" & VbCrLf 
          
          ' Trata a quebra de linha ao final de cada semana
          If ((l_cont Mod 7) = 0) Then
             l_html = l_html & "  </tr>" & VbCrLf 
             l_html = l_html & "  <tr align=""center"">" & VbCrLf 
          End If
          l_cont = l_cont + 1
       End If
   Next
   l_html = l_html & "</table>" & VbCrLf
   
   ' Devolve o calendário montado
   MontaCalendario = l_html
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------

REM =========================================================================
REM Converte CFLF para <BR>
REM -------------------------------------------------------------------------
Function CRLF2BR(expressao)
   If IsNull(expressao) or expressao = "" Then
      CRLF2BR = ""
   Else
      CRLF2BR = Replace(expressao, VbCrLf, "<BR>")
   End If
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------

REM =========================================================================
REM Trata valores nulos
REM -------------------------------------------------------------------------
Function Nvl(expressao,valor)
   If IsNull(expressao) or expressao = "" Then
      Nvl = valor
   Else
      Nvl = expressao
   End If
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------

REM =========================================================================
REM Retorna valores nulos se chegar cadeia vazia
REM -------------------------------------------------------------------------
Function Tvl(expressao)
   If IsNull(expressao) or expressao = "" Then
      Tvl = null
   Else
      Tvl = expressao
   End If
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------.

REM =========================================================================
REM Retorna valores nulos se chegar cadeia vazia
REM -------------------------------------------------------------------------
Function Cvl(expressao)
   If IsNull(expressao) or expressao = "" Then
      Cvl = 0
   Else
      Cvl = expressao
   End If
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------

REM =========================================================================
REM Retorna o caminho físico para o diretório  do cliente informado
REM -------------------------------------------------------------------------
Function DiretorioCliente(p_Cliente)
   DiretorioCliente = Request.ServerVariables("APPL_PHYSICAL_PATH") & "files\" & p_cliente
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------.

REM =========================================================================
REM Montagem de URL a partir da sigla da opção do menu
REM -------------------------------------------------------------------------
Function MontaURL (p_sigla)
  Dim RS_montaUrl, l_imagem, l_imagemPadrao

  DB_GetLinkData RS_montaUrl, Session("p_cliente"), "MESA"

  l_ImagemPadrao = "images/folder/SheetLittle.gif"
  
  If RS_montaUrl.EOF Then
     MontaURL = ""
  Else
     If RS_montaUrl("IMAGEM") > "" Then l_Imagem = RS_montaUrl("IMAGEM") Else l_Imagem = l_ImagemPadrao End If
     MontaURL = RS_montaUrl("LINK") & "&P1="&RS_montaUrl("P1")&"&P2="&RS_montaUrl("P2")&"&P3="&RS_montaUrl("P3")&"&P4="&RS_montaUrl("P4")&"&TP=<img src="&l_imagem&" BORDER=0>"&RS_montaUrl("NOME")&"&SG="&RS_montaUrl("SIGLA")
  End If
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem de cabeçalho padrão de formulário
REM -------------------------------------------------------------------------
Sub AbreForm (p_Name,p_Action,p_Method,p_onSubmit,p_Target)
    If IsNull(p_Target) Then
       ShowHTML "<FORM action=""" & p_action & """ method=""" & p_Method & """ name=""" & p_Name & """ onSubmit=""" & p_onSubmit & """>"
    Else
       ShowHTML "<FORM action=""" & p_action & """ method=""" & p_Method & """ name=""" & p_Name & """ onSubmit=""" & p_onSubmit & """ target=""" & p_Target & """>"
    End If
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem de campo do tipo radio com padrão Não
REM -------------------------------------------------------------------------
Sub MontaRadioNS (Label, Chave, Campo)
    ShowHTML "          <td><font size=""1"">"
    If Nvl(Label,"") > "" Then
       ShowHTML Label & "</b><br>"
    End If
    If chave = "Sim" or chave = "S" Then
       ShowHTML "              <input " & w_Disabled & " type=""radio"" name=""" & campo & """ value=""S"" checked> Sim <input " & w_Disabled & " type=""radio"" name=""" & campo & """ value=""N""> Não"
    Else
       ShowHTML "              <input " & w_Disabled & " type=""radio"" name=""" & campo & """ value=""S""> Sim <input " & w_Disabled & " type=""radio"" name=""" & campo & """ value=""N"" checked> Não"
    End If
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem de campo do tipo radio com padrão Sim
REM -------------------------------------------------------------------------
Sub MontaRadioSN (Label, Chave, Campo)
    ShowHTML "          <td><font size=""1"">"
    If Nvl(Label,"") > "" Then
       ShowHTML Label & "</b><br>"
    End If
    If chave = "Nao" or chave = "N" Then
       ShowHTML "              <input " & w_Disabled & " type=""radio"" name=""" & campo & """ value=""Sim""> Sim <input " & w_Disabled & " type=""radio"" name=""" & campo & """ value=""Nao"" checked> Não"
    Else
       ShowHTML "              <input " & w_Disabled & " type=""radio"" name=""" & campo & """ value=""Sim"" checked> Sim <input " & w_Disabled & " type=""radio"" name=""" & campo & """ value=""Nao""> Não"
    End If
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da seleção de escolas
REM -------------------------------------------------------------------------
REM =========================================================================
REM Montagem da seleção de escolas
REM -------------------------------------------------------------------------
Sub SelecaoEscola (label, accesskey, hint, chave, chaveAux, campo, restricao, atributo)

    
    SQL = "SELECT a.sq_cliente cliente, a.sq_tipo_cliente, case b.tipo when 1 then upper(a.ds_username) else a.ds_cliente end ds_cliente, b.tipo " & VbCrLf & _
          "  FROM escCLIENTE            a" & VbCrLf & _
          "      inner join escTipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente) " & VbCrLf & _
          " WHERE PUBLICA = 'S' and upper(a.ds_username) <> 'SBPI' "
    If chaveAux > "" Then
       SQL = SQL & "   and IsNull(sq_cliente_pai,0) = " & chaveAux & VbCrLf
    End If
    SQL = SQL & "ORDER BY b.tipo, ds_cliente" & VbCrLf
    ConectaBD SQL
    If IsNull(hint) Then
       ShowHTML "          <td valign=""top""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    Else
       ShowHTML "          <td valign=""top"" ONMOUSEOVER=""popup('" & hint & "','white')""; ONMOUSEOUT=""kill()""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    End If
    ShowHTML "          <option value="""">---"
    While Not RS.EOF
       If cDbl(nvl(RS("cliente"),-1)) = cDbl(nvl(chave,-1)) Then
          ShowHTML "          <option value=""" & RS("cliente") & """ SELECTED>" & RS("ds_cliente")
       Else
          ShowHTML "          <option value=""" & RS("cliente") & """>" & RS("ds_cliente")
       End If
       RS.MoveNext
    Wend
    ShowHTML "          </select>"
    DesconectaBD
End Sub

REM =========================================================================
REM Função que formata dias, horas, minutos e segundos a partir dos segundos
REM -------------------------------------------------------------------------
Function FormataTempo(p_segundos)
  Dim l_dias, l_horas, l_minutos, l_segundos,  l_tempo

  l_horas    = Int(p_segundos/3600)
  l_minutos  = Int((p_segundos - (l_horas*3600))/60)
  l_segundos = p_segundos - (l_horas*3600) - (l_minutos*60)
  FormataTempo = Mid(1000+l_horas,2,3) & ":" & Mid(100+l_minutos,2,2) & ":" & Mid(100+l_segundos,2,2)

  Set l_tempo       = Nothing
  Set l_dias        = Nothing
  Set l_horas       = Nothing
  Set l_minutos     = Nothing
  Set l_segundos    = Nothing
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina que encerra a sessão e fecha a janela do SIW
REM -------------------------------------------------------------------------
Sub EncerraSessao
  ScriptOpen "JavaScript"
  ShowHTML " alert('Tempo máximo de inatividade atingido! Autentique-se novamente.'); "
  ScriptClose
  Response.End()
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Função que formata um texto para exibição em HTML
REM -------------------------------------------------------------------------
Function ExibeTexto(p_texto)
    ExibeTexto = Replace(Replace(p_texto,VbCrLf, "<br>"),"  ","&nbsp;&nbsp;")
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------

REM =========================================================================
REM Função que retorna a data/hora do banco
REM -------------------------------------------------------------------------
Function DataHora()
    DataHora = FormatDateTime(Date(),1) & ", " & FormatDateTime(Time(),3)
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina que monta a máscara do beneficiário
REM -------------------------------------------------------------------------
Function MascaraBeneficiario(cgccpf)
  ' Se o campo tiver máscara, retira
  If Instr(cgccpf,".") > 0 Then
     MascaraBeneficiario = Replace(Replace(Replace(cgccpf,".",""),"-",""),"/","")
  ' Caso contrário, aplica a máscara, dependendo do tamanho do parâmetro
  ElseIf Len(cgccpf) = 11 Then
     MascaraBeneficiario = Mid(cgccpf,1,3) & "." & Mid(cgccpf,4,3) & "." & Mid(cgccpf,7,3) & "-" & Mid(cgccpf,10,2)
  ElseIf Len(cgccpf) = 14 Then
     MascaraBeneficiario = Mid(cgccpf,1,2) & "." & Mid(cgccpf,3,3) & "." & Mid(cgccpf,6,3) & "/" & Mid(cgccpf,9,4) & "-" & Mid(cgccpf,13,2)
  End If
End Function
REM =========================================================================
REM Final da rotina 
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina de envio de e-mail
REM -------------------------------------------------------------------------
Function EnviaMail(w_subject, w_mensagem, w_recipients, w_attachments)

  Dim JMail
  Dim Attachments(500), Recipients(500), i, j
  Set JMail  = Server.CreateObject("Persits.MailSender")
 
  JMail.Host     = conSMTPServer
  JMail.From     = "siw.sbpi"
  JMail.FromName = "SEDF - Secretaria de Estado de Educação do DF"

  JMail.Subject  = w_subject
  JMail.isHTML   = True
  JMail.Body     = w_mensagem
  i = 0
  w_recipients = w_recipients & ";"
  Do While Instr(w_recipients,";") > 0
     If Len(w_recipients) > 2 Then Recipients(i) = Mid(w_recipients,1,Instr(w_recipients,";")-1) End If
     w_recipients = Mid(w_recipients,Instr(w_recipients,";")+1,Len(w_recipients))
     i = i+1
  Loop
  For j = 0 To i-1
     JMail.AddAddress Recipients(j)
  Next

  If Nvl(w_attachments,"nulo") <> "nulo" Then
     w_attachments = w_attachments & ";"
     Do While Instr(w_attachments,";") > 0
        If Len(w_attachments) > 2 Then attachments(i) = Mid(w_attachments,1,Instr(w_attachments,";")-1) End If
        w_attachments = Mid(w_attachments,Instr(w_attachments,";")+1,Len(w_attachments))
        i = i+1
     Loop
     For j = 0 To i-1
        JMail.AddAttachment Attachments(j)
     Next
  End If

  On Error Resume Next

  JMail.Send

  If JMail.ErrorCode > 0 Then
    EnviaMail = "Erro: " & JMail.ErrorCode & "\nMensagem: " & JMail.ErrorMessage & "\nFonte: " & JMail.ErrorSource
  Else
    EnviaMail = ""
  End If
  
End Function
REM =========================================================================
REM Fim da rotina de envio de email
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina de envio de e-mail
REM -------------------------------------------------------------------------
Function EnviaMail1(w_subject, w_mensagem, w_recipients)

  w_recipients = w_recipients & ";"
  Dim JMail
  Dim Recipients(500), i, j

  Set JMail = Server.CreateObject("JMail.Message")
 
  JMail.Silent   = True
  JMail.FromName = conHead
  JMail.Logging  = False

  'Recupera os dados da mensagem
  SQL = "select a.ds_email email, a.ds_cliente " & VbCrLf & _
        "from escCliente      a " & VbCrLf & _
        "where a.sq_cliente = " & CL
  ConectaBD SQL
  JMail.From     = RS("email")
  JMail.MailServerUserName = RS("email")
  DesconectaBD
  
  JMail.Subject  = w_subject
  JMail.HtmlBody = w_mensagem
  JMail.ClearRecipients()
  i = 0
  Do While Instr(w_recipients,";") > 0
     If Len(w_recipients) > 2 Then Recipients(i) = Mid(w_recipients,1,Instr(w_recipients,";")-1) End If
     w_recipients = Mid(w_recipients,Instr(w_recipients,";")+1,Len(w_recipients))
     i = i+1
  Loop

  For j = 0 To i-1
     If j > 0 or i = 1 Then ' Se for o primeiro, coloca como destinatário principal. Caso contrário, coloca como CC.
        JMail.AddRecipient Recipients(j)
     Else
        JMail.AddRecipientBCC Recipients(j)
     End If
  Next

  JMail.Send conSMTPServer

  If JMail.ErrorCode <> 0 Then
    'EnviaMail = Replace("Erro: " & JMail.ErrorCode & "\nServidor: " & Session("smtp_server") & "\nConta: " & Session("siw_email_conta") & "\nSenha: " & Nvl(Session("siw_email_senha"),"---") & "\nMensagem: " & JMail.ErrorMessage & "\nFonte: " & JMail.ErrorSource,VbCrLf,"\n")
    EnviaMail1 = Replace("Erro: " & JMail.ErrorCode & "\nMensagem: " & JMail.ErrorMessage & "\nFonte: " & JMail.ErrorSource,VbCrLf,"\n")
  Else
    EnviaMail1 = ""
  End If
  
End Function
REM =========================================================================
REM Fim da rotina de envio de email
REM -------------------------------------------------------------------------

' Cria a tag Body
Function BodyOpenMail(cProperties)
Dim l_html
l_html = "<BASEFONT FACE=""Verdana"" SIZE=""2""> " & VbCrLf
l_html = l_html & "<style> " & VbCrLf
l_html = l_html &  " .SS{text-decoration:none;font:bold 8pt} " & VbCrLf
l_html = l_html &  " .SS:HOVER{text-decoration: underline;} " & VbCrLf
l_html = l_html &  " .HL{text-decoration:none;font:Arial;color=""#0000FF""} " & VbCrLf
l_html = l_html &  " .HL:HOVER{text-decoration: underline;} " & VbCrLf
l_html = l_html &  " .TTM{font: 10pt Arial}" & VbCrLf
l_html = l_html &  " .BTM{font: 8pt Verdana}" & VbCrLf
l_html = l_html &  " .XTM{font: 12pt Verdana}" & VbCrLf
l_html = l_html &  "</style> " & VbCrLf
l_html = l_html &  "<body Text=""" & conBodyText & """ Link=""" & conBodyLink & """ Alink=""" & conBodyALink & """ " & _
	"Vlink=""" & conBodyVLink & """ Bgcolor=""" & conBodyBgcolor & """ Background=""" & conBodyBackground & """ " & _
	"Bgproperties=""" & conBodyBgproperties & """ Topmargin=""" & conBodyTopmargin & """ " & _
	"Leftmargin=""" & conBodyLeftmargin & """ " & cProperties & "> " & VbCrLf
BodyOpenMail = l_html
Set l_html = Nothing
End Function

REM =========================================================================
REM Rotina que extrai a última parte da variável TP
REM -------------------------------------------------------------------------
Function RemoveTP(TP)
  Dim w_TP
  w_TP = TP
  While InStr(w_TP,"-") > 0
     w_TP = Mid(w_TP,InStr(w_TP,"-")+1,Len(w_TP))
  Wend
  RemoveTP = replace(TP," -"&w_TP,"")
End Function
REM =========================================================================
REM Final da rotina 
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina que extrai o nome de um arquivo, removendo o caminho
REM -------------------------------------------------------------------------
Function ExtractFileName(arquivo)
  Dim fsa
  fsa = arquivo
  While InStr(fsa,"\") > 0
     fsa = Mid(fsa,InStr(fsa,"\")+1,Len(fsa))
  Wend
  While InStr(fsa,"/") > 0
     fsa = Mid(fsa,InStr(fsa,"/")+1,Len(fsa))
  Wend
  ExtractFileName = fsa
End Function
REM =========================================================================
REM Final da rotina 
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina de deleção de arquivos em disco
REM -------------------------------------------------------------------------
Sub DeleteAFile(filespec)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(filespec) Then
     fso.DeleteFile(filespec)
  End If
End Sub
REM =========================================================================
REM Final da rotina de deleção de arquivos em disco
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina de tratamento de erros
REM -------------------------------------------------------------------------
Sub TrataErro(p_query)
  If instr(Err.description,"CKH_") > 0 Then ' REGISTRO TEM FILHOS
    ScriptOpen "JavaScript"
    ShowHTML " alert('Não é permitido usar as palavras ""script"" ou "".js"" em nenhum campo.');"
    Err.clear
    ShowHTML " history.back(1);"
    ScriptClose
  Elseif instr(Err.description,"ORA-02291") > 0 Then ' REGISTRO NÃO ENCONTRADO
    Err.clear
    ScriptOpen "JavaScript"
    ShowHTML " alert('Registro não encontrado.');"
    ShowHTML " history.back(1);"
    ScriptClose
  Elseif instr(Err.description,"ORA-00001") > 0 Then ' REGISTRO JÁ EXISTENTE
    ScriptOpen "JavaScript"
    ShowHTML " alert('Um dos campos digitados já existe no banco de dados e é único.');"'\n\n" & Mid(Err.Description,1,Instr(Err.Description,Chr(10))-1) & "');"
    Err.clear
    ShowHTML " history.back(1);"
    ScriptClose
  Elseif instr(Err.description,"ORA-03113") > 0 _
      or instr(Err.description,"ORA-03114") > 0 _
      or instr(Err.description,"ORA-12224") > 0 _
      or instr(Err.description,"ORA-12514") > 0 _
      or instr(Err.description,"ORA-12541") > 0 _
      or instr(Err.description,"ORA-12545") > 0 _
    Then
    ScriptOpen "JavaScript"
    ShowHTML " alert('Banco de dados fora do ar. Aguarde alguns instantes e tente novamente!');"'\n\n" & Mid(Err.Description,1,Instr(Err.Description,Chr(10))-1) & "');"
    Err.clear
    ShowHTML " history.back(1);"
    ScriptClose
  Else
    Dim w_item, w_html
    w_html =  "<html><BASEFONT FACE=""Arial""><body BGCOLOR=""#FF5555"" TEXT=""#FFFFFF"">"
    w_html = w_html & "<CENTER><H2>ATENÇÃO</H2></CENTER>"
    w_html = w_html & "<BLOCKQUOTE>"
    w_html = w_html & "<P ALIGN=""JUSTIFY"">Erro não previsto. <b>Uma cópia desta tela foi enviada por e-mail para os responsáveis pela correção. Favor tentar novamente mais tarde.</P>"
    w_html = w_html & "<TABLE BORDER=""2"" BGCOLOR=""#FFCCCC"" CELLPADDING=""5""><TR><TD><FONT COLOR=""#000000"">" 
    w_html = w_html & "<DL><DT>Data e hora da ocorrência: <FONT FACE=""courier"">" & Mid(100+Day(date()),2,2) & "/" & Mid(100+Month(date()),2,2) & "/" & Year(date()) & ", " & time() & "<br><br></font></DT>" 
    w_html = w_html & "<DT>Query submetida: <FONT FACE=""courier""><br>" & replace(replace(p_query,VbCrLf,"<BR>")," ","&nbsp;") & "<br><br></font></DT>"
    w_html = w_html & "<DT>Número do erro: <FONT FACE=""courier"">" & Err.Number & "<br><br></font></DT>"
    w_html = w_html & "<DT>Descrição do erro:<DD><FONT FACE=""courier"">" & replace(Err.description,"SQL execution error, ","") & "<br><br></font>"
    w_html = w_html & "<DT>Identificado por:<DD><FONT FACE=""courier"">" & Err.Source & "<br><br></font>"
    w_html = w_html & "<DT>Dados submetidos:<DD><FONT FACE=""courier"" size=1>"
    For Each w_Item IN Request.QueryString 
        w_html = w_html & w_Item & " => [" & Request.QueryString(w_Item) & "]<br>"
    Next
    For Each w_Item IN Request.Form
        w_html = w_html & w_Item & " => [" & Request.Form(w_Item) & "]<br>"
    Next
    w_html = w_html & "   <br><br></font>"
    w_html = w_html & "</DT>"
    w_html = w_html & "<DT>Variáveis de servidor:<DD><FONT FACE=""courier"" size=1>"
    w_html = w_html & " SCRIPT_NAME => [" & Request.ServerVariables("SCRIPT_NAME") & "]<br>" 
    w_html = w_html & " SERVER_NAME => [" & Request.ServerVariables("SERVER_NAME") & "]<br>" 
    w_html = w_html & " SERVER_PORT => [" & Request.ServerVariables("SERVER_PORT") & "]<br>" 
    w_html = w_html & " SERVER_PROTOCOL => [" & Request.ServerVariables("SERVER_PROTOCOL") & "]<br>" 
    w_html = w_html & " HTTP_ACCEPT_LANGUAGE => [" & Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") & "]<br>" 
    w_html = w_html & " HTTP_USER_AGENT => [" & Request.ServerVariables("HTTP_USER_AGENT") & "]<br>" 
    w_html = w_html & "</DT>"
    w_html = w_html & "   <br><br></font>"
    w_html = w_html & "<DT>Variáveis de sessão:<DD><FONT FACE=""courier"" size=1>"
    For Each w_Item IN Session.Contents
        w_html = w_html & w_Item & " => [" & Session(w_Item) & "]<br>"
    Next
    w_html = w_html & "</DT>"
    w_html = w_html & "   <br><br></font>"
    w_html = w_html & "</FONT></TD></TR></TABLE><BLOCKQUOTE>"
    w_html = w_html & "</body></html>"
    ShowHTML w_HTML
    Err.clear
    EnviaMail "ERRO SIGE-WEB", w_html, "alex@sbpi.com.br"
  end if               
  Response.End()
End Sub
REM =========================================================================
REM Fim da rotina de tratamento de erros
REM -------------------------------------------------------------------------

REM =========================================================================
REM Rotina de cabeçalho
REM -------------------------------------------------------------------------
Sub Cabecalho
   ShowHTML "<HTML>"
end sub
REM -------------------------------------------------------------------------
REM Final da rotina de cabeçalho
REM =========================================================================

REM =========================================================================
REM Rotina de rodapé
REM -------------------------------------------------------------------------
Sub Rodape
   ShowHTML "<HR>"
   'ShowHTML "<CENTER><FONT FACE=""ARIAL"" SIZE=1>Um produto da <A HREF=""http://www.sbpi.com.br"" TARGET=""_BLANK"">SBPI&reg;</A> Sociedade Brasileira para a Pesquisa em Informática.<br>Dúvidas, sugest&otilde;es ou problemas: encaminhar mensagem para <A HREF=""mailto:sbpi@sbpi.com.br""><img src=images/icone/mailto.gif border=0> <i>sbpi@sbpi.com.br</i></A><br>1996-2001 <A HREF=""http://www.sbpi.com.br"" TARGET=""_BLANK"">SBPI&reg;</a> todos os direitos reservados.</font></center"
   ShowHTML "</BODY>"
   ShowHTML "</HTML>"
end sub
REM -------------------------------------------------------------------------
REM Final da rotina de rodapé
REM =========================================================================

REM =========================================================================
REM Abre conexão com o banco de dados
REM -------------------------------------------------------------------------
Sub AbreSessao
   Set dbms = Server.CreateObject("ADODB.Connection")
   with dbms
      .ConnectionString = conConnectionString
      .Open
      .CursorLocation = adUseClient
   end with
   Set SP               = Server.CreateObject("ADODB.Command")
   Set RS               = Server.CreateObject("ADODB.RecordSet")
   sp.ActiveConnection  = dbms
   sp.CommandType       = adCmdStoredProc
end sub

Sub AbreSessaoManut
   Set dbms = Server.CreateObject("ADODB.Connection")
   with dbms
      .ConnectionString = conConnectionManut
      .Open
      .CursorLocation = adUseClient
   end with
   Set SP               = Server.CreateObject("ADODB.Command")
   Set RS               = Server.CreateObject("ADODB.RecordSet")
   sp.ActiveConnection  = dbms
   sp.CommandType       = adCmdStoredProc
end sub
REM -------------------------------------------------------------------------
REM Final do procedimento que obtém sessão do pool de conexões
REM =========================================================================

REM =========================================================================
REM Fecha conexão com o banco de dados
REM -------------------------------------------------------------------------
Sub FechaSessao
   dbms.close
end sub
REM -------------------------------------------------------------------------
REM Final do procedimento que obtém sessão do pool de conexões
REM =========================================================================

REM =========================================================================
REM Cria objeto gráfico
REM -------------------------------------------------------------------------
Sub CriaGrafico
   Set graph = Server.CreateObject("GraphLib.Graph")
end sub
REM -------------------------------------------------------------------------
REM Final da rotina
REM =========================================================================

REM =========================================================================
REM Cria parâmetro apenas para OLE/DB da Oracle
REM -------------------------------------------------------------------------
Sub SetProperty (state)
    sp.Properties("PLSQLRSet")= state
end sub
REM -------------------------------------------------------------------------
REM Final do procedimento que obtém sessão do pool de conexões
REM =========================================================================

REM =========================================================================
REM Rotina de execução de queries
REM -------------------------------------------------------------------------
Sub ExecutaSQL(p_SQL)
   On Error Resume Next
   If instr(uCase(p_SQL),"SCRIPT")>0 or instr(uCase(p_SQL),".JS")>0 Then ' CÓDIGO MALICIOSO
     ScriptOpen "JavaScript"
     ShowHTML " alert('Não é permitido usar as palavras ""script"" ou "".js"" em nenhum campo.');"
     Err.clear
     ShowHTML " history.back(1);"
     ScriptClose
     Response.End()
   End If
   dbms.Execute(p_sql)
   If Err.Number <> 0 Then 
      'dbms.RollbackTrans()
      TrataErro p_sql
      Response.End()
      Exit Sub
   End If
end sub
REM -------------------------------------------------------------------------
REM Final da rotina de execução de queries 
REM =========================================================================

REM =========================================================================
REM Rotina de execução de queries sem indicação de erro
REM -------------------------------------------------------------------------
Sub ExecutaSQL_Resume(p_SQL)
   On Error Resume Next
   dbms.Execute(p_sql)
end sub
REM -------------------------------------------------------------------------
REM Final da rotina de execução de queries sem indicação de erro
REM =========================================================================

REM =========================================================================
REM Rotina de Abertura do BD para execução de queries
REM -------------------------------------------------------------------------
Sub ConectaBD(p_Query)
   On Error Resume Next
   If RS.State = 1 Then 
      RS.Close 
   End If
   RS.Open p_Query, dbms, adOpenStatic
   If Err.number <> 0 Then
      TrataErro(p_query)
      Err.Clear
      Response.End()
      Exit Sub
   End If
end sub
REM -------------------------------------------------------------------------
REM Final do procedimento de abertura do BD
REM =========================================================================

REM =========================================================================
REM Rotina de Abertura do BD para execução de queries read only
REM -------------------------------------------------------------------------
Sub ConectaBDLock(p_Query)
   On Error Resume Next
   If RS.State = 1 Then 
      RS.Close 
   End If
   RS.Open sql, sobjConn, adOpenForwardOnly
   If Err.number <> 0 Then
      TrataErro(p_query)
      Err.Clear
      Response.End()
      Exit Sub
   End If
end sub
REM -------------------------------------------------------------------------
REM Final do procedimento de abertura do BD
REM =========================================================================

REM =========================================================================
REM Rotina de Fechamento do BD
REM -------------------------------------------------------------------------
Sub DesConectaBD
   If RS.State = 0 Then 
      RS.Close 
   End If
end sub
REM -------------------------------------------------------------------------
REM Final do procedimento de fechamento do BD
REM =========================================================================

REM -------------------------------------------------------------------------
REM Verifica se a Assinatura Eletronica do usuário está correta
REM =========================================================================
Function VerificaAssinaturaEletronica(Usuario,Senha)

   If Senha > "" Then
      If DB_VerificaAssinatura(Session("p_cliente"), Usuario, Senha) = 0 Then
         VerificaAssinaturaEletronica = True
      Else
         VerificaAssinaturaEletronica = False
      End If
   Else
      VerificaAssinaturaEletronica = True
   End If

End Function
REM -------------------------------------------------------------------------
REM Fim da verificação se a Assinatura Eletronica do usuário está correta
REM =========================================================================

REM -------------------------------------------------------------------------
REM Verifica se a senha de acesso do usuário está correta
REM =========================================================================
Function VerificaSenhaAcesso(Usuario,Senha)

   If DB_VerificaSenha(Session("p_cliente"), Usuario, Senha) = 0 Then
      VerificaSenhaAcesso = True
   Else
      VerificaSenhaAcesso = False
   End If

End Function
REM -------------------------------------------------------------------------
REM Fim da verificação se a senha de acesso do usuário está correta
REM =========================================================================

REM =========================================================================
REM Função que formata dias, horas, minutos e segundos a partir dos segundos
REM -------------------------------------------------------------------------
Function FormataDataEdicao(w_dt_grade)
  Dim l_dt_grade
  l_dt_grade = w_dt_grade
  If l_dt_grade > "" Then
     If Len(l_dt_grade) < 10 Then
        If Right(Mid(l_dt_grade,1,2),1) = "/" Then
           l_dt_grade = "0"&l_dt_grade
        End If
        If Len(l_dt_grade) < 10 and Right(Mid(l_dt_grade,4,2),1) = "/" Then
           l_dt_grade = Left(l_dt_grade,3)&"0"&Right(l_dt_grade,6)
        End If 
     End If
  Else
     l_dt_grade = ""
  End If

  FormataDataEdicao = l_dt_grade

  Set l_dt_grade       = Nothing
End Function
REM =========================================================================
REM Final da função
REM -------------------------------------------------------------------------



'Limpa Mascara para gravar os dados no banco de dados
Function LimpaMascara(Campo)
   LimpaMascara = replace(replace(replace(replace(replace(replace(campo,",",""),";",""),".",""),"-",""),"/",""),"\","")
End Function

' Cria a tag Body
Sub BodyOpen(cProperties)

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
ShowHtml " .STB {font-size: 8pt; color: #000000; border: 1pt solid #000000; background-color: #C0C0C0; }"  & VbCrLf
ShowHtml " .STS {font-size: 8pt; border-top: 1px solid #000000; background-color: #F5F5F5}"  & VbCrLf
ShowHtml " .STR {font-size: 8pt; border-top: 0px}"  & VbCrLf
ShowHtml " .STC {font-size: 8pt; border-top: 0px}"  & VbCrLf
ShowHTML "</style> "
ShowHTML "<STYLE TYPE=""text/css"">"
ShowHTML "<!--"
ShowHTML "BODY {OVERFLOW:scroll;OVERFLOW-X:hidden}"
ShowHTML ".DEK {POSITION:absolute;VISIBILITY:hidden;Z-INDEX:200;}"
ShowHTML "//-->"
ShowHTML "</STYLE>"
ShowHTML "<body Text=""" & conBodyText & """ Link=""" & conBodyLink & """ Alink=""" & conBodyALink & """ " & _
    "Vlink=""" & conBodyVLink & """ Bgcolor=""" & conBodyBgcolor & """ Background=""" & conBodyBackground & """ " & _
    "Bgproperties=""" & conBodyBgproperties & """ Topmargin=""" & conBodyTopmargin & """ " & _
    "Leftmargin=""" & conBodyLeftmargin & """ " & cProperties & "> " 
ShowHTML "<DIV ID=""dek"" CLASS=""dek""></DIV>"
ShowHTML "<SCRIPT TYPE=""text/javascript"">"
ShowHTML "<!--"
ShowHTML "Xoffset=-100;    // modify these values to ..."
ShowHTML "Yoffset= 20;    // change the popup position."
ShowHTML "var nav,old,iex=(document.all),yyy=-1000;"
ShowHTML "if(navigator.appName==""Netscape""){(document.layers)?nav=true:old=true;}"
ShowHTML "if(!old){"
ShowHTML "var skn=(nav)?document.dek:dek.style;"
ShowHTML "if(nav)document.captureEvents(Event.MOUSEMOVE);"
ShowHTML "document.onmousemove=get_mouse;"
ShowHTML "}"
ShowHTML "function popup(msg,bak){"
ShowHTML "var content='<TABLE  WIDTH=200 BORDER=1 BORDERCOLOR=black CELLPADDING=2 CELLSPACING=0 BGCOLOR='+bak+'><TD><DIV ALIGN=""JUSTIFY""><FONT COLOR=black SIZE=1>'+msg+'</FONT></DIV></TD></TABLE>';"
ShowHTML "if(old){alert(msg);return;} "
ShowHTML "else{yyy=Yoffset;"
ShowHTML " if(nav){skn.document.write(content);skn.document.close();skn.visibility=""visible""}"
ShowHTML " if(iex){document.all(""dek"").innerHTML=content;skn.visibility=""visible""}"
ShowHTML " }"
ShowHTML "}"
ShowHTML "function popup1(msg,bak){"
ShowHTML "var content='<TABLE  WIDTH=450 BORDER=1 BORDERCOLOR=black CELLPADDING=2 CELLSPACING=0 BGCOLOR='+bak+'><TD><DIV ALIGN=""JUSTIFY""><FONT COLOR=black SIZE=1>'+msg+'</FONT></DIV></TD></TABLE>';"
ShowHTML "if(old){alert(msg);return;} "
ShowHTML "else{yyy=Yoffset;"
ShowHTML " if(nav){skn.document.write(content);skn.document.close();skn.visibility=""visible""}"
ShowHTML " if(iex){document.all(""dek"").innerHTML=content;skn.visibility=""visible""}"
ShowHTML " }"
ShowHTML "}"
ShowHTML "function get_mouse(e){"
ShowHTML "var x=(nav)?e.pageX:event.x+document.body.scrollLeft;skn.left=x+Xoffset;"
ShowHTML "var y=(nav)?e.pageY:event.y+document.body.scrollTop;skn.top=y+yyy;"
ShowHTML "}"
ShowHTML "function kill(){"
ShowHTML "if(!old){yyy=-1000;skn.visibility=""hidden"";}"
ShowHTML "}"
ShowHTML "//-->"
ShowHTML "</SCRIPT>"

End Sub

Sub BodyOpenImage(cProperties, cImage, cFixed)
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
ShowHtml " .STB {font-size: 8pt; color: #000000; border: 1pt solid #000000; background-color: #C0C0C0; }"  & VbCrLf
ShowHtml " .STS {font-size: 8pt; border-top: 1px solid #000000; background-color: #F5F5F5}"  & VbCrLf
ShowHtml " .STR {font-size: 8pt; border-top: 0px}"  & VbCrLf
ShowHtml " .STC {font-size: 8pt; border-top: 0px}"  & VbCrLf
ShowHTML "</style> "
ShowHTML "<STYLE TYPE=""text/css"">"
ShowHTML "<!--"
ShowHTML "BODY {OVERFLOW:scroll;OVERFLOW-X:hidden}"
ShowHTML ".DEK {POSITION:absolute;VISIBILITY:hidden;Z-INDEX:200;}"
ShowHTML "//-->"
ShowHTML "</STYLE>"
ShowHTML "<body Text=""" & conBodyText & """ Link=""" & conBodyLink & """ Alink=""" & conBodyALink & """ " & _
    "Vlink=""" & conBodyVLink & """ Bgcolor=""" & conBodyBgcolor & """ Background=""" & cImage & """ " & _
    "Bgproperties=""" & cFixed & """ Topmargin=""" & conBodyTopmargin & """ " & _
    "Leftmargin=""" & conBodyLeftmargin & """ " & cProperties & "> " 
ShowHTML "<DIV ID=""dek"" CLASS=""dek""></DIV>"
ShowHTML "<SCRIPT TYPE=""text/javascript"">"
ShowHTML "<!--"
ShowHTML "Xoffset=-100;    // modify these values to ..."
ShowHTML "Yoffset= 20;    // change the popup position."
ShowHTML "var nav,old,iex=(document.all),yyy=-1000;"
ShowHTML "if(navigator.appName==""Netscape""){(document.layers)?nav=true:old=true;}"
ShowHTML "if(!old){"
ShowHTML "var skn=(nav)?document.dek:dek.style;"
ShowHTML "if(nav)document.captureEvents(Event.MOUSEMOVE);"
ShowHTML "document.onmousemove=get_mouse;"
ShowHTML "}"
ShowHTML "function popup(msg,bak){"
ShowHTML "var content='<TABLE  WIDTH=200 BORDER=1 BORDERCOLOR=black CELLPADDING=2 CELLSPACING=0 BGCOLOR='+bak+'><TD><DIV ALIGN=""JUSTIFY""><FONT COLOR=black SIZE=1>'+msg+'</FONT></DIV></TD></TABLE>';"
ShowHTML "if(old){alert(msg);return;} "
ShowHTML "else{yyy=Yoffset;"
ShowHTML " if(nav){skn.document.write(content);skn.document.close();skn.visibility=""visible""}"
ShowHTML " if(iex){document.all(""dek"").innerHTML=content;skn.visibility=""visible""}"
ShowHTML " }"
ShowHTML "}"
ShowHTML "function popup1(msg,bak){"
ShowHTML "var content='<TABLE  WIDTH=450 BORDER=1 BORDERCOLOR=black CELLPADDING=2 CELLSPACING=0 BGCOLOR='+bak+'><TD><DIV ALIGN=""JUSTIFY""><FONT COLOR=black SIZE=1>'+msg+'</FONT></DIV></TD></TABLE>';"
ShowHTML "if(old){alert(msg);return;} "
ShowHTML "else{yyy=Yoffset;"
ShowHTML " if(nav){skn.document.write(content);skn.document.close();skn.visibility=""visible""}"
ShowHTML " if(iex){document.all(""dek"").innerHTML=content;skn.visibility=""visible""}"
ShowHTML " }"
ShowHTML "}"
ShowHTML "function get_mouse(e){"
ShowHTML "var x=(nav)?e.pageX:event.x+document.body.scrollLeft;skn.left=x+Xoffset;"
ShowHTML "var y=(nav)?e.pageY:event.y+document.body.scrollTop;skn.top=y+yyy;"
ShowHTML "}"
ShowHTML "function kill(){"
ShowHTML "if(!old){yyy=-1000;skn.visibility=""hidden"";}"
ShowHTML "}"
ShowHTML "//-->"
ShowHTML "</SCRIPT>"

End Sub

' Imprime uma linha HTML
Sub ShowHtml(Line)
  Response.Write Line & CHR(13) & CHR(10)
End Sub
%>

