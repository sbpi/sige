<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/Constants_ADO.inc"-->
<!--#INCLUDE VIRTUAL="/modelos/Constants.inc"-->
<!--#INCLUDE VIRTUAL="/esc.inc"-->
<!--#INCLUDE VIRTUAL="/Funcoes.asp"-->
<%
Response.Expires = 0
REM =========================================================================
REM  /Aluno.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: Exibição dos dados do aluno
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 23/03/2000 16:55
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Informática
REM -------------------------------------------------------------------------

Private RS, RS1, RS2
Private RSCabecalho
Private RSRodape

Private sobjConn

Private sstrSN
Private sstrCL
Private w_ano_letivo

sstrSN = "Default.Asp"
sstrCL = "Aluno.Asp"
w_dir  = "Modelos/Mod16/"

Public sstrEA
Public sstrEF
Public sstrEW
Public sstrIN
Public CL
Dim SQL

Set RS = Server.CreateObject("ADODB.RecordSet")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
Set RS2 = Server.CreateObject("ADODB.RecordSet")
Set RSCabecalho = Server.CreateObject("ADODB.RecordSet")
Set RSRodape = Server.CreateObject("ADODB.RecordSet")
Set sobjConn  = Server.CreateObject("ADODB.Connection")

Server.ScriptTimeOut = Session("ScriptTimeOut")

sobjConn.ConnectionTimeout = 300
sobjConn.CommandTimeout = 300
sobjConn.Open conConnectionString

CL = Request.QueryString("CL")
sstrEA = Request("EA")
sstrEF = "sq_cliente=" & CL
sstrEW = Request("EW")
sstrIN = Request("IN")
If Request("w_ano_letivo") > "" Then
   w_ano_letivo = Request("w_ano_letivo")
Else
   sql = "SELECT max(ano_letivo) AS ano_letivo FROM escAluno_Turma WHERE " & sstrEA
   RS.Open sql, sobjConn, adOpenForwardOnly
   If IsNull(RS("ano_letivo")) Then
      w_ano_letivo = Year(Date())
   Else
      w_ano_letivo = RS("ano_letivo")
   End If
   RS.Close()
End If


ShowHTML "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
ShowHTML "<html xmlns=""http://www.w3.org/1999/xhtml"">"
ShowHTML "<head>"
ShowHTML "   <title>Secretaria de Estado de Educa&ccedil;&atilde;o</title>"
ShowHTML "   <meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> "
If sstrEW <> "conWhatExBoletimImp" Then  
   ShowHTML "   <link href=""/css/estilo.css"" media=""screen"" rel=""stylesheet"" type=""text/css"" />"
   ShowHTML "   <link href=""/css/print.css""  media=""print""  rel=""stylesheet"" type=""text/css"" />"
   ShowHTML "   <script language=""javascript"" src=""inc/scripts.js""> </script>"
End If
ShowHTML "</head>"

ShowHTML "<BASE HREF=""" & conSite & "/"">"
ShowHTML "<body>"
If sstrEW <> "conWhatExBoletimImp" Then  
   ShowHTML "<div id=""container"">"
   ShowHTML "  <div id=""cab"">" 
   ShowHTML "    <div id=""cabtopo"">"
   ShowHTML "      <div id=""logoesq""><img src=""img/fundo_logoesq.gif"" border=0></div>"
   ShowHTML "      <div id=""logodir""><a href=""http://www.se.df.gov.br""><img src=""img/fundo_logodir.jpg"" border=0></a></div>"   
   ShowHTML "    </div>"
   
   sql = "SELECT a.sq_cliente, a.ds_cliente, b.ds_mensagem, c.no_aluno, c.nr_matricula " & VbCrLf & _
         "  FROM escCliente a " & VbCrLf & _
         "       inner join escCliente_site b on (a.sq_cliente = b.sq_cliente) " & VbCrLf & _
         "       inner join escAluno        c on (b.sq_site_cliente = c.sq_site_cliente) " & VbCrLf & _
         " WHERE c." & sstrEA
   RS.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly
   ShowHTML "    <div id=""cabbase"">"
   ShowHTML "      <div id=""busca"" valign=""center""><marquee width=""100%"" align=""middle""><font color=""white"" face=""Arial"" size=""2""><b>" & server.HTMLEncode(RS("ds_mensagem")) & "</b></font></marquee></div> "
   ShowHTML "    </div>"
   ShowHTML "  </div>"
   ShowHTML "  <div id=""corpo"">"
   ShowHTML "    <div id=""menuesq"">"
   ShowHTML "      <div id=""logomenuesq""><img src=""img/fundo_logomenuesq.gif"" border=0></div>"
   ShowHTML "      <ul id=""menugov"">"
   ShowHTML "      <script language=""JavaScript"" src=""inc/mm_menu.js"" type=""text/JavaScript""></script>"
   ShowHTML "      <li><a href=""" & w_dir & sstrCL & "?EW=118&CL=" & CL & "&EA=" & sstrEA & """ >Inicial</a> </li>"
   ShowHTML "      <li><a href=""" & w_dir & sstrCL & "?EW=119&CL=" & CL & "&EA=" & sstrEA & """ id=""link1"">Troca de senha</a> </li>"
   ShowHTML "      <li><a href=""" & w_dir & sstrCL & "?EW=120&CL=" & CL & "&EA=" & sstrEA & """ >Boletim </a></li>"
   ShowHTML "      <li><a href=""" & w_dir & sstrCL & "?EW=121&CL=" & CL & "&EA=" & sstrEA & """ id=""link2"">Grade horária</a> </li>"
   ShowHTML "      <li><a href=""" & w_dir & sstrCL & "?EW=122&CL=" & CL & "&EA=" & sstrEA & """ id=""link5"">Mensagens</a>"
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
   ShowHTML "      <li><b>" & RS("no_aluno") & " (" & RS("nr_matricula") & ")</b></li>"
   Select Case sstrEW
     Case conWhatMenu        ShowHTML "      <li><b>Inicial</b></li>"
     Case conWhatExBoletim   ShowHTML "      <li><b>Boletim</b></li>"
     Case conWhatExGrade     ShowHTML "      <li><b>Grade horária</b></li>"
     Case conWhatSenha       ShowHTML "      <li><b>Troca de senha</b></li>"
     Case conWhatExMens      ShowHTML "      <li><b>Mensagens</b></li>"
   End Select

   ShowHTML "	    </ul>"
   ShowHTML "	  </div>"
   RS.Close
End If
ShowHTML "    <div id=""texto""><!-- Conteúdo -->"
ShowHTML "        <table width=""570"" border=""0"">"

Main

ShowHTML "        </table>"
If sstrEW <> "conWhatExBoletimImp" Then  
   ShowHTML "    </div>"
   ShowHTML "  </div> "
End If

ShowHTML "  <br clear=""all"" />"
ShowHTML "</div>"

If sstrEW <> "conWhatExBoletimImp" Then
   ShowHTML "<div id=""rodape"">"
   ShowHTML "  <div id=""endereco"">"
   ShowHTML "      <div></div>"
   ShowHTML "      <div></div>"
   ShowHTML "      <div><br /></div> "
   ShowHTML "  </div>"
   ShowHTML "</div>"
End If
ShowHTML "</body>"
ShowHTML "</html>"

Set RS = nothing
Set RSCabecalho = nothing
Set RSRodape = nothing
Set sobjConn  = nothing

REM =========================================================================
REM Troca nulos por hífen na exibição do boletim
REM -------------------------------------------------------------------------
Function TrocaNulo(valor)


   if trim(valor) > "" then 
      TrocaNulo = valor
   else TrocaNulo = "-"
   end if
   
End Function
REM -------------------------------------------------------------------------
REM Final da Function TrocaNulo
REM =========================================================================

REM =========================================================================
REM Monta a tela principal
REM -------------------------------------------------------------------------
Public Sub ShowMenu

  Dim sql, strNome

  ShowHTML "<tr><td>"

  ShowHTML "<p align=""justify""><b>Instruções:</b></p>"
  ShowHTML "<ul class=""texto"">"
  ShowHTML "<li>Selecione a opção desejada no menu acima para ver o boletim do aluno, sua grade horária ou histórico de mensagens."
  ShowHTML "<li>Abaixo poderão aparecer mensagens dirigidas pela escola ou pela rede de ensino. Leia-a atentamente e clique sobre <i>[Arquivar]</i> para enviá-la ao histórico de mensagens."
  ShowHTML "<li>Suas informações cadastrais também serão exibidas. Qualquer informação incompleta ou incorreta pode ser alterada junto à secretaria de sua escola."
  ShowHTML "<li>Se o conteúdo da página for maior que sua altura máxima, use a barra de rolagem, à direita, para subir ou descer o conteúdo."
  ShowHTML "<li>Se você tiver uma impressora disponível e quiser imprimir o conteúdo de qualquer página, clique com o botão direito do <i>mouse</i> e, em seguida, sobre a opção <i>Imprimir</i>."
  ShowHTML "</ul>"

  ' Exibe eventuais mensagens dirigidas ao aluno
  sql = "SELECT a.no_aluno,b.* " & _
        "FROM escAluno AS a LEFT OUTER JOIN escMensagem_aluno AS b ON (a.sq_aluno = b.sq_aluno) " & _
        "WHERE substring(b.in_lida,1,1) = 'N'" & _
        "  AND a." & sstrEA & " " & _
        "ORDER BY b.DT_MENSAGEM"

  RS.Open sql, sobjConn, adOpenForwardOnly

  If Not RS.EOF Then

    ShowHTML "<br><div align=""justify""><b>Novas mensagens:</b><hr></div>"

    ShowHTML "<TABLE BORDER=0 WIDTH=""100%"">"
    Do While Not RS.EOF

      If RS("sq_mensagem") > "" Then
  	     ShowHTML "<TR valign=""top""><TD width=""2%"">::<TD width=""98%""><div align=justify>" & RS("ds_mensagem") & " (" & Mid(100+Day(RS("dt_mensagem")),2,2) & "/" & Mid(100+Month(RS("dt_mensagem")),2,2) & "/" &Year(RS("dt_mensagem")) & ")" & "&nbsp;&nbsp;&nbsp;<a href=""" & w_dir & sstrCL & "?EW=" & conWhatGravaMens & "&EF=" & sstrEF & "&CL=" & CL & "&EA=" & sstrEA & "&IN=sq_mensagem=" & RS("sq_mensagem") & """><b><font size=""2"">[Arquivar]</font></b></a></div></font></TD></TR>"
  	  End If

      RS.MoveNext

    Loop
    ShowHTML "</TABLE>"


  End If

  RS.Close

  ' Recupera os dados cadastrais do aluno
  sql = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, a.dt_nascimento,  " & _
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end as in_sexo, " & _
        "       a.ds_naturalidade, a.no_mae, a.nr_fone_mae, a.no_pai, a.nr_fone_pai, " & _
        "       a.no_resposavel, a.nr_fone_responsavel, a.ds_email_responsavel, nr_fone_1, nr_fone_2 " & _
        "FROM escAluno AS a " & _
        "WHERE a." & sstrEA 

  RS.Open sql, sobjConn, adOpenForwardOnly

  If Not RS.EOF Then

    ShowHTML "<b>Informações cadastrais:</b><hr>"
    ShowHTML "<TABLE align=""center"" WIDTH=""100%"" BORDER=0>"
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><B>Nome:</b><BR>" & RS("NO_ALUNO") & "</TD>"
    ShowHTML "    <TD><B>Matrícula:</b><BR>" & RS("NR_MATRICULA") & "</TD>"
    ShowHTML "  </TR>"

    ' Recupera os dados cadastrais do aluno
    sql = "SELECT c.ano_letivo, c.st_aluno, " & _
          "       b.ds_grau, b.ds_serie, b.ds_turma, replace(b.ds_curso,'¬','ª') ds_curso, " & _
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end as ds_turno " & _
          "FROM escAluno_Turma      AS c " & _
          "     INNER JOIN escTurma AS b ON (b.sq_turma = c.sq_turma and b.sq_site_cliente = c.sq_site_cliente and c.ano_letivo = " & w_ano_letivo & ") " & _
          "WHERE c.sq_aluno = " & RS("SQ_ALUNO")
    RS1.Open sql, sobjConn, adOpenForwardOnly
    ShowHTML "  <TR valign=""TOP""><TD COLSPAN=""3""><TABLE WIDTH=""100%"" BORDER=""1"">"
    ShowHTML "    <TR valign=""TOP""><TD COLSPAN=""6""><B>Turma(s):"
    ShowHTML "    <TR valign=""TOP"" align=""CENTER"">"
    ShowHTML "    <TD><B>Ano Letivo</b></TD>"
    ShowHTML "    <TD><B>Série</b></TD>"
    ShowHTML "    <TD><B>Turno</b></TD>"
    ShowHTML "    <TD><B>Turma</b></TD>"
    ShowHTML "    <TD><B>Etapa/Curso/Modalidade</b></TD>"
    ShowHTML "    <TD><B>Situação</b></TD>"
    Do While Not RS1.EOF
      ShowHTML "  <TR valign=""TOP"">"
      'ShowHTML "    <TD><B>Grau:</b><BR>" & RS1("DS_GRAU") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER"">" & RS1("ANO_LETIVO") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER"">" & RS1("DS_SERIE") & "</TD>"
      ShowHTML "    <TD>" & RS1("DS_TURNO") & "</TD>"
      ShowHTML "    <TD>" & RS1("DS_TURMA") & "</TD>"
      ShowHTML "    <TD>"
      If Trim(RS1("DS_CURSO")) > "" Then ShowHTML RS1("DS_CURSO") Else ShowHTML "-" End If
      ShowHTML "    <TD>" & RS1("ST_ALUNO") & "</TD>"
      RS1.MoveNext
    Loop
    RS1.Close
    ShowHTML "  </TR></TABLE>"

    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><B>Data de nascimento:</b><BR>" 
    If Trim(RS("DT_NASCIMENTO")) > "" Then ShowHTML RS("DT_NASCIMENTO") Else ShowHTML "-" End If
    ShowHTML "    <TD><B>Sexo:</b><BR>" & RS("IN_SEXO") & "</TD>"
    ShowHTML "    <TD><B>Naturalidade:</b><BR>"
    If Trim(RS("DS_NATURALIDADE")) > "" Then ShowHTML RS("DS_NATURALIDADE") Else ShowHTML "-" End If
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><B>Mãe:</b><BR>"
    If Trim(RS("NO_MAE")) > "" Then ShowHTML RS("NO_MAE") Else ShowHTML "-" End If
    ShowHTML "    <TD><B>Telefone:</b><BR>"
    If Trim(RS("NR_FONE_MAE")) > "" Then ShowHTML RS("NR_FONE_MAE") Else ShowHTML "-" End If
    ShowHTML "  </TR>"
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><B>Pai:</b><BR>"
    If Trim(RS("NO_PAI")) > "" Then ShowHTML RS("NO_PAI") Else ShowHTML "-" End If
    ShowHTML "    <TD><B>Telefone:</b><BR>"
    If Trim(RS("NR_FONE_PAI")) > "" Then ShowHTML RS("NR_FONE_PAI") Else ShowHTML "-" End If
    ShowHTML "  </TR>"
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><B>Responsável pelo aluno:</b><BR>"
    If Trim(RS("NO_RESPOSAVEL")) > "" Then ShowHTML RS("NO_RESPOSAVEL") Else ShowHTML "-" End If
    ShowHTML "    <TD><B>Telefone:</b><BR>"
    If Trim(RS("NR_FONE_RESPONSAVEL")) > "" Then ShowHTML RS("NR_FONE_RESPONSAVEL") Else ShowHTML "-" End If
    ShowHTML "    <TD><B>e-Mail:</b><BR>"
    If Trim(RS("DS_EMAIL_RESPONSAVEL")) > "" Then ShowHTML RS("DS_EMAIL_RESPONSAVEL") Else ShowHTML "-" End If
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><B>Outros telefones do aluno:</b></TD>"
    ShowHTML "    <TD><B>Telefone 1:</b><BR>"
    If Trim(RS("NR_FONE_1")) > "" Then ShowHTML RS("NR_FONE_1") Else ShowHTML "-" End If
    ShowHTML "    <TD><B>Telefone 2:</b><BR>"
    If Trim(RS("NR_FONE_2")) > "" Then ShowHTML RS("NR_FONE_2") Else ShowHTML "-" End If
    ShowHTML "</TABLE></CENTER>"

  End If

  RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Página Principal
REM =========================================================================

REM =========================================================================
REM Monta a tela de troca de senha
REM -------------------------------------------------------------------------
Public Sub ShowSenha

  Dim sql, strNome

  ShowHTML "<tr><td><p align=""justify""><b>Instruções:</b></p>"
  ShowHTML "<ul>"
  ShowHTML "<li><div align=""justify"">Informe abaixo sua senha atual e a nova senha.</div>"
  ShowHTML "<li><div align=""justify"">A nova senha deve ser digitada nos campos ""Informe a nova senha"" e ""Redigite a nova senha"", para evitar erros de digitação. Deve ter pelo menos 6 letras ou números, não sendo aceitos espaços em branco, ponto, traço etc.</div>"
  ShowHTML "<li><div align=""justify"">Se ocorrer algum erro, tal como senha atual inválida ou valores diferentes nos campos relativos à nova senha, você receberá uma mensagem de alerta.</div>"
  ShowHTML "<li><div align=""justify"">Se os dados informados estiverem corretos, a senha atual será substituída pela nova senha e você receberá uma mensagem de confirmação.</div>"
  ShowHTML "<li><div align=""justify"">A nova senha entra em vigor assim que a troca for concluída. Assim, novos acessos à sua página pessoal devem ser feitos usando a nova senha.</div>"
  ShowHTML "</ul>"

  ShowHTML "<div align=""justify""><b>Troca da senha de acesso:</b><hr></div>" & VbCrLf

  ShowHTML "<script Language=""JavaScript"">" & chr(13)
  ShowHTML "<!--" & chr(13)
  ShowHTML "function Validacao(theForm)" & chr(13)
  ShowHTML "{" & chr(13)
  ShowHTML "  if (theForm.senha_atual.value == '')" & chr(13)
  ShowHTML "  {" & chr(13)
  ShowHTML "    alert(""ERRO!!!\nFavor informar um valor para o campo \""Senha atual\""."");" & chr(13)
  ShowHTML "    theForm.senha_atual.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)
  ShowHTML "  if (theForm.senha_atual.value.length < 5)" & chr(13)
  ShowHTML "  {" & chr(13)
  ShowHTML "    alert(""ERRO!!!\nFavor informar pelo menos 5 letras ou números no campo \""Senha atual\""."");" & chr(13)
  ShowHTML "    theForm.senha_atual.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)
  ShowHTML "  var checkOK = ""ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-"";" & chr(13)
  ShowHTML "  var checkStr = theForm.senha_atual.value;" & chr(13)
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
  ShowHTML "    alert(""ERRO!!!\nInforme apenas letras e números no campo \""Senha atual\"", sem espaços em branco."");" & chr(13)
  ShowHTML "    theForm.senha_atual.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)

  ShowHTML "  if (theForm.senha_nova.value == '')" & chr(13)
  ShowHTML "  {" & chr(13)
  ShowHTML "    alert(""ERRO!!!\nFavor informar um valor para o campo \""Nova senha\""."");" & chr(13)
  ShowHTML "    theForm.senha_nova.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)
  ShowHTML "  if (theForm.senha_nova.value.length < 5)" & chr(13)
  ShowHTML "  {" & chr(13)
  ShowHTML "    alert(""ERRO!!!\nFavor informar pelo menos 5 letras ou números no campo \""Nova senha\""."");" & chr(13)
  ShowHTML "    theForm.senha_nova.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)
  ShowHTML "  var checkOK = ""ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"";" & chr(13)
  ShowHTML "  var checkStr = theForm.senha_nova.value;" & chr(13)
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
  ShowHTML "    alert(""ERRO!!!\nInforme apenas letras e números no campo \""Nova senha\"", sem espaços em branco."");" & chr(13)
  ShowHTML "    theForm.senha_nova.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)

  ShowHTML "  if (theForm.senha_nova1.value == '')" & chr(13)
  ShowHTML "  {" & chr(13)
  ShowHTML "    alert(""ERRO!!!\nFavor informar um valor para o campo \""Redigite a nova senha\""."");" & chr(13)
  ShowHTML "    theForm.senha_nova1.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)
  ShowHTML "  if (theForm.senha_nova1.value.length < 5)" & chr(13)
  ShowHTML "  {" & chr(13)
  ShowHTML "    alert(""ERRO!!!\nFavor informar pelo menos 5 letras ou números no campo \""Redigite a nova senha\""."");" & chr(13)
  ShowHTML "    theForm.senha_nova1.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)
  ShowHTML "  var checkOK = ""ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"";" & chr(13)
  ShowHTML "  var checkStr = theForm.senha_nova1.value;" & chr(13)
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
  ShowHTML "    alert(""ERRO!!!\nInforme apenas letras e números no campo \""Redigite a nova senha\"", sem espaços em branco."");" & chr(13)
  ShowHTML "    theForm.senha_nova1.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)

  ShowHTML "  if (theForm.senha_nova.value != theForm.senha_nova1.value)" & chr(13)
  ShowHTML "  {" & chr(13)
  ShowHTML "    alert(""ERRO!!!\nOs campos \""Nova senha\"" e \""Redigite a nova senha\"" devem ter o mesmo valor."");" & chr(13)
  ShowHTML "    theForm.senha_nova1.value='';" & chr(13)
  ShowHTML "    theForm.senha_nova.value='';" & chr(13)
  ShowHTML "    theForm.senha_nova.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)

  ShowHTML "  if (theForm.senha_atual.value == theForm.senha_nova.value)" & chr(13)
  ShowHTML "  {" & chr(13)
  ShowHTML "    alert(""ERRO!!!\nA nova senha deve ser diferente da senha atual."");" & chr(13)
  ShowHTML "    theForm.senha_nova1.value='';" & chr(13)
  ShowHTML "    theForm.senha_nova.value='';" & chr(13)
  ShowHTML "    theForm.senha_nova.focus();" & chr(13)
  ShowHTML "    return (false);" & chr(13)
  ShowHTML "  }" & chr(13)
  ShowHTML "  return (true);" & chr(13)
  ShowHTML "}" & chr(13)
  ShowHTML "//-->" & chr(13)
  ShowHTML "</script>" & chr(13)

  ShowHTML "<FORM ACTION=""" & w_dir & sstrCL & "?EW=119&CL=" & CL & "&EA=" & sstrEA & """ METHOD=""POST"" onSubmit=""return(Validacao(this))"">" & VbCrLf
  ShowHTML "<TABLE align=""center"" WIDTH=""100%"" BORDER=0>" & VbCrLf
  ShowHTML "  <TR valign=""TOP""><TD ALIGN=""RIGHT""><B>Senha atual:</b><TD><INPUT TYPE=""PASSWORD"" NAME=""senha_atual"" SIZE=""14"" MAXLENTH=""14"" VALUE=""""></TD></TR>" & VbCrLf
  ShowHTML "  <TR valign=""TOP""><TD ALIGN=""RIGHT""><B>Informe a nova senha:</b><TD><INPUT TYPE=""PASSWORD"" NAME=""senha_nova"" SIZE=""14"" MAXLENTH=""14"" VALUE=""""></TD></TR>" & VbCrLf
  ShowHTML "  <TR valign=""TOP""><TD ALIGN=""RIGHT""><B>Redigite a nova senha:</b><TD><INPUT TYPE=""PASSWORD"" NAME=""senha_nova1"" SIZE=""14"" MAXLENTH=""14"" VALUE=""""></TD></TR>" & VbCrLf
  ShowHTML "  <TR valign=""TOP""><TD colspan=""2"" align=""center""><INPUT TYPE=""SUBMIT"" class=""botao"" NAME=""BOTAO"" VALUE=""Confirmar troca da senha de acesso""></TD></TR>" & VbCrLf
  ShowHTML "</TABLE>" & VbCrLf
  ShowHTML "</FORM>" & VbCrLf

  If Request("senha_nova") > "" Then
     
     ShowHTML "<SCRIPT LANGUAGE=""JAVASCRIPT""><!--" & VbCrLf

     sql = "select count(*) as existe from escAluno where ds_senha_acesso = '" & Trim(uCase(Request("senha_atual"))) & "' and " & sstrEA

     RS.Open sql, sobjConn, adOpenForwardOnly
     
     If RS("existe") = 0 Then
        ShowHTML "  alert('ERRO!!!\nSenha atual inválida.');" & VbCrLf
     Else
        On Error Resume Next
        sql = "update escAluno set ds_senha_acesso = '" & Trim(uCase(Request("senha_nova"))) & "' where " & sstrEA
        sobjConn.Execute sql
        If Err.number <> 0 Then
           ShowHTML "  alert('ERRO!!!\nPermissão de troca de senha negada. Favor contatar a escola e comunicar este problema!');" & VbCrLf
        Else
           ShowHTML "  alert('SUCESSO!!!\nSenha atualizada! A partir de agora, use sempre a nova senha.');" & VbCrLf
        End If
     End If
     ShowHTML "--></SCRIPT>" & VbCrLf

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Tela de Senha
REM =========================================================================

REM =========================================================================
REM Grava mensagens lidas
REM -------------------------------------------------------------------------
Public Sub GravaMens

  Dim sql

  sql = "update escMensagem_Aluno set in_lida = 'Sim' where " & sstrIN

  sobjConn.Execute sql

  Response.Redirect sstrCL & "?EW=118&EF=" & sstrEF & "&CL=" & CL & "&EA=" & sstrEA

End Sub
REM -------------------------------------------------------------------------
REM Final da GravaMens
REM =========================================================================

REM =========================================================================
REM Monta a tela de Boletim
REM -------------------------------------------------------------------------
Public Sub ShowBoletim

  Dim sql, w_ds_mensagem_boletim

  ShowHTML "<tr><td>"

  ' Recupera os dados cadastrais do aluno
  sql = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, a.dt_nascimento,  " & _
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end as in_sexo, " & _
        "       a.ds_naturalidade, a.no_mae, a.nr_fone_mae, a.no_pai, a.nr_fone_pai, " & _
        "       a.no_resposavel, a.nr_fone_responsavel, a.ds_email_responsavel, nr_fone_1, nr_fone_2 " & _
        "FROM escAluno AS a " & _
        "WHERE a." & sstrEA 
  RS.Open sql, sobjConn, adOpenForwardOnly
  If Not RS.EOF Then
    ShowHTML "<TABLE align=""center"" WIDTH=""100%"" BORDER=0>"
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD COLSPAN=""2"" ALIGN=""RIGHT""><FONT FACE=VERDANA SIZE=1><A HREF=""#"" onClick=""window.open('" & sstrCL & "?EW=conWhatExBoletimImp&CL=" & CL & "&EA=" & sstrEA & "','Boletim','width=600, height=350, top=50, left=50, toolbar=no, scrollbars=yes, resizable=yes, status=no'); return false;"" title=""Clique para visualizar a versão de impressão do boletim!""><B>Versão p/ Impressão</B><IMG ALIGN=""CENTER"" BORDER=0   TITLE=""Imprimir"" SRC=""img/bt_imprimir.gif""></A></TD>"
    ShowHTML "  </TR>"
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Nome:</b> " & RS("NO_ALUNO") & "</TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Matrícula:</b> " & RS("NR_MATRICULA") & "</TD>"
    ShowHTML "  </TR>"
    ShowHTML "  <TR valign=""TOP""><TD COLSPAN=""3""><TABLE WIDTH=""100%"" BORDER=""1"">"
    ShowHTML "    <TR valign=""TOP"" align=""CENTER"">"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Ano Letivo</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Série</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Turno</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Turma</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Etapa/Curso/Modalidade</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Situação</b></TD>"
    sql = "SELECT c.ano_letivo, c.st_aluno, " & VbCrLf & _
          "       b.ds_grau, b.ds_serie, b.ds_turma, replace(b.ds_curso,'¬','ª') ds_curso, " & VbCrLf & _
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end as ds_turno " & VbCrLf & _
          "FROM escAluno_Turma      AS c " & VbCrLf & _
          "     INNER JOIN escTurma AS b ON (b.sq_turma = c.sq_turma and b.sq_site_cliente = c.sq_site_cliente and c.ano_letivo = " & w_ano_letivo & ") " & VbCrLf & _
          "WHERE c.sq_aluno = " & RS("SQ_ALUNO") & VbCrLf 
    RS1.Open sql, sobjConn, adOpenForwardOnly
    Do While Not RS1.EOF
      ShowHTML "  <TR valign=""TOP"">"
      'ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Grau:</b><BR>" & RS1("DS_GRAU") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER""><FONT FACE=VERDANA SIZE=1>" & RS1("ANO_LETIVO") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER""><FONT FACE=VERDANA SIZE=1>" & RS1("DS_SERIE") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("DS_TURNO") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("DS_TURMA") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>"
      If Trim(RS1("DS_CURSO")) > "" Then ShowHTML RS1("DS_CURSO") Else ShowHTML "-" End If
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("ST_ALUNO") & "</TD>"
      RS1.MoveNext
    Loop
    RS1.Close
    ShowHTML "     </TR>"
    ShowHTML "  </TABLE>"
    ShowHTML "</TABLE>"
  End If
  RS.Close

  ' Apresenta caixa de seleção do ano letivo
  SQL = "select distinct ano_letivo from escBoletim where " & sstrEA & " order by ano_letivo desc"
  RS2.Open sql, sobjConn, adOpenForwardOnly
  If Not RS2.EOF Then
     RS2.MoveFirst
     If Request("w_ano_letivo") > "" Then
        w_ano_letivo = Request("w_ano_letivo")
     Else
        w_ano_letivo = RS2("ano_letivo")
     End If
     RS2.MoveFirst
  End If

  ShowHTML "<tr><td><TABLE border=1 cellpadding=5 cellspacing=0 width=""100%"" BGCOLOR=""#F7F7F7"">"
  ShowHTML "<FORM ACTION=""" & w_dir & sstrCL & """ NAME=""Form"" METHOD=""GET"">"
  ShowHTML "  <INPUT TYPE=""HIDDEN"" NAME=""EW"" VALUE=""" & sstrEW & """>"
  ShowHTML "  <INPUT TYPE=""HIDDEN"" NAME=""CL"" VALUE=""" & sstrCL & """>"
  ShowHTML "  <INPUT TYPE=""HIDDEN"" NAME=""EA"" VALUE=""" & sstrEA & """>"
  'ShowHTML "  <TR valign=""TOP""><TD COLSPAN=14><B><FONT FACE=VERDANA SIZE=1>Observação: para trocar o ano letivo, apenas escolha o ano desejado na caixa de seleção.</b></font></td>"
  ShowHTML "  <TR ALIGN=""CENTER"">"
  ShowHTML "    <TD ALIGN=""CENTER""><FONT FACE=VERDANA SIZE=1>"
  'ShowHTML "       Ano: <SELECT class=""texto"" NAME=""w_ano_letivo"" SIZE=1 onChange=""javascript:document.Form.submit();"">"
  'Do While Not RS2.EOF
  '   If w_ano_letivo & "a" = RS2("ano_letivo")  & "a" Then
  '      ShowHTML "       <OPTION VALUE=""" & RS2("ANO_LETIVO") & """ SELECTED> " & RS2("ANO_LETIVO")
  '   Else
  '      ShowHTML "       <OPTION VALUE=""" & RS2("ANO_LETIVO") & """> " & RS2("ANO_LETIVO")
  '   End If
  '   RS2.MoveNext
  'Loop
  'ShowHTML "       </SELECT></TD>"
  ShowHTML "    <TD colspan=2 nowrap>1º Bim"
  ShowHTML "    <TD colspan=2 nowrap>2º Bim"
  ShowHTML "    <TD colspan=2 nowrap>3º Bim"
  ShowHTML "    <TD colspan=2 nowrap>4º Bim"
  ShowHTML "    <TD rowspan=2>Média Anual"
  ShowHTML "    <TD rowspan=2>Recup. Final"
'  ShowHTML "    <TD rowspan=2>Recup. Esp."
  ShowHTML "    <TD rowspan=2>Faltas"
  ShowHTML "    <TD rowspan=2>Média Final"
  ShowHTML "    <TD rowspan=2>Result."
  ShowHTML "  </TR>"
  ShowHTML "  <TR>"
  ShowHTML "    <TD nowrap>Comp.Curric."
  ShowHTML "    <TD>Nt"
'  ShowHTML "    <TD>Recup."
  ShowHTML "    <TD>Flt"
  ShowHTML "    <TD>Nt"
'  ShowHTML "    <TD>Recup."
  ShowHTML "    <TD>Flt"
  ShowHTML "    <TD>Nt"
'  ShowHTML "    <TD>Recup."
  ShowHTML "    <TD>Flt"
  ShowHTML "    <TD>Nt"
'  ShowHTML "    <TD>Recup."
  ShowHTML "    <TD>Flt"
  ShowHTML "  <TR>"
  ShowHTML "    <TD COLSPAN=""14"" HEIGHT=""1"" BGCOLOR=""##DAEABD"">"
  ShowHTML "  </TR>"

  sql = "SELECT a.*, b.sg_disciplina, c.ds_mensagem_boletim " & VbCrLf & _
        "FROM escBoletim AS a " & VbCrLf & _
        "     INNER JOIN escDisciplina AS b on (a.sq_disciplina = b.sq_disciplina) " & VbCrLf & _
        "     INNER JOIN escAluno AS c on (a.sq_aluno = c.sq_aluno) " & VbCrLf & _
        "WHERE a." & sstrEA & VbCrLf & _
        "  AND a.ano_letivo = " & w_ano_letivo & " " & VbCrLf & _
        "ORDER BY b.sg_disciplina" & VbCrLf
  RS.Open sql, sobjConn, adOpenForwardOnly

  If Not RS.EOF Then

     w_ds_mensagem_boletim = RS("DS_MENSAGEM_BOLETIM")
    
     While Not RS.EOF
	   ShowHTML "    <TR>"
	   ShowHTML "    <TD><center><b>" & RS("sg_disciplina")
	   ShowHTML "    <TD>" & trocaNulo(RS("b1_nota"))
   '	ShowHTML "    <TD>" & trocaNulo(RS("b1_recup"))
	   ShowHTML "    <TD>" & trocaNulo(RS("b1_falta"))
	   ShowHTML "    <TD>" & trocaNulo(RS("b2_nota"))
   '	ShowHTML "    <TD>" & trocaNulo(RS("b2_recup"))
	   ShowHTML "    <TD>" & trocaNulo(RS("b2_falta"))
	   ShowHTML "    <TD>" & trocaNulo(RS("b3_nota"))
   '	ShowHTML "    <TD>" & trocaNulo(RS("b3_recup"))
	   ShowHTML "    <TD>" & trocaNulo(RS("b3_falta"))
	   ShowHTML "    <TD>" & trocaNulo(RS("b4_nota"))
   '	ShowHTML "    <TD>" & trocaNulo(RS("b4_recup"))
	   ShowHTML "    <TD>" & trocaNulo(RS("b4_falta"))
	   ShowHTML "    <TD>" & trocaNulo(RS("ds_media_anual"))
	   ShowHTML "    <TD>" & trocaNulo(RS("ds_recup_final"))
	   ShowHTML "    <TD>" & trocaNulo(RS("ds_falta_anual"))
   '	ShowHTML "    <TD>" & trocaNulo(RS("ds_recup_esp"))
	   ShowHTML "    <TD>" & trocaNulo(RS("ds_media_final"))
	   ShowHTML "    <TD><b>" 
	   If instr(1,lcase(trocaNulo(RS("ds_resultado"))),"rep") then
	      ShowHTML "<font color=""#FF0000"">" & trocaNulo(RS("ds_resultado"))
	   Else
	      ShowHTML "" & trocaNulo(RS("ds_resultado"))
	   End If
    
       RS.MoveNext
  
     Wend
     If w_ds_mensagem_boletim > "" Then
        ShowHTML "    <TR><TD colspan=14 align=""left""><b>Mensagem:<BR><font color=""#FF0000""><b>" & w_ds_mensagem_boletim
     End If

  Else
     ShowHTML "    <TR><TD colspan=14><b>Não há notas informadas.</b>"
  End If
  ShowHTML "</FORM>"
  ShowHTML "    </TABLE>"
	 
  RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Boletim
REM =========================================================================

REM =========================================================================
REM Monta a tela de Boletim para impressão
REM -------------------------------------------------------------------------
Public Sub ShowBoletimImp
  
  Dim sql, w_ds_mensagem_boletim
  
  ShowHTML "<tr><td>"
  
  ' Recupera o nome da escola
  sql = "SELECT b.ds_cliente " & _
        "FROM escAluno AS a " & _
        "     LEFT OUTER JOIN escCliente b on (a.sq_site_cliente = b.sq_cliente)" & _
        "WHERE a." & sstrEA 
  RS.Open sql, sobjConn, adOpenForwardOnly
  
  ShowHTML "<TABLE align=""center"" WIDTH=""100%"" BORDER=0>"
  ShowHTML "  <TR valign=""TOP"">"
  ShowHTML "    <TD align=""right""><IMG ALIGN=""CENTER"" TITLE=""Imprimir"" SRC=""img/bt_imprimir.gif"" onClick=""window.print();window.close();""></TD>"
  ShowHTML "  </TR>"
  ShowHTML "  <TR valign=""TOP"">"
  ShowHTML "    <TD ALIGN=""center""><FONT FACE=VERDANA SIZE=2><B>" & RS("DS_CLIENTE") & "</TD>"
  ShowHTML"   </TR>"
  ShowHTML "</TABLE>"
  RS.Close
  
  ' Recupera os dados cadastrais do aluno
  sql = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, a.dt_nascimento,  " & _
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end as in_sexo, " & _
        "       a.ds_naturalidade, a.no_mae, a.nr_fone_mae, a.no_pai, a.nr_fone_pai, " & _
        "       a.no_resposavel, a.nr_fone_responsavel, a.ds_email_responsavel, nr_fone_1, nr_fone_2 " & _
        "FROM escAluno AS a " & _
        "WHERE a." & sstrEA 
  RS.Open sql, sobjConn, adOpenForwardOnly
  
  If Not RS.EOF Then
    ShowHTML "<TABLE align=""center"" WIDTH=""100%"" BORDER=0>"
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Nome:</b> " & RS("NO_ALUNO") & "</TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Matrícula:</b> " & RS("NR_MATRICULA") & "</TD>"
    ShowHTML "  </TR>"
    ShowHTML "  <TR valign=""TOP""><TD COLSPAN=""3""><TABLE WIDTH=""100%"" BORDER=""1"">"
    ShowHTML "    <TR valign=""TOP"" align=""CENTER"">"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Ano Letivo</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Série</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Turno</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Turma</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Etapa/Curso/Modalidade</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Situação</b></TD>"
    sql = "SELECT c.ano_letivo, c.st_aluno, " & VbCrLf & _
          "       b.ds_grau, b.ds_serie, b.ds_turma, replace(b.ds_curso,'¬','ª') ds_curso, " & VbCrLf & _
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end as ds_turno " & VbCrLf & _
          "FROM escAluno_Turma      AS c " & VbCrLf & _
          "     INNER JOIN escTurma AS b ON (b.sq_turma = c.sq_turma and b.sq_site_cliente = c.sq_site_cliente and c.ano_letivo = " & w_ano_letivo & ") " & VbCrLf & _
          "WHERE c.sq_aluno = " & RS("SQ_ALUNO") & VbCrLf 
    RS1.Open sql, sobjConn, adOpenForwardOnly
    Do While Not RS1.EOF
      ShowHTML "  <TR valign=""TOP"">"
      'ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Grau:</b><BR>" & RS1("DS_GRAU") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER""><FONT FACE=VERDANA SIZE=1>" & RS1("ANO_LETIVO") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER""><FONT FACE=VERDANA SIZE=1>" & RS1("DS_SERIE") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("DS_TURNO") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("DS_TURMA") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>"
      If Trim(RS1("DS_CURSO")) > "" Then ShowHTML RS1("DS_CURSO") Else ShowHTML "-" End If
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("ST_ALUNO") & "</TD>"
      RS1.MoveNext
    Loop
    RS1.Close
    ShowHTML "     </TR>"
    ShowHTML "  </TABLE>"
    ShowHTML "</TABLE>"
  End If
  RS.Close

  ' Apresenta caixa de seleção do ano letivo
  SQL = "select distinct ano_letivo from escBoletim where " & sstrEA & " order by ano_letivo desc"
  RS2.Open sql, sobjConn, adOpenForwardOnly
  If Not RS2.EOF Then
     RS2.MoveFirst
     If Request("w_ano_letivo") > "" Then
        w_ano_letivo = Request("w_ano_letivo")
     Else
        w_ano_letivo = RS2("ano_letivo")
     End If
     RS2.MoveFirst
  End If

  ShowHTML "<tr><td><TABLE border=1 cellpadding=5 cellspacing=0 width=""100%"" BGCOLOR=""#F7F7F7"">"
  ShowHTML "<FORM ACTION=""" & w_dir & sstrCL & """ NAME=""Form"" METHOD=""GET"">"
  ShowHTML "  <INPUT TYPE=""HIDDEN"" NAME=""EW"" VALUE=""" & sstrEW & """>"
  ShowHTML "  <INPUT TYPE=""HIDDEN"" NAME=""CL"" VALUE=""" & sstrCL & """>"
  ShowHTML "  <INPUT TYPE=""HIDDEN"" NAME=""EA"" VALUE=""" & sstrEA & """>"
  'ShowHTML "  <TR valign=""TOP""><TD COLSPAN=14><B><FONT FACE=VERDANA SIZE=1>Observação: para trocar o ano letivo, apenas escolha o ano desejado na caixa de seleção.</b></font></td>"
  ShowHTML "  <TR ALIGN=""CENTER"">"
  ShowHTML "    <TD ALIGN=""CENTER"" ROWSPAN=""2""><FONT FACE=VERDANA SIZE=1>Comp.Curric."
  'ShowHTML "       Ano: <SELECT class=""texto"" NAME=""w_ano_letivo"" SIZE=1 onChange=""javascript:document.Form.submit();"">"
  'Do While Not RS2.EOF
  '   If w_ano_letivo & "a" = RS2("ano_letivo")  & "a" Then
  '      ShowHTML "       <OPTION VALUE=""" & RS2("ANO_LETIVO") & """ SELECTED> " & RS2("ANO_LETIVO")
  '   Else
  '      ShowHTML "       <OPTION VALUE=""" & RS2("ANO_LETIVO") & """> " & RS2("ANO_LETIVO")
  '   End If
  '   RS2.MoveNext
  'Loop
  'ShowHTML "       </SELECT></TD>"
  ShowHTML "    <TD colspan=2 nowrap><FONT FACE=VERDANA SIZE=1>1º Bim"
  ShowHTML "    <TD colspan=2 nowrap><FONT FACE=VERDANA SIZE=1>2º Bim"
  ShowHTML "    <TD colspan=2 nowrap><FONT FACE=VERDANA SIZE=1>3º Bim"
  ShowHTML "    <TD colspan=2 nowrap><FONT FACE=VERDANA SIZE=1>4º Bim"
  ShowHTML "    <TD rowspan=2><FONT FACE=VERDANA SIZE=1>Média Anual"
  ShowHTML "    <TD rowspan=2><FONT FACE=VERDANA SIZE=1>Recup. Final"
'  ShowHTML "    <TD rowspan=2><FONT FACE=VERDANA SIZE=1>Recup. Esp."
  ShowHTML "    <TD rowspan=2><FONT FACE=VERDANA SIZE=1>Faltas"
  ShowHTML "    <TD rowspan=2><FONT FACE=VERDANA SIZE=1>Média Final"
  ShowHTML "    <TD rowspan=2><FONT FACE=VERDANA SIZE=1>Result."
  ShowHTML "  </TR>"
  ShowHTML "  <TR>"
  'ShowHTML "    <TD nowrap><FONT FACE=VERDANA SIZE=1>Comp.Curric."
  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Nt"
'  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Recup."
  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Flt"
  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Nt"
'  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Recup."
  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Flt"
  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Nt"
'  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Recup."
  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Flt"
  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Nt"
'  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Recup."
  ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>Flt"
  ShowHTML "  <TR>"
  ShowHTML "    <TD COLSPAN=""14"" HEIGHT=""1"" BGCOLOR=""##DAEABD"">"
  ShowHTML "  </TR>"

  sql = "SELECT a.*, b.sg_disciplina, c.ds_mensagem_boletim " & VbCrLf & _
        "FROM escBoletim AS a " & VbCrLf & _
        "     INNER JOIN escDisciplina AS b on (a.sq_disciplina = b.sq_disciplina) " & VbCrLf & _
        "     INNER JOIN escAluno AS c on (a.sq_aluno = c.sq_aluno) " & VbCrLf & _
        "WHERE a." & sstrEA & VbCrLf & _
        "  AND a.ano_letivo = " & w_ano_letivo & " " & VbCrLf & _
        "ORDER BY b.sg_disciplina" & VbCrLf
  RS.Open sql, sobjConn, adOpenForwardOnly

  If Not RS.EOF Then

     w_ds_mensagem_boletim = RS("DS_MENSAGEM_BOLETIM")
    
     While Not RS.EOF
	   ShowHTML "    <TR>"
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><center><b>" & RS("sg_disciplina")
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b1_nota"))
   '	ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b1_recup"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b1_falta"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b2_nota"))
   '	ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b2_recup"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b2_falta"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b3_nota"))
   '	ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b3_recup"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b3_falta"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b4_nota"))
   '	ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b4_recup"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("b4_falta"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("ds_media_anual"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("ds_recup_final"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("ds_falta_anual"))
   '	ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("ds_recup_esp"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & trocaNulo(RS("ds_media_final"))
	   ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><b>" 
	   If instr(1,lcase(trocaNulo(RS("ds_resultado"))),"rep") then
	      ShowHTML "<font color=""#FF0000"">" & trocaNulo(RS("ds_resultado"))
	   Else
	      ShowHTML "" & trocaNulo(RS("ds_resultado"))
	   End If
    
       RS.MoveNext
  
     Wend
     If w_ds_mensagem_boletim > "" Then
        ShowHTML "    <TR><TD colspan=14 align=""left""><FONT FACE=VERDANA SIZE=1><b>Mensagem:<BR><font color=""#FF0000""><b>" & w_ds_mensagem_boletim
     End If

  Else
     ShowHTML "    <TR><TD colspan=14><FONT FACE=VERDANA SIZE=1><b>Não há notas informadas.</b>"
  End If
  ShowHTML "</FORM>"
  ShowHTML "    </TABLE>"
	 
  RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Boletim para impressão
REM =========================================================================

REM =========================================================================
REM Monta a tela de Grade Horária
REM -------------------------------------------------------------------------
Public Sub ShowGrade

  Dim sql, strHorario, wCont, wAtual
  Dim i, j
  Dim Linha,Coluna
  Dim grade(11, 8)

  Linha  = 11
  Coluna = 8
  grade(0,0) = "Horário"
  grade(0,1) = "Segunda"
  grade(0,2) = "Terça"
  grade(0,3) = "Quarta"
  grade(0,4) = "Quinta"
  grade(0,5) = "Sexta"
  grade(0,6) = "Sábado"
  grade(0,7) = "Domingo"
  grade(1,0) = "Primeiro"
  grade(2,0) = "Segundo"
  grade(3,0) = "Terceiro"
  grade(4,0) = "Quarto"
  grade(5,0) = "Quinto"
  grade(6,0) = "Sexto"
  grade(7,0) = "Sétimo"
  grade(8,0) = "Oitavo"
  grade(9,0) = "Nono"
  grade(10,0) = "Décimo"
  FOR i=1 TO Linha-1
      FOR j = 1 to Coluna-1
          grade(i, j)="-"
      NEXT
  NEXT

  ShowHTML "<tr><td>"

  ' Recupera os dados cadastrais do aluno
  sql = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, a.dt_nascimento,  " & VbCrLf & _
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end as in_sexo, " & VbCrLf & _
        "       a.ds_naturalidade, a.no_mae, a.nr_fone_mae, a.no_pai, a.nr_fone_pai, " & VbCrLf & _
        "       a.no_resposavel, a.nr_fone_responsavel, a.ds_email_responsavel, nr_fone_1, nr_fone_2 " & VbCrLf & _
        "FROM escAluno AS a " & VbCrLf & _
        "WHERE a." & sstrEA & VbCrLf
  RS.Open sql, sobjConn, adOpenForwardOnly
  If Not RS.EOF Then
    ShowHTML "<TABLE align=""center"" WIDTH=""100%"" BORDER=0>"
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Nome:</b> " & RS("NO_ALUNO") & "</TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Matrícula:</b> " & RS("NR_MATRICULA") & "</TD>"
    ShowHTML "  </TR>"
    ShowHTML "  <TR valign=""TOP""><TD COLSPAN=""3""><TABLE WIDTH=""100%"" BORDER=""1"">"
    ShowHTML "    <TR valign=""TOP"" align=""CENTER"">"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Ano Letivo</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Série</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Turno</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Turma</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Etapa/Curso/Modalidade</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Situação</b></TD>"
    sql = "SELECT c.ano_letivo, c.st_aluno, " & VbCrLf & _
          "       b.ds_grau, b.ds_serie, b.ds_turma, replace(b.ds_curso,'¬','ª') ds_curso, " & VbCrLf & _
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end as ds_turno " & VbCrLf & _
          "FROM escAluno_Turma      AS c " & VbCrLf & _
          "     INNER JOIN escTurma AS b ON (b.sq_turma = c.sq_turma and b.sq_site_cliente = c.sq_site_cliente and c.ano_letivo = " & w_ano_letivo & ") " & VbCrLf & _
          "WHERE c.sq_aluno = " & RS("SQ_ALUNO") & VbCrLf 
    RS1.Open sql, sobjConn, adOpenForwardOnly
    Do While Not RS1.EOF
      ShowHTML "  <TR valign=""TOP"">"
      'ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Grau:</b><BR>" & RS1("DS_GRAU") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER""><FONT FACE=VERDANA SIZE=1>" & RS1("ANO_LETIVO") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER""><FONT FACE=VERDANA SIZE=1>" & RS1("DS_SERIE") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("DS_TURNO") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("DS_TURMA") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>"
      If Trim(RS1("DS_CURSO")) > "" Then ShowHTML RS1("DS_CURSO") Else ShowHTML "-" End If
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("ST_ALUNO") & "</TD>"
      RS1.MoveNext
    Loop
    RS1.Close
    ShowHTML "     </TR>"
    ShowHTML "  </TABLE>"
    ShowHTML "</TABLE>"
  End If
  RS.Close

  ShowHTML "<tr><td align=""center""><TABLE border=1 cellpadding=5 cellspacing=0 BGCOLOR=""#F7F7F7"" width=""100%"">"
  ShowHTML "<FORM ACTION=""" & w_dir & sstrCL & """ NAME=""Form"" METHOD=""GET"">"
  ShowHTML "  <INPUT TYPE=""HIDDEN"" NAME=""EW"" VALUE=""" & sstrEW & """>"
  ShowHTML "  <INPUT TYPE=""HIDDEN"" NAME=""CL"" VALUE=""" & sstrCL & """>"
  ShowHTML "  <INPUT TYPE=""HIDDEN"" NAME=""EA"" VALUE=""" & sstrEA & """>"
  'ShowHTML "  <TR valign=""TOP""><TD COLSPAN=14><B><FONT FACE=VERDANA SIZE=1>Observação: para trocar o ano letivo, apenas escolha o ano desejado na caixa de seleção.</b></font></td>"

  ' Apresenta caixa de seleção do ano letivo
  sql = "SELECT distinct a.ano_letivo " & VbCrLf &  _
        "  FROM escAluno_Turma AS a INNER JOIN escGrade_horaria AS b ON (a.sq_turma = b.sq_turma and a.sq_site_cliente = b.sq_site_cliente and a.ano_letivo = b.ano_letivo) " & VbCrLf & _
        "                           INNER JOIN escDisciplina AS c ON (b.sq_disciplina = c.sq_disciplina) " & VbCrLf & _
        "WHERE a." & sstrEA & VbCrLf & _
        "ORDER BY a.ano_letivo desc" & VbCrLf 
  RS2.Open sql, sobjConn, adOpenForwardOnly
  If Not RS2.EOF Then
     RS2.MoveFirst
     If Request("w_ano_letivo") > "" Then
        w_ano_letivo = Request("w_ano_letivo")
     Else
        w_ano_letivo = RS2("ano_letivo")
     End If
     RS2.MoveFirst
  End If

  ShowHTML "    <TR><TD COLSPAN=14 ALIGN=""LEFT""><FONT FACE=VERDANA SIZE=1>"
  'ShowHTML "       Ano: <SELECT class=""texto"" NAME=""w_ano_letivo"" SIZE=1 onChange=""javascript:document.Form.submit();"">"
  'Do While Not RS2.EOF
  '   If w_ano_letivo & "a" = RS2("ano_letivo")  & "a" Then
  '      ShowHTML "       <OPTION VALUE=""" & RS2("ANO_LETIVO") & """ SELECTED> " & RS2("ANO_LETIVO")
  '   Else
  '      ShowHTML "       <OPTION VALUE=""" & RS2("ANO_LETIVO") & """> " & RS2("ANO_LETIVO")
  '   End If
  '   RS2.MoveNext
  'Loop
  'ShowHTML "       </SELECT></TD></TR>"

  sql = "SELECT b.*, lower(c.ds_disciplina) ds_disciplina " & VbCrLf & _
        "  FROM escAluno_Turma              a " & VbCrLf & _
        "       INNER JOIN escGrade_horaria b ON (a.sq_turma        = b.sq_turma and  " & VbCrLf & _
        "                                         a.sq_site_cliente = b.sq_site_cliente and " & VbCrLf & _
        "                                         a.ano_letivo      = b.ano_letivo and " & VbCrLf & _
        "                                         a.ano_letivo      = " & w_ano_letivo & VbCrLf & _
        "                                        ) " & VbCrLf & _
        "       INNER JOIN escDisciplina    c ON (b.sq_disciplina = c.sq_disciplina) " & VbCrLf & _
        "WHERE a." & sstrEA & " " & VbCrLf & _
        "ORDER BY b.ds_horario,b.ds_dia_semana" & VbCrLf
  RS.Open sql, sobjConn, adOpenForwardOnly

  If Not RS.EOF Then
     Do While Not RS.EOF
        grade(RS("DS_HORARIO"), RS("DS_DIA_SEMANA")) = RS("DS_DISCIPLINA")
        If RS("DS_DIA_SEMANA") > Coluna Then Coluna = RS("DS_DIA_SEMANA") End If
        If RS("DS_HORARIO")    > Linha  Then Linha = RS("DS_HORARIO")     End If
        RS.MoveNext
     Loop
     ShowHTML "  <TR valign=""top"" align=""center"">"
     For I = 0 TO Coluna
         ShowHTML "    <TD><b>" & grade(0,I)
     Next
     ShowHTML "  </TR>"
     ShowHTML "  <TR><TD COLSPAN=""8"" HEIGHT=""1"" BGCOLOR=""##DAEABD""></TR>"
     For I = 1 TO Linha-1
        ShowHTML "  <TR valign=""top"" align=""center"">"
        FOR J = 0 TO Coluna
	        ShowHTML "    <TD>" & grade(I,J)
	    NEXT
     ShowHTML "  </TR>"
     Next
  Else
     ShowHTML "    <TR><TD colspan=14><b>A grade horária não foi informada.</b>"
  End If
  ShowHTML "    </TABLE>"
	 
  RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Grade Horária
REM =========================================================================

REM =========================================================================
REM Monta a tela de Mensagens
REM -------------------------------------------------------------------------
Public Sub ShowMensagem

  Dim sql, strNome, wCont

  ShowHTML "<tr><td>"

  ' Recupera os dados cadastrais do aluno
  sql = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, a.dt_nascimento,  " & _
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end as in_sexo, " & _
        "       a.ds_naturalidade, a.no_mae, a.nr_fone_mae, a.no_pai, a.nr_fone_pai, " & _
        "       a.no_resposavel, a.nr_fone_responsavel, a.ds_email_responsavel, nr_fone_1, nr_fone_2 " & _
        "FROM escAluno AS a " & _
        "WHERE a." & sstrEA
  RS.Open sql, sobjConn, adOpenForwardOnly
  If Not RS.EOF Then
    ShowHTML "<TABLE align=""center"" WIDTH=""100%"" BORDER=0>"
    ShowHTML "  <TR valign=""TOP"">"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Nome:</b> " & RS("NO_ALUNO") & "</TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Matrícula:</b> " & RS("NR_MATRICULA") & "</TD>"
    ShowHTML "  </TR>"
    sql = "SELECT c.st_aluno, c.ano_letivo, " & _
          "       b.ds_grau, b.ds_serie, b.ds_turma, replace(b.ds_curso,'¬','ª') ds_curso, " & _
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end as ds_turno " & _
          "FROM escAluno_Turma AS c INNER JOIN escTurma AS b ON (b.sq_turma = c.sq_turma and b.sq_site_cliente = c.sq_site_cliente and b.ano_letivo = " & w_ano_letivo & ") " & _
          "WHERE c.sq_aluno = " & RS("SQ_ALUNO")
    RS1.Open sql, sobjConn, adOpenForwardOnly
    ShowHTML "  <TR valign=""TOP""><TD COLSPAN=""3""><TABLE WIDTH=""100%"" BORDER=""1"">"
    ShowHTML "    <TR valign=""TOP"" align=""CENTER"">"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Ano letivo</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Série</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Turno</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Turma</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Etapa/Curso/Modalidade</b></TD>"
    ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Situação</b></TD>"
    Do While Not RS1.EOF
      ShowHTML "  <TR valign=""TOP"">"
      'ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1><B>Grau:</b><BR>" & RS1("DS_GRAU") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER""><FONT FACE=VERDANA SIZE=1>" & RS1("ano_letivo") & "</TD>"
      ShowHTML "    <TD ALIGN=""CENTER""><FONT FACE=VERDANA SIZE=1>" & RS1("DS_SERIE") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("DS_TURNO") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("DS_TURMA") & "</TD>"
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>"
      If Trim(RS1("DS_CURSO")) > "" Then ShowHTML RS1("DS_CURSO") Else ShowHTML "-" End If
      ShowHTML "    <TD><FONT FACE=VERDANA SIZE=1>" & RS1("ST_ALUNO") & "</TD>"
      RS1.MoveNext
    Loop
    RS1.Close
    ShowHTML "  </TR></TABLE>"
    ShowHTML "</TABLE><br><br>"
  End If
  RS.Close

  ShowHTML "<tr><td align=""center"">"
  ShowHTML "<TABLE border=1 cellSpacing=1 width=""100%"" BGCOLOR=""#F7F7F7"">"
  ShowHTML "  <TR valign=""top"" align=""center"">"
  ShowHTML "    <TD width=""15%""><b>Data"
  ShowHTML "    <TD width=""75%"" align=""left""><b>Texto"
  ShowHTML "    <TD width=""10%""><b>Situação"
  ShowHTML "  </TR>"

  ' Recupera as mensagems da escola
  sql = "SELECT 'Escola' AS Origem, b.dt_mensagem, b.ds_mensagem, b.in_lida " & _
        "FROM escAluno AS a INNER JOIN escMensagem_aluno AS b ON (a.sq_aluno = b.sq_aluno) " & _
        "WHERE a." & sstrEA & " " & _
        "ORDER BY dt_mensagem DESC"

  RS.Open sql, sobjConn, adOpenForwardOnly

  wCont = 0
  If Not RS.EOF Then
 	wCont = 1
 	Do While Not RS.EOF

	ShowHTML "  <TR valign=""top"" align=""left"">"
	ShowHTML "    <TD align=""center"">" & Mid(100+Day(RS("dt_mensagem")),2,2) & "/" & Mid(100+Month(RS("dt_mensagem")),2,2) & "/" &Year(RS("dt_mensagem"))
	ShowHTML "    <TD>" & RS("ds_mensagem")
	If Trim(RS("in_lida")) = "Sim" Then
 	   ShowHTML "    <TD align=""center"">Lida"
	ElseIf Trim(RS("in_lida")) = "-" Then
 	   ShowHTML "    <TD align=""center"">-"
 	Else
 	   ShowHTML "    <TD align=""center"">Nova"
 	End If
	ShowHTML "  </TR>"

	RS.MoveNext

	Loop

  Else

	ShowHTML "  <TR>"
	ShowHTML "    <TD colspan=4 align=""center"">Não há mensagens"
	ShowHTML "  </TR>"

  End If

  ShowHTML "    </TABLE>"
	 
  RS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Página de Mensagens
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub Main

  If sstrEW = conWhatGravaMens Then
     GravaMens
  End If

  If sstrEW <> conWhatCabecalho Then
     ShowHTML "<style>" & chr(13) & chr(10)
     ShowHTML "<//" & chr(13) & chr(10)
     ShowHTML "        a                       { color: ##DAEABD; text-decoration: none; }" & chr(13) & chr(10)
     ShowHTML "        a:hover                 { color:##DAEABD; text-decoration: underline; }" & chr(13) & chr(10)
     ShowHTML ">" & chr(13) & chr(10)
     ShowHTML "</style>" & chr(13) & chr(10)
     If sstrEW <> conWhatSenha Then
        RS.Open "select count(*) as existe from escAluno where ds_senha_acesso = '12345' and " & sstrEA, sobjConn, adOpenForwardOnly
        If RS("existe") > 0 Then
           ShowHTML "<SCRIPT LANGUAGE=""JAVASCRIPT""><!--" & VbCrLf
           ShowHTML "  alert('ALERTA!!!\nUse a opção ""Senha"" para alterar sua senha de acesso.');" & VbCrLf
           ShowHTML "--></SCRIPT>" & VbCrLf
        End If
        RS.Close
     End If
  End If

  Select Case sstrEW
    Case conWhatMenu           ShowMenu
    Case conWhatExBoletim      ShowBoletim
    Case conWhatExGrade        ShowGrade
    Case conWhatSenha          ShowSenha
    Case conWhatExMens         ShowMensagem
    Case "conWhatExBoletimImp" ShowBoletimImp
  End Select
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================
%>
