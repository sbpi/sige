<%
Dim w_sq_tipo_cliente
REM =========================================================================
REM Montagem da sele��o de regionais de ensino
REM -------------------------------------------------------------------------
Sub SelecaoRegional (label, accesskey, hint, chave, chaveAux, campo, restricao, atributo)

    If Mid(Session("username"),1,2) = "RE" Then
       SQL = "SELECT  a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " & VbCrLf & _
             "  FROM  escCLIENTE a " & VbCrLf & _
             " WHERE  a." & CL & VbCrLf & _
             "ORDER BY a.ds_cliente "
    Else
       SQL = "SELECT a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " & VbCrLf & _
             "  FROM escCLIENTE a " & VbCrLf & _
             "       inner join escTipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente) " & VbCrLf & _
             " WHERE b.tipo = 2 " & VbCrLf & _
             "    OR (b.tipo = 1 and a.sq_cliente_pai is null) " & VbCrLf & _
             "ORDER BY b.tipo, a.ds_cliente" & VbCrLf
    End If
    ConectaBD SQL
    
    If IsNull(hint) Then
       ShowHTML "          <td valign=""top""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    Else
       ShowHTML "          <td valign=""top"" ONMOUSEOVER=""popup('" & hint & "','white')""; ONMOUSEOUT=""kill()""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    End If
    If restricao = "CADASTRO" Then
       ShowHTML "          <option value="""">---"
    Else
       If RS.RecordCount > 1 Then ShowHTML "          <option value="""">Tanto faz" End If
    End If
    While Not RS.EOF
       If cDbl(nvl(RS("sq_cliente"),-1)) = cDbl(nvl(chave,-1)) Then
          ShowHTML "          <option value=""" & RS("sq_cliente") & """ SELECTED>" & RS("ds_cliente")
       Else
          ShowHTML "          <option value=""" & RS("sq_cliente") & """>" & RS("ds_cliente")
       End If
       RS.MoveNext
    Wend
    ShowHTML "          </select>"
    DesconectaBD
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da sele��o de regionais de ensino
REM -------------------------------------------------------------------------
Sub SelecaoRegionalEscola (label, accesskey, hint, chave, chaveAux, campo, restricao, atributo)
    
    If IsNull(hint) Then
       ShowHTML "          <td valign=""top""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    Else
       ShowHTML "          <td valign=""top"" ONMOUSEOVER=""popup('" & hint & "','white')""; ONMOUSEOUT=""kill()""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    End If
    ShowHTML "          <option value="""">---"

    If Nvl(chaveAux,"nulo") <> "nulo" then
    
       SQL = "SELECT b.tipo, a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " & VbCrLf & _
             "  FROM escCLIENTE                      a " & VbCrLf & _
             "       inner      join escTipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente and b.tipo = 2) " & VbCrLf & _
             "       left outer join escCliente      c on (a.sq_cliente = c.sq_cliente_pai and c.sq_cliente = " & chaveAux & ") " & VbCrLf & _
             "ORDER BY b.tipo, a.ds_cliente" & VbCrLf
       ConectaBD SQL
       
       While Not RS.EOF
          If cInt(nvl(chave,0)) = cInt(RS("sq_cliente")) Then
             ShowHTML "          <option value=""" & RS("sq_cliente") & """ SELECTED>" & RS("ds_cliente")
          Else
             ShowHTML "          <option value=""" & RS("sq_cliente") & """>" & RS("ds_cliente")
          End If
          RS.MoveNext
       Wend
       DesconectaBD
    End If
    ShowHTML "          </select>"
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da sele��o de regi�es administrativas
REM -------------------------------------------------------------------------
Sub SelecaoRegiaoAdm (label, accesskey, hint, chave, chaveAux, campo, restricao, atributo)
    If IsNull(hint) Then
       ShowHTML "          <td valign=""top""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    Else
       ShowHTML "          <td valign=""top"" ONMOUSEOVER=""popup('" & hint & "','white')""; ONMOUSEOUT=""kill()""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    End If
    ShowHTML "          <option value="""">---"

    If Nvl(chaveAux,"nulo") <> "nulo" then
    
       SQL = "SELECT a.sq_regiao_adm, a.no_regiao " & VbCrLf & _
             "  FROM escRegiao_Administrativa a " & VbCrLf & _
             " WHERE a.ativo = 'S' " & VbCrLf & _
             "ORDER BY a.no_regiao" & VbCrLf
       ConectaBD SQL
       
       While Not RS.EOF
          If cInt(nvl(chave,0)) = cInt(RS("sq_regiao_adm")) Then
             ShowHTML "          <option value=""" & RS("sq_regiao_adm") & """ SELECTED>" & RS("no_regiao")
          Else
             ShowHTML "          <option value=""" & RS("sq_regiao_adm") & """>" & RS("no_regiao")
          End If
          RS.MoveNext
       Wend
       DesconectaBD
    End If
    ShowHTML "          </select>"
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da sele��o de regionais de ensino
REM -------------------------------------------------------------------------
Sub SelecaoTipoEscola (label, accesskey, hint, chave, chaveAux, campo, restricao, atributo)
    
    If IsNull(hint) Then
       ShowHTML "          <td valign=""top""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    Else
       ShowHTML "          <td valign=""top"" ONMOUSEOVER=""popup('" & hint & "','white')""; ONMOUSEOUT=""kill()""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    End If
    ShowHTML "          <option value="""">---"

    If Nvl(chaveAux,"nulo") <> "nulo" then
       
       SQL2 = "select sq_cliente, sq_tipo_cliente cliente " & VbCrLf & _
              "  from escCliente " & VbCrLf & _
              "  where sq_cliente = " & chaveAux
              
       ConectaBD SQL2
       w_sq_tipo_cliente = RS("cliente")
       DesconectaBD
       
       SQL = "select sq_tipo_cliente, ds_tipo_cliente, ds_registro, tipo, " & VbCrLf & _
             "       case tipo when 1 then 'Secretaria' " & VbCrLf & _
             "                 when 2 then 'Regional'   " & VbCrLf & _
             "                 when 3 then 'Escola'     " & VbCrLf & _
             "       end nm_tipo                        " & VbCrLf & _
             "  from escTipo_Cliente                    " & VbCrLf & _
             "  where tipo = 3                          " & VbCrLf
       
       ConectaBD SQL
       While Not RS.EOF
          If cdbl(w_sq_tipo_cliente) = cdbl(RS("sq_tipo_cliente")) Then
             ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """ SELECTED>" & RS("ds_tipo_cliente")
          Else
             ShowHTML "          <option value=""" & RS("sq_tipo_cliente") & """>" & RS("ds_tipo_cliente")
          End If
          RS.MoveNext
       Wend
       DesconectaBD
    End If
    ShowHTML "          </select>"
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da sele��o de bloco de dados
REM -------------------------------------------------------------------------
Sub SelecaoBlocoDados (label, accesskey, hint, chave, chaveAux, campo, restricao, atributo)
    If IsNull(hint) Then
       ShowHTML "          <td valign=""top""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    Else
       ShowHTML "          <td valign=""top"" ONMOUSEOVER=""popup('" & hint & "','white')""; ONMOUSEOUT=""kill()""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    End If
    ShowHTML "          <option value="""">Todos"
    If Nvl(chave,"") = "223" Then ShowHTML "          <option value=""223"" SELECTED>Administrativo"     Else ShowHTML "          <option value=""223"">Administrativo"     End If
    If Nvl(chave,"") = "135" Then ShowHTML "          <option value=""135"" SELECTED>Arquivos"           Else ShowHTML "          <option value=""135"">Arquivos"           End If
    If Nvl(chave,"") = "133" Then ShowHTML "          <option value=""133"" SELECTED>Calend�rio"         Else ShowHTML "          <option value=""133"">Calend�rio"         End If
    If Nvl(chave,"") = "220" Then ShowHTML "          <option value=""220"" SELECTED>Dados adicionais"   Else ShowHTML "          <option value=""220"">Dados adicionais"   End If
    If Nvl(chave,"") = "123" Then ShowHTML "          <option value=""123"" SELECTED>Dados b�sicos"      Else ShowHTML "          <option value=""123"">Dados b�sicos"      End If
    If Nvl(chave,"") = "144" Then ShowHTML "          <option value=""144"" SELECTED>Dados do site"      Else ShowHTML "          <option value=""144"">Dados do site"      End If
    If Nvl(chave,"") = "221" Then ShowHTML "          <option value=""221"" SELECTED>Mensagens"          Else ShowHTML "          <option value=""221"">Mensagens"          End If
    If Nvl(chave,"") = "127" Then ShowHTML "          <option value=""127"" SELECTED>Modalidades"        Else ShowHTML "          <option value=""127"">Modalidades"        End If
    If Nvl(chave,"") = "222" Then ShowHTML "          <option value=""222"" SELECTED>Not�cias"           Else ShowHTML "          <option value=""222"">Not�cias"           End If
    ShowHTML "          </select>"
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da exibicao de bloco de dados
REM -------------------------------------------------------------------------
Function ExibeBlocoDados (chave)
    Select Case Chave
        Case 135 ExibeBlocoDados = "Arquivos"
        Case 133 ExibeBlocoDados = "Calend�rio"
        Case 220 ExibeBlocoDados = "Dados adicionais"
        Case 123 ExibeBlocoDados = "Dados b�sicos"
        Case 144 ExibeBlocoDados = "Dados do site"
        Case 221 ExibeBlocoDados = "Mensagens"
        Case 127 ExibeBlocoDados = "Modalidades"
        Case 222 ExibeBlocoDados = "Not�cias"
        Case Else  ExibeBlocoDados = "Consulta"
    End Select
End Function
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da sele��o de bloco de dados
REM -------------------------------------------------------------------------
Sub SelecaoOperacao (label, accesskey, hint, chave, chaveAux, campo, restricao, atributo)
    If IsNull(hint) Then
       ShowHTML "          <td valign=""top""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    Else
       ShowHTML "          <td valign=""top"" ONMOUSEOVER=""popup('" & hint & "','white')""; ONMOUSEOUT=""kill()""><font size=""1""><b>" & Label & "</b><br><SELECT ACCESSKEY=""" & accesskey & """ CLASS=""STS"" NAME=""" & campo & """ " & w_Disabled & " " & atributo & ">"
    End If
    ShowHTML "          <option value="""">Todas"
    If Nvl(chave,"") = "0" Then ShowHTML "          <option value=""0"" SELECTED>Consulta"           Else ShowHTML "          <option value=""0"">Consulta"           End If
    If Nvl(chave,"") = "1" Then ShowHTML "          <option value=""1"" SELECTED>Inclus�o"           Else ShowHTML "          <option value=""1"">Inclus�o"           End If
    If Nvl(chave,"") = "2" Then ShowHTML "          <option value=""2"" SELECTED>Altera��o"           Else ShowHTML "          <option value=""2"">Altera��o"           End If
    If Nvl(chave,"") = "3" Then ShowHTML "          <option value=""3"" SELECTED>Exclus�o"           Else ShowHTML "          <option value=""3"">Exclus�o"           End If
    ShowHTML "          </select>"
End Sub
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------

REM =========================================================================
REM Montagem da exibicao de opera��o
REM -------------------------------------------------------------------------
Function ExibeOperacao (chave)
    Select Case Chave
        Case 0    ExibeOperacao = "Consulta"
        Case 1    ExibeOperacao = "Inclus�o"
        Case 2    ExibeOperacao = "Altera��o"
        Case 3    ExibeOperacao = "Exclus�o"
        Case Else ExibeOperacao = "Erro"
    End Select
End Function
REM =========================================================================
REM Final da rotina
REM -------------------------------------------------------------------------
%>