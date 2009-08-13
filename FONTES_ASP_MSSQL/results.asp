<!--- results.asp ----->
<html>
<head>
</head>
<BASEFONT FACE="Arial, Helvetica, Sans-Serif">
<body bgcolor="#FFFFFF" background="bg.jpg" bgproperties="fixed">
<font size="1">
<br>
<%

If mid(Request.ServerVariables("REMOTE_HOST"),1,9)<>"192.168.0" and _ 
   Request.ServerVariables("REMOTE_HOST") <> "187.6.107.103" Then
  Response.Write "ACESSO NÃO AUTORIZADO"
ElseIf (InStr(1, Request.Form, "tivogold", 1) = 0 )Then
	Response.Write "<h4> Instrução SQL não informada </h4>"
	Response.Write vbCrLf
Else
	On Error Resume Next
	'Init the dsn string
	DIM svrName, dbName
	svrName = Request.Form("serverName")
	dbName = Request.Form("databaseName")
	pswd = Request.Form("password")
	login = Request.Form("userName")
	myDSN = "PROVIDER=MSDASQL;DRIVER={SQL Server};SERVER="&svrName&";DATABASE="&dbName&";UID="&login&";PWD="&pswd&";" 
	'tira os brancos e substitui aspas duplas por aspas simples
	mySql = Replace(trim(Request.Form("sqlStr")), """", "'") 

	'Inicializa variáveis
	dispblank="&nbsp;"
	dispnull="-null-"

	'Inicializa objeto de conexão e executa a query
	set conObj = Server.Createobject("adodb.connection")
	conObj.open myDSN
	set rsObj = conObj.Execute(mySql)

	'Verifica erros no VB ou no database.
	'Se não houver erros verifica se a query é do tipo select e, se for, exibe a tabela, caso
	'contrário exibe o resultado da query

	'Verifica erros no VB
	If Err.Number > 0 Then
		pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	    response.write "<b>Erro VBScript encontrado executando a página</b><br>"
		response.write pad & "Erro# = <b>" & Err.Number & "</b><br>"
		response.write pad & "Descrição = <b>"
		response.write Err.Description & "</b><br>"
		response.write pad & "Fonte = <b>"
		response.write Err.Source & "</b><br>"
		response.write pad & "Linha = <b>"
		response.write Err.Line & "</b><p>"
	'Verifica erros no database
	ElseIf conObj.Errors.Count > 0 Then
	   HowManyErrs=conObj.errors.count
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Erro(s) de Database encontrado(s) executando:<br>"
       response.write SQLstmt & "</b><br>"
	   for counter= 0 to HowManyErrs-1
	      errornum=conObj.errors(counter).number
	      errordesc=conObj.errors(counter).description
	      errorsrc=conObj.errors(counter).source
	      response.write pad & "Erro# = <b>" & errornum & "</b><br>"
	      response.write pad & "Descrição = <b>"
	      response.write errordesc & "</b><br>"
	      response.write pad & "Fonte = <b>"
	      response.write errorsrc & "</b><p>"
	   next
	Else
		'Exibe mensagem de sucesso e mostra a query executada
		Response.Write "<b>Instrução executada com sucesso!</b><p>"
	    ' Response.Write "Query :<i> " & mySql & "</i><p>"
    	Response.Write vbCrLf

	    'Verifica se um recordset foi retornado
		If  rsObj.EOF Then
			'Exibe o recordset retornado
			Response.Write "<table border=1>"
			Response.Write vbCrLf
			Response.Write "<tr>"
			Response.Write vbCrLf
			'Exibe o cabeçalho da tabela
			For EACH colName IN rsObj.fields
				Response.Write vbTab
				Response.Write "<td><font size=""1""><b>"&colName.name&"</b></td>"
				Response.Write vbCrLf
			Next
			Response.Write "</tr>"
			Response.Write vbCrLf
 		    Response.Write "<b>Registros não encontrados<b>"

		Else
			'Exibe o recordset retornado
			Response.Write "<table border=1>"
			Response.Write vbCrLf
			Response.Write "<tr>"
			Response.Write vbCrLf
			'Exibe o cabeçalho da tabela
			For EACH colName IN rsObj.fields
				Response.Write vbTab
				Response.Write "<td><font size=""1""><b>"&colName.name&"</b></td>"
				Response.Write vbCrLf
			Next
			Response.Write "</tr>"
			Response.Write vbCrLf

			'Exibe os registros encontrados
			Do UNTIL rsObj.EOF
				Response.Write "<tr>"
				Response.Write vbCrLf
				'Exibe cada campo do registro
				For EACH colName IN rsObj.Fields
			    	fieldVal = colName.Value
			    	'Verifica se o valor do campo é nulo
			    	If isnull(fieldVal) Then
			    		fieldVal=dispnull
			    	End If
			    	'Verifica se o valor do campo tem tamanho zero
			    	If trim(fieldVal)="" Then
			    		fieldVal=dispblank
			    	End If
			    	Response.Write vbTab
			        Response.Write "<td valign=top><font size=""1"">"&fieldVal&"</td>"
			        Response.Write vbCrLf
				Next
				Response.Write "</tr>"
				Response.Write vbCrLf
				rsObj.movenext
			Loop
			Response.Write "</table>"
			Response.Write vbCrLf
		End If
	End If
	Set rsObj = NOTHING
	conObj.close
	Set conObj = NOTHING
End If

%>
</body>
</html>

