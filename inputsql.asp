<%@ LANGUAGE="VBSCRIPT" %>
<html>
<head>
</head>
<BASEFONT FACE="Arial, Helvetica, Sans-Serif">
<body bgcolor="#FFFFFF" background="bg.jpg" bgproperties="fixed">
<STYLE> .BTM{font: 8pt Arial}</STYLE>
<center>
<%
if mid(Request.ServerVariables("REMOTE_HOST"),1,9)<>"192.168.0" and _ 
   mid(Request.ServerVariables("REMOTE_HOST"),1,11)<>"200.162.126" Then
%>ACESSO N�O AUTORIZADO<%
Else
%>
	<form name='AcceptQuery' action='results.asp' method='POST' target='content'>
	<center><font face="Times New Roman" size="3"><b>Controle Central BD</b></font></center>
		<table width="100%">
		<tr><td valign="top"><table>
		<tr>
			<td align="right"><font size="1">Servidor</td>
			<td><input type=text name='serverName' CLASS=BTM></td>
		<tr>
			<td align="right"><font size="1">Database</td>
			<td><input type=text name='databaseName' CLASS=BTM></td>
		</tr>
		<tr>
			<td align="right"><font size="1">Usu�rio</td>
			<td><input type=text name='userName' CLASS=BTM></td>
		<tr>
			<td align="right"><font size="1">Senha</td>
			<td><input type=password name='password' CLASS=BTM></td>
		</tr>
		</table><td rowspan="2"><table>
		<tr>
			<td align=center>
				<input type='hidden' name='tivohidden' value='tivogold'>
				<textarea name="sqlStr" rows=9 cols=100 wrap="soft" CLASS=BTM></textarea>
			</td>
			<td align="center" valign="bottom"><input type='submit' name='exec' value='exec' CLASS=BTM onClick='sqlStr.focus()'></td>
		</tr>
		</table>
		</table>
	</form>
</body>
</html>
<%End If
%>