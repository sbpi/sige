<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="Constants_ADO.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="Funcoes.asp"-->
<%
Server.ScriptTimeout = 300
REM =========================================================================
REM  /GeraInserts.asp
REM ------------------------------------------------------------------------
REM Nome     : Alexandre Vinhadelli Papadópolis
REM Descricao: Página do portal 
REM Home     : http://www.sbpi.com.br/
REM Criacao  : 19/12/2008 17:21
REM Autor    : SBPI
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by SBPI - Sociedade Brasileira para a Pesquisa em Informática
REM -------------------------------------------------------------------------

Private sobjRS
Private sobjRS1
Private sobjConn
Private sstrSN
Private sstrES
Private sstrCA
Private sstrCL
Private strTitulo

sstrSN = "GeraInserts.asp"

Public sstrEF
Public sstrEW
Public sstrIN
Public sstrDiretorio
Public sstrModelo
  
Set sobjRS = Server.CreateObject("ADODB.RecordSet")
Set sobjRS1 = Server.CreateObject("ADODB.RecordSet")
Set sobjConn  = Server.CreateObject("ADODB.Connection")

sobjConn.ConnectionTimeout = 300
sobjConn.CommandTimeout = 300
sobjConn.Open conConnectionString

GeraInserts

REM =========================================================================
REM Gera os Diretórios
REM -------------------------------------------------------------------------
Public Sub GeraInserts

CONST ForReading = 1, ForWriting = 2, ForAppend = 8
CONST TristateUsedefault = -2 'Abre o arquivo usando o sistema default
CONST TristateTrue = -1 'Abre o arquivo como Unicode
CONST TristateFalse = 0 'Abre o arquivo como ASCII
DIM   sql[80], wCont, wCriado, wAlterado, w_unidade


w_unidade = Request("unidade")
w_todos   = Request("todos")

  
' Recupera as usernames de escCliente
If w_unidade > "" Then
   sql = "SELECT a.sq_cliente, a.ds_username,b.sq_modelo from escCliente as a inner join escCliente_Site as b on (a.sq_cliente = b.sq_cliente) where upper(a.ds_username) = upper('" & w_unidade & "')"
Else
   sql = "SELECT a.sq_cliente, a.ds_username,b.sq_modelo from escCliente as a inner join escCliente_Site as b on (a.sq_cliente = b.sq_cliente) order by ds_username"
End If
sobjRS.Open sql, sobjConn, adOpenForwardOnly

' Inicio ****************************************
' Gera o arquivo default.asp do cliente, a partir do default.asp do modelo de site escolhido
Dim FS, f1, f2, StrLine, strDir, strFile, strArq, strMod, sqCliente

wCont     = 0
wCriado   = 0
wAlterado = 0
While Not sobjRS.EOF

    strDir = "sedf\" & sobjRS("ds_username")
    strArq = "default.asp"
    strMod = sobjRS("sq_modelo")
    sqCliente = sobjRS("sq_cliente")
    strFile = replace(conFilePhysical & "\" & strDir & "\" & strArq,"\\","\")
    strDir = replace(conFilePhysical & "\" & strDir,"\\","\")
    strMod = replace(conFilePhysical & "\" & "Modelos\Mod" & strMod & "\Default_cliente.asp","\\","\")
    Set FS = CreateObject("Scripting.FileSystemObject")
    If (Not FS.FolderExists (strDir)) or w_unidade > "" or w_todos = "Sim" then
       If (Not FS.FolderExists (strDir)) Then
          'Response.Write strdir
          'Response.End()
          Set f1 = FS.CreateFolder(strDir)
          wCriado = wCriado + 1
       End If

       If FS.FileExists (strFile) then
          FS.DeleteFile (strFile)
          wAlterado = wAlterado + 1
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
    Elseif FS.FolderExists (strDir) Then
       If (Not FS.FileExists (strFile)) then
          Set f1 = FS.CreateTextFile(strFile)

          Set f2 = FS.OpenTextFile(strMod, ForReading)

          Do While Not f2.AtEndOfStream 
             strLine = f2.ReadLine  
             strLine = replace(strLine,"= *%*","= " & sqCliente)
             f1.WriteLine strLine
          Loop
          wAlterado = wAlterado + 1
                    
          f2.Close
          f1.Close
       End If
    End If
    wCont = wCont + 1
    sobjRS.MoveNext
    
Wend

SET f1 = Nothing
SET f2 = Nothing
SET FS = Nothing
' Fim ****************************************


ShowHTML "<HTML>"
ShowHTML "<TITLE>Geração de diretórios</TITLE>"

ShowHTML "<BASEFONT FACE=""Arial, Helvetica, Sans-Serif"">"
ShowHTML "<body topmargin=""0"" leftmargin=""0"" bgcolor=""#F0FFFF"">"
ShowHTML "<font size=""5"" color=""#0000FF"">Sucesso!</font><hr><font size=""2"">"
ShowHTML "<p align=""JUSTIFY"">Foram processadas " & wCont & " escolas.</p> " 
ShowHTML "<p align=""JUSTIFY""> " & wCriado & " novas escolas foram inseridas.</p> " 
ShowHTML "<p align=""JUSTIFY""> " & wAlterado & " escolas existentes tiveram sua página inicial recriada (demais arquivos mantidos).</p> " 
ShowHTML "</BODY>"
ShowHTML "</HTML>"

End Sub
REM -------------------------------------------------------------------------
REM Final da geração dos diretórios
REM =========================================================================


%>