<?php
// Garante que a sessão será reinicializada.
session_start();

$w_dir_volta = '';
$w_pagina       = 'controle.php?par=';
$w_disabled     = 'ENABLED';
$w_dir          = '';
$w_troca        = $_REQUEST['w_troca'];

if (isset($_SESSION['LOGON1'])) {
    echo '<SCRIPT LANGUAGE="JAVASCRIPT">';
    echo ' alert("Já existe outra sessão ativa!\nEncerre o sistema na outra janela do navegador ou aguarde alguns instantes.\nUSE SEMPRE A OPÇÃO \"SAIR DO SISTEMA\" para encerrar o uso da aplicação.");';
    echo ' history.back();';
    echo '</SCRIPT>';
    exit();
}

// Define o banco de dados a ser utilizado
$_SESSION['DBMS']=1;

include_once('constants.inc');
include_once('jscript.php');
include_once('funcoes.php');
include_once($w_dir_volta.'classes/db/abreSessao.php');
include_once($w_dir_volta.'classes/sp/db_exec.php');

// Abre conexão com o banco de dados
if (isset($_SESSION['DBMS'])){
    $dbms = abreSessao::getInstanceOf($_SESSION['DBMS']);
}

// =========================================================================
//  /controle.php
// ------------------------------------------------------------------------
// Nome     : Cesar Martin
// Descricao: Página de controle da aplicação
// Mail     : cesar@sbpi.com.br
// Criacao  : 16/03/2005 16:14PM
// Versao   : 1.0.0.0
// Local    : Brasília - DF
// -------------------------------------------------------------------------
//
// Declaração de variáveis

// Carrega variáveis de controle
$RS  = null;
$CL  = $_SESSION['CL'];
$par = substr(strtoupper($_REQUEST['par']),0,30);
$O   = substr(strtoupper($_REQUEST['O']),0,1);
if(nvl($O,'') == ''){
  $O = 'L';
}
$w_Data = Substr(100+Day(Time()),1,2) . "/" . Substr(100+Month(Time()),1,2) . "/" . Year(Time());

if (nvl($par,'')=='') Logon();
else main();

FechaSessao($dbms);

exit;

// =========================================================================
// Rotina de criação da tela de logon (backup)
// -------------------------------------------------------------------------
function LogOn() {
  extract($GLOBALS);

  $w_username = $_REQUEST['Login'];
  ShowHTML ('<HTML>');
  ShowHTML ('<HEAD>');
  ShowHTML ('<link rel="shortcut icon" href="'.$conRootSIW.'favicon.ico" type="image/ico" />');
  ShowHTML ('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
  ShowHTML ('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
  ShowHTML ('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
  ShowHTML ('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
  ShowHTML ('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
  ShowHTML ('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
  ShowHTML ('<TITLE>'.$conSgSistema.' - Autenticação</TITLE>');
  ScriptOpen('JavaScript');
  ShowHTML ('$(document).ready(function(){');
  ShowHTML ('	$("#Login1").change(function(){');
  ShowHTML ('		formataCampo();');
  ShowHTML ('	})');
  ShowHTML ('});');
  ShowHTML ('function formataCampo(){');
  ShowHTML ('	$("#Login1").val(trim($("#Login1").val()));');
  ShowHTML ('	if(  $("#Login1").val().length==11 &&  caracterAceito( $("#Login1").val() ,  "0123456789") ){');
  ShowHTML ('		$("#Login1").val( mascaraGlobal(\'###.###.###-##\',$("#Login1").val()) );');
  ShowHTML ('	}');
  ShowHTML ('}');
  ShowHTML ('function caracterAceito(string , checkOK){');
  ShowHTML ('	 //var checkOK = \'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-!@#$%&*()+/\';');
  ShowHTML ('		  var checkStr = string;');
  ShowHTML ('		  var allValid = true;');
  ShowHTML ('		  for (i = 0;  i < checkStr.length;  i++)');
  ShowHTML ('		  {');
  ShowHTML ('			ch = checkStr.charAt(i);');
  ShowHTML ('			if ((checkStr.charCodeAt(i) != 13) && (checkStr.charCodeAt(i) != 10) && (checkStr.charAt(i) != "\\\\")) {');
  ShowHTML ('			   for (j = 0;  j < checkOK.length;  j++) {');
  ShowHTML ('				 if (ch == checkOK.charAt(j))');
  ShowHTML ('				   break;');
  ShowHTML ('			   }');
  ShowHTML ('			   if (j == checkOK.length)');
  ShowHTML ('			   {');
  ShowHTML ('				 allValid = false;');
  ShowHTML ('				 break;');
  ShowHTML ('			   }');
  ShowHTML ('			}');
  ShowHTML ('		  }');
  ShowHTML ('		  return allValid;');
  ShowHTML ('}');
  ShowHTML ('function Ajuda() ');
  ShowHTML ('{ ');
  ShowHTML ('  document.Form.Botao.value = "Ajuda"; ');
  ShowHTML ('} ');
  Modulo();
  SaltaCampo();
  ValidateOpen('Validacao');
  Validate('Login1','Nome de usuário','','1','2','30','1','1');
  Validate('Password1','Senha','1','1','3','19','1','1');
  ShowHTML ('  theForm.Login.value = theForm.Login1.value; ');
  ShowHTML ('  theForm.Password.value = theForm.Password1.value; ');
  ShowHTML ('  theForm.Login1.value = ""; ');
  ShowHTML ('  theForm.Password1.value = ""; ');
  ValidateClose();
  ScriptClose();
  ShowHTML ('<link rel="stylesheet" type="text/css" href="'.$conRootSIW.'classes/menu/xPandMenu.css">');
  ShowHTML ('<style>');
  ShowHTML (' .cText {font-size: 8pt; border: 1px solid #000000; background-color: #F5F5F5}');
  ShowHTML (' .cButton {font-size: 8pt; color: #FFFFFF; border: 1px solid #000000; background-color: #669966; }');
  ShowHTML ('</style>');
  ShowHTML ('</HEAD>');
  // Se receber a username, dá foco na senha
  if (nvl($w_username,'nulo')=='nulo') {
      ShowHTML ('<body topmargin=0 leftmargin=10 onLoad=\'document.Form.Login1.focus();\'>');
  } else {
      ShowHTML ('<body topmargin=0 leftmargin=10 onLoad=\'document.Form.Password1.focus();\'>');
  }
  ShowHTML ('<CENTER>');
  ShowHTML ('<form method="post" action="controle.php?par=valida" onsubmit="return(Validacao(this));" name="Form"> ');
  ShowHTML ('<INPUT TYPE="HIDDEN" NAME="Login" VALUE=""> ');
  ShowHTML ('<INPUT TYPE="HIDDEN" NAME="Password" VALUE=""> ');
  ShowHTML ('<INPUT TYPE="HIDDEN" NAME="p_dbms" VALUE="1"> ');
  ShowHTML ('<TABLE cellSpacing=0 cellPadding=0 width="760" height=550 border=1  background="files/" & p_cliente & "/img/fundo.jpg" bgproperties="fixed"><tr><td width="100%" valign="top">');
  ShowHTML ('  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 background="img/cabecalho.gif">');
  ShowHTML ('    <TR>');
  ShowHTML ('      <TD vAlign=TOP align=middle width="65%">');
  ShowHTML ('        <B><FONT face=Arial size=5 color=#000088>Secretaria de Estado de Educação');
  ShowHTML ('        <br><br>SIGE - Módulo Web</B></FONT>');
  ShowHTML ('     </TD>');
  ShowHTML ('    </TR>');
  ShowHTML ('    <TR><TD colspan=3 borderColor=#ffffff height=22><HR align=center color=#808080></TD></TR>');
  ShowHTML ('  </TABLE>');
  ShowHTML ('  <table width="100%" border="0">');
  ShowHTML ('    <tr><td valign="middle" width="100%" height="100%">');
  ShowHTML ('        <table width="100%" height="100%" border="0">');
  ShowHTML ('          <tr><td align="center" colspan=2><font size="1" color="#990000"><b>Esta aplicação é de uso interno da Secretaria de Estado de Educação.<br>As informações contidas nesta aplicação são restritas e de uso exclusivo.<br>O uso indevido acarretará ao infrator penalidades de acordo com a legislação em vigor.<br><br>Informe seu nome de usuário, senha de acesso e clique no botão <i>OK</i> para ser autenticado pela aplicação.</b></font>');
  ShowHTML ('          <tr><td align="right" width="43%"><font size="2"><b>Nome de usuário:<td><input class="sti" name="Login1" size="14" maxlength="14">');
  ShowHTML ('          <tr><td align="right"><font size="2"><b>Senha:<td><input class="sti" type="Password" name="Password1" size="19">');
  ShowHTML ('          <tr><td align="right"><td><font size="2"><b><input class="stb" type="submit" value="OK" name="Botao"> ');
  ShowHTML ('          </font></td> </tr> ');
  ShowHTML ('          <TR><TD colspan=2 align="center"><br><table border=0 cellpadding=0 cellspacing=0><tr><td>');
  ShowHTML ('              <P><IMG height=37 src="img/ajuda.jpg" width=629><br>');
  ShowHTML ('              <font face="Arial" size=1><b>PARA ACESSAR A PÁGINA DE ATUALIZAÇÃO</b></font>');
  ShowHTML ('              <FONT face="Verdana, Arial, Helvetica, sans-serif" size=1>');
  ShowHTML ('              <li>Nome de usuário - Informe seu nome de usuário');
  ShowHTML ('              <li>Senha - Informe sua senha de acesso');
  ShowHTML ('              <li>Se esqueceu ou não foi informado dos dados acima, favor entrar em contato com a SEDF / SUBIP / Diretoria de Sistemas de Informação Educacional - DSIE');
  ShowHTML ('              </FONT></P>');
  ShowHTML ('              <P><font face="Arial" size=1><b>DOCUMENTAÇÃO - LEIA COM ATENÇÃO</b></font><br>');
  ShowHTML ('              <FONT face="Verdana" size=1>');
  ShowHTML ('              . <a class="SS" href="sedf/Orientacoes_Acesso.pdf" target="_blank" title="Abre arquivo que descreve as novas características e funcionalidades do SIGE-WEB.">Apresentação da nova versão do SIGE-WEB (PDF - 130KB - 4 páginas)</a><BR>');
  ShowHTML ('              . <a class="SS" href="manuais/operacao/" target="_blank" title="Exibe manual de operação do SIGE-WEB">Manual SIGE-WEB (HTML)</A><BR>');
  ShowHTML ('              <br></FONT></P>');
  ShowHTML ('              </TD></TR>');
  ShowHTML ('          </table> ');
  ShowHTML ('        </table> ');
  ShowHTML ('    </tr> ');
  ShowHTML ('  </table>');
  ShowHTML ('</table>');
  ShowHTML ('</form> ');
  ShowHTML ('</CENTER>');
  ShowHTML ('</body>');
  ShowHTML ('</html>');
}

// =========================================================================
// Rotina de autenticação dos usuários
// -------------------------------------------------------------------------
function Valida(){
  extract($GLOBALS);

  $w_uid = str_replace('""',"",str_replace("'","",Trim(strtoupper($_REQUEST["Login"]))));
  $w_pwd = str_replace('""',"",str_replace("'","",Trim($_REQUEST["Password"])));
  $w_erro = 0;
  $SQL = "select count(*) existe from sbpi.Cliente where upper(ds_username) = upper('" . $w_uid . "')";
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows); 
  foreach($RS as $row) { $RS = $row; break; }

  if (f($RS,"existe") == 0){
      $w_erro = 1;
  } else {
      $SQL = "select count(*) existe " . $crlf . 
             "  from sbpi.Cliente           a " . $crlf . 
             "       join sbpi.Tipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente and " . $crlf . 
             "                                  b.tipo            = 1 " . $crlf . 
             "                                 ) " . $crlf . 
             " where upper(ds_username) = upper('" . $w_uid . "')";
      $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
      foreach($RS as $row) { $RS = $row; break; }
      if(f($RS,"existe") == 0){
          $w_erro = 3;
      }else{
          $SQL = "select count(*) existe from sbpi.Cliente where upper(ds_username) = upper('" . $w_uid . "') and ds_senha_acesso = '" . $w_pwd . "'";
          $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows); 
          foreach($RS as $row) { $RS = $row; break; }
          if(f($RS,"existe") == 0){
              $w_erro = 2;
          }
      }        
  }
  ScriptOpen ('JavaScript');
  if ($w_erro > 0) {
    if ($w_erro == 1)     ShowHTML (' alert("Usuário inexistente!");');
    elseif ($w_erro == 2) ShowHTML (' alert("Senha inválida!");');
    else                  ShowHTML (' alert("Usuário sem permissão para acessar esta página!");');
    ShowHTML ('  history.back(1);');
  } else {
    //Recupera informações a serem usadas na montagem das telas para o usuário
    $SQL = "select ds_username, sq_cliente from sbpi.Cliente where upper(ds_username) = upper('" . $w_uid . "') and ds_senha_acesso = '" . $w_pwd . "'";
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows); 
    foreach($RS as $row) { $RS = $row; break; }
    $_SESSION['USERNAME'] = strtoupper(f($RS, 'ds_username'));
    $_SESSION['CL']       = f($RS, 'sq_cliente');
    
    If ($_SESSION['USERNAME'] != 'SBPI'){
      //Grava o acesso na tabela de log
      $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql) " . $crlf . 
            "values ( " . $crlf . 
            "         sbpi.sq_cliente_log.nextval, " . $crlf . 
            "         " . $_SESSION['CL'] . ", " . $crlf . 
            "         sysdate, " . $crlf . 
            "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf . 
            "         0, " . $crlf . 
            "         'Usuário " . $w_uid . " - acesso à tela de manutenção dos dados da rede de ensino.', " . $crlf . 
            "         null " . $crlf . 
            "       ) " . $crlf;
      $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    }
    ShowHTML (' location.href="controle.php?par=frames";');
  }
  ScriptClose();
}

// =========================================================================
// Exportação dos dados administrativos
// -------------------------------------------------------------------------
function administrativo(){
  extract($GLOBALS);
  Cabecalho();
  ShowHTML ('<HEAD>');
  ScriptOpen('Javascript');
  ValidateOpen('Validacao');
  ShowHTML ('  if (theForm.w_arquivo[0].checked == false && theForm.w_arquivo[1].checked == false) {');
  ShowHTML ('     alert(\'Você deve escolher uma das opções apresentadas antes de gerar o arquivo!\');');
  ShowHTML ('     return false;');
  ShowHTML ('  }');
  ShowHTML ('  return(confirm(\'Confirma a geração do arquivo com os dados indicados?\'));');
  ValidateClose();
  ScriptClose();
  ShowHTML ('</HEAD>');
  BodyOpen('onLoad=\'document.focus();\'');
  ShowHTML ('<B><FONT COLOR="#000000">Exportação dos dados administrativos</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
  AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
  ShowHTML(MontaFiltro('POST'));
  ShowHTML ('<input type="hidden" name="R" value="'.$R.'">');
  ShowHTML ('<tr bgcolor="'.'#EFEFEF'.'"><td align="center">');
  ShowHTML ('    <table width="97%" border="0">');
  ShowHTML ('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Exportação dos dados administrativos</td></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td><font size=1><ul><li>Esta tela permite a exportação, para um arquivo que pode ser aberto no Excel, dos dados administrativos preenchidos pelas unidades de ensino.<li>Permite também exportar as tabelas de apoio utilizadas pelo formulário.<li>Selecione uma das opções exibidas abaixo e clique no botão "Gerar arquivo" para que os dados sejam convertidos para um arquivo.</ul></font></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top"><font size="1"><b>O arquivo a ser gerado deve conter dados:</b>');
  ShowHTML ('          <br><INPUT '.$w_disabled.' class="BTM" type="radio" name="w_arquivo" value="Escola"> das unidades de ensino');
  ShowHTML ('          <br><INPUT '.$w_disabled.' class="BTM" type="radio" name="w_arquivo" value="Tipo"> da tabela de equipamentos');
  ShowHTML ('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');

  // Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
  ShowHTML ('      <tr><td align="center"><input class="STB" type="submit" name="Botao" value="Gerar arquivo"></td>');
  ShowHTML ('      </tr>');
  ShowHTML ('    </table>');
  ShowHTML ('    </TD>');
  ShowHTML ('</tr>');
  ShowHTML ('</FORM>');
  ShowHTML ('</table>');
  Rodape();
}


// =========================================================================
// Monta a Frame
// -------------------------------------------------------------------------
function frames(){
    extract($GLOBALS);
    $_SESSION["BodyWidth"] = "620";

    ShowHTML ('<html>');
    ShowHTML ('<head>');
    ShowHTML ('    <title>Controle Central</title>');
    ShowHTML ('</head>');

    ShowHTML ('<frameset cols="200,*">');
    ShowHTML ('    <frame name="menu" src="controle.php?par=menu" scrolling="auto" marginheight="0" marginwidth="0">');
    ShowHTML ('    <frame name="content" src="controle.php?par=escolas" scrolling="auto" marginheight="0" marginwidth="0">');
    ShowHTML ('</frameset>');
    ShowHTML ('</html>');
}

// =========================================================================
// Cadastro de modalidades de ensino
// -------------------------------------------------------------------------
function Modalidades(){

  extract($GLOBALS);
  
  $w_chave            = $_REQUEST["w_chave"];
  $w_troca = $_REQUEST["w_troca"];
  $w_ew    = '&w_ew=127';
  
  $SQL = "Select nr_ordem, ds_especialidade from sbpi.Especialidade order by nr_ordem, ds_especialidade";  
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows); 
  foreach($RS as $row) { $RS = $row; break; }  
  if (count($RS)>0) {
     $w_texto = '<b>Nºs de ordem em uso para esta subordinação:</b>:<br>' .
                '<table border=1 width=100% cellpadding=0 cellspacing=0>' .
                '<tr><td align=center><b><font size=1>Ordem' .
                '    <td><b><font size=1>Descrição';
    foreach($RS as $row) {
      $w_texto .= '<tr><td valign=top align=center><font size=1>' . f($row, 'nr_ordem') . '<td valign=top><font size=1>' . f($row, 'ds_especialidade');
    }
  $w_texto .= "</table>";
  }else{
     $w_texto = "Não há outros números de ordem vinculados à subordinação desta opção";
  }

  If ($w_troca > '') { // Se for recarga da página
    $w_ds_especialidade = $_REQUEST['w_ds_especialidade'];
    $w_nr_ordem         = $_REQUEST['w_nr_ordem'];    
  }else if($O == "L"){
    //Recupera todos os registros para a listagem
    $SQL = "select nr_ordem, ds_especialidade, sq_especialidade  from sbpi.Especialidade order by nr_ordem, ds_especialidade";
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows); 
  }else if(strpos("AEV",$O) !== false and $w_troca == ''){
    //Recupera os dados do endereço informado
    $SQL = "select nr_ordem, ds_especialidade, sq_especialidade from sbpi.Especialidade where sq_especialidade = " . $w_chave;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows); 
    foreach($RS as $row) { $RS = $row; break; }
    $w_ds_especialidade = f($RS, "ds_especialidade");
    $w_nr_ordem         = f($RS, "nr_ordem");
  }  
  
  Cabecalho();
  ShowHTML ('<HEAD>');
  if (strpos("IAEP",$O) > 0){
    ScriptOpen('JavaScript');
    CheckBranco();
    FormataData();
    ValidateOpen("Validacao");
    if (strpos("IA",$O) > 0){
    Validate ("w_ds_especialidade" , "Descrição"    , "" , "1" , "2" , "70" , "1" , "1");
    Validate ("w_nr_ordem"         , "Nr. de ordem" , "" , "1" , "1" , "4"  , ""  , "0123546789");
    }
    ShowHTML (' theForm.Botao[0].disabled=true;');
    ShowHTML ('  theForm.Botao[1].disabled=true;');
    ValidateClose();
    ScriptClose();
  }

  ShowHTML ('</HEAD>');
  if($w_troca > ""){
     BodyOpen ('onLoad=\'document.Form.' . $w_troca . '.focus()\';');
  } else if($O == "I" or $O == "A"){
     BodyOpen ("onLoad='document.Form.w_ds_especialidade.focus()';");
  } else {
     BodyOpen ("onLoad='document.focus()';");
  }
  ShowHTML ('<B><FONT COLOR="#000000">Cadastro de modalidades de ensino</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
  If ($O == 'L'){
    // Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML ('<tr><td><font size="2"><a accesskey="I" class="SS" href="' .$dir.$w_pagina.$par. $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
    ShowHTML ('    <td align="right"><b>Registros existentes: '.count($RS));
    ShowHTML ('<tr><td align="center" colspan=3>');
    ShowHTML ('    <TABLE WIDTH="100%" bgcolor="'.$conTableBgColor.'" BORDER="'.$conTableBorder.'" CELLSPACING="'.$conTableCellSpacing.'" CELLPADDING="'.$conTableCellPadding.'" BorderColorDark="'.$conTableBorderColorDark.'" BorderColorLight="'.$conTableBorderColorLight.'">');
    ShowHTML ('        <tr bgcolor="'.$conTrBgColor.'" align="center">');
    ShowHTML ('          <td><font size="1"><b>Ordem</font></td>');
    ShowHTML ('          <td><font size="1"><b>Modalidade</font></td>');
    ShowHTML ('          <td><font size="1"><b>Operações</font></td>');
    ShowHTML ('        </tr>');
    if (count($RS)<=0) {
      // Se não foram selecionados registros, exibe mensagem
      ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
    } else {
      foreach($RS as $row){
        $w_cor = ($w_cor==$conTrBgColor || $w_cor=='') ? $w_cor=$conTrAlternateBgColor : $w_cor=$conTrBgColor;
        ShowHTML ('      <tr bgcolor="' . $w_cor . '" valign="top">');
        if (Nvl(f($row, "nr_ordem"),"nulo") <> "nulo"){
           ShowHTML ('        <td align="CENTER"><font size="1">' . f($row, "nr_ordem") . '</td>');
        } else {
           ShowHTML ('        <td align="center"><font size="1">---</td>');
        }
        ShowHTML ('        <td><font size="1">' . f($row, "ds_especialidade") . '</td>');
        ShowHTML ('        <td align="top" nowrap><font size="1">');
        ShowHTML ('          <A class="HL" HREF="' . $w_pagina .'modalidades' . $w_ew . "&R=" . $w_pagina .'modalidades' . $w_ew . "&O=A&CL=" . $CL . "&w_chave=" . f($row, "sq_especialidade") . '">Alterar</A>&nbsp');
        ShowHTML ('          <A class="HL" HREF="' . $w_pagina . "GRAVA&R=" . $w_ew . "&O=E&CL=" . $CL . "&w_sq_cliente=" . str_replace($CL,"sq_cliente=","") . "&w_chave=" . f($row, "sq_especialidade") . 'onClick=\'return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
        ShowHTML ('        </td>');
        ShowHTML ('      </tr>');
      }
    }
    ShowHTML ('      </center>');
    ShowHTML ('    </table>');
    ShowHTML ('  </td>');
    ShowHTML ('</tr>');
  }else if (strpos('IAEV',$O)!==false){
    If (strpos('EV',$O)){
       $w_disabled = ' DISABLED ';
  }
    AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
    ShowHTML ('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
    ShowHTML ('<INPUT type="hidden" name="CL" value="' . $CL . '">');
    ShowHTML ('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
    ShowHTML ('<INPUT type="hidden" name="w_sq_cliente" value="' . str_replace('sq_cliente=',$CL,'sq_cliente=') . '">');
    ShowHTML ('<INPUT type="hidden" name="O" value="' . $O . '">');
    
    ShowHTML ('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
    ShowHTML ('    <table width=""95%"" border=""0"">');
    ShowHTML ('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML ('        <tr valign="top"><td valign="top"><font size="1"><b>D<u>e</u>scrição:</b><br><input ' . $w_disabled . ' accesskey="E" type="text" name="w_ds_especialidade" class="STI" SIZE="70" MAXLENGTH="70" VALUE="' . $w_ds_especialidade . '"></td>');
    ShowHTML ('              <td valign="top" align="left"><font size="1"><b><u>O</u>rdem:<br><INPUT ACCESSKEY="O" TYPE="TEXT" CLASS="STI" NAME="w_nr_ordem" SIZE=4 MAXLENGTH=4 VALUE="' . $w_nr_ordem . '" " . $w_disabled . "></td>');
    ShowHTML ('        </table>');
    ShowHTML ('      <tr>');
    ShowHTML ('      <tr><td align="center" colspan=4><hr>');
    If ($O == "E"){
       ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir">');
    } else {
       If ($O == "I"){
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Incluir">');
       } else {
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
       }
    }
    ShowHTML ('            <input class="STB" type="button" onClick="location.href=\'' .$dir.$w_pagina.$par. $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
    ShowHTML ('          </td>');
    ShowHTML ('      </tr>');
    ShowHTML ('    </table>');
    ShowHTML ('    </TD>');
    ShowHTML ('</tr>');
    ShowHTML ('</FORM>');
  }else{
    ScriptOpen('JavaScript');
    ShowHTML (' alert(\'Opção não disponível\');');
    ShowHTML (' history.back(1);');
    ScriptClose();
  }
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();

}
// =========================================================================
// Fim do cadastro de modalidades de ensino
// -------------------------------------------------------------------------

// =========================================================================
// Cadastro de calendario base
// -------------------------------------------------------------------------
function calend_base(){
  extract($GLOBALS);
  
  $w_chave            = $_REQUEST["w_chave"];
  $w_troca            = $_REQUEST["w_troca"];
  
  if ( $w_troca > "" ){ // Se for recarga da página
     $w_dt_ocorrencia = $_REQUEST["w_dt_ocorrencia"];
     $w_ds_ocorrencia = $_REQUEST["w_ds_ocorrencia"];
     $w_tipo          = $_REQUEST["w_tipo"];
  } else if ( $O == "L" ){
     //Recupera todos os registros para a listagem
     $SQL = 'select a.*, b.nome from sbpi.calendario_base a left join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) order by sbpi.year(dt_ocorrencia) desc, dt_ocorrencia';
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);	 
  } else if ( strpos("AEV",$O) !== false and $w_troca == "" ){
     //Recupera os dados do endereço informado
     $SQL = "select * from sbpi.calendario_base where sq_ocorrencia = " . $w_chave;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);	 
     foreach($RS as $row) { $RS = $row; break; }
     $w_dt_ocorrencia = FormataDataEdicao(f($row, "dt_ocorrencia"));
     $w_ds_ocorrencia = f($row, "ds_ocorrencia");
     $w_tipo          = f($row, "sq_tipo_data");
  }
  
  Cabecalho();
  ShowHTML ('<HEAD>');
  if ( strpos("IAEP",O) !== false ){
     ScriptOpen ('JavaScript');
     CheckBranco();
     FormataData();
     ValidateOpen("Validacao");
     if ( strpos("IA",$O) !== false ){
        Validate ("w_dt_ocorrencia" , "Data"      , "DATA" , "1" , "10" , "10" , "1" , "1");
        Validate ("w_ds_ocorrencia" , "Descrição" , ""     , "1" , "2"  , "60" , "1" , "1");
        Validate ("w_tipo" , "Tipo" , "SELECT"     , "1" , "1"  , "4" , "" , "1");
     }
     ShowHTML ('  theForm.Botao[0].disabled=true;');
     ShowHTML ('  theForm.Botao[1].disabled=true;');
     ValidateClose();
     ScriptClose();
  }
  ShowHTML ('</HEAD>');
  if ( $w_troca > "" ){
     BodyOpen ('onLoad=\'document.Form.' . $w_troca . '.focus()\';');
  } else if ( $O == "I" or $O == "A" ){
     BodyOpen ('onLoad=\'document.Form.w_dt_ocorrencia.focus()\';');
  } else {
     BodyOpen ('onLoad=\'document.focus()\';');
  }
  ShowHTML ('<B><FONT COLOR=""#000000"">Cadastro do calendário oficial</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">');
  if ( $O == "L" ){
    //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML ('<tr><td><font size="2"><a accesskey="I" class="SS" href="' .$dir.$w_pagina.$par. $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
    ShowHTML ('<tr><td align="center" colspan=3>');
    ShowHTML ('    <TABLE WIDTH="100%" bgcolor="'.$conTableBgColor.'" BORDER="'.$conTableBorder.'" CELLSPACING="'.$conTableCellSpacing.'" CELLPADDING="'.$conTableCellPadding.'" BorderColorDark="'.$conTableBorderColorDark.'" BorderColorLight="'.$conTableBorderColorLight.'">');
    ShowHTML ('        <tr bgcolor="'.$conTrBgColor.'" align="center">');
    ShowHTML ('          <td><font size="1"><b>Data</font></td>');
    ShowHTML ('          <td><font size="1"><b>Tipo</font></td>');
    ShowHTML ('          <td><font size="1"><b>Ocorrência</font></td>');
    ShowHTML ('          <td><font size="1"><b>Operações</font></td>');
    ShowHTML ('        </tr>');

    if (count($RS)<=0) {
      // Se não foram selecionados registros, exibe mensagem
      ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
    } else {
      foreach($RS as $row) {      
        $w_cor = ($w_cor==$conTrBgColor || $w_cor=='') ? $w_cor=$conTrAlternateBgColor : $w_cor=$conTrBgColor;
        if ( $wAno != year(f($row, "dt_ocorrencia")) ){
           ShowHTML ('      <tr bgcolor=""#C0C0C0"" valign=""top""><TD colspan=4 align="center"><font size=2><B>' . year(f($row, "dt_ocorrencia")) . '</b></font></td></tr>');
           $wAno = year(f($row, "dt_ocorrencia"));
        }
        ShowHTML ('      <tr bgcolor="'. $w_cor . '" valign="top">');
        ShowHTML ('        <td align="center"><font size="1">' . Substr(FormataDataEdicao(FormatDateTime(f($row, "dt_ocorrencia"),2)),0,5) . '</td>');
        ShowHTML ('        <td><font size="1">' . nvl(f($row, "nome"),"---") . '</td>');
        ShowHTML ('        <td><font size="1">' . f($row, "ds_ocorrencia") . '</td>');
        ShowHTML ('        <td align="top" nowrap><font size="1">');
        ShowHTML ('          <A class="HL" HREF="' . $w_pagina .'calend_base' . $w_ew . '&R=' . $w_pagina .'calend_base' . $w_ew . '&O=A&CL=' . $CL . '&w_chave=' . f($row, "sq_ocorrencia") . '">Alterar</A>');
        ShowHTML ('          <A class="HL" HREF="' . $w_pagina . "GRAVA&R=" . $w_ew . '&O=E&CL=' . $CL . '&w_sq_cliente=' . str_replace($CL,"sq_cliente=","") . '&w_chave=' . f($row, "sq_ocorrencia") . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
        ShowHTML ('        </td>');
        ShowHTML ('      </tr>');
      }
    }
    ShowHTML ('      </center>');
    ShowHTML ('    </table>');
    ShowHTML ('  </td>');
    ShowHTML ('</tr>');
  } else if ( strpos("IAEV",$O) !== false ){
    if ( strpos("EV",$O) ){
       $w_disabled = ' DISABLED ';
    }
    AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
    ShowHTML ('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
    ShowHTML ('<INPUT type="hidden" name="CL" value="' . $CL . '">');
    ShowHTML ('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
    ShowHTML ('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
    ShowHTML ('<INPUT type="hidden" name="O" value="' . $O . '">');

    ShowHTML ('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
    ShowHTML ('    <table width="95%" border="0">');
    ShowHTML ('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML ('        <tr valign="top">');
    ShowHTML ('          <td valign="top"><font size="1"><b><u>D</u>ata:</b><br><input accesskey="D" type="text" name="w_dt_ocorrencia" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_ocorrencia,Time()),2)) . '" onKeyDown="FormataData(this,event);"></td>');
    ShowHTML ('          <td valign="top"><font size="1"><b>D<u>e</u>scrição:</b><br><input ' . $w_disabled . ' accesskey="E" type="text" name="w_ds_ocorrencia" class="STI" SIZE="60" MAXLENGTH="60" VALUE="' . $w_ds_ocorrencia . '"></td>');
    $SQL = 'SELECT * FROM sbpi.Tipo_Data a WHERE a.abrangencia <> \'U\' ORDER BY a.nome' . $crlf;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);	 
    ShowHTML ('          <td><font size="1"><b>Tipo da ocorrência:</b><br><SELECT CLASS="STI" NAME="w_tipo">');
    ShowHTML ('          <option value=""> ---');
    foreach($RS as $row) {
       if ( doubleval(nvl(f($row, "sq_tipo_data"),0)) == doubleval(nvl($w_tipo,0)) ){
          ShowHTML ('          <option value=' . f($row, "sq_tipo_data") . ' SELECTED>' . f($row, "nome"));
       } else {
          ShowHTML ('          <option value=' . f($row, "sq_tipo_data") . '>' . f($row, "nome"));
       }
    }
    ShowHTML ('          </select>');
    ShowHTML ('        </table>');
    ShowHTML ('      <tr>');
    ShowHTML ('      <tr><td align="center" colspan=4><hr>');
    if ( $O == "E" ){
       ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir">');
    } else {
       if ( $O == "I" ){
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Incluir">');
       } else {
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
       }
    }
    ShowHTML ('            <input class="STB" type="button" onClick="location.href=\'' .$dir.$w_pagina.$par. $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
    ShowHTML ('          </td>');
    ShowHTML ('      </tr>');
    ShowHTML ('    </table>');
    ShowHTML ('    </TD>');
    ShowHTML ('</tr>');
    ShowHTML ('</FORM>');
  } else {
    ScriptOpen ('JavaScript');
    ShowHTML (' alert(\'Opção não disponível\');');
    ShowHTML (' history.back(1);');
    ScriptClose();
  }
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();

}
// =========================================================================
// Fim do cadastro de calendário base
// -------------------------------------------------------------------------

// =========================================================================
// Cadastro de calendario
// -------------------------------------------------------------------------
function calend_rede(){
  extract($GLOBALS);
  $w_chave            = $_REQUEST["w_chave"];
  $w_troca            = $_REQUEST["w_troca"];
  
  if( $w_troca > "" ){ //Se for recarga da página
     $w_dt_ocorrencia = $_REQUEST["w_dt_ocorrencia"];
     $w_ds_ocorrencia = $_REQUEST["w_ds_ocorrencia"];
     $w_tipo          = $_REQUEST["w_tipo"];
  }else if( $O == "L" ){
     //Recupera todos os registros para a listagem
     $SQL = "select a.sq_ocorrencia as chave, a.ds_ocorrencia, a.dt_ocorrencia, a.sq_tipo_data, b.nome from sbpi.Calendario_Cliente a left join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) where sq_cliente = " . $CL . " order by sbpi.year(dt_ocorrencia) desc, dt_ocorrencia";
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  }else if( strpos("AEV",$O) !== false and $w_troca == '' ){
     //Recupera os dados do endereço informado
     $SQL = "select dt_ocorrencia, ds_ocorrencia, sq_tipo_data from sbpi.Calendario_Cliente where sq_ocorrencia = " . $w_chave;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
     foreach($RS as $row) { $RS = $row; break;}
     $w_dt_ocorrencia = FormataDataEdicao(f($row, "dt_ocorrencia"));
     $w_ds_ocorrencia = f($row, "ds_ocorrencia");
     $w_tipo          = f($row, "sq_tipo_data");
  }
  
  Cabecalho();
  ShowHTML ('<HEAD>');
  if( strpos("IAEP",O) !== false ){
     ScriptOpen("JavaScript");
     CheckBranco();
     FormataData();
     ValidateOpen("Validacao");
     ShowHTML ('  theForm.Botao[0].disabled=true;');
     ShowHTML ('  theForm.Botao[1].disabled=true;');
     ValidateClose();
     ScriptClose();
  }
  ShowHTML ('</HEAD>');
  if( $w_troca > "" ){
     BodyOpen ('onLoad=\'document.Form.' . $w_troca . '.focus()\';');
  }else if( $O == "I" or $O == "A" ){
     BodyOpen ('onLoad=\'document.Form.w_dt_ocorrencia.focus()\';');
  } else {
     BodyOpen ('onLoad=\'document.focus()\';');
  }
  ShowHTML ('<B><FONT COLOR="#000000">Cadastro do calendário da rede de ensino</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
  if( $O == "L" ){
    //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML ('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $w_pagina.'calend_rede' . $w_ew . '&R=' . $w_pagina.'calend_rede' . '&O=I&CL=' . $CL . '"><u>I</u>ncluir</a>&nbsp;');
    ShowHTML ('    <td align="right"><font size="1"><b>Registros existentes: ' . count($RS));
    ShowHTML ('<tr><td align="center" colspan=3>');
    ShowHTML ('    <TABLE WIDTH="100%" bgcolor="'.$conTableBgColor.'" BORDER="'.$conTableBorder.'" CELLSPACING="'.$conTableCellSpacing.'" CELLPADDING="'.$conTableCellPadding.'" BorderColorDark="'.$conTableBorderColorDark.'" BorderColorLight="'.$conTableBorderColorLight.'">');
    ShowHTML ('        <tr bgcolor="'.$conTrBgColor.'" align="center">');
    ShowHTML ('          <td><font size="1"><b>Data</font></td>');
    ShowHTML ('          <td><font size="1"><b>Tipo</font></td>');
    ShowHTML ('          <td><font size="1"><b>Ocorrência</font></td>');
    ShowHTML ('          <td><font size="1"><b>Operações</font></td>');
    ShowHTML ('        </tr>');

    if (count($RS)<=0) {
      // Se não foram selecionados registros, exibe mensagem
      ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
    } else {
      foreach ($RS as $row) {
        $w_cor = ($w_cor==$conTrBgColor || $w_cor=='') ? $w_cor=$conTrAlternateBgColor : $w_cor=$conTrBgColor;
        if( $wAno != year(f($row, "dt_ocorrencia")) ){
           ShowHTML ('      <tr bgcolor="#C0C0C0" valign="top"><TD colspan=4 align="center"><font size=2><B>' . year(f($row, "dt_ocorrencia")) . '</b></font></td></tr>');
           $wAno = year(f($row, "dt_ocorrencia"));
        }
        ShowHTML ('      <tr bgcolor="'.$w_cor.'" valign="top">');
        ShowHTML ('        <td align="center"><font size="1">' . substr(FormataDataEdicao(FormatDateTime(f($row, "dt_ocorrencia"),2)),0,5) . '</td>');
        ShowHTML ('        <td><font size="1">' . nvl(f($row, "nome"),"---") . '</td>');
        ShowHTML ('        <td><font size="1">' . f($row, "ds_ocorrencia") . '</td>');
        ShowHTML ('        <td align="top" nowrap>');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=A&w_chave='.f($row,'chave').'">Alterar</A>&nbsp');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=E&w_chave='.f($row,'chave').'" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
        ShowHTML ('        </td>');
        ShowHTML ('      </tr>'); 
      }
    }
    ShowHTML ('      </center>');
    ShowHTML ('    </table>');
    ShowHTML ('  </td>');
    ShowHTML ('</tr>');
  } elseif (!(strpos('IAEV',$O)===false)) {
    if (!(strpos('EV',$O)===false)) $w_disabled=' DISABLED ';
    AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);    
    ShowHTML ('<INPUT type="hidden" name="w_chave" value="'.$w_chave.'">');
    ShowHTML ('<INPUT type="hidden" name="w_troca" value="">');
    ShowHTML ('<tr bgcolor="'.$conTrBgColor.'"><td align="center">');
    ShowHTML ('    <table width="97%" border="0">');
    ShowHTML ('      <tr><td><table border=0 width="100%"><tr valign="top">');
    ShowHTML ('           <td colspan=2><b><u>D</u>ata:</b><br><input '.$w_disabled.' accesskey="D" type="text" name="w_dt_ocorrencia" class="sti" SIZE="10" MAXLENGTH="10" VALUE="'.FormataDataEdicao(FormatDateTime(Nvl($w_dt_ocorrencia,Time()),2)).'"></td>');
    ShowHTML ('      <tr><td valign="top" colspan="2"><table border="0" width="100%" cellspacing=0>');
    ShowHTML ('        <tr valign="top">');
    ShowHTML ('           <td colspan=2><b><u>D</u>escrição:</b><br><input '.$w_disabled.' accesskey="E" type="text" name="w_ds_ocorrencia" class="sti" SIZE="60" MAXLENGTH="60" VALUE="'.$w_ds_ocorrencia.'"></td>');
    $SQL = 'SELECT * FROM sbpi.Tipo_Data a WHERE a.abrangencia <> \'U\' ORDER BY a.nome' . $crlf;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    ShowHTML ('          <td><font size="1"><b>Tipo da ocorrência:</b><br><SELECT CLASS="STI" NAME="w_tipo">');
    ShowHTML ('          <option value=""> ---');
    foreach($RS as $row) {
       if( doubleval(nvl(f($row, "sq_tipo_data"),0)) == doubleval(nvl($w_tipo,0)) ){
          ShowHTML ('          <option value="' . f($row, "sq_tipo_data") . '" SELECTED>' . f($row, "nome"));
       } else {
          ShowHTML ('          <option value="' . f($row, "sq_tipo_data") . '">' . f($row, "nome"));
       }
    }
    ShowHTML ('          </select>');
    ShowHTML ('        </table>');
    ShowHTML ('      <tr>');
    ShowHTML ('      <tr><td align="center" colspan=4><hr>');
    if( $O == "E" ){
       ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir">');
    } else {
       if( $O == "I" ){
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Incluir">');
       } else {
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
       }
    }

    ShowHTML ('            <input class="STB" type="button" onClick="location.href=\'' .$dir.$w_pagina.$par. $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
    ShowHTML ('          </td>');
    ShowHTML ('      </tr>');
    ShowHTML ('    </table>');
    ShowHTML ('    </TD>');
    ShowHTML ('</tr>');
    ShowHTML ('</FORM>');
  } else {
    ScriptOpen("JavaScript");
    ShowHTML (' alert(\'Opção não disponível\');');
    ShowHTML (' history.back(1);');
    ScriptClose();
  }
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();
}
// =========================================================================
// Fim do cadastro de calendário
// -------------------------------------------------------------------------

// =========================================================================
// Monta a tela de Homologação do Calendário das Escolas Particulares
// -------------------------------------------------------------------------
function ligacao(){
  $p_regiao = $_REQUEST["p_regiao"];
  $w_homologado = $_REQUEST["w_homologado"];
  if($w_homologado != "S") {
    $w_homologado = 'Não';
  } else {
    $w_homologado = 'Sim';
  }

  if ($p_tipo == 'W' ){
      //Response.ContentType = "application/msword"
      //HeaderWord p_layout
      ShowHTML ('<TABLE WIDTH="100%" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR="#000000">SIGE-WEB<TD ALIGN="RIGHT"><B><FONT SIZE=4 COLOR="#000000">');
      ShowHTML ('Consulta a escolas');
      ShowHTML ('</FONT><TR><TD ALIGN="RIGHT"><B><FONT SIZE=2 COLOR="#000000">" . DataHora(); . "</B></TD></TR>');
      ShowHTML ('</FONT></B></TD></TR></TABLE>');
      ShowHTML ('<HR>');
  } else {
     Cabecalho();
     ShowHTML ('<HEAD>');
     ShowHTML ('   <link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
     ShowHTML ('</HEAD>');
     if ($_REQUEST["pesquisa"] > '' ){
        BodyOpen ('onLoad="location.href=\'#lista\'"');
     //} else {
        //BodyOpen "onLoad='document.Form.p_regional.focus()';"
     }
  }
  ShowHTML ('<B><FONT COLOR="#000000">'.$w_tp.'</FONT></B>');
  ShowHTML ('<B><FONT size="2" COLOR="#000000">Vinculação e tipologia</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border="0" cellpadding="0" cellspacing="0" width="100%">');
  ShowHTML ('<tr bgcolor="" & conTrBgColor & ""><td align="center">'); 
  ShowHTML ('<tr bgcolor="'.$conTrBgColor.'"><td align="center">');
  ShowHTML ('    <table width="90%" cellspacing=0>');  
  AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
  ShowHTML ('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
  ShowHTML ('<INPUT type="hidden" name="w_ea" value="" & w_ea & "">');
  ShowHTML ('<INPUT type="hidden" name="w_troca" value="">');
  ShowHTML ('<INPUT type="hidden" name="CL" value="' . $CL . '">');
  ShowHTML ('      <tr><td colspan=2><table border=0 width="90%" cellspacing=0>');
  ShowHTML ('        <tr valign="top">');
  SelecaoEscolaParticular         ('Unidad<u>e</u> de ensino:', 'E', 'Selecione unidade.' , $p_escola_particular, null, "p_escola_particular", null, "onChange='document.Form.action=\'" . $w_pagina . $w_ew . "'; document.Form.w_ea.value='P'; document.Form.w_troca.value='p_calendario'; document.Form.submit();");
  //SelecaoEscolaParticular         ("Unidad<u>e</u> de ensino:", "E", "Selecione unidade." , $p_escola_particular, null, "p_escola_particular", null, "onChange="document.Form.action='" & w_pagina & w_ew & "'; document.Form.w_ea.value='P'; document.Form.w_troca.value='p_calendario'; document.Form.submit();"");
  //ShowHTML ('        <tr valign="top">');
  //SelecaoCalendarioParticular         "<u>C</u>alendário:", "E", "Selecione unidade." , p_calendario, p_escola_particular, "p_calendario", null, "onChange="document.Form.action='" & w_pagina & w_ew & "'; document.Form.w_ea.value='P'; document.Form.w_troca.value='w_homologado'; document.Form.submit();""
  //ShowHTML ('        <tr valign="top">"
  //SelecaoRegionalEscola "Regional de en<u>s</u>ino:", "S", "Indique a regional de ensino.", p_regional, p_escola_particular, "p_regional", null, null
  //ShowHTML ('        <tr valign="top">"
  //SelecaoTipoEscola     "<u>T</u>ipo de Escola:", "T", "Selecione o tipo da Escola.", p_tipo_escola, p_escola_particular, "p_tipo_escola", null, null
  if($p_escola_particular > ''){
  $SQL = 'SELECT sq_particular_calendario, sq_cliente as cliente, nome, homologado FROM escParticular_Calendario WHERE sq_cliente = ' . $p_escola_particular;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  
  ShowHTML ('<font color="#ff3737"><strong><a href="javascript:this.status.value;" onClick="window.open(\'calendario.asp?CL=sq_cliente=" & RS("cliente") & "&w_ea=L&w_ew=formcal&w_ee=1','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no\');">Acessar o(s) calendário(s)</a></strong></font>');
    
  ShowHTML ('<table id="tbhomologacao" border="1">');
  ShowHTML ('<tr><td>Título do Calendário</td><td>Homologado?</td></tr>');
  
  foreach ($RS as $row) {
     $homologado = RS("homologado");
     ShowHTML ('<INPUT type="hidden" name="w_cliente" value="' . RS("cliente") . '">');
     ShowHTML ('<tr><td>' .  $RS("nome") . '</td>');     
     ShowHTML ('<INPUT type="hidden" name="w_chave" value="' . RS("sq_particular_calendario") .'">');
     ShowHTML ('<td><select name="w_homologado">');          
     if(strpos(strtoupper($homologado),'N')){
        ShowHTML ('<option value="N" SELECTED>Não');
        ShowHTML ('<option value="S">Sim');
     }else if((strpos(strtoupper($homologado),"S")) ){
        ShowHTML ('<option value="S" SELECTED>Sim');
        ShowHTML ('<option value="N">Não');
     }     
     ShowHTML ('</select></td>');
     ShowHTML ('<td align="center" colspan="2">');
     //ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Gravar">');

     ShowHTML ('          </td>');
     ShowHTML ('</tr>'); 
  }  
  ShowHTML ('</table>');
    
  }
  ShowHTML ('         <tr valign="top">');

  ShowHTML ('          </table>');
  ShowHTML ('      <tr valign="top">');
  ShowHTML ('      <tr>');
  ShowHTML ('      <tr><td align="center" colspan="2" height="1" bgcolor="#000000">');
  ShowHTML ('      <tr><td align="center" colspan="2">');
  ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Gravar">');
  ShowHTML ('          </td>');
  ShowHTML ('      </tr>');
  ShowHTML ('    </table>');
  ShowHTML ('    </TD>');
  ShowHTML ('</tr>');
  ShowHTML ('</FORM>');
  ShowHTML ('</table>');
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();
}
//Fim da Pesquisa de Escolas Particulares

function newsletter(){
  extract($GLOBALS);
  
  $w_chave           = $_REQUEST["w_chave"];
  $w_troca           = $_REQUEST["w_troca"];
  
  if( $w_troca > '' ){ //Se for recarga da página
    $w_nome       = $_REQUEST["w_nome"];
    $w_email      = $_REQUEST["w_email"];
    $w_tipo       = $_REQUEST["w_tipo"];
    $w_envia_mail = $_REQUEST["w_envia_mail"];
  }else if( $O == 'L' || $O == 'G' ){
     if( $_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI" ){
        $SQL = "select sq_newsletter as chave, nome, email, tipo, envia_mail, data_inclusao, data_alteracao, " . $crlf . 
               "       case tipo when '1' then 'Responsável' " . $crlf . 
               "                 when '2' then 'Aluno' " . $crlf . 
               "                 when '3' then 'Outro' " . $crlf . 
               "       end nm_tipo, " . $crlf . 
               "       case envia_mail when 'S' then 'Sim' else 'Não' end nm_envia " . $crlf . 
               "  from sbpi.Newsletter " . $crlf . 
               " where sq_cliente = 0 " . $crlf;
     } else {
        //Recupera todos os registros para a listagem
        $SQL = "select sq_newsletter as chave, nome, email, tipo, envia_mail, data_inclusao, data_alteracao, " . $crlf . 
               "       case tipo when '1' then 'Responsável' " . $crlf . 
               "                 when '2' then 'Aluno' " . $crlf . 
               "                 when '3' then 'Outro' " . $crlf . 
               "       end nm_tipo, " . $crlf . 
               "       case envia_mail when 'S' then 'Sim' else 'Não' end nm_envia " . $crlf . 
               "  from sbpi.Newsletter " . $crlf . 
               " where sq_cliente= " . $CL . " " . $crlf;
     }
     if( $O == "G" ){
        $SQL .= '   and envia_mail = \'S\' ' . $crlf;
     }
     $SQL .= 'order by nome';
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  }else if( strpos("AEV",$O) !== false and $w_troca == '' ){
     //Recupera os dados do endereço informado
     $SQL = 'select * from sbpi.Newsletter where sq_newsletter = ' . $w_chave;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
     foreach($RS as $row) { $RS = $row; break;}
     $w_nome       = f($RS, "nome");
     $w_email      = f($RS, "email");
     $w_tipo       = f($RS, "tipo");
     $w_envia_mail = f($RS, "envia_mail");
  }
  
  Cabecalho();
  ShowHTML ('<HEAD>');
  if( strpos("IAEP",$O) !== false ){
     ScriptOpen ("JavaScript");
     CheckBranco();
     FormataData();
     ValidateOpen ("Validacao");
     if( strpos("IA",$O) !== false ){
        Validate ("w_nome"  , "Nome"   , "" , "1" , "3" , "60" , "1" , "1");
        Validate ("w_email" , "e-Mail" , "" , "1" , "4" , "60" , "1" , "1");
        ShowHTML ('  if (theForm.w_tipo[0].checked==false && theForm.w_tipo[1].checked==false && theForm.w_tipo[2].checked==false) {');
        ShowHTML ('     alert(\'Você deve selecionar uma das opções apresentadas no formulário!\');');
        ShowHTML ('     return false;');
        ShowHTML ('  }');
     }
     ShowHTML ('  theForm.Botao[0].disabled=true;');
     ShowHTML ('  theForm.Botao[1].disabled=true;');
     ValidateClose();
     ScriptClose();
  }
  ShowHTML ('</HEAD>');
  if( $w_troca > '' ){
     BodyOpen ('onLoad=\'document.Form.' . $w_troca . '.focus()\';');
  }else if( $O == 'I' || $O == 'A' ){
     BodyOpen ('onLoad=\'document.Form.w_nome.focus()\';');
  } else {
     BodyOpen ('onLoad=\'document.focus()\';');
  }
  ShowHTML ('<B><FONT COLOR="#000000">Lista de distribuição de informativos</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
  if( $O == "L" ){
    //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML ('<tr><td><font size="2">');
    ShowHTML ('<tr><td><a accesskey="I" class="SS" href="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=I&w_chave='.$w_chave.'"><u>I</u>ncluir</a>&nbsp;');
    ShowHTML ('        <a accesskey="L" class="SS" href="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=G&&w_chave='.$w_chave.'"><u>L</u>istar e-mails </a>&nbsp;');
    ShowHTML ('    <td align="right"><font size="1"><b>Registros existentes: ' .count($RS));
    ShowHTML ('<tr><td align="center" colspan=3>');
    ShowHTML ('    <TABLE WIDTH="100%" bgcolor="'.$conTableBgColor.'" BORDER="'.$conTableBorder.'" CELLSPACING="'.$conTableCellSpacing.'" CELLPADDING="'.$conTableCellPadding.'" BorderColorDark="'.$conTableBorderColorDark.'" BorderColorLight="'.$conTableBorderColorLight.'">');
    ShowHTML ('        <tr bgcolor="'.$conTrBgColor.'" align="center">');
    ShowHTML ('          <td><font size="1"><b>Nome</font></td>');
    ShowHTML ('          <td><font size="1"><b>Tipo</font></td>');
    ShowHTML ('          <td><font size="1"><b>Envia</font></td>');
    ShowHTML ('          <td><font size="1"><b>Cadastro</font></td>');
    ShowHTML ('          <td><font size="1"><b>Alteração</font></td>');
    ShowHTML ('          <td><font size="1"><b>Operações</font></td>');
    ShowHTML ('        </tr>');

    if (count($RS)<=0) {
      // Se não foram selecionados registros, exibe mensagem
      ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td colspan=4 align="center"><b>Não foram encontrados registros.</b></td></tr>');
    } else {
      foreach ($RS as $row) {
        $w_cor = ($w_cor==$conTrBgColor || $w_cor=='') ? $w_cor=$conTrAlternateBgColor : $w_cor=$conTrBgColor;
        ShowHTML ('      <tr bgcolor="'.$w_cor.'" valign="top">');
        ShowHTML ('        <td><font size="1"><a class="HL" href="mailto:' . f($row, "email") . '" title="' . f($row, "email") . '">' . f($row, "nome") . '</a></td>');
        ShowHTML ('        <td align="center"><font size="1">' . f($row, "nm_Tipo") . '</td>');
        ShowHTML ('        <td align="center"><font size="1">' . f($row, "nm_envia") . '</td>');
        ShowHTML ('        <td align="center"><font size="1">' . FormataDataEdicao(FormatDateTime(f($row, "data_inclusao"),2)) . '</td>');
        ShowHTML ('        <td align="center"><font size="1">');
        if( f($row, "data_alteracao") > "" ){
          ShowHTML (FormataDataEdicao(FormatDateTime(f($row, "data_alteracao"),2)));
        } else { 
          ShowHTML ('---');
        }
        ShowHTML ('</td>');
        ShowHTML ('        <td align="top" nowrap><font size="1">');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=A&w_chave='.f($row,'chave').'">Alterar</A>&nbsp');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=E&w_chave='.f($row,'chave').'" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
        ShowHTML ('        </td>');
        ShowHTML ('      </tr>');
      }
    }
    ShowHTML ('      </center>');
    ShowHTML ('    </table>');
    ShowHTML ('  </td>');
    ShowHTML ('</tr>');
  }else if( $O == 'G' ){
    //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML ('<tr><td><font size="2">');
    ShowHTML ('       <a accesskey="V" class="SS" href="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=L&w_chave='.f($row,'chave').'"><u>V</u>oltar</a>&nbsp;');
    ShowHTML ('    <td align="right"><font size="1"><b>Registros existentes: ' .count($RS));
    ShowHTML ('<tr><td align="center" colspan=3>');
    ShowHTML ('    <TABLE WIDTH="100%" bgcolor="'.$conTableBgColor.'" BORDER="'.$conTableBorder.'" CELLSPACING="'.$conTableCellSpacing.'" CELLPADDING="'.$conTableCellPadding.'" BorderColorDark="'.$conTableBorderColorDark.'" BorderColorLight="'.$conTableBorderColorLight.'">');
    ShowHTML ('        <tr bgcolor="'.$conTrBgColor.'" align="center">');
    ShowHTML ('          <td><font size="1"><b>Lista (cada linha com 20 e-mails)</font></td>');
    ShowHTML ('        </tr>');

    if (count($RS)<=0) {
      // Se não foram selecionados registros, exibe mensagem
      ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
    } else {
      $i = 0;
      $j = 0;
      foreach($RS as $row) {
        if( $i == 0 or $j >= 20 ){
           $i = 1;
           ShowHTML ('      <tr valign="top"><td><font size="1">');
           $j = 0;
        }
        ShowHTML ('        "' . f($row, "email") . '"; ');
        $j++;
      }
    }
    ShowHTML ('      </center>');
    ShowHTML ('    </table>');
    ShowHTML ('  </td>');
    ShowHTML ('</tr>');
  }else if( strpos("IAEV",$O) !== false ){
    if( strpos("EV",$O) ){
       $w_disabled = " DISABLED ";
    }
    AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
    ShowHTML ('<INPUT type="hidden" name="R" value="'.$w_ew.'">');
    ShowHTML ('<INPUT type="hidden" name="CL" value="'.$CL.'">');
    ShowHTML ('<INPUT type="hidden" name="w_chave" value="'.$w_chave.'">');
    ShowHTML ('<INPUT type="hidden" name="w_sq_cliente" value="' . str_replace($CL,"sq_cliente=","") . '">');
    ShowHTML ('<INPUT type="hidden" name="O" value="'.$O.'">');

    ShowHTML ('        <tr bgcolor="'.$conTrBgColor.'" align="center">');
    ShowHTML ('    <table width=""95%"" border=""0"">');
    ShowHTML ('      <TR><TD><font size="2" CLASS="BTM"><b>Nome completo:</b><br><input type="text" size="60" maxlength="60" name="w_nome" value="' . $w_nome . '" CLASS="STI">');
    ShowHTML ('      <TR><TD><font size="2" CLASS="BTM"><b>e-Mail:</b><br><input type="text" size="60" maxlength="60" name="w_email" value="' . $w_email . '" CLASS="STI">');
    ShowHTML ('      <TR><TD><font size="2" CLASS="BTM"><b>Tipo da pessoa:</b> ');
    if( $w_tipo == 1 ){
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="1" checked> Pai, mãe ou responsável por aluno da rede de ensino ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="2"> Aluno da rede de ensino ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="3"> Outro ');
    }else if( $w_tipo == 2 ){
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="1"> Pai, mãe ou responsável por aluno da rede de ensino ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="2" checked> Aluno da rede de ensino ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="3"> Outro ');
    }else if( $w_tipo == 3 ){
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="1"> Pai, mãe ou responsável por aluno da rede de ensino ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="2"> Aluno da rede de ensino ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="3" checked> Outro ');
    } else {
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="1"> Pai, mãe ou responsável por aluno da rede de ensino ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="2"> Aluno da rede de ensino ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="3"> Outro ');
    }
    ShowHTML ('      <TR> ');
    MontaRadioSN ("<b>Envia newsletter para este e-mail?</b>", $w_envia_mail, "w_envia_mail");
    ShowHTML ('      </TR> ');
    ShowHTML ('      <tr><td align=""center"" colspan=4><hr>');
    if( $O == 'E' ){
       ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir">');
    } else {
       if( $O == "I" ){
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Incluir">');
       } else {
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
       }
    }
    ShowHTML ('            <input class="STB" type="button" onClick="location.href=\'' .$dir.$w_pagina.$par. $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
    ShowHTML ('          </td>');
    ShowHTML ('      </tr>');
    ShowHTML ('    </table>');
    ShowHTML ('    </TD>');
    ShowHTML ('</tr>');
    ShowHTML ('</FORM>');
  } else {
    ScriptOpen ("JavaScript");
    ShowHTML (' alert(\'Opção não disponível\');');
    //ShowHTML (' history.back(1);');
    ScriptClose();
  }
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();
}
// =========================================================================
// Fim do cadastro de newsletter
// -------------------------------------------------------------------------

// =========================================================================
// Cadastro de notícias
// -------------------------------------------------------------------------
function Noticias(){
  extract($GLOBALS);

  $w_chave         = $_REQUEST["w_chave"];
  $w_troca         = $_REQUEST["w_troca"];
  
  if($w_troca > '' ){ // Se for recarga da página
     $w_dt_noticia = $_REQUEST["w_dt_noticia"];
     $w_ds_titulo  = $_REQUEST["w_ds_titulo"];
     $w_ds_noticia = $_REQUEST["w_ds_noticia"];
     $w_ln_noticia = $_REQUEST["w_ln_noticia"];
     $w_in_ativo   = $_REQUEST["w_in_ativo"];
     $w_in_exibe   = $_REQUEST["w_in_exibe"];
  }else if( $O == 'L' ){
     //Recupera todos os registros para a listagem
    if( $_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI" ){
        $SQL = "select sq_noticia as chave, ds_titulo, ds_noticia, dt_noticia, ativo, in_exibe from sbpi.Noticia_Cliente where sq_cliente = 0 order by dt_noticia desc";
    } else {
      $SQL = 'select sq_noticia as chave, ds_titulo, ds_noticia, dt_noticia, ativo, in_exibe from sbpi.Noticia_Cliente where sq_cliente= ' .$CL. ' order by dt_noticia desc';
    }
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  }else if( strpos("AEV",$O) !== false && $w_troca == '' ){
    //Recupera os dados do endereço informado
    $SQL = "select sq_noticia as chave, ds_titulo, ds_noticia, dt_noticia, ativo, in_exibe from sbpi.Noticia_Cliente where sq_noticia = " . $w_chave;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    foreach($RS as $row) { $RS = $row; break;}
    $w_dt_noticia = FormataDataEdicao(f($RS, "dt_noticia"));
    $w_ds_titulo  = f($RS, "ds_titulo");
    $w_ds_noticia = f($RS, "ds_noticia");
    $w_in_ativo   = f($RS, "in_ativo");
    $w_in_exibe   = f($RS, "in_exibe");
  }
  Cabecalho();
  ShowHTML ('<HEAD>');
  if( strpos("IAEP",$O) !== false ){
    ScriptOpen ("JavaScript");
    CheckBranco();
    FormataData();
    ValidateOpen ("Validacao");
    if( strpos("IA",$O) !== false ){
      Validate("w_dt_noticia" , "Data"      , "DATA" , "1" , "10" , "10"   , "1" , "1");
      Validate("w_ds_titulo"  , "Título"    , ""     , "1" , "2"  , "50"   , "1" , "1");
      Validate("w_ds_noticia" , "Descrição" , ""     , "1" , "2"  , "4000" , "1" , "1");
    }
    ShowHTML ('  theForm.Botao[0].disabled=true;');
    ShowHTML ('  theForm.Botao[1].disabled=true;');
    ValidateClose();
    ScriptClose();
  }
  
  ShowHTML ('</HEAD>');
  if ( $w_troca > "" ) {
     BodyOpen ('onLoad="document.Form.' . $w_troca . '.focus()";');
  }else if( $O == "I" or $O == "A" ) {
     BodyOpen ("onLoad='document.Form.w_ds_titulo.focus()';");
  } else {
     BodyOpen ("onLoad='document.focus()';");
  }
  ShowHTML ('<B><FONT COLOR="#000000">Cadastro de notícias da rede de ensino</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
  If ($O == 'L'){
    // Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML ('<tr><td><font size="2"><a accesskey="I" class="SS" href="' .$dir.$w_pagina.$par. $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
    ShowHTML ('    <td align="right"><b>Registros existentes: '.count($RS));
    ShowHTML ('<tr><td align="center" colspan=3>');
    ShowHTML ('    <TABLE WIDTH="100%" bgcolor="'.$conTableBgColor.'" BORDER="'.$conTableBorder.'" CELLSPACING="'.$conTableCellSpacing.'" CELLPADDING="'.$conTableCellPadding.'" BorderColorDark="'.$conTableBorderColorDark.'" BorderColorLight="'.$conTableBorderColorLight.'">');
    ShowHTML ('        <tr bgcolor="'.$conTrBgColor.'" align="center">');
    ShowHTML ('          <td><font size="1"><b>Data</font></td>');
    ShowHTML ('          <td><font size="1"><b>Título</font></td>');
    ShowHTML ('          <td><font size="1"><b>Ativo</font></td>');
    ShowHTML ('          <td><font size="1"><b>Operações</font></td>');
    ShowHTML ('        </tr>');

    if (count($RS)<=0) {
      // Se não foram selecionados registros, exibe mensagem
      ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
    } else {
      foreach($RS as $row){
        $w_cor = ($w_cor==$conTrBgColor || $w_cor=='') ? $w_cor=$conTrAlternateBgColor : $w_cor=$conTrBgColor;
        ShowHTML ('      <tr bgcolor="' . $w_cor . '" valign="top">');
        ShowHTML ('        <td align="center"><font size="1">' . FormataDataEdicao(FormatDateTime(f($row, "dt_noticia"),2)) . '</td>');
        ShowHTML ('        <td><font size="1">' . f($row, "ds_titulo") . '</td>');
        ShowHTML ('        <td align="center"><font size="1">' . f($row, "ativo") . '</td>');
        ShowHTML ('        <td align="top" nowrap><font size="1">');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=A&w_chave='.f($row,'chave').'">Alterar</A>&nbsp');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=E&w_chave='.f($row,'chave').'" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
        ShowHTML ('        </td>');
        ShowHTML ('      </tr>');
      }
    }
    ShowHTML ('      </center>');
    ShowHTML ('    </table>');
    ShowHTML ('  </td>');
    ShowHTML ('</tr>');
  }else if( strpos("IAEV",$O) !== false ) {
      if ( strpos("EV",$O) ) {
         $w_disabled = ' DISABLED ';
      }
      AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
      ShowHTML ('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
      ShowHTML ('<INPUT type="hidden" name="CL" value="' . $CL . '">');
      ShowHTML ('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML ('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
      ShowHTML ('<INPUT type="hidden" name="O" value="' . $O . '">');

      ShowHTML ('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
      ShowHTML ('    <table width="95%" border="0">');
      ShowHTML ('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
      ShowHTML ('        <tr valign="top">');
      ShowHTML ('          <td valign="top"><font size="1"><b><u>D</u>ata:</b><br><input accesskey="D" type="text" name="w_dt_noticia" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_noticia,Time()),2)) . '" onKeyDown="FormataData(this,event);" ></td>"');
      ShowHTML ('          <td valign="top"><font size="1"><b><u>T</u>ítulo:</b><br><input ' . $w_disabled . ' accesskey="T" type="text" name="w_ds_titulo" class="STI" SIZE="50" MAXLENGTH="50" VALUE="' . $w_ds_titulo . '"></td>');
      ShowHTML ('        </table>');
      ShowHTML ('      <tr><td valign="top"><font size="1"><b>D<u>e</u>scrição:</b><br><textarea " & w_Disabled & " accesskey="E" name="w_ds_noticia" class="STI" ROWS=5 cols=65>' . $w_ds_noticia . '</TEXTAREA></td>');
      ShowHTML ('      <tr>');
      ShowHTML ('      </tr>');
      ShowHTML ('      <tr>');
      MontaRadioSN ("<b>Exibir no site?</b>", $w_in_ativo, "w_in_ativo");
      ShowHTML ('      <tr>');
      ShowHTML ('      <tr><td align="center" colspan=4><hr>');
      if( $O == 'E' ){
        ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir">');
      } else {
        if( $O == "I" ){
        ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Incluir">');
        } else {
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }
  ShowHTML ('            <input class="STB" type="button" onClick="location.href=\'' .$dir.$w_pagina.$par. $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
  ShowHTML ('          </td>');
  ShowHTML ('      </tr>');
  ShowHTML ('    </table>');
  ShowHTML ('    </TD>');
  ShowHTML ('</tr>');
  ShowHTML ('</FORM>');
  } else {
    ScriptOpen ("JavaScript");
    ShowHTML (' alert(\'Opção não disponível\');');
    //ShowHTML (' history.back(1);');
    ScriptClose();
  }
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();
}
// =========================================================================
// Fim do cadastro de notícias
// -------------------------------------------------------------------------

// =========================================================================
// Cadastro de mensagens ao aluno
// -------------------------------------------------------------------------
function msgalunos(){
  extract($GLOBALS);

  $w_chave         = $_REQUEST["w_chave"];
  $w_troca         = $_REQUEST["w_troca"];
  $w_texto         = $_REQUEST["w_texto"];
  
  if($w_troca > '' ){ // Se for recarga da página
     $w_dt_mensagem  = $_REQUEST["w_dt_mensagem");
     $w_ds_mensagem  = $_REQUEST["w_ds_mensagem");
     $w_sq_aluno     = $_REQUEST["w_sq_aluno");
  }else if( $O == 'L' ){
     //Recupera todos os registros para a listagem
     $SQL = "select a.sq_mensagem as chave, b.no_aluno, b.nr_matricula, c.ds_cliente " . $crlf . 
            "  from sbpi.Mensagem_Aluno     a " . $crlf . 
            "       inner join sbpi.Aluno   b on (a.sq_aluno        = b.sq_aluno) " . $crlf . 
            "       inner join sbpi.Cliente c on (b.sq_cliente = c.sq_cliente) " . $crlf . 
            "order by c.ds_cliente, a.dt_mensagem desc, b.no_aluno, a.in_lida " . $crlf;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  }else if( strpos("AEV",$O) !== false && $w_troca == '' ){
     //Recupera os dados do endereço informado
     $SQL = 'select * from sbpi.Mensagem_Aluno where sq_mensagem = ' . $w_chave;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
     foreach($RS as $row) { $RS = $row; break;}
     $w_dt_mensagem  = FormataDataEdicao(f($row, "dt_mensagem"));
     $w_ds_mensagem  = f($row, "ds_mensagem");
     $w_sq_aluno     = f($row, "sq_aluno");
  }
  
  Cabecalho();
  ShowHTML ('<HEAD>');
  if( strpos("IAEP",$O) !== false ){
    ScriptOpen ("JavaScript");
    CheckBranco();
    FormataData();
    ValidateOpen ("Validacao");
    if( strpos("IA",$O) !== false ){
        Validate ("w_dt_mensagem" , "Data"     , "DATA"  , "1" , "10" , "10"   , "1" , "1");
        Validate ("w_ds_mensagem" , "Mensagem" , ""      , "1" , "2"  , "4000" , "1" , "1");
        if( $O == "I"){
          Validate ("w_sq_aluno" , "Aluno"    , "SELECT", "1" , "1"  , "10"   , ""  , "1");
        }
     }
    ShowHTML ('  theForm.Botao[0].disabled=true;');
    ShowHTML ('  theForm.Botao[1].disabled=true;');
    ValidateClose();
    ScriptClose();
  }
  ShowHTML ('</HEAD>');
  if ( $w_troca > "" ) {
     BodyOpen ('onLoad="document.Form.' . $w_troca . '.focus()";');
  }else if( $O == "I" or $O == "A" ) {
     BodyOpen ("onLoad='document.Form.w_ds_titulo.focus()';");
  } else {
     BodyOpen ("onLoad='document.focus()';");
  }
  ShowHTML ('<B><FONT COLOR="#000000">Mensagens a alunos da rede de ensino</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
  If ($O == 'L'){
    //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML ('<tr><td><font size="2"><a accesskey="I" class="SS" href="' .$dir.$w_pagina.$par. $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
    ShowHTML ('    <td align="right"><b>Registros existentes: '.count($RS));
    ShowHTML ('<tr><td align="center" colspan=3>');
    ShowHTML ('    <TABLE WIDTH="100%" bgcolor="'.$conTableBgColor.'" BORDER="'.$conTableBorder.'" CELLSPACING="'.$conTableCellSpacing.'" CELLPADDING="'.$conTableCellPadding.'" BorderColorDark="'.$conTableBorderColorDark.'" BorderColorLight="'.$conTableBorderColorLight.'">');
    ShowHTML ('        <tr bgcolor="'.$conTrBgColor.'" align="center">');
    ShowHTML ('          <td><font size="1"><b>Unidade</font></td>');
    ShowHTML ('          <td><font size="1"><b>Data</font></td>');
    ShowHTML ('          <td><font size="1"><b>Matrícula</font></td>');
    ShowHTML ('          <td><font size="1"><b>Aluno</font></td>');
    ShowHTML ('          <td><font size="1"><b>Lida</font></td>');
    ShowHTML ('          <td><font size="1"><b>Operações</font></td>');
    ShowHTML ('        </tr>');

    if (count($RS)<=0) {
      // Se não foram selecionados registros, exibe mensagem
      ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
    } else {
      foreach($RS as $row){
        $w_cor = ($w_cor==$conTrBgColor || $w_cor=='') ? $w_cor=$conTrAlternateBgColor : $w_cor=$conTrBgColor;
        ShowHTML ('      <tr bgcolor="' . $w_cor . '" valign="top">');
        ShowHTML ('        <td><font size="1">' . strtolower(f($row, "ds_cliente")) . '</td>');
        ShowHTML ('        <td align="center"><font size="1">' . FormataDataEdicao(FormatDateTime(f($row, "dt_mensagem"),2)) . '</td>');
        ShowHTML ('        <td align="center" nowrap><font size="1">' . f($row, "nr_matricula") . '</td>');
        ShowHTML ('        <td><font size="1">' . strtolower(f($row, "no_aluno")) . '</td>');
        ShowHTML ('        <td align="center"><font size="1">' . f($row, "in_lida") . '</td>');
        ShowHTML ('        <td align="top" nowrap><font size="1">');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=A&w_chave='.f($row,'chave').'">Alterar</A>&nbsp');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=E&w_chave='.f($row,'chave').'" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
        ShowHTML ('        </td>');
        ShowHTML ('      </tr>');
      }
    }
  ShowHTML ('      </center>');
  ShowHTML ('    </table>');
  ShowHTML ('  </td>');
  ShowHTML ('</tr>');
  }else if( strpos("IAEV",$O) !== false ) {
  if ( strpos("EV",$O) ) {
     $w_disabled = ' DISABLED ';
  }
  AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
  ShowHTML ('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
  ShowHTML ('<INPUT type="hidden" name="CL" value="' . $CL . '">');
  ShowHTML ('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
  ShowHTML ('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
  ShowHTML ('<INPUT type="hidden" name="O" value="' . $O . '">');

  ShowHTML ('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
  ShowHTML ('    <table width="95%" border="0">');
  ShowHTML ('      <tr><td><font size="1"><b><u>D</u>ata:</b><br><input accesskey="D" type="text" name="w_dt_mensagem" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_mensagem,Time()),2)) . '" onKeyDown="FormataData(this,event);"></td>');
  ShowHTML ('      <tr><td><font size="1"><b>M<u>e</u>nsagem:</b><br><textarea '.$w_disabled.' accesskey="E" name="w_ds_mensagem" class="STI" ROWS=5 cols=65>'.$w_ds_mensagem.'</TEXTAREA>');
  if( $O == "I" ){
       ShowHTML ('      <tr><td><font size="1"><b><u>P</u>rocurar por:</b><br><input accesskey="P" type="text" name="w_texto" class="STI" SIZE="50" MAXLENGTH="50" VALUE="' . $w_texto . '" >');
       ShowHTML ('          <input type="Button" class="STB" name="Pesquisa" value="Procurar" onClick="document.Form.w_troca.value=\''.$w_sq_aluno.'\'; document.Form.action=\'"' . $dir.$w_pagina.$par . $w_ew . '"\'; document.Form.submit();"></td>');
       ShowHTML ('      <tr><td><font size="1"><b><u>A</u>luno:</b><br><select accesskey="A" name="w_sq_aluno" class="STS" SIZE="1" >');
       ShowHTML ('          <OPTION VALUE="">---');
       If ($w_texto > ''){
          $SQL = "select sq_aluno, nr_matricula, no_aluno " . $crlf . 
                 "  from sbpi.Aluno " . $crlf . 
                 " where (upper(no_aluno) like \'%" . strtoupper($w_texto) . "%\' or " . $crlf . 
                 "        nr_matricula like \'%" . $w_texto . "%\') " . $crlf . 
                 "order by no_aluno, nr_matricula" . $crlf;
          $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
          foreach($RS as $row){
             ShowHTML ('          <OPTION VALUE="' . RS("sq_aluno") . '">' . RS("no_aluno") . '" ("' . RS("nr_matricula") . '")"');
          }
       }
       ShowHTML ('          </select>');
    } else {
       $SQL = 'select nr_matricula, no_aluno from sbpi.Aluno where sq_aluno = ' . $w_sq_aluno;
       $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
       foreach($RS as $row) { $RS = $row; break;}
       ShowHTML ('      <tr><td><font size="1"><b>Aluno:<br>' . f($RS, "no_aluno") . ' (' . f($RS, "nr_matricula") . ')</td>');
    }
    ShowHTML ('      <tr>');
    ShowHTML ('      <tr><td align=""center"" colspan=4><hr>');
    
    if( $O == 'E' ){
       ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir">');
    } else {
       if( $O == "I" ){
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Incluir">');
       } else {
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
       }
    }
    ShowHTML ('            <input class="STB" type="button" onClick="location.href=\'' .$dir.$w_pagina.$par. $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
    ShowHTML ('          </td>');
    ShowHTML ('      </tr>');
    ShowHTML ('    </table>');
    ShowHTML ('    </TD>');
    ShowHTML ('</tr>');
    ShowHTML ('</FORM>');
  } else {
    ScriptOpen ("JavaScript");
    ShowHTML (' alert(\'Opção não disponível\');');
    //ShowHTML (' history.back(1);');
    ScriptClose();
  }
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();

}
// =========================================================================
// Fim do cadastro de mensagens ao aluno
// -------------------------------------------------------------------------

// =========================================================================
// Cadastro de modalidades de ensino
// -------------------------------------------------------------------------
function tipoCliente(){
  extract($GLOBALS);

  $w_chave         = $_REQUEST["w_chave"];
  $w_troca         = $_REQUEST["w_troca"];
  
  if($w_troca > '' ){ // Se for recarga da página
     $w_tipo         = $_REQUEST["w_tipo"];
     $w_nome         = $_REQUEST["w_nome"];
  }else if( $O == 'L' ){
  //Recupera todos os registros para a listagem
     $SQL = "select sq_tipo_cliente as chave, ds_tipo_cliente, ds_registro, tipo, " . $crlf . 
           "       case tipo when '1' then 'Secretaria' " . $crlf . 
           "                 when '2' then 'Regional'   " . $crlf . 
           "                 when '3' then 'Escola'     " . $crlf . 
           "       end nm_tipo                        " . $crlf . 
           "  from sbpi.Tipo_Cliente                    " . $crlf;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  }else if( strpos("AEV",$O) !== false && $w_troca == '' ){
  //Recupera os dados do endereço informado
         $SQL = "select sq_tipo_cliente as chave, ds_tipo_cliente, ds_registro, tipo, " . $crlf . 
           "       case tipo when '1' then 'Secretaria' " . $crlf . 
           "                 when '2' then 'Regional'   " . $crlf . 
           "                 when '3' then 'Escola'     " . $crlf . 
           "       end nm_tipo                        " . $crlf . 
           "  from sbpi.Tipo_Cliente                    " . $crlf .
           "  where sq_tipo_Cliente = " . $w_chave;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    foreach($RS as $row) { $RS = $row; break;}
    $w_tipo         = f($RS, "tipo");
    $w_nome         = f($RS, "ds_tipo_cliente");
  }
 
  Cabecalho();
  ShowHTML ('<HEAD>');
  if( strpos("IAEP",$O) !== false ){
    ScriptOpen ("JavaScript");
    CheckBranco();
    FormataData();
    ValidateOpen ("Validacao");
    if( strpos("IA",$O) !== false ){
        Validate ("w_dt_mensagem" , "Data"     , "DATA"  , "1" , "10" , "10"   , "1" , "1");
        Validate ("w_ds_mensagem" , "Mensagem" , ""      , "1" , "2"  , "4000" , "1" , "1");
        if( $O == "I"){
          Validate ("w_sq_aluno" , "Aluno"    , "SELECT", "1" , "1"  , "10"   , ""  , "1");
        }
     }
    ShowHTML ('  theForm.Botao[0].disabled=true;');
    ShowHTML ('  theForm.Botao[1].disabled=true;');
    ValidateClose();
    ScriptClose();
  }
  ShowHTML ('</HEAD>');
  if ( $w_troca > "" ) {
     BodyOpen ('onLoad="document.Form.' . $w_troca . '.focus()";');
  }else if( $O == "I" or $O == "A" ) {
     BodyOpen ("onLoad='document.Form.w_ds_titulo.focus()';");
  } else {
     BodyOpen ("onLoad='document.focus()';");
  }
  ShowHTML ('<B><FONT COLOR="#000000">Cadastro de tipos de Instituições </FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
  If ($O == 'L'){
    //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML ('<tr><td><font size="2"><a accesskey="I" class="SS" href="' .$dir.$w_pagina.$par. $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
    ShowHTML ('    <td align="right"><b>Registros existentes: '.count($RS));
    ShowHTML ('<tr><td align="center" colspan=3>');
    ShowHTML ('    <TABLE WIDTH="100%" bgcolor="'.$conTableBgColor.'" BORDER="'.$conTableBorder.'" CELLSPACING="'.$conTableCellSpacing.'" CELLPADDING="'.$conTableCellPadding.'" BorderColorDark="'.$conTableBorderColorDark.'" BorderColorLight="'.$conTableBorderColorLight.'">');
    ShowHTML ('        <tr bgcolor="'.$conTrBgColor.'" align="center">');
    ShowHTML ('          <td><font size="1"><b>Descrição</font></td>');
    ShowHTML ('          <td><font size="1"><b>Tipo</font></td>');
    ShowHTML ('          <td><font size="1"><b>Operações</font></td>');
    ShowHTML ('        </tr>');
    if (count($RS)<=0) {
      // Se não foram selecionados registros, exibe mensagem
      ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
    } else {
      foreach($RS as $row){
        $w_cor = ($w_cor==$conTrBgColor || $w_cor=='') ? $w_cor=$conTrAlternateBgColor : $w_cor=$conTrBgColor;
        ShowHTML ('      <tr bgcolor="' . $w_cor . '" valign="top">');
        ShowHTML ('        <td><font size="1">' . f($row, "ds_tipo_cliente") . '</td>');
        switch(f($row, "tipo")){
           case 1:
              ShowHTML ('          <td align="center"><font size="1">Secretaria</font></td>'); break;
           case 2:
              ShowHTML ('          <td align="center"><font size="1">Regional</font></td>');break;
           case 3:
              ShowHTML ('          <td align="center"><font size="1">Escola</font></td>');break;
           default:
              ShowHTML ('        <td><font size="1">---</td>');break;
        }
        ShowHTML ('        <td align="top" nowrap><font size="1">');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=A&w_chave='.f($row,'chave').'">Alterar</A>&nbsp');
        ShowHTML ('          <A class="HL" HREF="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=E&w_chave='.f($row,'chave').'" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
        ShowHTML ('        </td>');
        ShowHTML ('      </tr>');        
      }
    }
  ShowHTML ('      </center>');
  ShowHTML ('    </table>');
  ShowHTML ('  </td>');
  ShowHTML ('</tr>');
  }else if( strpos("IAEV",$O) !== false ) {
  if ( strpos("EV",$O) ) {
     $w_disabled = ' DISABLED ';
  }
  AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
  ShowHTML ('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
  ShowHTML ('<INPUT type="hidden" name="CL" value="' . $CL . '">');
  ShowHTML ('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
  ShowHTML ('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
  ShowHTML ('<INPUT type="hidden" name="O" value="' . $O . '">');

  ShowHTML ('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
  ShowHTML ('    <table width="95%" border="0">');
  ShowHTML ('      <tr><td valign=""top"" colspan=""3""><table border=0 width=""100%"" cellspacing=0>');
  ShowHTML ('        <tr valign="top"><td valign="top"><font size="1"><b>D<u>e</u>scrição:</b><br><input ' . $w_disabled . ' accesskey="E" type="text" name="w_nome" class="STI" SIZE="60" MAXLENGTH="60" VALUE="' . $w_nome . '" ></td></tr>');
  ShowHTML ('      <TD><font size="2" CLASS="BTM"><b>Tipo:</b> ');
    if ($w_tipo == 1 ){
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="1" checked> Secretaria ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="2"> Regional ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="3"> Escola ');
    }else if( $w_tipo == 2 ){
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="1"> Secretaria ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="2" checked> Regional ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="3"> Escola ');
    }else if( $w_tipo == 3 ){
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="1"> Secretaria ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="2"> Regional ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="3" checked> Escola ');
    } else {
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="1"> Secretaria de Ensino do DF ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="2"> Rede de Ensino ');
       ShowHTML ('            <br><input type="Radio" name="w_tipo" value="3" checked> Escola ');
    }
    ShowHTML ('        </table>');
    ShowHTML ('      <tr>');
    ShowHTML ('      <tr><td align=""center"" colspan=4><hr>');
    
    if( $O == 'E' ){
       ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir">');
    } else {
       if( $O == "I" ){
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Incluir">');
       } else {
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
       }
    }
    ShowHTML ('            <input class="STB" type="button" onClick="location.href=\'' .$dir.$w_pagina.$par. $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
    ShowHTML ('          </td>');
    ShowHTML ('      </tr>');
    ShowHTML ('    </table>');
    ShowHTML ('    </TD>');
    ShowHTML ('</tr>');
    ShowHTML ('</FORM>');
  } else {
    ScriptOpen ("JavaScript");
    ShowHTML (' alert(\'Opção não disponível\');');
    //ShowHTML (' history.back(1);');
    ScriptClose();
  }
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();

}
// =========================================================================
// Fim do cadastro de modalidades de ensino
// -------------------------------------------------------------------------

// =========================================================================
// Tela de alteração da Senha
// -------------------------------------------------------------------------
function senha(){
extract($GLOBALS);
  if($O != 'A'){
    $O = "A";
  }

  If ($O == "A"){
    $SQL = "select * from sbpi.Cliente where sq_cliente = " . $CL;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    foreach($RS as $row) { $RS = $row; break;}
    $w_sq_cliente      = f($row, 'sq_cliente');
    $w_ds_senha_acesso = f($row, 'ds_senha_acesso');
  }
  
  Cabecalho();
  ShowHTML ('<HEAD>');
  ScriptOpen ("JavaScript");
  CheckBranco();
  FormataData();
  ValidateOpen ("Validacao");
  If ($O == "A"){
     Validate ("w_ds_senha_acesso", "Senha de acesso", "1", "1", "4", "14", "1", "1");
  }
  ValidateClose();
  ScriptClose();
  ShowHTML ('</HEAD>');
  BodyOpen ('onLoad=\'document.Form.w_ds_senha_acesso.focus();\'');
  ShowHTML ('<B><FONT COLOR="#000000">Atualização da senha de acesso</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
  AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
  ShowHTML(MontaFiltro('POST'));
  ShowHTML ('<INPUT type="hidden" name="R" value="' . $w_ew . '">');

  ShowHTML ('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
  ShowHTML ('    <table width="95%" border="0">');
  ShowHTML ('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Senha de acesso</td></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td><font size=1>Esta tela permite alterar a senha de acesso à tela de atualização dos dados da rede de ensino. Assim que a nova senha for gravada, ela já deverá ser utilizada para acessar as telas desta aplicação.</font></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top"><font size="1"><b><u>S</u>enha de acesso:</b><br><INPUT ACCESSKEY="S" ' . $w_disabled . ' class="STI" type="text" name="w_ds_senha_acesso" size="14" maxlength="14" value="' . $w_ds_senha_acesso . '"></td>');
  ShowHTML ('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');

  //Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
  ShowHTML ('      <tr><td align="center" colspan="3">');
  ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Gravar">');
  ShowHTML ('          </td>');
  ShowHTML ('      </tr>');
  ShowHTML ('    </table>');
  ShowHTML ('    </TD>');
  ShowHTML ('</tr>');
  ShowHTML ('</FORM>');
  ShowHTML ('</table>');
  Rodape();
}
// =========================================================================
// Fim da tela de dados básicos
// -------------------------------------------------------------------------

// =========================================================================
// Monta a tela de senhas especiais
// -------------------------------------------------------------------------
function senhaesp(){
  extract($GLOBALS);
  Cabecalho();
  BodyOpen ('onLoad=\'document.focus()\';');
  ShowHTML ('<B><FONT COLOR=""#000000"">Senhas especiais - Listagem</FONT></B>');
  ShowHTML ('<div align=center><center>');
  $SQL = "SELECT a.ds_username, a.sq_cliente, a.ds_cliente,  " . $crlf . 
        "       a.ds_senha_acesso, a.no_municipio, a.sg_uf, a.dt_alteracao " . $crlf .  
        "  from sbpi.Cliente a " . $crlf .  
        " where a.publica = 'S' and a.sq_cliente_pai is null or a.sq_cliente_pai = 0 and a.ds_username <> 'SBPI'" . $crlf . 
        "ORDER BY a.sq_cliente_pai, a.ds_cliente " . $crlf;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  

  ShowHTML ('<TR><TD valign="top"><table border=0 width="100%" cellpadding=0 cellspacing=0>');
  if (count($RS) > 0) {

     //RS.PageSize = RS.RecordCount + 1
     //rs.AbsolutePage = 1    

     ShowHTML ('<tr><td><td align="right"><b><font face=Verdana size=1>Registros encontrados: ' . count($RS) . '</font></b>');
     ShowHTML ('<tr><td><td>');
     ShowHTML ('<table border="1" cellspacing="0" cellpadding="0" width="100%">');
     ShowHTML ('<tr align="center" valign="top">');
     ShowHTML ('    <td><font face="Verdana" size="1"><b>Escola</b></td>');
     ShowHTML ('    <td><font face="Verdana" size="1"><b>Username</b></td>');
     ShowHTML ('    <td><font face="Verdana" size="1"><b>Senha</b></td>');
  
    // While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(Nvl($_REQUEST["P3"),RS.AbsolutePage))
    foreach($RS as $row){
       $w_cor = ($w_cor==$conTrBgColor || $w_cor=='') ? $w_cor=$conTrAlternateBgColor : $w_cor=$conTrBgColor;
       ShowHTML ('      <tr bgcolor="' . $w_cor . '" valign="top">');
       ShowHTML ('    <td><font face="Verdana" size="1">' . f($row, "ds_cliente") . '</font></td>');
       ShowHTML ('    <td><font face="Verdana" size="1">' . f($row, "ds_username") . '</font></td>');
       ShowHTML ('    <td align="center"><font face="Verdana" size="1">' . f($row, "ds_senha_acesso") . '</font></td>');
    }
    
     ShowHTML ('</table>');
     ShowHTML ('<tr><td><td colspan=""4"" align=""center""><hr>');
  } else {
    ShowHTML ('<TR><TD><TD colspan="3"><p align="justify"><img src="img/ico_educacao.gif" width="16" height="16" border="0" align="center">&nbsp;<font size="2"><b>Nenhuma ocorrência encontrada para as opções acima.');
  }

  ShowHTML ('</TABLE>');
  
}
// -------------------------------------------------------------------------
// Final da Página de Pesquisa
// =========================================================================

// =========================================================================
// Monta a tela de Homologação do Calendário das Escolas Particulares
// -------------------------------------------------------------------------
function escPartHomolog(){
extract($GLOBALS);
  //exibevariaveis();
  
  $$p_escola_particular = $_REQUEST["p_escola_particular"];
  $p_calendario         = $_REQUEST["p_calendario"];
  $p_regiao             = $_REQUEST["p_regiao"];
  $w_homologado         = $_REQUEST["w_homologado"];
  if($w_homologado != 'S') {
    $w_homologado = 'Não';
  } else {
    $w_homologado = 'Sim';
  }
  if ($p_tipo == "W"){
      //Response.ContentType = "application/msword"
      //HeaderWord p_layout
      ShowHTML ('<TABLE WIDTH="100%" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR="#000000">SIGE-WEB<TD ALIGN="RIGHT"><B><FONT SIZE=4 COLOR="#000000">');
      ShowHTML ('Consulta a escolas');
      ShowHTML ('</FONT><TR><TD ALIGN="RIGHT"><B><FONT SIZE=2 COLOR="#000000">' . DataHora() . '</B></TD></TR>');
      ShowHTML ('</FONT></B></TD></TR></TABLE>');
      ShowHTML ('<HR>');
  } else {
     Cabecalho();
     ShowHTML ('<HEAD>');
     ShowHTML ('   <link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
     ShowHTML ('</HEAD>');
     if ($_REQUEST["pesquisa"] > '' ){
        BodyOpen (' onLoad="location.href=\'#lista\'');
     }
  }
  ShowHTML ('<B><FONT COLOR="#000000">' . $w_tp . '</FONT></B>');
  ShowHTML ('<B><FONT COLOR="#000000">Homologação de Calendário</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border="0" cellpadding="0" cellspacing="0" width="100%">');
  ShowHTML ('<tr bgcolor="'.$conTrBgColor.'"><td align="center">'); 
  ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td align="center" valign="top"><table border=0 width="90%" cellspacing=0>');
  AbreForm('Form', $w_dir.$w_pagina.'Grava', 'POST', 'return(Validacao(this));', null);
  ShowHTML ('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
  ShowHTML ('<INPUT type="hidden" name="CL" value="' . $CL . '">');
  ShowHTML ('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
  ShowHTML ('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
  ShowHTML ('<INPUT type="hidden" name="O" value="' . $O . '">');
  ShowHTML ('      <tr><td colspan=2><table border=0 width="90%" cellspacing=0>');
  ShowHTML ('        <tr valign="top">');
  SelecaoEscolaParticular ('Unidad<u>e</u> de ensino:', 'E', 'Selecione unidade.' , $p_escola_particular, null, "p_escola_particular", null, "onChange='document.Form.action=\'" .$dir.$w_pagina.$par.$w_ew . "'; document.Form.w_ea.value='P'; document.Form.w_troca.value='p_calendario'; document.Form.submit();");
  
  if($p_escola_particular > '') {
  $SQL = "SELECT sq_particular_calendario, sq_cliente as cliente, nome, homologado FROM escParticular_Calendario WHERE sq_cliente = " . $p_escola_particular;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  foreach($RS as $row) { $RS = $row; break;}  
  ShowHTML ('<font color="#ff3737"><strong><a href="javascript:this.status.value;" onClick="window.open(\'calendario.asp?CL=sq_cliente=' . RS("cliente") . '&w_ea=L&w_ew=formcal&w_ee=1','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no\');">Acessar o(s) calendário(s)</a></strong></font>');    
  ShowHTML ('<table id="tbhomologacao" border="1">');
  ShowHTML ('<tr><td>Título do Calendário</td><td>Homologado?</td></tr>');
  
  foreach($RS as $row){
     $homologado = RS("homologado");
     ShowHTML ('<INPUT type="hidden" name="w_cliente"    value="" & RS("cliente") & "">');
     ShowHTML ('<tr><td>" &  RS("nome") & "</td>');     
     ShowHTML ('<INPUT type="hidden" name="w_chave" value="" & RS("sq_particular_calendario") &"">');
     ShowHTML ('<td><select name="w_homologado">');          
     if(strpos(strtoupper(homologado),"N")){
        ShowHTML ('<option value="N" SELECTED>Não');
        ShowHTML ('<option value="S">Sim');
     } else if(strpos(strtoupper(homologado),"S")){
        ShowHTML ('<option value="S" SELECTED>Sim');
        ShowHTML ('<option value="N">Não');
     }     
     ShowHTML ('</select></td>'); 
     ShowHTML ('<td align="center" colspan="2">');

     ShowHTML ('          </td>');
     ShowHTML ('</tr>'); 
  }  
  ShowHTML ('</table>');
  
  
  }
  ShowHTML ('         <tr valign="top">');

  ShowHTML ('          </table>');
  ShowHTML ('      <tr valign="top">');
  ShowHTML ('      <tr>');
  ShowHTML ('      <tr><td align="center" colspan="2" height="1" bgcolor="#000000">');
  ShowHTML ('      <tr><td align="center" colspan="2">');
  ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Gravar">');
  ShowHTML ('          </td>');
  ShowHTML ('      </tr>');
  ShowHTML ('    </table>');
  ShowHTML ('    </TD>');
  ShowHTML ('</tr>');
  ShowHTML ('</FORM>');
  ShowHTML ('</table>');
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();
}
// Fim da Pesquisa de Escolas Particulares

function Grava() {
exibevariaveis();
}

// =========================================================================
// Monta a tela de Pesquisa
// -------------------------------------------------------------------------
function escPart(){

 
  //p_regiao = $_REQUEST["p_regiao")

  if( $p_tipo == "W" ){/*
      Response.ContentType = "application/msword"
      HeaderWord p_layout
      ShowHTML ('<TABLE WIDTH="100%" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR="#000000">SIGE-WEB<TD ALIGN="RIGHT"><B><FONT SIZE=4 COLOR="#000000">"
      ShowHTML ('Consulta a escolas"
      ShowHTML ('</FONT><TR><TD ALIGN="RIGHT"><B><FONT SIZE=2 COLOR="#000000">" & DataHora() & "</B></TD></TR>"
      ShowHTML ('</FONT></B></TD></TR></TABLE>"
      ShowHTML ('<HR>"*/
  } else {
     Cabecalho();
     ShowHTML ('<HEAD>');
     
     ShowHTML ('</HEAD>');
     if( $_REQUEST["pesquisa"] > '' ){
        BodyOpen (' onLoad="location.href=\'#lista\'');
     } else {
        BodyOpen ('onLoad=\'document.focus()\';');
     }
  }
  ShowHTML ('<B><FONT COLOR="#000000">' . $w_TP . '</FONT></B>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<tr bgcolor="' . $conTrBgColor .'"><td align="center">');
  ShowHTML ('    <table width="95%" border="0">');
  if( $p_tipo == "H" ){
     Showhtml ('<FORM ACTION="controle.asp" name="Form" METHOD="POST">');
     ShowHTML ('<INPUT TYPE="HIDDEN" NAME="w_ew" VALUE="ESCPART">');
     ShowHTML ('<INPUT TYPE="HIDDEN" NAME="CL" VALUE="" & CL &  "">');
     ShowHTML ('<INPUT TYPE="HIDDEN" NAME="pesquisa" VALUE="X">');
     ShowHTML ('<input type="Hidden" name="P3" value="1">');
     ShowHTML ('<input type="Hidden" name="P4" value="15">');
     ShowHTML ('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
     ShowHTML ('    <table width="100%" border="0">');
     ShowHTML ('          <TR><TD valign="top"><table border=0 width="100%" cellpadding=0 cellspacing=0>');
  } else {
     ShowHTML ('<tr><td><div align="justify"><font size=1><b>Filtro:</b><ul>');
  }
  if( $p_tipo == "H" ){
     ShowHTML ('          <tr valign="top"><td>');
     SelecaoRegiaoAdm ("Região a<u>d</u>ministrativa:", "D", "Indique a região administrativa.", $p_regiao, "p_regiao", "p_regiao", null, null);
  } else if( nvl($p_regiao,0) > 0 ){
     $SQL = "SELECT  a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " . $crlf . 
           "  FROM  escCLIENTE a " . $crlf . 
           " WHERE  a.sq_cliente = " . p_regiao . $crlf . 
           "ORDER BY a.ds_cliente ";
     ConectaBD SQL
     ShowHTML('          <li><font size="1"><b>Escolas da ' . RS("ds_cliente") . '</b>');
  }
  if( $p_tipo == "H" ){
     ShowHTML ('  <TR><TD><TD><font size="1"><br>');
     if( $_REQUEST["C"] > '' ){
        ShowHTML ('          <input type="checkbox" name="C" value="X" CLASS="BTM" checked> Exibir apenas escolas com calendário(s) cadastrado(s) ');
     } else {
        ShowHTML ('          <input type="checkbox" name="C" value="X" CLASS="BTM"> Exibir apenas escolas com calendário(s) cadastrado(s)  ');
     }
  } else if( $_REQUEST["C"] > '' ){
     ShowHTML ('  <li><font size="1"><b>Apenas escolas com alunos carregados </b>');
  }
  
  ShowHTML ('          </tr>');
  ShowHTML ('          </table>');
  if ($p_tipo = "H" ){ 
     ShowHTML ('  <TR><TD colspan=2><font size="1"><b>Campos a serem exibidos');
     if( p_layout = "PORTRAIT" ){ 
        ShowHTML ('          (<input type="radio" name="p_layout" value="PORTRAIT" CLASS="BTM" checked> Retrato <input type="radio" name="p_layout" value="LANDSCAPE" CLASS="BTM"> Paisagem)');
     } else {
        ShowHTML ('          (<input type="radio" name="p_layout" value="PORTRAIT" CLASS="BTM"> Retrato <input type="radio" name="p_layout" value="LANDSCAPE" CLASS="BTM" checked> Paisagem)');
     }
     ShowHTML ('     <table width="100%" border=0>');
     ShowHTML ('       <tr valign="top">');
      if( $_SESSION["USERNAME"] = "SEDF" or $_SESSION["USERNAME"] = "SBPI" or $_SESSION["username") = "CTIS" or Substr($_SESSION["username"),0,2) = "RE" ){
        if( strpos(p_campos,"username") !== false ){ 
          ShowHTML (' <td><font size=1><input type="checkbox" name="p_campos" value="username" CLASS="BTM" checked>Username');           
        } else {
          ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="username" CLASS="BTM">Username');   
        }
      }
      if( strpos(p_campos,"senha")       !== false ){
        ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="senha" CLASS="BTM" checked>Senha');                 
      } else {
        ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="senha" CLASS="BTM">Senha');                
      }
  }
     if( strpos($p_campos,"alteracao")   !== false ){ 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="alteracao" CLASS="BTM" checked>Última alteração');  
     } else {
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="alteracao" CLASS="BTM">Última alteração'); 
     }
     if( strpos($p_campos,"diretor")     !== false ){ 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="diretor" CLASS="BTM" checked>Diretor');             
     } else { 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="diretor" CLASS="BTM">Diretor');            
     }
     if( strpos($p_campos,"secretario")  !== false ){
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="secretario" CLASS="BTM" checked>Secretário');
     } else { 
      ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="secretario" CLASS="BTM">Secretário');      
     }
     if( strpos($p_campos,"contato")     !== false ){
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="contato" CLASS="BTM" checked>Contato');
     } else { 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="contato" CLASS="BTM">Contato');
     }
     ShowHTML ('       <tr valign="top">');
     if( strpos($p_campos,"mail") !== false ){
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="mail" CLASS="BTM" checked>e-Mail');
     } else { 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="mail" CLASS="BTM">e-Mail');                
     }
     if( strpos($p_campos,"telefone")    !== false ){
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="telefone" CLASS="BTM" checked>Telefone');           
     } else { ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="telefone" CLASS="BTM">Telefone');          }
     if( strpos($p_campos,"endereco")    !== false ){ 
      ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="endereco" CLASS="BTM" checked>Endereço');           
     } else { 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="endereco" CLASS="BTM">Endereço');          
     }
     if( strpos($p_campos,"localizacao") !== false ){ 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="localizacao" CLASS="BTM" checked>Localização');     
     } else { 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="localizacao" CLASS="BTM">Localização');    
     }
     if( strpos($p_campos,"cep")         !== false ){ 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="cep" CLASS="BTM" checked>CEP');                     
     } else { 
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos" value="cep" CLASS="BTM">CEP');                    
     }
     ShowHTML ('     </table>');
     ShowHTML ('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000">');
     ShowHTML ('      <tr><td align="center" colspan="3">');
     ShowHTML ('            <input class="BTM" type="submit" name="Botao" value="Aplicar filtro">');
     if( $_SESSION["USERNAME"] == "SBPI" ){
        ShowHTML ('            <input class="BTM" type="button" name="Botao" onClick="location.href=\'"'. .$dir.$w_pagina.'cadastroescola'.'&CL='.$CL.MontaFiltro("GET").'&w_ea=I\'; value="Nova escola">');
     }
     ShowHTML ('          </td>');
     ShowHTML ('      </tr>');
  }
  ShowHTML ('    </table>');
  ShowHTML ('    </TD>');
  ShowHTML ('</tr>');
  if ($p_tipo = "H" ){ 
    ShowHTML ('</form>'); 
  }
  if( $_REQUEST["pesquisa") > '' ){
    $SQL = " SELECT DISTINCT a.sq_cliente, a.ds_cliente, a.ds_apelido, a.ln_internet, a.ds_username, a.ds_senha_acesso, a.no_municipio, a.sg_uf, a.dt_alteracao, d.diretor, d.secretario, d.telefone_1, d.fax, d.cep, d.endereco, d.email_1 " . $crlf .  
           "   from escCliente a " . $crlf .  
           "        INNER JOIN escTipo_Cliente b ON (a.sq_tipo_cliente = b.sq_tipo_cliente and b.tipo=4) "
          if ($_REQUEST["c"] > '') { 
              $SQL .= "      INNER JOIN escCalendario_cliente c ON (a.sq_cliente = c.sq_site_cliente) " . $crlf .  
              "        INNER JOIN escCliente_Particular d ON (a.sq_cliente = d.sq_cliente) " . $crlf;
          } else {
              $SQL .= $crlf .  
              "        INNER JOIN escCliente_Particular d ON (a.sq_cliente = d.sq_cliente) " . $crlf
          }

          if( ($p_regiao > ''){
              $SQL = $SQL . "and a.sq_regiao_adm = " . $p_regiao . $crlf .  
              "        order by a.ds_cliente ";
           } else {
              $SQL = $SQL . $crlf .  
              "        order by a.ds_cliente ";
           }           
        
     ConectaBD SQL
     /**************************************************************************/
     /**************************************************************************/
     /**************************************************************************/
     /**************************************************************************/
     /********                                                                                                                 **********/
     /********                                                                                                                 **********/
     /********                                                                                                                 **********/
     /********                                                                                                                 **********/
     /********                                        Continuar daqui!                                             **********/
     /********                                                                                                                 **********/
     /********                                                                                                                 **********/
     /********                                                                                                                 **********/
     /********                                                                                                                 **********/
     /********                                                                                                                 **********/
     /********                                                                                                                 **********/
     /**************************************************************************/
     /**************************************************************************/
     /**************************************************************************/
     /**************************************************************************/
     
     ShowHTML ('<TR><TD valign="top"><br><table border=0 width="100%" cellpadding=0 cellspacing=0>');
     if( Not RS.EOF ){

        if( p_tipo = "H" ){ 
           if( $_REQUEST["P4") > " ){ RS.PageSize = cDbl($_REQUEST["P4")) } else { RS.PageSize = 15 }
           rs.AbsolutePage = Nvl($_REQUEST["P3"),1)
        } else {
           RS.PageSize = RS.RecordCount + 1
           rs.AbsolutePage = 1
        }
      
/*
        ShowHTML ('<tr><td><td align="right"><b><font face=Verdana size=1>Registros encontrados: " & RS.RecordCount & "</font></b>"
        if( p_Tipo = "H" ){ ShowHTML ('     &nbsp;&nbsp;<A TITLE="Clique aqui para gerar arquivo Word com a listagem abaixo" class="SS" href="#"  onClick="window.open('Controle.asp?p_tipo=W&w_ew=" & w_ew & "&Q=" & $_REQUEST["Q") & "&C=" & $_REQUEST["C") & "&D=" & $_REQUEST["D") & "&U=" & $_REQUEST["U") & w_especialidade & MontaFiltro("GET") & "','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no');">Gerar Word<IMG ALIGN="CENTER" border=0 SRC="img/word.gif"></A>" }
        ShowHTML ('<tr><td><td>"
        ShowHTML ('<table border="1" cellspacing="0" cellpadding="0" width="100%">"
        ShowHTML ('<tr align="center" valign="top">"
        ShowHTML ('    <td><font face="Verdana" size="1"><b>Escola</b></td>"
        if( Session("username") = "SEDF" or Session("username") = "CTIS" or Mid(Session("username"),1,2) = "RE" or Session("username") = "SBPI" ){
           if( strpos(p_campos,"username") > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1"><b>Username</b></td>" }
           if( strpos(p_campos,"senha") > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1"><b>Senha</b></td>" }
        }
        'ShowHTML ('    <td><font face="Verdana" size="1"><b>Alunos</b></td>"
        if( strpos(p_campos,"alteracao")   > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1"><b>Última alteração</b></td>" }
        if( strpos(p_campos,"diretor")     > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1"><b>Diretor</b></td>" }
        if( strpos(p_campos,"secretario")  > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1"><b>Secretario</b></td>" }
        if( strpos(p_campos,"mail")        > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1"><b>e-Mail</b></td>" }
        if( strpos(p_campos,"telefone")    > 0 ){ 
           ShowHTML ('    <td><font face="Verdana" size="1"><b>Telefone</b></td>" 
           ShowHTML ('    <td><font face="Verdana" size="1"><b>Fax</b></td>" 
        }
        if( strpos(p_campos,"endereco")    > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1"><b>Endereço</b></td>" }
        if( strpos(p_campos,"localizacao") > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1"><b>Localização</b></td>" }
        if( strpos(p_campos,"cep")         > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1"><b>CEP</b></td>" }
        w_cor = "#FDFDFD"
        While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(Nvl($_REQUEST["P3"),RS.AbsolutePage))
          if( w_cor = "#EFEFEF" or w_cor = " ){ w_cor = "#FDFDFD" } else { w_cor = "#EFEFEF" }
          ShowHTML ('<tr valign="top" bgcolor="" & w_cor & "">"
          if( p_tipo = "H" ){
          
             if((RS("LN_INTERNET") > ") ){
                ShowHTML ('    <td><font face="Verdana" size="1"><a href="" & RS("LN_INTERNET") & "" target="_blank">" & RS("DS_CLIENTE") & "</a></b></font></td>"
             } else {
                ShowHTML ('    <td><font face="Verdana" size="1"><b>" & RS("DS_CLIENTE") & "</b></font></td>"
             End if                
          } else {
             ShowHTML ('    <td><font face="Verdana" size="1">" & RS("DS_CLIENTE") & "</font></td>"
          }
          if( Session("username") = "SEDF" or Session("username") = "CTIS" or Mid(Session("username"),1,2) = "RE" or Session("username") = "SBPI" ){
             if( strpos(p_campos,"username") > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1">" & RS("DS_USERNAME") & "</font></td>" }
             if( strpos(p_campos,"senha") > 0 ){ ShowHTML ('    <td align="center"><font face="Verdana" size="1">" & RS("DS_SENHA_ACESSO") & "</font></td>" }
          }
          'ShowHTML ('    <td align="right"><font face="Verdana" size="1">" & RS("alunos") & "</font></td>"
          if( strpos(p_campos,"alteracao")   > 0 ){ ShowHTML ('    <td align="center"><font face="Verdana" size="1">" & Nvl(RS("dt_alteracao"),"---") & "</font></td>" }
          if( strpos(p_campos,"diretor")     > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1">" & Nvl(RS("diretor"),"---") & "</td>" }
          if( strpos(p_campos,"secretario")  > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1">" & Nvl(RS("secretario"),"---") & "</td>" }
          if( strpos(p_campos,"mail")        > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1">" & Nvl(RS("email_1"),"---") & "</td>" }
          if( strpos(p_campos,"telefone")    > 0 ){ 
             ShowHTML ('    <td><font face="Verdana" size="1">" & Nvl(RS("telefone_1"),"---") & "</td>" 
             ShowHTML ('    <td><font face="Verdana" size="1">" & Nvl(RS("fax"),"---") & "</td>" 
          }
          if( strpos(p_campos,"endereco")    > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1">" & Nvl(RS("endereco"),"---") & "</td>" }
          if( strpos(p_campos,"localizacao") > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1">" & RS("no_municipio") & "-" & RS("sg_uf") & "</font></td>" }
          if( strpos(p_campos,"cep")         > 0 ){ ShowHTML ('    <td><font face="Verdana" size="1">" & Nvl(RS("cep"),"---") & "</td>" }
          if( p_tipo = "H" ){
             ShowHTML ('    <td><font face="Verdana" size="1">"
             if( Session("username") = "SBPI" ){
                ShowHTML ('       <A CLASS="SS" HREF="" & w_Pagina & "CadastroEscola" & "&w_chave=" & RS("sq_cliente") & "&w_ea=A" & MontaFiltro("GET") & "" Title="Alteração dos dados da escola!">Alt</A>"
             }             
'             if( nvl(RS("adm"),"nulo") <> "nulo" ){
'                ShowHTML ('       <A CLASS="SS" HREF="controle.asp?CL=sq_cliente=" & RS("sq_cliente") & "&w_ea=L&w_ew=Adm&w_ee=1&P3=1&P4=30" Title="Exibe o formulário de dados administrativos preenchido pela escola!" target="_blank">Adm</A>"
'             }
          }
          RS.MoveNext
          Wend
    
        ShowHTML ('</table>');
        ShowHTML ('<tr><td><td colspan="5" align="center"><hr>');

        if( $p_tipo = "H" ){
           ShowHTML ('<tr><td align="center" colspan=5>');
           MontaBarra "Controle.asp?CL=" & CL & "&w_ew=" & w_ew & "&Q=" & $_REQUEST["Q") & "&C=" & $_REQUEST["C") & "&D=" & $_REQUEST["D") & "&U=" & $_REQUEST["U") & w_especialidade, cDbl(RS.PageCount), cDbl($_REQUEST["P3")), cDbl($_REQUEST["P4")), cDbl(RS.RecordCount)
           ShowHTML ('</tr>"
        }

     } else {

        ShowHTML ('<TR><TD><TD colspan="3"><p align="justify"><img src="img/ico_educacao.gif" width="16" height="16" border="0" align="center">&nbsp;<font size="2"><b>Nenhuma ocorrência encontrada para as opções acima."

     }

  }
  ShowHTML ('</TABLE>"
 
*/
}
//Fim da Pesquisa de Escolas Particulares 

// =========================================================================
// Cadastro de arquivos
// -------------------------------------------------------------------------
function Arquivos(){
  extract($GLOBALS);
  $w_chave           = $_REQUEST["w_chave"];
  $w_troca           = $_REQUEST["w_troca"];
  
  if ( $w_troca > "" ) { // Se for recarga da página
     $w_dt_arquivo      = $_REQUEST["w_dt_arquivo"];    
     $w_ds_titulo       = $_REQUEST["w_ds_titulo"];
     $w_in_ativo        = $_REQUEST["w_in_ativo"];   
     $w_ds_arquivo      = $_REQUEST["w_ds_arquivo"];    
     $w_ln_arquivo      = $_REQUEST["w_ln_arquivo"];    
     $w_in_destinatario = $_REQUEST["w_in_destinatario"];    
     $w_nr_ordem        = $_REQUEST["w_nr_ordem"];    
  }else if( $O == "L" ) {
     //Recupera todos os registros para a listagem
     if ( $_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] = "SBPI" ) {
        $SQL = 'select * from sbpi.Cliente_Arquivo where sq_cliente = 0 order by nr_ordem, ltrim(upper(ds_titulo))';
     } else {
        $SQL = 'select * from sbpi.Cliente_Arquivo where " & replace($CL,"sq_cliente","sq_cliente") & " order by in_ativo, nr_ordem';
     }
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  }else if( strpos('AEV',$O)!==false and $w_troca == "" ) {
     //Recupera os dados do endereço informado
     $SQL = "select b.ds_diretorio, a.* from sbpi.Cliente_Arquivo a inner join sbpi.Cliente_Site b on (a.sq_cliente = b.sq_cliente) where a.sq_arquivo = " . $w_chave;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
     foreach($RS as $row) { $RS = $row; break; }
     $w_dt_arquivo      = FormataDataEdicao(f($RS, "dt_arquivo"));
     $w_ds_titulo       = f($RS, "ds_titulo");
     $w_in_ativo        = f($RS, "in_ativo");
     $w_ds_arquivo      = f($RS, "ds_arquivo");
     $w_ln_arquivo      = f($RS, "ln_arquivo");
     $w_in_destinatario = f($RS, "in_destinatario");
     $w_nr_ordem        = f($RS, "nr_ordem");
     $w_ds_diretorio    = f($RS, "ds_diretorio");
  }
  
  Cabecalho();
  ShowHTML ('<HEAD>');
  if ( strpos("IAEP",$O) !== false ) {
     ScriptOpen ('JavaScript');
     ValidateOpen ('Validacao');
     if ( strpos("IA",$O) !== false ) {
        Validate ("w_ds_titulo" , "Título"      , "", "1", "2", "50"  , "1", "1");
        Validate ("w_ds_arquivo", "Descrição"   , "", "1", "2", "200" , "1", "1");
        Validate ("w_ln_arquivo", "Link"        , "", "",  "2", "200" , "1", "1");
        Validate ("w_nr_ordem"  , "Nr. de ordem", "", "1", "1", "2"   , "1", "0123546789");
     }
     ShowHTML (' if (theForm.w_ln_arquivo.value > ""){');
     ShowHTML ('    if((theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.DLL\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.SH\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.BAT\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.EXE\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.ASP\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.PHP\')!=-1)) {');
     ShowHTML ('       alert(\'Tipo de arquivo não permitido!\');');
     ShowHTML ('       theForm.w_ln_arquivo.focus(); ');
     ShowHTML ('       return false;');
     ShowHTML ('    }');
     ShowHTML ('  }');           
     ShowHTML ('  theForm.Botao[0].disabled=true;');
     ShowHTML ('  theForm.Botao[1].disabled=true;');
     ValidateClose();
     ScriptClose();
  }
  ShowHTML ('</HEAD>');
  if ( $w_troca > "" ) {
     BodyOpen ('onLoad="document.Form.' . $w_troca . '.focus()";');
  }else if( $O == "I" or $O == "A" ) {
     BodyOpen ("onLoad='document.Form.w_ds_titulo.focus()';");
  } else {
     BodyOpen ("onLoad='document.focus()';");
  }
  ShowHTML ('<B><FONT COLOR="#000000">Cadastro de arquivos da rede de ensino</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<div align=center><center>');
  ShowHTML ('<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""95%"">');
  if ( $O == "L" ) {
    // Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
    ShowHTML ('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $w_pagina .'arquivos'. $w_ew . "&R=" . $w_pagina.'arquivos' . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
    ShowHTML ('    <td align="right"><font size="1"><b>Registros existentes: ' . count($RS));
    ShowHTML ('<tr><td align="center" colspan=3>');
    ShowHTML ('    <TABLE WIDTH="100%" bgcolor="" & conTableBgColor & "" BORDER="" & conTableBorder & "" CELLSPACING="" & conTableCellSpacing & "" CELLPADDING="" & conTableCellPadding & "" BorderColorDark="" & conTableBorderColorDark & "" BorderColorLight="" & conTableBorderColorLight & "">');
    ShowHTML ('        <tr bgcolor="" & "#EFEFEF" & "" align="center">');
    ShowHTML ('          <td><font size="1"><b>Ordem</font></td>');
    ShowHTML ('          <td><font size="1"><b>Arquivo</font></td>');
    ShowHTML ('          <td><font size="1"><b>Descrição</font></td>');
    ShowHTML ('          <td><font size="1"><b>Ativo</font></td>');
    ShowHTML ('          <td><font size="1"><b>Operações</font></td>');
    ShowHTML ('        </tr>');

    if (count($RS)<=0) {
      // Se não foram selecionados registros, exibe mensagem
      ShowHTML ('      <tr bgcolor="'.$conTrBgColor.'"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
    }else {
      foreach($RS as $row) {
        $w_cor = ($w_cor==$conTrBgColor || $w_cor=='') ? $w_cor=$conTrAlternateBgColor : $w_cor=$conTrBgColor;
        ShowHTML ('      <tr bgcolor="' . $w_cor . '" valign="top">');
        ShowHTML ('        <td align="center"><font size="1">' . f($row, "nr_ordem") . '</td>');
        ShowHTML ('        <td><font size="1">' . f($row, "ds_titulo") . '</td>');
        ShowHTML ('        <td><font size="1">' . f($row, "ds_arquivo") . '</td>');
        ShowHTML ('        <td align="center"><font size="1">' . f($row, "ativo") . '</td>');
        ShowHTML ('        <td align="top" nowrap><font size="1">');
        ShowHTML ('          <A class="HL" HREF="' . $w_pagina .'arquivos'. $w_ew . "&R=" . $w_pagina .'arquivos'. $w_ew . "&O=A&CL=" . $CL . "&w_chave=" . f($row, "sq_arquivo") . '">Alterar</A>&nbsp');
        ShowHTML ('          <A class="HL" HREF="' . $w_pagina . "GRAVA&R=" . $w_ew . "&O=E&CL=" . $CL . '"&w_sq_cliente="' . str_replace($CL,"sq_cliente=","") . '"&w_chave="' . f($row, "sq_arquivo") . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
        ShowHTML ('        </td>');
        ShowHTML ('      </tr>');
      }
    }
    ShowHTML ('      </center>');
    ShowHTML ('    </table>');
    ShowHTML ('  </td>');
    ShowHTML ('</tr>');
  }else if( strpos("IAEV",$O) !== false ) {
    if ( strpos("EV",$O) ) {
       $w_disabled = ' DISABLED ';
    }
    ShowHTML ('<FORM action="' . $w_pagina . '"Grava" method="POST" name="Form" onSubmit="return(Validacao(this));" enctype="multipart/form-data">');
    ShowHTML ('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
    ShowHTML ('<INPUT type="hidden" name="$CL" value="' . $CL . '">');
    ShowHTML ('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
    ShowHTML ('<INPUT type="hidden" name="w_sq_cliente" value="' . str_replace($CL,"sq_cliente=","") . '">');
    ShowHTML ('<INPUT type="hidden" name="O" value="' . $O . '">');

    ShowHTML ('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
    ShowHTML ('    <table width="95%" border="0">');
    ShowHTML ('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML ('        <tr valign="top">');
    ShowHTML ('          <td valign="top"><font size="1"><b><u>T</u>ítulo:</b><br><input "' . $w_disabled . '" accesskey="T" type="text" name="w_ds_titulo" class="STI" SIZE="50" MAXLENGTH="50" VALUE="' . $w_ds_titulo . '" ></td>');
    if ( O == "I" ) {
       ShowHTML ('          <td valign="top"><font size="1"><b>Cadastramento:</b><br><input disabled type="text" name="w_dt_arquivo" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(time(),2)) . '"></td>');
    } else {
       ShowHTML ('          <td valign="top"><font size="1"><b>Última alteração:</b><br><input disabled type="text" name="w_dt_arquivo" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime($w_dt_arquivo,2)) . '"></td>');
    }
    ShowHTML ('        </table>');
    ShowHTML ('      <tr><td valign="top"><font size="1"><b><u>D</u>escrição:</b><br><textarea " & w_disabled & " accesskey="D" name="w_ds_arquivo" class="STI" ROWS=5 cols=65>"' . $w_ds_arquivo . '"</TEXTAREA></td>');
    ShowHTML ('      <tr>');
    ShowHTML ('      </tr>');
    ShowHTML ('      <tr><td valign="top"><font size="1"><b><u>L</u>ink:</b><br><input " & w_disabled & " accesskey="L" type="file" name="w_ln_arquivo" class="STI" SIZE="80" MAXLENGTH="100" VALUE="" >');
    if ( $w_ln_arquivo > '' ) {
       ShowHTML ('              <b><a class="SS" href="' . $w_ds_diretorio . '/' . $w_ln_arquivo . '" target="_blank" title="Clique para exibir o arquivo atual.">Exibir</a></b>');
    }
    ShowHTML ('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML ('        <tr valign="top">');
    ShowHTML ('          <td><font size="1"><b><u>N</u>r. de ordem:</b><br><input "' . $w_disabled . '" accesskey="N" type="text" name="w_nr_ordem" class="STI" SIZE="2" MAXLENGTH="2" VALUE="' . $w_nr_ordem . '"></td>');
    ShowHTML ('          <td><font size="1"><b><u>D</u>estinatários:</b><br><select "' . $w_disabled . '" accesskey="D" name="w_in_destinatario" class="STS" SIZE="1">');
    if ( $w_in_destinatario == 'A' ) {
       ShowHTML ('            <OPTION VALUE="A" SELECTED>Apenas alunos <OPTION VALUE="P">Apenas professores <OPTION VALUE="T">Professores e alunos <OPTION VALUE="E">Escola');
    }else if( $w_in_destinatario == "E" ) {
       ShowHTML ('            <OPTION VALUE="A">Apenas alunos <OPTION VALUE="P">Apenas professores <OPTION VALUE="T">Professores e alunos <OPTION VALUE="E" SELECTED>Escola');
    }else if( $w_in_destinatario == "P" ) {
       ShowHTML ('            <OPTION VALUE="A">Apenas alunos <OPTION VALUE="P" SELECTED>Apenas professores <OPTION VALUE="T">Professores e alunos <OPTION VALUE="E">Escola');
    } else {
       ShowHTML ('            <OPTION VALUE="A">Apenas alunos <OPTION VALUE="P">Apenas professores <OPTION VALUE="T" SELECTED>Professores e alunos <OPTION VALUE="E">Escola');
    }
    ShowHTML ('            </SELECTED></TD>');
    MontaRadioSN ('<b>Exibir no site?</b>', $w_in_ativo, 'w_in_ativo');
    ShowHTML ('        </table>');
    ShowHTML ('      <tr>');
    ShowHTML ('      <tr><td align="center" colspan=4><hr>');
    if ( $O == "E" ) {
       ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir" onClick="return confirm(\'Confirma a exclusão do registro?\');">');    
       //ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir">"
    } else {
       if ( O == "I" ) {
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Incluir">');
       } else {
          ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
       }
    }
    ShowHTML ('            <input class="STB" type="button" onClick="location.href=\'' .$dir.$w_pagina.$par. $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
    ShowHTML ('          </td>');
    ShowHTML ('      </tr>');
    ShowHTML ('    </table>');
    ShowHTML ('    </TD>');
    ShowHTML ('</tr>');
    ShowHTML ('</FORM>');
  } else {
    ScriptOpen ('JavaScript');
    ShowHTML (' alert(\'Opção não disponível\');');
    //ShowHTML (' history.back(1);');
    ScriptClose();
  }
  ShowHTML ('</table>');
  ShowHTML ('</center>');
  Rodape();

}
// =========================================================================
// Fim do cadastro de arquivos
// -------------------------------------------------------------------------


// =========================================================================
// Monta o menu principal da aplicação
// -------------------------------------------------------------------------
function menu() {
  // Inclusão do arquivo da classe
  include_once("classes/menu/xPandMenu.php");
  // Instanciando a classe menu
  $root = new XMenu();
  
  $w_imagem = 'img/SheetLittle.gif';
  $i    = 0;
  $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Manual SIGE-WEB\',\'manuais/operacao/\',$w_Imagem,$w_Imagem,\'_blank\', null));');
  if ($_SESSION['USERNAME']=='ADMINISTRATIVO' || $_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='SBPI') { 
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Administrativo\',$w_pagina.\'administrativo\',$w_Imagem,$w_Imagem,\'content\', null));');
  }
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='IMPRENSA' || $_SESSION['USERNAME']=='SBPI') { 
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Arquivos (<i>download</i>)\',$w_pagina.\'arquivos\',$w_Imagem,$w_Imagem,\'content\', null));');
  }
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='SBPI') { 
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Calendário Base\',$w_pagina.\'calend_base\',$w_Imagem,$w_Imagem,\'content\', null));');
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Calendário da Rede\',$w_pagina.\'calend_rede\',$w_Imagem,$w_Imagem,\'content\', null));');
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Dados da escola\',$w_pagina.\'ligacao\',$w_Imagem,$w_Imagem,\'content\', null));');
  }
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='IMPRENSA' || $_SESSION['USERNAME']=='SBPI') { 
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Envio de e-Mail\',$w_pagina.\'email\',$w_Imagem,$w_Imagem,\'content\', null));');
  }
  if ($_SESSION['USERNAME'] != 'ADMINISTRATIVO' || $_SESSION['USERNAME']=='SBPI') { 
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Escolas\',$w_pagina.\'escolas\',$w_Imagem,$w_Imagem,\'content\', null));');  
  }
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='SBPI') { 
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Escolas Particulares\',$w_pagina.\'escpart\',$w_Imagem,$w_Imagem,\'content\', null));');  
  }
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='SBPI') { 
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Homologação de Calendário\',$w_pagina.\'escparthomolog\',$w_Imagem,$w_Imagem,\'content\', null));');  
  }
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='IMPRENSA' || $_SESSION['USERNAME']=='SBPI') {  
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Informativo\',$w_pagina.\'newsletter\',$w_Imagem,$w_Imagem,\'content\', null));');  
  }  
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='SBPI') { 
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Modalidades de Ensino\',$w_pagina.\'modalidades\',$w_Imagem,$w_Imagem,\'content\', null));');  
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Mensagens a alunos\',$w_pagina.\'msgalunos\',$w_Imagem,$w_Imagem,\'content\', null));');      
  }
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='IMPRENSA' || $_SESSION['USERNAME']=='SBPI') {  
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Notícias\',$w_pagina.\'noticias\',$w_Imagem,$w_Imagem,\'content\', null));');  
  }  
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='SBPI') {   
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Rede Particular\',$w_pagina.\'redepart\',$w_Imagem,$w_Imagem,\'content\', null));');    
  }
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='SBPI') {   
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Tipo de Instituição\',$w_pagina.\'tipocliente\',$w_Imagem,$w_Imagem,\'content\', null));');    
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Senha\',$w_pagina.\'senha\',$w_Imagem,$w_Imagem,\'content\', null));');    
  }  
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='SBPI') {   
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Senhas Especiais\',$w_pagina.\'senhaesp\',$w_Imagem,$w_Imagem,\'content\', null));');    
  }    
  if ($_SESSION['USERNAME']=='SBPI') {   
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Verif. Arquivos\',$w_pagina.\'verifarq\',$w_Imagem,$w_Imagem,\'content\', null));');    
  }  
  if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='SBPI') {   
    $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Log\',$w_pagina.\'log\',$w_Imagem,$w_Imagem,\'content\', null));');    
  }      
  $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Sair do sistema\',$w_pagina.\'Sair\',$w_Imagem,$w_Imagem,\'_top\', \'onClick="return(confirm(\\\'Confirma saída do sistema?\\\'));"\' ));');

  // Quando for concluída a montagem dos nós, chame a função generateTree(), usando o objeto raiz, para gerar o código HTML.
  // Essa função não possui argumentos.
  // No código da função pode ser verificado que há um parâmetro opcional, usado internamente para chamadas recursivas, necessárias à montagem de toda a árvore.
  $menu_html_code = $root->generateTree();

  // A função retornou o código HTML para exibir o menu


  // Montando a página:
  // 3 pontos:
  // - Referencie o arquivo Javascript
  // - Referencie o arquivo CSS
  // - Exiba o código HTML gerado anteriormente
  ShowHTML ('<html>');
  ShowHTML ('<head>');
  ShowHTML ('  <!-- CSS FILE for my tree-view menu -->');
  ShowHTML ('  <link rel="stylesheet" type="text/css" href="classes/menu/xPandMenu.css">');
  ShowHTML ('  <!-- JS FILE for my tree-view menu -->');
  ShowHTML ('  <script src="classes/menu/xPandMenu.js"></script>');
  ShowHTML ('</head>');
  ShowHTML ('<BASEFONT FACE="Verdana, Helvetica, Sans-Serif" SIZE="2">');
  // Decide se montará o body do menu principal ou o body do sub-menu de uma opção a partir do valor de w_sq_pagina
  echo('<BODY topmargin=0 bgcolor="#FFFFFF" BACKGROUND="img/background.gif" BGPROPERTIES="FIXED" text="#000000" link="#000000" vlink="#000000" alink="#FF0000" ');
  ShowHTML ('onLoad="javascript:top.content.location=\''.$w_pagina.'escolas\';"> ');
  ShowHTML ('  <CENTER><table border=0 cellpadding=0 height="80" width="100%">');
  ShowHTML ('      <tr><td colspan=2 width="100%"><table border=0 width="100%" cellpadding=0 cellspacing=0><tr valign="top">');
  ShowHTML ('          <td>Usuário:<b>'.$_SESSION['USERNAME'].'</b></TD>');
  ShowHTML ('          <td align="right"><a class="hl" href="help.php?par=Menu&TP=<img src=images/Folder/hlp.gif border=0> SIW - Visão Geral&SG=MESA&O=L" target="content" title="Exibe informações sobre os módulos do sistema."><img src="images/Folder/hlp.gif" border=0></a></TD>');
  ShowHTML ('          </table>');
  ShowHTML ('      <tr><td height=1><tr><td height=2 bgcolor="#000000">');
  ShowHTML ('  </table></CENTER>');
  ShowHTML ('  <table border=0 cellpadding=0 height="80" width="100%"><tr><td nowrap><b>');
  ShowHTML ('  <div id="container">');
  echo $menu_html_code;
  ShowHTML ('  </div>');
  ShowHTML ('  </table>');
  ShowHTML ('</body>');
  ShowHTML ('</html>');
}

// =========================================================================
// Rotina de encerramento da sessão
// -------------------------------------------------------------------------
function Sair() {
  extract($GLOBALS);
  // Eliminar todas as variáveis de sessão.
  $_SESSION = array();
  // Finalmente, destruição da sessão.
  session_destroy();

  ScriptOpen('JavaScript');
  ShowHTML ('  top.location.href=\''.$w_pagina.'\';');

  ScriptClose();
}

// =========================================================================
// Rotina principal
// -------------------------------------------------------------------------
function Main(){
  extract($GLOBALS);
  switch ($par) {
    Case 'MENU':                   Menu();            break;
    Case 'ADMINISTRATIVO':         Administrativo();  break;
    Case 'ARQUIVOS':               Arquivos();        break;
    Case 'CALEND_REDE':            Calend_rede();     break;
    Case 'CALEND_BASE':            Calend_base();     break;
    Case 'ESCPARTHOMOLOG':         escPartHomolog();  break;
    Case 'GRAVA':                  Grava();           break;
    Case 'LIGACAO':                Ligacao();         break;
    Case 'MSGALUNOS':              Msgalunos();       break;
    Case 'MODALIDADES':            Modalidades();     break;
    Case 'NEWSLETTER':             Newsletter();      break;
    Case 'NOTICIAS':               Noticias();        break;
    Case 'SENHA':                  Senha();           break;
    Case 'SENHAESP':               Senhaesp();        break;
    Case 'TIPOCLIENTE':            TipoCliente();     break;
    Case 'VALIDA':                 Valida();          break;
    Case 'FRAMES':                 Frames();          break;
    Case 'SAIR':                   Sair();            break;
    /*
    Case 'ADM':                        ShowAdmin; break;
    Case 'BASE':                       GetCalendarioBase; break;
    Case conWhatAdmin:                 Administrativo; break;
    Case conWhatCliente:               GetCliente; break;
    Case conWhatDadosAdicionais:       GetDadosAdicionais; break;
    Case conWhatSite:                  GetSite; break;
    Case conWhatNotCliente:            GetNoticiaCliente; break;
    Case conWhatCalendario:            GetCalendario; break;
    Case conWhatEspecialidadeCliente:  GetEspecialidadeCliente; break;
    Case conWhatDocumento:             GetDocumento; break;
    Case conWhatCalendario:            GetCalendario; break;
    Case conWhatMensagem:              GetMensagem; break;
    Case 'SENHAESP':                   ShowSenhaEspecial; break;
    Case 'NEWSLETTER':                 GetNewsletter; break;
    Case conWhatEspecialidade:         GetEspecialidade; break;
    Case conRelEscolas:                ShowEscolas; break;
    Case 'ESCPART':                    ShowEscolaParticular; break;
    Case 'ESCPARTHOMOLOG':             ShowEscolaParticularHomolog; break;
    Case 'GETESCOLAS':                 GetEscolas; break;
    Case 'CADASTROESCOLA':             CadastroEscolas; break;
    Case conWhatSGE:                   GetSistema; break;
    Case 'COMPONENTE':                 GetComponente; break;
    Case 'VERSAO':                     GetVersao; break;
    Case 'TIPOCLIENTE':                GetTipoCliente; break;
    Case 'VERIFARQ':                   GetVerifArquivo; break;
    Case 'REDEPART':                   GetRedeParticular; break;
    Case 'GRAVA':                      Grava; break;
    */
    default:
      Cabecalho();
      ShowHTML ('<BASE HREF="'.$conRootSIW.'">');
      BodyOpen('onLoad=this.focus();');
      Estrutura_Topo_Limpo();
      Estrutura_Menu();
      Estrutura_Corpo_Abre();
      Estrutura_Texto_Abre();
      ShowHTML ('<div align=center><center><br><br><br><br><br><br><br><br><br><br><img src="images/icone/underc.gif" align="center"> <b>Esta opção está sendo desenvolvida.</b><br><br><br><br><br><br><br><br><br><br></center></div>');
      Estrutura_Texto_Fecha();
      Estrutura_Fecha();
      Estrutura_Fecha();
      Estrutura_Fecha();
      Rodape();
      break;
  }
}
?>