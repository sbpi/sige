<?php
// Garante que a sessão será reinicializada.
session_start();

$w_dir_volta = '';
$w_pagina       = 'manut.php?par=';
$w_Disabled     = 'ENABLED';
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
//  /manut.php
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

$w_Data = date('d/m/Y');

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
  ShowHTML('<HTML>');
  ShowHTML('<HEAD>');
  ShowHTML('<link rel="shortcut icon" href="'.$conRootSIW.'favicon.ico" type="image/ico" />');
  ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
  ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
  ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
  ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
  ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
  ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
  ShowHTML('<TITLE>'.$conSgSistema.' - Autenticação</TITLE>');
  ScriptOpen('JavaScript');
  ShowHTML('$(document).ready(function(){');
  ShowHTML('  $("#Login1").change(function(){');
  ShowHTML('    formataCampo();');
  ShowHTML('  })');
  ShowHTML('});');
  ShowHTML('function formataCampo(){');
  ShowHTML('  $("#Login1").val(trim($("#Login1").val()));');
  ShowHTML('  if(  $("#Login1").val().length==11 ..  caracterAceito( $("#Login1").val() ,  "0123456789") ){');
  ShowHTML('    $("#Login1").val( mascaraGlobal(\'###.###.###-##\',$("#Login1").val()) );');
  ShowHTML('  }');
  ShowHTML('}');
  ShowHTML('function caracterAceito(string , checkOK){');
  ShowHTML('   //var checkOK = \'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-!@#$%.*()+/\';');
  ShowHTML('      var checkStr = string;');
  ShowHTML('      var allValid = true;');
  ShowHTML('      for (i = 0;  i < checkStr.length;  i++)');
  ShowHTML('      {');
  ShowHTML('      ch = checkStr.charAt(i);');
  ShowHTML('      if ((checkStr.charCodeAt(i) != 13) .. (checkStr.charCodeAt(i) != 10) .. (checkStr.charAt(i) != "\\\\")) {');
  ShowHTML('         for (j = 0;  j < checkOK.length;  j++) {');
  ShowHTML('         if (ch == checkOK.charAt(j))');
  ShowHTML('           break;');
  ShowHTML('         }');
  ShowHTML('         if (j == checkOK.length)');
  ShowHTML('         {');
  ShowHTML('         allValid = false;');
  ShowHTML('         break;');
  ShowHTML('         }');
  ShowHTML('      }');
  ShowHTML('      }');
  ShowHTML('      return allValid;');
  ShowHTML('}');
  ShowHTML('function Ajuda() ');
  ShowHTML('{ ');
  ShowHTML('  document.Form.Botao.value = "Ajuda"; ');
  ShowHTML('} ');
  Modulo();
  SaltaCampo();
  ValidateOpen('Validacao');
  Validate('Login1','Nome de usuário','','1','2','30','1','1');
  Validate('Password1','Senha','1','1','3','19','1','1');
  ShowHTML('  theForm.Login.value = theForm.Login1.value; ');
  ShowHTML('  theForm.Password.value = theForm.Password1.value; ');
  ShowHTML('  theForm.Login1.value = ""; ');
  ShowHTML('  theForm.Password1.value = ""; ');
  ValidateClose();
  ScriptClose();
  ShowHTML('<link rel="stylesheet" type="text/css" href="'.$conRootSIW.'classes/menu/xPandMenu.css">');
  ShowHTML('<style>');
  ShowHTML(' .cText {font-size: 8pt; border: 1px solid #000000; background-color: #F5F5F5}');
  ShowHTML(' .cButton {font-size: 8pt; color: #FFFFFF; border: 1px solid #000000; background-color: #669966; }');
  ShowHTML('</style>');
  ShowHTML('</HEAD>');
  // Se receber a username, dá foco na senha
  if (nvl($w_username,'nulo')=='nulo') {
      ShowHTML('<body topmargin=0 leftmargin=10 onLoad=\'document.Form.Login1.focus();\'>');
  } else {
      ShowHTML('<body topmargin=0 leftmargin=10 onLoad=\'document.Form.Password1.focus();\'>');
  }
  ShowHTML ('<CENTER>');
  ShowHTML ('<form method="post" action="manut.php?par=valida" onsubmit="return(Validacao(this));" name="Form"> ');
  ShowHTML('<INPUT TYPE="HIDDEN" NAME="Login" VALUE=""> ');
  ShowHTML('<INPUT TYPE="HIDDEN" NAME="Password" VALUE=""> ');
  ShowHTML('<INPUT TYPE="HIDDEN" NAME="p_dbms" VALUE="1"> ');
  ShowHTML ('<TABLE cellSpacing=0 cellPadding=0 width="760" height=550 border=1  background="files/" . p_cliente . "/img/fundo.jpg" bgproperties="fixed"><tr><td width="100%" valign="top">');
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
  ShowHTML ('              . <a class="SS" href="manuais/regionais/" target="_blank" title="Exibe versão HTML do manual de operação do SIGE-WEB (Diretorias Regionais de Ensino)">Manual SIGE-WEB para DRE</A><BR>');
  ShowHTML ('              . <a class="SS" href="manuais/operacao/" target="_blank" title="Exibe  versão HTML do manual de operação do SIGE-WEB (Outros)">Manual SIGE-WEB para demais I.E.</A><BR>');
  ShowHTML ('              <br></FONT></P>');
  ShowHTML ('              </TD></TR>');
  ShowHTML ('</table>');
  ShowHTML ('              </TD></TR>');
  ShowHTML ('          </table>   ');
  $wAno =  $_REQUESR["wAno"];
  
  If ($wAno == "" ){
     $wAno = date('Y');
  }

  $SQL = "SELECT case in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end  in_destinatario, " . $crlf . 
        "       dt_arquivo, ds_titulo, nr_ordem, ds_arquivo, ln_arquivo, 'SEDF' AS Origem, x.ln_internet diretorio " . $crlf . 
        " From sbpi.Cliente_Arquivo a INNER JOIN sbpi.Cliente  x ON (a.sq_cliente = x.sq_cliente)" . $crlf . 
        " WHERE x.ativo = 'Sim'" . $crlf . 
        "  AND x.sq_cliente = 0" . $crlf . 
        "  AND in_destinatario = 'E'" . $crlf . 
        "  and sbpi.YEAR(a.dt_arquivo) = " . $wAno . $crlf . 
        " ORDER BY origem, nr_ordem, dt_arquivo desc, in_destinatario " . $crlf ;

  ShowHTML (' <tr><td><table width="100%" border="0">');
  ShowHTML ('          <TR><TD align="center"><br><table border=0 cellpadding=0 cellspacing=0><tr><td>');
  ShowHTML ('              <P><font face="Arial" size=1><b>ARQUIVOS INSERIDOS PELA SEDF</b></font><br>');
  ShowHTML ('<tr><td><TABLE border=0 cellSpacing=5 width="95%">');
  ShowHTML ('  <TR>');
  ShowHTML ('    <TD><FONT face="Verdana" size=1><b>Origem');
  ShowHTML ('    <TD><FONT face="Verdana" size=1><b>Alvo');
  ShowHTML ('    <TD><FONT face="Verdana" size=1><b>Data');
  ShowHTML ('    <TD><FONT face="Verdana" size=1><b>Componente curricular');
  ShowHTML ('  </TR>');
  ShowHTML ('  <TR>');
  ShowHTML ('    <TD COLSPAN="4" HEIGHT="1" BGCOLOR="#DAEABD">');
  ShowHTML ('  </TR>');
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);      
  if(count($RS) > 0){
    foreach($RS as $row){

        ShowHTML ('  <TR valign="top">');
        ShowHTML ('    <TD><FONT face="Verdana" size=1>' . $row["origem"]);
        ShowHTML ('    <TD><FONT face="Verdana" size=1>' . $row["in_destinatario"]);
       // ShowHTML ('    <TD><FONT face="Verdana" size=1>' . substr(100+Day($row["dt_arquivo"]),1,2) . "/" . substr(100+Month($row["dt_arquivo"]),1,2) . "/" . Year($row["dt_arquivo"]);
      ShowHTML ('    <TD><FONT face="Verdana" size=1>' . $row["dt_arquivo"]);
//   REM  ShowHTML "    <TD><FONT face="Verdana" size=1><a href="http://" . replace(RS("diretorio"),"http://",") . "//" . RS("ln_arquivo") . "" target="_blank">" . RS("ds_titulo") . "</a><br><div align="justify"><font size=1>.:. " . RS("ds_arquivo") . "</div>" Original
    ShowHTML ('    <TD><FONT face="Verdana" size=1><a href="http://' . str_replace(str_replace($row["diretorio"],"http://",'') , "se.df.gov.br" , "gdfsige.df.gov.br") . "/sedf/sedf/" . $row["ln_arquivo"] . ' target="_blank">' . $row["ds_titulo"] . '</a><br><div align="justify"><font size=1>.:. ' . $row["ds_arquivo"] . '</div>');
        ShowHTML ('  </TR>');
    }
  }Else{
     ShowHTML ('  <TR><TD COLSPAN=4 ALIGN=CENTER><FONT face="Verdana" size=1>Não há arquivos disponíveis no momento para o ano de ' . $wAno . ' </TR>');
  }


  $SQL = "SELECT sbpi.year(dt_arquivo) ano " . $crlf . 
        " From sbpi.Cliente_Arquivo a INNER JOIN sbpi.Cliente x ON (a.sq_cliente = x.sq_cliente)" . $crlf . 
        " WHERE x.ativo = 'Sim'" . $crlf . 
        "  AND x.sq_cliente = 0" . $crlf . 
        "  AND in_destinatario = 'E'" . $crlf . 
        "  and sbpi.YEAR(a.dt_arquivo) <> " . $wAno . $crlf . 
        " ORDER BY sbpi.year(dt_arquivo) desc " . $crlf ;

    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    if(count($RS) > 0){
    ShowHTML ('  <TR><TD COLSPAN=4 ><FONT face="Verdana" size=1><b>Arquivos de outros anos</b><br>');
    foreach($RS as $row){
                 ShowHTML ('     <li><a href="'.  $w_dir . 'Manut.php?$wAno=' . $row["ano"] . '" >Arquivos de ' . $row["ano"] . '</a></TR>');
        }
       ShowHTML ('</TD></TR>');
  }

  ShowHTML ('</table>');
  ShowHTML ('              </FONT></P>');
  ShowHTML ('              </TD></TR>');
  ShowHTML ('          </table>');  
  ShowHTML ('        </table>');
  ShowHTML ('    </tr>');
  ShowHTML ('  </table>');
  ShowHTML ('</table>');   
  ShowHTML ('</form>');  
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
    $SQL = "select count(*) existe from sbpi.Cliente where upper(ds_username) = upper('" . $w_uid . "') and upper(ds_senha_acesso) = upper('" . $w_pwd . "')";
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows); 
    foreach($RS as $row) { $RS = $row; break; }
    if(f($RS,"existe") == 0){
      $w_erro = 2;          
      }        
  }
  ScriptOpen ('JavaScript');
  if ($w_erro > 0) {
    if ($w_erro == 1)     ShowHTML (' alert("Usuário inexistente!");');
    elseif ($w_erro == 2) ShowHTML (' alert("Senha inválida!");');
    
    ShowHTML ('  history.back(1);');
  } else {
    //Recupera informações a serem usadas na montagem das telas para o usuário
    $SQL = "select a.ds_username , a.sq_cliente , b.tipo from sbpi.Cliente a inner join sbpi.tipo_cliente b on a.sq_tipo_cliente = b.sq_tipo_cliente where upper(ds_username) = upper('" . $w_uid . "') and upper(ds_senha_acesso) = upper('" . $w_pwd . "')";
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows); 
    foreach($RS as $row) { $RS = $row; break; }

    $_SESSION['USERNAME'] = strtoupper(f($RS, 'ds_username'));
    $_SESSION['CL']       = f($RS, 'sq_cliente');
   $w_tipo       = f($RS,'tipo');
    
    If ($_SESSION['USERNAME'] != 'SBPI'){
      //Grava o acesso na tabela de log
      $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log,sq_cliente, data, ip_origem, tipo, abrangencia, sql) " . $crlf . 
            "values ( " . $crlf .       
            "         sbpi.sq_cliente_log.nextval , " . $crlf . 
            "         " . $_SESSION['CL'] . ", " . $crlf . 
            "         sysdate, " . $crlf . 
            "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf . 
            "         0, " . $crlf . 
            "         'Acesso à tela de atualização da escola.', " . $crlf . 
            "         null " . $crlf . 
            "       ) " . $crlf;
      $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    }
    ShowHTML (' location.href="'.$w_pagina.'frames";');
  }
  ScriptClose();
}

// =========================================================================
// Exportação dos dados administrativos
// -------------------------------------------------------------------------
function administrativo(){
  extract($GLOBALS);
  Cabecalho();
  ShowHTML('<HEAD>');
  ScriptOpen('Javascript');
  ValidateOpen('Validacao');
  ShowHTML('  if (theForm.w_arquivo[0].checked == false .. theForm.w_arquivo[1].checked == false) {');
  ShowHTML('     alert(\'Você deve escolher uma das opções apresentadas antes de gerar o arquivo!\');');
  ShowHTML('     return false;');
  ShowHTML('  }');
  ShowHTML('  return(confirm(\'Confirma a geração do arquivo com os dados indicados?\'));');
  ValidateClose();
  ScriptClose();
  ShowHTML('</HEAD>');
  BodyOpen('onLoad=\'document.focus();\'');
  ShowHTML('<B><FONT COLOR="#000000">Exportação dos dados administrativos</FONT></B>');
  ShowHTML('<HR>');
  ShowHTML('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
  AbreForm('Form', $w_Pagina.'Grava', 'POST', 'return(Validacao(this));', null);
  ShowHTML(MontaFiltro('POST'));
  ShowHTML('<input type="hidden" name="R" value="'.$R.'">');
  ShowHTML('<tr bgcolor="'.'#EFEFEF'.'"><td align="center">');
  ShowHTML('    <table width="97%" border="0">');
  ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
  ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Exportação dos dados administrativos</td></td></tr>');
  ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML('      <tr><td><font size=1><ul><li>Esta tela permite a exportação, para um arquivo que pode ser aberto no Excel, dos dados administrativos preenchidos pelas unidades de ensino.<li>Permite também exportar as tabelas de apoio utilizadas pelo formulário.<li>Selecione uma das opções exibidas abaixo e clique no botão "Gerar arquivo" para que os dados sejam convertidos para um arquivo.</ul></font></td></tr>');
  ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML('      <tr><td valign="top"><font size="1"><b>O arquivo a ser gerado deve conter dados:</b>');
  ShowHTML('          <br><INPUT '.$w_Disabled.' class="BTM" type="radio" name="w_arquivo" value="Escola"> das unidades de ensino');
  ShowHTML('          <br><INPUT '.$w_Disabled.' class="BTM" type="radio" name="w_arquivo" value="Tipo"> da tabela de equipamentos');
  ShowHTML('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');

  // Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
  ShowHTML('      <tr><td align="center"><input class="STB" type="submit" name="Botao" value="Gerar arquivo"></td>');
  ShowHTML('      </tr>');
  ShowHTML('    </table>');
  ShowHTML('    </TD>');
  ShowHTML('</tr>');
  ShowHTML('</FORM>');
  ShowHTML('</table>');
  Rodape();
}

// =========================================================================
// Monta a Frame
// -------------------------------------------------------------------------
Function frames(){
    extract($GLOBALS);
    $_SESSION["BodyWidth"] = "620";

    ShowHTML ('<html>');
    ShowHTML ('<head>');
    ShowHTML ('    <title>Controle Central</title>');
    ShowHTML ('</head>');
    ShowHTML ('<frameset cols="200,*">');
    ShowHTML ('    <frame name="menu" src="'.$w_pagina.'menu" scrolling="auto" marginheight="0" marginwidth="0">');
    ShowHTML ('    <frame name="content" src="'.$w_pagina.'getsite" scrolling="auto" marginheight="0" marginwidth="0">');
    ShowHTML ('</frameset>');
    ShowHTML ('</html>');
}

function escolas(){
    extract($GLOBALS);
  echo "aaaa";
}


// =========================================================================
// Tela de dados do site
// -------------------------------------------------------------------------
function GetSite(){
  extract($GLOBALS);


    
   $SQL = "select a.*, b.ds_diretorio tipo from sbpi.Cliente_Site a inner join sbpi.Modelo b on (a.sq_modelo = b.sq_modelo) where a.sq_cliente = " . $CL; 
   $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);


   foreach($RS as $row) { $RS = $row; break; }

     $w_sq_cliente           = $RS["sq_cliente"];
     $w_no_contato_internet  = $RS["no_contato_internet"];
     $w_ds_email_internet    = $RS["ds_email_internet"];
     $w_nr_fone_internet     = $RS["nr_fone_internet"];
     $w_nr_fax_internet      = $RS["nr_fax_internet"];
     $w_ds_texto_abertura    = $RS["ds_texto_abertura"];
     $w_ds_institucional     = $RS["ds_institucional"];
     $w_ds_mensagem          = $RS["ds_mensagem"];
     $w_ds_diretorio         = $RS["ds_diretorio"];
     $w_pedagogica           = $RS["ln_prop_pedagogica"];
     //$w_tipo                 = str_replace("Mod","",$RS["tipo"]);
   

  Cabecalho();
  ShowHTML ("<HEAD>");
  ScriptOpen ("Javascript");
  ValidateOpen ("Validacao");

     Validate ("w_no_contato_internet", "Nome", "1", "1", "2", "35", "1", "1");
     Validate ("w_ds_email_internet", "e-Mail", "1", "1", "6", "60", "1", "1");
     Validate ("w_nr_fone_internet", "Telefone", "1", "1", "6", "20", "1", "1");
     Validate ("w_nr_fax_internet", "Fax", "1", "", "6", "20", "1", "1");
     Validate ("w_ds_texto_abertura", "Texto de abertura", "1", "1", "4", "8000", "1", "1");
     Validate ("w_ds_institucional", "Texto da seção \"Quem somos\" ", "1", "1", "4", "8000", "1", "1");
     if($w_tipo != 13){
        If ($w_tipo == 2) { //' Se for regional
           Validate ("w_pedagogica", "Composição administrativa", "1", "", "5", "100", "1", "1");
        }Else{
           Validate ("w_pedagogica", "Projeto", "1", "", "5", "100", "1", "1");
        }
        ShowHTML (" if (theForm.w_pedagogica.value != ''){");
        ShowHTML ("    if((theForm.w_pedagogica.value.lastIndexOf('.PDF')==-1) .. (theForm.w_pedagogica.value.lastIndexOf('.pdf')==-1) .. (theForm.w_pedagogica.value.lastIndexOf('.DOC')==-1) .. (theForm.w_pedagogica.value.lastIndexOf('.doc')==-1)) {");
        ShowHTML ("       alert('Esolha arquivos com as extensões \'.doc\' ou \'.pdf\'!');");
        ShowHTML ("       theForm.w_pedagogica.value=''; ");
        ShowHTML ("       theForm.w_pedagogica.focus(); ");
        ShowHTML ("       return false;");
        ShowHTML ("    }");
        ShowHTML ("  }");
     }
     Validate ("w_ds_mensagem", "Texto da mensagem em destaque", "1", "1", "4", "80", "1", "1");
	
  ValidateClose();
  ScriptClose();
  ShowHTML ("</HEAD>");
  BodyOpen ("onLoad='document.Form.w_no_contato_internet.focus();'");
  ShowHTML ('<B><FONT COLOR="#000000">Atualização de dados do site da unidade</FONT></B>');
  ShowHTML ('<HR>');
  ShowHTML ('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
  ShowHTML ('<FORM action="manut.php?par=Grava" method="POST" name="Form" onSubmit="return(Validacao(this));" enctype="multipart/form-data">');
  ShowHTML (MontaFiltro("POST"));
  ShowHTML ('<input type="hidden" name="SG" value="getsite">');
  ShowHTML ('<tr bgcolor="" . "#EFEFEF" . ""><td align="center">');
  ShowHTML ('    <table width="97%" border="0">');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Contato da unidade para divulgação no site</td></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td><font size=1>Informe os dados da pessoa a ser exibida como contato da unidade para divulgação no site.</font></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
  ShowHTML ('        <tr valign="top">');
  ShowHTML ('          <td><font size="1"><b>No<u>m</u>e:</b><br><INPUT ACCESSKEY="M"  '. $w_Disabled . ' class="STI" type="text" name="w_no_contato_internet" size="35" maxlength="35" value="'. $w_no_contato_internet .'" title="OBRIGATÓRIO. Informe o nome da pessoa a ser exibida na seção \"Quem somos\" do site da unidade."></td>');
  ShowHTML ('          <td><font size="1"><b><u>e</u>-Mail:</b><br><INPUT ACCESSKEY="E" '. $w_Disabled . ' class="STI" type="text" name="w_ds_email_internet" size="40" maxlength="60" value="' . $w_ds_email_internet .'" title="OBRIGATÓRIO. Informe o e-mail da pessoa." ></td');
  ShowHTML ('        <tr valign="top">');
  ShowHTML ('          <td><font size="1"><b>Tele<u>f</u>one:</b><br><INPUT ACCESSKEY="F" '. $w_Disabled .' class="STI" type="text" name="w_nr_fone_internet" size="20" maxlength="20" value="'. $w_nr_fone_internet .'" title="OBRIGATÓRIO. Informe o telefone da pessoa."></td>');
  ShowHTML ('          <td><font size="1"><b>Fa<u>x</u>:</b><br><INPUT ACCESSKEY="X" '. $w_Disabled .' class="STI" type="text" name="w_nr_fax_internet" size="20" maxlength="20" value="' . $w_nr_fax_internet . '" title="OPCIONAL. Informe o fax da pessoa."></td>');
  ShowHTML ('        </table>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Página de abertura do site</td></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td><font size=1>Informe o texto a ser colocado na página de abertura do site.</font></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td><font size="1"><b>Texto da <u>p</u>ágina de abertura:</b><br><TEXTAREA ACCESSKEY="P" ' . $w_Disabled . ' class="STI" type="text" name="w_ds_texto_abertura" rows=5 cols=65 title="OBRIGATÓRIO. Informe o texto a ser exibido na página de abertura do site.">' . $w_ds_texto_abertura . '</TEXTAREA></td>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Página "Quem somos"</td></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td><font size=1>Informe o texto ser colocado na página "Quem somos" do site.</font></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
  ShowHTML ('      <tr><td><font size="1"><b>Texto da seção "<u>Q</u>uem somos":</b><br><TEXTAREA ACCESSKEY="Q" ' . $w_Disabled . ' class="STI" type="text" name="w_ds_institucional" rows=5 cols=65 title="OBRIGATÓRIO. Informe o texto a ser exibido na seção.">' . $w_ds_institucional . "</TEXTAREA></td>");
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  If ($w_tipo == 2) { // ' Se for regional
     ShowHTML ('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Página "Composição administrativa"</td></td></tr>');
     ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
     ShowHTML ('      <tr><td><font size=1>Informe o arquivo Word ou PDF a ser exibido na página "Composição administrativa" do site.');
     ShowHTML ('          <br><font color="red"><b>IMPORTANTE: <a href="sedf/orientacoes_word.pdf" class="hl" target="_blank">Para documentos Word, clique aqui para ler as orientações sobre a formatação e a proteção do texto</a></b></font>.');
     ShowHTML ('      </font></td></tr>');
     ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
     ShowHTML ('      <tr><td><font size="1"><b>Composição adminis<u>t</u>rativa (arquivo Word ou PDF):</b><br><INPUT ACCESSKEY="T" '.  $w_Disabled . ' class="STI" type="file" name="w_pedagogica" size="60" maxlength="100" value="" title="OPCIONAL. Clique no botão ao lado para localizar o arquivo que contém a composição administrativa da regional. Ele será transferido automaticamente para o servidor.">');
  } Else{
     //'SQL = "select b.ds_especialidade from escEspecialidade_Cliente a inner join escEspecialidade b on (a.sq_codigo_espec = b.sq_especialidade and a." . CL . ")"
     If ($w_tipo != 13){
     //'   If uCase(RS("ds_especialidade")) <> uCase("Biblioteca") Then
           ShowHTML ('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Página "Projeto"</td></td></tr>');
           ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
           ShowHTML ('      <tr><td><font size=1>Informe o arquivo Word ou PDF a ser exibido na página "Projeto" do site.');
           ShowHTML ('          <br><font color="red"><b>IMPORTANTE: <a href="sedf/orientacoes_word.pdf" class="hl" target="_blank">Para documentos Word, clique aqui para ler as orientações sobre a formatação e a proteção do texto</a></b></font>.');
           ShowHTML ('      </font></td></tr>');
           ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
           ShowHTML ('      <tr><td><font size="1"><b>Proje<u>t</u>o (arquivo Word):</b><br><INPUT ACCESSKEY="T" ' . $w_Disabled . ' class="STI" type="file" name="w_pedagogica" size="60" maxlength="100" value="" title="OPCIONAL. Clique no botão ao lado para localizar o arquivo que contém o projeto da escola. Ele será transferido automaticamente para o servidor.">');
   }
  }
  If ($w_pedagogica > ''){
     ShowHTML ('              <b><a class="SS" href="http://" '. str_replace(strtolower("http://"),'',strtolower($w_ds_diretorio)) . "/" . $w_pedagogica . ' target="_blank" title="Clique para abrir o arquivo atual.">Exibir</a></b>');
  }
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Diversos</td></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td><font size=1>Informe os dados abaixo de exibição geral no site.</font></td></tr>');
  ShowHTML ('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
  ShowHTML ('      <tr><td><font size="1"><b>Texto da men<u>s</u>agem em destaque:</b><br><INPUT ACCESSKEY="S" '. $w_Disabled . ' class="STI" type="text" name="w_ds_mensagem" size="80" maxlength="80" value="'. $w_ds_mensagem .'" title="OBRIGATÓRIO. Informe um texto que será exibido na parte superior do site, numa barra rolante."></td>');
  ShowHTML ('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');

  //' Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
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
// Monta o menu principal da aplicação
// -------------------------------------------------------------------------
function menu() {
    extract($GLOBALS);


  // Inclusão do arquivo da classe
  include_once("classes/menu/xPandMenu.php");

  // Instanciando a classe menu
  $root = new XMenu();
  
  $w_imagem = 'img/SheetLittle.gif';
  $i    = 0;

  $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Manual SIGE-WEB\',\'manuais/operacao/\',$w_Imagem,$w_Imagem,\'_blank\', null));');
  if ($_SESSION['TIPO']==3) { 
    $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Administrativo\',$w_pagina.\'administrativo\',$w_Imagem,$w_Imagem,\'content\', null));');
  }
  $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Fotos\',$w_pagina.\'fotos\',$w_Imagem,$w_Imagem,\'content\', null));');
  $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Dados básicos\',$w_pagina.\'basico\',$w_Imagem,$w_Imagem,\'content\', null));');
  $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Dados adicionais\',$w_pagina.\'adicionais\',$w_Imagem,$w_Imagem,\'content\', null));');
  $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Dados do site\',$w_pagina.\'getsite\',$w_Imagem,$w_Imagem,\'content\', null));');
  if ($_SESSION['TIPO']==3) { 
    $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Áreas de atuação\',$w_pagina.\'atuacao\',$w_Imagem,$w_Imagem,\'content\', null));');
  }
  $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Arquivos (<i>download</i>)\',$w_pagina.\'arquivos\',$w_Imagem,$w_Imagem,\'content\', null));');

   if ($_SESSION['TIPO']==3) { 
     $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Calendário\',$w_pagina.\'calendario\',$w_Imagem,$w_Imagem,\'content\', null));');
     $SQL = "select count(*) as qtd from sbpi.cliente_site a inner join sbpi.modelo b on (a.sq_modelo = b.sq_modelo and upper(b.ds_diretorio) <> 'MOD13') where a.sq_cliente = ". $CL  ;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows); 

     foreach($RS as $row){  $RS = $row ; break; }               
       if(f($RS,'qtd') > 0){          
         $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Mensagens\',$w_pagina.\'mensagens\',$w_Imagem,$w_Imagem,\'content\', null));');         
     }         
   }   
   $i++; eval('$node'.$i.' = .$root->addItem(new XNode(\'Sair do sistema\',$w_pagina.\'Sair\',$w_Imagem,$w_Imagem,\'_top\', \'onClick="return(confirm(\\\'Confirma saída do sistema?\\\'));"\' ));');

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
  ShowHTML('<html>');
  ShowHTML('<head>');
  ShowHTML('  <!-- CSS FILE for my tree-view menu -->');
  ShowHTML('  <link rel="stylesheet" type="text/css" href="classes/menu/xPandMenu.css">');
  ShowHTML('  <!-- JS FILE for my tree-view menu -->');
  ShowHTML('  <script src="classes/menu/xPandMenu.js"></script>');
  ShowHTML('</head>');
  ShowHTML('<BASEFONT FACE="Verdana, Helvetica, Sans-Serif" SIZE="2">');
  // Decide se montará o body do menu principal ou o body do sub-menu de uma opção a partir do valor de w_sq_pagina
  echo('<BODY topmargin=0 bgcolor="#FFFFFF" BACKGROUND="img/background.gif" BGPROPERTIES="FIXED" text="#000000" link="#000000" vlink="#000000" alink="#FF0000" ');
  ShowHTML('onLoad="javascript:top.content.location=\''.$w_pagina.'escolas\';"> ');
  ShowHTML('  <CENTER><table border=0 cellpadding=0 height="80" width="100%">');
  ShowHTML('      <tr><td colspan=2 width="100%"><table border=0 width="100%" cellpadding=0 cellspacing=0><tr valign="top">');
  ShowHTML('          <td>Usuário:<b>'.$_SESSION['USERNAME'].'</b></TD>');
  ShowHTML('          <td align="right"><a class="hl" href="help.php?par=Menu.TP=<img src=images/Folder/hlp.gif border=0> SIW - Visão Geral.SG=MESA.O=L" target="content" title="Exibe informações sobre os módulos do sistema."><img src="images/Folder/hlp.gif" border=0></a></TD>');
  ShowHTML('          </table>');
  ShowHTML('      <tr><td height=1><tr><td height=2 bgcolor="#000000">');
  ShowHTML('  </table></CENTER>');
  ShowHTML('  <table border=0 cellpadding=0 height="80" width="100%"><tr><td nowrap><b>');
  ShowHTML('  <div id="container">');
  echo $menu_html_code;
  ShowHTML('  </div>');
  ShowHTML('  </table>');
  ShowHTML('</body>');
  ShowHTML('</html>');
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
  ShowHTML('  top.location.href=\''.$w_pagina.'\';');

  ScriptClose();
}

// =========================================================================
// Rotina Gravação
// -------------------------------------------------------------------------
function Grava(){
extract($GLOBALS);
$SG             = strtoupper($_REQUEST['SG']); 

  switch ($SG) {
    Case 'GETSITE':   
		/*
	  $SQL = "   update sbpi.Cliente_Site set " .
             "   no_contato_internet    = '" . trim($_REQUEST["w_no_contato_internet"]) . "', " .
             "   ds_email_internet      = '" . trim($_REQUEST["w_ds_email_internet"]) . "', " .
             "   nr_fone_internet       = '" . trim($_REQUEST["w_nr_fone_internet"]) . "', " .
             "   ds_texto_abertura      = '" . trim($_REQUEST["w_ds_texto_abertura"]) . "', " .
             "   ds_institucional       = '" . trim($_REQUEST["w_ds_institucional"]) . "', " ;
		
       If ( $_REQUEST["w_nr_fax_internet"] > "" ){ $SQL .=  "  nr_fax_internet         = '" . trim($_REQUEST["w_nr_fax_internet"]) . "', " ;      }Else{ $SQL .=  "   nr_fax_internet        = null,  " ;}
       If ( $_REQUEST["w_ds_mensagem"] > "" )    { $SQL .=  "   ds_mensagem            = '" . trim($_REQUEST["w_ds_mensagem"]) . "' "    ;        }Else{ $SQL .=  "   ds_mensagem            = null  "  ;}
       $SQL .= "where sq_cliente = " . $CL;

       db_exec::getInstanceOf($dbms, $SQL, &$numRows); 
	*/


	break;
  }
}	


// =========================================================================
// Rotina principal
// -------------------------------------------------------------------------
function Main(){
  extract($GLOBALS);
  switch ($par) {
    Case 'MENU':                   Menu();            break;
    Case 'ADMINISTRATIVO':         Administrativo();  break;
    Case 'VALIDA':                 Valida();          break;
    Case 'FRAMES':                 Frames();          break;
    Case 'GETSITE':                GetSite();         break; 
    Case 'GRAVA':                  Grava();           break;
    Case 'ESCOLAS':                escolas();         break;
    Case 'fotos':                  fotos();           break;
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
      ShowHTML('<BASE HREF="'.$conRootSIW.'">');
      BodyOpen('onLoad=this.focus();');
      Estrutura_Topo_Limpo();
      Estrutura_Menu();
      Estrutura_Corpo_Abre();
      Estrutura_Texto_Abre();
      ShowHTML('<div align=center><center><br><br><br><br><br><br><br><br><br><br><img src="images/icone/underc.gif" align="center"> <b>Esta opção está sendo desenvolvida.</b><br><br><br><br><br><br><br><br><br><br></center></div>');
      Estrutura_Texto_Fecha();
      Estrutura_Fecha();
      Estrutura_Fecha();
      Estrutura_Fecha();
      Rodape();
      break;
  }
}
?>