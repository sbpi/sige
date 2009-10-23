<?php
  // Garante que a sessão será reinicializada.
  session_start();

  $w_dir_volta = '';
  $w_pagina = 'controle.php?par=';
  $w_Disabled = 'ENABLED';
  $w_dir = '';
  $w_troca = $_REQUEST['w_troca'];
  $SG = $_REQUEST['SG'];

  if (isset ($_SESSION['LOGON1'])) {
    echo '<SCRIPT LANGUAGE="JAVASCRIPT">';
    echo ' alert("Já existe outra sessão ativa!\nEncerre o sistema na outra janela do navegador ou aguarde alguns instantes.\nUSE SEMPRE A OPÇÃO \"SAIR DO SISTEMA\" para encerrar o uso da aplicação.");';
    echo ' history.back();';
    echo '</SCRIPT>';
    exit ();
  }

  // Define o banco de dados a ser utilizado
  $_SESSION['DBMS'] = 2;

  include_once ('constants.inc');
  include_once ('jscript.php');
  include_once ('funcoes.php');
  include_once ($w_dir_volta . 'classes/db/abreSessao.php');
  include_once ($w_dir_volta . 'classes/sp/db_exec.php');
  include_once ($w_dir_volta . 'funcoes/selecaoRegiaoAdm.php');
  include_once ($w_dir_volta . 'funcoes/selecaoRegionalEnsino.php');
  include_once ($w_dir_volta . 'funcoes/selecaoTipoInstituicao.php');

  $P3 = nvl($_REQUEST['P3'], 1);
  $P4 = nvl($_REQUEST['P4'], $conPageSize);
  $p_modalidade = explodeArray($_REQUEST['p_modalidade']);
  $p_campos = explodeArray($_REQUEST['p_campos']);

  // Abre conexão com o banco de dados
  if (isset ($_SESSION['DBMS'])) {
    $dbms = abreSessao :: getInstanceOf($_SESSION['DBMS']);
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
  $RS = null;
  $CL = $_SESSION['CL'];
  $par = substr(strtoupper($_REQUEST['par']), 0, 30);
  $O = substr(strtoupper($_REQUEST['O']), 0, 1);
  $R = $_REQUEST['R'];
  if (nvl($O, '') == '') {
    $O = 'L';
  }
  $w_Data = Substr(100 + Day(Time()), 1, 2) . "/" . Substr(100 + Month(Time()), 1, 2) . "/" . Year(Time());

  if (nvl($par, '') == '')
    Logon();
  else
    main();

  FechaSessao($dbms);

  exit;

  // =========================================================================
  // Rotina de criação da tela de logon (backup)
  // -------------------------------------------------------------------------
  function LogOn() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_username = $_REQUEST['Login'];
    ShowHTML('<HTML>');
    ShowHTML('<HEAD>');
    ShowHTML('<link rel="shortcut icon" href="' . $conRootSIW . 'favicon.ico" type="image/ico" />');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    ShowHTML('<TITLE>' . $conSgSistema . ' - Autenticação</TITLE>');
    ScriptOpen('JavaScript');
    ShowHTML('$(document).ready(function(){');
    ShowHTML('	$("#Login1").change(function(){');
    ShowHTML('		formataCampo();');
    ShowHTML('	})');
    ShowHTML('});');
    ShowHTML('function formataCampo(){');
    ShowHTML('	$("#Login1").val(trim($("#Login1").val()));');
    ShowHTML('	if(  $("#Login1").val().length==11 &&  caracterAceito( $("#Login1").val() ,  "0123456789") ){');
    ShowHTML('		$("#Login1").val( mascaraGlobal(\'###.###.###-##\',$("#Login1").val()) );');
    ShowHTML('	}');
    ShowHTML('}');
    ShowHTML('function caracterAceito(string , checkOK){');
    ShowHTML('	 //var checkOK = \'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-!@#$%&*()+/\';');
    ShowHTML('		  var checkStr = string;');
    ShowHTML('		  var allValid = true;');
    ShowHTML('		  for (i = 0;  i < checkStr.length;  i++)');
    ShowHTML('		  {');
    ShowHTML('			ch = checkStr.charAt(i);');
    ShowHTML('			if ((checkStr.charCodeAt(i) != 13) && (checkStr.charCodeAt(i) != 10) && (checkStr.charAt(i) != "\\\\")) {');
    ShowHTML('			   for (j = 0;  j < checkOK.length;  j++) {');
    ShowHTML('				 if (ch == checkOK.charAt(j))');
    ShowHTML('				   break;');
    ShowHTML('			   }');
    ShowHTML('			   if (j == checkOK.length)');
    ShowHTML('			   {');
    ShowHTML('				 allValid = false;');
    ShowHTML('				 break;');
    ShowHTML('			   }');
    ShowHTML('			}');
    ShowHTML('		  }');
    ShowHTML('		  return allValid;');
    ShowHTML('}');
    ShowHTML('function Ajuda() ');
    ShowHTML('{ ');
    ShowHTML('  document.Form.Botao.value = "Ajuda"; ');
    ShowHTML('} ');
    Modulo();
    SaltaCampo();
    ValidateOpen('Validacao');
    Validate('Login1', 'Nome de usuário', '', '1', '2', '30', '1', '1');
    Validate('Password1', 'Senha', '1', '1', '3', '19', '1', '1');
    ShowHTML('  theForm.Login.value = theForm.Login1.value; ');
    ShowHTML('  theForm.Password.value = theForm.Password1.value; ');
    ShowHTML('  theForm.Login1.value = ""; ');
    ShowHTML('  theForm.Password1.value = ""; ');
    ValidateClose();
    ScriptClose();
    ShowHTML('<link rel="stylesheet" type="text/css" href="' . $conRootSIW . 'classes/menu/xPandMenu.css">');
    ShowHTML('<style>');
    ShowHTML(' .cText {font-size: 8pt; border: 1px solid #000000; background-color: #F5F5F5}');
    ShowHTML(' .cButton {font-size: 8pt; color: #FFFFFF; border: 1px solid #000000; background-color: #669966; }');
    ShowHTML('</style>');
    ShowHTML('</HEAD>');
    // Se receber a username, dá foco na senha
    if (nvl($w_username, 'nulo') == 'nulo') {
      ShowHTML('<body topmargin=0 leftmargin=10 onLoad=\'document.Form.Login1.focus();\'>');
    } else {
      ShowHTML('<body topmargin=0 leftmargin=10 onLoad=\'document.Form.Password1.focus();\'>');
    }
    ShowHTML('<CENTER>');
    ShowHTML('<form method="post" action="controle.php?par=valida" onsubmit="return(Validacao(this));" name="Form"> ');
    ShowHTML('<INPUT TYPE="HIDDEN" NAME="Login" VALUE=""> ');
    ShowHTML('<INPUT TYPE="HIDDEN" NAME="Password" VALUE=""> ');
    ShowHTML('<INPUT TYPE="HIDDEN" NAME="p_dbms" VALUE="1"> ');
    ShowHTML('<TABLE cellSpacing=0 cellPadding=0 width="760" height=550 border=1  background="files/" & p_cliente & "/img/fundo.jpg" bgproperties="fixed"><tr><td width="100%" valign="top">');
    ShowHTML('  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 background="img/cabecalho.gif">');
    ShowHTML('    <TR>');
    ShowHTML('      <TD vAlign=TOP align=middle width="65%">');
    ShowHTML('        <B><FONT face=Arial size=5 color=#000088>Secretaria de Estado de Educação');
    ShowHTML('        <br><br>SIGE - Módulo Web</B></FONT>');
    ShowHTML('     </TD>');
    ShowHTML('    </TR>');
    ShowHTML('    <TR><TD colspan=3 borderColor=#ffffff height=22><HR align=center color=#808080></TD></TR>');
    ShowHTML('  </TABLE>');
    ShowHTML('  <table width="100%" border="0">');
    ShowHTML('    <tr><td valign="middle" width="100%" height="100%">');
    ShowHTML('        <table width="100%" height="100%" border="0">');
    ShowHTML('          <tr><td align="center" colspan=2><font size="1" color="#990000"><b>Esta aplicação é de uso interno da Secretaria de Estado de Educação.<br>As informações contidas nesta aplicação são restritas e de uso exclusivo.<br>O uso indevido acarretará ao infrator penalidades de acordo com a legislação em vigor.<br><br>Informe seu nome de usuário, senha de acesso e clique no botão <i>OK</i> para ser autenticado pela aplicação.</b></font>');
    ShowHTML('          <tr><td align="right" width="43%"><font size="2"><b>Nome de usuário:<td><input class="sti" name="Login1" size="14" maxlength="14">');
    ShowHTML('          <tr><td align="right"><font size="2"><b>Senha:<td><input class="sti" type="Password" name="Password1" size="19">');
    ShowHTML('          <tr><td align="right"><td><font size="2"><b><input class="stb" type="submit" value="OK" name="Botao"> ');
    ShowHTML('          </font></td> </tr> ');
    ShowHTML('          <TR><TD colspan=2 align="center"><br><table border=0 cellpadding=0 cellspacing=0><tr><td>');
    ShowHTML('              <P><IMG height=37 src="img/ajuda.jpg" width=629><br>');
    ShowHTML('              <font face="Arial" size=1><b>PARA ACESSAR A PÁGINA DE ATUALIZAÇÃO</b></font>');
    ShowHTML('              <FONT face="Verdana, Arial, Helvetica, sans-serif" size=1>');
    ShowHTML('              <li>Nome de usuário - Informe seu nome de usuário');
    ShowHTML('              <li>Senha - Informe sua senha de acesso');
    ShowHTML('              <li>Se esqueceu ou não foi informado dos dados acima, favor entrar em contato com a SEDF / SUBIP / Diretoria de Sistemas de Informação Educacional - DSIE');
    ShowHTML('              </FONT></P>');
    ShowHTML('              <P><font face="Arial" size=1><b>DOCUMENTAÇÃO - LEIA COM ATENÇÃO</b></font><br>');
    ShowHTML('              <FONT face="Verdana" size=1>');
    ShowHTML('              . <a class="SS" href="sedf/Orientacoes_Acesso.pdf" target="_blank" title="Abre arquivo que descreve as novas características e funcionalidades do SIGE-WEB.">Apresentação da nova versão do SIGE-WEB (PDF - 130KB - 4 páginas)</a><BR>');
    ShowHTML('              . <a class="SS" href="manuais/operacao/" target="_blank" title="Exibe manual de operação do SIGE-WEB">Manual SIGE-WEB (HTML)</A><BR>');
    ShowHTML('              <br></FONT></P>');
    ShowHTML('              </TD></TR>');
    ShowHTML('          </table> ');
    ShowHTML('        </table> ');
    ShowHTML('    </tr> ');
    ShowHTML('  </table>');
    ShowHTML('</table>');
    ShowHTML('</form> ');
    ShowHTML('</CENTER>');
    ShowHTML('</body>');
    ShowHTML('</html>');
  }

  // =========================================================================
  // Rotina de autenticação dos usuários
  // -------------------------------------------------------------------------
  function Valida() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_uid = str_replace('""', "", str_replace("'", "", Trim(strtoupper($_REQUEST["Login"]))));
    $w_pwd = str_replace('""', "", str_replace("'", "", Trim($_REQUEST["Password"])));
    $w_erro = 0;
    $SQL = "select count(*) existe from sbpi.Cliente where upper(ds_username) = upper('" . $w_uid . "')";
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    foreach ($RS as $row) {
      $RS = $row;
      break;
    }

    if (f($RS, "existe") == 0) {
      $w_erro = 1;
    } else {
      $SQL = "select count(*) existe " . $crlf .
             "  from sbpi.Cliente           a " . $crlf .
             "       join sbpi.Tipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente and " . $crlf .
             "                                  b.tipo            = 1 " . $crlf .
             "                                 ) " . $crlf .
             " where upper(ds_username) = upper('" . $w_uid . "')";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      if (f($RS, "existe") == 0) {
        $w_erro = 3;
      } else {
        $SQL = "select count(*) existe from sbpi.Cliente where upper(ds_username) = upper('" . $w_uid . "') and upper(ds_senha_acesso) = upper('" . $w_pwd . "')";

        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        foreach ($RS as $row) {
          $RS = $row;
          break;
        }
        if (f($RS, "existe") == 0) {
          $w_erro = 2;
        }
      }
    }
    ScriptOpen('JavaScript');
    if ($w_erro > 0) {
      if ($w_erro == 1)
        ShowHTML(' alert("Usuário inexistente!");');
      elseif ($w_erro == 2) ShowHTML(' alert("Senha inválida!");');
      else
        ShowHTML(' alert("Usuário sem permissão para acessar esta página!");');
      ShowHTML('  history.back(1);');
    } else {
      //Recupera informações a serem usadas na montagem das telas para o usuário
      $SQL = "select ds_username, sq_cliente from sbpi.Cliente where upper(ds_username) = upper('" . $w_uid . "') and ds_senha_acesso = '" . $w_pwd . "'";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $_SESSION['USERNAME'] = strtoupper(f($RS, 'ds_username'));
      $_SESSION['CL'] = f($RS, 'sq_cliente');
      If ($_SESSION['USERNAME'] != 'SBPI') {
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
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      }
      ShowHTML(' location.href="controle.php?par=frames";');
    }
    ScriptClose();
  }

  // =========================================================================
  // Exportação dos dados administrativos
  // -------------------------------------------------------------------------
  function administrativo() {
    extract($GLOBALS);
    global $w_Disabled;

    Cabecalho();
    ShowHTML('<HEAD>');
    ScriptOpen('Javascript');
    ValidateOpen('Validacao');
    ShowHTML('  if (theForm.w_arquivo[0].checked == false && theForm.w_arquivo[1].checked == false) {');
    ShowHTML('     alert(\'Você deve escolher uma das opções apresentadas antes de gerar o arquivo!\');');
    ShowHTML('     return false;');
    ShowHTML('  }');
    ShowHTML('  return(confirm(\'Confirma a geração do arquivo com os dados indicados?\'));');
    ValidateClose();
    ScriptClose();
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    ShowHTML('</HEAD>');
    BodyOpen('onLoad=\'document.focus();\'');
    ShowHTML('<B><FONT COLOR="#000000">Exportação dos dados administrativos</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
    AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
    ShowHTML(MontaFiltro('POST'));
    ShowHTML('<input type="hidden" name="SG" value="ADM">');
    ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
    ShowHTML('    <table width="97%" border="0">');
    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><b>Exportação dos dados administrativos</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1><ul><li>Esta tela permite a exportação, para um arquivo que pode ser aberto no Excel, dos dados administrativos preenchidos pelas unidades de ensino.<li>Permite também exportar as tabelas de apoio utilizadas pelo formulário.<li>Selecione uma das opções exibidas abaixo e clique no botão "Gerar arquivo" para que os dados sejam convertidos para um arquivo.</ul></font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top"><b>O arquivo a ser gerado deve conter dados:</b>');
    ShowHTML('          <br><INPUT ' . $w_Disabled . ' class="BTM" type="radio" name="w_arquivo" value="Escola"> das unidades de ensino');
    ShowHTML('          <br><INPUT ' . $w_Disabled . ' class="BTM" type="radio" name="w_arquivo" value="Tipo"> da tabela de equipamentos');
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
  function frames() {
    extract($GLOBALS);
    $_SESSION["BodyWidth"] = "620";

    ShowHTML('<html>');
    ShowHTML('<head>');
    ShowHTML('    <title>Controle Central</title>');
    ShowHTML('</head>');

    ShowHTML('<frameset cols="200,*">');
    ShowHTML('    <frame name="menu" src="controle.php?par=menu" scrolling="auto" marginheight="0" marginwidth="0">');
    ShowHTML('    <frame name="content" src="controle.php?par=escolas" scrolling="auto" marginheight="0" marginwidth="0">');
    ShowHTML('</frameset>');
    ShowHTML('</html>');
  }

  // =========================================================================
  // Cadastro de modalidades de ensino
  // -------------------------------------------------------------------------
  function Modalidades() {

    extract($GLOBALS);
    global $w_Disabled;

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];

    $SQL = "Select nr_ordem, ds_especialidade from sbpi.Especialidade order by nr_ordem, ds_especialidade";
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    foreach ($RS as $row) {
      $RS = $row;
      break;
    }
    if (count($RS) > 0) {
      $w_texto = '<b>Nºs de ordem em uso para esta subordinação:</b>:<br>' .
      '<table border=1 width=100% cellpadding=0 cellspacing=0>' .
      '<tr><td align=center><b><font size=1>Ordem' .
      '    <td><b><font size=1>Descrição';
      foreach ($RS as $row) {
        $w_texto .= '<tr><td valign=top align=center><font size=1>' . f($row, 'nr_ordem') . '<td valign=top><font size=1>' . f($row, 'ds_especialidade');
      }
      $w_texto .= "</table>";
    } else {
      $w_texto = "Não há outros números de ordem vinculados à subordinação desta opção";
    }

    If ($w_troca > '') { // Se for recarga da página
      $w_ds_especialidade = $_REQUEST['w_ds_especialidade'];
      $w_nr_ordem = $_REQUEST['w_nr_ordem'];
    }
    elseif ($O == "L") {
      //Recupera todos os registros para a listagem
      $SQL = "select nr_ordem, ds_especialidade, sq_especialidade  from sbpi.Especialidade order by nr_ordem, ds_especialidade";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }
    elseif (strpos("AEV", $O) !== false and $w_troca == '') {
      //Recupera os dados do endereço informado
      $SQL = "select nr_ordem, ds_especialidade, sq_especialidade from sbpi.Especialidade where sq_especialidade = " . $w_chave;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_ds_especialidade = f($RS, "ds_especialidade");
      $w_nr_ordem = f($RS, "nr_ordem");
    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    if (strpos("IAEP", $O) !== false) {
      ScriptOpen('JavaScript');
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      if (strpos("IA", $O) !== false) {
        Validate("w_ds_especialidade", "Descrição", "", "1", "2", "70", "1", "1");
        Validate("w_nr_ordem", "Nr. de ordem", "", "1", "1", "4", "", "0123546789");
      }
      ShowHTML(' theForm.Botao[0].disabled=true;');
      ShowHTML('  theForm.Botao[1].disabled=true;');
      ValidateClose();
      ScriptClose();
    }

    ShowHTML('</HEAD>');
    if ($w_troca > "") {
      BodyOpen('onLoad=\'document.Form.' . $w_troca . '.focus()\';');
    }
    elseif ($O == "I" or $O == "A") {
      BodyOpen("onLoad='document.Form.w_ds_especialidade.focus()';");
    } else {
      BodyOpen("onLoad='document.focus()';");
    }
    ShowHTML('<B><FONT COLOR="#000000">Cadastro de modalidades de ensino</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    If ($O == 'L') {
      // Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $dir . $w_pagina . $par . $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><b>Ordem</font></td>');
      ShowHTML('          <td><b>Modalidade</font></td>');
      ShowHTML('          <td><b>Operações</font></td>');
      ShowHTML('        </tr>');
      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          if (Nvl(f($row, "nr_ordem"), "nulo") <> "nulo") {
            ShowHTML('        <td align="CENTER">' . f($row, "nr_ordem") . '</td>');
          } else {
            ShowHTML('        <td align="center">---</td>');
          }
          ShowHTML('        <td>' . f($row, "ds_especialidade") . '</td>');
          ShowHTML('        <td align="top" nowrap>');
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . 'modalidades' . $w_ew . "&R=" . $w_pagina . 'modalidades' . $w_ew . "&O=A&CL=" . $CL . "&w_chave=" . f($row, "sq_especialidade") . '">Alterar</A>&nbsp');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=E&w_chave=' . f($row, 'sq_especialidade') . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
        }
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif (strpos('IAEV', $O) !== false) {
      If (strpos('EV', $O) !== false) {
        $w_Disabled = ' DISABLED ';
      }
      AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
      ShowHTML('<INPUT type="hidden" name="SG" value="MODALIDADES">');
      ShowHTML('<INPUT type="hidden" name="CL" value="' . $CL . '">');
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML('<INPUT type="hidden" name="w_sq_cliente" value="' . str_replace('sq_cliente=', $CL, 'sq_cliente=') . '">');
      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');

      ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
      ShowHTML('        <tr valign="top"><td valign="top"><b>D<u>e</u>scrição:</b><br><input ' . $w_Disabled . ' accesskey="E" type="text" name="w_ds_especialidade" class="STI" SIZE="70" MAXLENGTH="70" VALUE="' . $w_ds_especialidade . '"></td>');
      ShowHTML('              <td valign="top" align="left"><b><u>O</u>rdem:<br><INPUT ' . $w_Disabled . ' ACCESSKEY="O" TYPE="TEXT" CLASS="STI" NAME="w_nr_ordem" SIZE=4 MAXLENGTH=4 VALUE="' . $w_nr_ordem . '" " . $w_Disabled . "></td>');
      ShowHTML('        </table>');
      ShowHTML('      <tr>');
      ShowHTML('      <tr><td align="center" colspan=4><hr>');
      If ($O == "E") {
        ShowHTML('   <input class="STB" type="submit" name="Botao" value="Excluir">');
      } else {
        If ($O == "I") {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Incluir">');
        } else {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }
      ShowHTML('            <input class="STB" type="button" onClick="location.href=\'' . $dir . $w_pagina . $par . $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
      ShowHTML('          </td>');
      ShowHTML('      </tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');
      ShowHTML('</FORM>');
    } else {
      ScriptOpen('JavaScript');
      ShowHTML(' alert(\'Opção não disponível\');');
      ShowHTML(' history.back(1);');
      ScriptClose();
    }
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();

  }
  // =========================================================================
  // Fim do cadastro de modalidades de ensino
  // -------------------------------------------------------------------------

  // =========================================================================
  // Cadastro de calendario base
  // -------------------------------------------------------------------------
  function calend_base() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];

    if ($w_troca > "") { // Se for recarga da página
      $w_dt_ocorrencia = $_REQUEST["w_dt_ocorrencia"];
      $w_ds_ocorrencia = $_REQUEST["w_ds_ocorrencia"];
      $w_tipo = $_REQUEST["w_tipo"];
    }
    elseif ($O == "L") {
      //Recupera todos os registros para a listagem
      $SQL = 'select a.*, b.nome from sbpi.calendario_base a left join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) order by sbpi.year(dt_ocorrencia) desc, dt_ocorrencia';
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }
    elseif (strpos("AEV", $O) !== false and $w_troca == "") {
      //Recupera os dados do endereço informado
      $SQL = "select * from sbpi.calendario_base where sq_ocorrencia = " . $w_chave;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_dt_ocorrencia = FormataDataEdicao(f($row, "dt_ocorrencia"));
      $w_ds_ocorrencia = f($row, "ds_ocorrencia");
      $w_tipo = f($row, "sq_tipo_data");
    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('   <link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    if (strpos("IAEP", $O) !== false) {
      ScriptOpen('JavaScript');
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      if (strpos("IA", $O) !== false) {
        Validate("w_dt_ocorrencia", "Data", "DATA", "1", "10", "10", "1", "1");
        Validate("w_ds_ocorrencia", "Descrição", "", "1", "2", "60", "1", "1");
        Validate("w_tipo", "Tipo", "SELECT", "1", "1", "4", "", "1");
      }
      ShowHTML('  theForm.Botao[0].disabled=true;');
      ShowHTML('  theForm.Botao[1].disabled=true;');
      ValidateClose();
      ScriptClose();
    }
    ShowHTML('</HEAD>');
    if ($w_troca > "") {
      BodyOpen('onLoad=\'document.Form.' . $w_troca . '.focus()\';');
    }
    elseif ($O == "I" or $O == "A") {
      BodyOpen('onLoad="document.Form.w_dt_ocorrencia.focus()";');
    } else {
      BodyOpen('onLoad="document.focus()";');
    }
    ShowHTML('<B><FONT COLOR=""#000000"">Cadastro do calendário oficial</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    if ($O == "L") {
      //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $dir . $w_pagina . $par . $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><b>Data</font></td>');
      ShowHTML('          <td><b>Tipo</font></td>');
      ShowHTML('          <td><b>Ocorrência</font></td>');
      ShowHTML('          <td><b>Operações</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          if ($wAno != year(f($row, "dt_ocorrencia"))) {
            ShowHTML('      <tr bgcolor=""#C0C0C0"" valign=""top""><TD colspan=4 align="center"><font size=2><B>' . year(f($row, "dt_ocorrencia")) . '</b></font></td></tr>');
            $wAno = year(f($row, "dt_ocorrencia"));
          }
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td align="center">' . Substr(FormataDataEdicao(FormatDateTime(f($row, "dt_ocorrencia"), 2)), 0, 5) . '</td>');
          ShowHTML('        <td>' . nvl(f($row, "nome"), "---") . '</td>');
          ShowHTML('        <td>' . f($row, "ds_ocorrencia") . '</td>');
          ShowHTML('        <td align="top" nowrap>');
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . 'calend_base' . $w_ew . '&R=' . $w_pagina . 'calend_base' . $w_ew . '&O=A&CL=' . $CL . '&w_chave=' . f($row, "sq_ocorrencia") . '&SG=' . $SG . '">Alterar</A>');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=E&w_chave=' . f($row, 'sq_ocorrencia') . '&SG=' . $SG . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
        }
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif (strpos("IAEV", $O) !== false) {
      if (strpos("EV", $O) !== false) {
        $w_Disabled = ' DISABLED ';
      }
      AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
      ShowHTML('<input type="hidden" name="SG" value="CALEND_BASE">');
      ShowHTML('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
      ShowHTML('<INPUT type="hidden" name="CL" value="' . $CL . '">');
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');

      ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td valign="top"><b><u>D</u>ata:</b><br><input ' . $w_Disabled . ' accesskey="D" type="text" name="w_dt_ocorrencia" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_ocorrencia, Time()), 2)) . '" onKeyDown="FormataData(this,event);"></td>');
      ShowHTML('          <td valign="top"><b>D<u>e</u>scrição:</b><br><input ' . $w_Disabled . ' accesskey="E" type="text" name="w_ds_ocorrencia" class="STI" SIZE="60" MAXLENGTH="60" VALUE="' . $w_ds_ocorrencia . '"></td>');
      $SQL = 'SELECT * FROM sbpi.Tipo_Data a WHERE a.abrangencia <> \'U\' ORDER BY a.nome' . $crlf;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      ShowHTML('          <td><b>Tipo da ocorrência:</b><br><SELECT CLASS="STI" NAME="w_tipo" ' . $w_Disabled . '>');
      ShowHTML('          <option value=""> ---');
      foreach ($RS as $row) {
        if (doubleval(nvl(f($row, "sq_tipo_data"), 0)) == doubleval(nvl($w_tipo, 0))) {
          ShowHTML('          <option value=' . f($row, "sq_tipo_data") . ' SELECTED>' . f($row, "nome"));
        } else {
          ShowHTML('          <option value=' . f($row, "sq_tipo_data") . '>' . f($row, "nome"));
        }
      }
      ShowHTML('          </select>');
      ShowHTML('        </table>');
      ShowHTML('      <tr>');
      ShowHTML('      <tr><td align="center" colspan=4><hr>');
      if ($O == "E") {
        ShowHTML('   <input class="STB" type="submit" name="Botao" value="Excluir">');
      } else {
        if ($O == "I") {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Incluir">');
        } else {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }
      ShowHTML('            <input class="STB" type="button" onClick="location.href=\'' . $dir . $w_pagina . $par . $w_ew . "&CL=" . $CL . '&O=L&SG=\';" name="Botao" value="Cancelar">');
      ShowHTML('          </td>');
      ShowHTML('      </tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');
      ShowHTML('</FORM>');
    } else {
      ScriptOpen('JavaScript');
      ShowHTML(' alert(\'Opção não disponível\');');
      ShowHTML(' history.back(1);');
      ScriptClose();
    }
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();

  }
  // =========================================================================
  // Fim do cadastro de calendário base
  // -------------------------------------------------------------------------

  // =========================================================================
  // Cadastro de calendario
  // -------------------------------------------------------------------------
  function calend_rede() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];

    if ($w_troca > "") { //Se for recarga da página
      $w_dt_ocorrencia = $_REQUEST["w_dt_ocorrencia"];
      $w_ds_ocorrencia = $_REQUEST["w_ds_ocorrencia"];
      $w_tipo = $_REQUEST["w_tipo"];
    }
    elseif ($O == "L") {
      //Recupera todos os registros para a listagem
      $SQL = "select a.sq_ocorrencia as chave, a.ds_ocorrencia, a.dt_ocorrencia, a.sq_tipo_data, b.nome from sbpi.Calendario_Cliente a left join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) where sq_cliente = " . $CL . " order by sbpi.year(dt_ocorrencia) desc, dt_ocorrencia";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }
    elseif (strpos("AEV", $O) !== false and $w_troca == '') {
      //Recupera os dados do endereço informado
      $SQL = "select dt_ocorrencia, ds_ocorrencia, sq_tipo_data from sbpi.Calendario_Cliente where sq_ocorrencia = " . $w_chave;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_dt_ocorrencia = FormataDataEdicao(f($row, "dt_ocorrencia"));
      $w_ds_ocorrencia = f($row, "ds_ocorrencia");
      $w_tipo = f($row, "sq_tipo_data");
    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('   <link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');

    if (strpos("IAEP", $O) !== false) {
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      if (strpos("IA", $O) !== false) {
        Validate("w_dt_ocorrencia", "Data", "DATA", "1", "10", "10", "1", "1");
        Validate("w_ds_ocorrencia", "Descrição", "", "1", "2", "60", "1", "1");
        Validate("w_tipo", "Tipo", "SELECT", "1", "1", "4", "", "1");
      }
      ShowHTML('  theForm.Botao[0].disabled=true;');
      ShowHTML('  theForm.Botao[1].disabled=true;');
      ValidateClose();
      ScriptClose();
    }
    ShowHTML('</HEAD>');
    if ($w_troca > "") {
      BodyOpen('onLoad=\'document.Form.' . $w_troca . '.focus()\';');
    }
    elseif ($O == "I" or $O == "A") {
      BodyOpen('onLoad=\'document.Form.w_dt_ocorrencia.focus()\';');
    } else {
      BodyOpen('onLoad=\'document.focus()\';');
    }
    ShowHTML('<B><FONT COLOR="#000000">Cadastro do calendário da rede de ensino</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    if ($O == "L") {
      //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $w_pagina . $par . '&O=I&CL=' . $CL . '"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><b>Data</font></td>');
      ShowHTML('          <td><b>Tipo</font></td>');
      ShowHTML('          <td><b>Ocorrência</font></td>');
      ShowHTML('          <td><b>Operações</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          if ($wAno != year(f($row, "dt_ocorrencia"))) {
            ShowHTML('      <tr bgcolor="#C0C0C0" valign="top"><TD colspan=4 align="center"><font size=2><B>' . year(f($row, "dt_ocorrencia")) . '</b></font></td></tr>');
            $wAno = year(f($row, "dt_ocorrencia"));
          }
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td align="center">' . substr(FormataDataEdicao(FormatDateTime(f($row, "dt_ocorrencia"), 2)), 0, 5) . '</td>');
          ShowHTML('        <td>' . nvl(f($row, "nome"), "---") . '</td>');
          ShowHTML('        <td>' . f($row, "ds_ocorrencia") . '</td>');
          ShowHTML('        <td align="top" nowrap>');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=A&w_chave=' . f($row, 'chave') . '&SG=' . $SG . '">Alterar</A>&nbsp');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=E&w_chave=' . f($row, 'chave') . '&SG=' . $SG . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
        }
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif ((strpos('IAEV', $O) !== false)) {
      if ((strpos('EV', $O) !== false))
        $w_Disabled = ' DISABLED ';
      AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML('<INPUT type="hidden" name="w_troca" value="">');
      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');
      ShowHTML('<INPUT type="hidden" name="SG" value="CALEND_REDE">');
      ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
      ShowHTML('    <table width="97%" border="0">');
      ShowHTML('      <tr><td><table border=0 width="100%"><tr valign="top">');
      ShowHTML('           <td colspan=2><b><u>D</u>ata:</b><br><input ' . $w_Disabled . ' accesskey="D" type="text" name="w_dt_ocorrencia" class="sti" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_ocorrencia, Time()), 2)) . '"  onKeyDown="FormataData(this,event);"></td>');
      ShowHTML('      <tr><td valign="top" colspan="2"><table border="0" width="100%" cellspacing=0>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('           <td colspan=2><b><u>D</u>escrição:</b><br><input ' . $w_Disabled . ' accesskey="E" type="text" name="w_ds_ocorrencia" class="sti" SIZE="60" MAXLENGTH="60" VALUE="' . $w_ds_ocorrencia . '"></td>');
      $SQL = 'SELECT * FROM sbpi.Tipo_Data a WHERE a.abrangencia <> \'U\' ORDER BY a.nome' . $crlf;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      ShowHTML('          <td><b>Tipo da ocorrência:</b><br><SELECT CLASS="STI" NAME="w_tipo" ' . $w_Disabled . '>');
      ShowHTML('          <option value=""> ---');
      foreach ($RS as $row) {
        if (doubleval(nvl(f($row, "sq_tipo_data"), 0)) == doubleval(nvl($w_tipo, 0))) {
          ShowHTML('          <option value="' . f($row, "sq_tipo_data") . '" SELECTED>' . f($row, "nome"));
        } else {
          ShowHTML('          <option value="' . f($row, "sq_tipo_data") . '">' . f($row, "nome"));
        }
      }
      ShowHTML('          </select>');
      ShowHTML('        </table>');
      ShowHTML('      <tr>');
      ShowHTML('      <tr><td align="center" colspan=4><hr>');
      if ($O == "E") {
        ShowHTML('   <input class="STB" type="submit" name="Botao" value="Excluir">');
      } else {
        if ($O == "I") {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Incluir">');
        } else {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }

      ShowHTML('            <input class="STB" type="button" onClick="location.href=\'' . $dir . $w_pagina . $par . $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
      ShowHTML('          </td>');
      ShowHTML('      </tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');
      ShowHTML('</FORM>');
    } else {
      ScriptOpen("JavaScript");
      ShowHTML(' alert(\'Opção não disponível\');');
      ShowHTML(' history.back(1);');
      ScriptClose();
    }
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();
  }
  // =========================================================================
  // Fim do cadastro de calendário
  // -------------------------------------------------------------------------

  // =========================================================================
  // Monta a tela de Homologação do Calendário das Escolas Particulares
  // -------------------------------------------------------------------------
  function ligacao() {
    extract($GLOBALS);
    global $w_Disabled;
    $p_regiao = $_REQUEST["p_regiao"];
    $w_homologado = $_REQUEST["w_homologado"];
    if ($w_homologado != "S") {
      $w_homologado = 'Não';
    } else {
      $w_homologado = 'Sim';
    }

    if ($p_tipo == 'W') {
      //Response.ContentType = "application/msword"
      //HeaderWord p_layout
      ShowHTML('<TABLE WIDTH="100%" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR="#000000">SIGE-WEB<TD ALIGN="RIGHT"><B><FONT SIZE=4 COLOR="#000000">');
      ShowHTML('Consulta a escolas');
      ShowHTML('</FONT><TR><TD ALIGN="RIGHT"><B><FONT SIZE=2 COLOR="#000000">" . DataHora(); . "</B></TD></TR>');
      ShowHTML('</FONT></B></TD></TR></TABLE>');
      ShowHTML('<HR>');
    } else {
      Cabecalho();
      ShowHTML('<HEAD>');
      ShowHTML('   <link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
      ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
      ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
      ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
      ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
      ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
      ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
      ShowHTML('</HEAD>');
      if ($_REQUEST["pesquisa"] > '') {
        BodyOpen('onLoad="location.href=\'#lista\'"');
        //} else {
        //BodyOpen "onLoad='document.Form.p_regional.focus()';"
      }
    }
    ShowHTML('<B><FONT COLOR="#000000">' . $w_tp . '</FONT></B>');
    ShowHTML('<B><FONT size="2" COLOR="#000000">Vinculação e tipologia</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="100%">');
    ShowHTML('<tr bgcolor="" & conTrBgColor & ""><td align="center">');
    ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
    ShowHTML('    <table width="90%" cellspacing=0>');
    AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
    ShowHTML('<INPUT type="hidden" name="R" value="' . $w_ew . '">');
    ShowHTML('<INPUT type="hidden" name="O" value="' . $w_ea . '">');
    ShowHTML('<INPUT type="hidden" name="w_troca" value="">');
    ShowHTML('<INPUT type="hidden" name="CL" value="' . $CL . '">');
    ShowHTML('      <tr><td colspan=2><table border=0 width="90%" cellspacing=0>');
    ShowHTML('        <tr valign="top">');
    SelecaoEscolaParticular('Unidad<u>e</u> de ensino:', 'E', 'Selecione unidade.', $p_escola_particular, null, "p_escola_particular", null, "onChange='document.Form.action=\'" . $w_pagina . $par . "'; document.Form.O.value='P'; document.Form.w_troca.value='p_calendario'; document.Form.submit();");
    //SelecaoEscolaParticular         ("Unidad<u>e</u> de ensino:", "E", "Selecione unidade." , $p_escola_particular, null, "p_escola_particular", null, "onChange="document.Form.action='" & w_pagina & w_ew & "'; document.Form.w_ea.value='P'; document.Form.w_troca.value='p_calendario'; document.Form.submit();"");
    //ShowHTML ('        <tr valign="top">');
    //SelecaoCalendarioParticular         "<u>C</u>alendário:", "E", "Selecione unidade." , p_calendario, p_escola_particular, "p_calendario", null, "onChange="document.Form.action='" & w_pagina & w_ew & "'; document.Form.w_ea.value='P'; document.Form.w_troca.value='w_homologado'; document.Form.submit();""
    //ShowHTML ('        <tr valign="top">"
    //SelecaoRegionalEscola "Regional de en<u>s</u>ino:", "S", "Indique a regional de ensino.", p_regional, p_escola_particular, "p_regional", null, null
    //ShowHTML ('        <tr valign="top">"
    //SelecaoTipoEscola     "<u>T</u>ipo de Escola:", "T", "Selecione o tipo da Escola.", p_tipo_escola, p_escola_particular, "p_tipo_escola", null, null
    if ($p_escola_particular > '') {
      $SQL = 'SELECT sq_particular_calendario, sq_cliente as cliente, nome, homologado FROM sbpi.Particular_Calendario WHERE sq_cliente = ' . $p_escola_particular;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

      ShowHTML('<font color="#ff3737"><strong><a href="javascript:this.status.value;" onClick="window.open(\'calendario.php?par=formcal&CL=' . $RS[0]["cliente"] . '&O=L&w_ew=formcal&w_ee=1&controle=s\',\'MetaWord\',\'width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no\');">Acessar o(s) calendário(s)</a></strong></font>');

      ShowHTML('<table id="tbhomologacao" border="1">');
      ShowHTML('<tr><td>Título do Calendário</td><td>Homologado?</td></tr>');

      foreach ($RS as $row) {
        $homologado = f($row, "homologado");
        ShowHTML('<INPUT type="hidden" name="w_cliente" value="' . RS("cliente") . '">');
        ShowHTML('<tr><td>' . $RS ("nome") . '</td>');
        ShowHTML('<INPUT type="hidden" name="w_chave" value="' . RS("sq_particular_calendario") . '">');
        ShowHTML('<td><select name="w_homologado">');
        if (strpos(strtoupper($homologado), 'N')) {
          ShowHTML('<option value="N" SELECTED>Não');
          ShowHTML('<option value="S">Sim');
        }
        elseif ((strpos(strtoupper($homologado), "S"))) {
          ShowHTML('<option value="S" SELECTED>Sim');
          ShowHTML('<option value="N">Não');
        }
        ShowHTML('</select></td>');
        ShowHTML('<td align="center" colspan="2">');
        //ShowHTML ('            <input class="STB" type="submit" name="Botao" value="Gravar">');

        ShowHTML('          </td>');
        ShowHTML('</tr>');
      }
      ShowHTML('</table>');

    }
    ShowHTML('         <tr valign="top">');

    ShowHTML('          </table>');
    ShowHTML('      <tr valign="top">');
    ShowHTML('      <tr>');
    ShowHTML('      <tr><td align="center" colspan="2" height="1" bgcolor="#000000">');
    ShowHTML('      <tr><td align="center" colspan="2">');
    ShowHTML('            <input class="STB" type="submit" name="Botao" value="Gravar">');
    ShowHTML('          </td>');
    ShowHTML('      </tr>');
    ShowHTML('    </table>');
    ShowHTML('    </TD>');
    ShowHTML('</tr>');
    ShowHTML('</FORM>');
    ShowHTML('</table>');
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();
  }
  //Fim da Pesquisa de Escolas Particulares

  function newsletter() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];

    if ($w_troca > '') { //Se for recarga da página
      $w_nome = $_REQUEST["w_nome"];
      $w_email = $_REQUEST["w_email"];
      $w_tipo = $_REQUEST["w_tipo"];
      $w_envia_mail = $_REQUEST["w_envia_mail"];
    }
    elseif ($O == 'L' || $O == 'G') {
      if ($_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI") {
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
      if ($O == "G") {
        $SQL .= '   and envia_mail = \'S\' ' . $crlf;
      }
      $SQL .= 'order by nome';
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }
    elseif (strpos("AEV", $O) !== false and $w_troca == '') {
      //Recupera os dados do endereço informado
      $SQL = 'select * from sbpi.Newsletter where sq_newsletter = ' . $w_chave;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_nome = f($RS, "nome");
      $w_email = f($RS, "email");
      $w_tipo = f($RS, "tipo");
      $w_envia_mail = f($RS, "envia_mail");
    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    if (strpos("IAEP", $O) !== false) {
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      if (strpos("IA", $O) !== false) {
        Validate("w_nome", "Nome", "", "1", "3", "60", "1", "1");
        Validate("w_email", "e-Mail", "", "1", "4", "60", "1", "1");
        ShowHTML('  if (theForm.w_tipo[0].checked==false && theForm.w_tipo[1].checked==false && theForm.w_tipo[2].checked==false) {');
        ShowHTML('     alert(\'Você deve selecionar uma das opções apresentadas no formulário!\');');
        ShowHTML('     return false;');
        ShowHTML('  }');
      }
      ShowHTML('  theForm.Botao[0].disabled=true;');
      ShowHTML('  theForm.Botao[1].disabled=true;');
      ValidateClose();
      ScriptClose();
    }
    ShowHTML('</HEAD>');
    if ($w_troca > '') {
      BodyOpen('onLoad=\'document.Form.' . $w_troca . '.focus()\';');
    }
    elseif ($O == 'I' || $O == 'A') {
      BodyOpen('onLoad=\'document.Form.w_nome.focus()\';');
    } else {
      BodyOpen('onLoad=\'document.focus()\';');
    }
    ShowHTML('<B><FONT COLOR="#000000">Lista de distribuição de informativos</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    if ($O == "L") {
      //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2">');
      ShowHTML('<tr><td><a accesskey="I" class="SS" href="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=I&w_chave=' . $w_chave . '"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('        <a accesskey="L" class="SS" href="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=G&&w_chave=' . $w_chave . '"><u>L</u>istar e-mails </a>&nbsp;');
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><b>Nome</font></td>');
      ShowHTML('          <td><b>Tipo</font></td>');
      ShowHTML('          <td><b>Envia</font></td>');
      ShowHTML('          <td><b>Cadastro</font></td>');
      ShowHTML('          <td><b>Alteração</font></td>');
      ShowHTML('          <td><b>Operações</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=4 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        $i = 0;
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td><a class="HL" href="mailto:' . f($row, "email") . '" title="' . f($row, "email") . '">' . f($row, "nome") . '</a></td>');
          ShowHTML('        <td align="center">' . f($row, "nm_Tipo") . '</td>');
          ShowHTML('        <td align="center">' . f($row, "nm_envia") . '</td>');
          ShowHTML('        <td align="center">' . FormataDataEdicao(FormatDateTime(f($row, "data_inclusao"), 2)) . '</td>');
          ShowHTML('        <td align="center">');
          if (f($row, "data_alteracao") > "") {
            ShowHTML(FormataDataEdicao(FormatDateTime(f($row, "data_alteracao"), 2)));
          } else {
            ShowHTML('---');
          }
          ShowHTML('</td>');
          ShowHTML('        <td align="top" nowrap>');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=A&w_chave=' . f($row, 'chave') . '">Alterar</A>&nbsp');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=E&w_chave=' . f($row, 'chave') . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
          $i++;
          if ($i > 500)
            break;
        }
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif ($O == 'G') {
      //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2">');
      ShowHTML('       <a accesskey="V" class="SS" href="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=L&w_chave=' . f($row, 'chave') . '"><u>V</u>oltar</a>&nbsp;');
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><b>Lista (cada linha com 20 e-mails)</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        $i = 0;
        $j = 0;
        foreach ($RS as $row) {
          if ($i == 0 or $j >= 20) {
            $i = 1;
            ShowHTML('      <tr valign="top"><td>');
            $j = 0;
          }
          ShowHTML('        "' . f($row, "email") . '"; ');
          $j++;
        }
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif (strpos("IAEV", $O) !== false) {
      if (strpos("EV", $O) !== false) {
        $w_Disabled = " DISABLED ";
      }
      AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
      ShowHTML('<INPUT type="hidden" name="SG" value="NEWSLETTER">');
      ShowHTML('<INPUT type="hidden" name="CL" value="' . $CL . '">');
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML('<INPUT type="hidden" name="w_sq_cliente" value="' . str_replace($CL, "sq_cliente=", "") . '">');
      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');

      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('    <table width=""95%"" border=""0"">');
      ShowHTML('      <TR><TD><font size="2" CLASS="BTM"><b>Nome completo:</b><br><input ' . $w_Disabled . ' type="text" size="60" maxlength="60" name="w_nome" value="' . $w_nome . '" CLASS="STI">');
      ShowHTML('      <TR><TD><font size="2" CLASS="BTM"><b>e-Mail:</b><br><input ' . $w_Disabled . ' type="text" size="60" maxlength="60" name="w_email" value="' . $w_email . '" CLASS="STI">');
      ShowHTML('      <TR><TD><font size="2" CLASS="BTM"><b>Tipo da pessoa:</b> ');
      if ($w_tipo == 1) {
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="1" checked> Pai, mãe ou responsável por aluno da rede de ensino ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="2"> Aluno da rede de ensino ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="3"> Outro ');
      }
      elseif ($w_tipo == 2) {
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="1"> Pai, mãe ou responsável por aluno da rede de ensino ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="2" checked> Aluno da rede de ensino ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="3"> Outro ');
      }
      elseif ($w_tipo == 3) {
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="1"> Pai, mãe ou responsável por aluno da rede de ensino ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="2"> Aluno da rede de ensino ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="3" checked> Outro ');
      } else {
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="1"> Pai, mãe ou responsável por aluno da rede de ensino ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="2"> Aluno da rede de ensino ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="3"> Outro ');
      }
      ShowHTML('      <TR> ');
      MontaRadioSN("<b>Envia newsletter para este e-mail?</b>", $w_envia_mail, "w_envia_mail");
      ShowHTML('      </TR> ');
      ShowHTML('      <tr><td align=""center"" colspan=4><hr>');
      if ($O == 'E') {
        ShowHTML('   <input class="STB" type="submit" name="Botao" value="Excluir">');
      } else {
        if ($O == "I") {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Incluir">');
        } else {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }
      ShowHTML('            <input class="STB" type="button" onClick="location.href=\'' . $dir . $w_pagina . $par . $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
      ShowHTML('          </td>');
      ShowHTML('      </tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');
      ShowHTML('</FORM>');
    } else {
      ScriptOpen("JavaScript");
      ShowHTML(' alert(\'Opção não disponível\');');
      //ShowHTML (' history.back(1);');
      ScriptClose();
    }
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();
  }
  // =========================================================================
  // Fim do cadastro de newsletter
  // -------------------------------------------------------------------------

  // =========================================================================
  // Cadastro de notícias
  // -------------------------------------------------------------------------
  function Noticias() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];

    if ($w_troca > '') { // Se for recarga da página
      $w_dt_noticia = $_REQUEST["w_dt_noticia"];
      $w_ds_titulo  = str_replace('"','&quot;',$_REQUEST["w_ds_titulo"]);
      $w_ds_noticia = $_REQUEST["w_ds_noticia"];
      $w_ln_noticia = $_REQUEST["w_ln_noticia"];
      $w_in_ativo   = $_REQUEST["w_in_ativo"];
      $w_in_exibe   = $_REQUEST["w_in_exibe"];
      $w_ln_externo = str_replace('"','&quot;',$_REQUEST["w_ln_externo"]);
    }
    elseif ($O == 'L') {
      //Recupera todos os registros para a listagem
      if ($_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI") {
        $SQL = "select sq_noticia as chave, ds_titulo, ds_noticia, dt_noticia, ln_externo, ativo, in_exibe from sbpi.Noticia_Cliente where sq_cliente = 0 order by dt_noticia desc";
      } else {
        $SQL = 'select sq_noticia as chave, ds_titulo, ds_noticia, dt_noticia, ln_externo, ativo, in_exibe from sbpi.Noticia_Cliente where sq_cliente= ' . $CL . ' order by dt_noticia desc';
      }
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }
    elseif (strpos("AEV", $O) !== false && $w_troca == '') {
      //Recupera os dados do endereço informado
      $SQL = "select sq_noticia as chave, ds_titulo, ds_noticia, dt_noticia, ln_externo, ativo, in_exibe from sbpi.Noticia_Cliente where sq_noticia = " . $w_chave;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_dt_noticia = FormataDataEdicao(f($RS, "dt_noticia"));
      $w_ds_titulo  = str_replace('"','&quot;',f($RS, "ds_titulo"));
      $w_ds_noticia = f($RS, "ds_noticia");
      $w_in_ativo   = f($RS, "in_ativo");
      $w_in_exibe   = f($RS, "in_exibe");
      $w_ln_externo = str_replace('"','&quot;',f($RS, "ln_externo"));
    }
    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('   <link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    if (strpos("IAEP", $O) !== false) {
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      if (strpos("IA", $O) !== false) {
        Validate("w_dt_noticia", "Data", "DATA", "1", "10", "10", "1", "1");
        Validate("w_ds_titulo", "Título", "", "1", "2", "50", "1", "1");
        Validate("w_ds_noticia", "Descrição", "", "1", "2", "4000", "1", "1");
      }
      ShowHTML('  theForm.Botao[0].disabled=true;');
      ShowHTML('  theForm.Botao[1].disabled=true;');
      ValidateClose();
      ScriptClose();
    }

    ShowHTML('</HEAD>');
    if ($w_troca > "") {
      BodyOpen('onLoad="document.Form.' . $w_troca . '.focus()";');
    }
    elseif ($O == "I" or $O == "A") {
      BodyOpen("onLoad='document.Form.w_ds_titulo.focus()';");
    } else {
      BodyOpen("onLoad='document.focus()';");
    }
    ShowHTML('<B><FONT COLOR="#000000">Cadastro de notícias da rede de ensino</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    If ($O == 'L') {
      // Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $dir . $w_pagina . $par . $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '>');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><b>Data</font></td>');
      ShowHTML('          <td align="center"><b>Título</font></td>');
      ShowHTML('          <td align="center"><b>Descrição</font></td>');
      ShowHTML('          <td><b>Ativo</font></td>');
      ShowHTML('          <td><b>Operações</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td align="center">' . FormataDataEdicao(FormatDateTime(f($row, "dt_noticia"), 2)) . '</td>');
          if(f($row,"ln_externo") != ""){
            ShowHTML('        <td><font size="1"><a href="'.f($row,"ln_externo").'" target="_blank">' . f($row, "ds_titulo") . '</a></td>');
          }else{
            ShowHTML('        <td><font size="1">' . f($row, "ds_titulo") . '</td>');
          }
          ShowHTML('        <td>' . f($row, "ds_noticia") . '</td>');
          ShowHTML('        <td align="center">' . f($row, "ativo") . '</td>');
          ShowHTML('        <td align="top" nowrap>');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=A&w_chave=' . f($row, 'chave') . '">Alterar</A>&nbsp');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=E&w_chave=' . f($row, 'chave') . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
        }
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif (strpos("IAEV", $O) !== false) {
      if (strpos("EV", $O) !== false) {
        $w_Disabled = ' DISABLED ';
      }
      AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
      ShowHTML('<INPUT type="hidden" name="SG" value="NOTICIAS">');
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');

      ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td valign="top"><b><u>D</u>ata:</b><br><input ' . $w_Disabled . ' accesskey="D" type="text" name="w_dt_noticia" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_noticia, Time()), 2)) . '" onKeyDown="FormataData(this,event);" ></td>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td valign="top"><b><u>T</u>ítulo:</b><br><input ' . $w_Disabled . ' accesskey="T" type="text" name="w_ds_titulo" class="STI" SIZE="90" MAXLENGTH="100" VALUE="' . $w_ds_titulo . '"></td>');
      ShowHTML('      <tr valign="top">');
      ShowHTML('          <td valign="top"><font size="1"><b><u>L</u>ink externo:</b><br><input " . w_Disabled . " accesskey="L" type="text" name="w_ln_externo" class="STI" SIZE="90" MAXLENGTH="255" VALUE="' . $w_ln_externo . '"></td>');
      ShowHTML('      </tr>');
      ShowHTML('        </table>');
      ShowHTML('      <tr><td valign="top"><b>D<u>e</u>scrição:</b><br><textarea ' . $w_Disabled . ' accesskey="E" name="w_ds_noticia" class="STI" ROWS=5 cols=65>' . $w_ds_noticia . '</TEXTAREA></td>');
      ShowHTML('      <tr>');
      ShowHTML('      </tr>');
      ShowHTML('      <tr>');
      MontaRadioSN("<b>Exibir no site?</b>", $w_in_ativo, "w_in_ativo");
      ShowHTML('      <tr>');
      ShowHTML('      <tr><td align="center" colspan=4><hr>');
      if ($O == 'E') {
        ShowHTML('   <input class="STB" type="submit" name="Botao" value="Excluir">');
      } else {
        if ($O == "I") {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Incluir">');
        } else {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }
      ShowHTML('            <input class="STB" type="button" onClick="location.href=\'' . $dir . $w_pagina . $par . $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
      ShowHTML('          </td>');
      ShowHTML('      </tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');
      ShowHTML('</FORM>');
    } else {
      ScriptOpen("JavaScript");
      ShowHTML(' alert(\'Opção não disponível\');');
      //ShowHTML (' history.back(1);');
      ScriptClose();
    }
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();
  }
  // =========================================================================
  // Fim do cadastro de notícias
  // -------------------------------------------------------------------------

  // =========================================================================
  // Cadastro de mensagens ao aluno
  // -------------------------------------------------------------------------
  function msgalunos() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];
    $w_texto = $_REQUEST["w_texto"];

    if ($w_troca > '') { // Se for recarga da página
      $w_dt_mensagem = $_REQUEST["w_dt_mensagem"];
      $w_ds_mensagem = $_REQUEST["w_ds_mensagem"];
      $w_sq_aluno = $_REQUEST["w_sq_aluno"];
    }
    elseif ($O == 'L') {
      //Recupera todos os registros para a listagem
      $SQL = "select a.sq_mensagem as chave, a.in_lida, b.no_aluno, b.nr_matricula, c.ds_cliente " . $crlf .
             "  from sbpi.Mensagem_Aluno     a " . $crlf .
             "       inner join sbpi.Aluno   b on (a.sq_aluno        = b.sq_aluno) " . $crlf .
             "       inner join sbpi.Cliente c on (b.sq_cliente = c.sq_cliente) " . $crlf .
             "order by c.ds_cliente, a.dt_mensagem desc, b.no_aluno, a.in_lida " . $crlf;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }
    elseif (strpos("AEV", $O) !== false && $w_troca == '') {
      //Recupera os dados do endereço informado
      $SQL = 'select * from sbpi.Mensagem_Aluno where sq_mensagem = ' . $w_chave;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_dt_mensagem = FormataDataEdicao(f($row, "dt_mensagem"));
      $w_ds_mensagem = f($row, "ds_mensagem");
      $w_sq_aluno = f($row, "sq_aluno");
    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('   <link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    if (strpos("IAEP", $O) !== false) {
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      if (strpos("IA", $O) !== false) {
        Validate("w_dt_mensagem", "Data", "DATA", "1", "10", "10", "1", "1");
        Validate("w_ds_mensagem", "Mensagem", "", "1", "2", "4000", "1", "1");
        if ($O == "I") {
          Validate("w_sq_aluno", "Aluno", "SELECT", "1", "1", "10", "", "1");
        }
      }
      ShowHTML('  theForm.Botao[0].disabled=true;');
      ShowHTML('  theForm.Botao[1].disabled=true;');
      ValidateClose();
      ScriptClose();
    }
    ShowHTML('</HEAD>');
    if ($w_troca > "") {
      BodyOpen('onLoad="document.Form.' . $w_troca . '.focus()";');
    }
    elseif ($O == "I" or $O == "A") {
      BodyOpen("onLoad='document.Form.w_dt_mensagem.focus()';");
    } else {
      BodyOpen("onLoad='document.focus()';");
    }
    ShowHTML('<B><FONT COLOR="#000000">Mensagens a alunos da rede de ensino</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    If ($O == 'L') {
      //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $dir . $w_pagina . $par . $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><b>Unidade</font></td>');
      ShowHTML('          <td><b>Data</font></td>');
      ShowHTML('          <td><b>Matrícula</font></td>');
      ShowHTML('          <td><b>Aluno</font></td>');
      ShowHTML('          <td><b>Lida</font></td>');
      ShowHTML('          <td><b>Operações</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td>' . strtolower(f($row, "ds_cliente")) . '</td>');
          ShowHTML('        <td align="center">' . FormataDataEdicao(FormatDateTime(f($row, "dt_mensagem"), 2)) . '</td>');
          ShowHTML('        <td align="center" nowrap>' . f($row, "nr_matricula") . '</td>');
          ShowHTML('        <td>' . strtolower(f($row, "no_aluno")) . '</td>');
          ShowHTML('        <td align="center">' . f($row, "in_lida") . '</td>');
          ShowHTML('        <td align="top" nowrap>');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=A&w_chave=' . f($row, 'chave') . '">Alterar</A>&nbsp');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=E&w_chave=' . f($row, 'chave') . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
        }
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif (strpos("IAEV", $O) !== false) {
      if (strpos("EV", $O) !== false) {
        $w_Disabled = ' DISABLED ';
      }
      AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
      ShowHTML('<INPUT type="hidden" name="SG" value="MSGALUNOS">');
      ShowHTML('<INPUT type="hidden" name="CL" value="' . $CL . '">');
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML('<INPUT type="hidden" name="w_troca" value="">');
      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');

      ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('      <tr><td><b><u>D</u>ata:</b><br><input ' . $w_Disabled . ' accesskey="D" type="text" name="w_dt_mensagem" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_mensagem, Time()), 2)) . '" onKeyDown="FormataData(this,event);"></td>');
      ShowHTML('      <tr><td><b>M<u>e</u>nsagem:</b><br><textarea ' . $w_Disabled . ' accesskey="E" name="w_ds_mensagem" class="STI" ROWS=5 cols=65>' . $w_ds_mensagem . '</TEXTAREA>');
      if ($O == "I") {
        ShowHTML('      <tr><td><b><u>P</u>rocurar por:</b><br><input accesskey="P" type="text" name="w_texto" class="STI" SIZE="50" MAXLENGTH="50" VALUE="' . $w_texto . '" >');
        ShowHTML('          <input type="Button" class="STB" name="Pesquisa" value="Procurar" onClick="document.Form.w_troca.value=\'w_sq_aluno\'; document.Form.action=\'' . $w_dir . $w_pagina . $par . '\'; document.Form.submit();"></td>');
        ShowHTML('      <tr><td><b><u>A</u>luno:</b><br><select accesskey="A" name="w_sq_aluno" class="STS" SIZE="1" >');
        ShowHTML('          <OPTION VALUE="">---');
        If ($w_texto > '') {
          $SQL = "select sq_aluno, nr_matricula, no_aluno " . $crlf .
          "  from sbpi.Aluno " . $crlf .
          " where (upper(no_aluno) like '%" . strtoupper($w_texto) . "%' or " . $crlf .
          "        nr_matricula like '%" . $w_texto . "%') " . $crlf .
          "order by no_aluno, nr_matricula" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            ShowHTML('          <OPTION VALUE="' . f($row, "sq_aluno") . '">' . f($row, "no_aluno") . '" ("' . f($row, "nr_matricula") . '")"');
          }
        }
        ShowHTML('          </select>');
      } else {
        $SQL = 'select nr_matricula, no_aluno from sbpi.Aluno where sq_aluno = ' . $w_sq_aluno;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        foreach ($RS as $row) {
          $RS = $row;
          break;
        }
        ShowHTML('      <tr><td><b>Aluno:<br>' . f($RS, "no_aluno") . ' (' . f($RS, "nr_matricula") . ')</td>');
      }
      ShowHTML('      <tr>');
      ShowHTML('      <tr><td align=""center"" colspan=4><hr>');

      if ($O == 'E') {
        ShowHTML('   <input class="STB" type="submit" name="Botao" value="Excluir">');
      } else {
        if ($O == "I") {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Incluir">');
        } else {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }
      ShowHTML('            <input class="STB" type="button" onClick="location.href=\'' . $dir . $w_pagina . $par . $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
      ShowHTML('          </td>');
      ShowHTML('      </tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');
      ShowHTML('</FORM>');
    } else {
      ScriptOpen("JavaScript");
      ShowHTML(' alert(\'Opção não disponível\');');
      //ShowHTML (' history.back(1);');
      ScriptClose();
    }
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();

  }
  // =========================================================================
  // Fim do cadastro de mensagens ao aluno
  // -------------------------------------------------------------------------

  // =========================================================================
  // Cadastro de modalidades de ensino
  // -------------------------------------------------------------------------
  function tipoCliente() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];

    if ($w_troca > '') { // Se for recarga da página
      $w_tipo = $_REQUEST["w_tipo"];
      $w_nome = $_REQUEST["w_nome"];
    }
    elseif ($O == 'L') {
      //Recupera todos os registros para a listagem
      $SQL = "select sq_tipo_cliente as chave, ds_tipo_cliente, ds_registro, tipo, " . $crlf .
      "       case tipo when '1' then 'Secretaria' " . $crlf .
      "                 when '2' then 'Regional'   " . $crlf .
      "                 when '3' then 'Escola'     " . $crlf .
      "       end nm_tipo                        " . $crlf .
      "  from sbpi.Tipo_Cliente                    " . $crlf;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }
    elseif (strpos("AEV", $O) !== false && $w_troca == '') {
      //Recupera os dados do endereço informado
      $SQL = "select sq_tipo_cliente as chave, ds_tipo_cliente, ds_registro, tipo, " . $crlf .
      "       case tipo when '1' then 'Secretaria' " . $crlf .
      "                 when '2' then 'Regional'   " . $crlf .
      "                 when '3' then 'Escola'     " . $crlf .
      "       end nm_tipo                        " . $crlf .
      "  from sbpi.Tipo_Cliente                    " . $crlf .
      "  where sq_tipo_Cliente = " . $w_chave;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_tipo = f($RS, "tipo");
      $w_nome = f($RS, "ds_tipo_cliente");
    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    if (strpos("IAEP", $O) !== false) {
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      if (strpos("IA", $O) !== false) {
        //Validate ("w_dt_mensagem" , "Data"     , "DATA"  , "1" , "10" , "10"   , "1" , "1");
        Validate("w_nome", "Nome", "", "1", "2", "50", "1", "1");
      }
      ShowHTML('  theForm.Botao[0].disabled=true;');
      ShowHTML('  theForm.Botao[1].disabled=true;');
      ValidateClose();
      ScriptClose();
    }
    ShowHTML('</HEAD>');
    if ($w_troca > "") {
      BodyOpen('onLoad="document.Form.' . $w_troca . '.focus()";');
    }
    elseif ($O == "I" or $O == "A") {
      BodyOpen("onLoad='document.Form.w_nome.focus()';");
    } else {
      BodyOpen("onLoad='document.focus()';");
    }
    ShowHTML('<B><FONT COLOR="#000000">Cadastro de tipos de Instituições </FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    If ($O == 'L') {
      //Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $dir . $w_pagina . $par . $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><b>Descrição</font></td>');
      ShowHTML('          <td><b>Tipo</font></td>');
      ShowHTML('          <td><b>Operações</font></td>');
      ShowHTML('        </tr>');
      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td>' . f($row, "ds_tipo_cliente") . '</td>');
          switch (f($row, "tipo")) {
            case 1 :
              ShowHTML('          <td align="center">Secretaria</font></td>');
              break;
            case 2 :
              ShowHTML('          <td align="center">Regional</font></td>');
              break;
            case 3 :
              ShowHTML('          <td align="center">Escola Pública</font></td>');
              break;
            case 4 :
              ShowHTML('          <td align="center">Escola Privada</font></td>');
              break;
            default :
              ShowHTML('        <td>---</td>');
              break;
          }
          ShowHTML('        <td align="top" nowrap>');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=A&w_chave=' . f($row, 'chave') . '">Alterar</A>&nbsp');
          ShowHTML('          <A class="HL" HREF="' . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=E&w_chave=' . f($row, 'chave') . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
        }
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif (strpos("IAEV", $O) !== false) {
      if (strpos("EV", $O) !== false) {
        $w_Disabled = ' DISABLED ';
      }
      AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
      ShowHTML('<INPUT type="hidden" name="SG" value="TIPOCLIENTE">');
      ShowHTML('<INPUT type="hidden" name="CL" value="' . $CL . '">');
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');

      ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('      <tr><td valign=""top"" colspan=""3""><table border=0 width=""100%"" cellspacing=0>');
      ShowHTML('        <tr valign="top"><td valign="top"><b>D<u>e</u>scrição:</b><br><input ' . $w_Disabled . ' accesskey="E" type="text" name="w_nome" class="STI" SIZE="60" MAXLENGTH="60" VALUE="' . $w_nome . '" ></td></tr>');
      ShowHTML('      <TD><font size="2" CLASS="BTM"><b>Tipo:</b> ');
      if ($w_tipo == 1) {
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="1" checked> Secretaria ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="2"> Regional ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="3"> Escola Pública ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="4"> Escola Privada ');
      }
      elseif ($w_tipo == 2) {
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="1"> Secretaria ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="2" checked> Regional ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="3"> Escola Pública ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="4"> Escola Privada ');
      }
      elseif ($w_tipo == 3) {
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="1"> Secretaria ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="2"> Regional ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="3" checked> Escola Pública ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="4"> Escola Privada ');
      }
      elseif ($w_tipo == 4) {
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="1"> Secretaria ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="2"> Regional ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="3"> Escola Pública ');
        ShowHTML('            <br><input ' . $w_Disabled . ' type="Radio" name="w_tipo" value="4" checked> Escola Privada ');
      } else {
        ShowHTML('            <br><input type="Radio" name="w_tipo" value="1"> Secretaria de Ensino do DF ');
        ShowHTML('            <br><input type="Radio" name="w_tipo" value="2"> Rede de Ensino ');
        ShowHTML('            <br><input type="Radio" name="w_tipo" value="3" checked> Escola Pública');
        ShowHTML('            <br><input type="Radio" name="w_tipo" value="4"> Escola Privada ');
      }
      ShowHTML('        </table>');
      ShowHTML('      <tr>');
      ShowHTML('      <tr><td align=""center"" colspan=4><hr>');

      if ($O == 'E') {
        ShowHTML('   <input class="STB" type="submit" name="Botao" value="Excluir">');
      } else {
        if ($O == "I") {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Incluir">');
        } else {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }
      ShowHTML('            <input class="STB" type="button" onClick="location.href=\'' . $dir . $w_pagina . $par . $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
      ShowHTML('          </td>');
      ShowHTML('      </tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');
      ShowHTML('</FORM>');
    } else {
      ScriptOpen("JavaScript");
      ShowHTML(' alert(\'Opção não disponível\');');
      //ShowHTML (' history.back(1);');
      ScriptClose();
    }
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();

  }
  // =========================================================================
  // Fim do cadastro de modalidades de ensino
  // -------------------------------------------------------------------------

  // =========================================================================
  // Tela de alteração da Senha
  // -------------------------------------------------------------------------
  function senha() {
    extract($GLOBALS);
    global $w_Disabled;

    if ($O != 'A') {
      $O = "A";
    }

    If ($O == "A") {
      $SQL = "select * from sbpi.Cliente where sq_cliente = " . $CL;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_sq_cliente = f($row, 'sq_cliente');
      $w_ds_senha_acesso = f($row, 'ds_senha_acesso');
    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    ScriptOpen("JavaScript");
    CheckBranco();
    FormataData();
    ValidateOpen("Validacao");
    If ($O == "A") {
      Validate("w_ds_senha_acesso", "Senha de acesso", "1", "1", "4", "14", "1", "1");
    }
    ValidateClose();
    ScriptClose();
    ShowHTML('   <link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    ShowHTML('</HEAD>');
    BodyOpen('onLoad=\'document.Form.w_ds_senha_acesso.focus();\'');
    ShowHTML('<B><FONT COLOR="#000000">Atualização da senha de acesso</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
    AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
    ShowHTML(MontaFiltro('POST'));
    ShowHTML('<INPUT type="hidden" name="SG" value="SENHA">');

    ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
    ShowHTML('    <table width="95%" border="0">');
    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><b>Senha de acesso</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Esta tela permite alterar a senha de acesso à tela de atualização dos dados da rede de ensino. Assim que a nova senha for gravada, ela já deverá ser utilizada para acessar as telas desta aplicação.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top"><b><u>S</u>enha de acesso:</b><br><INPUT ACCESSKEY="S" ' . $w_Disabled . ' class="STI" type="text" name="w_ds_senha_acesso" size="14" maxlength="14" value="' . $w_ds_senha_acesso . '"></td>');
    ShowHTML('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');

    //Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
    ShowHTML('      <tr><td align="center" colspan="3">');
    ShowHTML('            <input class="STB" type="submit" name="Botao" value="Gravar">');
    ShowHTML('          </td>');
    ShowHTML('      </tr>');
    ShowHTML('    </table>');
    ShowHTML('    </TD>');
    ShowHTML('</tr>');
    ShowHTML('</FORM>');
    ShowHTML('</table>');
    Rodape();
  }
  // =========================================================================
  // Fim da tela de dados básicos
  // -------------------------------------------------------------------------

  // =========================================================================
  // Monta a tela de senhas especiais
  // -------------------------------------------------------------------------
  function senhaesp() {
    extract($GLOBALS);
    global $w_Disabled;

    Cabecalho();
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    BodyOpen('onLoad=\'document.focus()\';');
    ShowHTML('<B><FONT COLOR=""#000000"">Senhas especiais - Listagem</FONT></B>');
    ShowHTML('<div align=center><center>');
    $SQL = "SELECT a.ds_username, a.sq_cliente, a.ds_cliente,  " . $crlf .
    "       a.ds_senha_acesso, a.no_municipio, a.sg_uf, a.dt_alteracao " . $crlf .
    "  from sbpi.Cliente a " . $crlf .
    " where a.publica = 'S' and a.sq_cliente_pai is null or a.sq_cliente_pai = 0 and a.ds_username <> 'SBPI'" . $crlf .
    "ORDER BY a.sq_cliente_pai, a.ds_cliente " . $crlf;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    ShowHTML('<TR><TD valign="top"><table border=0 width="100%" cellpadding=0 cellspacing=0>');
    if (count($RS) > 0) {

      //RS.PageSize = RS.RecordCount + 1
      //rs.AbsolutePage = 1

      ShowHTML('<tr><td><td align="right"><b><font face=Verdana size=1>Registros encontrados: ' . count($RS) . '</font></b>');
      ShowHTML('<tr><td><td>');
      ShowHTML('<table border="1" cellspacing="0" cellpadding="0" width="100%">');
      ShowHTML('<tr align="center" valign="top">');
      ShowHTML('    <td><font face="Verdana" size="1"><b>Escola</b></td>');
      ShowHTML('    <td><font face="Verdana" size="1"><b>Username</b></td>');
      ShowHTML('    <td><font face="Verdana" size="1"><b>Senha</b></td>');

      // While Not RS.EOF and doubleval(RS.AbsolutePage) = doubleval(Nvl($_REQUEST["P3"),RS.AbsolutePage))
      foreach ($RS as $row) {
        $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
        ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
        ShowHTML('    <td><font face="Verdana" size="1">' . f($row, "ds_cliente") . '</font></td>');
        ShowHTML('    <td><font face="Verdana" size="1">' . f($row, "ds_username") . '</font></td>');
        ShowHTML('    <td align="center"><font face="Verdana" size="1">' . f($row, "ds_senha_acesso") . '</font></td>');
      }

      ShowHTML('</table>');
      ShowHTML('<tr><td><td colspan=""4"" align=""center""><hr>');
    } else {
      ShowHTML('<TR><TD><TD colspan="3"><p align="justify"><img src="img/ico_educacao.gif" width="16" height="16" border="0" align="center">&nbsp;<font size="2"><b>Nenhuma ocorrência encontrada para as opções acima.');
    }

    ShowHTML('</TABLE>');

  }
  // -------------------------------------------------------------------------
  // Final da Página de Pesquisa
  // =========================================================================

  // =========================================================================
  // Monta a tela de Homologação do Calendário das Escolas Particulares
  // -------------------------------------------------------------------------
  function escPartHomolog() {
    extract($GLOBALS);
    global $w_Disabled;
    //exibevariaveis();

    $p_escola_particular = $_REQUEST["p_escola_particular"];
    $p_calendario = $_REQUEST["p_calendario"];
    $p_regiao = $_REQUEST["p_regiao"];
    $w_homologado = $_REQUEST["w_homologado"];
    if ($w_homologado != 'S') {
      $w_homologado = 'Não';
    } else {
      $w_homologado = 'Sim';
    }
    if ($p_tipo == "W") {
      //Response.ContentType = "application/msword"
      //HeaderWord p_layout
      ShowHTML('<TABLE WIDTH="100%" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR="#000000">SIGE-WEB<TD ALIGN="RIGHT"><B><FONT SIZE=4 COLOR="#000000">');
      ShowHTML('Consulta a escolas');
      ShowHTML('</FONT><TR><TD ALIGN="RIGHT"><B><FONT SIZE=2 COLOR="#000000">' . DataHora() . '</B></TD></TR>');
      ShowHTML('</FONT></B></TD></TR></TABLE>');
      ShowHTML('<HR>');
    } else {
      Cabecalho();
      ShowHTML('<HEAD>');
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      Validate("p_escola_particular", "Unidade de Ensino", "SELECT", "1", "1", "4", "", "1");
      ValidateClose();
      ScriptClose();
      ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
      ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
      ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
      ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
      ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
      ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
      ShowHTML('   <link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
      ShowHTML('</HEAD>');
      if ($_REQUEST["pesquisa"] > '') {
        BodyOpen(' onLoad="location.href=\'#lista\'"');
      }
    }
    ShowHTML('<B><FONT COLOR="#000000">' . $w_tp . '</FONT></B>');
    ShowHTML('<B><FONT COLOR="#000000">Homologação de Calendário</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="100%">');
    ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
    ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td align="center" valign="top"><table border=0 width="90%" cellspacing=0>');
    AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
    ShowHTML('<INPUT type="hidden" name="SG" value="ESCPARTHOMOLOG">');
    ShowHTML('<INPUT type="hidden" name="CL" value="' . $CL . '">');
    ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
    ShowHTML('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
    ShowHTML('<INPUT type="hidden" name="w_troca" value="">');
    ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');
    ShowHTML('      <tr><td colspan=2><table border=0 width="90%" cellspacing=0>');
    ShowHTML('        <tr valign="top">');
    SelecaoEscolaParticular('Unidad<u>e</u> de ensino:', 'E', 'Selecione unidade.', $_REQUEST["p_escola_particular"], null, "p_escola_particular", null, 'onChange="document.Form.action=\'' . $w_pagina . $par . '\'; document.Form.O.value=\'P\'; document.Form.w_troca.value=\'p_calendario\'; document.Form.submit();"');
    if ($p_escola_particular > '') {
      $SQL = "SELECT sq_particular_calendario, sq_cliente as cliente, nome, homologado FROM sbpi.Particular_Calendario WHERE sq_cliente = " . $p_escola_particular;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      //  ShowHTML ('<font color="#ff3737"><strong><a href="javascript:this.status.value;" onClick="window.open(\'calendario.php?CL=sq_cliente=' . $RS[0]["cliente"] . '&w_ea=L&w_ew=formcal&w_ee=1\',\'MetaWord\',\'width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no\');">Acessar o(s) calendário(s)</a></strong></font>');
      ShowHTML('<font color="#ff3737"><strong><a href="javascript:this.status.value;" onClick="window.open(\'calendario.php?par=formcal&CL=' . $RS[0]["cliente"] . '&O=L&w_ew=formcal&w_ee=1\',\'MetaWord\',\'width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no\');">Acessar o(s) calendário(s)</a></strong></font>');
      ShowHTML('<table id="tbhomologacao" border="1">');
      ShowHTML('<tr><td>Título do Calendário</td><td>Homologado?</td></tr>');

      foreach ($RS as $row) {
        $homologado = f($row, "homologado");
        ShowHTML('<INPUT type="hidden" name="w_cliente"    value="' . f($row, "cliente") . '">');
        ShowHTML('<tr><td>' . f($row, "nome") . '</td>');
        ShowHTML('<INPUT type="hidden" name="w_chave[]" value="' . f($row, "sq_particular_calendario") . '">');
        ShowHTML('<td><select name="w_homologado[]">');
        if (strpos(strtoupper($homologado), "N") !== false) {
          ShowHTML('<option value="N" SELECTED>Não');
          ShowHTML('<option value="S">Sim');
        }
        elseif (strpos(strtoupper($homologado), "S") !== false) {
          ShowHTML('<option value="S" SELECTED>Sim');
          ShowHTML('<option value="N">Não');
        }
        ShowHTML('</select></td>');
        ShowHTML('<td align="center" colspan="2">');

        ShowHTML('          </td>');
        ShowHTML('</tr>');
      }
      ShowHTML('</table>');

    }
    ShowHTML('         <tr valign="top">');

    ShowHTML('          </table>');
    ShowHTML('      <tr valign="top">');
    ShowHTML('      <tr>');
    ShowHTML('      <tr><td align="center" colspan="2" height="1" bgcolor="#000000">');
    ShowHTML('      <tr><td align="center" colspan="2">');
    ShowHTML('            <input class="STB" type="submit" name="Botao" value="Gravar">');
    ShowHTML('          </td>');
    ShowHTML('      </tr>');
    ShowHTML('    </table>');
    ShowHTML('    </TD>');
    ShowHTML('</tr>');
    ShowHTML('</FORM>');
    ShowHTML('</table>');
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();
  }
  // Fim da Pesquisa de Escolas Particulares

  // =========================================================================
  // Monta a tela de Pesquisa
  // -------------------------------------------------------------------------
  function GetVerifArquivo() { /*

      Dim RS1, p_regional

      Dim $SQL, $SQL2, $wcont, $SQL1, wAtual, wIN, $w_especialidade

      Set RS1 = Server.CreateObject("ADODB.RecordSet")

      p_regional = $_REQUEST["p_regional")

      if( $p_tipo == "W" ){
      Response.ContentType = "application/msword"
      HeaderWord p_layout
      ShowHTML ('<TABLE WIDTH="100%" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR="#000000">SIGE-WEB<TD ALIGN="RIGHT"><B><FONT SIZE=4 COLOR="#000000">"
      ShowHTML ('Consulta a escolas"
      ShowHTML ('</FONT><TR><TD ALIGN="RIGHT"><B><FONT SIZE=2 COLOR="#000000">".DataHora()."</B></TD></TR>"
      ShowHTML ('</FONT></B></TD></TR></TABLE>"
      ShowHTML ('<HR>"
      } else {
      Cabecalho
      ShowHTML ('<HEAD>"
      ScriptOpen("JavaScript")
      ShowHTML (' function marca() {"
      ShowHTML ('   if (document.Form2.arquivo.length==undefined) {"
      ShowHTML ('      if (document.Form2.dummy.checked) {"
      ShowHTML ('         document.Form2.arquivo.checked=true;"
      ShowHTML ('      } else {"
      ShowHTML ('         document.Form2.arquivo.checked=false;"
      ShowHTML ('      }"
      ShowHTML ('   } else {"
      ShowHTML ('      for (var i = 0; i < document.Form2.arquivo.length; i++) {"
      ShowHTML ('         if (document.Form2.dummy.checked) {"
      ShowHTML ('            document.Form2.arquivo[i].checked=true;"
      ShowHTML ('         } else {"
      ShowHTML ('            document.Form2.arquivo[i].checked=false;"
      ShowHTML ('         }"
      ShowHTML ('      }"
      ShowHTML ('   }"
      ShowHTML (' }"
      ValidateOpen "Validacao2"
      ShowHTML ('   if (document.Form2.arquivo.length==undefined) {"
      ShowHTML ('     if (!document.Form2.arquivo.checked) {"
      ShowHTML ('       alert('Indique pelo menos um arquivo a ser excluído!');"
      ShowHTML ('       return false;"
      ShowHTML ('     }"
      ShowHTML ('   } else {"
      ShowHTML ('      var w_erro = true; "
      ShowHTML ('      for (var i = 0; i < theForm.arquivo.length; i++) {"
      ShowHTML ('         if (theForm.arquivo[i].checked) {"
      ShowHTML ('            w_erro = false; "
      ShowHTML ('            break;"
      ShowHTML ('         }"
      ShowHTML ('      }"
      ShowHTML ('     if (w_erro) {"
      ShowHTML ('       alert('Indique pelo menos um arquivo a ser excluído!');"
      ShowHTML ('       return false;"
      ShowHTML ('     }"
      ShowHTML ('  }"
      ValidateClose();
      ScriptClose
      ShowHTML ('</HEAD>"
      if( $_REQUEST["pesquisa") > " ){
      BodyOpen " onLoad="location.href='#lista'""
      } else {
      BodyOpen "onLoad='document.Form.p_regional.focus()';"
      }
      }
      ShowHTML ('<B><FONT COLOR="#000000">".w_TP."</FONT></B>"
      ShowHTML ('<div align=center><center>"
      ShowHTML ('<tr bgcolor="".conTrBgColor.""><td align="center">"
      ShowHTML ('    <table width="95%" border="0">"
      if( $p_tipo == "H" ){
      Showhtml "<FORM ACTION="controle.asp" name="Form" METHOD="POST">"
      ShowHTML ('<INPUT TYPE="HIDDEN" NAME="w_ew" VALUE="".w_ew. "">"
      ShowHTML ('<INPUT TYPE="HIDDEN" NAME="CL" VALUE="".CL. "">"
      ShowHTML ('<INPUT TYPE="HIDDEN" NAME="pesquisa" VALUE="X">"
      ShowHTML ('<input type="Hidden" name="P3" value="1">"
      ShowHTML ('<input type="Hidden" name="P4" value="15">"
      ShowHTML ('<tr bgcolor="".conTrBgColor.""><td align="center">"
      ShowHTML ('    <table width="100%" border="0">"
      ShowHTML ('          <TR><TD valign="top"><table border=0 width="100%" cellpadding=0 cellspacing=0>"
      } else {
      ShowHTML ('<tr><td><div align="justify"><font size=1><b>Filtro:</b><ul>"
      }
      if( $p_tipo == "H" ){
      ShowHTML ('          <tr valign="top"><td>"
      SelecaoRegional "<u>S</u>ubordinação:", "S", "Indique a subordinação da escola.", p_regional, null, "p_regional", null, null
      } elseif( Nvl(p_regional,0) > 0 ){
      SQL = "SELECT  a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente ".VbCrLf._
      "  FROM  escCLIENTE a ".VbCrLf._
      " WHERE  a.sq_cliente = ".p_regional.VbCrLf._
      "ORDER BY a.ds_cliente "
      ConectaBD SQL
      Response.Write "          <li><b>Escolas da ".RS("ds_cliente")."</b>"
      DesconectaBD
      }
      SQL = "SELECT * FROM escTipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente".VbCrLf
      ConectaBD SQL
      if( $p_tipo == "H" ){
      ShowHTML ('          <tr valign="top"><td><td><br><b>Tipo de instituição:</b><br><SELECT CLASS="STI" NAME="p_tipo_cliente">"
      if( RS.RecordCount > 1 ){ ShowHTML ('          <option value="">Todos" }
      While Not RS.EOF
      if( doubleval(nvl(RS("sq_tipo_cliente"),0)) = doubleval(nvl($_REQUEST["p_tipo_cliente"),0)) ){
      ShowHTML ('          <option value="".RS("sq_tipo_cliente")."" SELECTED>".RS("ds_tipo_cliente")
      } else {
      ShowHTML ('          <option value="".RS("sq_tipo_cliente")."">".RS("ds_tipo_cliente")
      }
      RS.MoveNext
      Wend
      ShowHTML ('          </select>"
      } elseif( nvl($_REQUEST["p_tipo_cliente"),0) > 0 ){
      ShowHTML ('          <li><b>Tipo de instituição: "
      While Not RS.EOF
      if( doubleval(nvl(RS("sq_tipo_cliente"),0)) = doubleval(nvl($_REQUEST["p_tipo_cliente"),0)) ){ ShowHTML RS("ds_tipo_cliente") }
      RS.MoveNext
      Wend
      ShowHTML ('</b>"
      }
      if( $p_tipo == "H" ){
      ShowHTML ('  <TR><TD><TD><br><b>Lay-out do arquivo Word:</b><br>"
      if( nvl($_REQUEST["p_layout"),"LANDSCAPE") = "LANDSCAPE" ){ ShowHTML ('          <input type="radio" name="p_layout" value="LANDSCAPE" CLASS="BTM" checked> Paisagem<br>"  } else { ShowHTML ('          <input type="radio" name="p_layout" value="LANDSCAPE" CLASS="BTM"> Paisagem<br>"  }
      if( $_REQUEST["p_layout")                  = "PORTRAIT"  ){ ShowHTML ('          <input type="radio" name="p_layout" value="PORTRAIT" CLASS="BTM" checked> Retrato<br>"    } else { ShowHTML ('          <input type="radio" name="p_layout" value="PORTRAIT" CLASS="BTM"> Retrato<br> "   }
      }
      if( $p_tipo == "H" ){
      ShowHTML ('  <TR><TD><TD><br><b>Verificar arquivos que:</b><br>"
      if( nvl($_REQUEST["E"),"S") = "S" ){ ShowHTML ('          <input type="radio" name="E" value="S" CLASS="BTM" checked> existem somente no banco<br>"     } else { ShowHTML ('          <input type="radio" name="E" value="S" CLASS="BTM"> existem somente no banco<br>"      }
      if( $_REQUEST["E")          = "N" ){ ShowHTML ('          <input type="radio" name="E" value="N" CLASS="BTM" checked> existem somente no file system<br>" } else { ShowHTML ('          <input type="radio" name="E" value="N" CLASS="BTM"> existem somente no file system<br> " }
      } elseif( $_REQUEST["E") > " ){
      ShowHTML ('  <li><b>"
      if( $_REQUEST["E") = "S" ){ ShowHTML ('          Existe somente no banco</b>"        }
      if( $_REQUEST["E") = "N" ){ ShowHTML ('          Existe somente no file system</b> " }
      }
      if( $p_tipo <> "W" ){
      ShowHTML ('          </table>"
      }
      DesconectaBD
      $wcont = 0
      $SQL1 = "

      ' Seleção de etapas/modalidades
      $SQL = "SELECT DISTINCT a.curso as sq_especialidade, a.curso as ds_especialidade, 1 as nr_ordem, 'M' as tp_especialidade ".VbCrLf._
      " from escTurma_Modalidade                 AS a ".VbCrLf._
      "      INNER JOIN escTurma                 AS c ON (a.serie           = c.ds_serie) ".VbCrLf._
      "      INNER JOIN escCliente               AS d ON (c.sq_site_cliente = d.sq_cliente) ".VbCrLf._
      "UNION ".VbCrLf._
      "SELECT DISTINCT cast(a.sq_especialidade as varchar) as sq_especialidade, a.ds_especialidade,  ".VbCrLf._
      "       case a.tp_especialidade when 'J' then '1' else a.nr_ordem end as nr_ordem, ".VbCrLf._
      "       case a.tp_especialidade when 'J' then 'M' else a.tp_especialidade end as tp_especialidade".VbCrLf._
      " from escEspecialidade AS a ".VbCrLf._
      "      INNER JOIN escEspecialidade_cliente AS c ON (a.sq_especialidade = c.sq_codigo_espec) ".VbCrLf._
      "      INNER JOIN escCliente               AS d ON (c.sq_cliente       = d.sq_cliente) ".VbCrLf._
      " where a.tp_especialidade <> 'M' ".VbCrLf._
      "ORDER BY a.nr_ordem, a.ds_especialidade ".VbCrLf
      ConectaBD $SQL

      if( $p_tipo == "H" ){
      if( Not RS.EOF ){
      $wcont = 0
      wAtual = "

      ShowHTML ('          <TD valign="top"><table border="0" align="left" cellpadding=0 cellspacing=0>"
      Do While Not RS.EOF
      if( wAtual = " or wAtual <> RS("tp_especialidade") ){
      wAtual = RS("tp_especialidade")
      if( wAtual = "M" ){
      ShowHTML ('            <TR><TD colspan=2><b>Etapas/Modalidades de ensino:</b>"
      } elseif( wAtual = "R" ){
      ShowHTML ('            <TR><TD colspan=2><b>Em Regime de Intercomplementaridade:</b>"
      } else {
      ShowHTML ('            <TR><TD colspan=2><b>Outras:</b>"
      }
      }

      $wcont = $wcont + 1
      marcado = "N"
      For i = 1 to $_REQUEST["p_modalidade").Count
      if( RS("sq_especialidade") = $_REQUEST["p_modalidade")(i) ){
      marcado = "S"
      $SQL1 = ",'".$_REQUEST["p_modalidade")(i)."'".$SQL1
      }
      Next

      if( marcado = "S" ){
      ShowHTML chr(13)."           <tr><td><input type="checkbox" name="p_modalidade" value="".RS("sq_especialidade")."" checked><td><font size=1>".RS("ds_especialidade")
      wIN = 1
      } else {
      ShowHTML chr(13)."           <tr><td><input type="checkbox" name="p_modalidade" value="".RS("sq_especialidade").""><td><font size=1>".RS("ds_especialidade")
      }
      RS.MoveNext

      if( ($wcont Mod 2) = 0 ){
      $wcont = 0
      }

      Loop
      DesconectaBD
      $SQL1 = Mid($SQL1,2)
      }
      } elseif( Nvl($_REQUEST["p_modalidade"), ") > " ){
      if( Not RS.EOF ){
      $wcont = 0

      ShowHTML ('          <li><b>Modalidades de ensino:</b><ul>"
      Do While Not RS.EOF

      ShowHTML chr(13)."           <li><font size=1>".RS("ds_especialidade")
      if( RS("sq_especialidade") = $_REQUEST["p_modalidade")(i) ){
      $SQL1 = ",'".$_REQUEST["p_modalidade")(i)."'".$SQL1
      }
      wIN = 1
      RS.MoveNext
      Loop
      DesconectaBD
      $SQL1 = Mid($SQL1,2)
      }
      }
      ShowHTML ('          </tr>"
      ShowHTML ('          </table>"
      if $p_tipo == "H" ){
      ShowHTML ('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000">"
      ShowHTML ('      <tr><td align="center" colspan="3">"
      ShowHTML ('            <input class="BTM" type="submit" name="Botao" value="Aplicar filtro">"
      if( Session("username") = "SBPI" ){
      ShowHTML ('            <input class="BTM" type="button" name="Botao" onClick="location.href='".w_Pagina."CadastroEscola"."&CL=".CL.MontaFiltro("GET")."&w_ea=I';" value="Nova escola">"
      }
      ShowHTML ('          </td>"
      ShowHTML ('      </tr>"
      }
      ShowHTML ('    </table>"
      ShowHTML ('    </TD>"
      ShowHTML ('</tr>"
      if $p_tipo == "H" ){ ShowHTML ('</form>" }

      ' INÍCIO DA VERIFICAÇÃO

      if( $_REQUEST["pesquisa") > " ){
      if( $_REQUEST["E") = "S" ){
      $SQL = "SELECT DISTINCT 'ARQUIVO' as tipo, d.ds_username, d.sq_cliente, d.ds_cliente, d.ds_apelido, d.ln_internet, f.ds_diretorio,".VbCrLf._
      "       h.sq_arquivo as chave, h.ds_titulo, h.nr_ordem, h.ln_arquivo ".VbCrLf._
      "  from escCliente                                 d ".VbCrLf._
      "       INNER      JOIN escCliente                 b ON (b.sq_cliente       = d.sq_cliente_pai) ".VbCrLf._
      "       INNER      JOIN escCliente_Site            f ON (d.sq_cliente       = f.sq_cliente) ".VbCrLf._
      "       INNER      JOIN escTipo_Cliente            g ON (d.sq_tipo_cliente  = g.sq_tipo_cliente) ".VbCrLf._
      "       INNER      JOIN escCliente_Arquivo         h ON (d.sq_cliente       = h.sq_site_cliente) ".VbCrLf._
      " where 1 = 1 ".VbCrLf
      if( Mid(Session("username"),1,2) = "RE" ){
      SQL = SQL."   and b.ds_username = '".Session("USERNAME")."' ".VbCrLf
      }
      if( $_REQUEST["D") > " ){ SQL = SQL."   and d.dt_alteracao     is not null ".VbCrLf }

      if( $_REQUEST["p_regional") > " ){
      $SQL = $SQL + "    and d.sq_cliente_pai = ".$_REQUEST["p_regional").VbCrLf
      } else {
      $SQL = $SQL + "    and g.tipo = 3".VbCrLf
      }

      if( $_REQUEST["p_tipo_cliente") > "          ){ $SQL = $SQL + "    and d.sq_tipo_cliente= ".$_REQUEST["p_tipo_cliente")         .VbCrLf }

      if $SQL1 > " then
      $SQL = $SQL._
      "    and (0 < (select count(*) from escEspecialidade_Cliente where sq_cliente = d.sq_cliente and cast(sq_codigo_espec as varchar) in (" + $SQL1 + ")) or ".VbCrLf._
      "         0 < (select count(*) from escTurma_Modalidade  w INNER JOIN escTurma x ON (w.serie = x.ds_serie) where x.sq_site_cliente = d.sq_cliente and w.curso in (" + $SQL1 + ")) ".VbCrLf._
      "        ) ".VbCrLf
      end if
      $SQL = $SQL._
      "UNION SELECT DISTINCT 'FOTO' as tipo, d.ds_username, d.sq_cliente, d.ds_cliente, d.ds_apelido, d.ln_internet, f.ds_diretorio,".VbCrLf._
      "       h.sq_cliente_foto as chave, h.ds_foto as ds_titulo, h.nr_ordem, h.ln_foto as ln_arquivo ".VbCrLf._
      "  from escCliente                                 d ".VbCrLf._
      "       INNER      JOIN escCliente                 b ON (b.sq_cliente       = d.sq_cliente_pai) ".VbCrLf._
      "       INNER      JOIN escCliente_Site            f ON (d.sq_cliente       = f.sq_cliente) ".VbCrLf._
      "       INNER      JOIN escTipo_Cliente            g ON (d.sq_tipo_cliente  = g.sq_tipo_cliente) ".VbCrLf._
      "       INNER      JOIN escCliente_Foto            h ON (d.sq_cliente       = h.sq_cliente) ".VbCrLf._
      " where 1 = 1 ".VbCrLf
      if( Mid(Session("username"),1,2) = "RE" ){
      SQL = SQL."   and b.ds_username = '".Session("USERNAME")."' ".VbCrLf
      }
      if( $_REQUEST["D") > " ){ SQL = SQL."   and d.dt_alteracao     is not null ".VbCrLf }

      if( $_REQUEST["p_regional") > " ){
      $SQL = $SQL + "    and d.sq_cliente_pai = ".$_REQUEST["p_regional").VbCrLf
      } else {
      $SQL = $SQL + "    and g.tipo = 3".VbCrLf
      }

      if( $_REQUEST["p_tipo_cliente") > "          ){ $SQL = $SQL + "    and d.sq_tipo_cliente= ".$_REQUEST["p_tipo_cliente")         .VbCrLf }

      if $SQL1 > " then
      $SQL = $SQL._
      "    and (0 < (select count(*) from escEspecialidade_Cliente where sq_cliente = d.sq_cliente and cast(sq_codigo_espec as varchar) in (" + $SQL1 + ")) or ".VbCrLf._
      "         0 < (select count(*) from escTurma_Modalidade  w INNER JOIN escTurma x ON (w.serie = x.ds_serie) where x.sq_site_cliente = d.sq_cliente and w.curso in (" + $SQL1 + ")) ".VbCrLf._
      "        ) ".VbCrLf
      end if
      $SQL = $SQL + "ORDER BY d.ds_cliente, tipo desc, h.nr_ordem, h.ds_titulo ".VbCrLf
      ConectaBD SQL

      ShowHTML ('<TR><TD valign="top"><br><table border=0 width="100%" cellpadding=0 cellspacing=0>"
      if( Not RS.EOF ){

      if( $p_tipo == "H" ){
      if( $_REQUEST["P4") > " ){ RS.PageSize = doubleval($_REQUEST["P4")) } else { RS.PageSize = 15 }
      rs.AbsolutePage = Nvl($_REQUEST["P3"),1)
      } else {
      RS.PageSize = RS.RecordCount + 1
      rs.AbsolutePage = 1
      }


      ShowHTML ('<tr><td><td align="right"><b><font face=Verdana size=1></font></b>"
      if( p_Tipo = "H" ){ ShowHTML ('     &nbsp;&nbsp;<A TITLE="Clique aqui para gerar arquivo Word com a listagem abaixo" class="SS" href="#"  onClick="window.open('controle.asp?$p_tipo=W&w_ew=".w_ew."&Q=".$_REQUEST["Q")."&C=".$_REQUEST["C")."&D=".$_REQUEST["D")."&U=".$_REQUEST["U").$w_especialidade.MontaFiltro("GET")."','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no');">Gerar Word<IMG ALIGN="CENTER" border=0 SRC="img/word.gif"></A>" }
      ShowHTML ('<tr><td><td>"
      ShowHTML ('<table border="1" cellspacing="0" cellpadding="0" width="100%">"
      AbreForm "Form2", w_Pagina."Grava", "POST", "return(Validacao2(this));", null
      ShowHTML ('<INPUT TYPE="HIDDEN" NAME="R" VALUE="VERIFBANCO">"
      ShowHTML ('<INPUT TYPE="HIDDEN" NAME="CL" VALUE="".CL. "">"
      ShowHTML ('<INPUT TYPE="HIDDEN" NAME="pesquisa" VALUE="X">"
      ShowHTML ('<input type="Hidden" name="P3" value="1">"
      ShowHTML ('<input type="Hidden" name="P4" value="15">"

      ShowHTML ('<tr align="center" valign="top">"
      ShowHTML ('    <td><font face="Verdana" size="1"><b>Escola</b></td>"
      ShowHTML ('    <td><font face="Verdana" size="1"><b>Tipo</b></td>"
      ShowHTML ('    <td><font face="Verdana" size="1"><b>Ordem</b></td>"
      ShowHTML ('    <td><font face="Verdana" size="1"><b>Título</b></td>"
      ShowHTML ('    <td><font face="Verdana" size="1"><b>Link</b></td>"
      ShowHTML ('    <td id="checkbox"><font face="Verdana" size="1"><input type="CHECKBOX" name="dummy" value="none" onClick="marca()"></b></td>"
      ShowHTML ('    <script>document.getElementById('checkbox').style.display='none';</script>"

      $w_cor   = "#FDFDFD"
      $w_atual = "
      checkbox = "unok"

      Set FS = CreateObject("Scripting.FileSystemObject")

      While Not RS.EOF
      strFile = replace(conFilePhysical."\sedf\".RS("ds_username")."\".RS("ln_arquivo"),"\\","\")

      ' Remove o arquivo, caso ele já exista
      if( Not FS.FileExists (strFile) then
      if( checkbox <> "ok" ){
      ShowHTML ('    <script>document.getElementById('checkbox').style.display='block';</script>"
      checkbox = "ok"
      }
      if( $w_atual = " or $w_atual <> RS("DS_CLIENTE") ){
      if( $w_cor = "#EFEFEF" or $w_cor = " ){ $w_cor = "#FDFDFD" } else { $w_cor = "#EFEFEF" }
      ShowHTML ('<tr valign="top" bgcolor="".$w_cor."">"
      if( Not IsNull (RS("LN_INTERNET")) ){
      if( inStr(lcase(RS("LN_INTERNET")),"http://") > 0 ){
      ShowHTML ('                <a href="http://".replace(RS("LN_INTERNET"),"http://",")."" target="_blank">".RS("DS_CLIENTE")."</a></b>"
      } else {
      ShowHTML ('                <a href="".RS("LN_INTERNET")."" target="_blank">".RS("DS_CLIENTE")."</a></b>"
      }
      } else {
      ShowHTML ('    <td><font face="Verdana" size="1">".RS("DS_CLIENTE")."</font></td>"
      }
      $w_atual = RS("DS_CLIENTE")
      } else {
      ShowHTML ('<tr valign="top" bgcolor="".$w_cor."">"
      ShowHTML ('    <td><font face="Verdana" size="1">&nbsp;</font></td>"
      }
      ShowHTML ('    <TD align="center"><font face="Verdana" size="1">".RS("tipo")
      ShowHTML ('    <TD align="center"><font face="Verdana" size="1">".RS("nr_ordem")
      ShowHTML ('    <TD><font face="Verdana" size="1">".RS("ds_titulo")
      ShowHTML ('    <TD><font face="Verdana" size="1"><a href="".RS("ln_internet")."/".RS("ln_arquivo")."" target="_blank">".RS("ln_arquivo")."</a>"
      ShowHTML ('<td align="center" width="1%" nowrap><input type="checkbox" name="arquivo" value="".RS("tipo")."=|=".RS("chave").""></td>"
      }

      RS.MoveNext
      Wend

      ShowHTML ('</table>"
      ShowHTML ('<tr><td><td colspan="5" align="center"><input type="SUBMIT" name="Botao" value="Remover todos os registros indicados">"
      ShowHTML ('</FORM>"
      ShowHTML ('<tr><td><td colspan="5" align="center"><hr>"

      } else {

      ShowHTML ('<TR><TD><TD colspan="3"><p align="justify"><img src="img/ico_educacao.gif" width="16" height="16" border="0" align="center">&nbsp;<font size="2"><b>Nenhuma ocorrência encontrada para as opções acima."

      }
      } else {
      $SQL = "SELECT d.ds_username, d.sq_cliente, d.ds_cliente, d.ds_apelido, d.ln_internet, f.ds_diretorio ".VbCrLf._
      "  from escCliente                                 d ".VbCrLf._
      "       INNER      JOIN escCliente                 b ON (b.sq_cliente       = d.sq_cliente_pai) ".VbCrLf._
      "       INNER      JOIN escCliente_Site            f ON (d.sq_cliente       = f.sq_cliente) ".VbCrLf._
      "       INNER      JOIN escTipo_Cliente            g ON (d.sq_tipo_cliente  = g.sq_tipo_cliente) ".VbCrLf._
      " where 1 = 1 ".VbCrLf
      if( Mid(Session("username"),1,2) = "RE" ){
      SQL = SQL."   and b.ds_username = '".Session("USERNAME")."' ".VbCrLf
      }
      if( $_REQUEST["D") > " ){ SQL = SQL."   and d.dt_alteracao     is not null ".VbCrLf }

      if( $_REQUEST["p_regional") > " ){
      $SQL = $SQL + "    and d.sq_cliente_pai = ".$_REQUEST["p_regional").VbCrLf
      } else {
      $SQL = $SQL + "    and g.tipo = 3".VbCrLf
      }

      if( $_REQUEST["p_tipo_cliente") > "          ){ $SQL = $SQL + "    and d.sq_tipo_cliente= ".$_REQUEST["p_tipo_cliente")         .VbCrLf }

      if $SQL1 > " then
      $SQL = $SQL._
      "    and (0 < (select count(*) from escEspecialidade_Cliente where sq_cliente = d.sq_cliente and cast(sq_codigo_espec as varchar) in (" + $SQL1 + ")) or ".VbCrLf._
      "         0 < (select count(*) from escTurma_Modalidade  w INNER JOIN escTurma x ON (w.serie = x.ds_serie) where x.sq_site_cliente = d.sq_cliente and w.curso in (" + $SQL1 + ")) ".VbCrLf._
      "        ) ".VbCrLf
      end if
      $SQL = $SQL + "ORDER BY d.ds_cliente ".VbCrLf
      ConectaBD SQL

      ShowHTML ('<TR><TD valign="top"><br><table border=0 width="100%" cellpadding=0 cellspacing=0>"
      if( Not RS.EOF ){

      if( $p_tipo == "H" ){
      if( $_REQUEST["P4") > " ){ RS.PageSize = doubleval($_REQUEST["P4")) } else { RS.PageSize = 15 }
      rs.AbsolutePage = Nvl($_REQUEST["P3"),1)
      } else {
      RS.PageSize = RS.RecordCount + 1
      rs.AbsolutePage = 1
      }


      ShowHTML ('<tr><td><td align="right"><b><font face=Verdana size=1></font></b>"
      if( p_Tipo = "H" ){ ShowHTML ('     &nbsp;&nbsp;<A TITLE="Clique aqui para gerar arquivo Word com a listagem abaixo" class="SS" href="#"  onClick="window.open('controle.asp?$p_tipo=W&w_ew=".w_ew."&Q=".$_REQUEST["Q")."&C=".$_REQUEST["C")."&D=".$_REQUEST["D")."&U=".$_REQUEST["U").$w_especialidade.MontaFiltro("GET")."','MetaWord','width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no');">Gerar Word<IMG ALIGN="CENTER" border=0 SRC="img/word.gif"></A>" }
      ShowHTML ('<tr><td><td>"
      ShowHTML ('<table border="1" cellspacing="0" cellpadding="0" width="100%">"
      ShowHTML ('<tr align="center" valign="top">"
      AbreForm "Form2", w_Pagina."Grava", "POST", "return(Validacao2(this));", null
      ShowHTML ('<INPUT TYPE="HIDDEN" NAME="R" VALUE="".w_ew. "">"
      ShowHTML ('<INPUT TYPE="HIDDEN" NAME="CL" VALUE="".CL. "">"
      ShowHTML ('<INPUT TYPE="HIDDEN" NAME="pesquisa" VALUE="X">"
      ShowHTML ('<input type="Hidden" name="P3" value="1">"
      ShowHTML ('<input type="Hidden" name="P4" value="15">"

      ShowHTML ('    <td><font face="Verdana" size="1"><b>Escola</b></td>"
      ShowHTML ('    <td><font face="Verdana" size="1"><b>Arquivo</b></td>"
      ShowHTML ('    <td><font face="Verdana" size="1"><b>KB</b></td>"
      ShowHTML ('    <td id="checkbox"><font face="Verdana" size="1"><input type="CHECKBOX" name="dummy" value="none" onClick="marca()"></b></td>"
      ShowHTML ('    <script>document.getElementById('checkbox').style.display='none';</script>"

      $w_cor   = "#FDFDFD"
      $w_atual = "
      dim checkbox
      checkbox = "unok"

      Set FS = CreateObject("Scripting.FileSystemObject")
      While Not RS.EOF
      strDir = replace(conFilePhysical."\sedf\".RS("ds_username")."\","\\","\")

      dim FileSystem
      dim File
      dim Folder
      dim w_total


      set FileSystem = Server.CreateObject("Scripting.FileSystemObject")
      set Folder = FileSystem.GetFolder(strDir)

      For each File in Folder.files
      if( File.name <> "default.asp" ){
      sqlstr = "select coalesce(a.qtd,0) + coalesce(b.qtd,0) as quantidade ".VbCrLf._
      "from (select count(*) as qtd from escCliente_Foto where sq_cliente = ".RS("SQ_CLIENTE")." and ln_foto = '".File.name."') as a, ".VbCrLf._
      "     (select count(*) as qtd from escCliente_Arquivo where sq_site_cliente = ".RS("SQ_CLIENTE")." and ln_arquivo = '".File.name."') as b ".VbCrLf

      Set RsQtd = dbms.Execute(sqlstr)
      if( RsQtd("quantidade") = 0 ){
      if( checkbox <> "ok" ){
      ShowHTML ('    <script>document.getElementById('checkbox').style.display='block';</script>"
      checkbox = "ok"
      }
      if( $w_atual = " or $w_atual <> RS("DS_CLIENTE") ){
      if( $w_cor = "#EFEFEF" or $w_cor = " ){ $w_cor = "#FDFDFD" } else { $w_cor = "#EFEFEF" }
      ShowHTML ('<tr valign="top" bgcolor="".$w_cor."">"
      ShowHTML ('<td width="1%" nowrap><font face="Verdana" size="1">".RS("DS_CLIENTE")."<br>".strDir."</td>"
      $w_atual = RS("DS_CLIENTE")
      } else {
      ShowHTML ('<tr valign="top" bgcolor="".$w_cor."">"
      ShowHTML ('<td></td>"
      }

      ShowHTML ('<td><font face="Verdana" size="1">".(File.name)
      ShowHTML ('<td align="right" width="1%" nowrap><font face="Verdana" size="1">".formatNumber(File.size/(1024),0)."&nbsp;&nbsp;"
      ShowHTML ('<td align="center" width="1%" nowrap><input type="checkbox" name="arquivo" value="".RS("DS_USERNAME")."=|=".File.name.""></td>"
      ShowHTML ('</td>"
      w_total = w_total + File.size
      }
      checkbox = "unok"
      }
      Next

      RS.MoveNext
      Wend
      ShowHTML ('<tr><td colspan=2 align="right"><font face="Verdana" size="1">Total<td align="right" width="1%" nowrap><font face="Verdana" size="1">".formatnumber(w_total/1024,0)."&nbsp;&nbsp;</td>"
      ShowHTML ('</table>"
      ShowHTML ('<tr><td><td colspan="5" align="center"><input type="SUBMIT" name="Botao" value="Remover todos os arquivos indicados">"
      ShowHTML ('</FORM>"
      ShowHTML ('<tr><td><td colspan="5" align="center"><hr>"
      } else {

      ShowHTML ('<TR><TD><TD colspan="3"><p align="justify"><img src="img/ico_educacao.gif" width="16" height="16" border="0" align="center">&nbsp;<font size="2"><b>Nenhuma ocorrência encontrada para as opções acima."

      }
      }

      }
      ShowHTML ('</TABLE>"
      Set p_regional = Nothing*/
  }

  function Grava() {
    extract($GLOBALS);
    if ($SG != 'ADM') {
      Cabecalho();
      ShowHTML('</HEAD>');
      ShowHTML('<BASE HREF="' . $conRootSIW . '">');
      BodyOpen('onLoad=this.focus();');
    }

    switch ($SG) {

      case "CADASTROESCOLA" :

        If ($O == "I") {
          $SQL = "SELECT ds_username from sbpi.Cliente where ds_username = '" . $_REQUEST["w_ds_username"] . "'";
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          If (count($RS) > 0) {
            ScriptOpen("JavaScript");
            ShowHTML("alert('Username já existente!');");
            ScriptClose();
            retornaFormulario('w_ds_username');
            die();
          }

          $SQL = "SELECT sbpi.sq_cliente.nextval  as sq_cliente  from dual";
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          $RS = $RS[0];
          $w_chave = ((int) $RS["sq_cliente"]);

        }
        ElseIf (Nvl($O, "") == "A") {
          If ($_REQUEST["w_ds_username"] != $_REQUEST["w_username_atual"]) {
            $SQL = "SELECT ds_username from sbpi.Cliente where ds_username = '" . $_REQUEST["w_ds_username"] . "'";
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            If (count($RS) > 0) {
              ScriptOpen("JavaScript");
              ShowHTML("alert('Username já existente!');");
              ScriptClose();
              retornaFormulario('w_ds_username');
              die();
            }

          }
          $w_chave = $_REQUEST["w_chave"];

        }

        $w_diretorio = $conFilePhysical . strtolower(trim($_REQUEST['w_ds_username']));
        $w_cliente = $w_chave;
        $w_origem = str_replace('sedf/', '', $conFilePhysical) . 'modelos/mod' . strtolower($_REQUEST['w_sq_modelo']) . '/default_cliente.php';
        $w_destino = $w_diretorio . '/default.php';

        // Verifica a necessidade de criação dos diretórios do cliente
        if (!file_exists($w_diretorio)) {

          mkdir($w_diretorio);

          // Carrega o conteúdo do arquivo origem em uma variável local
          $conteudo = '';
          $handle = fopen($w_origem, 'r');
          while (!feof($handle)) {
            $conteudo .= fgets($handle, 4096);
          }
          fclose($handle);

          // Gera a página inicial do cliente, sustituindo *%* pelo código do cliente
          $handle = fopen($w_destino, 'w');
          fwrite($handle, str_replace('*%*', $w_chave, $conteudo));
          fclose($handle);
        }

        If (Nvl($O, "") == "I") {

          //'Criação das escolas na tabela principal
          $SQL = "insert into sbpi.Cliente (  " . $crlf .
                 "       sq_cliente,              sq_cliente_pai,              sq_tipo_cliente,                            ds_cliente,  " . $crlf .
                 "       ds_apelido,              no_municipio,                sg_uf,                                      ds_username, " . $crlf .
                 "       ds_senha_acesso,         ln_internet,                 ds_email,                                   dt_cadastro,  " . $crlf .
                 "       ativo  ,                 sq_regiao_adm ,              localizacao" . $crlf .
                 " ) values ( " . $crlf .
                 "       " . $w_chave . "," . $_REQUEST["w_sq_cliente_pai"] . "," . $_REQUEST["w_sq_tipo_cliente"] . ",'" . $_REQUEST["w_ds_cliente"] . "'," . $crlf .
                 "       '" . $_REQUEST["w_ds_apelido"] . "','" . $_REQUEST["w_no_municipio"] . "','" . $_REQUEST["w_sg_uf"] . "','" . $_REQUEST["w_ds_username"] . "'," . $crlf .
                 "       '9" . $w_chave . "','" . $_REQUEST["w_ln_internet"] . "','" . Nvl($_REQUEST["w_ds_email"], "Não informado") . "', sysdate,'" . $_REQUEST["w_ativo"] . "' , '" . $crlf .
                 "       " . $_REQUEST['w_regiao'] . ",'" . $_REQUEST['w_local'] . "') ";

          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          //'Criação das escolas na tabela que contém dados complementares'
          $SQL = "insert into sbpi.Cliente_Dados (" . $crlf .
                 "       sq_cliente,                  nr_cnpj,                 tp_registro,  " . $crlf .
                 "       ds_ato,                      nr_ato,              dt_ato,                  ds_orgao,         ds_logradouro, " . $crlf .
                 "       no_bairro,                   nr_cep,              no_contato,              ds_email_contato, nr_fone_contato, " . $crlf .
                 "       nr_fax_contato,              no_diretor,          no_secretario " . $crlf .
                 " ) values ( " . $crlf .
                 "       " . $w_chave . ",null,null, " . $crlf .
                 "       null,                        null,                null,                    null,'" . $_REQUEST["w_ds_logradouro"] . "'," . $crlf .
                 "       '" . $_REQUEST["w_no_bairro"] . "','" . trim($_REQUEST["w_nr_cep"]) . "','" . $_REQUEST["w_no_contato"] . "', '" . $_REQUEST["w_email_contato"] . "','" . $_REQUEST["w_nr_fone_contato"] . "'," . $crlf .
                 "       '" . $_REQUEST["w_nr_fax_contato"] . "','" . Nvl($_REQUEST["w_no_diretor"], "Não informado") . "','" . Nvl($_REQUEST["w_no_secretario"], "Não informado") . "') ";

          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          //'Criação das escolas na tabela que contém informações sobre o site web
          $SQL = "insert into sbpi.Cliente_Site (" . $crlf .
                 "       sq_cliente,                   sq_modelo,                      no_contato_internet,              " . $crlf .
                 "       nr_fone_internet,              nr_fax_internet,              ds_email_internet,              ds_diretorio,                     " . $crlf .
                 "       sq_siscol,                     ds_mensagem,                  ds_institucional,               ds_texto_abertura                 " . $crlf .
                 " ) values ( " . $crlf .
                 "       " . $w_chave . "," . $_REQUEST["w_sq_modelo"] . ",'" . $_REQUEST["w_no_contato_internet"] . "', " . $crlf .
                 "       '" . $_REQUEST["w_nr_fone_internet"] . "','" . $_REQUEST["w_nr_fax_internet"] . "','" . $_REQUEST["w_ds_email_internet"] . "','" . $_REQUEST["w_ds_diretorio"] . "'," . $crlf .
                 "       '" . $_REQUEST["w_sq_siscol"] . "','" . Nvl($_REQUEST["w_ds_mensagem"], "A ser inserido") . "','" . Nvl($_REQUEST["w_ds_institucional"], "A ser inserido") . "','" . Nvl($_REQUEST["w_ds_texto_abertura"], "A ser inserido") . "')";

          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          // 'Criação das escolas na tabela que contém as modalidades de ensino oferecidas pela escola

          if (is_array($_REQUEST['w_sq_codigo_espec'])) {
            foreach ($_REQUEST['w_sq_codigo_espec'] as $row) {
              $SQL = "insert into sbpi.Especialidade_Cliente (sq_cliente , sq_especialidade) values (" . $w_chave . "," . $row . ") ";
              db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }
          }
        } elseif (Nvl($O, "") == "A") {

          //'Altera os dados das escolas na tabela principal
          $SQL = "update sbpi.Cliente set " . $crlf .
                 "   sq_cliente_pai  = " . $_REQUEST["w_sq_cliente_pai"] . "," . $crlf .
                 "   sq_tipo_cliente = " . $_REQUEST["w_sq_tipo_cliente"] . "," . $crlf .
                 "   ds_cliente      = '" . $_REQUEST["w_ds_cliente"] . "'," . $crlf .
                 "   ds_apelido      = '" . $_REQUEST["w_ds_apelido"] . "'," . $crlf .
                 "   no_municipio    = '" . $_REQUEST["w_no_municipio"] . "'," . $crlf .
                 "   sg_uf           = '" . $_REQUEST["w_sg_uf"] . "'," . $crlf .
                 "   ds_username     = '" . $_REQUEST["w_ds_username"] . "'," . $crlf .
                 "   ln_internet     = '" . $_REQUEST["w_ln_internet"] . "'," . $crlf .
                 "   ds_email        = '" . $_REQUEST["w_ds_email"] . "'," . $crlf .
                 "   sq_regiao_adm   = '" . $_REQUEST["w_regiao"] . "'," . $crlf .
                 "   localizacao     = '" . $_REQUEST["w_local"] . "'," . $crlf .
                 "   ativo        = '" . $_REQUEST["w_ativo"] . "'" . $crlf .
                 " where sq_cliente = " . $w_chave;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          //'Altera as escolas na tabela que contém dados complementares
          $SQL = "update sbpi.Cliente_Dados set " . $crlf .
                 "   ds_logradouro   = '" . $_REQUEST["w_ds_logradouro"] . "'," . $crlf .
                 "   no_bairro       = '" . $_REQUEST["w_no_bairro"] . "'," . $crlf .
                 "   nr_cep          = '" . trim($_REQUEST["w_nr_cep"]) . "'," . $crlf .
                 "   no_contato      = '" . $_REQUEST["w_no_contato"] . "'," . $crlf .
                 "   ds_email_contato= '" . $_REQUEST["w_email_contato"] . "'," . $crlf .
                 "   nr_fone_contato = '" . $_REQUEST["w_nr_fone_contato"] . "'," . $crlf .
                 "   nr_fax_contato  = '" . $_REQUEST["w_nr_fax_contato"] . "'," . $crlf .
                 "   no_diretor      = '" . Nvl($_REQUEST["w_no_diretor"], "Não informado") . "'," . $crlf .
                 "   no_secretario   = '" . Nvl($_REQUEST["w_no_secretario"], "Não informado") . "'" . $crlf .
                 " where sq_cliente = " . $w_chave;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          //'Altera as escolas na tabela que contém informações sobre o site web
          $SQL = "update sbpi.Cliente_Site set " . $crlf .
                 "   sq_modelo             = '" . $_REQUEST["w_sq_modelo"] . "'," . $crlf .
                 "   no_contato_internet   = '" . $_REQUEST["w_no_contato_internet"] . "'," . $crlf .
                 "   nr_fone_internet      = '" . $_REQUEST["w_nr_fone_internet"] . "'," . $crlf .
                 "   nr_fax_internet       = '" . $_REQUEST["w_nr_fax_internet"] . "'," . $crlf .
                 "   ds_email_internet     = '" . $_REQUEST["w_ds_email_internet"] . "'," . $crlf .
                 "   ds_diretorio          = '" . $_REQUEST["w_ds_diretorio"] . "'," . $crlf .
                 "   sq_siscol             = '" . $_REQUEST["w_sq_siscol"] . "'," . $crlf .
                 "   ds_mensagem           = '" . Nvl($_REQUEST["w_ds_mensagem"], "A ser inserido") . "'," . $crlf .
                 "   ds_institucional      = '" . Nvl($_REQUEST["w_ds_institucional"], "A ser inserido") . "'," . $crlf .
                 "   ds_texto_abertura     = '" . Nvl($_REQUEST["w_ds_texto_abertura"], "A ser inserido") . "'" . $crlf .
                 " where sq_cliente = " . $w_chave;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          //'Apaga as escolas na tabela que contém as modalidades de ensino oferecidas pela escola
          $SQL = "Delete sbpi.Especialidade_Cliente where sq_cliente = " . $w_chave;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          //'Criação das escolas na tabela que contém as modalidades de ensino oferecidas pela escola

          if (is_array($_REQUEST['w_sq_codigo_espec'])) {
            foreach ($_REQUEST['w_sq_codigo_espec'] as $row) {
              $SQL = "insert into sbpi.Especialidade_Cliente (sq_cliente , sq_especialidade " .
              " ) values ( " . $w_chave . "," . $row . ") ";
              db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }
          }
        }
        ScriptOpen("JavaScript");
        ShowHTML("  location.href=\"" . $w_pagina . "escolas&O=L&pesquisa=Sim" . MontaFiltro("GET") ."\"; ");
        ScriptClose();

        break;

      Case "REDEPART" :
        set_time_limit(0);
        $w_diretorio = $conFilePhysical . 'particular';
        $w_arquivo = $_FILES["w_no_arquivo"]['name'];

        if ($_FILES["w_no_arquivo"]['size'] > 0) {
          //Remove o arquivo físico
          if (file_exists($w_diretorio . '/' . $w_arquivo)) {
            unlink($w_diretorio . '/' . $w_arquivo);
          }
        }

        move_uploaded_file($_FILES["w_no_arquivo"]['tmp_name'], $w_diretorio . '/' . $_FILES["w_no_arquivo"]['name']);

        $file = $w_diretorio . '/' . $w_arquivo;
        $arr = csv($file);

        $rede = $arr['[REDE]'];
        $os = $arr['[OS]'];
        $cursos = $arr['[CURSOS]'];
        $arquivos = $arr['[ARQUIVOS]'];
        $profissionais = $arr['[PROFISSIONAIS]'];
        $portarias = $arr['[PORTARIAS]'];

        //Remove os registros vinculados à escola
        $SQL = "DELETE from sbpi.Particular_Portaria";
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        //echo $SQL.'<br/><br/>';

        $SQL = "DELETE from sbpi.Particular_OS";
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        //echo $SQL.'<br/><br/>';

        $SQL = "DELETE from sbpi.Particular_Curso";
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        //echo $SQL.'<br/><br/>';

        $SQL = "DELETE from sbpi.Curso";
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        //echo $SQL.'<br/><br/>';

        $SQL = "update sbpi.Cliente_Particular SET situacao = 0";
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        $w_cont = 0;
        ShowHTML('<br>&nbsp;&nbsp;&nbsp;<b> - Rede</b>');
        foreach ($rede as $w_campo) {
          $w_cont++;
          $idPasta = $w_campo[0];
          $idInstituicao = $w_campo[1];
          $idDiretor = $w_campo[2];
          $idMantenedora = $w_campo[3];
          $idEndereco = $w_campo[4];
          $idTelefone_1 = $w_campo[5];
          $idFax = $w_campo[6];
          $idVencimento = $w_campo[7];
          $idObservacao = $w_campo[8];
          $idEscola = $w_campo[9];
          $idCodinep = $w_campo[10];
          $idTelefone_2 = $w_campo[11];
          $idEmail_1 = $w_campo[12];
          $idEmail_2 = $w_campo[13];
          $idLocalizacao = $w_campo[14];
          $idCnpjExecutora = $w_campo[15];
          $idCnpjEscola = $w_campo[16];
          $idSecretario = $w_campo[17];
          $idCep = $w_campo[18];
          $idAutHabSecretario = $w_campo[19];
          $idEinf = $w_campo[20];
          $idEf = $w_campo[21];
          $idEja = $w_campo[22];
          $idEm = $w_campo[23];
          $idEDA = $w_campo[24];
          $idEprof = $w_campo[25];
          $idRegiao = $w_campo[26];
          $idEndMantenedora = $w_campo[27];
          $SiteEscola = $w_campo[28];
          $Situacao = $w_campo[29];
          $credenc = $w_campo[30];
          $Login = $w_campo[31];
          $Senha = $w_campo[32];

          if ($idMantenedora == "NULL") {
            $idMantenedora = "'Sem informação'";
          }
          if ($idPasta == "NULL") {
            $idPasta = 0;
          }
          if ($idParecereResolucao == "NULL") {
            $idParecereResolucao = "'Sem informação'";
          }
          if ($idTelefone_1 == "NULL") {
            $idTelefone_1 = "'Sem informação'";
          }
          if ($idEndereco == "NULL") {
            $idEndereco = "'Sem informação'";
          }
          if ($idCep == "NULL") {
            $idCep = "'Sem informação'";
          }
          if ($idAutHabSecretario == "NULL") {
            $idAutHabSecretario = "0";
          }
          if ($idCodinep == "NULL") {
            $idCodinep = "0";
          }
          if ($idRegiao == "'0'") {
            $idRegiao = "1";
          }
          if ($idVencimento == "'0000-00-00'") {
            $idVencimento = "NULL";
          }

          ShowHTML('<br>&nbsp;&nbsp;&nbsp;Linha ' . $w_cont . ': ' . str_replace("'", "", $idInstituicao));

          $SQL = "SELECT count(idescola) Registros from sbpi.Cliente_Particular WHERE idEscola = " . $idEscola;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          foreach ($RS as $row) {
            $RS = $row;
            break;
          }

          if (intval(f($RS, "Registros")) == 0) {
            $SQL = "SELECT sbpi.sq_cliente.nextval MaxValue from sbpi.Cliente WHERE sq_cliente < 99000";
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            $sq_cliente = intval(f($RS, "MaxValue"));

            $SQL = "INSERT into sbpi.Cliente (sq_cliente,ativo,sq_tipo_cliente, ds_cliente, no_municipio, sg_uf, ds_username, ds_senha_acesso, localizacao, publica, sq_regiao_adm) " . $crlf .
            "(SELECT " . $sq_cliente . ", 'S', a.sq_tipo_cliente, " . $idInstituicao . ", " . $crlf .
            "        substr(b.no_regiao, instr(' ',b.no_regiao)+1,500), 'DF', " . $crlf .
            "        " . $login . ", " . $senha . ", " . $idLocalizacao . ", 'N', " . $idRegiao . $crlf .
            "   from sbpi.Tipo_Cliente a, " . $crlf .
            "        escRegiao_Administrativa b " . $crlf .
            "  WHERE a.tipo = 4 " . $crlf .
            "    and b.sq_regiao_adm = " . $idRegiao . $crlf .
            ")";
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            //echo $SQL.'<br/><br/>';

            $SQL = "INSERT into sbpi.Cliente_Particular (sq_cliente, Pasta, Diretor, Mantenedora, Endereco, Telefone_1, Fax, Vencimento, Observacao,sbpi.Escola, Codinep, Telefone_2, Email_1, Email_2, Cnpj_Executora, Cnpj_Escola, Secretario, Cep, Aut_Hab_Secretario, Infantil, Fundamental, EJA, Medio, Distancia, Profissional, Regiao, mantenedora_endereco, url, situacao,primeiro_credenc) " . $crlf .
            "VALUES(" . $sq_cliente . ", " . $idPasta . ", " . $idDiretor . ", " . $idMantenedora . ", " . $idEndereco . ", " . $idTelefone_1 . ", " . $idFax . ", to_date(" . $idVencimento . ",'yyyy-mm-dd'), " . $idObservacao . ", " . $idEscola . ", " . $idCodinep . ", " . $idTelefone_2 . ", " . $idEmail_1 . ", " . $idEmail_2 . ", " . $idCnpjExecutora . ", " . $idCnpjEscola . ", " . $idSecretario . ", " . $idCep . ", " . $idAutHabSecretario . ", " . $idEinf . ", " . $idEf . ", " . $idEja . ", " . $idEm . ", " . $idEDA . ", " . $idEprof . ", " . $idRegiao . ", " . $idEndMantenedora . ", " . $SiteEscola . ", " . $situacao . ",to_date(" . $credenc . ",'yyyy-mm-dd'));";
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            //echo $SQL.'<br/><br/>';
          } else {

            $SQL = "SELECT sq_cliente from sbpi.Cliente_Particular WHERE idescola = " . $idEscola;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            foreach ($RS as $row) {
              $RS = $row;
              break;
            }

            $sq_cliente = f($RS, "sq_cliente");

            //Remove os arquivos vinculados à escola
            $SQL = "DELETE from sbpi.Cliente_Arquivo WHERE sq_cliente = " . $sq_cliente;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            $SQL = "update sbpi.Cliente SET " . $crlf .
            "       ds_cliente      = " . $idInstituicao . ", " . $crlf .
            "       ds_username     = " . $Login . ", " . $crlf .
            "       ds_senha_acesso = " . $Senha . ", " . $crlf .
            "       no_municipio    = (select substr(no_regiao, instr(' ',no_regiao)+1,500) from sbpi.Regiao_Administrativa where sq_regiao_adm = " . $idRegiao . "), " . $crlf .
            "       localizacao     = " . $idLocalizacao . ", " . $crlf .
            "       sq_regiao_adm   = " . $idRegiao . $crlf .
            "  WHERE sq_cliente = " . $sq_cliente;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            //echo $SQL.'<br/><br/>';

            $SQL = "update sbpi.Cliente_Particular SET " . $crlf .
            "       Pasta                = " . $idPasta . ", " . $crlf .
            "       Diretor              = " . $idDiretor . ", " . $crlf .
            "       Mantenedora          = " . $idMantenedora . ", " . $crlf .
            "       Endereco             = " . $idEndereco . ", " . $crlf .
            "       Telefone_1           = " . $idTelefone_1 . ", " . $crlf .
            "       Fax                  = " . $idFax . ", " . $crlf .
            "       Vencimento           = to_date(" . $idVencimento . ",'yyyy-mm-dd'), " . $crlf .
            "       Observacao           = " . $idObservacao . ", " . $crlf .
            "       Codinep              = " . $idCodinep . ", " . $crlf .
            "       Telefone_2           = " . $idTelefone_2 . ", " . $crlf .
            "       Email_1              = " . $idEmail_1 . ", " . $crlf .
            "       Email_2              = " . $idEmail_2 . ", " . $crlf .
            "       Cnpj_Executora       = " . $idCnpjExecutora . ", " . $crlf .
            "       Cnpj_Escola          = " . $idCnpjEscola . ", " . $crlf .
            "       Secretario           = " . $idSecretario . ", " . $crlf .
            "       Cep                  = " . $idCep . ", " . $crlf .
            "       Aut_Hab_Secretario   = " . $idAutHabSecretario . ", " . $crlf .
            "       Infantil             = " . $idEinf . ", " . $crlf .
            "       Fundamental          = " . $idEf . ", " . $crlf .
            "       EJA                  = " . $idEja . ", " . $crlf .
            "       Medio                = " . $idEm . ", " . $crlf .
            "       Distancia            = " . $idEDA . ", " . $crlf .
            "       Profissional         = " . $idEprof . ", " . $crlf .
            "       Regiao               = " . $idRegiao . ", " . $crlf .
            "       Mantenedora_Endereco = " . $idEndMantenedora . ", " . $crlf .
            "       URL                  = " . $SiteEscola . ", " . $crlf .
            "       Situacao             = " . $Situacao . ", " . $crlf .
            "       primeiro_credenc     = to_date(" . $credenc . ",'yyyy-mm-dd') " . $crlf .
            "  WHERE sq_cliente = " . $sq_cliente;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          }

          $SQL = "DELETE from sbpi.Especialidade_Cliente WHERE SQ_CLIENTE = " . $sq_cliente;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          $SQL = "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " . $crlf .
          "  from sbpi.Cliente                       a " . $crlf .
          "       inner join sbpi.Cliente_Particular b on (a.sq_cliente = b.sq_cliente), " . $crlf .
          "       sbpi.Especialidade                 c " . $crlf .
          " where a.sq_cliente = " . $sq_cliente . $crlf .
          "   and (b.infantil  = c.sq_especialidade or  " . $crlf .
          "        (b.infantil = 3 and c.sq_especialidade in (1,2)) " . $crlf .
          "       ) " . $crlf .
          "UNION " . $crlf .
          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " . $crlf .
          "  from sbpi.Cliente                       a " . $crlf .
          "       inner join sbpi.Cliente_Particular b on (a.sq_cliente = b.sq_cliente), " . $crlf .
          "       sbpi.Especialidade                 c " . $crlf .
          " where a.sq_cliente = " . $sq_cliente . $crlf .
          "   and ((b.fundamental <> 0 and c.sq_especialidade = 5) or  " . $crlf .
          "        (b.fundamental in (1,2,3,7,11) and c.sq_especialidade =6) " . $crlf .
          "       ) " . $crlf .
          "UNION " . $crlf .
          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " . $crlf .
          "  from sbpi.Cliente                       a " . $crlf .
          "       inner join sbpi.Cliente_Particular b on (a.sq_cliente = b.sq_cliente), " . $crlf .
          "       sbpi.Especialidade                 c " . $crlf .
          " where a.sq_cliente       = " . $sq_cliente . $crlf .
          "   and b.medio            = 1  " . $crlf .
          "   and c.sq_especialidade = 21 " . $crlf .
          "UNION " . $crlf .
          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " . $crlf .
          "  from sbpi.Cliente                       a " . $crlf .
          "       inner join sbpi.Cliente_Particular b on (a.sq_cliente = b.sq_cliente), " . $crlf .
          "       sbpi.Especialidade                 c " . $crlf .
          " where a.sq_cliente       = " . $sq_cliente . $crlf .
          "   and b.eja              <> 0  " . $crlf .
          "   and c.sq_especialidade = 22 " . $crlf .
          "UNION " . $crlf .
          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " . $crlf .
          "  from sbpi.Cliente                       a " . $crlf .
          "       inner join sbpi.Cliente_Particular b on (a.sq_cliente = b.sq_cliente), " . $crlf .
          "       sbpi.Especialidade                 c " . $crlf .
          " where a.sq_cliente       = " . $sq_cliente . $crlf .
          "   and b.profissional     = 1  " . $crlf .
          "   and c.sq_especialidade = 8 " . $crlf .
          "UNION " . $crlf .
          "select a.sq_cliente, c.sq_especialidade, c.ds_especialidade " . $crlf .
          "  from sbpi.Cliente                       a " . $crlf .
          "       inner join sbpi.Cliente_Particular b on (a.sq_cliente = b.sq_cliente), " . $crlf .
          "       sbpi.Especialidade                 c " . $crlf .
          " where a.sq_cliente       = " . $sq_cliente . $crlf .
          "   and b.distancia        = 1  " . $crlf .
          "   and c.sq_especialidade = 24";
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          foreach ($RS as $row) {
            $SQL = "INSERT into sbpi.Especialidade_Cliente (sq_cliente, sq_especialidade) " . $crlf .
            "VALUES (" . $row["sq_cliente"] . ", " . $row["sq_especialidade"] . ") ";
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            //echo $SQL . '<br><br>';
          }

        }

        ShowHTML('<br>&nbsp;&nbsp;&nbsp; - <b>Cursos</b>');
        foreach ($cursos as $w_campo) {
          $w_cont++;
          $Curso = $w_campo[0];
          $idCursos = $w_campo[1];
          ShowHTML('<br>&nbsp;&nbsp;&nbsp;Linha ' . $w_cont . ": " . str_replace("'", "", $Curso));

          $SQL = "INSERT into sbpi.Curso (sq_curso, ds_curso, idcurso) VALUES (" . $idCursos . ", " . $Curso . ", " . $idCursos . ")";
          //echo $SQL. '<br><br>';
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        }

        ShowHTML('<br>&nbsp;&nbsp;&nbsp; - <b>Arquivos</b>');
        foreach ($arquivos as $w_campo) {
          $w_cont++;
          $idEscola = $w_campo[0];
          $idNomeArquivo = $w_campo[1];
          $idarquivos = $w_campo[2];
          $descricao = $w_campo[3];
          ShowHTML('<br>&nbsp;&nbsp;&nbsp;Linha ' . $w_cont . ": id " . str_replace("'", "", $idArquivos));

          $SQL = "INSERT into sbpi.CLIENTE_ARQUIVO (SQ_ARQUIVO,sq_cliente,DT_ARQUIVO,DS_TITULO,DS_ARQUIVO,LN_ARQUIVO,ativo,IN_DESTINATARIO,NR_ORDEM) " . $crlf .
          "(SELECT sbpi.sq_arquivo.nextval," . $crlf .
          "        a.sq_cliente," . $crlf .
          "        sysdate," . $crlf .
          "        " . $descricao . "," . $crlf .
          "        " . $descricao . "," . $crlf .
          "        " . $idNomeArquivo . "," . $crlf .
          "        'S'," . $crlf .
          "        'A'," . $crlf .
          "        1" . $crlf .
          "   from sbpi.Cliente_Particular a " . $crlf .
          "  WHERE a.idEscola = " . $idEscola . $crlf .
          ")";
          //echo $SQL . '<br><br>';
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }
        ShowHTML('<br>&nbsp;&nbsp;&nbsp; - <b>Profissionais</b>');
        foreach ($profissionais as $w_campo) {
          $cont++;
          $idProfissionais = $w_campo[0];
          $idEscola = $w_campo[1];
          $Pasta = 0; //$w_campo[2];
          $Parecer = "'Não Inf'"; //$w_campo[3];
          $Portaria = "'Não Inf'"; //$w_campo[4];
          $Observacao = "'Não Inf'"; //$w_campo[5];
          $idCursos = $w_campo[2];

          ShowHTML('<br>&nbsp;&nbsp;&nbsp;Linha ' . $w_cont . ": id " . str_replace("'", "", $idProfissionais));

          $SQL = "INSERT into sbpi.Particular_Curso (sq_particular_curso, sq_cliente, sq_curso, pasta, parecer, portaria, observacao, idProfissional) " . $crlf .
          "(SELECT " . $idProfissionais . ", a.sq_cliente, b.sq_curso, " . $Pasta . ", " . $Parecer . ", " . $Portaria . ", " . $Observacao . ", " . $idProfissionais . $crlf .
          "   from sbpi.Cliente_Particular a, " . $crlf .
          "        sbpi.Curso b " . $crlf .
          "  WHERE a.idEscola = " . $idEscola . $crlf .
          "    and b.idCurso  = " . $idCursos . $crlf .
          ")";

          //echo $SQL . '<br><br>';
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }
        ShowHTML('<br>&nbsp;&nbsp;&nbsp; - <b>OS</b>');
        foreach ($os as $w_campo) {
          $w_cont++;
          $idOS = $w_campo[0];
          $idEscola = $w_campo[1];
          $Numero = $w_campo[2];
          $Data = str_replace("-00-00", "-01-01", $w_campo[3]);
          $Dodf = $w_campo[4];
          $PagDodf = $w_campo[5];
          $Observacao = $w_campo[6];
          $DataDodf = str_replace("-00-00", "-01-01", $w_campo[7]);

          if ($Data == "NULL" or $Data == "'0000-01-01'") {
            $Data = "NULL";
          }
          if ($DataDodf == "NULL" or $DataDodf == "'0000-01-01'") {
            $DataDodf = "NULL";
          }
          ShowHTML('<br>&nbsp;&nbsp;&nbsp;Linha ' . $w_cont . ": id " . str_replace("'", "", $idOS));

          $SQL = "INSERT into sbpi.Particular_OS (sq_particular_os, sq_cliente, numero, data, dodf, dodf_pagina, dodf_data, observacao) " . $crlf .
          "(SELECT " . $idOS . ", a.sq_cliente, " . $Numero . ", to_date(" . $Data . ",'yyyy-mm-dd'), " . $Dodf . ", " . $PagDodf . ", to_date(" . $DataDodf . ",'yyyy-mm-dd'), " . $Observacao . $crlf .
          "   from sbpi.Cliente_Particular a " . $crlf .
          "  WHERE a.idEscola = " . $idEscola . ")";

          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }
        ShowHTML('<br>&nbsp;&nbsp;&nbsp; - <b>Portarias</b>');
        foreach ($portarias as $w_campo) {
          $w_cont++;
          $idPortaria = $w_campo[0];
          $idEscola = $w_campo[1];
          $Numero = $w_campo[2];
          $Data = str_replace("-00-00", "-01-01", $w_campo[3]);
          $Dodf = $w_campo[4];
          $PagDodf = $w_campo[5];
          $Observacao = $w_campo[6];
          $DataDodf = str_replace("-00-00", "-01-01", $w_campo[7]);

          if ($Data == "NULL" or $Data == "'0000-01-01'") {
            $Data = "NULL";
          }
          if ($DataDodf == "NULL" or $DataDodf == "'0000-01-01'") {
            $DataDodf = "NULL";
          }
          ShowHTML('<br>&nbsp;&nbsp;&nbsp;Linha ' . $w_cont . ": id " . str_replace("'", "", $idPortaria));

          $SQL = "INSERT into sbpi.Particular_Portaria (sq_particular_portaria, sq_cliente, numero, data, dodf, dodf_pagina, dodf_data, observacao) " . $crlf .
          "(SELECT " . $idPortaria . ", a.sq_cliente, " . $Numero . ", to_date(" . $Data . ",'yyyy-mm-dd'), " . $Dodf . ", " . $PagDodf . ", to_date(" . $DataDodf . ",'yyyy-mm-dd'), " . $Observacao . $crlf .
          "   from sbpi.Cliente_Particular a " . $crlf .
          "  WHERE a.idEscola = " . $idEscola .
          ")";
          //echo $SQL .'<br><br>';
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }

        //'Encerra a transação

        ScriptOpen('JavaScript');
        ShowHTML("  alert('Atualização da rede particular executada com sucesso.');");
        ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '";');
        ScriptClose();
        Rodape();

        break;

      case 'DADOSESCOLA' :

        $SQL = "update sbpi.Cliente set " .
        "     sq_regiao_adm   = " . $_REQUEST["p_regiao"] . ", " .
        "     sq_tipo_cliente = " . $_REQUEST["p_tipo_escola"] . ", " .
        "     localizacao     = " . $_REQUEST["p_local"];
        If (nvl($_REQUEST['p_regional'], "") != "")
          $SQL = $SQL . "     , sq_cliente_pai  = " . $_REQUEST["p_regional"];
        Else
          $SQL = $SQL . "     , sq_cliente_pai  = null ";
        $SQL = $SQL . "   where sq_cliente  = " . $_REQUEST["p_escola"];
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        ScriptOpen("JavaScript");
        ShowHTML("  location.href='" . $w_pagina . "dadosescola&p_escola=" . $_REQUEST["p_escola"] . "&p_regional=" . $_REQUEST["p_regional"] . "&p_tipo_escola=" . $_REQUEST["p_tipo_escola"] . "';");
        ScriptClose();

        break;

      case 'CALEND_BASE' :
        $w_funcionalidade = 17;
        if ($O == 'I') {
          //Recupera o valor da próxima chave primária
          $SQL = "select sbpi.sq_ocorrencia.nextval chave from sbpi.Calendario_Base" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");

          //Insere o arquivo
          $SQL = " insert into sbpi.Calendario_Base " . $crlf .
          "    (sq_ocorrencia, in_abrangencia, dt_ocorrencia, ds_ocorrencia, sq_tipo_data) " . $crlf .
          " values ( " . $w_chave . ", 'T', " . $crlf .
          "     to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_ocorrencia"], 2)) . "','dd/mm/yyyy'), " . $crlf .
          "     '" . $_REQUEST["w_ds_ocorrencia"] . "', " . $crlf .
          "     '" . $_REQUEST["w_tipo"] . "' " . $crlf .
          " )" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          If ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $_REQUEST["w_sq_cliente"] . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf .
            "         1, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["username"]) . " - inclusão de data no calendário base.', " . $crlf .
            "         '" . str_replace($w_sql, "'", "''") . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }
        elseif ($O == "A") {
          $SQL = " update sbpi.Calendario_Base set " . $crlf .
          "     dt_ocorrencia  = to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_ocorrencia"], 2)) . "','dd/mm/yyyy'), " . $crlf .
          "     ds_ocorrencia  = '" . $_REQUEST["w_ds_ocorrencia"] . "', " . $crlf .
          "     sq_tipo_data   = '" . $_REQUEST["w_tipo"] . "' " . $crlf .
          "where sq_ocorrencia = " . $_REQUEST["w_chave"] . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          If ($_SESSION["username"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $_REQUEST["w_sq_cliente"] . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf .
            "         2, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - alteração de data no calendário base.', " . $crlf .
            "         '" . str_replace($w_sql, "'", "''") . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }
        elseif ($O == "E") {
          $SQL = " delete sbpi.Calendario_Base where sq_ocorrencia = " . $_REQUEST["w_chave"] . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          If ($_SESSION["USERNAME"] != "SBPI") { /*
                      //Grava o acesso na tabela de log
                      $w_sql = $SQL;
                      $SQL = 'insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) ' . $crlf .
                      '( select sbpi.sq_cliente_log.nextval, ' . $crlf .
                      '         ' . $_REQUEST["w_sq_cliente"] . ', ' . $crlf .
                      '         sysdate, ' . $crlf .
                      '         \'' . $_SERVER["REMOTE_ADDR"] . '\', ' . $crlf .
                      '         3, ' . $crlf .
                      '         \'Usuário ' . strtoupper($_SESSION["username"] . ' - exclusão de data no calendário base.\', ' . $crlf .
                      '         \'' . str_replace($w_sql,"'", "''") . '\', ' . $crlf .
                      '         ' . $w_funcionalidade . ' ' . $crlf .
                      '   from dual) ' . $crlf;
                      $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
                      */
          }
        }

        ScriptOpen("JavaScript");
        ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '&$O=L&CL=' . $CL . '";');
        ScriptClose();
        Break;

      Case 'CALEND_REDE' :
        $w_funcionalidade = 17;
        if ($O == 'I') {
          //Recupera o valor da próxima chave primária
          $SQL = "select sbpi.sq_ocorrencia.nextval chave from sbpi.Calendario_Cliente" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");

          //Insere o arquivo
          $SQL = " insert into sbpi.Calendario_Cliente " . $crlf .
          "    (sq_ocorrencia, sq_cliente, dt_ocorrencia, ds_ocorrencia, sq_tipo_data) " . $crlf .
          " values ( " . $w_chave . ", " . $crlf .
          "     " . $CL . ", " . $crlf .
          "     to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_ocorrencia"], 2)) . "','dd/mm/yyyy'), " . $crlf .
          "     '" . $_REQUEST["w_ds_ocorrencia"] . "', '" . $_REQUEST["w_tipo"] . "' " . $crlf .
          " )" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $CL . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         1, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - inclusão de data no calendário da rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        } else
          if ($O == "A") {
            $SQL = " update sbpi.Calendario_Cliente set " . $crlf .
            "     dt_ocorrencia  = to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_ocorrencia"], 2)) . "','dd/mm/yyyy'), " . $crlf .
            "     ds_ocorrencia  = '" . $_REQUEST["w_ds_ocorrencia"] . "', " . $crlf .
            "     sq_tipo_data  = '" . $_REQUEST["w_tipo"] . "' " . $crlf .
            "where sq_ocorrencia = " . $_REQUEST["w_chave"] . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            if ($_SESSION["USERNAME"] != "SBPI") {
              //Grava o acesso na tabela de log
              $w_sql = $SQL;
              $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
              "( select sbpi.sq_cliente_log.nextval, " . $crlf .
              "         " . $CL . ", " . $crlf .
              "         sysdate, " . $crlf .
              "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
              "         2, " . $crlf .
              "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - alteração de data no calendário da rede.', " . $crlf .
              "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
              "         " . $w_funcionalidade . " " . $crlf .
              "   from dual) " . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }
          } else
            if ($O == "E") {
              $SQL = " delete sbpi.Calendario_Cliente where sq_ocorrencia = " . $_REQUEST["w_chave"] . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

              if ($_SESSION["USERNAME"] != "SBPI") {
                //Grava o acesso na tabela de log
                $w_sql = $SQL;
                $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                "( select sbpi.sq_cliente_log.nextval, " . $crlf .
                "         " . $CL . ", " . $crlf .
                "         sysdate, " . $crlf .
                "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
                "         3, " . $crlf .
                "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - exclusão de data no calendário da rede.', " . $crlf .
                "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                "         " . $w_funcionalidade . " " . $crlf .
                "   from dual) " . $crlf;
                $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
              }
            }

        ScriptOpen('JavaScript');
        ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '&$O=L";');
        ScriptClose();
        Break;

      Case 'NOTICIAS' :

        $w_funcionalidade = 20;
        if ($O == 'I') {
          //Recupera o valor da próxima chave primária
          $SQL = "select sbpi.sq_noticia.nextval chave from dual" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");

          //Insere o arquivo
          $SQL = " insert into sbpi.Noticia_Cliente " . $crlf .
          "    (sq_noticia, sq_cliente, dt_noticia, ds_titulo, ds_noticia, ln_externo, ativo) " . $crlf .
          " values (" . $w_chave . ", " . $crlf;
          if ($_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI") {
            $SQL .= "0," . $crlf;
          } else {
            $SQL .= "     " . $CL . ", " . $crlf;
          }
          $SQL .= "     to_date('" .
          FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_noticia"], 2)) . "','dd/mm/yyyy'), " . $crlf .
          "     '" . $_REQUEST["w_ds_titulo"] . "', " . $crlf .
          "     '" . $_REQUEST["w_ds_noticia"] . "', " . $crlf .
          "     '" . $w_ln_externo . "', " . $crlf .
          "     '" . $_REQUEST["w_in_ativo"] . "' " . $crlf .
          " )" . $crlf;
                    echo $SQL;
          exit();
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf;
            if ($_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI") {
              $SQL .= "0," . $crlf;
            } else {
              $SQL .= "     " . $CL . ", " . $crlf;
            }
            $SQL .= "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         1, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - inclusão de notícia da rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }
        elseif ($O == "A") {
          $SQL = " update sbpi.Noticia_Cliente set " . $crlf .
          "     dt_noticia  = to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_noticia"], 2)) . "','dd/mm/yyyy'), " . $crlf .
          "     ds_titulo   = '" . $_REQUEST["w_ds_titulo"] . "', " . $crlf .
          "     ds_noticia  = '" . $_REQUEST["w_ds_noticia"] . "', " . $crlf .
          "     ln_externo  = '" . $_REQUEST["w_ln_externo"] . "', " . $crlf .          
          "     ativo       = '" . $_REQUEST["w_in_ativo"] . "' " . $crlf .
          "where sq_noticia = " . $_REQUEST["w_chave"] . $crlf;
          ;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf;
            if ($_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI") {
              $SQL .= "0," . $crlf;
            } else {
              $SQL .= "     " . $CL . ", " . $crlf;
            }
            $SQL .= "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         2, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - alteração de notícia da rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }
        elseif ($O == "E") {
          $SQL = " delete sbpi.Noticia_Cliente where sq_noticia = " . $_REQUEST["w_chave"] . $crlf;
          ;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf;
            if ($_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI") {
              $SQL .= "0," . $crlf;
            } else {
              $SQL .= "     " . $CL . ", " . $crlf;
            }
            $SQL .= "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         3, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - exclusão de notícia da rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }

        ScriptOpen('JavaScript');
        ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '&$O=L"');
        ScriptClose();
        break;

      case 'ADM' :
        If ($_REQUEST["w_arquivo"] == "Escola") {
          $SQL = "select count(*) from sbpi.Cliente_Admin";
        } Else {
          $SQL = "select count(*) from sbpi.Equipamento";
        }
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        If (count($RS[0]) == 0) {
          ScriptOpen("JavaScript");
          ShowHTML("  alert('Não há informações a serem exportadas para a opção selecionada!);");
          ShowHTML("  history.back ;");
          ScriptClose();
        } Else {
          If ($_REQUEST["w_arquivo"] == "Escola") {
            $w_Arq = "Escola_" . date('Ymd') . "_" . date('His') . ".csv";
          } Else {
            $w_Arq = "Equipamento_" . date('Ymd') . "_" . date('His') . ".csv";
          }
          header("Content-type: application/octet-stream");
          header("Content-Disposition: attachment; filename=\"$w_Arq\"");

          If ($_REQUEST["w_arquivo"] == "Escola") {
            showHTML("Estrutura de dados:");
            showHTML("cd_escola; Código da unidade de ensino");
            showHTML("nm_escola; Nome da unidade de ensino");
            showHTML("st_limpeza_terceirizada; S - Limpeza com pessoal terceirizado, N - Limpeza com pessoal próprio");
            showHTML("st_oferece_merenda; S - A escola oferece merenda, N - A escola não oferece merenda");
            showHTML("qt_banheiro; Quantidade de quadros magnéticos informada pela unidade de ensino");
            showHTML("cd_tipo_equipamento; Código do tipo de equipamento");
            showHTML("nm_tipo_equipamento; Nome do tipo de equipamento");
            showHTML("st_tipo_equipamento; Situação do tipo de equipamento: S-Ativo, N-Inativo");
            showHTML("cd_equipamento; código do equipamento");
            showHTML("nm_equipamento; Nome do equipamento");
            showHTML("st_equipamento; Situação do equipamento: S-Ativo, N-Inativo");
            showHTML("qt_equipamento; Quantidade existente do equipamento");
            showHTML("");
            showHTML("Observações:");
            showHTML("1 - A lista está ordenada pelo nome da escola (1), pelo nome do tipo de equipamento (2) e pelo nome do equipamento (3);");
            showHTML("2 - As 5 primeiras colunas são duplicadas para cada um dos equipamentos da unidade de ensino;");
            showHTML("3 - Equipamentos ou tipos de equipamentos inativos não aparecem para a unidade de ensino informar sua quantidade.");
            showHTML("");
            showHTML("");

            //' Recupera os dados a serem exportados
            $SQL = "select f.sq_siscol cd_escola, e.ds_cliente nm_escola, " .
            "       d.limpeza_terceirizada st_limpeza_terceirizada, d.merenda_terceirizada st_oferece_merenda,  " .
            "       d.banheiro qt_banheiro, " .
            "       a.codigo cd_tipo_equipamento, a.nome nm_tipo_equipamento, a.ativo st_tipo_equipamento, " .
            "       b.codigo cd_equipamento,      b.nome nm_equipamento,      b.ativo st_equipamento,  " .
            "       c.quantidade qt_equipamento " .
            "  from sbpi.Cliente                                   e " .
            "       inner            join sbpi.Cliente_Site        f on (e.sq_cliente          = f.sq_cliente) " .
            "       inner            join sbpi.Cliente_Admin       d on (e.sq_cliente          = d.sq_cliente) " .
            "         left outer     join sbpi.Cliente_Equipamento c on (d.sq_cliente          = c.sq_cliente) " .
            "           left outer   join sbpi.Equipamento         b on (c.sq_equipamento      = b.sq_equipamento) " .
            "             left outer join sbpi.Tipo_Equipamento    a on (b.sq_tipo_equipamento = a.sq_tipo_equipamento) " .
            " order by e.ds_cliente, a.nome, b.nome ";
          } Else {
            showHTML("Estrutura de dados:");
            showHTML("cd_tipo_equipamento; Código do tipo de equipamento");
            showHTML("nm_tipo_equipamento; Nome do tipo de equipamento");
            showHTML("st_tipo_equipamento; Situação do tipo de equipamento: S-Ativo, N-Inativo");
            showHTML("cd_equipamento; código do equipamento");
            showHTML("nm_equipamento; Nome do equipamento");
            showHTML("st_equipamento; Situação do equipamento: S-Ativo, N-Inativo");
            showHTML("");
            showHTML("Observações:");
            showHTML("1 - A lista está ordenada pelo pelo nome do tipo de equipamento (1) e pelo nome do equipamento (2);");
            showHTML("2 - Equipamentos ou tipos de equipamentos inativos não aparecem para a unidade de ensino informar sua quantidade.");
            showHTML("");
            showHTML("");

            //' Recupera os dados a serem exportados
            $SQL = "select a.codigo cd_tipo_equipamento, a.nome nm_tipo_equipamento, a.ativo st_tipo_equipamento, " .
            "       b.codigo cd_equipamento,      b.nome nm_equipamento,      b.ativo st_equipamento " .
            "  from sbpi.Tipo_Equipamento a " .
            "       inner join sbpi.Equipamento b on (a.sq_tipo_equipamento = b.sq_tipo_equipamento) " .
            "order by a.nome, b.nome ";
          }
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          $primeiro = true;
          foreach ($RS as $row) {
            $col = array_keys($row);
            if ($primeiro) {
              $primeiro = false;
              showHTML(implode(';', $col));
            }
            $linha = "";
            foreach ($col as $colname) {
              $linha .= $row[$colname] . ';';
            }
            showHTML($linha);
          }
        }

        break;

      case 'MSGALUNOS' :
        $w_funcionalidade = 19;
        if ($O == 'I') {
          //Recupera o valor da próxima chave primária
          $SQL = "select sbpi.sq_mensagem.nextval chave from sbpi.Mensagem_Aluno" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");

          //Insere o arquivo
          $SQL = " insert into sbpi.Mensagem_Aluno " . $crlf .
          "    (sq_mensagem, sq_aluno, dt_mensagem, ds_mensagem, in_lida) " . $crlf .
          " values ( " . $w_chave . ", " . $crlf .
          "     " . $_REQUEST["w_sq_aluno"] . ", " . $crlf .
          "     to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_mensagem"], 2)) . "','dd/mm/yyyy'), " . $crlf .
          "     '" . $_REQUEST["w_ds_mensagem"] . "', " . $crlf .
          "     'N' " . $crlf .
          " )" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $CL . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         1, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - inclusão de mensagem para aluno da rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }
        elseif ($O == "A") {
          $SQL = " update sbpi.Mensagem_Aluno set " . $crlf .
          "     dt_mensagem  = to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_mensagem"], 2)) . "','dd/mm/yyyy'), " . $crlf .
          "     ds_mensagem  = '" . $_REQUEST["w_ds_mensagem"] . "' " . $crlf .
          "where sq_mensagem = " . $_REQUEST["w_chave"] . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $CL . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         2, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - alteração de mensagem para aluno da rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }
        elseif ($O == "E") {
          $SQL = " delete sbpi.Mensagem_Aluno where sq_mensagem = " . $_REQUEST["w_chave"] . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $CL . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         3, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - exclusão de mensagem para aluno da rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }

        ScriptOpen('JavaScript');
        ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '&$O=L"');
        ScriptClose();
        Break;

      Case 'ARQUIVOS' :

        $w_funcionalidade = 16;
        $SQL = "select a.ds_username, b.ds_username pai " . $crlf .
        "  from sbpi.Cliente a, " . $crlf .
        "       sbpi.CLiente b " . $crlf .
        " where b.sq_cliente_pai is null " . $crlf .
        "   and a.sq_cliente = " . $CL;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        $w_diretorio = $conFilePhysical;

        if ($O == "I") {

          if ($_FILES["w_ln_arquivo"]['size'] == 0) {
            ScriptOpen('JavaScript');
            ShowHTML("  alert('Informe um arquivo!');");
            ShowHTML("  history.back(1);");
            ScriptClose();
            exit ();
          }

          //Verifica se o nome desse arquivo já existe

          $SQL = "select count(*) existe from sbpi.Cliente_Arquivo where ln_arquivo = '" . $_FILES["w_ln_arquivo"]['name'] . "'" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          if (intval($RS[0]["existe"]) > 0) {
            ScriptOpen('JavaScript');
            ShowHTML('  alert(\'Nome do arquivo já existe!\');');
            ShowHTML('  history.back(1);');
            ScriptClose();
            exit ();
          }

          move_uploaded_file($_FILES["w_ln_arquivo"]['tmp_name'], $w_diretorio . '/' . $_FILES["w_ln_arquivo"]['name']);
          $w_imagem = $_FILES["w_ln_arquivo"]['name'];

          //Recupera o valor da próxima chave primária
          $SQL = "select sbpi.sq_arquivo.nextval chave from sbpi.Cliente_Arquivo" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");

          //Insere o arquivo
          $SQL = " insert into sbpi.Cliente_Arquivo " . $crlf .
          "    (sq_arquivo, sq_CLiente, dt_arquivo, ds_arquivo, ln_arquivo, " . $crlf .
          "     ativo,   in_destinatario, nr_ordem,   ds_titulo) " . $crlf .
          " values ( " . $w_chave . ", " . $crlf .
          "     " . $CL . ", " . $crlf .
          "     to_date('" . FormataDataEdicao(FormatDateTime(time(), 2)) . "','dd/mm/yyyy'), " . $crlf .
          "     '" . $_REQUEST["w_ds_arquivo"] . "', " . $crlf .
          "     '" . $w_imagem . "', " . $crlf .
          "     '" . $_REQUEST["w_in_ativo"] . "', " . $crlf .
          "     '" . $_REQUEST["w_in_destinatario"] . "', " . $crlf .
          "     '" . $_REQUEST["w_nr_ordem"] . "', " . $crlf .
          "     '" . $_REQUEST["w_ds_titulo"] . "' " . $crlf .
          " )" . $crlf;

          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          //Grava o acesso na tabela de log
          $w_sql = $SQL;
          If ($_SESSION['USERNAME'] != 'SBPI') {
            $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $CL . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         1, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - inclusão de arquivo.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }

        } else
          if ($O == "A") {
            //Verifica se o nome desse arquivo já existe
            $SQL = "select count(*) existe from sbpi.Cliente_Arquivo where ln_arquivo = '" . $_FILES["w_ln_arquivo"]['name'] . "' and sq_arquivo <> " . $_REQUEST['w_chave'] . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            if (intval($RS[0]["existe"]) > 0) {
              ScriptOpen('JavaScript');
              ShowHTML("  alert('Nome do arquivo já existe!');");
              ShowHTML("  history.back(1);");
              ScriptClose();
              exit ();
            }

            //Remove o arquivo físico
            $SQL = "select ln_arquivo arquivo from sbpi.Cliente_Arquivo where sq_arquivo = " . $_REQUEST['w_chave'];
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            foreach ($RS as $row) {
              $RS = $row;
              break;
            }
            $w_arquivo = f($RS, "arquivo");

            $SQL = " update sbpi.Cliente_Arquivo set " . $crlf .
            "     ds_titulo       = '" . $_REQUEST["w_ds_titulo"] . "', " . $crlf .
            "     dt_arquivo      = to_date('" . FormataDataEdicao(FormatDateTime(time(), 2)) . "','dd/mm/yyyy'), " . $crlf .
            "     ds_arquivo      = '" . $_REQUEST["w_ds_arquivo"] . "', " . $crlf .
            "     ativo           = '" . $_REQUEST["w_in_ativo"] . "', " . $crlf .
            "     in_destinatario = '" . $_REQUEST["w_in_destinatario"] . "', " . $crlf .
            "     nr_ordem        = '" . $_REQUEST["w_nr_ordem"] . "' " . $crlf .
            " where sq_arquivo = " . $_REQUEST["w_chave"] . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            If ($_SESSION['USERNAME'] != 'SBPI') {
              $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
              "( select sbpi.sq_cliente_log.nextval, " . $crlf .
              "         " . $CL . ", " . $crlf .
              "         sysdate, " . $crlf .
              "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
              "         2, " . $crlf .
              "         'Alteração de arquivo.', " . $crlf .
              "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
              "         " . $w_funcionalidade . " " . $crlf .
              "   from dual) " . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }

            if ($_FILES["w_ln_arquivo"]['size'] > 0) {
              //Remove o arquivo físico
              unlink($w_diretorio . '/' . $w_arquivo);

              move_uploaded_file($_FILES["w_ln_arquivo"]['tmp_name'], $w_diretorio . '/' . $_FILES["w_ln_arquivo"]['name']);
              $w_imagem = $_FILES["w_ln_arquivo"]['name'];

              $SQL = " update sbpi.Cliente_Arquivo set " . $crlf .
              "     ln_arquivo      = '" . $w_imagem . "' " . $crlf .
              "where sq_arquivo = " . $_REQUEST["w_chave"] . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }

          } else
            if ($O == "E") {
              //Remove o arquivo físico
              $SQL = "select ln_arquivo from sbpi.Cliente_Arquivo where sq_arquivo = " . $_REQUEST["w_chave"];
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

              unlink($w_diretorio . '/' . $RS[0]["ln_arquivo"]);

              $SQL = " delete sbpi.Cliente_Arquivo where sq_arquivo = " . $_REQUEST["w_chave"] . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

              //Grava o acesso na tabela de log
              $w_sql = $SQL;
              If ($_SESSION['USERNAME'] != 'SBPI') {
                $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                "( select sbpi.sq_cliente_log.nextval, " . $crlf .
                "         " . $CL . ", " . $crlf .
                "         sysdate, " . $crlf .
                "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
                "         3, " . $crlf .
                "         'ExCLusão de arquivo.', " . $crlf .
                "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                "         " . $w_funcionalidade . " " . $crlf .
                "   from dual) " . $crlf;
                $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
              }
            }

        $SQL = "update sbpi.Cliente set dt_alteracao = sysdate where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        ScriptOpen('JavaScript');

        ShowHTML('  location.href=\'' . $w_dir . $w_pagina . $SG . '\';');
        ScriptClose();
        Break;

      Case 'MODALIDADES' :

        $w_funcionalidade = 18;
        if ($O == 'I') {
          //Recupera o valor da próxima chave primária
          $SQL = "select sbpi.sq_especialidade.nextval chave from sbpi.Especialidade" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");

          //Insere o arquivo
          $SQL = " insert into sbpi.Especialidade " . $crlf .
          "    (sq_especialidade, ds_especialidade, nr_ordem) " . $crlf .
          " values ( " . $w_chave . ", " . $crlf .
          "     '" . $_REQUEST["w_ds_especialidade"] . "', " . $crlf .
          "     '" . $_REQUEST["w_nr_ordem"] . "' " . $crlf .
          " )" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $CL . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         1, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - inclusão de modalidade de ensino na rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }
        elseif ($O == "A") {
          $SQL = " update sbpi.Especialidade set " . $crlf .
          "     ds_especialidade  = '" . $_REQUEST["w_ds_especialidade"] . "', " . $crlf .
          "             nr_ordem  = '" . $_REQUEST["w_nr_ordem"] . "' " . $crlf .
          "where sq_especialidade = " . $_REQUEST["w_chave"] . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $CL . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         2, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - alteração de modalidade de ensino na rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }
        elseif ($O == "E") {
          $SQL = " delete sbpi.Especialidade_Cliente where sq_especialidade = " . $_REQUEST["w_chave"] . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $CL . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         3, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - exclusão de modalidade de ensino na rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }

          $SQL = " delete sbpi.Especialidade where sq_especialidade = " . $_REQUEST["w_chave"] . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          if ($_SESSION["USERNAME"] != "SBPI") {
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
            "( select sbpi.sq_cliente_log.nextval, " . $crlf .
            "         " . $CL . ", " . $crlf .
            "         sysdate, " . $crlf .
            "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
            "         3, " . $crlf .
            "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - exclusão de modalidade de ensino na rede.', " . $crlf .
            "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
            "         " . $w_funcionalidade . " " . $crlf .
            "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        }

        ScriptOpen('JavaScript');
        ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '&$O=L"');
        ScriptClose();
        Break;

      Case "NEWSLETTER" :
        if ($O == 'I') {

          $SQL = 'select sq_newsletter as chave from sbpi.newsletter from dual';
          //Recupera o valor da próxima chave primária
          $SQL = "select sbpi.sq_newsletter.nextval chave from sbpi.newsletter" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");
          //Insere o registro
          $SQL = "insert into sbpi.Newsletter (sq_newsletter, sq_cliente, nome, email, tipo, envia_mail, data_inclusao, data_alteracao) " . $crlf .
          "  values (" . $w_chave . ", 0," . $crlf .
          "          '" . trim($_REQUEST["w_nome"]) . "', " . $crlf .
          "          '" . trim($_REQUEST["w_email"]) . "', " . $crlf .
          "          '" . $_REQUEST["w_tipo"] . "', " . $crlf .
          "          '" . substr($_REQUEST["w_envia_mail"], 0, 1) . "', " . $crlf .
          "          sysdate, " . $crlf .
          "          null " . $crlf .
          "         )" . $crlf;
        }
        elseif ($O == "A") {
          $SQL = "update sbpi.Newsletter set " . $crlf .
          "   envia_mail     = '" . substr($_REQUEST["w_envia_mail"], 0, 1) . "', " . $crlf .
          "   nome           = '" . trim($_REQUEST["w_nome"]) . "', " . $crlf .
          "   email          = '" . trim($_REQUEST["w_email"]) . "', " . $crlf .
          "   tipo           = '" . $_REQUEST["w_tipo"] . "', " . $crlf .
          "   data_alteracao = sysdate " . $crlf .
          "where sq_newsletter = " . $_REQUEST["w_chave"] . $crlf;
        }
        elseif ($O == "E") {
          $SQL = " delete sbpi.Newsletter where sq_newsletter = " . $_REQUEST["w_chave"] . $crlf;
        }
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        ScriptOpen('JavaScript');
        ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '&$O=L"');
        ScriptClose();
        Break;

      Case "TIPOCLIENTE" :

        if ($O == 'I') {
          $SQL = "select sbpi.sq_tipo_cliente.nextval chave from sbpi.Tipo_Cliente" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");

          //Insere o registro
          $SQL = "insert into sbpi.Tipo_Cliente (sq_tipo_cliente, ds_tipo_cliente, ds_registro, tipo) " . $crlf .
          " values ( " . $w_chave . " ," . $crlf .
          "         '" . $_REQUEST["w_nome"] . "'," . $crlf .
          "              null                  ," . $crlf .
          "          " . $_REQUEST["w_tipo"] . "  " . $crlf .
          "         )" . $crlf;
        }
        elseif ($O == "A") {
          $SQL = "update sbpi.Tipo_Cliente set " . $crlf .
          "   ds_tipo_Cliente    = '" . ($_REQUEST["w_nome"]) . "'," . $crlf .
          "   ds_registro        =     null                     ," . $crlf .
          "   tipo               = " . $_REQUEST["w_tipo"] . $crlf .
          "where sq_tipo_cliente = " . $_REQUEST["w_chave"] . $crlf;
        }
        elseif ($O == "E") {
          $SQL = " delete sbpi.Tipo_Cliente where sq_tipo_cliente = " . $_REQUEST["w_chave"] . $crlf;
        }
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        ScriptOpen('JavaScript');
        ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '&$O=L"');
        ScriptClose();
        Break;

      Case 'SENHA' :
        $w_funcionalidade = 15;
        $SQL = "update sbpi.Cliente set " . $crlf .
        "   ds_senha_acesso = '" . trim($_REQUEST["w_ds_senha_acesso"]) . "' " . $crlf;
        $SQL .= "where sq_cliente =  " .
        $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        if ($_SESSION["USERNAME"] != "SBPI") {
          $w_sql = $SQL;
          ;
          $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
          "( select sbpi.sq_cliente_log.nextval, " . $crlf .
          "         " . $CL . ", " . $crlf .
          "         sysdate, " . $crlf .
          "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
          "         2, " . $crlf .
          "         'Usuário " . strtoupper($_SESSION["USERNAME"]) . " - atualização da senha de acesso.', " . $crlf .
          "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
          "         " . $w_funcionalidade . " " . $crlf .
          "   from dual) " . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }

        ScriptOpen('JavaScript');
        ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '&$O=L"');
        ScriptClose();

      Case "ESCPARTHOMOLOG" :
        if ($_REQUEST["p_escola_particular"] > '') {
          for ($i = 0; $i < count($_REQUEST["w_chave"]); $i++) {
            $SQL = "update sbpi.Particular_Calendario SET homologado = '" . $_REQUEST["w_homologado"][$i] . "' , ultima_homologacao = sysdate WHERE sq_particular_calendario = " . $_REQUEST["w_chave"][$i];
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            ScriptOpen('JavaScript');
            ShowHTML('location.href="' . $w_dir . $w_pagina . $SG . '&$O=L"');
            ScriptClose();
          }
        } else {
          ScriptOpen('JavaScript');
          ShowHTML('alert(\'Selecione, primeiramente, uma escola. Em seguida, seu(s) calendário(s).\')');
          ShowHTML(" history.back(1);");
          ScriptClose();
        }
    }
  }

  // =========================================================================
  // Monta a tela de Pesquisa
  // -------------------------------------------------------------------------
  function escPart() {
    extract($GLOBALS);
    global $w_Disabled;
    //print_r($_REQUEST);

    $p_tipo = 'H';
    $p_regiao = $_REQUEST["p_regiao"];
    $p_campos = $_REQUEST["p_campos"];
    if ($p_tipo == "W") { /*
          Response.ContentType = "application/msword"
          HeaderWord p_layout
          ShowHTML ('<TABLE WIDTH="100%" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR="#000000">SIGE-WEB<TD ALIGN="RIGHT"><B><FONT SIZE=4 COLOR="#000000">"
          ShowHTML ('Consulta a escolas"
          ShowHTML ('</FONT><TR><TD ALIGN="RIGHT"><B><FONT SIZE=2 COLOR="#000000">" & DataHora() & "</B></TD></TR>"
          ShowHTML ('</FONT></B></TD></TR></TABLE>"
          ShowHTML ('<HR>"*/
    } else {
      Cabecalho();
      ScriptOpen('JavaScript');
      ValidateOpen('Validacao');
      ValidateClose();
      ScriptClose();
      ShowHTML('<HEAD>');
      ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
      ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
      ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
      ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
      ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
      ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
      ShowHTML('</HEAD>');
      if ($_REQUEST["pesquisa"] > '') {
        BodyOpen(' onLoad="location.href=\'#lista\'"');
      } else {
        BodyOpen('onLoad=\'document.focus()\';');
      }
    }
    ShowHTML('<B><FONT COLOR="#000000">' . $w_tp . '</FONT></B>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
    ShowHTML('    <table width="95%" border="0">');
    if ($p_tipo == "H") {
      //Showhtml ('<FORM ACTION="controle.php" name="Form" METHOD="POST">');
      AbreForm('Form', $w_dir . $w_pagina . $par, 'POST', 'return(Validacao(this));', null);
      ShowHTML('<INPUT TYPE="HIDDEN" NAME="w_ew" VALUE="ESCPART">');
      ShowHTML('<INPUT TYPE="HIDDEN" NAME="CL" VALUE="" & CL &  "">');
      ShowHTML('<INPUT TYPE="HIDDEN" NAME="pesquisa" VALUE="X">');
      ShowHTML('<input type="Hidden" name="P3" value="1">');
      ShowHTML('<input type="Hidden" name="P4" value="15">');
      ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
      ShowHTML('    <table width="100%" border="0">');
      ShowHTML('          <TR><TD valign="top"><table border=0 width="100%" cellpadding=0 cellspacing=0>');
    } else {
      ShowHTML('<tr><td><div align="justify"><font size=1><b>Filtro:</b><ul>');
    }
    if ($p_tipo == "H") {
      ShowHTML('          <tr valign="top"><td>');
      SelecaoRegiaoAdm('Região a<u>d</u>ministrativa:', 'D', 'Indique a região administrativa.', $p_regiao, 'p_regiao', "p_regiao", null, null);
    }
    elseif (nvl($p_regiao, 0) > 0) {
      $SQL = "SELECT  a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " . $crlf .
      "  FROM  sbpi.CLIENTE a " . $crlf .
      " WHERE  a.sq_cliente = " . $p_regiao . $crlf .
      "ORDER BY a.ds_cliente ";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      ShowHTML('          <li><b>Escolas da ' . f($RS, "ds_cliente") . '</b>');
    }
    if ($p_tipo == "H") {
      ShowHTML('  <TR><TD><TD><br>');
      if ($_REQUEST["C"] > '') {
        ShowHTML('          <input type="checkbox" name="C" value="X" CLASS="BTM" checked> Exibir apenas escolas com calendário(s) cadastrado(s) ');
      } else {
        ShowHTML('          <input type="checkbox" name="C" value="X" CLASS="BTM"> Exibir apenas escolas com calendário(s) cadastrado(s)  ');
      }
    }
    elseif ($_REQUEST["C"] > '') {
      ShowHTML('  <li><b>Apenas escolas com alunos carregados </b>');
    }

    ShowHTML('          </tr>');
    ShowHTML('          </table>');
    if ($p_tipo = "H") {
      ShowHTML('  <TR><TD colspan=2><b>Campos a serem exibidos');
      if ($p_layout = "PORTRAIT") {
        ShowHTML('          (<input type="radio" name="p_layout" value="PORTRAIT" CLASS="BTM" checked> Retrato <input type="radio" name="p_layout" value="LANDSCAPE" CLASS="BTM"> Paisagem)');
      } else {
        ShowHTML('          (<input type="radio" name="p_layout" value="PORTRAIT" CLASS="BTM"> Retrato <input type="radio" name="p_layout" value="LANDSCAPE" CLASS="BTM" checked> Paisagem)');
      }
      ShowHTML('     <table width="100%" border=0>');
      ShowHTML('       <tr valign="top">');
      if ($_SESSION["USERNAME"] == "SEDF" or $_SESSION["USERNAME"] == "SBPI" or Substr($_SESSION["USERNAME"], 0, 2) == "RE") {
        //if($_REQUEST['p_campos'] != '' and  in_array('username', $_REQUEST['p_campos'])){
        if ($_REQUEST['p_campos'] != '' and in_array('username', $_REQUEST['p_campos'])) {
          ShowHTML(' <td><font size=1><input type="checkbox" name="p_campos[]" value="username" CLASS="BTM" checked>Username');
        } else {
          ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="username" CLASS="BTM">Username');
        }
      }
      //if( strpos($p_campos,"senha") !== false ){
      if ($_REQUEST['p_campos'] != '' and in_array('senha', $_REQUEST['p_campos'])) {
        ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="senha" CLASS="BTM" checked>Senha');
      } else {
        ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="senha" CLASS="BTM">Senha');
      }
    }
    //if( strpos($p_campos,"alteracao") !== false ){
    if ($_REQUEST['p_campos'] != '' and in_array('acao', $_REQUEST['p_campos'])) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="alteracao" CLASS="BTM" checked>Última alteração');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="alteracao" CLASS="BTM">Última alteração');
    }
    //if( strpos($p_campos,"diretor") !== false ){
    if ($_REQUEST['p_campos'] != '' and in_array('diretor', $_REQUEST['p_campos'])) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="diretor" CLASS="BTM" checked>Diretor');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="diretor" CLASS="BTM">Diretor');
    }
    //if( strpos($p_campos,"secretario") !== false ){
    if ($_REQUEST['p_campos'] != '' and in_array('secretario', $_REQUEST['p_campos'])) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="secretario" CLASS="BTM" checked>Secretário');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="secretario" CLASS="BTM">Secretário');
    }
    //if( strpos($p_campos,"contato") !== false ){
    /*if($_REQUEST['p_campos'] != '' and  in_array('contato', $_REQUEST['p_campos'])){
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos[]" value="contato" CLASS="BTM" checked>Contato');
     } else {
     ShowHTML ('          <td><font size=1><input type="checkbox" name="p_campos[]" value="contato" CLASS="BTM">Contato');
     }*/
    ShowHTML('       <tr valign="top">');
    //if( strpos($p_campos,"mail") !== false ){
    if ($_REQUEST['p_campos'] != '' and in_array('mail', $_REQUEST['p_campos'])) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="mail" CLASS="BTM" checked>e-Mail');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="mail" CLASS="BTM">e-Mail');
    }
    //if( strpos($p_campos,"telefone") !== false ){
    if ($_REQUEST['p_campos'] != '' and in_array('telefone', $_REQUEST['p_campos'])) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="telefone" CLASS="BTM" checked>Telefone');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="telefone" CLASS="BTM">Telefone');
    }
    //if( strpos($p_campos,"endereco") !== false ){
    if ($_REQUEST['p_campos'] != '' and in_array('endereco', $_REQUEST['p_campos'])) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="endereco" CLASS="BTM" checked>Endereço');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="endereco" CLASS="BTM">Endereço');
    }
    //if( strpos($p_campos,"localizacao") !== false ){
    if ($_REQUEST['p_campos'] != '' and in_array('localizacao', $_REQUEST['p_campos'])) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="localizacao" CLASS="BTM" checked>Localização');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="localizacao" CLASS="BTM">Localização');
    }
    //if( strpos($p_campos,"cep") !== false ){
    if ($_REQUEST['p_campos'] != '' and in_array('cep', $_REQUEST['p_campos'])) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="cep" CLASS="BTM" checked>CEP');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="cep" CLASS="BTM">CEP');
    }
    ShowHTML('     </table>');
    ShowHTML('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000">');
    ShowHTML('      <tr><td align="center" colspan="3">');
    ShowHTML('            <input class="BTM" type="submit" name="Botao" value="Aplicar filtro">');
    // if( $_SESSION["USERNAME"] == "SBPI" ){
    // ShowHTML ('            <input class="BTM" type="button" name="Botao" onClick="location.href=\'"'.$dir.$w_pagina.'cadastroescola'.'&CL='.$CL.MontaFiltro("GET").'&w_ea=I\'; value="Nova escola">');
    // }
    ShowHTML('          </td>');
    ShowHTML('      </tr>');
    ShowHTML('    </table>');
    ShowHTML('    </TD>');
    ShowHTML('</tr>');
    if ($p_tipo = "H") {
      ShowHTML('</form>');
    }
    if ($_REQUEST["pesquisa"] > '') {
      $SQL = " SELECT DISTINCT a.sq_cliente, a.ds_cliente, a.ds_apelido, a.ln_internet, a.ds_username, a.ds_senha_acesso, a.no_municipio, a.sg_uf, a.dt_alteracao, d.diretor, d.secretario, d.telefone_1, d.fax, d.cep, d.endereco, d.email_1 " . $crlf .
      "   from sbpi.Cliente a " . $crlf .
      "        INNER JOIN sbpi.Tipo_Cliente b ON (a.sq_tipo_cliente = b.sq_tipo_cliente and b.tipo=4) ";
      if ($_REQUEST["C"] > '') {
        $SQL .= "      INNER JOIN sbpi.Calendario_cliente c ON (a.sq_cliente = c.sq_cliente) " . $crlf .
        "        INNER JOIN sbpi.Cliente_Particular d ON (a.sq_cliente = d.sq_cliente) " . $crlf;
      } else {
        $SQL .= $crlf .
        "        INNER JOIN sbpi.Cliente_Particular d ON (a.sq_cliente = d.sq_cliente) " . $crlf;
      }
      if ($p_regiao > '') {
        $SQL = $SQL . "and a.sq_regiao_adm = " . $p_regiao . $crlf .
        "        order by a.ds_cliente ";
      } else {
        $SQL = $SQL . $crlf .
        "        order by a.ds_cliente ";
      }
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      ShowHTML('<TR><TD valign="top"><br><table border=0 width="100%" cellpadding=0 cellspacing=0>');
      if (count($RS) > 0) {
        if ($p_tipo == "H") {
          if ($_REQUEST["P4"] > '') {
            //RS.PageSize = doubleval($_REQUEST["P4"])
          } else {
            //RS.PageSize = 15;
          }
          //rs.AbsolutePage = Nvl($_REQUEST["P3"),1)
        } else {
          //RS.PageSize = RS.RecordCount + 1
          //rs.AbsolutePage = 1
        }

        ShowHTML('<tr><td><td align="right"><b><font face=Verdana size=1>Registros encontrados: ' . count($RS) . '</font></b>');
        if ($p_Tipo = "H") {
          //      ShowHTML ('     &nbsp;&nbsp;<A TITLE="Clique aqui para gerar arquivo Word com a listagem abaixo" class="SS" href="#"  onClick="window.open(\'Controle.asp?p_tipo=W&w_ew=' . $w_ew . '&Q=' . $_REQUEST["Q"] . '&C=' . $_REQUEST["C"] . '&D=' . $_REQUEST["D"] . '&U=' . $_REQUEST["U"] . $w_especialidade . MontaFiltro("GET") . '\',\'MetaWord\',\'width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no\');">Gerar Word<IMG ALIGN="CENTER" border=0 SRC="img/word.gif"></A>');
        }
        ShowHTML('<tr><td><td>');
        ShowHTML('<table border="1" cellspacing="0" cellpadding="0" width="100%">');
        ShowHTML('<tr align="center" valign="top">');
        ShowHTML('    <td><font face="Verdana" size="1"><b>Escola</b></td>');
        if ($_SESSION["USERNAME"] == "SEDF" or SubStr($_SESSION["USERNAME"], 0, 2) == "RE" or $_SESSION["USERNAME"] == "SBPI") {
          //if( strpos($p_campos,"username") !== false ){
          if ($_REQUEST['p_campos'] != '' and in_array('username', $_REQUEST['p_campos'])) {
            ShowHTML('    <td><font face="Verdana" size="1"><b>Username</b></td>');
          }
          //if( strpos($p_campos,"senha") !== false ){
          if ($_REQUEST['p_campos'] != '' and in_array('senha', $_REQUEST['p_campos'])) {
            ShowHTML('    <td><font face="Verdana" size="1"><b>Senha</b></td>');
          }
        }
        //if( strpos($p_campos,"alteracao") !== false ){
        if ($_REQUEST['p_campos'] != '' and in_array('acao', $_REQUEST['p_campos'])) {
          ShowHTML('    <td><font face="Verdana" size="1"><b>Última alteração</b></td>');
        }
        //if( strpos($p_campos,"diretor") !== false ){
        if ($_REQUEST['p_campos'] != '' and in_array('diretor', $_REQUEST['p_campos'])) {

          ShowHTML('    <td><font face="Verdana" size="1"><b>Diretor</b></td>');
        }
        //if( strpos($p_campos,"secretario") !== false ){
        if ($_REQUEST['p_campos'] != '' and in_array('secretario', $_REQUEST['p_campos'])) {
          ShowHTML('    <td><font face="Verdana" size="1"><b>Secretario</b></td>');
        }
        //if( strpos($p_campos,"mail") !== false ){
        if ($_REQUEST['p_campos'] != '' and in_array('mail', $_REQUEST['p_campos'])) {
          ShowHTML('    <td><font face="Verdana" size="1"><b>e-Mail</b></td>');
        }
        //if( strpos($p_campos,"telefone") !== false ){
        if ($_REQUEST['p_campos'] != '' and in_array('telefone', $_REQUEST['p_campos'])) {
          ShowHTML('    <td><font face="Verdana" size="1"><b>Telefone</b></td>');
          ShowHTML('    <td><font face="Verdana" size="1"><b>Fax</b></td>');
        }
        //if( strpos($p_campos,"endereco") !== false ){
        if ($_REQUEST['p_campos'] != '' and in_array('endereco', $_REQUEST['p_campos'])) {
          ShowHTML('    <td><font face="Verdana" size="1"><b>Endereço</b></td>');
        }
        //if( strpos($p_campos,"localizacao") !== false ){
        if ($_REQUEST['p_campos'] != '' and in_array('localizacao', $_REQUEST['p_campos'])) {
          ShowHTML('    <td><font face="Verdana" size="1"><b>Localização</b></td>');
        }
        //if( strpos($p_campos,"cep") !== false ){
        if ($_REQUEST['p_campos'] != '' and in_array('cep', $_REQUEST['p_campos'])) {
          ShowHTML('    <td><font face="Verdana" size="1"><b>CEP</b></td>');
        }
        $w_cor = '#FDFDFD';

        foreach ($RS as $row) { // and doubleval(RS.AbsolutePage) = doubleval(Nvl($_REQUEST["P3"),RS.AbsolutePage))
          if ($w_cor == "#EFEFEF" or $w_cor == '') {
            $w_cor = "#FDFDFD";
          } else {
            $w_cor = "#EFEFEF";
          }
          ShowHTML('<tr valign="top" bgcolor="" & w_cor & "">');
          if ($p_tipo == "H") {
            if (f($row, "LN_INTERNET") > '') {
              ShowHTML('    <td><font face="Verdana" size="1"><a href="' . f($row, "LN_INTERNET") . '" target="_blank">' . f($row, "DS_CLIENTE") . '</a></b></font></td>');
            } else {
              ShowHTML('    <td><font face="Verdana" size="1"><b>' . f($row, "DS_CLIENTE") . '</b></font></td>');
            }
          } else {
            ShowHTML('    <td><font face="Verdana" size="1">' . f($row, "DS_CLIENTE") . '</font></td>');
          }
          if ($_SESSION["USERNAME"] == "SEDF" or SubStr($_SESSION["USERNAME"], 0, 2) == "RE" or $_SESSION["USERNAME"] == "SBPI") {
            // if( strpos($p_campos,"username") > 0 ){
            if ($_REQUEST['p_campos'] != '' and in_array('username', $_REQUEST['p_campos'])) {
              ShowHTML('    <td><font face="Verdana" size="1">' . f($row, "DS_USERNAME") . '</font></td>');
            }
            // if( strpos($p_campos,"senha") > 0 ){
            if ($_REQUEST['p_campos'] != '' and in_array('senha', $_REQUEST['p_campos'])) {
              ShowHTML('    <td align="center"><font face="Verdana" size="1">' . f($row, "DS_SENHA_ACESSO") . '</font></td>');
            }
          }
          //if( strpos($p_campos,"alteracao")   > 0 ){
          if ($_REQUEST['p_campos'] != '' and in_array('acao', $_REQUEST['p_campos'])) {
            ShowHTML('    <td align="center"><font face="Verdana" size="1">' . Nvl(FormataDataEdicao(f($row, "dt_alteracao")), "---") . '</font></td>');
          }
          // if( strpos($p_campos,"diretor")     > 0 ){
          if ($_REQUEST['p_campos'] != '' and in_array('diretor', $_REQUEST['p_campos'])) {
            ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "diretor"), "---") . '</td>');
          }
          // if( strpos($p_campos,"secretario")  > 0 ){
          if ($_REQUEST['p_campos'] != '' and in_array('secretario', $_REQUEST['p_campos'])) {
            ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "secretario"), "---") . '</td>');
          }
          // if( strpos($p_campos,"mail")        > 0 ){
          if ($_REQUEST['p_campos'] != '' and in_array('mail', $_REQUEST['p_campos'])) {
            ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "email_1"), "---") . '</td>');
          }
          // if( strpos($p_campos,"telefone")    > 0 ){
          if ($_REQUEST['p_campos'] != '' and in_array('telefone', $_REQUEST['p_campos'])) {
            ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "telefone_1"), "---") . '</td>');
            ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "fax"), "---") . '</td>');
          }
          // if( strpos($p_campos,"endereco")    > 0 ){
          if ($_REQUEST['p_campos'] != '' and in_array('endereco', $_REQUEST['p_campos'])) {
            ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "endereco"), "---") . '</td>');
          }
          // if( strpos($p_campos,"localizacao") > 0 ){
          if ($_REQUEST['p_campos'] != '' and in_array('localizacao', $_REQUEST['p_campos'])) {
            ShowHTML('    <td><font face="Verdana" size="1">' . f($row, "no_municipio") . '-' . f($row, "sg_uf") . '</font></td>');
          }
          //if( strpos($p_campos,"cep")         > 0 ){
          if ($_REQUEST['p_campos'] != '' and in_array('cep', $_REQUEST['p_campos'])) {
            ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "cep"), "---") . '</td>');
          }
          if ($p_tipo = "H") {
            ShowHTML('    <td><font face="Verdana" size="1">');
            if ($_SESSION["username"] == 'SBPI') {
              ShowHTML('       <A CLASS="SS" HREF="' . $w_dir . $w_pagina . $par . '&w_chave=' . f($row, "sq_cliente") . '&O=A' . MontaFiltro("GET") . '" Title="Alteração dos dados da escola!">Alt</A>');
            }
          }
        }

        ShowHTML('</table>');
        ShowHTML('<tr><td><td colspan="5" align="center"><hr>');

        if ($p_tipo = "H") {
          ShowHTML('<tr><td align="center" colspan=5>');
          //MontaBarra "Controle.asp?CL=" & CL & "&w_ew=" & w_ew & "&Q=" & $_REQUEST["Q") & "&C=" & $_REQUEST["C") & "&D=" & $_REQUEST["D") & "&U=" & $_REQUEST["U") & w_especialidade, doubleval(RS.PageCount), doubleval($_REQUEST["P3")), doubleval($_REQUEST["P4")), doubleval(RS.RecordCount)
          ShowHTML('</tr>');
        }

      } else {

        ShowHTML('<TR><TD><TD colspan="3"><p align="justify"><img src="img/ico_educacao.gif" width="16" height="16" border="0" align="center">&nbsp;<font size="2"><b>Nenhuma ocorrência encontrada para as opções acima.');
      }
    }
    ShowHTML('</TABLE>');

  }

  //Fim da Pesquisa de Escolas Particulares

  // =========================================================================
  // Monta a tela de Pesquisa
  // -------------------------------------------------------------------------
  function escolas() {
    extract($GLOBALS);
    global $w_Disabled;

    $p_tipo = 'H';
    $p_regiao = $_REQUEST["p_regiao"];

    if ($p_tipo == "W") {
      //Response.ContentType = "application/msword"
      // HeaderWord p_layout
      ShowHTML('<TABLE WIDTH="100%" BORDER=0><TR><TD ROWSPAN=2><FONT SIZE=4 COLOR="#000000">SIGE-WEB<TD ALIGN="RIGHT"><B><FONT SIZE=4 COLOR="#000000">');
      ShowHTML('Consulta a escolas');
      ShowHTML('</FONT><TR><TD ALIGN="RIGHT"><B><FONT SIZE=2 COLOR="#000000">" ' . DataHora() . ' "</B></TD></TR>');
      ShowHTML('</FONT></B></TD></TR></TABLE>');
      ShowHTML('<HR>');
    } else {
      Cabecalho();
      ShowHTML('<HEAD>');
      ShowHTML('</HEAD>');
      if ($_REQUEST["p_pesquisa"] > '') {
        BodyOpen(' onLoad="location.href=\'#lista\'"');
      } else {
        BodyOpen('onLoad=\'document.focus()\';');
      }
    }
    ShowHTML('<B><FONT COLOR="#000000">' . $w_tp . '</FONT></B>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
    ShowHTML('    <table width="95%" border="0">');
    if ($p_tipo == "H") {
      //Showhtml ('<FORM ACTION="controle.php" name="Form" METHOD="POST">');
      AbreForm('Form', $w_pagina . $par, 'POST', null, null);
      ShowHTML('<INPUT TYPE="HIDDEN" NAME="SG" VALUE="ESCOLAS">');
      ShowHTML('<INPUT TYPE="HIDDEN" NAME="p_pesquisa" VALUE="X">');
      // ShowHTML ('<input type="Hidden" name="P3" value="1">');
      // ShowHTML ('<input type="Hidden" name="P4" value="15">');
      ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
      ShowHTML('    <table width="100%" border="0">');
      ShowHTML('          <TR><TD valign="top"><table border=0 width="100%" cellpadding=0 cellspacing=0>');
    } else {
      ShowHTML('<tr><td><div align="justify"><font size=1><b>Filtro:</b><ul>');
    }
    if ($p_tipo == "H") {
      ShowHTML('          <tr valign="top"><td>');
      SelecaoRegional("<u>S</u>ubordinação:", "S", "Indique a subordinação da escola.", $p_regional, null, "p_regional", null, null);
    }
    elseif (nvl($p_regiao, 0) > 0) {
      $SQL = "SELECT  a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " . $crlf .
      "  FROM  sbpi.CLIENTE a " . $crlf .
      " WHERE  a.sq_cliente = " . $p_regiao . $crlf .
      "ORDER BY a.ds_cliente ";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      ShowHTML('          <li><b>Escolas da ' . f($RS, "ds_cliente") . '</b>');
    }
    $SQL = "SELECT * from sbpi.Tipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente" . $crlf;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    if ($p_tipo == "H") {
      ShowHTML('          <tr valign="top"><td><td><br><b>Tipo de instituição:</b><br><SELECT CLASS="STI" NAME="p_tipo_cliente">');
      if (Count($RS) > 1) {
        ShowHTML('          <option value="">Todos');
      }
      foreach ($RS as $row) {
        if (doubleval(nvl($row["sq_tipo_cliente"], 0)) == doubleval(nvl($_REQUEST["p_tipo_cliente"], 0))) {
          ShowHTML('          <option value="' . $row["sq_tipo_cliente"] . '" SELECTED>' . $row["ds_tipo_cliente"]);
        } else {
          ShowHTML('          <option value="' . $row["sq_tipo_cliente"] . '">' . $row["ds_tipo_cliente"]);
        }
      }
      ShowHTML('          </select>');
    }
    elseif (nvl($_REQUEST["p_tipo_cliente"], 0) > 0) {
      ShowHTML('          <li><b>Tipo de instituição: ');
      foreach ($RS as $row) {
        if (doubleval(nvl($row["sq_tipo_cliente"], 0)) == cDbl(nvl($_REQUEST["p_tipo_cliente"], 0))) {
          ShowHTML($row["ds_tipo_cliente"]);
        }
      }
      ShowHTML('</b>');
    }
    if ($p_tipo == "H") {
      ShowHTML('  <TR><TD><TD><br>');
      if ($_REQUEST["C"] > '') {
        ShowHTML('          <input type="checkbox" name="C" value="X" CLASS="BTM" checked> Exibir apenas escolas com calendário(s) cadastrado(s) ');
      } else {
        ShowHTML('          <input type="checkbox" name="C" value="X" CLASS="BTM"> Exibir apenas escolas com calendário(s) cadastrado(s)  ');
      }
    }
    elseif ($_REQUEST["C"] > '') {
      ShowHTML('  <li><b>Apenas escolas com alunos carregados </b>');
    }

    if ($p_tipo == "H") {
      ShowHTML('  <TR><TD><TD><br><b>Quanto às informações administrativas?</b><br>');
      if ($_REQUEST["E"] = "S") {
        ShowHTML('          <input type="radio" name="E" value="S" CLASS="BTM" checked> Listar apenas escolas que informaram<br>');
      } else {
        ShowHTML('          <input type="radio" name="E" value="S" CLASS="BTM"> Listar apenas escolas que informaram<br>');
      }
      if ($_REQUEST["E"] = "N") {
        ShowHTML('          <input type="radio" name="E" value="N" CLASS="BTM" checked> Listar apenas escolas que não informaram<br>');
      } else {
        ShowHTML('          <input type="radio" name="E" value="N" CLASS="BTM"> Listar apenas escolas que não informaram<br> ');
      }
      if ($_REQUEST["E"] = '') {
        ShowHTML('          <input type="radio" name="E" value=""  CLASS="BTM" checked> Tanto faz');
      } else {
        ShowHTML('          <input type="radio" name="E" value="" CLASS="BTM"> Tanto faz ');
      }
    }
    elseif ($_REQUEST["E"] > '') {
      ShowHTML('  <li><b>');
      if ($_REQUEST["E"] = "S") {
        ShowHTML('          Listar apenas escolas que informaram dados administrativos</b>');
      }
      if ($_REQUEST["E"] = "N") {
        ShowHTML('          Listar apenas escolas que não informaram dados administrativos</b> ');
      }
    }
    if ($p_tipo != 'W') {
      ShowHTML('          </table>');
    }
    $wCont = 0;
    $SQL1 = '';

    //Seleção de etapas/modalidades
    $SQL = "SELECT DISTINCT a.curso sq_especialidade, a.curso ds_especialidade, '1' nr_ordem, 'M' tp_especialidade " . $crlf .
    " from sbpi.Turma_Modalidade                 a " . $crlf .
    "      INNER join sbpi.Turma                 c ON (a.serie           = c.ds_serie) " . $crlf .
    "      INNER join sbpi.Cliente               d ON (c.sq_cliente = d.sq_cliente) " . $crlf .
    "UNION " . $crlf .
    "SELECT DISTINCT to_char(a.sq_especialidade) sq_especialidade, a.ds_especialidade,  " . $crlf .
    "       case a.tp_especialidade when 'J' then '1' else to_char(a.nr_ordem) end nr_ordem, " . $crlf .
    "       case a.tp_especialidade when 'J' then 'M' else a.tp_especialidade end tp_especialidade" . $crlf .
    " from sbpi.Especialidade a " . $crlf .
    "      INNER join sbpi.Especialidade_cliente c ON (a.sq_especialidade = c.sq_especialidade) " . $crlf .
    "      INNER join sbpi.Cliente               d ON (c.sq_cliente       = d.sq_cliente) " . $crlf .
    " where a.tp_especialidade <> 'M' " . $crlf .
    "ORDER BY nr_ordem, ds_especialidade " . $crlf;

    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    if ($p_tipo == 'H') {
      if (count($RS) > 0) {
        $wCont = 0;
        $wAtual = "";
        ShowHTML('          <TD valign="top"><table border="0" align="left" cellpadding=0 cellspacing=0>');
        foreach ($RS as $row) {
          if ($wAtual == '' or $wAtual != $row["tp_especialidade"]) {
            $wAtual = $row["tp_especialidade"];
            if ($wAtual == 'M') {
              ShowHTML('            <TR><TD colspan=2><b>Etapas/Modalidades de ensino:</b>');
            }
            elseif ($wAtual = "R") {
              ShowHTML('            <TR><TD colspan=2><b>Em Regime de Intercomplementaridade:</b>');
            } else {
              ShowHTML('            <TR><TD colspan=2><b>Outras:</b>');
            }
          }
          $wCont++;
          $marcado = 'N';

          if (is_array($_REQUEST['p_modalidade'])) {
            $SQL1 = '\'' . str_replace(',', '\',\'', implode(',', $_REQUEST['p_modalidade'])) . '\'';
          } else {
            $SQL1 = '\'' . str_replace(',', '\',\'', $p_modalidade) . '\'';
          }

          if (strpos($SQL1, '\'' . f($row, "sq_especialidade") . '\'') !== false) {
            $marcado = 'S';
          }

          if ($marcado == "S") {
            //ShowHTML chr(13) & "           <tr><td><input type=""checkbox" name=""p_modalidade" value=""" & RS("sq_especialidade") & """ checked><td><font size=1>" & RS("ds_especialidade")
            ShowHTML('           <tr><td><input type="checkbox" name="p_modalidade[]" value="' . $row["sq_especialidade"] . '" checked><td><font size=1>' . $row["ds_especialidade"]);
            $wIN = 1;
          } else {
            ShowHTML('           <tr><td><input type="checkbox" name="p_modalidade[]" value="' . $row["sq_especialidade"] . '"><td><font size=1>' . $row["ds_especialidade"]);
          }

          if (($wCont % 2) == 0) {
            $wCont = 0;
          }

        }
        //        $SQL1 = substr($SQL1,0,1);

      }
    }
    elseif (Nvl($_REQUEST["p_modalidade"], "") > '') {
      if (count($RS) > 0) {
        $wCont = 0;
        ShowHTML('          <li><b>Modalidades de ensino:</b><ul>');
        $SQL1 = "'" . str_replace(", ", "','", $_REQUEST["p_modalidade"]) . "'";

        foreach ($RS as $row) {
          For ($i = 0; $_REQUEST["p_modalidade"]; $i++) {
            if (strpos($SQL1, "'" . $row["sq_especialidade"] . "'") > 0) {
              ShowHTML('           <li><font size=1>' . $row["ds_especialidade"]);
            }
          }
          $wIN = 1;
        }
      }
    }
    ShowHTML('          </tr>');
    ShowHTML('          </table>');
    if ($p_tipo == 'H') {
      ShowHTML('  <TR><TD colspan=2><b>Campos a serem exibidos');
      if ($p_layout == "PORTRAIT") {
        ShowHTML('          (<input type="radio" name="p_layout" value="PORTRAIT" CLASS="BTM" checked> Retrato <input type="radio" name="p_layout" value="LANDSCAPE" CLASS="BTM"> Paisagem)');
      } else {
        ShowHTML('          (<input type="radio" name="p_layout" value="PORTRAIT" CLASS="BTM"> Retrato <input type="radio" name="p_layout" value="LANDSCAPE" CLASS="BTM" checked> Paisagem)');
      }
      ShowHTML('     <table width="100%" border=0>');
      ShowHTML('       <tr valign="top">');
      if ($_SESSION["USERNAME"] == "SEDF" or $_SESSION["USERNAME"] == "SBPI" or Substr($_SESSION["USERNAME"], 0, 2) == "RE") {

        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'username') !== false) {
          ShowHTML(' <td><font size=1><input type="checkbox" name="p_campos[]" value="username" CLASS="BTM" checked>Username');
        } else {
          ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="username" CLASS="BTM">Username');
        }
      }
      //if( strpos($p_campos,"senha") !== false ){
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'senha') !== false) {
        ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="senha" CLASS="BTM" checked>Senha');
      } else {
        ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="senha" CLASS="BTM">Senha');
      }
    }
    //if( strpos($p_campos,"alteracao") !== false ){
    if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'acao') !== false) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="alteracao" CLASS="BTM" checked>Última alteração');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="alteracao" CLASS="BTM">Última alteração');
    }
    //if( strpos($p_campos,"diretor") !== false ){
    if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'diretor') !== false) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="diretor" CLASS="BTM" checked>Diretor');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="diretor" CLASS="BTM">Diretor');
    }
    //if( strpos($p_campos,"secretario") !== false ){
    if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'secretario') !== false) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="secretario" CLASS="BTM" checked>Secretário');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="secretario" CLASS="BTM">Secretário');
    }
    //if( strpos($p_campos,"contato") !== false ){
    if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'contato') !== false) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="contato" CLASS="BTM" checked>Contato');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="contato" CLASS="BTM">Contato');
    }
    ShowHTML('       <tr valign="top">');
    //if( strpos($p_campos,"mail") !== false ){
    if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'mail') !== false) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="mail" CLASS="BTM" checked>e-Mail');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="mail" CLASS="BTM">e-Mail');
    }
    //if( strpos($p_campos,"telefone") !== false ){
    if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'telefone') !== false) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="telefone" CLASS="BTM" checked>Telefone');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="telefone" CLASS="BTM">Telefone');
    }
    //if( strpos($p_campos,"endereco") !== false ){
    if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'endereco') !== false) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="endereco" CLASS="BTM" checked>Endereço');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="endereco" CLASS="BTM">Endereço');
    }
    //if( strpos($p_campos,"localizacao") !== false ){
    if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'localizacao') !== false) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="localizacao" CLASS="BTM" checked>Localização');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="localizacao" CLASS="BTM">Localização');
    }
    //if( strpos($p_campos,"cep") !== false ){
    if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'cep') !== false) {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="cep" CLASS="BTM" checked>CEP');
    } else {
      ShowHTML('          <td><font size=1><input type="checkbox" name="p_campos[]" value="cep" CLASS="BTM">CEP');
    }
    ShowHTML('     </table>');
    ShowHTML('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000">');
    ShowHTML('      <tr><td align="center" colspan="3">');
    ShowHTML('            <input class="BTM" type="submit" name="Botao" value="Aplicar filtro">');
    if ($_SESSION["USERNAME"] == "SBPI") {
      ShowHTML('            <input class="BTM" type="button" name="Botao" onClick="location.href=\'' . $w_pagina . 'cadastroescola' . MontaFiltro("GET") . '&O=I\'"; value="Nova escola">');
    }
    ShowHTML('          </td>');
    ShowHTML('      </tr>');
    ShowHTML('    </table>');
    ShowHTML('    </TD>');
    ShowHTML('</tr>');
    if ($p_tipo == "H") {
      ShowHTML('</form>');
    }
    if ($_REQUEST["p_pesquisa"] > '') {
      $SQL = "SELECT DISTINCT b.ds_username, d.sq_cliente, d.ds_cliente, d.ds_apelido, d.ln_internet, " . $crlf .
      "       d.ds_username, d.ds_senha_acesso, d.no_municipio, d.sg_uf, d.dt_alteracao, " . $crlf .
      "       e.no_diretor, e.no_secretario, e.no_bairro, e.nr_cep, e.ds_logradouro, " . $crlf .
      "       f.ds_email_internet, f.no_contato_internet, nr_fone_internet, nr_fax_internet, " . $crlf .
      "       coalesce(h.existe,0) alunos, i.sq_cliente adm, d.ativo " . $crlf .
      "  from sbpi.Cliente                                 d " . $crlf .
      "       INNER      join sbpi.Cliente                 b ON (b.sq_cliente       = d.sq_cliente_pai) " . $crlf .
      "       LEFT OUTER join sbpi.Cliente_Dados           e ON (d.sq_cliente       = e.sq_cliente) " . $crlf .
      "       INNER      join sbpi.Cliente_Site            f ON (d.sq_cliente       = f.sq_cliente) " . $crlf .
      "       INNER      join sbpi.Tipo_Cliente            g ON (d.sq_tipo_cliente  = g.sq_tipo_cliente) " . $crlf .
      "       LEFT OUTER JOIN (select sq_cliente, count(*) existe " . $crlf .
      "                          from sbpi.Aluno " . $crlf .
      "                         group by sq_cliente " . $crlf .
      "                       )                          h on (f.sq_cliente  = h.sq_cliente) " . $crlf .
      "       LEFT OUTER join sbpi.Cliente_Admin           i on (d.sq_cliente       = i.sq_cliente) " . $crlf .
      " where 1 = 1 " . $crlf;
      if (Substr($_SESSION["USERNAME"], 0, 2) == "RE") {
        $SQL .= '   and b.ds_username = \'' . $_SESSION["USERNAME"] . '\'' . $crlf;
      }
      if ($_REQUEST["C"] > "") {
        $SQL .= '   and coalesce(h.existe,0) > 0 ' . $crlf;
      }
      if ($_REQUEST["D"] > "") {
        $SQL .= '   and d.dt_alteracao     is not null ' . $crlf;
      }
      if ($_REQUEST["E"] > "") {
        if ($_REQUEST["E"] == "S") {
          $SQL .= "   and i.sq_cliente       is not null " . $crlf;
        } else {
          $SQL .= "   and i.sq_cliente       is null " . $crlf;
        }
      }

      if ($_REQUEST["p_regional"] > '') {
        $SQL .= '    and d.sq_cliente_pai = ' . $_REQUEST["p_regional"] . $crlf;
      } else {
        $SQL .= '    and g.tipo = 3' . $crlf;
      }

      if ($_REQUEST["p_tipo_cliente"] > '') {
        $SQL .= '    and d.sq_tipo_cliente= ' . $_REQUEST["p_tipo_cliente"] . $crlf;
      }

      if (strlen(trim($SQL1)) > 3) {
        $SQL .= "    and (0 < (select count(*) from sbpi.Especialidade_Cliente where sq_cliente = d.sq_cliente and to_char(sq_especialidade) in (" . $SQL1 . ")) or " . $crlf .
        "         0 < (select count(*) from sbpi.Turma_Modalidade  w INNER join sbpi.Turma x ON (w.serie = x.ds_serie) where x.sq_cliente = d.sq_cliente and w.curso in (" . $SQL1 . ")) " . $crlf .
        "        ) " . $crlf;
      }
      $SQL .= 'ORDER BY d.ds_cliente ' . $crlf;

      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

      ShowHTML('<TR><TD valign="top"><br><table border=0 width="100%" cellpadding=0 cellspacing=0>');
      /*
       if( Count($RS)){

       if( $p_tipo == 'H' ){
       if( $_REQUEST["P4") > 0 ){ RS.PageSize = cDbl($_REQUEST["P4")) } else { RS.PageSize = 15 }
       rs.AbsolutePage = Nvl($_REQUEST["P3"),1)
       } else {
       RS.PageSize = RS.RecordCount + 1
       rs.AbsolutePage = 1
       }

       */
      ShowHTML('<tr><td><td align="right"><b><font face=Verdana size=1>Registros encontrados: ' . count($RS) . '</font></b>');
      if ($p_Tipo = "H") {
        //ShowHTML ('     &nbsp;&nbsp;<A TITLE="Clique aqui para gerar arquivo Word com a listagem abaixo" class="SS" href="#"  onClick="window.open(\'Controle.asp?p_tipo=W&w_ew=' . $w_ew . '&Q=' . $_REQUEST["Q"] . '&C=' . $_REQUEST["C"] . '&D=' . $_REQUEST["D"] . '&U=' . $_REQUEST["U"] . $w_especialidade . MontaFiltro("GET") . '\',\'MetaWord\',\'width=600, height=350, top=65, left=65, menubar=yes, scrollbars=yes, resizable=yes, status=no\');">Gerar Word<IMG ALIGN="CENTER" border=0 SRC="img/word.gif"></A>');
      }
      ShowHTML('<tr><td><td>');
      ShowHTML('<table border="1" cellspacing="0" cellpadding="0" width="100%">');
      ShowHTML('<tr align="center" valign="top">');
      ShowHTML('    <td><font face="Verdana" size="1"><b>Escola</b></td>');
      if ($_SESSION["USERNAME"] == "SEDF" or SubStr($_SESSION["USERNAME"], 0, 2) == "RE" or $_SESSION["USERNAME"] == "SBPI") {
        //if( strpos($p_campos,"username") !== false ){
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'username') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1"><b>Username</b></td>');
        }
        //if( strpos($p_campos,"senha") !== false ){
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'senha') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1"><b>Senha</b></td>');
        }
      }
      //if( strpos($p_campos,"alteracao") !== false ){
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'acao') !== false) {
        ShowHTML('    <td><font face="Verdana" size="1"><b>Última alteração</b></td>');
      }
      //if( strpos($p_campos,"diretor") !== false ){
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'diretor') !== false) {

        ShowHTML('    <td><font face="Verdana" size="1"><b>Diretor</b></td>');
      }
      //if( strpos($p_campos,"secretario") !== false ){
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'secretario') !== false) {
        ShowHTML('    <td><font face="Verdana" size="1"><b>Secretario</b></td>');
      }
      //if( strpos($p_campos,"mail") !== false ){
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'mail') !== false) {
        ShowHTML('    <td><font face="Verdana" size="1"><b>e-Mail</b></td>');
      }
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'contato') !== false) {
        ShowHTML('    <td><font face="Verdana" size="1"><b>Contato</b></td>');
      }
      //if( strpos($p_campos,"telefone") !== false ){
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'telefone') !== false) {
        ShowHTML('    <td><font face="Verdana" size="1"><b>Telefone</b></td>');
        ShowHTML('    <td><font face="Verdana" size="1"><b>Fax</b></td>');
      }
      //if( strpos($p_campos,"endereco") !== false ){
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'endereco') !== false) {
        ShowHTML('    <td><font face="Verdana" size="1"><b>Endereço</b></td>');
      }
      //if( strpos($p_campos,"localizacao") !== false ){
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'localizacao') !== false) {
        ShowHTML('    <td><font face="Verdana" size="1"><b>Localização</b></td>');
      }
      //if( strpos($p_campos,"cep") !== false ){
      if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'cep') !== false) {
        ShowHTML('    <td><font face="Verdana" size="1"><b>CEP</b></td>');
      }

      ShowHTML('    <td><font face="Verdana" size="1"><b>Outros</b></td>');
      $w_cor = '#FDFDFD';

      $RS1 = array_slice($RS, (($P3 -1) * $P4), $P4);
      foreach ($RS1 as $row) { // and doubleval(RS.AbsolutePage) = doubleval(Nvl($_REQUEST["P3"),RS.AbsolutePage))
        if (f($row, "ativo") == "N") {
          $w_cor = "#31BCBC";
        }
        elseif ($w_cor == "#EFEFEF" or $w_cor == "") {
          $w_cor = "#FDFDFD";
        } else {
          $w_cor = "#EFEFEF";
        }

        ShowHTML('<tr valign="top" bgcolor="' . $w_cor . '">');
        ShowHTML('    <td><font face="Verdana" size="1">');
        if ($row["ativo"] == 'N') {
          ShowHTML('      (DESATIVADA) ');
        }
        if ($p_tipo == "H") {
          ShowHTML('      <a href="' . $row["ln_internet"] . '" target="_blank">' . $row["ds_cliente"] . '</a></b></font></td>');
        } else {
          ShowHTML('      ' . $row["ds_cliente"] . '</font></td>');
        }
        if ($_SESSION["USERNAME"] == "SEDF" or $_SESSION["USERNAME"] == "CTIS" or SubStr($_SESSION["USERNAME"], 0, 2) == "RE" or $_SESSION["USERNAME"] == "SBPI") {
          if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'username') !== false) {
            ShowHTML('    <td><font face="Verdana" size="1">' . f($row, "ds_username") . '</font></td>');
          }
          if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'senha') !== false) {
            ShowHTML('    <td align="center"><font face="Verdana" size="1">' . f($row, "DS_SENHA_ACESSO") . '</font></td>');
          }
        }
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'acao') !== false) {
          ShowHTML('    <td align="center"><font face="Verdana" size="1">' . Nvl(FormataDataEdicao(f($row, "dt_alteracao")), "---") . '</font></td>');
        }
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'diretor') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "no_diretor"), "---") . '</td>');
        }
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'secretario') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "no_secretario"), "---") . '</td>');
        }
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'mail') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "ds_email_internet"), "---") . '</td>');
        }
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'contato') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "no_contato_internet"), "---") . '</td>');
        }
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'telefone') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "nr_fone_internet"), "---") . '</td>');
          ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "nr_fax_internet"), "---") . '</td>');
        }
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'endereco') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "ds_logradouro"), "---") . '</td>');
        }
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'localizacao') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1">' . ((nvl(f($row, "no_bairro"), '') == '') ? '' : f($row, "no_bairro") . ' - ') . f($row, "no_municipio") . '-' . f($row, "sg_uf") . '</font></td>');
        }
        if ($_REQUEST['p_campos'] != '' and strpos($p_campos, 'cep') !== false) {
          ShowHTML('    <td><font face="Verdana" size="1">' . Nvl(f($row, "nr_cep"), "---") . '</td>');
        }

        if ($p_tipo == "H") {
          ShowHTML('    <td><font face="Verdana" size="1">');
          if ($_SESSION["USERNAME"] == "SBPI") {
            ShowHTML('       <A CLASS="SS" HREF="' . $w_pagina . 'CADASTROESCOLA&p_chave=' . $row["sq_cliente"] . '&O=A' . MontaFiltro("GET") . '" Title="Alteração dos dados da escola!">Alt</A>');
          }

          ShowHTML('       <A CLASS="SS" HREF="manut.php?par=showlog&O=L&p_cliente=' . f($row, "sq_cliente") . '&w_ee=1&P3=1&P4=30" Title="Exibe o registro de ocorrências da escola!" target="_blank">Log</A>');
          if (nvl($row["adm"], "nulo") != "nulo") {
            ShowHTML('       <A CLASS="SS" HREF="controle.php?p_cliente=' . $row["sq_cliente"] . '&O=L&par=AdmLog&w_ee=1&P3=1&P4=30" Title="Exibe o formulário de dados administrativos preenchido pela escola!" target="_blank">Adm</A>');
          }
        }
      }
    }
    ShowHTML('</table>');
    ShowHTML('<tr><td><td colspan="5" align="center"><hr>');

    if ($p_tipo == "H") {
      ShowHTML('<tr><td align="center" colspan=5>');
      //           MontaBarra ('Controle.asp?CL=" & CL & "&w_ew=" & w_ew & "&Q=" & $_REQUEST["Q") & "&C=" & $_REQUEST["C") & "&D=" & $_REQUEST["D") & "&U=" & $_REQUEST["U") & w_especialidade, cDbl(RS.PageCount), cDbl($_REQUEST["P3")), cDbl($_REQUEST["P4")), cDbl(RS.RecordCount));
      if ($RS1)
        barra($w_dir .
        $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=' . $O . '&P1=' . $P1 . '&P2=' . $P2 . '&TP=' . $TP . '&SG=' . $SG . MontaFiltro('GET'), ceil(count($RS) / $P4), $P3, $P4, count($RS));
      ShowHTML('</tr>');
    } else {
      ShowHTML('<TR><TD><TD colspan="3"><p align="justify"><img src="img/ico_educacao.gif" width="16" height="16" border="0" align="center">&nbsp;<font size="2"><b>Nenhuma ocorrência encontrada para opções acima.');
    }

    ShowHTML('</TABLE>');

  }

  //Fim da Pesquisa de Escolas Públicas

  // =========================================================================
  // Cadastro de arquivos
  // -------------------------------------------------------------------------
  function Arquivos() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];

    if ($w_troca > "") { // Se for recarga da página
      $w_dt_arquivo = $_REQUEST["w_dt_arquivo"];
      $w_ds_titulo = $_REQUEST["w_ds_titulo"];
      $w_in_ativo = $_REQUEST["w_in_ativo"];
      $w_ds_arquivo = $_REQUEST["w_ds_arquivo"];
      $w_ln_arquivo = $_REQUEST["w_ln_arquivo"];
      $w_in_destinatario = $_REQUEST["w_in_destinatario"];
      $w_nr_ordem = $_REQUEST["w_nr_ordem"];
      $w_pasta    = $_REQUEST["w_pasta"];
    }
    elseif ($O == "L") {
      //Recupera todos os registros para a listagem
      if ($_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI") {
        $SQL = 'select * from sbpi.Cliente_Arquivo where sq_cliente = 0 order by nr_ordem, ltrim(upper(ds_titulo))';
      } else {
        $SQL = 'select * from sbpi.Cliente_Arquivo where sq_cliente = ' . $CL . ' order by ativo, nr_ordem';
      }
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }
    elseif (strpos('AEV', $O) !== false and $w_troca == "") {
      //Recupera os dados do endereço informado
      $SQL = "select b.ds_diretorio, a.* from sbpi.Cliente_Arquivo a inner join sbpi.Cliente_Site b on (a.sq_cliente = b.sq_cliente) where a.sq_arquivo = " . $w_chave;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_dt_arquivo = FormataDataEdicao(f($RS, "dt_arquivo"));
      $w_ds_titulo = f($RS, "ds_titulo");
      $w_in_ativo = f($RS, "ativo");
      $w_ds_arquivo = f($RS, "ds_arquivo");
      $w_ln_arquivo = f($RS, "ln_arquivo");
      $w_in_destinatario = f($RS, "in_destinatario");
      $w_nr_ordem = f($RS, "nr_ordem");
      $w_ds_diretorio = f($RS, "ds_diretorio");
      $w_pasta        = f($RS, "pasta");

    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('<link href="/css/particular.css" media="screen" rel="stylesheet" type="text/css" />');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    if (strpos("IAEP", $O) !== false) {
      ScriptOpen('JavaScript');
      ValidateOpen('Validacao');
      if (strpos("IA", $O) !== false) {
        Validate("w_ds_titulo", "Título", "", "1", "2", "50", "1", "1");
        Validate("w_ds_arquivo", "Descrição", "", "1", "2", "200", "1", "1");
        Validate ("w_pasta" , "Pasta"        , "SELECT" , "1" , "1" , "10"   , "1" , "1");
        if ($O == 'I') {
          Validate("w_ln_arquivo", "Link", "", "1", "2", "200", "1", "1");
        } else {
          Validate("w_ln_arquivo", "Link", "", "", "2", "200", "1", "1");
        }
        Validate("w_nr_ordem", "Nr. de ordem", "", "1", "1", "4", "1", "0123546789");
      }
      ShowHTML(' if (theForm.w_ln_arquivo.value > ""){');
      ShowHTML('    if((theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.DLL\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.SH\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.BAT\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.EXE\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.ASP\')!=-1) || (theForm.w_ln_arquivo.value.toUpperCase().lastIndexOf(\'.PHP\')!=-1)) {');
      ShowHTML('       alert(\'Tipo de arquivo não permitido!\');');
      ShowHTML('       theForm.w_ln_arquivo.focus(); ');
      ShowHTML('       return false;');
      ShowHTML('    }');
      ShowHTML('  }');
      ShowHTML('  theForm.Botao[0].disabled=true;');
      ShowHTML('  theForm.Botao[1].disabled=true;');
      ValidateClose();
      ScriptClose();
    }
    ShowHTML('</HEAD>');
    if ($w_troca > "") {
      BodyOpen('onLoad="document.Form.' . $w_troca . '.focus()";');
    }
    elseif ($O == "I" or $O == "A") {
      BodyOpen("onLoad='document.Form.w_ds_titulo.focus()';");
    } else {
      BodyOpen("onLoad='document.focus()';");
    }
    ShowHTML('<B><FONT COLOR="#000000">Cadastro de arquivos da rede de ensino</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    If ($O == 'L') {
      // Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $dir . $w_pagina . $par . $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><b>Ordem</font></td>');
      ShowHTML('          <td><b>Arquivo</font></td>');
      ShowHTML('          <td><b>Descrição</font></td>');
      ShowHTML('          <td><b>Ativo</font></td>');
      ShowHTML('          <td><b>Operações</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td align="center">' . f($row, "nr_ordem") . '</td>');
          ShowHTML('        <td>' . f($row, "ds_titulo") . '</td>');
          ShowHTML('        <td>' . f($row, "ds_arquivo") . '</td>');
          ShowHTML('        <td align="center">' . f($row, "ativo") . '</td>');
          ShowHTML('        <td align="top" nowrap>');
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . 'arquivos' . $w_ew . "&R=" . $w_pagina . 'arquivos' . $w_ew . "&O=A&CL=" . $CL . "&w_chave=" . f($row, "sq_arquivo") . '">Alterar</A>&nbsp');
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . "arquivos&R=" . $w_pagina . "&O=E&w_chave=" . f($row, "sq_arquivo") . '" onClick="return confirm(\'Confirma a exclusão do registro?\');">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
        }
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif (strpos("IAEV", $O) !== false) {
      if (strpos("EV", $O) !== false) {
        $w_Disabled = ' DISABLED ';
      }
      ShowHTML('<FORM action="' . $w_pagina . 'Grava" method="POST" name="Form" onSubmit="return(Validacao(this));" enctype="multipart/form-data">');
      ShowHTML('<INPUT type="hidden" name="SG" value="ARQUIVOS">');
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');

      ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td valign="top"><b><u>T</u>ítulo:</b><br><input ' . $w_Disabled . ' accesskey="T" type="text" name="w_ds_titulo" class="STI" SIZE="50" MAXLENGTH="50" VALUE="' . $w_ds_titulo . '" ></td>');
      if ($O == "I") {
        ShowHTML('          <td valign="top"><b>Cadastramento:</b><br><input ' . $w_Disabled . ' type="text" name="w_dt_arquivo" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(time(), 2)) . '"></td>');
      } else {
        ShowHTML('          <td valign="top"><b>Última alteração:</b><br><input ' . $w_Disabled . ' type="text" name="w_dt_arquivo" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime($w_dt_arquivo, 2)) . '"></td>');
      }
      ShowHTML('        </table>');
      ShowHTML('      <tr><td valign="top"><b><u>D</u>escrição:</b><br><textarea ' . $w_Disabled . ' accesskey="D" name="w_ds_arquivo" class="STI" ROWS=5 cols=65>' . $w_ds_arquivo . '</TEXTAREA></td>');
      ShowHTML('      <tr>');
      ShowHTML('      </tr>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td><font size="1"><b>Exibir na <u>p</u>asta:</b><br><select ' . $w_Disabled . ' accesskey="P" name="w_pasta" class="STS" SIZE="1">');
      if($w_pasta == "1"){
        ShowHTML('            <OPTION VALUE="" >Selecione uma opção...<OPTION SELECTED VALUE="1">Meses<OPTION VALUE="2">Formulários<OPTION VALUE="3">Diversos');
      }elseif($w_pasta == "2"){
        ShowHTML('            <OPTION VALUE="" >Selecione uma opção...<OPTION VALUE="1">Meses<OPTION SELECTED VALUE="2">Formulários<OPTION VALUE="3">Diversos');
      }elseif($w_pasta == "3"){
        ShowHTML('            <OPTION VALUE="" >Selecione uma opção...<OPTION VALUE="1">Meses<OPTION VALUE="2">Formulários<OPTION SELECTED VALUE="3">Diversos');
      }else{
        ShowHTML('            <OPTION SELECTED VALUE="" >Selecione uma opção...<OPTION VALUE="1">Meses<OPTION VALUE="2">Formulários<OPTION VALUE="3">Diversos');
      }
      ShowHTML('            </SELECTED></TD></TR>');
      ShowHTML('      <tr><td valign="top"><b><u>L</u>ink:</b><br><input ' . $w_Disabled . ' accesskey="L" type="file" name="w_ln_arquivo" class="STI" SIZE="80" MAXLENGTH="100" VALUE="" >');
      if ($w_ln_arquivo > '') {
        ShowHTML('              <b><a class="SS" href="' . $w_ds_diretorio . '/' . $w_ln_arquivo . '" target="_blank" title="Clique para exibir o arquivo atual.">Exibir</a></b>');
      }
      ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td><b><u>N</u>r. de ordem:</b><br><input "' . $w_Disabled . '" accesskey="N" type="text" name="w_nr_ordem" class="STI" SIZE="4" MAXLENGTH="4" VALUE="' . $w_nr_ordem . '"></td>');
      ShowHTML('          <td><b><u>D</u>estinatários:</b><br><select "' . $w_Disabled . '" accesskey="D" name="w_in_destinatario" class="STS" SIZE="1">');
      if ($w_in_destinatario == 'A') {
        ShowHTML('            <OPTION VALUE="A" SELECTED>Apenas alunos <OPTION VALUE="P">Apenas professores <OPTION VALUE="T">Professores e alunos <OPTION VALUE="E">Escola');
      }
      elseif ($w_in_destinatario == "E") {
        ShowHTML('            <OPTION VALUE="A">Apenas alunos <OPTION VALUE="P">Apenas professores <OPTION VALUE="T">Professores e alunos <OPTION VALUE="E" SELECTED>Escola');
      }
      elseif ($w_in_destinatario == "P") {
        ShowHTML('            <OPTION VALUE="A">Apenas alunos <OPTION VALUE="P" SELECTED>Apenas professores <OPTION VALUE="T">Professores e alunos <OPTION VALUE="E">Escola');
      } else {
        ShowHTML('            <OPTION VALUE="A">Apenas alunos <OPTION VALUE="P">Apenas professores <OPTION VALUE="T" SELECTED>Professores e alunos <OPTION VALUE="E">Escola');
      }
      ShowHTML('            </SELECTED></TD>');

      MontaRadioSN('<b>Exibir no site?</b>', $w_in_ativo, 'w_in_ativo');
      ShowHTML('        </table>');
      ShowHTML('      <tr>');
      ShowHTML('      <tr><td align="center" colspan=4><hr>');
      if ($O == "E") {
        ShowHTML('   <input class="STB" type="submit" name="Botao" value="Excluir" onClick="return confirm(\'Confirma a exclusão do registro?\');">');
      } else {
        if (O == "I") {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Incluir">');
        } else {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }
      ShowHTML('            <input class="STB" type="button" onClick="location.href=\'' . $dir . $w_pagina . $par . $w_ew . "&CL=" . $CL . '&O=L\';" name="Botao" value="Cancelar">');
      ShowHTML('          </td>');
      ShowHTML('      </tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');
      ShowHTML('</FORM>');
    } else {
      ScriptOpen('JavaScript');
      ShowHTML(' alert(\'Opção não disponível\');');
      ScriptClose();
    }
    ShowHTML('</table>');
    ShowHTML('</center>');
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
    include_once ("classes/menu/xPandMenu.php");
    // Instanciando a classe menu
    $root = new XMenu();

    $w_imagem = 'img/SheetLittle.gif';
    $i = 0;
    $i++;
    eval ('$node' . i . ' = &$root->addItem(new XNode(\'Manual SIGE-WEB\',\'manuais/operacao/\',$w_Imagem,$w_Imagem,\'_blank\', null));');
    if ($_SESSION['USERNAME'] == 'ADMINISTRATIVO' || $_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Administrativo\',$w_pagina.\'administrativo\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'IMPRENSA' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Arquivos (<i>download</i>)\',$w_pagina.\'arquivos\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Calendário Base\',$w_pagina.\'calend_base\',$w_Imagem,$w_Imagem,\'content\', null));');
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Calendário da Rede\',$w_pagina.\'calend_rede\',$w_Imagem,$w_Imagem,\'content\', null));');
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Dados da escola\',$w_pagina.\'dadosescola\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    // if ($_SESSION['USERNAME']=='SEDF' || $_SESSION['USERNAME']=='IMPRENSA' || $_SESSION['USERNAME']=='SBPI') {
    //  $i++; eval('$node'.i.' = &$root->addItem(new XNode(\'Envio de e-Mail\',$w_pagina.\'email\',$w_Imagem,$w_Imagem,\'content\', null));');
    //}
    if ($_SESSION['USERNAME'] != 'ADMINISTRATIVO' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Escolas\',$w_pagina.\'escolas\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Escolas Particulares\',$w_pagina.\'escpart\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Homologação de Calendário\',$w_pagina.\'escparthomolog\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'IMPRENSA' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Informativo\',$w_pagina.\'newsletter\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Modalidades de Ensino\',$w_pagina.\'modalidades\',$w_Imagem,$w_Imagem,\'content\', null));');
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Mensagens a alunos\',$w_pagina.\'msgalunos\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'IMPRENSA' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Notícias\',$w_pagina.\'noticias\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Rede Particular\',$w_pagina.\'redepart\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Tipo de Instituição\',$w_pagina.\'tipocliente\',$w_Imagem,$w_Imagem,\'content\', null));');
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Senha\',$w_pagina.\'senha\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Senhas Especiais\',$w_pagina.\'senhaesp\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Verif. Arquivos\',$w_pagina.\'verifarq\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    if ($_SESSION['USERNAME'] == 'SEDF' || $_SESSION['USERNAME'] == 'SBPI') {
      $i++;
      eval ('$node' . i . ' = &$root->addItem(new XNode(\'Log\',$w_pagina.\'log\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    $i++;
    eval ('$node' . i . ' = &$root->addItem(new XNode(\'Encerrar\',$w_pagina.\'Sair\',$w_Imagem,$w_Imagem,\'_top\', \'onClick="return(confirm(\\\'Confirma saída do sistema?\\\'));"\' ));');

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
    echo ('<BODY topmargin=0 bgcolor="#FFFFFF" BACKGROUND="img/background.gif" BGPROPERTIES="FIXED" text="#000000" link="#000000" vlink="#000000" alink="#FF0000" ');
    ShowHTML('onLoad="javascript:top.content.location=\'' . $w_pagina . 'escolas\';"> ');
    ShowHTML('  <CENTER><table border=0 cellpadding=0 height="80" width="100%">');
    ShowHTML('      <tr><td colspan=2 width="100%"><table border=0 width="100%" cellpadding=0 cellspacing=0><tr valign="top">');
    ShowHTML('          <td>Usuário:<b>' . $_SESSION['USERNAME'] . '</b></TD>');
    //ShowHTML ('          <td align="right"><a class="hl" href="help.php?par=Menu&TP=<img src=images/Folder/hlp.gif border=0> SIW - Visão Geral&SG=MESA&O=L" target="content" title="Exibe informações sobre os módulos do sistema."><img src="images/Folder/hlp.gif" border=0></a></TD>');
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
  // Monta barra de navegação
  // -------------------------------------------------------------------------
  function barra($p_link, $p_PageCount, $p_AbsolutePage, $p_PageSize, $p_RecordCount) {

    extract($GLOBALS);
    global $w_Disabled;

    global $P3;
    ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT">');
    ShowHTML('  function pagina (pag) {');
    ShowHTML('    document.Barra.P3.value = pag;');
    ShowHTML('    document.Barra.submit();');
    ShowHTML('  }');
    ShowHTML('</SCRIPT>');
    ShowHTML('<left><FORM ACTION="' . $p_link . '" METHOD="POST" name="Barra">');
    ShowHTML('<input type="Hidden" name="P4" value="' . $p_PageSize . '">');
    ShowHTML('<input type="Hidden" name="p_modalidade" value="' . $p_modalidade . '">');

    ShowHTML(MontaFiltro('POST'));
    if ($p_PageCount == $p_AbsolutePage) {
      ShowHTML('<br><b>' . ($p_RecordCount - (($p_PageCount -1) * $p_PageSize)) . '</b> resultados de <b>' . $p_RecordCount . '</b>');
    } else {
      ShowHTML('<br><b>' . $p_PageSize . '</b> resultados de <b>' . $p_RecordCount . '</b>');
    }
    if ($p_PageSize < $p_RecordCount) {
      ShowHTML('<br/>Página ');
      echo '<SELECT CLASS="texto_livre" NAME="P3" SIZE=1 onChange="document.Barra.submit();">';
      for ($i = 1; $i <= $p_PageCount; $i++) {
        echo '<option value="' . $i . '" ' . (($i == $p_AbsolutePage) ? 'SELECTED' : '') . '>' . $i;
      }
      echo '</SELECT>';
      ShowHTML(' de <span class="texto_livre">&nbsp;' . $p_PageCount . '&nbsp;</span>');
      ShowHTML('<p class="pag">');
      if ($p_AbsolutePage > 1) {
        ShowHTML('<A TITLE="Primeira página" HREF="javascript:pagina(1)"><span class="primeira">Primeira</span></A>');
        ShowHTML('<A TITLE="Página anterior" HREF="javascript:pagina(' . ($p_AbsolutePage -1) . ')"><span class="anterior">Anterior</span></A>');
      } else {
        ShowHTML('<span class="primeiraOff">Primeira</span>');
        ShowHTML('<span class="anteriorOff">Anterior</span>');
      }
      if ($p_PageCount == $p_AbsolutePage) {
        ShowHTML('<span class="proximaOff">Próxima</span>');
        ShowHTML('<span class="ultimaOff">Última</span>');
      } else {
        ShowHTML('<A TITLE="Página seguinte" HREF="javascript:pagina(' . ($p_AbsolutePage +1) . ')"><span class="proxima">Próxima</span></A>');
        ShowHTML('<A TITLE="Última página" HREF="javascript:pagina(' . $p_PageCount . ')"><span class="ultima">Última</span></A>');
      }
      ShowHTML('</p>');
    } else {
      ShowHTML('<br/>Página <b>' . $p_AbsolutePage . '</b> de <b>' . $p_PageCount . '</b>');
    }
    ShowHtml('</FORM></center>');
  }

  function AdmLog() {

    extract($GLOBALS);
    global $w_Disabled;

    $w_Disabled = ' DISABLED ';

    $SQL = "select a.ds_cliente, b.* from sbpi.Cliente a inner join sbpi.Cliente_Admin b on (a.sq_cliente = b.sq_cliente) where a.sq_cliente = " . $_REQUEST['p_cliente'];
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    foreach ($RS as $row) {
      $RS = $row;
      break;
    }
    if (count($RS) > 0) {
      $w_sq_cliente = f($RS, "sq_cliente");
      $ds_cliente = f($RS, "ds_cliente");
      $w_limpeza = f($RS, "limpeza_terceirizada");
      $w_merenda = f($RS, "merenda_terceirizada");
      $w_banheiro = f($RS, "banheiro");
    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('</HEAD>');
    //BodyOpen ('onLoad=\'document.Form.focus();\'');
    BodyOpen("onLoad='document.focus()';");
    ShowHTML('<B><FONT COLOR="#000000">' . $ds_cliente . ' - Dados administrativos </FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
    //AbreForm ("Form", $w_pagina."Grava", "POST", "return(Validacao(this));", null);
    ShowHTML('<input type="hidden" name="SG" value="administrativo">');
    ShowHTML('<input type="hidden" name="O" value="A">');
    ShowHTML('<input type="hidden" name="w_equipamento[]" value="">');
    ShowHTML('<input type="hidden" name="w_codigo[]" value="">');
    ShowHTML(MontaFiltro("POST"));

    ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
    ShowHTML('    <table width="97%" border="0">');
    ShowHTML('      <tr><td align="center"><font size=1 color="#BC3131"></td></tr>');
    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><b>Serviços</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1></td></tr>');

    ShowHTML('      <tr><td align="left" height="1" >');

    If ($w_limpeza == "S")
      ShowHTML("          <li>O serviço de limpeza da escola é terceirizado.");
    Else
      ShowHTML("          <li>O serviço de limpeza da escola não é terceirizado.");
    If ($w_merenda == "S")
      ShowHTML("          <li>A escola oferece merenda.");
    Else
      ShowHTML("          <li>A escola não oferece merenda.");
    If ($w_banheiro == "1")
      ShowHTML("          <li>Na escola existe 1 quadro magnético.");
    Else
      ShowHTML("          <li>Na escola existem " . $w_banheiro . " quadros magnéticos.");
    ShowHTML('      </td></tr>');
    ShowHTML('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');

    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><b>Equipamentos</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');

    $SQL = "select a.nome, b.sq_equipamento, b.nome nm_equipamento, coalesce(c.quantidade, 0) quantidade " . $crlf .
    "  from sbpi.Tipo_Equipamento                    a " . $crlf .
    "       inner      join sbpi.Equipamento         b on (a.sq_tipo_equipamento = b.sq_tipo_equipamento) " . $crlf .
    "       inner      join sbpi.Cliente_Equipamento c on (b.sq_equipamento      = c.sq_equipamento and " . $crlf .
    "                                                    c.sq_cliente          = " . $_REQUEST['p_cliente'] . " " . $crlf .
    "                                                   ) " . $crlf .
    "order by a.codigo, b.nome " . $crlf;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    if (count($RS) <= 0) {
      ShowHTML('      <tr><td><b>Não há equipamentos cadastrados.</td>');
    } else {
      ShowHTML('      <tr><td><table border=0 width="100%">');
      $w_atual = '';
      foreach ($RS as $row) {
        if ($w_atual != f($row, "nome")) {
          ShowHTML('        <TR><TD colspan=2><font size=2><b>' . f($row, "nome") . '</b>');
          $w_atual = f($row, "nome");
          $w_cont = 0;
        }
        if (($w_cont % 2) == 0) {
          ShowHTML('        <TR valign="top">');
          $w_cont = 0;
        }

        if (f($row, "quantidade") > 0) {
          ShowHTML('        <td><font size=1><INPUT ' . $w_Disabled . ' class="STI" type="text" name="w_equipamento[]" size="3" maxlength="3" value="' . f($row, "quantidade") . '"> ');
          ShowHTML('        <b>' . f($row, "nm_equipamento") . '</b>');
        } else {
          ShowHTML('        ' . f($row, "nm_equipamento"));
        }
        ShowHTML('                         <INPUT ' . $w_Disabled . ' type="hidden" name="w_codigo[]" value="' . f($row, "sq_equipamento") . '"> ' . '</td>');
        $w_cont++;
      }
      ShowHTML('      </table>');
    }

    //Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" colspan="3">');
    ShowHTML('          </td>');
    ShowHTML('      </tr>');
    ShowHTML('    </table>');
    ShowHTML('    </TD>');
    ShowHTML('</tr>');
    ShowHTML('</FORM>');
    ShowHTML('</table>');
    Rodape();

  }

  function redepart() {
    extract($GLOBALS);
    global $w_Disabled;

    ShowHTML('</HEAD>');
    BodyOpen("onLoad='document.focus()';");
    ShowHTML('<B><FONT COLOR="#000000">Cadastro da Rede Particular</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    ShowHTML('<FORM action=" ' . $w_pagina . 'Grava" method="POST" name="Form" onSubmit="return(Validacao(this));" enctype="multipart/form-data">');

    //  ShowHTML ('<INPUT type="hidden" name="w_sq_cliente" value="" & replace(CL,"sq_cliente=",") & "">"');
    ShowHTML('<INPUT type="hidden" name="SG" value="REDEPART">');

    ShowHTML('<tr bgcolor="#EFEFEF"><td align="center">');
    ShowHTML('    <table width="95%" border="0">');
    ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML('        <tr><td><b><u>A</u>rquivo:</b><br><input "' . $w_Disabled . '" accesskey="A" type="file" name="w_no_arquivo" class="STI" SIZE="80" MAXLENGTH="100" VALUE="" title="OBRIGATÓRIO. Clique no botão ao lado para localizar o arquivo que contém a versão do componente. Ele será transferido automaticamente para o servidor.">');
    ShowHTML('      <tr><td align="center" colspan=4><hr>');
    ShowHTML('            <input class="STB" type="submit" name="Botao" value="Enviar">');
    ShowHTML('          </td>');
    ShowHTML('      </tr>');
    ShowHTML('    </table>');
    ShowHTML('    </TD>');
    ShowHTML('</tr>');
    ShowHTML('</FORM>');
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();

  }

  // =========================================================================
  // Rotina de encerramento da sessão
  // -------------------------------------------------------------------------
  function Sair() {
    extract($GLOBALS);
    global $w_Disabled;

    // Eliminar todas as variáveis de sessão.
    $_SESSION = array ();
    // Finalmente, destruição da sessão.
    session_destroy();

    ScriptOpen('JavaScript');
    ShowHTML('  top.location.href=\'' . $w_pagina . '\';');

    ScriptClose();
  }

  function dadosEscola() {
    extract($GLOBALS);
    global $w_Disabled;

    $p_escola = $_REQUEST['p_escola'];

    If (nvl($p_escola, "") != "") {
      $SQL = "select * from sbpi.Cliente where sq_cliente = " . $p_escola;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      $RS = $RS[0];
      $p_regiao = f($RS, "sq_regiao_adm");
      $p_regional = f($RS, "sq_cliente_pai");
      $p_tipo_escola = f($RS, "sq_tipo_cliente");
      $p_local = f($RS, "localizacao");
    }

    Cabecalho();
    ShowHTML("<HEAD>");
    ScriptOpen("JavaScript");
    ValidateOpen("Validacao");
    Validate("p_escola", "Unidade de ensino", "SELECT", "1", "1", "18", "", "1");
    Validate("p_regiao", "Região administrativa", "SELECT", "1", "1", "18", "", "1");
    Validate("p_regional", "Regional de ensino", "SELECT", "1", "1", "18", "", "1");
    Validate("p_tipo_escola", "Tipo de unidade", "SELECT", "1", "1", "18", "", "1");
    ShowHTML("  theForm.Botao.disabled=true;");
    ValidateClose();
    ScriptClose();
    ShowHTML("</HEAD>");

    If ($w_troca > "") { //' Se for recarga da página
      BodyOpen("onLoad='document.Form." . $w_troca . ".focus();'");
    } Else {
      BodyOpen("onLoad=document.Form.p_escola.focus();");
    }

    ShowHTML('<B><FONT COLOR="#000000">' . $w_TP . '</FONT></B>');
    ShowHTML('<B><FONT size="2" COLOR="#000000">Vinculação e tipologia</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="100%">');
    ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
    ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td align="center" valign="top"><table border=0 width="90%" cellspacing=0>');
    AbreForm('Form', $w_pagina . "grava", "POST", "return(Validacao(this));", null);

    ShowHTML('<INPUT type="hidden" name="w_troca" value="">');
    ShowHTML('<INPUT type="hidden" name="SG" value="DADOSESCOLA">');

    ShowHTML('      <tr><td colspan=2><table border=0 width="90%" cellspacing=0>');
    ShowHTML('        <tr valign="top">');

    SelecaoEscola("Unidad<u>e</u> de ensino:", "E", "Selecione unidade.", $p_escola, null, "p_escola", "ESCOLA", "onChange='document.Form.action=\"controle.php?par=dadosescola&O=p \"; document.Form.submit();'");
    ShowHTML('        <tr valign="top">');
    SelecaoRegiaoAdm("Região a<u>d</u>ministrativa:", "D", "Indique a região administrativa.", $p_regiao, $p_escola, "p_regiao", null, null);
    ShowHTML('        <tr valign="top">');
    SelecaoRegionalEscola("Regional de en<u>s</u>ino:", "S", "Indique a regional de ensino.", $p_regional, $p_escola, "p_regional", null, null);
    ShowHTML('        <tr valign="top">');
    SelecaoTipoEscola("<u>T</u>ipo de Escola:", "T", "Selecione o tipo da Escola.", $p_tipo_escola, $p_escola, "p_tipo_escola", null, null);
    ShowHTML('         <tr valign="top">');
    ShowHTML('          <td><b>Localização</b><br>');
    If ($p_local == "1") {
      ShowHTML('              <input  type="radio" name="p_local" value="0"> Não informada <input  type="radio" name="p_local" value="1" checked> Urbana <input  type="radio" name="p_local" value="2"> Rural ');
    }
    ElseIf ($p_local == "2") {
      ShowHTML('              <input  type="radio" name="p_local" value="0"> Não informada <input  type="radio" name="p_local" value="1"> Urbana <input  type="radio" name="p_local" value="2" checked> Rural ');
    } Else {
      ShowHTML('              <input  type="radio" name="p_local" value="0" checked> Não informada <input  type="radio" name="p_local" value="1"> Urbana <input  type="radio" name="p_local" value="2"> Rural ');
    }
    ShowHTML('          </table>');
    ShowHTML('      <tr valign="top">');
    ShowHTML('      <tr>');
    ShowHTML('      <tr><td align="center" colspan="2" height="1" bgcolor="#000000">');
    ShowHTML('      <tr><td align="center" colspan="2">');
    ShowHTML('            <input class="STB" type="submit" name="Botao" value="Gravar">');
    ShowHTML("          </td>");
    ShowHTML("      </tr>");
    ShowHTML("    </table>");
    ShowHTML("    </TD>");
    ShowHTML("</tr>");
    ShowHTML("</FORM>");
    ShowHTML("</table>");
    ShowHTML("</table>");
    ShowHTML("</center>");
    Rodape();

  }

  function log_controle() {
    extract($GLOBALS);
    global $w_Disabled;

    $p_escola = $_REQUEST["p_escola"];
    $p_bdados = $_REQUEST['p_bdados'];
    $p_operacao = $_REQUEST['p_operacao'];
    $p_regional = $_REQUEST['p_regional'];
    $p_agrega = $_REQUEST['p_agrega'];
    $p_inicio = $_REQUEST['p_inicio'];
    $p_fim = $_REQUEST['p_fim'];
    $O = $_REQUEST['O'];
    if (is_null($O))
      $O = 'P';
    if ($O == 'L') {

      If ($p_regional > "") {

        $SQL = "select * from sbpi.Cliente where sq_cliente = " . $p_regional;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        $RS = $RS[0];
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Subordinação <td><font size=1>[<b>' . $RS["ds_cliente"] . '</b>]';
      }
      If ($p_escola > "") {
        $SQL = "select * from sbpi.Cliente where sq_cliente = " . $p_escola;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        $RS = $RS[0];
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Escola <td><font size=1>[<b>' . $RS["ds_cliente"] . '</b>]';
      }
      If ($p_bdados > "") {
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Bloco de dados <td><font size=1>[<b>' . ExibeBlocoDados($p_bdados) . '</b>]';
      }
      If ($p_operacao > "") {
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Operação <td><font size=1>[<b>' . ExibeOperacao($p_operacao) . '</b>]';
      }
      If ($p_inicio > "") {
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Período <td><font size=1>[<b>' . $p_inicio . " a " . $p_fim . "</b>]";
      }
      If ($w_filtro > "")
        $w_filtro = '<table border=0><tr valign="top"><td><font size=1><b>Filtro:</b><td nowrap><font size=1><ul>' . $w_filtro . '</ul></tr></table>';

      $SQL = "select a.*, " .
      "       case to_char(a.tipo) " .
      "            when '0' then 'Consulta' " .
      "            when '1' then 'Inclusão' " .
      "            when '2' then 'Alteração' " .
      "            when '3' then 'Exclusão' " .
      "            else        'Erro' " .
      "       end nm_tipo, " .
      "       coalesce(b.nome,'Consulta') nm_funcionalidade,   coalesce(b.codigo,0) cd_funcionalidade, " .
      "       c.sq_cliente sq_escola,     case h.tipo when '1' then '* '||upper(c.ds_username) when '2' then '** '||c.ds_cliente else c.ds_cliente end nm_escola, " .
      "       d.sq_cliente sq_regional,   d.ds_cliente nm_regional, " .
      "       coalesce(e.sq_cliente, d.sq_cliente) sq_secretaria, coalesce(e.ds_cliente, d.ds_cliente) nm_secretaria " .
      "  from sbpi.Cliente_Log                    a " .
      "       left outer join sbpi.Funcionalidade b on (a.sq_funcionalidade = b.sq_funcionalidade and " .
      "                                               b.tipo              = 2 " .
      "                                              ) " .
      "       inner      join sbpi.Cliente        c on (a.sq_cliente        = c.sq_cliente) " .
      "         inner   join sbpi.Tipo_Cliente    h on (c.sq_tipo_cliente   = h.sq_tipo_cliente) " .
      "         inner   join sbpi.Cliente         d on (coalesce(c.sq_cliente_pai,0)    = d.sq_cliente) " .
      "           left  join sbpi.Cliente         e on (d.sq_cliente_pai    = e.sq_cliente), " .
      "       sbpi.cliente                        f " .
      "         inner   join sbpi.Tipo_Cliente    g on (f.sq_tipo_cliente   = g.sq_tipo_cliente) " .
      " where f.sq_cliente = " . $CL . " " .
      "   and ((g.tipo = 2 and d.sq_cliente = " . $CL . ") or " .
      "        (g.tipo <> 2) " .
      "       ) ";
      If ($p_regional > "")
        $SQL = $SQL . "   and d.sq_cliente = " . $p_regional;
      If ($p_escola > "")
        $SQL = $SQL . "   and c.sq_cliente = " . $p_escola;
      If ($p_bdados > "")
        $SQL = $SQL . "   and b.codigo     = " . $p_bdados;
      If ($p_operacao > "")
        $SQL = $SQL . "   and a.tipo       = " . $p_operacao;
      If ($p_inicio > "")
        $SQL = $SQL . "   and a.data       between to_date('" . $p_inicio . "', 'dd/mm/yyyy') and to_date('" . $p_fim . "', 'dd/mm/yyyy') + 1 ";
      switch ($p_agrega) {
        Case "SECRETARIA" :
          $SQL = $SQL . " order by nm_secretaria";
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          $w_TP = $TP . "Histórico de acessos por secretaria";
          break;
        Case "REGIONAL" :
          $SQL = $SQL . " order by nm_regional";
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          $w_TP = $TP . "Histórico de acessos por regional";
          break;
        Case "ESCOLA" :
          $SQL = $SQL . " order by nm_escola";
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          $w_TP = $TP . "Histórico de acessos por escola";
          break;
        Case "BLOCODADOS" :
          $SQL = $SQL . " order by nm_funcionalidade";
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          $w_TP = $TP . "Histórico de acessos por bloco de dados";
          break;
      }
    }

    Cabecalho();
    ShowHTML("<HEAD>");
    If ($O == "P") {
      ScriptOpen("Javascript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      Validate("p_inicio", "Início do período", "DATA", "1", "10", "10", "", "0123456789/");
      Validate("p_fim", "Fim do período", "DATA", "1", "10", "10", "", "0123456789/");
      CompData("p_inicio", "Data inicial", "<=", "p_fim", "Data final");
      ShowHTML('  var w_data, w_data1, w_data2, w_dias;');
      ShowHTML('  w_dias = 90;');
      ShowHTML('  w_data = theForm.p_inicio.value;');
      ShowHTML('  w_data = w_data.substr(3,2) + \'/\' + w_data.substr(0,2) + \'/\' + w_data.substr(6,4);');
      ShowHTML('  w_data1  = new Date(Date.parse(w_data));');
      ShowHTML('  w_data = theForm.p_fim.value;');
      ShowHTML('  w_data = w_data.substr(3,2) + \'/\' + w_data.substr(0,2) + \'/\' + w_data.substr(6,4);');
      ShowHTML('  w_data2= new Date(Date.parse(w_data));');
      ShowHTML('  var MinMilli = 1000 * 60;');
      ShowHTML('  var HrMilli = MinMilli * 60;');
      ShowHTML('  var DyMilli = HrMilli * 24;');
      ShowHTML('  var Days = Math.round(Math.abs((w_data2 - w_data1) / DyMilli));');
      ShowHTML('  if (Days > w_dias) {');
      ShowHTML('     alert("Período deve ser de, no máximo, "+w_dias+" dias!");');
      ShowHTML('     theForm.p_inicio.focus();');
      ShowHTML('     return false;');
      ShowHTML('  }');
      ValidateClose();
      ScriptClose();
    }
    ShowHTML("</HEAD>");
    If ($w_Troca > "") { //' Se for recarga da página
      BodyOpen("onLoad='document.Form." . $w_Troca . ".focus();'");
    }
    ElseIf ($O == 'P') {
      BodyOpen("onLoad='document.Form.p_agrega.focus()';");
    } Else {
      BodyOpenClean("onLoad=document.focus();");
    }
    If ($O == "L") {
      ShowHTML('<B><FONT COLOR="#000000">' . $w_TP . "</FONT></B>");
      ShowHTML("<HR>");
      If ($w_filtro > "")
        ShowHTML($w_filtro);
    } Else {
      ShowHTML('<B><FONT COLOR="#000000">' . $w_TP . '</FONT></B>');
      ShowHTML("<HR>");
    }

    ShowHTML("<div align=center><center>");
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="100%">');
    If ($O == "L" or $O == "W") {
      If ($O == "L") {
        ShowHTML('<tr><td>');
        If (MontaFiltro("GET") > "") {
          ShowHTML(' <a accesskey="F" class="SS" href="' . $w_pagina . $par . '&O=' . $O . "&R=" . $w_pagina . $O . "&O=P&P3=1&P4=" . $P4 . "&TP=" . $TP . "&SG=" . $SG . MontaFiltro("GET") . '"><u><font color="#BC5100">F</u>iltrar (Ativo)</font></a>');
        } Else {
          ShowHTML(' <a accesskey="F" class="SS" href="' . $w_pagina . $par . '&O=' . $O . "&R=" . $w_pagina . $O . "&O=P&p3=1&P4=" . $P4 . "&TP=" . $TP . "&SG=" . $SG . MontaFiltro("GET") . '"><u>F</u>iltrar (Inativo)</a>');
        }
      }
      ImprimeCabecalho();
      If (count($RS) == 0) {
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=10 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } Else {
        If ($O == "L") {
          ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT">');
          ShowHTML("  function lista (filtro, oper) {");
          ShowHTML("    if (filtro != -1) {");
          switch ($p_agrega) {
            Case "SECRETARIA" :
              ShowHTML("      document.Form.p_secretaria.value=filtro;");
              break;
            Case "REGIONAL" :
              ShowHTML("      document.Form.p_regional.value=filtro;");
              break;
            Case "ESCOLA" :
              ShowHTML("      document.Form.p_escola.value=filtro;");
              break;
            Case "BLOCODADOS" :
              ShowHTML("      document.Form.p_bdados.value=filtro;");
              break;
          }
          ShowHTML("    }");
          switch ($p_agrega) {
            Case "SECRETARIA" :
              ShowHTML(" else document.Form.p_secretaria.value='" . $_REQUEST["p_secretaria"] . "';");
              break;
            Case "REGIONAL" :
              ShowHTML(" else document.Form.p_regional.value='" . $_REQUEST["p_regional"] . "';");
              break;
            Case "ESCOLA" :
              ShowHTML(" else document.Form.p_escola.value='" . $_REQUEST["p_escola"] . "';");
              break;
            Case "BLOCODADOS" :
              ShowHTML(" else document.Form.p_bdados.value='" . $_REQUEST["p_bdados"] . "';");
              break;
          }
          ShowHTML("    if (oper != -1) document.Form.p_operacao.value=oper; else document.Form.p_operacao.value=''; ");
          ShowHTML("    document.Form.submit();");
          ShowHTML("  }");
          ShowHTML("</SCRIPT>");
          AbreForm("Form", $w_pagina . "show_log&P3=1&P4=30&O=L", "POST", "return(Validacao(this));", "Lista");
          ShowHTML(MontaFiltro("POST"));
          If (Nvl($_REQUEST["p_operacao"], "nulo") == "nulo") {
            ShowHTML('<input type="Hidden" name="p_operacao" value="">');
          }
          switch ($p_agrega) {
            Case "SECRETARIA" :
              If ($_REQUEST["p_secretaria"] == "")
                ShowHTML('<input type="Hidden" name="p_secretaria" value="">');
              break;
            Case "REGIONAL" :
              If ($_REQUEST["p_regional"] == "")
                ShowHTML('<input type="Hidden" name="p_regional" value="">');
              break;
            Case "ESCOLA" :
              If ($_REQUEST["p_escola"] == "")
                ShowHTML('<input type="Hidden" name="p_escola" value="">');
              break;
            Case "BLOCODADOS" :
              If ($_REQUEST["p_bdados"] == "")
                ShowHTML('<input type="Hidden" name="p_bdados" value="">');
              break;
          }
        }

        //      RS.PageSize       = P4
        //     RS.AbsolutePage   = P3
        $w_nm_quebra = "";
        $w_qt_quebra = 0;
        $t_acesso = 0;
        $t_consulta = 0;
        $t_inc = 0;
        $t_alt = 0;
        $t_exc = 0;
        $t_totacesso = 0;
        $t_totconsulta = 0;
        $t_totinc = 0;
        $t_totalt = 0;
        $t_totexc = 0;
        foreach ($RS as $row) {
          switch ($p_agrega) {
            Case "SECRETARIA" :
              If ($w_nm_quebra != $row["nm_secretaria"]) {
                If ($w_qt_quebra > 0) {
                  ImprimeLinha($t_acesso, $t_consulta, $t_inc, $t_alt, $t_exc, $w_chave);
                  $w_linha = $w_linha +2;
                }
                If ($O <> "W" or ($O == "W" and $w_linha <= 25)) {
                  //' Se for geração de MS-Word, coloca a nova quebra somente se não estourou o limite
                  ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top"><td><font size=1><b>' . $row["nm_secretaria"]);
                }
                $w_nm_quebra = $row["nm_secretaria"];
                $w_chave = $row["sq_regional"];
                $w_qt_quebra = 0;
                $t_acesso = 0;
                $t_consulta = 0;
                $t_inc = 0;
                $t_alt = 0;
                $t_exc = 0;
                $w_linha = $w_linha +1;
              }
              break;
            Case "REGIONAL" :
              If ($w_nm_quebra != $row["nm_regional"]) {
                If ($w_qt_quebra > 0) {
                  ImprimeLinha($t_acesso, $t_consulta, $t_inc, $t_alt, $t_exc, $w_chave);
                  $w_linha = $w_linha +2;
                }
                If (O != "W" or ($O == "W" and w_linha <= 25)) {
                  //   ' Se for geração de MS-Word, coloca a nova quebra somente se não estourou o limite
                  ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top"><td><font size=1><b>' . $row["nm_regional"]);
                }
                $w_nm_quebra = $row["nm_regional"];
                $w_chave = $row["sq_regional"];
                $w_qt_quebra = 0;
                $t_acesso = 0;
                $t_consulta = 0;
                $t_inc = 0;
                $t_alt = 0;
                $t_exc = 0;
                $w_linha = $w_linha +1;
              }
              break;
            Case "ESCOLA" :
              If ($w_nm_quebra <> $row["nm_escola"]) {
                If ($w_qt_quebra > 0) {
                  ImprimeLinha($t_acesso, $t_consulta, $t_inc, $t_alt, $t_exc, $w_chave);
                  $w_linha = $w_linha +2;
                }
                If ($O <> "W" or ($O == "W" and $w_linha <= 25)) {
                  //   ' Se for geração de MS-Word, coloca a nova quebra somente se não estourou o limite
                  ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top"><td><font size=1><b>' . $row["nm_escola"]);
                }
                $w_nm_quebra = $row["nm_escola"];
                $w_chave = $row["sq_escola"];
                $w_qt_quebra = 0;
                $t_acesso = 0;
                $t_consulta = 0;
                $t_inc = 0;
                $t_alt = 0;
                $t_exc = 0;
                $w_linha = $w_linha +1;
              }
              break;
            Case "BLOCODADOS" :
              If ($w_nm_quebra != $row["nm_funcionalidade"]) {
                If ($w_qt_quebra > 0) {
                  ImprimeLinha($t_acesso, $t_consulta, $t_inc, $t_alt, $t_exc, $w_chave);
                  $w_linha = $w_linha +2;
                }
                If ($O <> "W" or ($O == "W" and $w_linha <= 25)) {
                  // ' Se for geração de MS-Word, coloca a nova quebra somente se não estourou o limite
                  ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top"><td><font size=1><b>' . $row["nm_funcionalidade"]);
                }
                $w_nm_quebra = $row["nm_funcionalidade"];
                $w_chave = $row["cd_funcionalidade"];
                $w_qt_quebra = 0;
                $t_acesso = 0;
                $t_consulta = 0;
                $t_inc = 0;
                $t_alt = 0;
                $t_exc = 0;
                $w_linha = $w_linha +1;
              }
          }
          If ($O == "W" and $w_linha > 25) { // ' Se for geração de MS-Word, quebra a página
            ShowHTML("    </table>");
            ShowHTML("  </td>");
            ShowHTML("</tr>");
            ShowHTML("</table>");
            ShowHTML("</center></div>");
            ShowHTML('    <br style="page-break-after:always">');
            $w_linha = 0;
            $w_pag = $w_pag +1;
            CabecalhoWord($w_cliente, $w_TP, $w_pag);
            If ($w_filtro > "")
              ShowHTML($w_filtro);
            ShowHTML("<div align=center><center>");
            ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="100%">');
            ImprimeCabecalho();
            switch ($p_agrega) {
              Case "SECRETARIA" :
                ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top"><td><font size=1><b>' . $row["nm_secretaria"]);
                break;
              Case "REGIONAL" :
                ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top"><td><font size=1><b>' . $row["nm_regional"]);
                break;
              Case "ESCOLA" :
                ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top"><td><font size=1><b>' . $row["nm_escola"]);
                break;
              Case "BLOCODADOS" :
                ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top"><td><font size=1><b>' . $row["nm_funcionalidade"]);
                break;
            }
            $w_linha = $w_linha +1;
          }
          If ($row["tipo"] == "0") {
            $t_consulta = $t_consulta +1;
            $t_totconsulta = $t_totconsulta +1;
          }
          ElseIf ($row["tipo"] == "1") {
            $t_inc = $t_inc +1;
            $t_totinc = $t_totinc +1;
          }
          ElseIf ($row["tipo"] == "2") {
            $t_alt = $t_alt +1;
            $t_totalt = $t_totalt +1;
          }
          ElseIf ($row["tipo"] == "3") {
            $t_exc = $t_exc +1;
            $t_totexc = $t_totexc +1;
          }
          $t_acesso = $t_acesso +1;
          $t_totacesso = $t_totacesso +1;
          $w_qt_quebra = $w_qt_quebra +1;

        }

        ImprimeLinha($t_acesso, $t_consulta, $t_inc, $t_alt, $t_exc, $w_chave);
        ShowHTML('      <tr bgcolor="#DCDCDC" valign="top" align="right">');
        ShowHTML('          <td><b>Totais</font></td>');
        ImprimeLinha($t_totacesso, $t_totconsulta, $t_totinc, $t_totalt, $t_totexc, -1);
      }
      ShowHTML('      </center>');
      ShowHTML('    </table>');
      ShowHTML('  </td>');
      ShowHTML('</tr>');
    }
    elseif ($O == "P") {
      ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="100%">');
      ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td><div align="justify"><font size=2>Informe nos campos abaixo os valores que deseja filtrar e clique sobre o botão <i>Aplicar filtro</i>. Clicando sobre o botão <i>Remover filtro</i>, o filtro existente será apagado.</div><hr>');
      ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
      ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td align="center" valign="top"><table border=0 width="90%" cellspacing=0>');
      AbreForm('Form', $w_dir . $w_pagina . $par, 'POST', 'return(Validacao(this));', null);
      ShowHTML('<INPUT type="hidden" name="O" value="L">');
      ShowHTML('<INPUT type="hidden" name="w_troca" value="">');

      //  ' Exibe parâmetros de apresentação
      ShowHTML('         <tr><td colspan="2" align="center" bgcolor="#D0D0D0" style="border: 2px solid rgb(0,0,0);"><b>Parâmetros de Apresentação</td>');
      ShowHTML('         <tr valign="top"><td colspan=2><table border=0 width="100%" cellpadding=0 cellspacing=0><tr valign="top">');
      ShowHTML('         <td><b><U>A</U>gregar por:<br><SELECT ACCESSKEY="O" ' . $w_Disabled . ' class="STS" name="p_agrega" size="1">');

      If ($w_tipo == 2) {
        switch ($p_agrega) {
          Case "REGIONAL" :
            ShowHTML(' <option value="REGIONAL" selected>Regional <option value="ESCOLA">Escola          <option value="BLOCODADOS">Bloco de Dados');
            break;
          Case "ESCOLA" :
            ShowHTML(' <option value="REGIONAL">Regional          <option value="ESCOLA" selected>Escola <option value="BLOCODADOS">Bloco de Dados');
            break;
          Case "BLOCODADOS" :
            ShowHTML(' <option value="REGIONAL">Regional          <option value="ESCOLA">Escola          <option value="BLOCODADOS" selected>Bloco de Dados');
            break;
          Default :
            ShowHTML(' <option value="REGIONAL" selected>Regional <option value="ESCOLA">Escola          <option value="BLOCODADOS">Bloco de Dados');
            break;
        }
      } Else {
        switch ($p_agrega) {
          Case "SECRETARIA" :
            ShowHTML(' <option value="SECRETARIA" selected>Secretaria <option value="REGIONAL">Regional          <option value="ESCOLA">Escola          <option value="BLOCODADOS">Bloco de Dados');
            break;
          Case "REGIONAL" :
            ShowHTML(' <option value="SECRETARIA">Secretaria          <option value="REGIONAL" selected>Regional <option value="ESCOLA">Escola          <option value="BLOCODADOS">Bloco de Dados');
            break;
          Case "ESCOLA" :
            ShowHTML(' <option value="SECRETARIA">Secretaria          <option value="REGIONAL">Regional          <option value="ESCOLA" selected>Escola <option value="BLOCODADOS">Bloco de Dados');
            break;
          Case "BLOCODADOS" :
            ShowHTML(' <option value="SECRETARIA">Secretaria          <option value="REGIONAL">Regional          <option value="ESCOLA">Escola          <option value="BLOCODADOS" selected>Bloco de Dados');
            break;
          Default :
            ShowHTML(' <option value="SECRETARIA" selected>Secretaria <option value="REGIONAL">Regional          <option value="ESCOLA">Escola          <option value="BLOCODADOS">Bloco de Dados');
            break;
        }
      }

      ShowHTML('          </select></td>');
      ShowHTML('           </table>');
      ShowHTML('         </tr>');
      ShowHTML('         <tr><td valign="top" colspan="2" align="center" bgcolor="#D0D0D0" style="border: 2px solid rgb(0,0,0);"><b>Critérios de Busca</td>');
      ShowHTML('      <tr><td colspan=2><table border=0 width="90%" cellspacing=0><tr valign="top">');
      SelecaoRegional("<u>S</u>ubordinação:", "S", "Se desejar, indique um item para repuperar apenas escolas subordinadas a ele.", $p_regional, null, "p_regional", null, "onChange='document.Form.O.value=\"P\"; document.Form.p_escola.value=\"\"; document.Form.w_troca.value=\"p_escola\"; document.Form.submit();'");
      SelecaoEscola("<u>E</u>scola:", "E", "Selecione a Escola.", $p_escola, $p_regional, "p_escola", null, null);
      ShowHTML('         <tr valign="top">');
      SelecaoBlocoDados("<u>B</u>loco de Dados:", "B", "Selecione o bloco de dados.", $p_bdados, null, "p_bdados", null, null);
      SelecaoOperacao("<u>O</u>peração:", "O", "Selecione a operação.", $p_operacao, null, "p_operacao", null, null);
      ShowHTML('          </table>');
      ShowHTML('      <tr valign="top">');
      ShowHTML('          <td valign="top"><b><u>P</u>eríodo:</b><br><input ' . $w_Disabled . ' accesskey="P" type="text" name="p_inicio" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . (Nvl($p_inicio, subDayIntoDate(date("Ymd"), 30, 2))) . '" onKeyDown="FormataData(this,event);"> e <input ' . $w_Disabled . ' accesskey="C" type="text" name="p_fim" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(Nvl($p_fim, time())) . '" onKeyDown="FormataData(this,event);"></td>');
      ShowHTML('      <tr>');
      ShowHTML('      <tr><td align="center" colspan="2" height="1" bgcolor="#000000">');
      ShowHTML('      <tr><td align="center" colspan="2">');
      ShowHTML('            <input class="STB" type="Submit" name="Botao" value="Exibir">');
      ShowHTML('          </td>');
      ShowHTML('      </tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');
      ShowHTML('</FORM>');
      ShowHTML('</body>');
      ShowHTML('</table>');
      ShowHTML('</html>');
    }

  }

  function ShowLog() {
    extract($GLOBALS);
    global $w_Disabled;

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];
    $P3 = $_REQUEST["P3"];
    $P4 = $_REQUEST["P4"];
    $p_operacao = $_REQUEST["p_operacao"];
    $p_inicio = $_REQUEST["p_inicio"];
    $p_fim = $_REQUEST["p_fim"];
    $p_escola = $_REQUEST["p_escola"];
    $p_secretaria = $_REQUEST["p_secretaria"];
    $p_regional = $_REQUEST["p_regional"];
    $p_bdados = $_REQUEST["p_bdados"];
    $O = $_REQUEST["O"];
    //var_dump($_REQUEST);die;

    If ($O == "L") {

      $w_filtro = "";

      If ($p_regional > "") {
        $SQL = "select * from sbpi.Cliente where sq_cliente = " . $p_regional;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        $RS = $RS[0];
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Regional <td><font size=1>[<b>' . $RS["ds_cliente"] . "</b>]";
      }
      If ($p_escola > "") {
        $SQL = "select * from sbpi.Cliente where sq_cliente = " . $p_escola;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        $RS = $RS[0];
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Escola <td><font size=1>[<b>' . $RS["ds_cliente"] . "</b>]";
      }
      If ($p_bdados > "") {
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Bloco de dados <td><font size=1>[<b>' . ExibeBlocoDados($p_bdados) . "</b>]";
      }
      If ($p_operacao > "") {
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Operação <td><font size=1>[<b>' . ExibeOperacao($p_operacao) . "</b>]";
      }
      If ($p_inicio > "") {
        $w_filtro = $w_filtro . '<tr valign="top"><td align="right"><font size=1>Período <td><font size=1>[<b>' . $p_inicio . " a " . $p_fim . "</b>]";
      }
      If ($w_filtro > "")
        $w_filtro = '<table border=0><tr valign="top"><td><font size=1><b>Filtro:</b><td nowrap><font size=1><ul>' . $w_filtro . "</ul></tr></table>";

      //' Recupera todos os registros para a listagem
      $SQL = "select  a.sq_cliente_log,a.abrangencia, to_char(a.data,'dd/mm/yyyy, hh24:mi:ss') as phpdt_data,a.ip_origem, " .
             "       case to_char(a.tipo) " .
             "            when '0' then 'Consulta' " .
             "            when '1' then 'Inclusão' " .
             "            when '2' then 'Alteração' " .
             "            when '3' then 'Exclusão' " .
             "            else        'Consulta' " .
             "       end nm_tipo, " .
             "       coalesce(b.nome,'Consulta') nm_funcionalidade,   coalesce(b.codigo,0) cd_funcionalidade, " .
             "       c.sq_cliente sq_escola,     c.ds_cliente nm_escola, " .
             "       d.sq_cliente sq_regional,   d.ds_cliente nm_regional, " .
             "       e.sq_cliente sq_secretaria, e.ds_cliente nm_secretaria " .
             "  from sbpi.Cliente_Log                    a " .
             "       left outer join sbpi.Funcionalidade b on (a.sq_funcionalidade = b.sq_funcionalidade and " .
             "                                               b.tipo              = 2 " .
             "                                              ) " .
             "       inner     join sbpi.Cliente         c on (a.sq_cliente        = c.sq_cliente) " .
             "         inner   join sbpi.Cliente         d on (coalesce(c.sq_cliente_pai,0)    = d.sq_cliente) " .
             "           left  join sbpi.Cliente         e on (d.sq_cliente_pai    = e.sq_cliente), " .
             "       sbpi.cliente                        f " .
             "         inner   join sbpi.Tipo_Cliente    g on (f.sq_tipo_cliente   = g.sq_tipo_cliente) " .
             " where f.sq_cliente = " . $CL . " " .
             "   and ((g.tipo = 2 and d.sq_cliente = " . $CL . ") or " .
             "        (g.tipo <> 2) " .
             "       ) ";
      If ($p_regional > "")
        $SQL = $SQL . "   and d.sq_cliente = " . $p_regional;
      If ($p_escola > "")
        $SQL = $SQL . "   and c.sq_cliente = " . $p_escola;
      If (intval(Nvl($p_bdados, "-1")) == 0) {
        $SQL = $SQL . "   and b.codigo     is null ";
      }
      ElseIf (Nvl($p_bdados, "NULO") != "NULO") {
        $SQL = $SQL . "   and b.codigo     = " . $p_bdados;
      }
      If ($p_operacao > "")
        $SQL = $SQL . "   and a.tipo       = " . $p_operacao;
      If ($p_inicio > "")
        $SQL = $SQL . "   and a.data       between to_date('" . $p_inicio . "', 'dd/mm/yyyy') and to_date('" . $p_fim . "', 'dd/mm/yyyy') + 1 ";
      $SQL = $SQL . " order by to_char(data,'yyyymmdd') desc, data desc ";

      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }

    Cabecalho();
    ShowHTML("<HEAD>");
    ShowHTML("<TITLE>Registro de ocorrências no site da escola</TITLE>");
    If (strpos('IAEP', $O) !== false) {
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      If (strpos('IA', $O) !== false) {
        Validate("w_dt_ocorrencia", "dt_ocorrencia", "DATA", "1", "10", "10", "1", "1");
        Validate("w_ds_ocorrencia", "ds_ocorrencia", "", "1", "2", "60", "1", "1");
      }
      ShowHTML("  theForm.Botao[0].disabled=true;");
      ShowHTML("  theForm.Botao[1].disabled=true;");
      ValidateClose();
      ScriptClose();
    }
    ShowHTML("</HEAD>");
    If ($w_troca > "") {
      BodyOpen("onLoad='document.Form." . $w_troca . ".focus()';");
    }
    ElseIf ($O == "I" or $O == "A") {
      BodyOpen("onLoad='document.Form.w_dt_ocorrencia.focus()';");
    } Else {
      BodyOpen("onLoad='document.focus()';");
    }
    ShowHTML('<B><FONT COLOR="#000000">Registro de ocorrências</FONT></B>');
    ShowHTML('<HR>');
    If ($w_filtro > "")
      ShowHTML($w_filtro);
    ShowHTML("<div align=center><center>");
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    If ($O == "L") {
      //' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('<tr><td><font size="2"><b>' . $w_ds_cliente);
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="#EFEFEF" align="center">');
      ShowHTML('          <td><b>Data</font></td>');
      ShowHTML('          <td><b>Hora</font></td>');
      ShowHTML('          <td><b>Origem</font></td>');
      ShowHTML('          <td><b>Regional / Escola</font></td>');
      ShowHTML('          <td><b>Ocorrência</font></td>');
      ShowHTML('          <td><b>Operações</font></td>');
      ShowHTML('        </tr>');

      If (count($RS) == 0) { //' Se não foram selecionados registros, exibe mensagem
        ShowHTML('<tr bgcolor="#EFEFEF"><td colspan=6 align="center"><font size="2"><b>Não foram encontrados registros.</b></td></tr>');
      } Else {

        //      rs.PageSize     = P4
        //    rs.AbsolutePage = P3
        $w_ano = "";
        // While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(P3)
        $RS1 = array_slice($RS, (($P3 -1) * $P4), $P4);
        foreach ($RS1 as $row) {

          If ($w_cor = "#EFEFEF" or $w_cor = "")
            $w_cor = "#FDFDFD";
          Else
            $w_cor = "#EFEFEF";

          If ($wAno != date('m/Y', $row["phpdt_data"])) {
            ShowHTML('<tr bgcolor="#C0C0C0" valign="top"><TD colspan=6 align="center"><font size=2><B>' . month($row["phpdt_data"]) . "/" . year($row["phpdt_data"]) . '</b></font></td></tr>');
            $wAno = date('m/Y', $row["phpdt_data"]);
          }
          ShowHTML('     <tr bgcolor="' . $w_cor . '" valign="top">');
          If ($w_data != FormataDataEdicao($row["phpdt_data"])) {
            ShowHTML('<td align="left">' . FormataDataEdicao($row["phpdt_data"]) . "</td>");
            $w_data = FormataDataEdicao($row["phpdt_data"]);
          } Else {
            ShowHTML('        <td align="center"></td>');
          }
          ShowHTML('        <td align="left">' . date('H:i:s', $row["phpdt_data"]) . "</td>");

          ShowHTML('        <td align="left">' . $row['ip_origem'] . '</td>');
          ShowHTML('        <td align="left">' . $row['nm_escola'] . '</td>');

          ShowHTML('        <td>' . $row["abrangencia"] . "</td>");
          ShowHTML('        <td align="top" nowrap>');
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . $par . "&O=V&sq_cliente_log=" . $row["sq_cliente_log"] . "&sq_cliente=" . $w_chave . "&R=" . $w_pagina . $par . '">Detalhes</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');

        }
      }
      ShowHTML("      </center>");
      ShowHTML("    </table>");

      ShowHTML("<tr><td colspan=2><br><hr>");

      ShowHTML('<tr><td align="center" colspan=2>');

      barra($w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=' . $O . '&P1=' . $P1 . '&P2=' . $P2 . '&TP=' . $TP . '&SG=' . $SG . MontaFiltro('GET'), ceil(count($RS) / $P4), $P3, $P4, count($RS));

      ShowHTML("</tr>");

      ShowHTML("  </td>");
      ShowHTML("</tr>");

    }
    ElseIf (strpos("IAEV", $O) !== false) {
      //   ' Recupera os dados da ocorrência selecionada
      $SQL = "select a.*, " .
             "       case to_char(a.tipo) " .
             "            when '0' then 'Consulta' " .
             "            when '1' then 'Inclusão' " .
             "            when '2' then 'Alteração' " .
             "            when '3' then 'Exclusão' " .
             "            else        'Erro' " .
             "       end nm_tipo, " .
             "       coalesce(b.nome,'Consulta') nm_funcionalidade,   coalesce(b.codigo,0) cd_funcionalidade, " .
             "       c.sq_cliente sq_escola,     c.ds_cliente nm_escola, " .
             "       d.sq_cliente sq_regional,   d.ds_cliente nm_regional, " .
             "       e.sq_cliente sq_secretaria, e.ds_cliente nm_secretaria " .
             "  from sbpi.Cliente_Log                    a " .
             "       left outer join sbpi.Funcionalidade b on (a.sq_funcionalidade = b.sq_funcionalidade and " .
             "                                               b.tipo              = 2 " .
             "                                              ) " .
             "       inner      join sbpi.Cliente        c on (a.sq_cliente        = c.sq_cliente) " .
             "         inner    join sbpi.Cliente         d on (coalesce(c.sq_cliente_pai,0)    = d.sq_cliente) " .
             "           left  join sbpi.Cliente         e on (d.sq_cliente_pai    = e.sq_cliente) " .
             " where sq_cliente_log = " . $_REQUEST['sq_cliente_log'];
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      $RS = $RS[0];
      ShowHTML('<tr bgcolor="#EFEFEF"><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td>Data:<br><b>' . formatadataedicao($RS["data"]) . '</b></td>');
      ShowHTML('          <td>IP de origem:<br><b>' . $RS["ip_origem"] . '</b></td>');
      ShowHTML('        <tr><td colspan=2>&nbsp;');
      ShowHTML('        <tr><td colspan=2>Tipo da ocorrencia: <b>');
      if ($RS['tipo'] == 0)
        ShowHTML('            Acesso');
      if ($RS['tipo'] == 1)
        ShowHTML('            Inclusão');
      if ($RS['tipo'] == 2)
        ShowHTML('            Alteração');
      if ($RS['tipo'] == 3)
        ShowHTML('            Exclusão');

      ShowHTML('            </td>');
      ShowHTML('        <tr><td colspan=2>&nbsp;');
      ShowHTML('        <tr><td colspan=2>Descrição:<br><b>' . $RS["abrangencia"] . '</b></td>');

      If (strlen($RS["query"]) > 0) {

        ShowHTML('        <tr><td colspan=2>&nbsp;');
        ShowHTML('        <tr><td colspan=2>Comando(s) executado(s):<br><b>' . str_replace($crlf, "<br>", str_replace(" ", "&nbsp;", $RS["query"])) . '</b></td>');
      }
      ShowHTML('      <tr><td colspan=2 height=1 bgcolor="#000000">');
      ShowHTML('      <tr><td colspan=2 align="center"><input class="STB" type="button" onClick="history.back(1);" name="Botao" value="Voltar"></td></tr>');
      ShowHTML('    </table>');
      ShowHTML('    </TD>');
      ShowHTML('</tr>');

    } Else {
      ScriptOpen("JavaScript");
      ShowHTML(" alert('Opção não disponível');");
      ShowHTML(" history.back(1);");
      ScriptClose();
    }
    ShowHTML("</table>");
    ShowHTML("</center>");
    Rodape();

  }

  // =========================================================================
  // Cadastro de escolas
  // -------------------------------------------------------------------------
  function CadastroEscolas() {
    extract($GLOBALS);
    //  Set RS2 = Server.CreateObject("ADODB.RecordSet")

    $w_chave = $_REQUEST["p_chave"];
    $w_troca = $_REQUEST["w_troca"];

    if (nvl($w_troca, '') != '') {
      extract($_REQUEST);
      $w_sq_codigo_espec = explode(',', $w_sq_codigo_espec[0]);
    }

    elseIf ((strpos("A", $O) !== false) and ($w_troca == "")) {

      //     ' Recupera os dados do cliente
      $SQL = " SELECT a.sq_cliente, sq_cliente_pai, sq_tipo_cliente, ds_cliente, ds_apelido,a.sq_regiao_adm,a.localizacao as localizacao ," .
             "        no_municipio, sg_uf, ds_username, ln_internet, ds_email, nr_cep, no_contato, " .
             "        ds_email_contato, nr_fone_contato, nr_fax_contato, no_diretor, no_secretario, " .
             "        ds_logradouro, no_bairro, sq_modelo, " .
             "        no_contato_internet, nr_fone_internet, nr_fax_internet, ds_email_internet, " .
             "        ds_diretorio, sq_siscol, ds_mensagem, " .
             "        ds_institucional, ds_texto_abertura, a.ativo " .
             "   FROM sbpi.Cliente a " .
             "        LEFT OUTER JOIN sbpi.Cliente_Dados b on (a.sq_cliente   = b.sq_cliente)  " .
             "        LEFT OUTER JOIN sbpi.Cliente_Site  c on (a.sq_cliente   = c.sq_cliente)  " .
             "  WHERE a.sq_cliente = " . $w_chave;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      $RS = $RS[0];

      $w_chave = $RS["sq_cliente"];
      $w_regiao = $RS["sq_regiao_adm"];
      $w_local  = $RS["localizacao"];
      $w_sq_cliente_pai = $RS["sq_cliente_pai"];
      $w_sq_tipo_cliente = $RS["sq_tipo_cliente"];
      $w_ds_cliente = $RS["ds_cliente"];
      $w_ds_apelido = $RS["ds_apelido"];
      $w_no_municipio = $RS["no_municipio"];
      $w_sg_uf = $RS["sg_uf"];
      $w_ds_username = $RS["ds_username"];
      $w_ln_internet = $RS["ln_internet"];
      $w_ds_email = $RS["ds_email"];
      $w_ds_logradouro = $RS["ds_logradouro"];
      $w_no_bairro = $RS["no_bairro"];
      $w_nr_cep = trim($RS["nr_cep"]);
      $w_no_contato = $RS["no_contato"];
      $w_email_contato = $RS["ds_email_contato"];
      $w_nr_fone_contato = $RS["nr_fone_contato"];
      $w_nr_fax_contato = $RS["nr_fax_contato"];
      $w_no_diretor = $RS["no_diretor"];
      $w_no_secretario = $RS["no_secretario"];
      $w_sq_modelo = $RS["sq_modelo"];
      $w_no_contato_internet = $RS["no_contato_internet"];
      $w_nr_fone_internet = $RS["nr_fone_internet"];
      $w_nr_fax_internet = $RS["nr_fax_internet"];
      $w_ds_email_internet = $RS["ds_email_internet"];
      $w_ds_diretorio = $RS["ds_diretorio"];
      $w_sq_siscol = $RS["sq_siscol"];
      $w_ds_mensagem = $RS["ds_mensagem"];
      $w_ds_institucional = $RS["ds_institucional"];
      $w_ds_texto_abertura = $RS["ds_texto_abertura"];
      $w_username_atual = $RS["ds_username"];
      $w_ativo = $RS["ativo"];
      $w_sq_regiao_adm = $RS["sq_regiao_adm"];
      $w_localizacao = $RS["localizacao"];

    }
    ElseIf (strpos("I", $O) !== false) {

    }
    $SQL = "select ds_username " .
    "  from sbpi.Cliente  a " .
    " where sq_cliente = 0 ";
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    $RS = $RS[0];
    $w_diretorio = $RS["ds_username"] . "/";

    Cabecalho();
    ShowHTML("<HEAD>");
    ScriptOpen("JavaScript");
    modulo();
    checkbranco();
    FormataCEP();
    ShowHTML("function montaLink() {");
    ShowHTML("  var link = '" . $conSite . $conVirtualPath . $w_diretorio . "';");
    ShowHTML("  document.Form.w_ln_internet.value = link + document.Form.w_ds_username.value;");
    ShowHTML("  document.Form.w_ds_diretorio.value = link + document.Form.w_ds_username.value;");
    ShowHTML("}");
    ValidateOpen("Validacao");
    Validate("w_sq_cliente_pai", "D.R.E.", "SELECT", "1", "1", "18", "1", "1");
    Validate("w_regiao", "Região Administrativa", "SELECT", "1", "1", "18", "1", "1");
    Validate("w_sq_tipo_cliente", "Tipo de instituição", "SELECT", "1", "1", "18", "1", "1");
    Validate("w_ds_cliente", "Nome da escola", "1", "1", "3", "60", "1", "1");
    Validate("w_ds_apelido", "Apelido", "1", "", "3", "30", "1", "1");
    Validate("w_ds_username", "Username", "1", "1", "3", "14", "1", "1");
    Validate("w_ln_internet", "Link para acesso", "1", "1", "10", "60", "1", "1");
    Validate("w_ds_email", "e_Mail", "1", "", "4", "60", "1", "1");
    Validate("w_ds_logradouro", "Logradouro", "1", "1", "4", "60", "1", "1");
    Validate("w_no_bairro", "Bairro", "1", "", "2", "30", "1", "1");
    Validate("w_no_municipio", "Cidade", "1", "1", "2", "30", "1", "1");
    Validate("w_sg_uf", "UF", "SELECT", "1", "1", "2", "1", "1");
    Validate("w_nr_cep", "CEP", "1", "1", "9", "9", "", "0123456789-");
    Validate("w_sq_regiao_adm", "Região Administrativa", "SELECT", "1", "1", "18", "1", "1");
    Validate("w_no_contato", "Contato técnico", "1", "1", "2", "35", "1", "1");
    Validate("w_nr_fone_contato", "Telefone contato tencnico", "1", "1", "6", "20", "1", "1");
    Validate("w_nr_fax_contato", "Fax contato técnico", "1", "", "6", "20", "1", "1");
    Validate("w_email_contato", "e-Mail contato técnico", "1", "", "6", "60", "1", "1");
    Validate("w_no_diretor", "Nome do(a) Diretor(a)", "1", "", "2", "40", "1", "1");
    Validate("w_no_secretario", "Nome do(a) Secretário(a)", "1", "", "2", "40", "1", "1");
    Validate("w_no_contato_internet", "Contato internet", "1", "1", "2", "35", "1", "1");
    Validate("w_nr_fone_internet", "Telefone contato internet", "1", "1", "6", "20", "1", "1");
    Validate("w_nr_fax_internet", "Fax contato internet", "1", "", "6", "20", "1", "1");
    Validate("w_ds_email_internet", "e-Mail contato internet", "1", "1", "6", "60", "1", "1");
    Validate("w_sq_modelo", "Modelo de site", "SELECT", "1", "1", "18", "1", "1");
    Validate("w_ds_diretorio", "Diretório", "1", "1", "4", "60", "1", "1");
    Validate("w_sq_siscol", "SISCOL", "1", "", "5", "6", "1", "1");
    Validate("w_ds_mensagem", "Mensagem", "1", "", "4", "80", "1", "1");
    Validate("w_ds_institucional", "Institucional", "1", "", "4", "7000", "1", "1");
    Validate("w_ds_texto_abertura", "Texto de abertura", "1", "", "4", "7000", "1", "1");

    ShowHTML("  theForm.Botao[0].disabled=true;");
    ShowHTML("  theForm.Botao[1].disabled=true;");
    ValidateClose();
    ScriptClose();
    ShowHTML("</HEAD>");
    If ($w_troca > "") {
      BodyOpen("onLoad='document.Form." . $w_troca . ".focus()';");
    }
    ElseIf ($O == "I" or $O == "A") {
      BodyOpen("onLoad='document.Form.w_sq_cliente_pai.focus()';");
    } Else {
      BodyOpen("onLoad='document.focus()';");
    }
    ShowHTML("<B><FONT COLOR='#000000'>Cadastro de escolas da rede de ensino</FONT></B>");
    ShowHTML("<HR>");
    ShowHTML("<div align=center><center>");
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    If (strpos("EV", $O) !== false) {
      $w_disabled = " DISABLED ";
    }
    ShowHTML('<FORM action="' . $w_pagina . 'Grava" method="POST" name="Form" onSubmit="return(Validacao(this));">');
    ShowHTML(MontaFiltro("POST"));
    ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');
    ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
    ShowHTML('<INPUT type="hidden" name="SG" value="' . $par . '">');
    //  ShowHTML ('<INPUT type="hidden" name="w_ea" value="" & w_ea & "">"
    If ($O == "A") {
      ShowHTML('<INPUT type="hidden" name="w_username_atual" value="' . $w_ds_username . '">');
    }
    ShowHTML('<tr bgcolor="#EFEFEF"><td align="center">');
    ShowHTML('    <table width="100%" border="0">');
    ShowHTML('        <tr valign="top">');
    SelecaoRegional("Regional de en<u>s</u>ino:", "S", "Indique a D.R.E. da escola.", $w_sq_cliente_pai, null, "w_sq_cliente_pai", "CADASTRO", null);
    SelecaoRegiaoAdm("Região a<u>d</u>ministrativa:", "D", "Indique a região administrativa.", $w_regiao, $w_chave, "w_regiao", null, null);
    ShowHTML('        <tr valign="top">');
    $SQL = "SELECT * FROM sbpi.Tipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente";
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    ShowHTML('              <td valign="top"><font size="1"><b>Tipo de in<u>s</u>tituição:</b><br><SELECT accesskey="S" class="STI" NAME="w_sq_tipo_cliente">');
    if (count($RS) > 0)
      ShowHTML('          <option value="">---');
    foreach ($RS as $row) {
      If ($row["sq_tipo_cliente"] == $w_sq_tipo_cliente) {
        ShowHTML('          <option value="' . $row["sq_tipo_cliente"] . '" SELECTED>' . $row["ds_tipo_cliente"]);
      } Else {
        ShowHTML('          <option value="' . $row["sq_tipo_cliente"] . '">' . $row["ds_tipo_cliente"]);
      }
    }
    ShowHTML('          </select>');
    ShowHTML('          <td><b>Localização</b><br>');
    If ($w_local == "1") {
      ShowHTML('              <input  type="radio" name="w_local" value="0"> Não informada <input  type="radio" name="w_local" value="1" checked> Urbana <input  type="radio" name="w_local" value="2"> Rural ');
    }
    ElseIf ($w_local == "2") {
      ShowHTML('              <input  type="radio" name="w_local" value="0"> Não informada <input  type="radio" name="w_local" value="1"> Urbana <input  type="radio" name="w_local" value="2" checked> Rural ');
    } Else {
      ShowHTML('              <input  type="radio" name="w_local" value="0" checked> Não informada <input  type="radio" name="w_local" value="1"> Urbana <input  type="radio" name="w_local" value="2"> Rural ');
    }
    
    ShowHTML('      <tr><td valign="top"><font size="1"><b>Nome <u>d</u>a escola:</b><br><input ' . $w_disabled . ' accesskey="D" type="text" name="w_ds_cliente" class="STI" SIZE="60" MAXLENGTH="60" VALUE="' . $w_ds_cliente . '" title="Informe uma descrição para a escola."></td>');
    ShowHTML('          <td valign="top"><font size="1"><b><u>A</u>pelido:</b><br><input ' . $w_disabled . ' accesskey="A" type="text" name="w_ds_apelido" class="STI" SIZE="30" MAXLENGTH="30" VALUE="' . $w_ds_apelido . '" title="Informe o apelido da escola."></td>');

    ShowHTML('      <tr><td valign="top"><font size="1"><b><u>U</u>sername:</b><br><input ' . $w_disabled . ' accesskey="U" type="text" name="w_ds_username" class="STI" SIZE="14" MAXLENGTH="14" VALUE="' . $w_ds_username . '" title="Informe o usernanme para acesso." ONBLUR="javascript:montaLink();"></td>');
    MontaRadioSN("<b>Ativo?</b>", $w_ativo, "w_ativo");
    ShowHTML('      <tr><td valign="top"><font size="1"><b><u>L</u>ink para acesso:</b><br><input ' . $w_disabled . ' accesskey="L" type="text" name="w_ln_internet" class="STI" SIZE="60" MAXLENGTH="60" VALUE="' . $w_ln_internet . '" title="Informe link de acesso para o site da escola."></td>');
    ShowHTML('          <td valign="top"><font size="1"><b><u>e</u>-Mail:</b><br><input ' . $w_disabled . ' accesskey="E" type="text" name="w_ds_email" class="STI" SIZE="50" MAXLENGTH="60" VALUE="' . $w_ds_email . '" title="Informe e-Mail da escola."></td>');
    ShowHTML('      <tr><td valign="top"><font size="1"><b>Lo<u>g</u>radouro:</b><br><input ' . $w_disabled . ' accesskey="G" type="text" name="w_ds_logradouro" class="STI" SIZE="50" MAXLENGTH="60" VALUE="' . $w_ds_logradouro . '" title="Informe o logradouro da escola."></td>');
    ShowHTML('          <td valign="top"><font size="1"><b><u>B</u>airro:</b><br><input ' . $w_disabled . ' accesskey="B" type="text" name="w_no_bairro" class="STI" SIZE="30" MAXLENGTH="60" VALUE="' . $w_no_bairro . '" title="Informe o bairro da escola."></td>');
    ShowHTML('      <tr valign="top"><td colspan="2"><font size="1">');
    ShowHTML('        <table width="100%" border="0">');
    ShowHTML('          <tr><td valign="top"><font size="1"><b><u>C</u>idade:</b><br><input ' . $w_disabled . '  accesskey="C" type="text" name="w_no_municipio" class="STI" SIZE="30" MAXLENGTH="30" VALUE="' . $w_no_municipio . '" title="Informe a cidade."></td>');

    switch ($w_sg_uf) {
      Case "AC" :
        $AC = "selected";
        break;
      Case "AM" :
        $AM = "selected";
        break;
      Case "AP" :
        $AP = "selected";
        break;
      Case "AL" :
        $AL = "selected";
        break;
      Case "BA" :
        $BA = "selected";
        break;
      Case "CE" :
        $CE = "selected";
        break;
      Case "DF" :
        $DF = "selected";
        break;
      Case "ES" :
        $ES = "selected";
        break;
      Case "GO" :
        $GO = "selected";
        break;
      Case "MA" :
        $MA = "selected";
        break;
      Case "MT" :
        $MT = "selected";
        break;
      Case "MS" :
        $MS = "selected";
        break;
      Case "MG" :
        $MG = "selected";
        break;
      Case "PA" :
        $PA = "selected";
        break;
      Case "PB" :
        $PB = "selected";
        break;
      Case "PR" :
        $PR = "selected";
        break;
      Case "PE" :
        $PE = "selected";
        break;
      Case "PI" :
        $PI = "selected";
        break;
      Case "RJ" :
        $RJ = "selected";
        break;
      Case "RN" :
        $RN = "selected";
        break;
      Case "RO" :
        $RO = "selected";
        break;
      Case "RR" :
        $RR = "selected";
        break;
      Case "RS" :
        $RGS = "selected";
        break;
      Case "SC" :
        $SC = "selected";
        break;
      Case "SP" :
        $SP = "selected";
        break;
      Case "SE" :
        $SE = "selected";
        break;
      Case "TO" :
        $TOC = "selected";
        break;
    }
    ShowHTML('         <td valign="top"><font  size="1"><b>U<U>F</U>:</b><br>');
    ShowHTML('             <select ACCESSKEY="F" name="w_sg_uf" class="STI">');
    ShowHTML('                <option value="">---</option>');
    ShowHTML('                <option value="AC" ' . $AC . '>AC</option>');
    ShowHTML('                <option value="AM" ' . $AM . '>AM</option>');
    ShowHTML('                <option value="AP" ' . $AP . '>AP</option>');
    ShowHTML('                <option value="AL" ' . $AL . '>AL</option>');
    ShowHTML('                <option value="BA" ' . $BA . '>BA</option>');
    ShowHTML('                <option value="CE" ' . $CE . '>CE</option>');
    ShowHTML('                <option value="DF" ' . $DF . '>DF</option>');
    ShowHTML('                <option value="ES" ' . $ES . '>ES</option>');
    ShowHTML('                <option value="GO" ' . $GO . '>GO</option>');
    ShowHTML('                <option value="MA" ' . $MA . '>MA</option>');
    ShowHTML('                <option value="MT" ' . $MT . '>MT</option>');
    ShowHTML('                <option value="MS" ' . $MS . '>MS</option>');
    ShowHTML('                <option value="MG" ' . $MG . '>MG</option>');
    ShowHTML('                <option value="PA" ' . $PA . '>PA</option>');
    ShowHTML('                <option value="PB" ' . $PB . '>PB</option>');
    ShowHTML('                <option value="PR" ' . $PR . '>PR</option>');
    ShowHTML('                <option value="PE" ' . $PE . '>PE</option>');
    ShowHTML('                <option value="PI" ' . $PI . '>PI</option>');
    ShowHTML('                <option value="RJ" ' . $RJ . '>RJ</option>');
    ShowHTML('                <option value="RN" ' . $RN . '>RN</option>');
    ShowHTML('                <option value="RO" ' . $RO . '>RO</option>');
    ShowHTML('                <option value="RR" ' . $RR . '>RR</option>');
    ShowHTML('                <option value="RS" ' . $RGS . '>RS</option>');
    ShowHTML('                <option value="SC" ' . $SC . '>SC</option>');
    ShowHTML('                <option value="SP" ' . $SP . '>SP</option>');
    ShowHTML('                <option value="SE" ' . $SE . '>SE</option>');
    ShowHTML('                <option value="TO" ' . $TOC . '>TO</option>');
    ShowHTML('             </select>');

    ShowHTML('          <td valign="top"><font size="1"><b>Ce<u>p</u>:</b></font><br><input ' . $w_disabled . ' accesskey="P" type="text" name="w_nr_cep" class="sti" SIZE="9" MAXLENGTH="9" VALUE="' . $w_nr_cep . '" onKeyDown="FormataCEP(this,event)" title="Informe o CEP deste endereço."></td></tr>');

    ShowHTML('      <tr><td valign="top" ><b>Localização</b><br>');

    if ($w_localizacao == 2) {
      ShowHTML('      <input  type="radio" name="w_localizacao" value="1" > Urbana <input  type="radio" name="w_localizacao" value="2" checked> Rural ');
    } else {
      ShowHTML('      <input  type="radio" name="w_localizacao" value="1" checked> Urbana <input  type="radio" name="w_localizacao" value="2"> Rural ');
    }
    SelecaoRegiaoAdm("<td valign='top'><b>Região adminis<u>t</u>rativa:<b>", "D", "Indique a região administrativa.", $w_sq_regiao_adm, $p_escola, "w_sq_regiao_adm", null, null);

    ShowHTML('      <tr><td valign="top"><font size="1"><b>Co<u>n</u>tato técnico:</b><br><input ' . $w_disabled . ' accesskey="N" type="text" name="w_no_contato" class="STI" SIZE="35" MAXLENGTH="35" VALUE="' . $w_no_contato . '" title="Informe o nome do contato técnico da escola."></td>');
    ShowHTML('          <td valign="top"><font size="1"><b><u>T</u>elefone contato técnico:</b><br><input ' . $w_disabled . ' accesskey="T" type="text" name="w_nr_fone_contato" class="STI" SIZE="13" MAXLENGTH="20" VALUE="' . $w_nr_fone_contato . '" title="Informe o telefone do contato técnico da escola."></td>');
    ShowHTML('          <td valign="top"><font size="1"><b>Fa<u>x</u> contato técnico:</b><br><input ' . $w_disabled . ' accesskey="X" type="text" name="w_nr_fax_contato" class="STI" SIZE="13" MAXLENGTH="20" VALUE="' . $w_nr_fax_contato . '" title="Informe o fax do contato técnico da escola."></td>');
    ShowHTML('      <tr><td valign="top"><font size="1"><b>e-Mail contato técnic<u>o</u>:</b><br><input ' . $w_disabled . ' accesskey="O" type="text" name="w_email_contato" class="STI" SIZE="35" MAXLENGTH="60" VALUE="' . $w_email_contato . '" title="Informe o e-Mail do contato técnico da escola."></td>');
    ShowHTML('          <td valign="top"><font size="1"><b>No<u>m</u>e do(a) Diretor(a):</b><br><input ' . $w_disabled . ' accesskey="M" type="text" name="w_no_diretor" class="STI" SIZE="40" MAXLENGTH="40" VALUE="' . $w_no_diretor . '" title="Informe o nome do diretor da escola."></td>');
    ShowHTML('          <td valign="top"><font size="1"><b>No<u>m</u>e do(a) <u>S</u>ecretário(a):</b><br><input ' . $w_disabled . ' accesskey="M" type="text" name="w_no_secretario" class="STI" SIZE="40" MAXLENGTH="40" VALUE="' . $w_no_secretario . '" title="Informe o telefone do contato técnico da escola."></td>');
    ShowHTML('      <tr><td valign="top"><font size="1"><b>Co<u>n</u>tato internet:</b><br><input ' . $w_disabled . ' accesskey="N" type="text" name="w_no_contato_internet" class="STI" SIZE="35" MAXLENGTH="35" VALUE="' . $w_no_contato_internet . '" title="Informe o nome do contato na internet da escola."></td>');
    ShowHTML('          <td valign="top"><font size="1"><b><u>T</u>elefone contato internet:</b><br><input ' . $w_disabled . ' accesskey="Y" type="text" name="w_nr_fone_internet" class="STI" SIZE="13" MAXLENGTH="20" VALUE="' . $w_nr_fone_internet . '" title="Informe o telefone do contato na internet da escola."></td>');
    ShowHTML('          <td valign="top"><font size="1"><b>Fa<u>x</u> contato internet:</b><br><input ' . $w_disabled . ' accesskey="X" type="text" name="w_nr_fax_internet" class="STI" SIZE="13" MAXLENGTH="20" VALUE="' . $w_nr_fax_internet . '" title="Informe o fax do contato na internet da escola."></td>');
    ShowHTML('      </table>');
    ShowHTML('      <tr valign="top"><td colspan="2"><font size="1">');
    ShowHTML('        <table width="100%" border="0">');
    ShowHTML('          <td valign="top"><font size="1"><b><u>e</u>-Mail contato internet:</b><br><input ' . $w_disabled . ' accesskey="E" type="text" name="w_ds_email_internet" class="STI" SIZE="35" MAXLENGTH="60" VALUE="' . $w_ds_email_internet . '" title="Informe o e-Mail do contato na internet da escola."></td>');

    $SQL = "SELECT a.sq_modelo, a.ds_modelo FROM sbpi.Modelo a ORDER BY a.sq_modelo desc";
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    ShowHTML('              <td valign="top"><font size="1"><b>Modelo de site:</b><br><SELECT class="STI" NAME="w_sq_modelo">');
    If (count($RS) > 0)
      ShowHTML('          <option value="">---');
    foreach ($RS as $row) {
      If ($row["sq_modelo"] == $w_sq_modelo) {
        ShowHTML('          <option value="' . $row["sq_modelo"] . '" SELECTED>' . $row["ds_modelo"]);
      } Else {
        ShowHTML('          <option value="' . $row["sq_modelo"] . '" >' . $row["ds_modelo"]);
      }
    }
    ShowHTML('              </select>');

    ShowHTML('              <td valign="top"><font size="1"><b><u>D</u>iretório:</b><br><input ' . $w_disabled . ' accesskey="D" type="text" name="w_ds_diretorio" class="STI" SIZE="40" MAXLENGTH="60" VALUE="' . $w_ds_diretorio . '" title="Informe o diretório físico da escola."></td>');
    ShowHTML('              <td valign="top"><font size="1"><b>SIS<u>C</u>OL:</b><br><input ' . $w_disabled . ' accesskey="C" type="text" name="w_sq_siscol" class="STI" SIZE="6" MAXLENGTH="6" VALUE="' . $w_sq_siscol . '" title="Informe o número siscol da escola."></td>');
    ShowHTML('      </table>');
    ShowHTML('      <tr><td valign="top"><font size="1"><b><u>M</u>ensagem:</b><br><input ' . $w_disabled . ' accesskey="M" type="text" name="w_ds_mensagem" class="STI" SIZE="80" MAXLENGTH="80" VALUE="' . Nvl($w_ds_mensagem, "Brasília - Patrimônio Cultural da Humanidade") . '" title="Informe o mensagem rolante da tela da escola."></td>');
    ShowHTML('      <tr><td valign="top" colspan="2"><font  size="1"><b><U>I</U>nstitucional:</b><br><TEXTAREA ACCESSKEY="I" ' . $w_disabled . ' class="STI" type="text" name="w_ds_institucional" ROWS=4 COLS=76>' . Nvl($w_ds_institucional, "Desenvolvemos essa solução a fim de prestar serviços eficientes e gratuitos à comunidade, criando um espaço de intercâmbio com a escola para, juntos, trabalharmos pela melhora do ensino público.") . "</TEXTAREA ></td>");
    ShowHTML('      <tr><td valign="top" colspan="2"><font  size="1"><b><U>T</U>exto de abertura:</b><br><TEXTAREA ACCESSKEY="X" ' . $w_disabled . ' class="STI" type="text" name="w_ds_texto_abertura" ROWS=4 COLS=76>' . Nvl($w_ds_texto_abertura, "O ensino moderno oferece amplo apoio tecnológico ao estudante. A Internet modificou, substancialmente, o paradigma existente na área educacional. Modernizamos nossas atividades a fim de garantir a nossos alunos, pais e responsáveis qualidade, eficiência e rapidez na prestação de nossos serviços.") . "</TEXTAREA ></td>");

    $SQL = "SELECT a.*,c.sq_cliente " .
           "  from sbpi.Especialidade a " .
           "       LEFT JOIN sbpi.Especialidade_cliente  c ON (a.sq_especialidade = c.sq_especialidade and " .
           "                                     c.sq_cliente  = " . Nvl($w_chave, 0) . ") " .
           "       LEFT JOIN sbpi.Cliente  d ON (c.sq_cliente = d.sq_cliente  and " .
           "                                     d.sq_cliente  = " . Nvl($w_chave, 0) . ") " .
           " where a.tp_especialidade <> 'M' or a.tp_especialidade = 'J' " .
           " ORDER BY a.nr_ordem, a.ds_especialidade ";

    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    if (count($RS) > 0) {
      $wAtual = "";

      ShowHTML('          <tr><TD colspan=2><table border="0" align="left" cellpadding=0 cellspacing=0>');
      foreach ($RS as $row) {
        If ($wAtual == "" or $wAtual != $row["tp_especialidade"]) {
          $wAtual = $row["tp_especialidade"];
          If ($wAtual == "M") {
            ShowHTML('            <TR><TD colspan=2><font size="1" CLASS="BTM"><b>Etapas/Modalidades de ensino:</b>');
          }
          ElseIf ($wAtual == "R") {
            ShowHTML('            <TR><TD colspan=2><font size="1" CLASS="BTM"><b>Em Regime de Intercomplementaridade:</b>');
          } Else {
            ShowHTML('            <TR><TD colspan=2><font size="1" CLASS="BTM"><b>Outras:</b>');
          }
        }

        If (trim($row["sq_cliente"]) != "" or (strlen(trim($w_troca)) > 0 and in_array($row["sq_especialidade"], $w_sq_codigo_espec))) {

          ShowHTML(chr(13) . '           <tr><td><input type="checkbox" name="w_sq_codigo_espec[]" value="' . $row["sq_especialidade"] . '" checked><td><font size=1>' . $row["ds_especialidade"]);
        } Else {
          ShowHTML(chr(13) . '           <tr><td><input type="checkbox" name="w_sq_codigo_espec[]" value="' . $row["sq_especialidade"] . '" ><td><font size=1>' . $row["ds_especialidade"]);
        }

      }
    }

    ShowHTML('      </table>');
    ShowHTML('      <tr><td align="center" colspan=4><hr>');
    If ($O == "E") {
      ShowHTML('   <input class="STB" type="submit" name="Botao" value="Excluir" onClick="return confirm(\'Confirma a exclusão do registro?\');">');
    } Else {
      If ($O == "I") {
        ShowHTML('            <input class="STB" type="submit" name="Botao" value="Incluir">');
      } Else {
        ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
      }
    }
    ShowHTML("            <input class='STB' type='button' onClick='location.href=\"" . $w_pagina . "escolas&O=L&pesquisa=Sim" . MontaFiltro("GET") . "\";' name='Botao' value='Cancelar'>");
    ShowHTML('          </td>');
    ShowHTML('      </tr>');
    ShowHTML('    </table>');
    ShowHTML('    </TD>');
    ShowHTML('</tr>');
    ShowHTML('</FORM>');
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();
  }
  // =========================================================================
  // Rotina principal
  // -------------------------------------------------------------------------
  function Main() {
    extract($GLOBALS);
    switch ($par) {
      Case 'MENU'           : Menu();               break;
      Case 'ADMINISTRATIVO' : Administrativo();     break;
      Case 'ADMLOG'         : AdmLog();             break;
      Case 'ARQUIVOS'       : Arquivos();           break;
      Case 'SHOW_LOG'       : showlog();            break;
      Case 'CALEND_REDE'    : Calend_rede();        break;
      Case 'CALEND_BASE'    : Calend_base();        break;
      Case 'ESCPART'        : escPart();            break;
      Case 'ESCOLAS'        : escolas();            break;
      Case 'ESCPARTHOMOLOG' : escPartHomolog();     break;
      Case 'GRAVA'          : Grava();              break;
      Case 'LOG'            : log_controle();       break;
      Case 'DADOSESCOLA'    : dadosEscola();        break;
      Case 'MSGALUNOS'      : Msgalunos();          break;
      Case 'MODALIDADES'    : Modalidades();        break;
      Case 'NEWSLETTER'     : Newsletter();         break;
      Case 'NOTICIAS'       : Noticias();           break;
      Case 'CADASTROESCOLA' : CadastroEscolas();    break;
      Case 'SENHA'          : Senha();              break;
      Case 'SENHAESP'       : Senhaesp();           break;
      Case 'TIPOCLIENTE'    : TipoCliente();        break;
      Case 'VALIDA'         : Valida();             break;
      Case 'FRAMES'         : Frames();             break;
      Case 'REDEPART'       : redepart();           break;
      Case 'SAIR'           : Sair();               break;
      default :
        Cabecalho();
        ShowHTML('<BASE HREF="' . $conRootSIW . '">');
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