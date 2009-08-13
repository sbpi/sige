<?php

  // Garante que a sessão será reinicializada.
  session_start();

  $w_dir_volta = '';
  $w_pagina = 'manut.php?par=';
  $w_Disabled = 'ENABLED';
  $w_dir = '';
  $w_troca = $_REQUEST['w_troca'];

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

  // Recupera o filtro selecionado
  $P3 = nvl($_REQUEST['P3'], 1);
  $P4 = nvl($_REQUEST['P4'], $conPageSize);

  // Abre conexão com o banco de dados
  if (isset ($_SESSION['DBMS'])) {
    $dbms = abreSessao :: getInstanceOf($_SESSION['DBMS']);
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
  $RS = null;
  $CL = $_SESSION['CL'];
  $par = substr(strtoupper($_REQUEST['par']), 0, 30);
  $O = substr(strtoupper($_REQUEST['O']), 0, 1);

  if (nvl($O, '') == '') {
    $O = 'L';
  }

  $w_Data = date('d/m/Y');

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
    ShowHTML('  $("#Login1").change(function(){');
    ShowHTML('    formataCampo();');
    ShowHTML('  })');
    ShowHTML('});');
    ShowHTML('function caracterAceito(string , checkOK){');
    ShowHTML('   //var checkOK = \'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-!@#$%.*()+/\';');
    ShowHTML('      var checkStr = string;');
    ShowHTML('      var allValid = true;');
    ShowHTML('      for (i = 0;  i < checkStr.length;  i++)');
    ShowHTML('      {');
    ShowHTML('      ch = checkStr.charAt(i);');
    ShowHTML('      if ((checkStr.charCodeAt(i) != 13) && (checkStr.charCodeAt(i) != 10) && (checkStr.charAt(i) != "\\\\")) {');
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
    ShowHTML('<form method="post" action="manut.php?par=valida" onsubmit="return(Validacao(this));" name="Form"> ');
    ShowHTML('<INPUT TYPE="HIDDEN" NAME="Login" VALUE=""> ');
    ShowHTML('<INPUT TYPE="HIDDEN" NAME="Password" VALUE=""> ');
    ShowHTML('<INPUT TYPE="HIDDEN" NAME="p_dbms" VALUE="1"> ');
    ShowHTML('<TABLE cellSpacing=0 cellPadding=0 width="760" height=550 border=1  background="files/" . p_cliente . "/img/fundo.jpg" bgproperties="fixed"><tr><td width="100%" valign="top">');
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
    ShowHTML('              . <a class="SS" href="manuais/regionais/" target="_blank" title="Exibe versão HTML do manual de operação do SIGE-WEB (Diretorias Regionais de Ensino)">Manual SIGE-WEB para DRE</A><BR>');
    ShowHTML('              . <a class="SS" href="manuais/operacao/" target="_blank" title="Exibe  versão HTML do manual de operação do SIGE-WEB (Outros)">Manual SIGE-WEB para demais I.E.</A><BR>');
    ShowHTML('              <br></FONT></P>');
    ShowHTML('              </TD></TR>');
    ShowHTML('</table>');
    ShowHTML('              </TD></TR>');
    ShowHTML('          </table>   ');

    $wAno = $_REQUESR["wAno"];

    If ($wAno == "") {
      $wAno = date('Y');
    }

    $SQL = "SELECT case in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end  in_destinatario, " . $crlf .
           "       dt_arquivo, ds_titulo, nr_ordem, ds_arquivo, ln_arquivo, 'SEDF' AS Origem, x.ln_internet diretorio " . $crlf .
           " From sbpi.Cliente_Arquivo a INNER JOIN sbpi.Cliente  x ON (a.sq_cliente = x.sq_cliente)" . $crlf .
           " WHERE x.ativo = 'S'" . $crlf .
           "  AND x.sq_cliente = 0" . $crlf .
           "  AND in_destinatario = 'E'" . $crlf .
           "  and sbpi.YEAR(a.dt_arquivo) = " . $wAno . $crlf .
           " ORDER BY origem, nr_ordem, dt_arquivo desc, in_destinatario " . $crlf;

    ShowHTML(' <tr><td><table width="100%" border="0">');
    ShowHTML('          <TR><TD align="center"><br><table border=0 cellpadding=0 cellspacing=0><tr><td>');
    ShowHTML('              <P><font face="Arial" size=1><b>ARQUIVOS INSERIDOS PELA SEDF</b></font><br>');
    ShowHTML('<tr><td><TABLE border=0 cellSpacing=5 width="95%">');
    ShowHTML('  <TR>');
    ShowHTML('    <TD><FONT face="Verdana" size=1><b>Origem');
    ShowHTML('    <TD><FONT face="Verdana" size=1><b>Alvo');
    ShowHTML('    <TD><FONT face="Verdana" size=1><b>Data');
    ShowHTML('    <TD><FONT face="Verdana" size=1><b>Componente curricular');
    ShowHTML('  </TR>');
    ShowHTML('  <TR>');
    ShowHTML('    <TD COLSPAN="4" HEIGHT="1" BGCOLOR="#DAEABD">');
    ShowHTML('  </TR>');

    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    if (count($RS) > 0) {
      foreach ($RS as $row) {
        ShowHTML('  <TR valign="top">');
        ShowHTML('    <TD><FONT face="Verdana" size=1>' . $row["origem"]);
        ShowHTML('    <TD><FONT face="Verdana" size=1>' . $row["in_destinatario"]);
        ShowHTML('    <TD><FONT face="Verdana" size=1>' . formatadataedicao($row["dt_arquivo"]));
        ShowHTML('    <TD><FONT face="Verdana" size=1><a href="sedf/' . $row["ln_arquivo"] . '" target="_blank">' . $row["ds_titulo"] . '</a><br><div align="justify"><font size=1>.:. ' . $row["ds_arquivo"] . '</div>');
        ShowHTML('  </TR>');
      }
    } Else {
      ShowHTML('  <TR><TD COLSPAN=4 ALIGN=CENTER><FONT face="Verdana" size=1>Não há arquivos disponíveis no momento para o ano de ' . $wAno . ' </TR>');
    }

    $SQL = "SELECT distinct sbpi.year(dt_arquivo) ano " . $crlf .
           " From sbpi.Cliente_Arquivo a INNER JOIN sbpi.Cliente x ON (a.sq_cliente = x.sq_cliente)" . $crlf .
           " WHERE x.ativo = 'S'" . $crlf .
           "  AND x.sq_cliente = 0" . $crlf .
           "  AND in_destinatario = 'E'" . $crlf .
           "  and sbpi.YEAR(a.dt_arquivo) <> " . $wAno . $crlf .
           " ORDER BY sbpi.year(dt_arquivo) desc " . $crlf;

    //var_dump($SQL);

    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    if (count($RS) > 0) {
      ShowHTML('  <TR><TD COLSPAN=4 ><FONT face="Verdana" size=1><b>Arquivos de outros anos</b><br>');
      foreach ($RS as $row) {
        ShowHTML('     <li><a href="' . $w_dir . 'Manut.php?$wAno=' . $row["ano"] . '" >Arquivos de ' . $row["ano"] . '</a>');
      }
      ShowHTML('</TD></TR>');
    }

    ShowHTML('</table>');
    ShowHTML('              </FONT></P>');
    ShowHTML('              </TD></TR>');
    ShowHTML('          </table>');
    ShowHTML('        </table>');
    ShowHTML('    </tr>');
    ShowHTML('  </table>');
    ShowHTML('</table>');
    ShowHTML('</form>');
    ShowHTML('</CENTER>');
    ShowHTML('</body>');
    ShowHTML('</html>');

  }

  // =========================================================================
  // Rotina de autenticação dos usuários
  // -------------------------------------------------------------------------
  function Valida() {
    extract($GLOBALS);

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
    ScriptOpen('JavaScript');
    if ($w_erro > 0) {
      if ($w_erro == 1)
        ShowHTML(' alert("Usuário inexistente!");');
      elseif ($w_erro == 2) ShowHTML(' alert("Senha inválida!");');

      ShowHTML('  history.back(1);');
    } else {
      //Recupera informações a serem usadas na montagem das telas para o usuário
      $SQL = "select a.ds_username , a.sq_cliente , b.tipo from sbpi.Cliente a inner join sbpi.tipo_cliente b on a.sq_tipo_cliente = b.sq_tipo_cliente where upper(ds_username) = upper('" . $w_uid . "') and upper(ds_senha_acesso) = upper('" . $w_pwd . "')";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }

      $_SESSION['USERNAME'] = strtoupper(f($RS, 'ds_username'));
      $_SESSION['CL'] = f($RS, 'sq_cliente');
      $_SESSION['TIPO'] = f($RS, 'tipo');

      If ($_SESSION['USERNAME'] != 'SBPI') {
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
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      }
      ShowHTML(' location.href="' . $w_pagina . 'frames";');
    }
    ScriptClose();
  }

  // =========================================================================
  // Monta a Frame
  // -------------------------------------------------------------------------
  Function frames() {
    extract($GLOBALS);
    $_SESSION["BodyWidth"] = "620";

    ShowHTML('<html>');
    ShowHTML('<head>');
    ShowHTML('    <title>Controle Central</title>');
    ShowHTML('</head>');
    ShowHTML('<frameset cols="200,*">');
    ShowHTML('    <frame name="menu" src="' . $w_pagina . 'menu" scrolling="auto" marginheight="0" marginwidth="0">');
    ShowHTML('    <frame name="content" src="' . $w_pagina . 'getsite" scrolling="auto" marginheight="0" marginwidth="0">');
    ShowHTML('</frameset>');
    ShowHTML('</html>');
  }

  function escolas() {
    extract($GLOBALS);
    echo "aaaa";
  }

  // =========================================================================
  // Tela de dados do site
  // -------------------------------------------------------------------------
  function GetSite() {
    extract($GLOBALS);

    $SQL = "select a.*, b.ds_diretorio tipo from sbpi.Cliente_Site a inner join sbpi.Modelo b on (a.sq_modelo = b.sq_modelo) where a.sq_cliente = " . $CL;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    foreach ($RS as $row) {
      $RS = $row;
      break;
    }

    $w_sq_cliente = $RS["sq_cliente"];
    $w_no_contato_internet = $RS["no_contato_internet"];
    $w_ds_email_internet = $RS["ds_email_internet"];
    $w_nr_fone_internet = $RS["nr_fone_internet"];
    $w_nr_fax_internet = $RS["nr_fax_internet"];
    $w_ds_texto_abertura = $RS["ds_texto_abertura"];
    $w_ds_institucional = $RS["ds_institucional"];
    $w_ds_mensagem = $RS["ds_mensagem"];
    $w_ds_diretorio = $RS["ds_diretorio"];
    $w_pedagogica = $RS["ln_prop_pedagogica"];
    //$w_tipo                 = str_replace("Mod","",$RS["tipo"]);

    Cabecalho();
    ShowHTML("<HEAD>");
    ScriptOpen("Javascript");
    ValidateOpen("Validacao");

    Validate("w_no_contato_internet", "Nome", "1", "1", "2", "35", "1", "1");
    Validate("w_ds_email_internet", "e-Mail", "1", "1", "6", "60", "1", "1");
    Validate("w_nr_fone_internet", "Telefone", "1", "1", "6", "20", "1", "1");
    Validate("w_nr_fax_internet", "Fax", "1", "", "6", "20", "1", "1");
    Validate("w_ds_texto_abertura", "Texto de abertura", "1", "1", "4", "8000", "1", "1");
    Validate("w_ds_institucional", "Texto da seção \"Quem somos\" ", "1", "1", "4", "8000", "1", "1");
    if ($w_tipo != 13) {
      If ($w_tipo == 2) { //' Se for regional
        Validate("w_pedagogica", "Composição administrativa", "1", "", "5", "100", "1", "1");
      } Else {
        Validate("w_pedagogica", "Projeto", "1", "", "5", "100", "1", "1");
      }
      ShowHTML(" if (theForm.w_pedagogica.value != ''){");
      ShowHTML("    if((theForm.w_pedagogica.value.lastIndexOf('.PDF')==-1) && (theForm.w_pedagogica.value.lastIndexOf('.pdf')==-1) && (theForm.w_pedagogica.value.lastIndexOf('.DOC')==-1) && (theForm.w_pedagogica.value.lastIndexOf('.doc')==-1)) {");
      ShowHTML("       alert('Esolha arquivos com as extensões \'.doc\' ou \'.pdf\'!');");
      ShowHTML("       theForm.w_pedagogica.value=''; ");
      ShowHTML("       theForm.w_pedagogica.focus(); ");
      ShowHTML("       return false;");
      ShowHTML("    }");
      ShowHTML("  }");
    }
    Validate("w_ds_mensagem", "Texto da mensagem em destaque", "1", "1", "4", "80", "1", "1");

    ValidateClose();
    ScriptClose();
    ShowHTML("</HEAD>");
    BodyOpen("onLoad='document.Form.w_no_contato_internet.focus();'");
    ShowHTML('<B><FONT COLOR="#000000">Atualização de dados do site da unidade</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
    ShowHTML('<FORM action="manut.php?par=Grava" method="POST" name="Form" onSubmit="return(Validacao(this));" enctype="multipart/form-data">');
    ShowHTML(MontaFiltro("POST"));
    ShowHTML('<input type="hidden" name="SG" value="getsite">');
    ShowHTML('<tr bgcolor="" . "#EFEFEF" . ""><td align="center">');
    ShowHTML('    <table width="97%" border="0">');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Contato da unidade para divulgação no site</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Informe os dados da pessoa a ser exibida como contato da unidade para divulgação no site.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML('        <tr valign="top">');
    ShowHTML('          <td><font size="1"><b>No<u>m</u>e:</b><br><INPUT ACCESSKEY="M"  ' . $w_Disabled . ' class="STI" type="text" name="w_no_contato_internet" size="35" maxlength="35" value="' . $w_no_contato_internet . '" title="OBRIGATÓRIO. Informe o nome da pessoa a ser exibida na seção \"Quem somos\" do site da unidade."></td>');
    ShowHTML('          <td><font size="1"><b><u>e</u>-Mail:</b><br><INPUT ACCESSKEY="E" ' . $w_Disabled . ' class="STI" type="text" name="w_ds_email_internet" size="40" maxlength="60" value="' . $w_ds_email_internet . '" title="OBRIGATÓRIO. Informe o e-mail da pessoa." ></td');
    ShowHTML('        <tr valign="top">');
    ShowHTML('          <td><font size="1"><b>Tele<u>f</u>one:</b><br><INPUT ACCESSKEY="F" ' . $w_Disabled . ' class="STI" type="text" name="w_nr_fone_internet" size="20" maxlength="20" value="' . $w_nr_fone_internet . '" title="OBRIGATÓRIO. Informe o telefone da pessoa."></td>');
    ShowHTML('          <td><font size="1"><b>Fa<u>x</u>:</b><br><INPUT ACCESSKEY="X" ' . $w_Disabled . ' class="STI" type="text" name="w_nr_fax_internet" size="20" maxlength="20" value="' . $w_nr_fax_internet . '" title="OPCIONAL. Informe o fax da pessoa."></td>');
    ShowHTML('        </table>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Página de abertura do site</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Informe o texto a ser colocado na página de abertura do site.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size="1"><b>Texto da <u>p</u>ágina de abertura:</b><br><TEXTAREA ACCESSKEY="P" ' . $w_Disabled . ' class="STI" type="text" name="w_ds_texto_abertura" rows=5 cols=65 title="OBRIGATÓRIO. Informe o texto a ser exibido na página de abertura do site.">' . $w_ds_texto_abertura . '</TEXTAREA></td>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Página "Quem somos"</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Informe o texto ser colocado na página "Quem somos" do site.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML('      <tr><td><font size="1"><b>Texto da seção "<u>Q</u>uem somos":</b><br><TEXTAREA ACCESSKEY="Q" ' . $w_Disabled . ' class="STI" type="text" name="w_ds_institucional" rows=5 cols=65 title="OBRIGATÓRIO. Informe o texto a ser exibido na seção.">' . $w_ds_institucional . "</TEXTAREA></td>");
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    If ($w_tipo == 2) { // ' Se for regional
      ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Página "Composição administrativa"</td></td></tr>');
      ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
      ShowHTML('      <tr><td><font size=1>Informe o arquivo Word ou PDF a ser exibido na página "Composição administrativa" do site.');
      ShowHTML('          <br><font color="red"><b>IMPORTANTE: <a href="sedf/orientacoes_word.pdf" class="hl" target="_blank">Para documentos Word, clique aqui para ler as orientações sobre a formatação e a proteção do texto</a></b></font>.');
      ShowHTML('      </font></td></tr>');
      ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
      ShowHTML('      <tr><td><font size="1"><b>Composição adminis<u>t</u>rativa (arquivo Word ou PDF):</b><br><INPUT ACCESSKEY="T" ' . $w_Disabled . ' class="STI" type="file" name="w_pedagogica" size="60" maxlength="100" value="" title="OPCIONAL. Clique no botão ao lado para localizar o arquivo que contém a composição administrativa da regional. Ele será transferido automaticamente para o servidor.">');
    } Else {
      //'SQL = "select b.ds_especialidade from escEspecialidade_Cliente a inner join escEspecialidade b on (a.sq_codigo_espec = b.sq_especialidade and a." . CL . ")"
      If ($w_tipo != 13) {
        //'   If uCase(RS("ds_especialidade")) <> uCase("Biblioteca") Then
        ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Página "Projeto"</td></td></tr>');
        ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
        ShowHTML('      <tr><td><font size=1>Informe o arquivo Word ou PDF a ser exibido na página "Projeto" do site.');
        ShowHTML('          <br><font color="red"><b>IMPORTANTE: <a href="sedf/orientacoes_word.pdf" class="hl" target="_blank">Para documentos Word, clique aqui para ler as orientações sobre a formatação e a proteção do texto</a></b></font>.');
        ShowHTML('      </font></td></tr>');
        ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
        ShowHTML('      <tr><td><font size="1"><b>Proje<u>t</u>o (arquivo Word):</b><br><INPUT ACCESSKEY="T" ' . $w_Disabled . ' class="STI" type="file" name="w_pedagogica" size="60" maxlength="100" value="" title="OPCIONAL. Clique no botão ao lado para localizar o arquivo que contém o projeto da escola. Ele será transferido automaticamente para o servidor.">');
      }
    }
    If ($w_pedagogica > '') {
      ShowHTML('              <b><a class="SS" href="' . $w_ds_diretorio . '/' . $w_pedagogica . '" target="_blank" title="Clique para abrir o arquivo atual.">Exibir</a></b>');
    }
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Diversos</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Informe os dados abaixo de exibição geral no site.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size="1"><b>Texto da men<u>s</u>agem em destaque:</b><br><INPUT ACCESSKEY="S" ' . $w_Disabled . ' class="STI" type="text" name="w_ds_mensagem" size="80" maxlength="80" value="' . $w_ds_mensagem . '" title="OBRIGATÓRIO. Informe um texto que será exibido na parte superior do site, numa barra rolante."></td>');
    ShowHTML('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');

    //' Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
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
  // Monta o menu principal da aplicação
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
    } else
      if ($O == 'L') {
        //Recupera todos os registros para a listagem
        if ($_SESSION["USERNAME"] == "IMPRENSA" or $_SESSION["USERNAME"] == "SBPI") {
          $SQL = "select sq_noticia as chave, ds_titulo, ds_noticia, dt_noticia, ln_externo, ativo, in_exibe from sbpi.Noticia_Cliente where sq_cliente = 0 order by dt_noticia desc";
        } else {
          $SQL = 'select sq_noticia as chave, ds_titulo, ds_noticia, dt_noticia, ln_externo, ativo, in_exibe from sbpi.Noticia_Cliente where sq_cliente= ' . $CL . ' order by dt_noticia desc';
        }
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      } else
        if (strpos("AEV", $O) !== false && $w_troca == '') {
          //Recupera os dados do endereço informado
          $SQL = "select sq_noticia as chave, ds_titulo, ds_noticia, dt_noticia, ativo, in_exibe, ln_externo from sbpi.Noticia_Cliente where sq_noticia = " . $w_chave;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }

          $w_dt_noticia = FormataDataEdicao(f($RS, "dt_noticia"));
          $w_ds_titulo  = str_replace('"','&quot;',f($RS, "ds_titulo"));
          $w_ds_noticia = f($RS, "ds_noticia");
          $w_in_ativo   = f($RS, "ativo");
          $w_in_exibe   = f($RS, "in_exibe");
          $w_ln_externo = str_replace('"','&quot;',f($RS, "ln_externo"));
        }
    Cabecalho();
    ShowHTML('<HEAD>');
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
    } else
      if ($O == "I" or $O == "A") {
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
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td width="1%"><font size="1"><b>Data</font></td>');
      ShowHTML('          <td align="center"><font size="1"><b>Título</font></td>');
      ShowHTML('          <td align="center"><font size="1"><b>Descrição</font></td>');      
      ShowHTML('          <td width="1%"><font size="1"><b>Ativo</font></td>');
      ShowHTML('          <td width="1%"><font size="1"><b>Operações</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td align="center"><font size="1">' . FormataDataEdicao(FormatDateTime(f($row, "dt_noticia"), 2)) . '</td>');
          if(f($row,"ln_externo") != ""){
            ShowHTML('        <td><font size="1"><a href="'.f($row,"ln_externo").'" target="_blank">' . f($row, "ds_titulo") . '</a></td>');
          }else{
            ShowHTML('        <td><font size="1">' . f($row, "ds_titulo") . '</td>');
          }
          ShowHTML('        <td>' . f($row, "ds_noticia") . '</td>');
          ShowHTML('        <td align="center"><font size="1">' . f($row, "ativo") . '</td>');
          ShowHTML('        <td align="top" nowrap><font size="1">');
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
    } else
      if (strpos("IAEV", $O) !== false) {
        if (strpos("EV", $O) !== false) {
          $w_Disabled = ' DISABLED ';
        }
        AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
        ShowHTML('<INPUT type="hidden" name="SG" value="NOTICIAS">');
        ShowHTML('<INPUT type="hidden" name="CL" value="' . $CL . '">');
        ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
        ShowHTML('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
        ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');

        ShowHTML('<tr bgcolor="#EFEFEF"><td align="center">');
        ShowHTML('    <table width="95%" border="0">');
        ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
        ShowHTML('        <tr valign="top">');
        ShowHTML('          <td valign="top"><font size="1"><b><u>D</u>ata:</b><br><input ' . $w_Disabled . ' accesskey="D" type="text" name="w_dt_noticia" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_noticia, Time()), 2)) . '" onKeyDown="FormataData(this,event);" ></td>');
        ShowHTML('        <tr valign="top">');
        ShowHTML('          <td valign="top"><font size="1"><b><u>T</u>ítulo:</b><br><input ' . $w_Disabled . ' accesskey="T" type="text" name="w_ds_titulo" class="STI" SIZE="90" MAXLENGTH="100" VALUE="' . $w_ds_titulo . '"></td>');
        ShowHTML('      <tr valign="top">');
        ShowHTML('          <td valign="top"><font size="1"><b><u>L</u>ink externo:</b><br><input " . w_Disabled . " accesskey="L" type="text" name="w_ln_externo" class="STI" SIZE="90" MAXLENGTH="255" VALUE="' . $w_ln_externo . '"></td>');
        ShowHTML('      </tr>');
        ShowHTML('        </table>');
        ShowHTML('      <tr><td valign="top"><font size="1"><b>D<u>e</u>scrição:</b><br><textarea " ' . $w_Disabled . ' " accesskey="E" name="w_ds_noticia" class="STI" ROWS=5 cols=65>' . $w_ds_noticia . '</TEXTAREA></td>');
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
  // ---------------------------------------------

  function menu() {
    extract($GLOBALS);

    // Inclusão do arquivo da classe
    include_once ("classes/menu/xPandMenu.php");

    // Instanciando a classe menu
    $root = new XMenu();

    $w_imagem = 'img/SheetLittle.gif';
    $i = 0;

    $i++;
    eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Manual SIGE-WEB\',\'manuais/operacao/\',$w_Imagem,$w_Imagem,\'_blank\', null));');
    if ($_SESSION['TIPO'] == 3) {
      $i++;
      eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Administrativo\',$w_pagina.\'administrativo\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    $i++;
    eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Fotos\',$w_pagina.\'fotos\',$w_Imagem,$w_Imagem,\'content\', null));');
    $i++;
    eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Dados básicos\',$w_pagina.\'basico\',$w_Imagem,$w_Imagem,\'content\', null));');
    $i++;
    eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Dados adicionais\',$w_pagina.\'adicionais\',$w_Imagem,$w_Imagem,\'content\', null));');
    $i++;
    eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Dados do site\',$w_pagina.\'getsite\',$w_Imagem,$w_Imagem,\'content\', null));');
    if ($_SESSION['TIPO'] == 3) {
      $i++;
      eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Áreas de atuação\',$w_pagina.\'atuacao\',$w_Imagem,$w_Imagem,\'content\', null));');
    }
    $i++;
    eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Arquivos (<i>download</i>)\',$w_pagina.\'arquivos\',$w_Imagem,$w_Imagem,\'content\', null));');

    if ($_SESSION['TIPO'] == 3) {
      $i++;
      eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Calendário\',$w_pagina.\'calendario\',$w_Imagem,$w_Imagem,\'content\', null));');
      $SQL = "select count(*) as qtd from sbpi.cliente_site a inner join sbpi.modelo b on (a.sq_modelo = b.sq_modelo and upper(b.ds_diretorio) <> 'MOD13') where a.sq_cliente = " . $CL;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      if (f($RS, 'qtd') > 0) {
        $i++;
        eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Mensagens\',$w_pagina.\'mensagens\',$w_Imagem,$w_Imagem,\'content\', null));');
      }
    }
    $i++;
    eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Notícias\',$w_pagina.\'noticias\',$w_Imagem,$w_Imagem,\'content\', null));');

    $i++;
    eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Log\',$w_pagina.\'log_manut\',$w_Imagem,$w_Imagem,\'content\', null));');

    $i++;
    eval ('$node' . $i . ' = &$root->addItem(new XNode(\'Encerrar\',$w_pagina.\'Sair\',$w_Imagem,$w_Imagem,\'_top\', \'onClick="return(confirm(\\\'Confirma saída do sistema?\\\'));"\' ));');

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
    ShowHTML('onLoad="javascript:top.content.location=\'' . $w_pagina . 'getsite\';"> ');
    ShowHTML('  <TR><TD><font size=2><b>Atualização</TD></TR>');
    ShowHTML('  <CENTER><table border=0 cellpadding=0 height="80" width="100%">');
    ShowHTML('      <tr><td colspan=2 width="100%"><table border=0 width="100%" cellpadding=0 cellspacing=0><tr valign="top">');
    ShowHTML('          <td align="center">Usuário:<b>' . $_SESSION['USERNAME'] . '</b></TD>');
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
    $_SESSION = array ();
    // Finalmente, destruição da sessão.
    session_destroy();

    ScriptOpen('JavaScript');
    ShowHTML('  top.location.href=\'' . $w_pagina . '\';');

    ScriptClose();
  }

  // =========================================================================
  // Cadastro de calendario
  // -------------------------------------------------------------------------
  function calend_rede() {
    extract($GLOBALS);
    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];

    if ($w_troca > "") { //Se for recarga da página
      $w_dt_ocorrencia = $_REQUEST["w_dt_ocorrencia"];
      $w_ds_ocorrencia = $_REQUEST["w_ds_ocorrencia"];
      $w_tipo = $_REQUEST["w_tipo"];
    } else
      if ($O == "L") {
        //Recupera todos os registros para a listagem
        $SQL = "select a.sq_ocorrencia as chave, a.ds_ocorrencia, a.dt_ocorrencia, a.sq_tipo_data, b.nome from sbpi.Calendario_Cliente a left join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) where sq_cliente = " . $CL . " order by sbpi.year(dt_ocorrencia) desc, dt_ocorrencia";
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      } else
        if (strpos("AEV", $O) !== false and $w_troca == '') {
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
    if (strpos("IAEP", $O) !== false) {
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      if (strpos("IA", $O) !== false) {
        Validate("w_dt_ocorrencia", "Data", "DATA", "1", "10", "10", "1", "1");
        Validate("w_ds_ocorrencia", "Descrição", "", "1", "2", "60", "1", "1");
        Validate("w_tipo", "Tipo da ocorrência", "SELECT", "1", "1", "18", "", "1");
      }
      ShowHTML('  theForm.Botao[0].disabled=true;');
      ShowHTML('  theForm.Botao[1].disabled=true;');
      ValidateClose();
      ScriptClose();
    }
    ShowHTML('</HEAD>');
    if ($w_troca > "") {
      BodyOpen('onLoad=\'document.Form.' . $w_troca . '.focus()\';');
    } else
      if ($O == "I" or $O == "A") {
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
      ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $w_pagina . 'calendario' . '&R=' . $w_pagina . 'calendario' . '&O=I"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('    <td align="right"><font size="1"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><font size="1"><b>Data</font></td>');
      ShowHTML('          <td><font size="1"><b>Tipo</font></td>');
      ShowHTML('          <td><font size="1"><b>Ocorrência</font></td>');
      ShowHTML('          <td><font size="1"><b>Operações</font></td>');
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
          ShowHTML('        <td align="center"><font size="1">' . substr(FormataDataEdicao(FormatDateTime(f($row, "dt_ocorrencia"), 2)), 0, 5) . '</td>');
          ShowHTML('        <td><font size="1">' . nvl(f($row, "nome"), "---") . '</td>');
          ShowHTML('        <td><font size="1">' . f($row, "ds_ocorrencia") . '</td>');
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
    elseif (!(strpos('IAEV', $O) === false)) {
      if (!(strpos('EV', $O) === false))
        $w_Disabled = ' DISABLED ';
      AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
      ShowHTML('<INPUT type="hidden" name="w_troca" value="">');
      ShowHTML('<INPUT type="hidden" name="SG" value="CALENDARIO">');
      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');

      ShowHTML('<tr bgcolor="' . $conTrBgColor . '"><td align="center">');
      ShowHTML('    <table width="97%" border="0">');
      ShowHTML('      <tr><td><table border=0 width="100%"><tr valign="top">');
      ShowHTML('           <td colspan=2><b><u>D</u>ata:</b><br><input ' . $w_Disabled . ' accesskey="D" type="text" name="w_dt_ocorrencia" class="sti" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_ocorrencia, Time()), 2)) . '"  onKeyDown="FormataData(this,event);"></td>');
      ShowHTML('      <tr><td valign="top" colspan="2"><table border="0" width="100%" cellspacing=0>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('           <td colspan=2><b><u>D</u>escrição:</b><br><input ' . $w_Disabled . ' accesskey="E" type="text" name="w_ds_ocorrencia" class="sti" SIZE="60" MAXLENGTH="60" VALUE="' . $w_ds_ocorrencia . '"></td>');
      $SQL = 'SELECT * FROM sbpi.Tipo_Data a WHERE a.abrangencia <> \'U\' ORDER BY a.nome' . $crlf;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      ShowHTML('          <td><font size="1"><b>Tipo da ocorrência:</b><br><SELECT ' . $w_Disabled . ' CLASS="STI" NAME="w_tipo">');
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
  // -----------------------

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
    } else
      if ($O == "L") {
        //Recupera todos os registros para a listagem

        $SQL = ' select * from sbpi.Cliente_Arquivo where  sq_cliente = ' . $CL . ' order by nr_ordem, ltrim(upper(ds_titulo))';
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

      } else
        if (strpos('AEV', $O) !== false and $w_troca == "") {
          //Recupera os dados do endereço informado
          $SQL = "select b.ds_diretorio, a.* from sbpi.Cliente_Arquivo a inner join sbpi.Cliente_Site b on (a.sq_cliente = b.sq_cliente) where a.sq_arquivo = " . $w_chave;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_dt_arquivo = FormataDataEdicao(f($RS, "dt_arquivo"));
          $w_ds_titulo = f($RS, "ds_titulo");
          $w_in_ativo = f($RS, "in_ativo");
          $w_ds_arquivo = f($RS, "ds_arquivo");
          $w_ln_arquivo = f($RS, "ln_arquivo");
          $w_in_destinatario = f($RS, "in_destinatario");
          $w_nr_ordem = f($RS, "nr_ordem");
          $w_ds_diretorio = f($RS, "ds_diretorio");
          $w_pasta        = f($RS, "pasta");
        }

    Cabecalho();
    ShowHTML('<HEAD>');
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
        Validate("w_nr_ordem", "Nr. de ordem", "", "1", "1", "4", "", "0123546789");
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
    } else
      if ($O == "I" or $O == "A") {
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
      ShowHTML('<tr><td align="center" colspan="2"><font size="1" color="red"><b>IMPORTANTE: <a href="sedf/orientacoes_word.pdf" class="hl" target="_blank">Para documentos Word, clique aqui para ler as orientações sobre a formatação e a proteção do texto</a></b></font>.');
      ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $dir . $w_pagina . $par . $w_ew . "&R=" . $w_pagina . $par . "&O=I&CL=" . $CL . '"><u>I</u>ncluir</a>&nbsp;');
      ShowHTML('    <td align="right"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="' . $conTrBgColor . '" align="center">');
      ShowHTML('          <td><font size="1"><b>Ordem</font></td>');
      ShowHTML('          <td><font size="1"><b>Arquivo</font></td>');
      ShowHTML('          <td><font size="1"><b>Descrição</font></td>');
      ShowHTML('          <td><font size="1"><b>Ativo</font></td>');
      ShowHTML('          <td><font size="1"><b>Operações</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td align="center"><font size="1">' . f($row, "nr_ordem") . '</td>');
          ShowHTML('        <td><font size="1">' . f($row, "ds_titulo") . '</td>');
          ShowHTML('        <td><font size="1">' . f($row, "ds_arquivo") . '</td>');
          ShowHTML('        <td align="center"><font size="1">' . f($row, "ativo") . '</td>');
          ShowHTML('        <td align="top" nowrap><font size="1">');
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
    } else
      if (strpos("IAEV", $O) !== false) {
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
        ShowHTML('          <td valign="top"><font size="1"><b><u>T</u>ítulo:</b><br><input "' . $w_Disabled . '" accesskey="T" type="text" name="w_ds_titulo" class="STI" SIZE="50" MAXLENGTH="50" VALUE="' . $w_ds_titulo . '" ></td>');
        if ($O == "I") {
          ShowHTML('          <td valign="top"><font size="1"><b>Cadastramento:</b><br><input disabled type="text" name="w_dt_arquivo" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(time(), 2)) . '"></td>');
        } else {
          ShowHTML('          <td valign="top"><font size="1"><b>Última alteração:</b><br><input disabled type="text" name="w_dt_arquivo" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime($w_dt_arquivo, 2)) . '"></td>');
        }
        ShowHTML('        </table>');
        ShowHTML('      <tr><td valign="top"><font size="1"><b><u>D</u>escrição:</b><br><textarea " ' . $w_Disabled . ' " accesskey="D" name="w_ds_arquivo" class="STI" ROWS=5 cols=65>' . $w_ds_arquivo . '</TEXTAREA></td>');
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
        ShowHTML('      <tr><td valign="top"><font size="1"><b><u>L</u>ink:</b><br><input "  ' . $w_Disabled . ' " accesskey="L" type="file" name="w_ln_arquivo" class="STI" SIZE="80" MAXLENGTH="100" VALUE="" >');
        if ($w_ln_arquivo > '') {
          ShowHTML('              <b><a class="SS" href="' . $w_ds_diretorio . '/' . $w_ln_arquivo . '" target="_blank" title="Clique para exibir o arquivo atual.">Exibir</a></b>');
        }
        ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
        ShowHTML('        <tr valign="top">');
        ShowHTML('          <td><font size="1"><b><u>N</u>r. de ordem:</b><br><input "' . $w_Disabled . '" accesskey="N" type="text" name="w_nr_ordem" class="STI" SIZE="4" MAXLENGTH="4" VALUE="' . $w_nr_ordem . '"></td>');
        ShowHTML('          <td><font size="1"><b><u>D</u>estinatários:</b><br><select "' . $w_Disabled . '" accesskey="D" name="w_in_destinatario" class="STS" SIZE="1">');
        if ($w_in_destinatario == 'A') {
          ShowHTML('            <OPTION VALUE="A" SELECTED>Apenas alunos <OPTION VALUE="P">Apenas professores <OPTION VALUE="T">Professores e alunos <OPTION VALUE="E">Escola');
        } else
          if ($w_in_destinatario == "E") {
            ShowHTML('            <OPTION VALUE="A">Apenas alunos <OPTION VALUE="P">Apenas professores <OPTION VALUE="T">Professores e alunos <OPTION VALUE="E" SELECTED>Escola');
          } else
            if ($w_in_destinatario == "P") {
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
          //ShowHTML ('   <input class="STB" type="submit" name="Botao" value="Excluir">"
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
        //ShowHTML (' history.back(1);');
        ScriptClose();
      }
    ShowHTML('</table>');
    ShowHTML('</center>');
    Rodape();

  }
  // =========================================================================
  // Fim do cadastro de arquivos
  // -----------------------

  // =========================================================================
  // Rotina Gravação
  // -------------------------------------------------------------------------
  function Grava() {
    extract($GLOBALS);
    $SG = strtoupper($_REQUEST['SG']);
    switch ($SG) {

      Case 'ADICIONAIS' :

        $SQL = "select count(*) existe from sbpi.Cliente_Dados where sq_cliente = " . $CL;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        foreach ($RS as $row) {
          $RS = $row;
          break;
        }
        if (intval(f($RS, "existe")) == 0) {
          $w_chave = str_replace("sq_CLiente=", "", $CL);
          //Criação das escolas na tabela que contém dados complementares'
          $SQL = "insert into sbpi.Cliente_Dados ( " . $crlf .
                 "     sq_CLiente_dados, sq_CLiente,       nr_cnpj,        tp_registro, " . $crlf .
                 "     ds_ato,           nr_ato,           dt_ato,         ds_orgao, " . $crlf .
                 "     no_bairro,        ds_email_contato, nr_fax_contato, no_diretor, " . $crlf .
                 "     no_secretario,    ds_logradouro,    nr_cep,         no_contato," . $crlf .
                 "     nr_fone_contato" . $crlf .
                 " ) values ( " . $crlf .
                 "     " . $w_chave . "," . $crlf .
                 "     " . $w_chave . "," . $crlf;
          if ($_REQUEST["w_nr_cnpj"] > "") {
            $SQL .= "   '" . $_REQUEST["w_nr_cnpj"] . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_tp_registro"] > "") {
            $SQL .= "   '" . $_REQUEST["w_tp_registro"] . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_ds_ato"] > "") {
            $SQL .= "   '" . trim($_REQUEST["w_ds_ato"]) . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_nr_ato"] > "") {
            $SQL .= "   '" . trim($_REQUEST["w_nr_ato"]) . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_dt_ato"] > "") {
            $SQL .= "   '" . cDate(trim($_REQUEST["w_dt_ato"])) . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_ds_orgao"] > "") {
            $SQL .= "   '" . trim($_REQUEST["w_ds_orgao"]) . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_no_bairro"] > "") {
            $SQL .= "   '" . trim($_REQUEST["w_no_bairro"]) . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_ds_email_contato"] > "") {
            $SQL .= "   '" . trim($_REQUEST["w_ds_email_contato"]) . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_nr_fax_contato"] > "") {
            $SQL .= "   '" . trim($_REQUEST["w_nr_fax_contato"]) . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_no_diretor"] > "") {
            $SQL .= "   '" . trim($_REQUEST["w_no_diretor"]) . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          if ($_REQUEST["w_no_secretario"] > "") {
            $SQL .= "   '" . trim($_REQUEST["w_no_secretario"]) . "', " . $crlf;
          } else {
            $SQL .= "   null, " . $crlf;
          }
          $SQL .= "   '" . trim($_REQUEST["w_ds_logradouro"]) . "', " . $crlf .
          "   '" . $_REQUEST["w_nr_cep"] . "', " . $crlf .
          "   '" . trim($_REQUEST["w_no_contato"]) . "', " . $crlf .
          "   '" . trim($_REQUEST["w_nr_fone_contato"]) . "' " . $crlf .
          ")" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        } else {
          $SQL = "update sbpi.Cliente_Dados set " . $crlf .
                 "   ds_logradouro  = '" . trim($_REQUEST["w_ds_logradouro"]) . "', " . $crlf .
                 "   nr_cep           = '" . $_REQUEST["w_nr_cep"] . "', " . $crlf .
                 "   no_contato     = '" . trim($_REQUEST["w_no_contato"]) . "', " . $crlf .
                 "   nr_fone_contato= '" . trim($_REQUEST["w_nr_fone_contato"]) . "', " . $crlf;
          if ($_REQUEST["w_no_bairro"] > "") {
            $SQL .= "   no_bairro          = '" . trim($_REQUEST["w_no_bairro"]) . "', " . $crlf;
          } else {
            $SQL .= "   no_bairro        = null, " . $crlf;
          }
          if ($_REQUEST["w_no_diretor"] > "") {
            $SQL .= "   no_diretor       = '" . trim($_REQUEST["w_no_diretor"]) . "', " . $crlf;
          } else {
            $SQL .= "   no_diretor       = null, " . $crlf;
          }
          if ($_REQUEST["w_no_secretario"] > "") {
            $SQL .= "   no_secretario    = '" . trim($_REQUEST["w_no_secretario"]) . "', " . $crlf;
          } else {
            $SQL .= "   no_secretario    = null, " . $crlf;
          }
          if ($_REQUEST["w_nr_cnpj"] > "") {
            $SQL .= "   nr_cnpj            = '" . $_REQUEST["w_nr_cnpj"] . "', " . $crlf;
          } else {
            $SQL .= "   nr_cnpj          = null, " . $crlf;
          }
          if ($_REQUEST["w_tp_registro"] > "") {
            $SQL .= "   tp_registro      = '" . $_REQUEST["w_tp_registro"] . "', " . $crlf;
          } else {
            $SQL .= "   tp_registro      = null, " . $crlf;
          }
          if ($_REQUEST["w_ds_ato"] > "") {
            $SQL .= "   ds_ato             = '" . trim($_REQUEST["w_ds_ato"]) . "', " . $crlf;
          } else {
            $SQL .= "   ds_ato           = null, " . $crlf;
          }
          if ($_REQUEST["w_nr_ato"] > "") {
            $SQL .= "   nr_ato             = '" . trim($_REQUEST["w_nr_ato"]) . "', " . $crlf;
          } else {
            $SQL .= "   nr_ato           = null, " . $crlf;
          }
          if ($_REQUEST["w_dt_ato"] > "") {
            $SQL .= "   dt_ato             = '" . cDate(trim($_REQUEST["w_dt_ato"])) . "', " . $crlf;
          } else {
            $SQL .= "   dt_ato           = null, " . $crlf;
          }
          if ($_REQUEST["w_ds_orgao"] > "") {
            $SQL .= "   ds_orgao           = '" . trim($_REQUEST["w_ds_orgao"]) . "', " . $crlf;
          } else {
            $SQL .= "   ds_orgao         = null, " . $crlf;
          }
          if ($_REQUEST["w_ds_email_contato"] > "") {
            $SQL .= "   ds_email_contato = '" . trim($_REQUEST["w_ds_email_contato"]) . "', " . $crlf;
          } else {
            $SQL .= "   ds_email_contato = null, " . $crlf;
          }
          if ($_REQUEST["w_nr_fax_contato"] > "") {
            $SQL .= "   nr_fax_contato   = '" . trim($_REQUEST["w_nr_fax_contato"]) . "' " . $crlf;
          } else {
            $SQL .= "   nr_fax_contato   = null " . $crlf;
          }
          $SQL .= "where sq_cliente=" . $CL . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }
        $w_sql = $SQL;
        $SQL = "update sbpi.Cliente set dt_alteracao = sysdate where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        $w_funcionalidade = 22;

        //Grava o acesso na tabela de log
        If ($_SESSION['USERNAME'] != 'SBPI') {
          $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                 "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
                 "           " . $CL . ", " . $crlf .
                 "           sysdate, " . $crlf .
                 "           '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
                 "           2, " . $crlf .
                 "           'Atualização de dados adicionais.', " . $crlf .
                 "           '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                 "           '" . $w_funcionalidade . "' " . $crlf .
                 "   from dual) " . $crlf;

          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }

        ScriptOpen('JavaScript');
        ShowHTML('location.href="' . $w_pagina . $_REQUEST["SG"] . '&O=L";');
        ScriptClose();

        break;

      case 'MENSAGENS' :
        $w_funcionalidade = 27;
        if ($O == "I") {
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

          //Grava o acesso na tabela de log
          $w_sql = $SQL;

          If ($_SESSION['USERNAME'] != 'SBPI') {
            $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                   "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
                   "         " . str_replace("sq_CLiente=", "", $CL) . ", " . $crlf .
                   "         sysdate, " . $crlf .
                   "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
                   "         1, " . $crlf .
                   "         'Inclusão de mensagem para aluno da escola.', " . $crlf .
                   "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                   "         " . $w_funcionalidade . " " . $crlf .
                   "   from dual) " . $crlf;
                   $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        } else
          if ($O == "A") {
            $SQL = " update sbpi.Mensagem_Aluno set " . $crlf .
            "     dt_mensagem  = to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_mensagem"], 2)) . "','dd/mm/yyyy'), " . $crlf .
            "     ds_mensagem  = '" . $_REQUEST["w_ds_mensagem"] . "' " . $crlf .
            "where sq_mensagem = " . $_REQUEST["w_chave"] . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            If ($_SESSION['USERNAME'] != 'SBPI') {
              $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                     "( select sbpi.sq_cliente_log.nextval, " . $crlf .
                     "         " . $CL . ", " . $crlf .
                     "         sysdate, " . $crlf .
                     "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf .
                     "         2, " . $crlf .
                     "         'Alteração de mensagem para aluno da escola.', " . $crlf .
                     "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                     "         " . $w_funcionalidade . " " . $crlf .
                     "   from dual) " . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }
          } else
            if ($O = "E") {
              $SQL = " delete sbpi.Mensagem_Aluno where sq_mensagem = " . $_REQUEST["w_chave"] . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

              //Grava o acesso na tabela de log
              $w_sql = $SQL;
              If ($_SESSION['USERNAME'] != 'SBPI') {
                $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                       "( select sbpi.sq_cliente_log.nextval, " . $crlf .
                       "         " . $CL . ", " . $crlf .
                       "         sysdate, " . $crlf .
                       "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf .
                       "         2, " . $crlf .
                       "         'Exclusão de mensagem para aluno da escola.', " . $crlf .
                       "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                       "         " . $w_funcionalidade . " " . $crlf .
                       "   from dual) " . $crlf;
                $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
              }
            }

        $SQL = "update sbpi.Cliente set dt_alteracao = sysdate where sq_cliente=" . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        ScriptOpen('JavaScript');
        ShowHTML(' location.href=\'' . $w_dir . $w_pagina . $SG . '&O=L&CL=' . $CL . '\';');
        ScriptClose();
        break;

      Case 'NOTICIAS' :
        $w_funcionalidade = 28;

        //	exibevariaveis();

        if ($O == "I") {
          //Recupera o valor da próxima chave primária
          $SQL = "select sbpi.sq_noticia.nextval chave from sbpi.Noticia_Cliente" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");

          //Insere o arquivo
          $SQL = " insert into sbpi.Noticia_Cliente " . $crlf .
                 "    (sq_noticia, sq_CLiente, dt_noticia, ds_titulo, ds_noticia, ln_externo, ativo) " . $crlf .
                 " values ( " . $w_chave . ", " . $crlf .
                 "     " . $_REQUEST["w_sq_cliente"] . ", " . $crlf .
                 "     to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_noticia"], 2)) . "','dd/mm/yyyy'), " . $crlf .
                 "     '" . $_REQUEST["w_ds_titulo"] . "', " . $crlf .
                 "     '" . $_REQUEST["w_ds_noticia"] . "', " . $crlf .
                 "     '" . $_REQUEST["w_ln_externo"] . "', " . $crlf .                 
                 "     '" . $_REQUEST["w_in_ativo"] . "' " . $crlf .
                 " )" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          //Grava o acesso na tabela de log
          $w_sql = $SQL;
          If ($_SESSION['USERNAME'] != 'SBPI') {
            $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                   "( select sbpi.sq_cliente_log.nextval, " . $crlf .
                   "         " . $_REQUEST["w_sq_cliente"] . ", " . $crlf .
                   "         sysdate, " . $crlf .
                   "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf .
                   "         1, " . $crlf .
                   "         'InCLusão de notícia da escola.', " . $crlf .
                   "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                   "         " . $w_funcionalidade . " " . $crlf .
                   "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }

        } else
          if ($O == "A") {
            $SQL = " update sbpi.Noticia_Cliente set " . $crlf .
                   "     dt_noticia  = to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_noticia"], 2)) . "','dd/mm/yyyy'), " . $crlf .
                   "     ds_titulo   = '" . $_REQUEST["w_ds_titulo"] . "', " . $crlf .
                   "     ds_noticia  = '" . $_REQUEST["w_ds_noticia"] . "', " . $crlf .
                   "     ln_externo  = '" . $_REQUEST["w_ln_externo"] . "', " . $crlf .
                   "     ativo    = '" . $_REQUEST["w_in_ativo"] . "' " . $crlf .
                   "where sq_noticia = " . $_REQUEST["w_chave"] . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            If ($_SESSION['USERNAME'] != 'SBPI') {
              $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                     "( select sbpi.sq_cliente_log.nextval, " . $crlf .
                     "         " . $_REQUEST["w_sq_cliente"] . ", " . $crlf .
                     "         sysdate, " . $crlf .
                     "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf .
                     "         2, " . $crlf .
                     "         'Alteração de notícia da escola.', " . $crlf .
                     "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                     "         " . $w_funcionalidade . " " . $crlf .
                     "   from dual) " . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }
          } else
            if ($O == "E") {
              $SQL = " delete sbpi.Noticia_Cliente where sq_noticia = " . $_REQUEST["w_chave"] . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

              //Grava o acesso na tabela de log
              $w_sql = $SQL;
              If ($_SESSION['USERNAME'] != 'SBPI') {
                $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                       "( select sbpi.sq_cliente_log.nextval, " . $crlf .
                       "         " . $_REQUEST["w_sq_cliente"] . ", " . $crlf .
                       "         sysdate, " . $crlf .
                       "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf .
                       "         3, " . $crlf .
                       "         'ExCLusão de notícia da escola.', " . $crlf .
                       "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                       "         " . $w_funcionalidade . " " . $crlf .
                       "   from dual) " . $crlf;
                $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
              }
            }

        $SQL = "update sbpi.Cliente set dt_alteracao = sysdate where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        ScriptOpen('JavaScript');
        ShowHTML('  location.href=\'' . $w_dir . $w_pagina . $SG . '&O=L&CL=' . $CL . '\';');
        ScriptClose();
        break;

      Case 'CALENDARIO' :
        $w_funcionalidade = 26;

        if ($O == "I") {
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
                 "     '" . $_REQUEST["w_ds_ocorrencia"] . "', " . $_REQUEST["w_tipo"] . $crlf .
                 " )" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          //Grava o acesso na tabela de log
          If ($_SESSION['USERNAME'] != 'SBPI') {
            $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                   "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
                   "           " . $CL . ", " . $crlf .
                   "           sysdate, " . $crlf .
                   "           '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
                   "           1, " . $crlf .
                   "           'Inclusão de data no calendário da escola.', " . $crlf .
                   "           '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                   "           " . $w_funcionalidade . " " . $crlf .
                   "   from dual) " . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          }
        } else
          if ($O == "A") {

            $SQL = " update sbpi.Calendario_Cliente set " . $crlf .
                   "     dt_ocorrencia  = to_date('" . FormataDataEdicao(FormatDateTime($_REQUEST["w_dt_ocorrencia"], 2)) . "','dd/mm/yyyy'), " . $crlf .
                   "     ds_ocorrencia  = '" . $_REQUEST["w_ds_ocorrencia"] . "', " . $crlf .
                   "     sq_tipo_data   = " . $_REQUEST["w_tipo"] . $crlf .
                   "where sq_ocorrencia = " . $_REQUEST["w_chave"] . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            If ($_SESSION['USERNAME'] != 'SBPI') {
              $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                     "( select sbpi.sq_cliente_log.nextval, " . $crlf .
                     "         " . $CL . ", " . $crlf .
                     "         sysdate, " . $crlf .
                     "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf .
                     "         2, " . $crlf .
                     "         'Alteração de data no calendário da escola.', " . $crlf .
                     "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                     "         " . $w_funcionalidade . " " . $crlf .
                     "   from dual) " . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }

          } else
            if ($O == "E") {
              $SQL = " delete sbpi.Calendario_Cliente where sq_ocorrencia = " . $_REQUEST["w_chave"] . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

              //Grava o acesso na tabela de log
              $w_sql = $SQL;
              If ($_SESSION['USERNAME'] != 'SBPI') {
                $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                       "( select sbpi.sq_cliente_log.nextval, " . $crlf .
                       "         " . $CL . ", " . $crlf .
                       "         sysdate, " . $crlf .
                       "         '" . $_SERVER["REMOTE_ADDR"] . "', " . $crlf .
                       "         3, " . $crlf .
                       "         'ExCLusão de data no calendário da escola.', " . $crlf .
                       "         '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                       "         " . $w_funcionalidade . " " . $crlf .
                       "   from dual) " . $crlf;
                $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
              }
            }

        $SQL = "update sbpi.Cliente set dt_alteracao = sysdate where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        ScriptOpen('JavaScript');
        ShowHTML('  location.href=\'' . $w_dir . $w_pagina . $SG . '&O=L&CL=' . $CL . '\';');
        ScriptClose();
        Break;

        //Áreas de Atuação
      Case 'ATUACAO' :
        $w_funcionalidade = 24;
        //Apaga especialidades existentes para a unidade
        $SQL = "delete sbpi.Especialidade_Cliente where sq_cliente = " . $CL . $crlf;
        $w_sql = $SQL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        //Em seguida, cria apenas especialidades indicadas pelo CLiente

        if (isset ($_REQUEST["w_chave"])) {
          foreach ($_REQUEST["w_chave"] as $row) {
            $SQL = "insert into sbpi.Especialidade_Cliente (sq_cliente, sq_especialidade) " . $crlf .
            "values (" . $crlf .
            "   " . $CL . ", " . $crlf .
            "   " . $row . " " . $crlf .
            ")";

            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            $w_sql = $w_sql . $SQL . $crlf;
          }
        }

        $SQL = "update sbpi.Cliente set dt_alteracao = sysdate where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        //Grava o acesso na tabela de log
        If ($_SESSION['USERNAME'] != 'SBPI') {
          $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                 "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
                 "         " . $CL . ", " . $crlf .
                 "         sysdate, " . $crlf .
                 "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
                 "         2, " . $crlf .
                 "         'Atualização das modalidades de ensino.', " . $crlf .
                 "           '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                 "           " . $w_funcionalidade . " " . $crlf .
                 "   from dual) " . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }
        ScriptOpen('JavaScript');
        ShowHTML('  location.href=\'' . $w_dir . $w_pagina . $SG . '&O=L.CL=' . $CL . '\';');
        ScriptClose();
        Break;

      Case 'BASICO' :

        $SQL = "update sbpi.Cliente set " . $crlf .
               "   ds_senha_acesso = '" . trim($_REQUEST["w_ds_senha_acesso"]) . "', " . $crlf .
               "   dt_alteracao    = sysdate, " . $crlf;
        if ($_REQUEST["w_ds_email"] > "") {
          $SQL .= "   ds_email        = '" . trim($_REQUEST["w_ds_email"]) . "' " . $crlf;
        } else {
          $SQL .= "   ds_email        = null " . $crlf;
        }
        $SQL .= "where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        $w_funcionalidade = 21;
        //Grava o acesso na tabela de log
        $w_sql = $SQL;
        If ($_SESSION['USERNAME'] != 'SBPI') {
          $SQL = "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                 "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
                 "         " . $CL . ", " . $crlf .
                 "         sysdate, " . $crlf .
                 "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
                 "         2, " . $crlf .
                 "         'Atualização de dados básicos.', " . $crlf .
                 "           '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                 "           " . $w_funcionalidade . " " . $crlf .
                 "   from dual) " . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }

        ScriptOpen('JavaScript');
        ShowHTML('  location.href=\'' . $w_dir . $w_pagina . $SG . '&O=A&$CL=' . $CL . '\';');
        ScriptClose();
        break;

      Case 'GETSITE' :

        $SQL = "select a.ds_username, b.ds_username pai " . $crlf .
               "  from sbpi.Cliente a, " . $crlf .
               "       sbpi.CLiente b " . $crlf .
               " where b.sq_CLiente_pai is null " . $crlf .
               "   and a.sq_cliente = " . $CL;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        //$w_diretorio = str_replace($conFilePhysical . "\\" . $RS["pai"] .  "\\" . $RS["ds_username"] . "\\","\\\\","\\");
        $w_diretorio = $conFilePhysical . strtolower($_SESSION['USERNAME']);

        $SQL = "update sbpi.Cliente_Site set " . $crlf .
               "   no_contato_internet    = '" . trim($_REQUEST["w_no_contato_internet"]) . "', " . $crlf .
               "   ds_email_internet      = '" . trim($_REQUEST["w_ds_email_internet"]) . "', " . $crlf .
               "   nr_fone_internet       = '" . trim($_REQUEST["w_nr_fone_internet"]) . "', " . $crlf .
               "   ds_texto_abertura      = '" . trim($_REQUEST["w_ds_texto_abertura"]) . "', " . $crlf .
               "   ds_institucional       = '" . trim($_REQUEST["w_ds_institucional"]) . "', " . $crlf;
        if ($_REQUEST["w_nr_fax_internet"] > "") {
          $SQL .= "   nr_fax_internet        = '" . trim($_REQUEST["w_nr_fax_internet"]) . "', " . $crlf;
        } else {
          $SQL .= "   nr_fax_internet        = null, " . $crlf;
        }
        if ($_REQUEST["w_ds_mensagem"] > "") {
          $SQL .= "   ds_mensagem            = '" . trim($_REQUEST["w_ds_mensagem"]) . "' " . $crlf;
        } else {
          $SQL .= "   ds_mensagem            = null " . $crlf;
        }
        $SQL .= "where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        $w_sql = $SQL;

        $SQL = "update sbpi.Cliente set dt_alteracao = sysdate where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        if ($_FILES["w_pedagogica"]['size'] > 0) {

          if (file_exists($w_diretorio) == false) {
            mkdir($w_diretorio);
          }
          move_uploaded_file($_FILES["w_pedagogica"]['tmp_name'], $w_diretorio . '/' . $_FILES["w_pedagogica"]['name']);

          $w_imagem = $_FILES["w_pedagogica"]['name'];

          $SQL = "update sbpi.Cliente_Site set " . $crlf .
                 "   ln_prop_pedagogica = '" . $w_imagem . "' " . $crlf .
                 "where sq_cliente = " . $CL . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        }

        //Grava o acesso na tabela de log

        $w_funcionalidade = 23;
        If ($_SESSION['USERNAME'] != 'SBPI') {
          $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                 "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
                 "         " . $CL . ", " . $crlf .
                 "         sysdate, " . $crlf .
                 "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
                 "         2, " . $crlf .
                 "         'Atualização de dados do site.', " . $crlf .
                 "           '" . str_replace("'", "''", $w_sql) . "', " . $crlf .
                 "           " . $w_funcionalidade . " " . $crlf .
                 "   from dual) " . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }
        //       dbms.CommitTrans()

        ScriptOpen('JavaScript');
        ShowHTML('  location.href=\'' . $w_dir . $w_pagina . $SG . '&O=A&$CL=' . $CL . '\';');
        ScriptClose();
        Break;

      Case 'ARQUIVOS' :

        $w_funcionalidade = 25;
        $SQL = "select a.ds_username, b.ds_username pai " . $crlf .
        "  from sbpi.Cliente a, " . $crlf .
        "       sbpi.CLiente b " . $crlf .
        " where b.sq_CLiente_pai is null " . $crlf .
        "   and a.sq_cliente = " . $CL;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        $w_diretorio = $conFilePhysical . strtolower($_SESSION['USERNAME']);

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
                 "     ativo,   in_destinatario, nr_ordem,   ds_titulo, pasta) " . $crlf .
                 " values ( " . $w_chave . ", " . $crlf .
                 "     " . $CL . ", " . $crlf .
                 "     to_date('" . FormataDataEdicao(FormatDateTime(time(), 2)) . "','dd/mm/yyyy'), " . $crlf .
                 "     '" . $_REQUEST["w_ds_arquivo"] . "', " . $crlf .
                 "     '" . $w_imagem . "', " . $crlf .
                 "     '" . $_REQUEST["w_in_ativo"] . "', " . $crlf .
                 "     '" . $_REQUEST["w_in_destinatario"] . "', " . $crlf .
                 "     '" . $_REQUEST["w_nr_ordem"] . "', " . $crlf .
                 "     '" . $_REQUEST["w_ds_titulo"] . "', " . $crlf .
                 "     '" . $_REQUEST["w_pasta"] . "' " . $crlf .                 
                 " )" . $crlf;

          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

          //Grava o acesso na tabela de log
          $w_sql = $SQL;
          If ($_SESSION['USERNAME'] != 'SBPI') {
            $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                   "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
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
                   "     nr_ordem        = '" . $_REQUEST["w_nr_ordem"] . "', " . $crlf .
                   "     pasta        = '" . $_REQUEST["w_pasta"] . "' " . $crlf .                   
                   " where sq_arquivo = " . $_REQUEST["w_chave"] . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            //Grava o acesso na tabela de log
            $w_sql = $SQL;
            If ($_SESSION['USERNAME'] != 'SBPI') {
              $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                     "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
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
                       "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
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

      Case 'ADMINISTRATIVO' :
        //Apaga os equipamentos existentes para a unidade
        $SQL = "delete sbpi.Cliente_Equipamento where sq_cliente = " . $CL . $crlf;
        $w_sql = $SQL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        //Atualiza os dados em sbpi.CLiente_Admin
        $SQL = "select count(*) existe from sbpi.Cliente_Admin where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        foreach ($RS as $row) {
          $RS = $row;
          break;
        }
        if (intval(f($RS, "existe")) == 0) {
          $w_oper = 1;
          $SQL = "insert into sbpi.Cliente_Admin (sq_CLiente, limpeza_terceirizada, merenda_terceirizada, banheiro) " . $crlf .
                 " values (" . $CL . ", " . $crlf .
                 "         '" . $_REQUEST["w_limpeza"] . "', " . $crlf .
                 "         '" . $_REQUEST["w_merenda"] . "', " . $crlf .
                 "         '" . $_REQUEST["w_banheiro"] . "' " . $crlf .
                 "        )" . $crlf;
          $w_sql = $SQL . $crlf;
        } else {
          $w_oper = 2;
          $SQL = "update sbpi.Cliente_Admin set " . $crlf .
                 "   limpeza_terceirizada = '" . $_REQUEST["w_limpeza"] . "', " . $crlf .
                 "   merenda_terceirizada = '" . $_REQUEST["w_merenda"] . "', " . $crlf .
                 "   banheiro             = '" . $_REQUEST["w_banheiro"] . "' " . $crlf .
                 " where sq_cliente =  " . $CL . " " . $crlf;
          $w_sql = $SQL . $crlf;
        }
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        //Em seguida, grava os equipamentos com quantidade maior que zero
        for ($i = 1; $i < count($_REQUEST['w_equipamento']); $i++) {
          //  if( strlen(trim($_REQUEST['w_equipamento'][$i])) == 0 ){
          $SQL = "insert into sbpi.Cliente_Equipamento (sq_CLiente, sq_equipamento, quantidade) " . $crlf .
                 "values (" . $crlf .
                 "   " . $CL . ", " . $crlf .
                 "   " . $_REQUEST["w_codigo"][$i] . ", " . $crlf .
                 "  " . $_REQUEST['w_equipamento'][$i] . " " . $crlf .
                 ")";
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          $w_sql = $w_sql . $SQL . $crlf;
          //}

        }

        $w_funcionalidade = 29;
        //Grava o acesso na tabela de log
        If ($_SESSION['USERNAME'] != 'SBPI') {
          $SQL = "insert into sbpi.Cliente_Log (sq_CLiente_log, sq_CLiente, data, ip_origem, tipo, abrangencia, sql, sq_funcionalidade) " . $crlf .
                 "( select sbpi.sq_CLiente_log.nextval, " . $crlf .
                 "         " . $CL . ", " . $crlf .
                 "         sysdate, " . $crlf .
                 "         '" . $_SERVER['REMOTE_ADDR'] . "', " . $crlf .
                 "         " . $w_oper . ", " . $crlf .
                 "         'Informações administrativas da escola.', " . $crlf .
                 "           '" . str_replace("'", "''", substr($w_sql, 0, 1000)) . "', " . $crlf .
                 "           " . $w_funcionalidade . " " . $crlf .
                 "   from dual) " . $crlf;

          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        }
        $SQL = "update sbpi.Cliente set dt_alteracao = sysdate where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        ScriptOpen('JavaScript');
        ShowHTML('location.href=\'' . $w_dir . $w_pagina . $SG . '&O=L&CL=' . $CL . '\';');
        ScriptClose();
        Break;

      Case 'FOTOS' :

        $SQL = "select a.ds_username, b.ds_username pai " . $crlf .
               "  from sbpi.Cliente a, " . $crlf .
               "       sbpi.CLiente b " . $crlf .
               " where b.sq_CLiente_pai is null " . $crlf .
               "   and a.sq_cliente = " . $CL;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        $w_diretorio = $conFilePhysical . strtolower($_SESSION['USERNAME']);
        if ($O == "I") {
          //Verifica se o nome desse arquivo já existe
          $SQL = "select count(*) existe from sbpi.Cliente_Foto where ln_foto = '" . $_FILES["w_ln_foto"]['name'] . "'" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          if (intval($RS[0]["existe"]) > 0) {
            ScriptOpen('JavaScript');
            ShowHTML('  alert(\'Nome do arquivo já existe!\');');
            ShowHTML('  history.back(1);');
            ScriptClose();
            exit ();
          }
          $SQL = "select count(*) existe from sbpi.Cliente_Arquivo where ln_arquivo = '" . $_FILES["w_ln_arquivo"]['name'] . "'" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          if (intval($RS[0]["existe"]) > 0) {
            ScriptOpen('JavaScript');
            ShowHTML('  alert(\'Nome do arquivo já existe!\');');
            ShowHTML('  history.back(1);');
            ScriptClose();
            exit ();
          }

          move_uploaded_file($_FILES["w_ln_foto"]['tmp_name'], $w_diretorio . '/' . $_FILES["w_ln_foto"]['name']);
          $w_imagem = $_FILES["w_ln_foto"]['name'];

          //Recupera o valor da próxima chave primária
          $SQL = "select sbpi.sq_CLiente_foto.nextval chave from sbpi.Cliente_Foto" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
          foreach ($RS as $row) {
            $RS = $row;
            break;
          }
          $w_chave = f($RS, "chave");

          //Insere o arquivo
          $SQL = " insert into sbpi.Cliente_Foto " . $crlf .
                 "    (sq_CLiente_foto, sq_CLiente, ds_foto, ln_foto, tp_foto, nr_ordem) " . $crlf .
                 " values ( " . $w_chave . ", " . $crlf .
                 "     " . $CL . ", " . $crlf .
                 "     '" . $_REQUEST["w_ds_foto"] . "', " . $crlf .
                 "     '" . $w_imagem . "', " . $crlf .
                 "     '" . $_REQUEST["w_tp_foto"] . "', " . $crlf .
                 "     '" . $_REQUEST["w_nr_ordem"] . "' " . $crlf .
                 " )" . $crlf;
          $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

        } else
          if ($O == "A") {

            //Verifica se o nome desse arquivo já existe
            $SQL = "select count(*) existe from sbpi.Cliente_Foto where ln_foto = '" . $_FILES["w_ln_foto"]['name'] . "' and sq_CLiente_foto <> " . $_REQUEST["w_chave"] . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            foreach ($RS as $row) {
              $RS = $row;
              break;
            }
            if (intval(f($RS, "existe")) > 0) {
              ScriptOpen('JavaScript');
              ShowHTML('  alert(\'Nome do arquivo já existe!\');');
              ShowHTML('  history.back(1);');
              ScriptClose();
              Exit ();
            }

            //Remove o arquivo físico
            $SQL = "select ln_foto foto from sbpi.Cliente_Foto where sq_CLiente_foto = " . $_REQUEST["w_chave"];
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            foreach ($RS as $row) {
              $RS = $row;
              break;
            }
            $w_foto = f($RS, "foto");

            $SQL = " update sbpi.Cliente_Foto set " . $crlf .
                   "     ds_foto       = '" . $_REQUEST["w_ds_foto"] . "', " . $crlf .
                   "     nr_ordem      = '" . $_REQUEST["w_nr_ordem"] . "' " . $crlf .
                   "where sq_CLiente_foto = " . $_REQUEST["w_chave"] . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

            if ($_FILES["w_ln_foto"]['size'] > 0) {
              //Remove o arquivo físico
              unlink($w_diretorio . '/' . $w_foto);

              if (file_exists($w_diretorio) == false) {
                mkdir($w_diretorio);
              }

              move_uploaded_file($_FILES["w_ln_foto"]['tmp_name'], $w_diretorio . '/' . $_FILES["w_ln_foto"]['name']);
              $w_imagem = $_FILES["w_ln_foto"]['name'];

              $SQL = " update sbpi.Cliente_Foto set " . $crlf .
                     "     ln_foto      = '" . $w_imagem . "' " . $crlf .
                     " where sq_CLiente_foto = " . $_REQUEST["w_chave"] . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }

          } else
            if ($O == "E") {
              //Remove o arquivo físico
              $SQL = "select ln_foto from sbpi.Cliente_Foto where sq_CLiente_foto = " . $_REQUEST["w_chave"];
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
              foreach ($RS as $row) {
                $RS = $row;
                break;
              }
              unlink($w_diretorio . '/' . f($RS, "ln_foto"));

              $SQL = " delete sbpi.Cliente_Foto where sq_CLiente_foto = " . $_REQUEST["w_chave"] . $crlf;
              $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            }

        $SQL = "update sbpi.Cliente set dt_alteracao = sysdate where sq_cliente = " . $CL . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
        foreach ($RS as $row) {
          $RS = $row;
          break;
        }

        ScriptOpen('JavaScript');
        ShowHTML(' location.href=\'' . $w_dir . $w_pagina . $SG . '&O=L\';');
        ScriptClose();

        Break;

      Default :
        ScriptOpen('JavaScript');
        ShowHTML('  alert(\'Bloco de dados não encontrado: ' . $SG . '\');');
        ShowHTML('  history.back(1);');
        ScriptClose();
        Break;

    }
  }

  // =========================================================================
  // Tela de dados adicionais
  // -------------------------------------------------------------------------
  function Adicionais() {

    extract($GLOBALS);
    $O = "A";
    if ($O == "A") {
      $SQL = "select * from sbpi.Cliente_Dados where sq_cliente = " . $CL;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      If (count($RS) > 0) {
        $w_sq_cliente = $RS["sq_cliente"];
        $w_nr_cnpj = $RS["nr_cnpj"];
        $w_tp_registro = $RS["tp_registro"];
        $w_ds_ato = $RS["ds_ato"];
        $w_nr_ato = $RS["nr_ato"];
        $w_dt_ato = $RS["dt_ato"];
        $w_ds_orgao = $RS["ds_orgao"];
        $w_ds_logradouro = $RS["ds_logradouro"];
        $w_no_bairro = $RS["no_bairro"];
        $w_nr_cep = trim(str_replace(".", "", $RS["nr_cep"]));
        $w_no_contato = $RS["no_contato"];
        $w_ds_email_contato = $RS["ds_email_contato"];
        $w_nr_fone_contato = $RS["nr_fone_contato"];
        $w_nr_fax_contato = $RS["nr_fax_contato"];
        $w_no_diretor = $RS["no_diretor"];
        $w_no_secretario = $RS["no_secretario"];
      } Else {
        $w_sq_cliente = "";
        $w_nr_cnpj = "";
        $w_tp_registro = "";
        $w_ds_ato = "";
        $w_nr_ato = "";
        $w_dt_ato = "";
        $w_ds_orgao = "";
        $w_ds_logradouro = "";
        $w_no_bairro = "";
        $w_nr_cep = "";
        $w_no_contato = "";
        $w_ds_email_contato = "";
        $w_nr_fone_contato = "";
        $w_nr_fax_contato = "";
        $w_no_diretor = "";
        $w_no_secretario = "";
      }
    }

    Cabecalho();
    ShowHTML("<HEAD>");
    ScriptOpen("Javascript");
    CheckBranco();
    FormataDATA();
    FormataCEP();
    ValidateOpen("Validacao");

    If ($O == "A") {
      Validate("w_ds_logradouro", "Logradouro", "1", "1", "4", "60", "1", "1");
      Validate("w_no_bairro", "Bairro", "1", "", "2", "30", "1", "1");
      Validate("w_nr_cep", "CEP", "1", "1", "9", "9", "", "0123456789-");
      Validate("w_no_diretor", "Nome do(a) Diretor(a)", "1", "", "2", "40", "1", "1");
      If ($_SESSION['TIPO'] == 3) {
        Validate("w_no_secretario", "Nome do(a) Secretário(a)", "1", "", "2", "40", "1", "1");
      }

      Validate("w_no_contato", "Nome", "1", "1", "2", "35", "1", "1");
      Validate("w_ds_email_contato", "e-Mail", "1", "", "6", "60", "1", "1");
      Validate("w_nr_fone_contato", "Telefone", "1", "1", "6", "20", "1", "1");
      Validate("w_nr_fax_contato", "Fax", "1", "", "6", "20", "1", "1");
    }
    ValidateClose();
    ScriptClose();
    ShowHTML("</HEAD>");
    BodyOpen("onLoad='document.Form.w_ds_logradouro.focus();'");
    ShowHTML('<B><FONT COLOR="#000000">Atualização de dados adicionais da unidade</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
    AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
    ShowHTML(MontaFiltro("POST"));
    ShowHTML('<input type="hidden" name="O" value="' . $O . '">');
    ShowHTML('<input type="hidden" name="SG" value="adicionais">');
    ShowHTML('<tr bgcolor="#EFEFEF"><td align="center">');
    ShowHTML('    <table width="97%" border="0">');
    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Endereço</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Informe o endereço da unidade, a ser exibido na seção "Quem Somos" do site.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top"><font size="1"><b><u>L</u>ogradouro:</b><br><INPUT ACCESSKEY="L" ' . $w_Disabled . '  class="STI" type="text" name="w_ds_logradouro" size="60" maxlength="60" value="' . $w_ds_logradouro . '" title="OBRIGATÓRIO. Informe o logradouro de funcionamento da unidade."></td>');
    ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML('        <tr valign="top">');
    ShowHTML('          <td><font size="1"><b><u>B</u>airro:</b><br><INPUT ACCESSKEY="B" ' . $w_Disabled . '  class="STI" type="text" name="w_no_bairro" size="30" maxlength="30" value="' . $w_no_bairro . '" title="OPCIONAL. Informe o bairro de funcionamento da unidade."></td>');
    ShowHTML('          <td><font size="1"><b>C<u>E</u>P:</b><br><INPUT ACCESSKEY="C" ' . $w_Disabled . '  class="STI" type="text" name="w_nr_cep" size="9" maxlength="9" value="' . $w_nr_cep . '" title="OBRIGATÓRIO. Informe o CEP da unidade." onKeyDown="FormataCEP(this,event);"></td>');
    ShowHTML('        </table>');
    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Dirigentes</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Os dados deste bloco são exibidos na seção "Quem Somos" do site.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML('        <tr valign="top">');
    ShowHTML('          <td><font size="1"><b>Nome do(a) <u>D</u>iretor:</b><br><INPUT ACCESSKEY="D" ' . $w_Disabled . ' class="STI" type="text" name="w_no_diretor" size="40" maxlength="40" value="' . $w_no_diretor . '" title="OPCIONAL. Informe o nome do(a) diretor(a) da unidade"></td>');
    If ($_SESSION['TIPO'] == 3) {
      ShowHTML('          <td><font size="1"><b>Nome do(a)<u>S</u>ecretário(a):</b><br><INPUT ACCESSKEY="S" ' . $w_Disabled . ' class="STI" type="text" name="w_no_secretario" size="40" maxlength="40" value="' . $w_no_secretario . '" title="OPCIONAL. Informe o nome do(a) secretário(a) da unidade."></td>');
    }
    ShowHTML('        </table>');

    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Contato técnico da unidade</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Informe os dados da pessoa de contato técnico para uso interno, não sendo utilizado para divulgação no site.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
    ShowHTML('       <tr valign="top">');
    ShowHTML('          <td><font size="1"><b>No<u>m</u>e:</b><br><INPUT ACCESSKEY="M" ' . $w_Disabled . '  class="STI" type="text" name="w_no_contato" size="35" maxlength="35" value="' . $w_no_contato . '" title="OBRIGATÓRIO. Informe o nome da pessoa a ser contactada pela equipe técnica. Os dados deste contato não serão exibidos no site."></td>');
    ShowHTML('          <td><font size="1"><b><u>e</u>-Mail:</b><br><INPUT ACCESSKEY="E" ' . $w_Disabled . '  class="STI" type="text" name="w_ds_email_contato" size="40" maxlength="60" value="' . $w_ds_email_contato . '" title="OPCIONAL. Informe o e-mail da pessoa de contato técnico."></td>');
    ShowHTML('        <tr valign="top">');
    ShowHTML('          <td><font size="1"><b>Tele<u>f</u>one:</b><br><INPUT ACCESSKEY="F" ' . $w_Disabled . '  class="STI" type="text" name="w_nr_fone_contato" size="20" maxlength="20" value="' . $w_nr_fone_contato . '" title="OBRIGATÓRIO. Informe o telefone da pessoa de contato técnico."></td>');
    ShowHTML('          <td><font size="1"><b>Fa<u>x</u>:</b><br><INPUT ACCESSKEY="X" ' . $w_Disabled . '  class="STI" type="text" name="w_nr_fax_contato" size="20" maxlength="20" value="' . $w_nr_fax_contato . '" title="OPCIONAL. Informe o fax da pessoa de contato técnico."></td>');
    ShowHTML('        </table>');
    ShowHTML('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');

    // Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
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
  // Fim da tela de dados adicionais
  // --------------------------------

  // =========================================================================
  // Rotina de alteração dos recursos da etapa
  // -------------------------------------------------------------------------
  function atuacao() {
    extract($GLOBALS);

    If ($O == "L") {
      $SQL = "select a.sq_especialidade, a.ds_especialidade, a.tp_especialidade, b.sq_especialidade as sq_codigo_cli " .
             "  from sbpi.Especialidade a " .
             "       left outer join sbpi.Especialidade_Cliente b on (a.sq_especialidade = b.sq_especialidade and b.sq_cliente = " . $CL . ") " .
             "where tp_especialidade <> 'M' order by a.nr_ordem, a.ds_especialidade ";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }

    Cabecalho();
    ShowHTML("<HEAD>");
    ScriptOpen("Javascript");
    ValidateOpen("Validacao");
    If ($O == "L") {

      ShowHTML('  var w_erro=true; ');
      ShowHTML('  if (theForm["w_chave[]"].value==undefined) {');
      ShowHTML('     for (i=0; i < theForm["w_chave[]"].length; i++) {');
      ShowHTML('       if (theForm["w_chave[]"][i].checked) w_erro=false;');
      ShowHTML('     }');
      ShowHTML('  }');
      ShowHTML('  else {');
      ShowHTML('     if (theForm["w_chave[]"].checked) w_erro=false;');
      ShowHTML('  }');
      ShowHTML('  if (w_erro) {');
      ShowHTML('    alert(\'Você deve selecionar pelo menos uma área de atuação!\'); ');
      ShowHTML('    return false;');
      ShowHTML('  }');
    }
    ValidateClose();
    ScriptClose();
    ShowHTML("</HEAD>");
    BodyOpen("onLoad='document.focus();'");
    ShowHTML('<B><FONT COLOR="#000000">Atualização de áreas de atuação</FONT></B>');
    ShowHTML("<HR>");
    ShowHTML('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">  ');
    AbreForm("Form", $w_pagina . "Grava", "POST", "return(Validacao(this));", null);
    ShowHTML(MontaFiltro("POST"));
    ShowHTML('<input type="hidden" name="O" value="' . $O . '">');
    ShowHTML('<input type="hidden" name="SG" value="ATUACAO">');
    ShowHTML('<tr bgcolor="' . "#EFEFEF" . '"><td align="center">');
    ShowHTML('    <table width="97%" border="0">');
    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Áreas de atuação</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Selecione áreas de atuação oferecidas pela unidade.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');

    if (count($RS) > 0) {
      $wAtual = '';

      foreach ($RS as $row) {
        If (($wAtual == '') or ($wAtual != $row["tp_especialidade"])) {
          $wAtual = $row["tp_especialidade"];
          If ($wAtual == "J") {
            ShowHTML('          <TR><TD><font size=1><b>Etapas / Modalidades de ensino</b>:');
          }
          ElseIf ($wAtual == "R") {
            ShowHTML('          <TR><TD><font size=1><b>Em Regime de Intercomplementaridade</b>:');
          } Else {
            ShowHTML('          <TR><TD><font size=1><b>Outras</b>:');
          }
        }

        If (!empty ($row["sq_codigo_cli"])) {
          ShowHTML('      <tr valign="top"><td><font size="1"><input type="checkbox" name="w_chave[]" value=" ' . $row["sq_especialidade"] . '" checked>' . $row["ds_especialidade"] . '</td></tr>');
        } Else {
          ShowHTML('      <tr valign="top"><td><font size="1"><input type="checkbox" name="w_chave[]" value=" ' . $row["sq_especialidade"] . '">' . $row["ds_especialidade"] . '</td></tr>');
        }
      }
    }

    ShowHTML("      </center>");
    ShowHTML("    </table>");
    ShowHTML("  </td>");
    ShowHTML("</tr>");

    ShowHTML('      <tr><td align="center"><font size=1>&nbsp;');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000">');
    ShowHTML('      <tr><td align="center">');
    ShowHTML('            <input class="STB" type="submit" name="Botao" value="Gravar">');
    ShowHTML("          </td>");
    ShowHTML("      </tr>");
    ShowHTML("</table>");
    ShowHTML("</center>");
    ShowHTML("</FORM>");
    Rodape();

  }
  // =========================================================================
  // Fim da rotina
  //-------------------------------------------------------------------------

  // =========================================================================
  // Cadastro de mensagens ao aluno
  // -------------------------------------------------------------------------
  function msgalunos() {
    extract($GLOBALS);

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];
    $w_texto = $_REQUEST["w_texto"];

    if ($w_troca > '') { // Se for recarga da página
      $w_dt_mensagem = $_REQUEST["w_dt_mensagem"];
      $w_ds_mensagem = $_REQUEST["w_ds_mensagem"];
      $w_sq_aluno = $_REQUEST["w_sq_aluno"];
    } else
      if ($O == 'L') {
        //Recupera todos os registros para a listagem
        $SQL = "select a.sq_mensagem as chave, a.in_lida, b.no_aluno, b.nr_matricula, c.ds_cliente " . $crlf .
               "  from sbpi.Mensagem_Aluno     a " . $crlf .
               "       inner join sbpi.Aluno   b on (a.sq_aluno        = b.sq_aluno) " . $crlf .
               "       inner join sbpi.Cliente c on (b.sq_cliente = c.sq_cliente) " . $crlf .
               "       where c.sq_cliente =  " . $CL . $crlf .
               "order by c.ds_cliente, a.dt_mensagem desc, b.no_aluno, a.in_lida " . $crlf;
        $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

      } else
        if (strpos("AEV", $O) !== false && $w_troca == '') {
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
    } else
      if ($O == "I" or $O == "A") {
        BodyOpen("onLoad='document.Form.w_ds_mensagem.focus()';");
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
      ShowHTML('          <td><font size="1"><b>Unidade</font></td>');
      ShowHTML('          <td><font size="1"><b>Data</font></td>');
      ShowHTML('          <td><font size="1"><b>Matrícula</font></td>');
      ShowHTML('          <td><font size="1"><b>Aluno</font></td>');
      ShowHTML('          <td><font size="1"><b>Lida</font></td>');
      ShowHTML('          <td><font size="1"><b>Operações</font></td>');
      ShowHTML('        </tr>');

      if (count($RS) <= 0) {
        // Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="' . $conTrBgColor . '"><td colspan=5 align="center"><b>Não foram encontrados registros.</b></td></tr>');
      } else {
        foreach ($RS as $row) {
          $w_cor = ($w_cor == $conTrBgColor || $w_cor == '') ? $w_cor = $conTrAlternateBgColor : $w_cor = $conTrBgColor;
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td><font size="1">' . strtolower(f($row, "ds_cliente")) . '</td>');
          ShowHTML('        <td align="center"><font size="1">' . FormataDataEdicao(f($row, "dt_mensagem"), 2) . '</td>');
          ShowHTML('        <td align="center" nowrap><font size="1">' . f($row, "nr_matricula") . '</td>');
          ShowHTML('        <td><font size="1">' . strtolower(f($row, "no_aluno")) . '</td>');
          ShowHTML('        <td align="center"><font size="1">' . f($row, "in_lida") . '</td>');
          ShowHTML('        <td align="top" nowrap><font size="1">');
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
    } else
      if (strpos("IAEV", $O) !== false) {
        if (strpos("EV", $O) !== false) {
          $w_Disabled = ' DISABLED ';
        }
        AbreForm('Form', $w_dir . $w_pagina . 'Grava', 'POST', 'return(Validacao(this));', null);
        ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');
        ShowHTML('<INPUT type="hidden" name="w_sq_cliente" value="' . $_SESSION["CL"] . '">');
        ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');
        ShowHTML('<INPUT type="hidden" name="SG" value="MENSAGENS">');
        ShowHTML('<INPUT type="hidden" name="w_troca" value="">');

        ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
        ShowHTML('    <table width="95%" border="0">');
        ShowHTML('      <tr><td><font size="1"><b><u>D</u>ata:</b><br><input ' . $w_Disabled . ' accesskey="D" type="text" name="w_dt_mensagem" class="STI" SIZE="10" MAXLENGTH="10" VALUE="' . FormataDataEdicao(FormatDateTime(Nvl($w_dt_mensagem, Time()), 2)) . '" onKeyDown="FormataData(this,event);"></td>');
        ShowHTML('      <tr><td><font size="1"><b>M<u>e</u>nsagem:</b><br><textarea ' . $w_Disabled . ' accesskey="E" name="w_ds_mensagem" class="STI" ROWS=5 cols=65>' . $w_ds_mensagem . '</TEXTAREA>');
        if ($O == "I") {
          ShowHTML('      <tr><td><font size="1"><b><u>P</u>rocurar por:</b><br><input accesskey="P" type="text" name="w_texto" class="STI" SIZE="50" MAXLENGTH="50" VALUE="' . $w_texto . '" >');
          ShowHTML('          <input type="Button" class="STB" name="Pesquisa" value="Procurar" onClick="document.Form.w_troca.value=\'' . $w_sq_aluno . '\'; document.Form.action=\'' . $dir . $w_pagina . $par . $w_ew . '\'; document.Form.submit();"></td>');
          ShowHTML('      <tr><td><font size="1"><b><u>A</u>luno:</b><br><select accesskey="A" name="w_sq_aluno" class="STS" SIZE="1" >');
          ShowHTML('          <OPTION VALUE="">---');
          If ($w_texto > '') {
            $SQL = "select sq_aluno, nr_matricula, no_aluno " . $crlf .
                   "  from sbpi.Aluno " . $crlf .
                   " where (upper(no_aluno) like '%" . strtoupper($w_texto) . "%' or " . $crlf .
                   "        nr_matricula like '%" . $w_texto . "%') and sq_cliente =  " . $CL . $crlf .
                   "order by no_aluno, nr_matricula" . $crlf;
            $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
            foreach ($RS as $row) {
              ShowHTML('          <OPTION VALUE="' . f($row, "sq_aluno") . '">' . f($row, "no_aluno") . ' ' . f($row, "nr_matricula"));
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
          ShowHTML('      <tr><td><font size="1"><b>Aluno:<br>' . f($RS, "no_aluno") . ' (' . f($RS, "nr_matricula") . ')</td>');
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
  // Tela de cadastro de fotos
  // -------------------------------------------------------------------------
  function fotos() {
    extract($GLOBALS);

    $O = $_REQUEST['O'];
    if (is_null($O))
      $O = "L";

    $w_chave = $_REQUEST["w_chave"];
    $w_troca = $_REQUEST["w_troca"];
    $w_tp_foto = $_REQUEST["w_tp_foto"];

    If ($w_troca > "") { //' Se for recarga da página
      $w_ds_foto = $_REQUEST["w_ds_foto"];
      $w_ln_foto = $_REQUEST["w_ln_foto"];
      $w_nr_ordem = $_REQUEST["w_nr_ordem"];

    }
    ElseIf ($O == "L") {

      //' Recupera todos os registros para a listagem
      $SQL = "select * from sbpi.Cliente_Foto where sq_cliente = " . $CL . " and tp_foto = 'P' order by nr_ordem";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

      $SQL = "select * from sbpi.Cliente_Foto where sq_cliente = " . $CL . " and tp_foto = 'Q' order by nr_ordem";
      $RS1 = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    }
    ElseIf (strpos("AEV", $O) !== false and is_null($w_troca)) {
      // ' Recupera os dados do endereço informado

      $SQL = "select b.ds_diretorio, a.* from sbpi.Cliente_Foto a inner join sbpi.Cliente_Site b on (a.sq_cliente = b.sq_cliente) where a.sq_cliente_foto = " . $w_chave;
      $RS3 = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

      foreach ($RS3 as $row) {
        $RS3 = $row;
        break;
      }

      $w_ds_foto = $row["ds_foto"];
      $w_tp_foto = $row["tp_foto"];
      $w_ln_foto = $row["ln_foto"];
      $w_nr_ordem = $row["nr_ordem"];
      $w_ds_diretorio = $row["ds_diretorio"];

    }

    Cabecalho();
    ShowHTML("<HEAD>");
    If (strpos("IAEP", $O) !== false) {
      ScriptOpen("JavaScript");
      ValidateOpen("Validacao");
      If (strpos("IA", $O) !== false) {
        Validate("w_ds_foto", "Título", "", "1", "2", "100", "1", "1");
        If ($O == "I" || $O == "A") {
          if ($O == "I") {
            Validate("w_ln_foto", "link", "1", "1", "5", "100", "1", "1");
          } else {
            Validate("w_ln_foto", "link", "1", "", "5", "100", "1", "1");
          }
          ShowHTML(" if (theForm.w_ln_foto.value > ''){");
          ShowHTML("    if((theForm.w_ln_foto.value.lastIndexOf('.jpg')==-1) && (theForm.w_ln_foto.value.lastIndexOf('.gif')==-1) && (theForm.w_ln_foto.value.lastIndexOf('.JPG')==-1) && (theForm.w_ln_foto.value.lastIndexOf('.GIF')==-1)) {");
          ShowHTML("       alert('Só é possível escolher arquivos com a extensão \'.jpg\' ou \'.gif\'!');");
          ShowHTML("       theForm.w_ln_foto.focus(); ");
          ShowHTML("       return false;");
          ShowHTML("    }");
          ShowHTML("  }");
        }
        Validate("w_nr_ordem", "Nr. de ordem", "", "1", "1", "2", "1", "0123546789");
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
    ElseIf (($O == "I") or ($O == "A")) {
      BodyOpen("onLoad='document.Form.w_ds_foto.focus()';");
    } Else {
      BodyOpen("onLoad='document.focus()';");
    }
    ShowHTML('<B><FONT COLOR="#000000">Cadastro de fotos da unidade</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');
    If ($O == "L") {

      ShowHTML('      <tr><td colspan="3" align="center" height="2" bgcolor="#000000"></td></tr>');
      ShowHTML('      <tr><td colspan="3" align="center" height="1" bgcolor="#000000"></td></tr>');
      ShowHTML('      <tr><td colspan="3" valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Fotos da página de abertura do site</td></td></tr>');
      ShowHTML('      <tr><td colspan="3" align="center" height="1" bgcolor="#000000"></td></tr>');
      ShowHTML('      <tr><td colspan="3"><font size=1>Informe até 5 imagens a serem colocadas na página de abertura do site. A exibição será feita no formato 302px (largura) por 206px (altura). Recomenda-se imagens com extensão JPG, de até 40KB de tamanho.</font></td></tr>');
      ShowHTML('      <tr><td colspan="3" align="center" height="1" bgcolor="#000000"></td></tr>');

      //    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem

      If (count($RS) < 5) {
        ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $w_pagina . $par . '&O=I&w_tp_foto=P&R=' . $w_pagina . $par . '"><u>I</u>ncluir</a>&nbsp;');
      } Else {
        ShowHTML('<tr><td><font size="2">&nbsp;');
      }
      ShowHTML('    <td align="right"><font size="1"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor= "' . $conTableBgColor . '" BORDER=" ' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING=" ' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="#EFEFEF" align="center">');
      ShowHTML('          <td><font size="1"><b>Ordem</font></td>');
      ShowHTML('          <td><font size="1"><b>Descrição</font></td>');
      ShowHTML('          <td><font size="1"><b>Operações</font></td>');
      ShowHTML('        </tr>');

      If (count($RS) == 0) { //' Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="#EFEFEF"><td colspan=7 align="center"><font size="2"><b>Não foram encontrados registros.</b></td></tr>');
      } Else {
        foreach ($RS as $row) {
          If ($w_cor = "#EFEFEF" or $w_cor = '')
            $w_cor = "#FDFDFD";
          Else
            $w_cor = "#EFEFEF";
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td align="center"><font size="1">' . $row["nr_ordem"] . "</td>");
          ShowHTML('        <td><font size="1">' . $row["ds_foto"] . "</td>");
          ShowHTML('        <td align="top" nowrap><font size="1">');
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . $par . '&O=A' . '&w_chave=' . $row["sq_cliente_foto"] . '&R=' . $w_pagina . $par . '">Alterar</A>&nbsp');
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . $par . '&O=E' . '&w_chave=' . $row["sq_cliente_foto"] . '&R=' . $w_pagina . $par . '">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
        }
      }
      ShowHTML("      </center>");
      ShowHTML("    </table>");
      ShowHTML("  </td>");
      ShowHTML("</tr>");
      ShowHTML('<tr><td colspan="3"><br><br></td></tr>');
      ShowHTML('      <tr><td colspan="3" align="center" height="2" bgcolor="#000000"></td></tr>');
      ShowHTML('      <tr><td colspan="3" align="center" height="1" bgcolor="#000000"></td></tr>');
      ShowHTML('      <tr><td colspan="3" valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Fotos da página "Quem somos"</td></td></tr>');
      ShowHTML('      <tr><td colspan="3" align="center" height="1" bgcolor="#000000"></td></tr>');
      ShowHTML('      <tr><td colspan="3"><font size=1>Informe até 10 imagens a serem colocadas na página "Quem somos" do site. A exibição será feita no formato 201px (largura) por 153px (altura). Recomenda-se imagens com extensão JPG, de até 40KB de tamanho.</font></td></tr>');
      ShowHTML('      <tr><td colspan="3" align="center" height="1" bgcolor="#000000"></td></tr>');
      //' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      If (count($RS1) < 10) {
        ShowHTML('<tr><td><font size="2"><a accesskey="I" class="SS" href="' . $w_pagina . $par . '&O=I&w_tp_foto=Q&R=' . $w_pagina . $par . '"><u>I</u>ncluir</a>&nbsp;');
      } Else {
        ShowHTML('<tr><td><font size="2">&nbsp;');
      }
      ShowHTML('    <td align="right"><font size="1"><b>Registros existentes: ' . count($RS1));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('    <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING=" ' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="EFEFEF" align="center">');
      ShowHTML('          <td><font size="1"><b>Ordem</font></td>');
      ShowHTML('          <td><font size="1"><b>Descrição</font></td>');
      ShowHTML('          <td><font size="1"><b>Operações</font></td>');
      ShowHTML('        </tr>');

      If (count($RS1) == 0) { //' Se não foram selecionados registros, exibe mensagem
        ShowHTML('      <tr bgcolor="#EFEFEF"><td colspan=7 align="center"><font size="2"><b>Não foram encontrados registros.</b></td></tr>');
      } Else {
        foreach ($RS1 as $row) {
          If ($w_cor = "#EFEFEF" or $w_cor = "")
            $w_cor = "#FDFDFD";
          Else
            $w_cor = "#EFEFEF";
          ShowHTML('      <tr bgcolor="' . $w_cor . '" valign="top">');
          ShowHTML('        <td align="center"><font size="1">' . $row["nr_ordem"] . "</td>");
          ShowHTML('        <td><font size="1">' . $row["ds_foto"] . "</td>");
          ShowHTML('       <td align="top" nowrap><font size="1">');
          //ShowHTML ('          <A class="HL" HREF="' . $w_pagina . "&R=" . $w_pagina . "&O=A&w_chave=" . $row["sq_cliente_foto"] . ">Alterar</A>&nbsp");
          //ShowHTML ('          <A class="HL" HREF="' . $w_pagina . "&R=" . $w_pagina . "&O=E&w_chave=" . $row["sq_cliente_foto"] . ">Excluir</A>&nbsp");
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . $par . '&O=A' . '&w_chave=' . $row["sq_cliente_foto"] . '&R=' . $w_pagina . $par . '">Alterar</A>&nbsp');
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . $par . '&O=E' . '&w_chave=' . $row["sq_cliente_foto"] . '&R=' . $w_pagina . $par . '">Excluir</A>&nbsp');
          ShowHTML('        </td>');
          ShowHTML('      </tr>');
        }
        ShowHTML("      </center>");
        ShowHTML("    </table>");
        ShowHTML("  </td>");
        ShowHTML("</tr>");
      }
    }
    elseif (strpos("IAEV", $O) !== false) {
      If (strpos("EV", $O) !== false) {
        $w_Disabled = " DISABLED ";
      }
      ShowHTML('<FORM action="' . $w_pagina . 'Grava" method="POST" name="Form" onSubmit="return(Validacao(this));" enctype="multipart/form-data">');
      ShowHTML('<INPUT type="hidden" name="SG" value="FOTOS">');
      ShowHTML('<INPUT type="hidden" name="w_chave" value="' . $w_chave . '">');

      ShowHTML('<INPUT type="hidden" name="O" value="' . $O . '">');
      ShowHTML('<INPUT type="hidden" name="w_tp_foto" value="' . $w_tp_foto . '">');

      ShowHTML('<tr bgcolor="#EFEFEF" ><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td valign="top"><font size="1"><b><u>T</u>ítulo:</b><br><input "' . $w_Disabled . '" accesskey="T" type="text" name="w_ds_foto" class="STI" SIZE="60" MAXLENGTH="100" VALUE="' . $w_ds_foto . '" title="OBRIGATÓRIO. Informe um título para o arquivo."></td>');
      ShowHTML('        </table>');
      ShowHTML('      <tr>');
      ShowHTML('     </tr>');
      ShowHTML('     <tr><td valign="top"><font size="1"><b><u>L</u>ink:</b><br><input "' . $w_Disabled . '" accesskey="L" type="file" name="w_ln_foto" class="STI" SIZE="80" MAXLENGTH="100" VALUE="" title="OBRIGATÓRIO. Clique no botão ao lado para localizar o arquivo. Ele será transferido automaticamente para o servidor.">');
      If ($w_ln_foto > "") {
        ShowHTML('        <b><a class="SS" href="' . strtolower($w_ds_diretorio) . "/" . $w_ln_foto . '" target="_blank" title="Clique para exibir o arquivo atual.<br><img src=\'' . strtolower($w_ds_diretorio) . "/" . $w_ln_foto . '\' hight=50px width=70px >">Exibir</a></b>');
      }
      ShowHTML('      <tr><td valign="top" colspan="2"><table border=0 width="100%" cellspacing=0>');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td><font size="1"><b><u>N</u>r. de ordem:</b><br><input "' . $w_Disabled . '" accesskey="N" type="text" name="w_nr_ordem" class="STI" SIZE="2" MAXLENGTH="2" VALUE="' . $w_nr_ordem . '" title="OBRIGATÓRIO. Informe a posição em que este arquivo deve aparecer na lista de arquivos disponíveis. Ex: 1, 2, 3 etc."></td>');
      ShowHTML('        </table>');
      ShowHTML('      <tr>');
      ShowHTML('      <tr><td align="center" colspan=4><hr>');
      If ($O == "E") {
        ShowHTML(' <input class="STB" type="submit" name="Botao" value="Excluir" onClick="return confirm(\'Confirma a exclusão do registro?\');">');
      } Else {
        If ($O == "I") {
          ShowHTML('           <input class="STB" type="submit" name="Botao" value="Incluir">');
        } Else {
          ShowHTML('            <input class="STB" type="submit" name="Botao" value="Atualizar">');
        }
      }
      ShowHTML('            <input class="STB" type="button" onClick="location.href=\'' . $w_pagina . $par . '\'"    name="Botao" value="Cancelar">');
      ShowHTML("          </td>");
      ShowHTML("      </tr>");
      ShowHTML("    </table>");
      ShowHTML("    </TD>");
      ShowHTML("</tr>");
      ShowHTML("</FORM>");

    } Else {

      ScriptOpen("JavaScript");
      ShowHTML(" alert('Opção não disponível');");
      ScriptClose();
    }
    ShowHTML("</table>");
    ShowHTML("</center>");
    Rodape();
  }

  // =========================================================================
  // Fim do cadastro de fotos
  // -------------------------------------------------------------------------

  // =========================================================================
  // Tela de dados básicos
  // -------------------------------------------------------------------------
  function DadosBasicos() {
    extract($GLOBALS);

    $O = "A";

    If ($O == "A") {
      $SQL = "select * from sbpi.Cliente where sq_cliente = " . $CL;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      foreach ($RS as $row) {
        $RS = $row;
        break;
      }
      $w_sq_cliente = $RS["sq_cliente"];
      $w_ds_senha_acesso = $RS["ds_senha_acesso"];
      $w_ds_email = $RS["ds_email"];
      $w_ds_apelido = $RS["ds_apelido"];
    }
    Cabecalho();
    ShowHTML("<HEAD>");
    ScriptOpen("Javascript");
    ValidateOpen("Validacao");
    If ($O == "A") {
      Validate("w_ds_senha_acesso", "Senha de acesso", "1", "1", "4", "14", "1", "1");
      Validate("w_ds_email", "e-Mail", "1", "", "6", "60", "1", "1");

    }
    ValidateClose();
    ScriptClose();
    ShowHTML("</HEAD>");
    BodyOpen("onLoad='document.Form.w_ds_senha_acesso.focus();'");
    ShowHTML('<B><FONT COLOR="#000000">Atualização de dados básicos</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">');
    AbreForm('Form', $w_pagina . "Grava", "POST", "return(Validacao(this));", null);
    ShowHTML(MontaFiltro("POST"));
    ShowHTML('<input type="hidden" name="O" value="' . $O . ' ">');
    ShowHTML('<input type="hidden" name="SG" value="BASICO">');

    ShowHTML('<tr bgcolor="#EFEFEF"><td align="center">');
    ShowHTML('    <table width="97%" border="0">');
    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Dados básicos da Unidade</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Os dados deste bloco são utilizados para acesso à página de atualização do site, para envio de mensagens ao responsável técnico e para pesquisa por apelido no mecanismo de busca.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top"><font size="1"><b><u>S</u>enha de acesso:</b><br><INPUT ACCESSKEY="S" "' . $w_Disabled . '" class="STI" type="text" name="w_ds_senha_acesso" size="14" maxlength="14" value="' . $w_ds_senha_acesso . '" title="OBRIGATÓRIO. Informe a senha desejada para acessar a tela de atualização dos dados do site."></td>');
    ShowHTML('      <tr><td valign="top"><font size="1"><b><u>e</u>-Mail:</b><br><INPUT ACCESSKEY="E" "' . $w_Disabled . '" class="STI" type="text" name="w_ds_email" size="60" maxlength="60" value="' . $w_ds_email . '" title="OPCIONAL. Informe o e-mail de sua unidade para contatos com a equipe técnica. Este e-mail não será divulgado no site."></td>');

    ShowHTML('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');

    // Verifica se poderá ser feito o envio da solicitação, a partir do resultado da validação
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
  // -------------------------

  function Administrativo() {
    extract($GLOBALS);
    $SQL = "select * from sbpi.Cliente_Admin where sq_cliente = " . $CL;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    foreach ($RS as $row) {
      $RS = $row;
      break;
    }
    if (count($RS) > 0) {
      $w_sq_cliente = f($RS, "sq_cliente");
      $w_limpeza = f($RS, "limpeza_terceirizada");
      $w_merenda = f($RS, "merenda_terceirizada");
      $w_banheiro = f($RS, "banheiro");
    }

    Cabecalho();
    ShowHTML('<HEAD>');
    ScriptOpen(Javascript);
    CheckBranco();
    FormataDATA();
    FormataCEP();
    ValidateOpen('Validacao');
    Validate("w_banheiro", "Quantidade de quadros magneticos", "1", "1", "1", "2", "", "0123456789");
    ShowHTML('  for (ind=1; ind < theForm["w_equipamento[]"].length; ind++) {');
    Validate('["w_equipamento[]"][ind]', "Quantidade de equipamentos", "1", "1", "1", "2", "", "0123456789");
    ShowHTML('  }');
    /*
      ShowHTML ('  for (i = 0; i < theForm["w_equipamento[]"].length; i++) {');
      ShowHTML ('      if (isNaN(parseInt(theForm["w_equipamento[]"][i].value))) {');
      ShowHTML ('         alert(\'Informe apenas números na quantidade dos equipamentos!\');');
      ShowHTML ('         theForm["w_equipamento[]"][i].focus();');
      ShowHTML ('         return false;');
      ShowHTML ('      }');
      ShowHTML ('  }');
    */
    ShowHTML('  theForm.Botao.disabled=true;');
    ValidateClose();
    ScriptClose();
    ShowHTML('</HEAD>');
    BodyOpen('onLoad=\'document.Form.focus();\'');
    ShowHTML('<B><FONT COLOR="#000000">Atualização de dados administrativos da unidade</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">');
    AbreForm("Form", $w_pagina . "Grava", "POST", "return(Validacao(this));", null);
    ShowHTML('<input type="hidden" name="SG" value="administrativo">');
    ShowHTML('<input type="hidden" name="O" value="A">');
    ShowHTML('<input type="hidden" name="w_equipamento[]" value="">');
    ShowHTML('<input type="hidden" name="w_codigo[]" value="">');
    ShowHTML(MontaFiltro("POST"));

    ShowHTML('<tr bgcolor="' . '#EFEFEF' . '"><td align="center">');
    ShowHTML('    <table width="97%" border="0">');
    ShowHTML('      <tr><td align="center"><font size=1 color="#BC3131"><b>ATENÇÃO: Após inserir ou alterar os dados abaixo, clique sobre o botão "<i>Gravar</i>", localizado no final desta página; caso contrário, seus dados serão perdidos.</td></tr>');
    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Serviços</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Informe os dados solicitados abaixo, sobre a terceirização de serviços e a quantidade de quadros magnéticos.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr>');
    MontaRadioNS("<b>O serviço de limpeza da escola é feita por empresa terceirizada?", $w_limpeza, "w_limpeza");
    ShowHTML('      <tr>');
    MontaRadioNS("<b>A escola oferece merenda?", $w_merenda, "w_merenda");
    ShowHTML('      <tr><td><font size="1"><b>Informe a quantidade total de QUADROS MAGNÉTICOS da escola:</b><br><INPUT ".$w_Disabled." class="STI" type="text" name="w_banheiro" size="2" maxlength="2" value="' . $w_banheiro . '"></td>');
    ShowHTML('      <tr><td align="center" colspan="3" height="1" bgcolor="#000000"></TD></TR>');
    ShowHTML('      <tr><td align="center" height="2" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td valign="top" align="center" bgcolor="#D0D0D0"><font size="1"><b>Equipamentos</td></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    ShowHTML('      <tr><td><font size=1>Verifique a lista abaixo e informe a quantidade de equipamentos existentes na escola.</font></td></tr>');
    ShowHTML('      <tr><td align="center" height="1" bgcolor="#000000"></td></tr>');
    $SQL = "select a.nome, b.sq_equipamento, b.nome nm_equipamento, coalesce(c.quantidade, 0) quantidade " . $crlf .
           "  from sbpi.Tipo_Equipamento                    a " . $crlf .
           "       inner      join sbpi.Equipamento         b on (a.sq_tipo_equipamento = b.sq_tipo_equipamento) " . $crlf .
           "       left outer join sbpi.Cliente_Equipamento c on (b.sq_equipamento      = c.sq_equipamento and " . $crlf .
           "                                                    c.sq_cliente          = " . $CL . " " . $crlf .
           "                                                   ) " . $crlf .
           "order by a.codigo, b.nome " . $crlf;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    if (count($RS) <= 0) {
      ShowHTML('      <tr><td><font size="1"><b>Não há equipamentos cadastrados.</td>');
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
        ShowHTML('        <td><font size=1><INPUT ' . $w_Disabled . ' class="STI" type="text" name="w_equipamento[]" size="3" maxlength="3" value="' . f($row, "quantidade") . '"> ');
        if (doubleval(f($row, "quantidade")) > 0) {
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
  // Fim da tela de dados administrativos
  // -------------------------------------------------------------------------

  function ShowLog() {
    extract($GLOBALS);

    $w_chave = $_REQUEST["p_cliente"];
    $w_troca = $_REQUEST["w_troca"];
    $sq_cliente_log = $_REQUEST['p_cliente_log'];

    //  ' Recupera o nome da escola
    $SQL = "select * from sbpi.Cliente where sq_cliente = " . $w_chave;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    $w_ds_cliente = $RS[0]["ds_cliente"];

    If ($O == "L") {
      //   ' Recupera todos os registros para a listagem
      $SQL = "select sq_cliente_log,abrangencia, to_char(data,'dd/mm/yyyy, hh24:mi:ss') as phpdt_data from sbpi.Cliente_Log where sq_cliente = " . $w_chave . " order by to_char(data,'yyyymmdd') desc, data desc";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }

    Cabecalho();
    ShowHTML("<HEAD>");
    ShowHTML("<TITLE>Registro de ocorrências no site da escola</TITLE>");
    If (strpos("IAEP", $O) !== false) {
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      If (strpos("IA", $O) !== false) {
        Validate("w_dt_ocorrencia", "data", "DATA", "1", "10", "10", "1", "1");
        Validate("w_ds_ocorrencia", "descriçao", "", "1", "2", "60", "1", "1");
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

    ShowHTML('<B><FONT COLOR="#000000">' . $w_ds_cliente . ' - Registro de ocorrências</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');

    If ($O == "L") {

      //    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('    <td align="right"><font size="1"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('   <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="#EFEFEF" align="center">');
      ShowHTML('          <td><font size="1"><b>Data</font></td>');
      ShowHTML('          <td><font size="1"><b>Hora</font></td>');
      ShowHTML('          <td><font size="1"><b>Ocorrência</font></td>');
      ShowHTML('          <td><font size="1"><b>Operações</font></td>');
      ShowHTML('        </tr>');

      If (count($RS) == 0) { //' Se não foram selecionados registros, exibe mensagem
        ShowHTML('<tr bgcolor="#EFEFEF"><td colspan=4 align="center"><font size="2"><b>Não foram encontrados registros.</b></td></tr>');
      } Else {

        //   rs.PageSize     = P4
        //   rs.AbsolutePage = P3
        $wAno = "";
        //      While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(P3)
        $RS1 = array_slice($RS, (($P3 -1) * $P4), $P4);
        foreach ($RS1 as $row) {

          If ($w_cor = "#EFEFEF" or $w_cor = "")
            $w_cor = "#FDFDFD";
          Else
            $w_cor = "#EFEFEF";

          If ($wAno != date('m/Y', $row["phpdt_data"])) {
            ShowHTML('<tr bgcolor="#C0C0C0" valign="top"><TD colspan=4 align="center"><font size=2><B>' . month($row["phpdt_data"]) . "/" . year($row["phpdt_data"]) . '</b></font></td></tr>');
            $wAno = date('m/Y', $row["phpdt_data"]);
          }
          ShowHTML('     <tr bgcolor="' . $w_cor . '" valign="top">');
          If ($w_data != FormataDataEdicao($row["phpdt_data"])) {
            ShowHTML('<td align="center"><font size="1">' . FormataDataEdicao($row["phpdt_data"]) . "</td>");
            $w_data = FormataDataEdicao($row["phpdt_data"]);
          } Else {
            ShowHTML('        <td align="center"></td>');
          }
          ShowHTML('        <td align="center"><font size="1">' . date('H:i:s', $row["phpdt_data"]) . "</td>");
          ShowHTML('        <td><font size="1">' . $row["abrangencia"] . "</td>");
          ShowHTML('        <td align="top" nowrap><font size="1">');
          ShowHTML('          <A class="HL" HREF="' . $w_pagina . $par . "&O=V&p_cliente_log=" . $row["sq_cliente_log"] . "&p_cliente=" . $w_chave . "&R=" . $w_pagina . $par . '">Detalhes</A>&nbsp');
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
      //    ' Recupera os dados da ocorrência selecionada
      $SQL = "select abrangencia,ip_origem ,data,sql as query from sbpi.Cliente_Log where sq_cliente_log =  " . $sq_cliente_log;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      $RS = $RS[0];
      ShowHTML('<tr bgcolor="#EFEFEF"><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td><font size="1">Data:<br><b>' . formatadataedicao($RS["data"]) . '</b></td>');
      ShowHTML('          <td><font size="1">IP de origem:<br><b>' . $RS["ip_origem"] . '</b></td>');
      ShowHTML('        <tr><td colspan=2><font size="1">&nbsp;');
      ShowHTML('        <tr><td colspan=2><font size="1">Tipo da ocorrencia: <b>');
      if ($RS['tipo'] == 0)
        ShowHTML('            Acesso');
      if ($RS['tipo'] == 1)
        ShowHTML('            Inclusão');
      if ($RS['tipo'] == 2)
        ShowHTML('            Alteração');
      if ($RS['tipo'] == 3)
        ShowHTML('            Exclusão');

      ShowHTML('            </td>');
      ShowHTML('        <tr><td colspan=2><font size="1">&nbsp;');
      ShowHTML('        <tr><td colspan=2><font size="1">Descrição:<br><b>' . $RS["abrangencia"] . '</b></td>');

      If (strlen($RS["query"]) > 0) {

        ShowHTML('        <tr><td colspan=2><font size="1">&nbsp;');
        ShowHTML('        <tr><td colspan=2><font size="1">Comando(s) executado(s):<br><b>' . str_replace($crlf, "<br>", str_replace(" ", "&nbsp;", $RS["query"])) . '</b></td>');
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

  function log_manut() {
    extract($GLOBALS);

    $w_chave = $CL;
    $w_troca = $_REQUEST["w_troca"];
    $sq_cliente_log = $_REQUEST['sq_cliente_log'];

    //  ' Recupera o nome da escola
    $SQL = "select * from sbpi.Cliente where sq_cliente = " . $w_chave;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    $w_ds_cliente = $RS[0]["ds_cliente"];

    If ($O == "L") {
      //   ' Recupera todos os registros para a listagem
      $SQL = "select sq_cliente_log,abrangencia, to_char(data,'dd/mm/yyyy, hh24:mi:ss') as phpdt_data from sbpi.Cliente_Log where sq_cliente = " . $w_chave . " order by to_char(data,'yyyymmdd') desc, data desc";
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    }

    Cabecalho();
    ShowHTML("<HEAD>");
    ShowHTML("<TITLE>Registro de ocorrências no site da escola</TITLE>");
    If (strpos("IAEP", $O) !== false) {
      ScriptOpen("JavaScript");
      CheckBranco();
      FormataData();
      ValidateOpen("Validacao");
      If (strpos("IA", $O) !== false) {
        Validate("w_dt_ocorrencia", "data", "DATA", "1", "10", "10", "1", "1");
        Validate("w_ds_ocorrencia", "descriçao", "", "1", "2", "60", "1", "1");
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

    ShowHTML('<B><FONT COLOR="#000000">' . $w_ds_cliente . ' - Registro de ocorrências</FONT></B>');
    ShowHTML('<HR>');
    ShowHTML('<div align=center><center>');
    ShowHTML('<table border="0" cellpadding="0" cellspacing="0" width="95%">');

    If ($O == "L") {

      //    ' Exibe a quantidade de registros apresentados na listagem e o cabeçalho da tabela de listagem
      ShowHTML('    <td align="right"><font size="1"><b>Registros existentes: ' . count($RS));
      ShowHTML('<tr><td align="center" colspan=3>');
      ShowHTML('   <TABLE WIDTH="100%" bgcolor="' . $conTableBgColor . '" BORDER="' . $conTableBorder . '" CELLSPACING="' . $conTableCellSpacing . '" CELLPADDING="' . $conTableCellPadding . '" BorderColorDark="' . $conTableBorderColorDark . '" BorderColorLight="' . $conTableBorderColorLight . '">');
      ShowHTML('        <tr bgcolor="#EFEFEF" align="center">');
      ShowHTML('          <td><font size="1"><b>Data</font></td>');
      ShowHTML('          <td><font size="1"><b>Hora</font></td>');
      ShowHTML('          <td><font size="1"><b>Ocorrência</font></td>');
      ShowHTML('          <td><font size="1"><b>Operações</font></td>');
      ShowHTML('        </tr>');

      If (count($RS) == 0) { //' Se não foram selecionados registros, exibe mensagem
        ShowHTML('<tr bgcolor="#EFEFEF"><td colspan=4 align="center"><font size="2"><b>Não foram encontrados registros.</b></td></tr>');
      } Else {

        //   rs.PageSize     = P4
        //   rs.AbsolutePage = P3
        $wAno = "";
        //      While Not RS.EOF and cDbl(RS.AbsolutePage) = cDbl(P3)
        $RS1 = array_slice($RS, (($P3 -1) * $P4), $P4);
        foreach ($RS1 as $row) {

          If ($w_cor = "#EFEFEF" or $w_cor = "")
            $w_cor = "#FDFDFD";
          Else
            $w_cor = "#EFEFEF";

          If ($wAno != date('m/Y', $row["phpdt_data"])) {
            ShowHTML('<tr bgcolor="#C0C0C0" valign="top"><TD colspan=4 align="center"><font size=2><B>' . month($row["phpdt_data"]) . "/" . year($row["phpdt_data"]) . '</b></font></td></tr>');
            $wAno = date('m/Y', $row["phpdt_data"]);
          }
          ShowHTML('     <tr bgcolor="' . $w_cor . '" valign="top">');
          If ($w_data != FormataDataEdicao($row["phpdt_data"])) {
            ShowHTML('<td align="center"><font size="1">' . FormataDataEdicao($row["phpdt_data"]) . "</td>");
            $w_data = FormataDataEdicao($row["phpdt_data"]);
          } Else {
            ShowHTML('        <td align="center"></td>');
          }
          ShowHTML('        <td align="center"><font size="1">' . date('H:i:s', $row["phpdt_data"]) . "</td>");
          ShowHTML('        <td><font size="1">' . $row["abrangencia"] . "</td>");
          ShowHTML('        <td align="top" nowrap><font size="1">');
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
      //    ' Recupera os dados da ocorrência selecionada
      $SQL = "select abrangencia,ip_origem ,data,sql as query from sbpi.Cliente_Log where sq_cliente_log =  " . $sq_cliente_log;
      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      $RS = $RS[0];
      ShowHTML('<tr bgcolor="#EFEFEF"><td align="center">');
      ShowHTML('    <table width="95%" border="0">');
      ShowHTML('        <tr valign="top">');
      ShowHTML('          <td><font size="1">Data:<br><b>' . formatadataedicao($RS["data"]) . '</b></td>');
      ShowHTML('          <td><font size="1">IP de origem:<br><b>' . $RS["ip_origem"] . '</b></td>');
      ShowHTML('        <tr><td colspan=2><font size="1">&nbsp;');
      ShowHTML('        <tr><td colspan=2><font size="1">Tipo da ocorrencia: <b>');
      if ($RS['tipo'] == 0)
        ShowHTML('            Acesso');
      if ($RS['tipo'] == 1)
        ShowHTML('            Inclusão');
      if ($RS['tipo'] == 2)
        ShowHTML('            Alteração');
      if ($RS['tipo'] == 3)
        ShowHTML('            Exclusão');

      ShowHTML('            </td>');
      ShowHTML('        <tr><td colspan=2><font size="1">&nbsp;');
      ShowHTML('        <tr><td colspan=2><font size="1">Descrição:<br><b>' . $RS["abrangencia"] . '</b></td>');

      If (strlen($RS["query"]) > 0) {

        ShowHTML('        <tr><td colspan=2><font size="1">&nbsp;');
        ShowHTML('        <tr><td colspan=2><font size="1">Comando(s) executado(s):<br><b>' . str_replace($crlf, "<br>", str_replace(" ", "&nbsp;", $RS["query"])) . '</b></td>');
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
  // Monta barra de navegação
  // -------------------------------------------------------------------------
  function barra($p_link, $p_PageCount, $p_AbsolutePage, $p_PageSize, $p_RecordCount) {
    extract($GLOBALS);
    ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT">');
    ShowHTML('  function pagina (pag) {');
    ShowHTML('    document.Barra.P3.value = pag;');
    ShowHTML('    document.Barra.submit();');
    ShowHTML('  }');
    ShowHTML('</SCRIPT>');
    ShowHTML('<left><FORM ACTION="' . $p_link . '" METHOD="POST" name="Barra">');
    ShowHTML('<input type="Hidden" name="P4" value="' . $p_PageSize . '">');
    ShowHTML('<input type="Hidden" name="p_modalidade" value="' . $p_modalidade . '">');
    ShowHTML('<input type="Hidden" name="w_ew" value="' . $w_ew . '">');
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

  // =========================================================================
  // Rotina principal
  // -------------------------------------------------------------------------
  function Main() {
    extract($GLOBALS);
    switch ($par) {
      Case 'MENU' :
        Menu();
        break;
      Case 'ADMINISTRATIVO' :
        Administrativo();
        break;
      Case 'VALIDA' :
        Valida();
        break;
      Case 'LOG_MANUT' :
        log_manut();
        break;
      Case 'FRAMES' :
        Frames();
        break;
      Case 'GETSITE' :
        GetSite();
        break;
      Case 'GRAVA' :
        Grava();
        break;
      Case 'NOTICIAS' :
        Noticias();
        break;
      Case 'ATUACAO' :
        Atuacao();
        break;
      Case 'ARQUIVOS' :
        Arquivos();
        break;
      Case 'ESCOLAS' :
        escolas();
        break;
      Case 'FOTOS' :
        fotos();
        break;
      Case 'CALENDARIO' :
        calend_rede();
        break;
      Case 'MENSAGENS' :
        Msgalunos();
        break;
      Case 'ADICIONAIS' :
        Adicionais();
        break;
      Case 'BASICO' :
        DadosBasicos();
        break;
      Case 'SHOWLOG' :
        ShowLog();
        break;
      Case 'SAIR' :
        Sair();
        break;
        /*
        
        */
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