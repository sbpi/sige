  <?php

  // Código do cliente
  $CL = * % *;

  // Define o banco de dados a ser utilizado
  $_SESSION['DBMS'] = 1;

  $w_dir_volta = '../../';
  $w_dir = 'modelos/mod13/';
  $w_pagina = 'default_cliente.php?';

  include_once ($w_dir_volta . 'constants.inc');
  include_once ($w_dir_volta . 'jscript.php');
  include_once ($w_dir_volta . 'funcoes.php');
  include_once ($w_dir_volta . 'classes/db/abreSessao.php');
  include_once ($w_dir_volta . 'classes/sp/db_exec.php');

  // =========================================================================
  //  /default.php
  // ------------------------------------------------------------------------
  // Nome     : Alexandre Vinhadelli Papadópolis
  // Descricao: Default para criação da página inicial da escola
  // Mail     : alex@sbpi.com.br
  // Criacao  : 21/01/2009, 11:11
  // Versao   : 1.0.0.0
  // Local    : Brasília - DF
  // SBPI Consultoria Ltda - Todos os direitos reservados
  // -------------------------------------------------------------------------
  //
  // Declaração de variáveis
  $RS = null;

  // Abre conexão com o banco de dados
  if (isset ($_SESSION['DBMS']))
    $dbms = abreSessao :: getInstanceOf($_SESSION['DBMS']);

  Main();

  // Fecha conexão com o banco de dados
  if (isset ($_SESSION['DBMS']))
    FechaSessao($dbms);

  exit;

  // =========================================================================
  // Rotina principal
  // -------------------------------------------------------------------------
  function Main() {
    extract($GLOBALS);

    // Verifica se há alguma escola particular no cadastro. Se não, marca para tratar apenas escolas públicas
    $SQL = "SELECT ds_cliente from sbpi.Cliente WHERE sq_cliente=" . $CL;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    foreach ($RS as $row) {
      $RS = $row;
      break;
    }

    ShowHTML('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">');
    ShowHTML('<html xmlns="http://www.w3.org/1999/xhtml">');
    ShowHTML('<head>');
    ShowHTML('  <title>' . f($RS, 'ds_cliente') . '</title>');
    ShowHTML('</head>');
    ShowHTML('<frameset framespacing="0" border="0" rows="*" frameborder="0">');
    ShowHTML('  <frame name="content" src="' . $conRootSIW . $w_dir . 'default.php?par=inicial&CL=' . $CL . '" target="_self" marginwidth="0" marginheight="0" frameborder="NO">');
    ShowHTML('  <noframes>');
    ShowHTML('  <body topmargin="0">');
    ShowHTML('  <p>Seu navegador não aceita quadros. Utilize o Netscape Navigator 4.x ou o Internet Explorer 5.x</p>');
    ShowHTML('  </body>');
    ShowHTML('  </noframes>');
    ShowHTML('</frameset>');
    ShowHTML('</html>');
  }
  ?>