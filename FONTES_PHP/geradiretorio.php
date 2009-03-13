  <?php

  // Define o banco de dados a ser utilizado
  $_SESSION['DBMS'] = 1;

  $w_dir_volta = '';
  $w_dir = '';
  $w_pagina = 'geradiretorio.php?';

  include_once ('constants.inc');
  include_once ('jscript.php');
  include_once ('funcoes.php');
  include_once ('classes/db/abreSessao.php');
  include_once ('classes/sp/db_exec.php');

  // =========================================================================
  //  /geradiretorio.php
  // ------------------------------------------------------------------------
  // Nome     : Alexandre Vinhadelli Papadópolis
  // Descricao: Gera diretórios e páginas iniciais para escolas
  // Mail     : alex@sbpi.com.br
  // Criacao  : 18/12/2008 13:02
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

  // Recupera o filtro selecionado
  $P3 = nvl($_REQUEST['P3'], 1);
  $P4 = nvl($_REQUEST['P4'], $conPageSize);

  $p_unidade = $_REQUEST['p_unidade'];
  $p_todos = $_REQUEST['p_todos'];

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

    if (nvl($p_unidade, '') != '') {
      $SQL = "SELECT a.sq_cliente, a.ds_username,b.sq_modelo from sbpi.Cliente a inner join sbpi.Cliente_Site b on (a.sq_cliente = b.sq_cliente) where upper(a.ds_username) = upper('" . $p_unidade . "')";
    } else {
      $SQL = "SELECT a.sq_cliente, a.ds_username,b.sq_modelo from sbpi.Cliente a inner join sbpi.Cliente_Site b on (a.sq_cliente = b.sq_cliente) order by ds_username";
    }
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    // Inicio ****************************************
    // Gera o arquivo default.php do cliente, a partir do default_cliente.php do modelo de site escolhido

    $w_cont = 0;
    $w_criado = 0;
    $w_alterado = 0;
    foreach ($RS as $row) {
      $w_diretorio = $conFilePhysical . strtolower(trim(f($row, 'ds_username')));
      $w_cliente = f($row, 'sq_cliente');
      $w_origem = str_replace('sedf/', '', $conFilePhysical) . 'modelos/mod' . strtolower(f($row, 'sq_modelo')) . '/default_cliente.php';
      $w_destino = $w_diretorio . '/default.php';

      // Verifica a necessidade de criação dos diretórios do cliente
      if ((!file_exists($w_diretorio)) || nvl($p_unidade, '') != '' || $p_todos == 'Sim') {
        if (!file_exists($w_diretorio)) {
          mkdir($w_diretorio);
          $w_criado++;
        }
        if (file_exists($w_destino)) {
          unlink($w_destino);
          $w_alterado++;
        }

        // Carrega o conteúdo do arquivo origem em uma variável local
        $conteudo = '';
        $handle = fopen($w_origem, 'r');
        while (!feof($handle)) {
          $conteudo .= fgets($handle, 4096);
        }
        fclose($handle);

        // Gera a página inicial do cliente, sustituindo *%* pelo código do cliente
        $handle = fopen($w_destino, 'w');
        fwrite($handle, str_replace('*%*', f($row, 'sq_cliente'), $conteudo));
        fclose($handle);
      }
      elseif (file_exists($w_diretorio)) {
        if (!file_exists($w_destino)) {
          // Carrega o conteúdo do arquivo origem em uma variável local
          $handle = fopen($w_origem, 'r');
          $conteudo = fread($handle, filesize($w_origem));
          fclose($handle);

          // Gera a página inicial do cliente, sustituindo *%* pelo código do cliente
          $handle = fopen($w_destino, 'w');
          fwrite($handle, str_replace('*%*', f($row, 'sq_cliente'), $conteudo));
          fclose($handle);
        }
      }
      $w_cont++;
    }
    // Fim ****************************************

    ShowHTML('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">');
    ShowHTML('<html xmlns="http://www.w3.org/1999/xhtml">');
    ShowHTML('<head>');
    ShowHTML('  <title>Geração de diretórios</title>');
    ShowHTML('</head>');
    ShowHTML('<body topmargin="0" leftmargin="0" bgcolor="#F0FFFF">');
    ShowHTML('<font size="5" color="#0000FF">Sucesso!</font><hr><font size="2">');
    ShowHTML('<p align="JUSTIFY">Foram processadas ' . $w_cont . ' escolas.</p>');
    ShowHTML('<p align="JUSTIFY"> ' . $w_criado . ' novas escolas foram inseridas.</p>');
    ShowHTML('<p align="JUSTIFY"> ' . $w_alterado . ' escolas existentes tiveram sua página inicial recriada (demais arquivos mantidos).</p>');
    ShowHTML('</body>');
    ShowHTML('</html>');
  }
  ?>