<?php
  setlocale(LC_ALL, 'pt_BR');
  mb_language('en');
  date_default_timezone_set('America/Sao_Paulo');

  //Este codigo deve ficar obrigatoriamente no inicio deste arquivo.
  $_REQUEST = array_antiInjection($_REQUEST);
  //var_dump($_REQUEST);

  function array_antiInjection($array) {
    $saida = array ();
    foreach ($array as $k => $v) {
      if (is_array($v)) {
        $v = array_antiInjection($v);
      } else {
        $v = antiInjection($v);
      }
      $saida[$k] = $v;
    }
    return $saida;
  }

  function antiInjection($string) {
    $string = trim($string);
    $string = preg_replace("/(procedure|alter|drop|union|create|from|alter table|select|insert|delete|update|where|drop table|show tables|#|\*|--|\\\\)/i", "", $string);
    //$string = strip_tags($string); // Retira Tags HTML
    $string = addslashes($string);
    $charS = array (
      "\'",
      '\"'
    );
    $charR = array (
      "\''",
      '\""'
    );
    $string = str_replace($charS, $charR, $string);
    return $string;
  }

  //$locale_info = localeconv();
  //echo "<pre>\n";
  //echo "--------------------------------------------\n";
  //echo "  Monetary information for current locale:  \n";
  //echo "--------------------------------------------\n\n";
  //echo "int_curr_symbol:   {$locale_info["int_curr_symbol"]}\n";
  //echo "currency_symbol:   {$locale_info["currency_symbol"]}\n";
  //echo "mon_decimal_point: {$locale_info["mon_decimal_point"]}\n";
  //echo "mon_thousands_sep: {$locale_info["mon_thousands_sep"]}\n";
  //echo "positive_sign:     {$locale_info["positive_sign"]}\n";
  //echo "negative_sign:     {$locale_info["negative_sign"]}\n";
  //echo "int_frac_digits:   {$locale_info["int_frac_digits"]}\n";
  //echo "frac_digits:       {$locale_info["frac_digits"]}\n";
  //echo "p_cs_precedes:     {$locale_info["p_cs_precedes"]}\n";
  //echo "p_sep_by_space:    {$locale_info["p_sep_by_space"]}\n";
  //echo "n_cs_precedes:     {$locale_info["n_cs_precedes"]}\n";
  //echo "n_sep_by_space:    {$locale_info["n_sep_by_space"]}\n";
  //echo "p_sign_posn:       {$locale_info["p_sign_posn"]}\n";
  //echo "n_sign_posn:       {$locale_info["n_sign_posn"]}\n";
  //echo "</pre>\n";

  // =========================================================================
  // Fun��o garante que as chaves de um array estar�o no caso indicado
  // -------------------------------------------------------------------------
  function array_key_case_change(& $array, $mode = 'CASE_LOWER') {
    // Make sure $array is really an array 
    if (!is_array($array))
      return false;

    $temp = $array;
    while (list ($key, $value) = each($temp)) {
      // First we unset the original so it's not lingering about 

      unset ($array[$key]);
      // Then modify the $key 
      switch ($mode) {
        case 'CASE_UPPER' :
          $value = array_change_key_case($value, CASE_UPPER);
          break;
        case 'CASE_LOWER' :
          $value = array_change_key_case($value, CASE_LOWER);
          break;
      }

      // Lastly read to the array using the new $key 
      $array[$key] = $value;
    }
    return true;
  }

  // =========================================================================
  // Fun��o para classifica��o de arrays
  // -------------------------------------------------------------------------
  function SortArray() {
    $arguments = array_change_key_case(func_get_args());
    $array = array_change_key_case($arguments[0]);
    $code = '';
    for ($c = 1; $c < count($arguments); $c += 2) {
      if (in_array($arguments[$c +1], array (
          "asc",
          "desc"
        ))) {
        $code .= 'if ($a["' . $arguments[$c] . '"] != $b["' . $arguments[$c] . '"]) {';
        if ($arguments[$c +1] == "asc") {
          $code .= 'return ($a["' . $arguments[$c] . '"] < $b["' . $arguments[$c] . '"] ? -1 : 1); }';
        } else {
          $code .= 'return ($a["' . $arguments[$c] . '"] < $b["' . $arguments[$c] . '"] ? 1 : -1); }';
        }
      }
    }
    $code .= 'return 0;';
    $compare = create_function('$a,$b', $code);
    usort($array, $compare);
    return $array;
  }

  // =========================================================================
  // Montagem do link para abrir o calend�rio
  // -------------------------------------------------------------------------
  function exibeCalendario($form, $campo) {
    extract($GLOBALS);
    return '   <a class="ss" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'calendario.php?nmForm=' . $form . '&nmCampo=' . $campo . '&vData=\'+document.' . $form . '.' . $campo . '.value,\'dp\',\'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,width=250,height=250,left=500,top=200\'); return false;" title="Visualizar calend�rio"><img src="images/icone/GotoTop.gif" border=0 align=top height=13 width=15></a>';
  }

  // =========================================================================
  // Retorna buffer de sa�da
  // -------------------------------------------------------------------------
  function callback($buffer) {
    return strip_tags($buffer, '<html><head><link><title><style><base><table><tr><td><div><b><p><hr><font>');
    ;
  }

  // =========================================================================
  // Abre e fecha a arvore.
  // -------------------------------------------------------------------------
  function colapsar($w_chave) {
    $saida = "&nbsp;";
    $saida .= "<img src='images/mais.jpg' style='cursor:pointer' alt='Expandir' onclick='colapsar(" . $w_chave . ",this)'/>";
    $saida .= "&nbsp;";
    return $saida;
  }

  // =========================================================================
  // Gera um link chamando o arquivo desejado
  // -------------------------------------------------------------------------
  function LinkArquivo($p_classe, $p_cliente, $p_arquivo, $p_target, $p_hint, $p_descricao, $p_retorno) {
    extract($GLOBALS);
    // Monta a chamada para a p�gina que retorna o arquivo
    $l_link = $conRootSIW . 'file.php?force=false&cliente=' . $p_cliente . '&id=' . $p_arquivo;

    If (strtoupper(Nvl($p_retorno, '')) == 'WORD') { // Se for gera�ao de Word, dispensa sess�o ativa
      // Altera a chamada padr�o, dispensando a sess�o
      $l_link = $conRootSIW . 'file.php?force=false&sessao=false&cliente=' & $p_cliente & '&id=' & $p_arquivo;
    }
    ElseIf (strtoupper(Nvl($p_retorno, '')) <> 'EMBED') { // Se n�o for objeto incorporado, monta tag anchor
      // Trata a possibilidade da chamada ter passado classe, target e hint
      If (Nvl($p_classe, '') > '')
        $l_classe = ' class="' . $p_classe . '" ';
      Else
        $l_classe = '';
      If (Nvl($p_target, '') > '')
        $l_target = ' target="' . $p_target . '" ';
      Else
        $l_target = '';
      If (Nvl($p_hint, '') > '')
        $l_hint = ' title="' . $p_hint . '" ';
      Else
        $l_hint = '';

      // Montagem da tag anchor
      $l_link = '<a' . $l_classe . 'href="' . str_replace('force=false', 'force=true', $l_link) . '"' . $l_target . $l_hint . '>' . $p_descricao . '</a>';
    }

    // Retorno ao chamador
    return $l_link;
  }

  // =========================================================================
  // Grava��o da imagem da solicita��o no log
  // -------------------------------------------------------------------------
  function CriaBaseLine($l_chave, $l_html, $l_nome, $l_tramite) {
    extract($GLOBALS);
    include_once ($w_dir_volta . 'classes/sp/dml_putBaseLine.php');
    $l_caminho = $conFilePhysical . $w_cliente . '/';
    $l_nome_arq = $l_chave . '_' . time() . '.html';
    $l_arquivo = $l_caminho . $l_nome_arq;
    // Abre o arquivo de log
    $l_arq = @ fopen($l_arquivo, 'w');
    $l_html = str_replace("display:none", "", $l_html);
    $l_html = str_replace("mais.jpg", "menos.jpg", $l_html);
    fwrite($l_arq, '<HTML>');
    fwrite($l_arq, '<HEAD>');
    fwrite($l_arq, '<TITLE>Visualiza��o de ' . $l_nome . '</TITLE>');
    fwrite($l_arq, '</HEAD>');
    fwrite($l_arq, '<BASE HREF="' . $conRootSIW . '">');
    fwrite($l_arq, '<link rel="stylesheet" type="text/css" href="' . $conRootSIW . 'classes/menu/xPandMenu.css">');
    fwrite($l_arq, '<BODY>');
    fwrite($l_arq, '<div align="center">');
    fwrite($l_arq, '<table width="95%" border="0" cellspacing="3">');
    fwrite($l_arq, '<tr><td colspan="2">');
    fwrite($l_arq, $l_html);
    fwrite($l_arq, '</table>');
    fwrite($l_arq, '</div>');
    fwrite($l_arq, '</BODY>');
    fwrite($l_arq, '</HTML>');
    @ fclose($l_arq);
    dml_putBaseLine :: getInstanceOf($dbms, $w_cliente, $l_chave, $w_usuario, $l_tramite, $l_nome_arq, filesize($l_arquivo), 'text/html', $l_nome_arq);
  }

  // =========================================================================
  // Gera um link para JavaScript, em fun��o do navegador
  // -------------------------------------------------------------------------
  function montaURL_JS($p_dir, $p_link) {
    extract($GLOBALS);
    $l_link = str_replace($conRootSIW, '', $p_link);
    if (nvl($p_dir, '') != '')
      $l_link = str_replace($p_dir, '', $l_link);
    return $conRootSIW . $p_dir . $l_link;
  }

  // =========================================================================
  // Gera c�digo de barras para o valor informado
  // -------------------------------------------------------------------------
  function geraCB($l_valor, $l_tamanho = 6, $l_fator = 0.6, $l_formato = 'C39') {
    extract($GLOBALS);
    if (strtoUpper($l_formato) == 'C39') {
      include_once ($w_dir_volta . 'classes/graph_barcode/C39Barcode_class.php');
      $cb = new c39Barcode('cb1', $l_valor);
    } else {
      include_once ($w_dir_volta . '/classes/graph_barcode/I25Barcode_class.php');
      $cb = new I25Barcode('cb1', $l_valor);
    }
    $cb->setFactor($l_fator); // Fator de aumento. Quanto maior, mais larga � cada barra do c�digo.
    return '<font size=' . intVal($l_tamanho) . '">' . $cb->getBarcode() . '</font>';
  }

  // =========================================================================
  // Fun��o para codificar strings em base64 com final "=="
  // -------------------------------------------------------------------------
  function base64encodeIdentificada($string) {
    return base64_encode($string) . "=|=";
  }

  // =========================================================================
  // Declara��o inicial para p�ginas OLE com Word
  // -------------------------------------------------------------------------
  function headerWord($p_orientation = 'LANDSCAPE') {
    extract($GLOBALS);
    header("Cache-Control: no-cache, must-revalidate", false);
    header('Content-type: application/msword', false);
    header('Content-Disposition: attachment; filename=arquivo.doc');
    ShowHTML('<html xmlns:o="urn:schemas-microsoft-com:office:office" ');
    ShowHTML('xmlns:w="urn:schemas-microsoft-com:office:word" ');
    ShowHTML('xmlns="http://www.w3.org/TR/REC-html40"> ');
    ShowHTML('<head> ');
    ShowHTML('<meta http-equiv=Content-Type content="text/html; charset=windows-1252"> ');
    ShowHTML('<meta name=ProgId content=Word.Document> ');
    ShowHTML('<!--[if gte mso 9]><xml> ');
    ShowHTML(' <w:WordDocument> ');
    ShowHTML('  <w:View>Print</w:View> ');
    ShowHTML('  <w:Zoom>BestFit</w:Zoom> ');
    ShowHTML('  <w:SpellingState>Clean</w:SpellingState> ');
    ShowHTML('  <w:GrammarState>Clean</w:GrammarState> ');
    ShowHTML('  <w:HyphenationZone>21</w:HyphenationZone> ');
    ShowHTML('  <w:IgnoreMixedContent>true</w:IgnoreMixedContent>');
    ShowHTML('  <w:Compatibility> ');
    ShowHTML('   <w:BreakWrappedTables/> ');
    ShowHTML('   <w:SnapToGridInCell/> ');
    ShowHTML('   <w:WrapTextWithPunct/> ');
    ShowHTML('   <w:UseAsianBreakRules/> ');
    ShowHTML('  </w:Compatibility> ');
    ShowHTML('  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel> ');
    //ShowHTML('  <w:DocumentProtection>forms</w:DocumentProtection> ');
    ShowHTML(' </w:WordDocument> ');
    ShowHTML('</xml><![endif]--> ');
    ShowHTML('<style> ');
    ShowHTML('<!-- ');
    ShowHTML(' /* Style Definitions */ ');
    ShowHTML('@page Section1 ');
    if (strtoupper(Nvl($p_orientation, 'LANDSCAPE')) == 'PORTRAIT') {
      ShowHTML('    {size:8.5in 11.0in; ');
      ShowHTML('    mso-page-orientation:portrait; ');
      ShowHTML('    margin:2.0cm 2.0cm 2.0cm 2.0cm; ');
      ShowHTML('    mso-header-margin:35.4pt; ');
      ShowHTML('    mso-footer-margin:35.4pt; ');
      ShowHTML('    mso-paper-source:0;} ');
    } else {
      ShowHTML('    {size:11.0in 8.5in; ');
      ShowHTML('    mso-page-orientation:landscape; ');
      ShowHTML('    margin:60.85pt 1.0cm 60.85pt 2.0cm; ');
      ShowHTML('    mso-header-margin:35.4pt; ');
      ShowHTML('    mso-footer-margin:35.4pt; ');
      ShowHTML('    mso-paper-source:0;} ');
    }
    ShowHTML('div.Section1 ');
    ShowHTML('    {page:Section1;} ');
    ShowHTML('--> ');
    ShowHTML('</style> ');
    ShowHTML('</head> ');
    BodyOpen('onLoad=this.focus();');
    ShowHTML('<div class=Section1> ');
    ShowHTML('<link rel="stylesheet" type="text/css" href="' . $conRootSIW . 'classes/menu/xPandMenu.css">');
    ShowHTML('<BASE HREF="' . $conRootSIW . '">');
  }

  // =========================================================================
  // Monta (+) e (-) na arvore de projeto
  // -------------------------------------------------------------------------
  function montaArvore($string) {
    $string = str_replace(".", "-", $string);
    $img = "<img src='images/mais.jpg' alt='Expandir' onclick='abreFecha(\"$string\")' id='img-$string' style='cursor:pointer'/>";
    return $img;
  }

  // =========================================================================
  // Declara��o inicial para p�ginas OLE com PDF
  // -------------------------------------------------------------------------

  function headerPdf($titulo, $pag = null) {
    extract($GLOBALS);
    header("Cache-Control: no-cache, must-revalidate", false);
    header("Expires: Mon, 26 Jul 2008 05:00:00 GMT");
    ob_end_clean();
    ob_start();
    Cabecalho();
    ShowHTML('<HEAD>');
    ShowHTML('<TITLE>' . $titulo . '</TITLE>');
    ShowHTML('<link rel="stylesheet" type="text/css" href="' . $conRootSIW . '/classes/menu/xPandMenu.css">');
    ShowHTML('</HEAD>');
    ShowHTML('<BASE HREF="' . $conRootSIW . '">');
    CabecalhoWord($w_cliente, $titulo, $pag);
    BodyOpenMail(null);
  }

  // =========================================================================
  // Declara��o inicial para p�ginas OLE com Word
  // -------------------------------------------------------------------------
  function headerExcel($p_orientation = 'LANDSCAPE') {
    extract($GLOBALS);
    header('Content-type: application/excel', false);
    header('Content-Disposition: attachment; filename=arquivo.xls');
    ShowHTML('<html xmlns:o="urn:schemas-microsoft-com:office:office" ');
    ShowHTML('xmlns:w="urn:schemas-microsoft-com:office:excel" ');
    ShowHTML('xmlns="http://www.w3.org/TR/REC-html40"> ');
    ShowHTML('<head> ');
    ShowHTML('<meta http-equiv=Content-Type content="text/html; charset=windows-1252"> ');
    ShowHTML('<meta name=ProgId content=Pdf.Document> ');
    ShowHTML('<!--[if gte mso 9]><xml> ');
    ShowHTML(' <w:ExcelDocument> ');
    ShowHTML('  <w:View>Print</w:View> ');
    ShowHTML('  <w:Zoom>BestFit</w:Zoom> ');
    ShowHTML('  <w:SpellingState>Clean</w:SpellingState> ');
    ShowHTML('  <w:GrammarState>Clean</w:GrammarState> ');
    ShowHTML('  <w:HyphenationZone>21</w:HyphenationZone> ');
    ShowHTML('  <w:Compatibility> ');
    ShowHTML('   <w:BreakWrappedTables/> ');
    ShowHTML('   <w:SnapToGridInCell/> ');
    ShowHTML('   <w:WrapTextWithPunct/> ');
    ShowHTML('   <w:UseAsianBreakRules/> ');
    ShowHTML('  </w:Compatibility> ');
    ShowHTML('  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel> ');
    //ShowHTML('  <w:DocumentProtection>forms</w:DocumentProtection> ');
    ShowHTML(' </w:ExcelDocument> ');
    ShowHTML('</xml><![endif]--> ');
    ShowHTML('<style> ');
    ShowHTML('<!-- ');
    ShowHTML(' /* Style Definitions */ ');
    ShowHTML('@page Section1 ');
    if (strtoupper(Nvl($p_orientation, 'LANDSCAPE')) == 'PORTRAIT') {
      ShowHTML('    {size:8.5in 11.0in; ');
      ShowHTML('    mso-page-orientation:portrait; ');
      ShowHTML('    margin:2.0cm 2.0cm 2.0cm 2.0cm; ');
      ShowHTML('    mso-header-margin:35.4pt; ');
      ShowHTML('    mso-footer-margin:35.4pt; ');
      ShowHTML('    mso-paper-source:0;} ');
    } else {
      ShowHTML('    {size:11.0in 8.5in; ');
      ShowHTML('    mso-page-orientation:landscape; ');
      ShowHTML('    margin:60.85pt 1.0cm 60.85pt 2.0cm; ');
      ShowHTML('    mso-header-margin:35.4pt; ');
      ShowHTML('    mso-footer-margin:35.4pt; ');
      ShowHTML('    mso-paper-source:0;} ');
    }
    ShowHTML('div.Section1 ');
    ShowHTML('    {page:Section1;} ');
    ShowHTML('--> ');
    ShowHTML('</style> ');
    ShowHTML('</head> ');
    BodyOpen('onLoad=this.focus();');
    ShowHTML('<div class=Section1> ');
    ShowHTML('<link rel="stylesheet" type="text/css" href="' . $conRootSIW . 'classes/menu/xPandMenu.css">');
    ShowHTML('<BASE HREF="' . $conRootSIW . '">');
  }

  // =========================================================================
  // Montagem do cabe�alho de documentos Word
  // -------------------------------------------------------------------------
  function CabecalhoWord($p_cliente, $p_titulo, $p_pagina) {
    extract($GLOBALS);
    include_once ($w_dir_volta . 'classes/sp/db_getCustomerData.php');
    $l_RS = db_getCustomerData :: getInstanceOf($dbms, $p_cliente);
    ShowHTML('<TABLE WIDTH="100%" BORDER=0>');
    ShowHTML('  <TR>');
    if (nvl($p_pagina, 0) > 0)
      $l_colspan = 4;
    else
      $l_colspan = 3;
    ShowHTML('    <TD ROWSPAN=' . $l_colspan . '><IMG ALIGN="LEFT" SRC="' . $conFileVirtual . $w_cliente . '/img/' . f($l_RS, 'LOGO') . '">');
    ShowHTML('    <TD ALIGN="RIGHT"><B><FONT SIZE=3 COLOR="#000000">' . $p_titulo . '</FONT>');
    ShowHTML('  </TR>');
    ShowHTML('  <TR><TD ALIGN="RIGHT"><B><FONT COLOR="#000000">' . DataHora() . '</B></TD></TR>');
    ShowHTML('  <TR><TD ALIGN="RIGHT"><B><FONT COLOR="#000000">Usu�rio: ' . $_SESSION['NOME_RESUMIDO'] . '</B></TD></TR>');
    if (nvl($p_pagina, 0) > 0)
      ShowHTML('  <TR><TD ALIGN="RIGHT"><B><FONT SIZE=2 COLOR="#000000">P�gina: ' . $p_pagina . '</B></TD></TR>');
    ShowHTML('  <TR><TD colspan=2><HR></td></tr>');
    ShowHTML('</TABLE>');
  }

  // =========================================================================
  // Montagem de link para ordena��o, usada nos t�tulos de colunas
  // -------------------------------------------------------------------------
  function LinkOrdena($p_label, $p_campo) {
    extract($GLOBALS);
    foreach ($_POST as $chv => $vlr) {
      if (nvl($vlr, '') > '' && (strtoupper(substr($chv, 0, 2)) == "W_" || strtoupper(substr($chv, 0, 2)) == "P_")) {
        if (strtoupper($chv) == "P_ORDENA") {
          $l_ordena = strtoupper($vlr);
        } else {
          if (is_array($vlr))
            $l_string .= '&' . $chv . "=" . explodeArray($vlr);
          else
            $l_string .= '&' . $chv . "=" . $vlr;
        }
      }
    }
    foreach ($_GET as $chv => $vlr) {
      if (nvl($vlr, '') > '' && (strtoupper(substr($chv, 0, 2)) == "W_" || strtoupper(substr($chv, 0, 2)) == "P_")) {
        if (strtoupper($chv) == "P_ORDENA") {
          $l_ordena = strtoupper($vlr);
        } else {
          if (is_array($vlr))
            $l_string .= '&' . $chv . "=" . explodeArray($vlr);
          else
            $l_string .= '&' . $chv . "=" . $vlr;
        }
      }
    }
    if (strtoupper($p_campo) == str_replace(' DESC', '', str_replace(' ASC', '', strtoupper($l_ordena)))) {
      if (strpos(strtoupper($l_ordena), ' DESC') !== false) {
        $l_string .= '&p_ordena=' . $p_campo . ' asc&';
        $l_img = '&nbsp;<img src="images/down.gif" width=8 height=8 border=0 align="absmiddle">';
      } else {
        $l_string .= '&p_ordena=' . $p_campo . ' desc&';
        $l_img = '&nbsp;<img src="images/up.gif" width=8 height=8 border=0 align="absmiddle">';
      }
    } else {
      $l_string .= '&p_ordena=' . $p_campo . ' asc&';
    }
    return '<a class="ss" href="' . $w_dir . $w_pagina . $par . '&R=' . $R . '&O=' . $O . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=1' . '&P4=' . $P4 . '&TP=' . $TP . '&SG=' . $SG . $l_string . '" title="Ordena a listagem por esta coluna.">' . $p_label . '</a>' . $l_img;
  }

  // =========================================================================
  // Montagem do cabe�alho de relat�rios
  // -------------------------------------------------------------------------
  function CabecalhoRelatorio($p_cliente, $p_titulo, $p_rowspan = 2, $l_chave = null) {
    extract($GLOBALS);

    include_once ($w_dir_volta . 'classes/sp/db_getCustomerData.php');
    $RS_Logo = db_getCustomerData :: getInstanceOf($dbms, $w_cliente);
    if (f($RS_Logo, 'logo') > '') {
      $p_logo = 'img/logo' . substr(f($RS_Logo, 'logo'), (strpos(f($RS_Logo, 'logo'), '.') ? strpos(f($RS_Logo, 'logo'), '.') + 1 : 0) - 1, 30);
    }
    ShowHTML('<TABLE WIDTH="100%" BORDER=0><TR><TD ROWSPAN=' . $p_rowspan . '><IMG ALIGN="LEFT" SRC="' . LinkArquivo(null, $p_cliente, $p_logo, null, null, null, 'EMBED') . '"><TD ALIGN="RIGHT"><B><FONT SIZE=4 COLOR="#000000">' . $p_titulo);
    ShowHTML('</FONT></TR><TR><TD ALIGN="RIGHT"><B><FONT COLOR="#000000">' . DataHora() . '</B></TD></TR>');
    ShowHTML('<TR><TD ALIGN="RIGHT"><B><font COLOR="#000000">Usu�rio: ' . $_SESSION['NOME_RESUMIDO'] . '</B></TD></TR>');
    if (($p_tipo != 'WORD' && $w_tipo != 'WORD')) {
      ShowHTML('<TR><TD ALIGN="RIGHT">');
      if (nvl($l_chave, '') > '') {
        if (RetornaGestor($l_chave, $w_usuario) == 'S')
          ShowHTML('&nbsp;<A  class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . montaURL_JS(null, $conRootSIW . 'seguranca.php?par=TelaAcessoUsuarios&w_chave=' . $l_chave . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=1&TP=' . $TP . '&SG=') . '\',\'Usuarios\',\'width=780,height=550,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;"><IMG border=0 ALIGN="CENTER" TITLE="Usu�rios com acesso a este documento" SRC="images/Folder/User.gif"></a>');
      }
      ShowHTML('&nbsp;<IMG ALIGN="CENTER" TITLE="Imprimir" SRC="images/impressora.gif" onClick="window.print();">');
      $word_par = montaurl_js($w_dir, $conRootSIW . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=' . $O . '&w_chave=' . $l_chave . '&w_acordo=' . $l_chave . '&p_plano=' . $l_chave . '&w_ano=' . $w_ano . '&p_tipo=WORD&w_tipo=WORD&w_tipo_rel=WORD&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=1&TP=' . $TP . '&SG=' . $SG . MontaFiltro('GET'));
      //ShowHTML('&nbsp;<a href="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O='.$O.'&w_chave='.$l_chave.'&w_acordo='.$l_chave.'&p_plano='.$l_chave.'&w_sq_pessoa='.$l_chave.'&w_ano='.$w_ano.'&p_tipo=WORD&w_tipo=WORD&w_tipo_rel=WORD&P1='.$P1.'&P2='.$P2.'&P3='.$P3.'&P4=1&TP='.$TP.'&SG='.$SG.MontaFiltro('GET').'"><IMG border=0 ALIGN="CENTER" TITLE="Gerar word" SRC="images/word.jpg"></a>');
      ShowHtml('<img  style="cursor:pointer" onclick=\' document.temp.opcao.value="W"; displayMessage(310,140,"funcoes/orientacao.php");\' border=0 ALIGN="CENTER" TITLE="Gerar Word" SRC="images/word.jpg" />');

      //ShowHTML('&nbsp;<a href="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O=L&w_chave='.$l_chave.'&w_acordo='.$l_chave.'&p_plano='.$l_chave.'&w_ano='.$w_ano.'&p_tipo=EXCEL&w_tipo=EXCEL&w_tipo_rel=EXCEL&P1='.$P1.'&P2='.$P2.'&P3='.$P3.'&P4=1&TP='.$TP.'&SG='.$SG.MontaFiltro('GET').'"><IMG border=0 ALIGN="CENTER" TITLE="Gerar Excel" SRC="images/excel.jpg"></a>');
      // ShowHTML('&nbsp;<a href="'.$w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O='.$O.'&w_chave='.$l_chave.'&w_acordo='.$l_chave.'&p_plano='.$l_chave.'&w_ano='.$w_ano.'&p_tipo=PDF&w_tipo=PDF&w_tipo_rel=WORD&P1='.$P1.'&P2='.$P2.'&P3='.$P3.'&P4=1&TP='.$TP.'&SG='.$SG.MontaFiltro('GET').'" target="_blank"><IMG border=0 ALIGN="CENTER" TITLE="Gerar PDF" SRC="images/pdf.png"></a>');
      $pdf_par = montaurl_js($w_dir, $conRootSIW . $w_dir . $w_pagina . $par . '&R=' . $w_pagina . $par . '&O=' . $O . '&w_chave=' . $l_chave . '&w_acordo=' . $l_chave . '&p_plano=' . $l_chave . '&w_ano=' . $w_ano . '&p_tipo=PDF&w_tipo=PDF&w_tipo_rel=WORD&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=1&TP=' . $TP . '&SG=' . $SG . MontaFiltro('GET'));
      // echo $parametros;
      ShowHtml('<img  style="cursor:pointer" onclick=\' document.temp.opcao.value="P"; displayMessage(310,140,"funcoes/orientacao.php");\' border=0 ALIGN="CENTER" TITLE="Gerar PDF" SRC="images/pdf.png" />');
      ShowHTML('</TD></TR>');
    }
    ShowHTML('</TABLE>');
    ShowHTML('<form name="temp">');
    ShowHTML('<input type="hidden" name="word" id="word" value="' . $word_par . '">');
    ShowHTML('<input type="hidden" name="pdf" id="pdf" value="' . $pdf_par . '">');
    ShowHTML('<input type="hidden" name="opcao" id="opcao" value="">');
    ShowHTML('</form>');
    flush();
  }

  // =========================================================================
  // Montagem da barra de navega��o de recordsets
  // -------------------------------------------------------------------------
  function MontaBarra($p_link, $p_PageCount, $p_AbsolutePage, $p_PageSize, $p_RecordCount) {
    extract($GLOBALS);
    ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT">');
    ShowHTML('  function pagina (pag) {');
    ShowHTML('    document.Barra.P3.value = pag;');
    ShowHTML('    document.Barra.submit();');
    ShowHTML('  }');
    ShowHTML('</SCRIPT>');
    ShowHTML('<FORM ACTION="' . $p_link . '" METHOD="POST" name="Barra">');
    ShowHTML('<input type="Hidden" name="P4" value="' . $p_PageSize . '">');
    ShowHTML('<input type="Hidden" name="P4" value="' . $p_PageSize . '">');
    ShowHTML('<input type="Hidden" name="w_ew" value="' . $w_ew . '">');
    ShowHTML(MontaFiltro('POST'));
    if ($p_PageSize < $p_RecordCount) {
      if ($p_PageCount == $p_AbsolutePage) {
        ShowHTML('<span class="BTM"><br>' . ($p_RecordCount - (($p_PageCount -1) * $p_PageSize)) . ' linhas apresentadas de ' . $p_RecordCount . ' linhas');
      } else {
        ShowHTML('<span class="BTM"><br>' . $p_PageSize . ' linhas apresentadas de ' . $p_RecordCount . ' linhas');
      }
      ShowHTML('<br>na p�gina ' . $p_AbsolutePage . ' de ' . $p_PageCount . ' p�ginas');
      if ($p_AbsolutePage > 1) {
        ShowHTML('<br>[<A class="ss" TITLE="Primeira p�gina" HREF="javascript:pagina(1)" onMouseOver="window.status=\'Primeira (1/' . $p_PageCount . ')\'; return true" onMouseOut="window.status=\'\'; return true;">Primeira</A>]&nbsp;');
        ShowHTML('[<A class="ss" TITLE="P�gina anterior" HREF="javascript:pagina(' . ($p_AbsolutePage -1) . ')" onMouseOver="window.status=\'Anterior (' . ($p_AbsolutePage -1) . '/' . $p_PageCount . ')\'; return true;" onMouseOut="window.status=\'\'; return true;">Anterior</A>]&nbsp;');
      } else {
        ShowHTML('<br>[Primeira]&nbsp;');
        ShowHTML('[Anterior]&nbsp;');
      }
      if ($p_PageCount == $p_AbsolutePage) {
        ShowHTML('[Pr�xima]&nbsp;');
        ShowHTML('[�ltima]');
      } else {
        ShowHTML('[<A class="ss" TITLE="P�gina seguinte" HREF="javascript:pagina(' . ($p_AbsolutePage +1) . ')"  onMouseOver="window.status=\'Pr�xima (' . ($p_AbsolutePage +1) . '/' . $p_PageCount . ')\'; return true" onMouseOut="window.status=\'\'; return true">Pr�xima</A>]&nbsp;');
        ShowHTML('[<A class="ss" TITLE="�ltima p�gina" HREF="javascript:pagina(' . $p_PageCount . ')"  onMouseOver="window.status=\'�ltima (' . $p_PageCount . '/' . $p_PageCount . ')\'; return true" onMouseOut="window.status=\'\'; return true">�ltima</A>]');
      }
      ShowHTML('</span>');
    }
    ShowHtml('</FORM>');
  }

  // =========================================================================
  // Retorna o n�vel de acesso que o usu�rio tem � solicita��o informada
  // -------------------------------------------------------------------------
  function SolicAcesso($p_solicitacao, $p_usuario) {
    extract($GLOBALS);
    $l_acesso = db_getSolicAcesso :: getInstanceOf($dbms, $p_solicitacao, $p_usuario);
    return $l_acesso;
  }

  // =========================================================================
  // Fun��o que retorna S/N indicando se h� expediente na data informada
  // -------------------------------------------------------------------------
  function RetornaExpediente($p_data, $p_cliente, $p_pais, $p_uf, $p_cidade) {
    extract($GLOBALS);
    include_once ($w_dir_volta . 'classes/sp/db_VerificaDataEspecial.php');
    $l_expediente = db_VerificaDataEspecial :: getInstanceOf($dbms, $p_data, $p_cliente, $p_pais, $p_uf, $p_cidade);
    return $l_expediente;
  }

  // =========================================================================
  // Retorna o tipo de recurso a partir do c�digo
  // -------------------------------------------------------------------------
  function RetornaTipoRecurso($l_chave) {
    extract($GLOBALS);

    switch ($l_chave) {
      case 0 :
        return 'Financeiro';
        break;
      case 1 :
        return 'Humano';
        break;
      case 2 :
        return 'Material';
        break;
      case 3 :
        return 'Metodol�gico';
        break;
      default :
        return 'Erro';
        break;
    }
  }
  // =========================================================================
  // Retorna o nome da base geogr�fica a partir do c�digo
  // -------------------------------------------------------------------------
  function retornaBaseGeografica($l_chave) {
    extract($GLOBALS);

    switch ($l_chave) {
      case 1 :
        return 'Nacional';
        break;
      case 2 :
        return 'Regional';
        break;
      case 3 :
        return 'Estadual';
        break;
      case 4 :
        return 'Municipal';
        break;
      case 5 :
        return 'Organizacional';
        break;
      default :
        return 'Erro';
        break;
    }
  }
  // =========================================================================
  // Fun�ao para retornar o tipo da data
  // -------------------------------------------------------------------------
  function RetornaTipoData($l_chave) {
    extract($GLOBALS);
    switch ($l_chave) {
      case 'I' :
        return 'Invari�vel';
        break;
      case 'E' :
        return 'Espec�fica';
        break;
      case 'S' :
        return 'Segunda Carnaval';
        break;
      case 'C' :
        return 'Ter�a Carnaval';
        break;
      case 'Q' :
        return 'Quarta Cinzas';
        break;
      case 'P' :
        return 'Sexta Santa';
        break;
      case 'D' :
        return 'Domingo P�scoa';
        break;
      case 'H' :
        return 'Corpus Christi';
        break;
      default :
        return 'Erro';
        break;
    }
  }
  // =========================================================================
  // Fun�ao para retornar o expediente da data
  // -------------------------------------------------------------------------
  function RetornaExpedienteData($l_chave) {
    extract($GLOBALS);
    switch ($l_chave) {
      case 'S' :
        return 'Sim';
        break;
      case 'N' :
        return 'N�o';
        break;
      case 'M' :
        return 'Somente manh�';
        break;
      case 'T' :
        return 'Somente tarde';
        break;
      default :
        return 'Sim';
        break;
    }
  }
  // =========================================================================
  // Retorna uma parte qualquer de uma linha delimitada
  // -------------------------------------------------------------------------
  // function Piece($p_line,$p_delimiter,$p_separator,$p_position) {
  // $l_array = explode($p_separator,$p_line);
  // return $l_array[($p_position-1)];
  // }

  // =========================================================================
  // Montagem da URL com os par�metros de filtragem
  // -------------------------------------------------------------------------
  function MontaFiltro($p_method) {
    extract($GLOBALS);
    if (strtoupper($p_method) == 'GET' || strtoupper($p_method) == 'POST') {
      $l_string = '';
      foreach ($_POST as $l_Item => $l_valor) {
        if (substr($l_Item, 0, 2) == 'p_' && $l_valor > '') {
          if (strtoupper($p_method) == 'GET') {
            if (is_array($_POST[$l_Item])) {
              $l_string .= '&' . $l_Item . '=' . explodeArray($_POST[$l_Item]);
            } else {
              $l_string .= '&' . $l_Item . '=' . $l_valor;
            }
          }
          elseif (strtoupper($p_method) == 'POST') {
            if (is_array($_POST[$l_Item])) {
              $l_string .= '<INPUT TYPE="HIDDEN" NAME="' . $l_Item . '" VALUE="' . explodeArray($_POST[$l_Item]) . '">';
            } else {
              $l_string .= '<INPUT TYPE="HIDDEN" NAME="' . $l_Item . '" VALUE="' . $l_valor . '">';
            }
          }
        }
      }
      foreach ($_GET as $l_Item => $l_valor) {
        if (substr($l_Item, 0, 2) == 'p_' && $l_valor > '') {
          if (strtoupper($p_method) == 'GET') {
            if (is_array($_GET[$l_Item])) {
              $l_string .= '&' . $l_Item . '=' . explodeArray($_GET[$l_Item]);
            } else {
              $l_string .= '&' . $l_Item . '=' . $l_valor;
            }
          }
          elseif (strtoupper($p_method) == 'POST') {
            if (is_array($_GET[$l_Item])) {
              $l_string .= '<INPUT TYPE="HIDDEN" NAME="' . $l_Item . '" VALUE="' . explodeArray($_GET[$l_Item]) . '">';
            } else {
              $l_string .= '<INPUT TYPE="HIDDEN" NAME="' . $l_Item . '" VALUE="' . $l_valor . '">';
            }
          }
        }
      }
    }
    return $l_string;
  }
  // =========================================================================
  // Montagem de formul�rio para retorno � p�gina anterior
  // rerecebendo o nome do campo
  // a ser dado focus no formul�rio original
  // -------------------------------------------------------------------------
  function RetornaFormulario($l_troca = null, $l_sg = null, $l_menu = null, $l_o = null, $l_dir = null, $l_pagina = null, $l_par = null, $l_p1 = null, $l_p2 = null, $l_p3 = null, $l_p4 = null, $l_tp = null, $l_r = null) {
    extract($GLOBALS);
    $l_form = '';
    // Os par�metros informados prevalecem sobre os valores default
    if (nvl($l_pagina, '') != '') {
      $l_form .= AbreForm('RetornaDados', $l_dir . $l_pagina . $l_par, 'POST', null, null, nvl($l_p1, $_POST['P1']), nvl($l_p2, $_POST['P2']), nvl($l_p3, $_POST['P3']), nvl($l_p4, $_POST['P4']), nvl($l_tp, $_POST['TP']), nvl($l_sg, $_POST['SG']), nvl($l_r, $_POST['R']), nvl($l_o, $_POST['O']), 'texto');
    } else {
      $l_form .= AbreForm('RetornaDados', nvl($w_dir . $_POST['R'], $_SERVER['HTTP_REFERER']), 'POST', null, null, nvl($l_p1, $_POST['P1']), nvl($l_p2, $_POST['P2']), nvl($l_p3, $_POST['P3']), nvl($l_p4, $_POST['P4']), nvl($l_tp, $_POST['TP']), nvl($l_sg, $_POST['SG']), nvl($l_r, $_POST['R']), nvl($l_o, $_POST['O']), 'texto');
    }
    if (nvl($l_troca, '') != '')
      $l_form .= chr(13) . '<INPUT TYPE="HIDDEN" NAME="w_troca" VALUE="' . $l_troca . '">';
    if (nvl($l_menu, '') != '')
      $l_form .= chr(13) . '<INPUT TYPE="HIDDEN" NAME="w_menu" VALUE="' . $l_menu . '">';
    if (nvl($w_dir . $_POST['R'], '') != '') {
      foreach ($_GET as $l_Item => $l_valor) {
        if ($l_Item != 'par') {
          if (is_array($_GET[$l_Item])) {
            $l_form .= chr(13) . '<INPUT TYPE="HIDDEN" NAME="' . $l_Item . '[]" VALUE="' . explodeArray($_GET[$l_Item]) . '">';
          } else {
            $l_form .= chr(13) . '<INPUT TYPE="HIDDEN" NAME="' . $l_Item . '" VALUE="' . $l_valor . '">';
          }
        }
      }
    }
    foreach ($_POST as $l_Item => $l_valor) {
      if (strpos($l_form, 'NAME="' . $l_Item . '"') === false) {
        if ($l_Item != 'w_troca' && $l_Item != 'w_assinatura' && $l_Item != 'Password' && $l_Item != 'R' && $l_Item != 'P1' && $l_Item != 'P2' && $l_Item != 'P3' && $l_Item != 'P4' && $l_Item != 'TP' && $l_Item != 'O') {
          if (is_array($_POST[$l_Item])) {
            $l_form .= chr(13) . '<INPUT TYPE="HIDDEN" NAME="' . $l_Item . '[]" VALUE="' . explodeArray($_POST[$l_Item]) . '">';
          } else {
            $l_form .= chr(13) . '<INPUT TYPE="HIDDEN" NAME="' . $l_Item . '" VALUE="' . $l_valor . '">';
          }
        }
      }
    }
    ShowHTML($l_form);
    ShowHTML('</FORM>');
    ScriptOpen('JavaScript');
    // Registra no servidor syslog erro na assinatura eletr�nica
    if (nvl($l_troca, 'x') == 'w_assinatura') {
      $w_resultado = enviaSyslog('AI', 'ASSINATURA INV�LIDA', '(' . $_SESSION['SQ_PESSOA'] . ') ' . $_SESSION['NOME_RESUMIDO']);
      if ($w_resultado > '')
        ShowHTML('  alert(\'ATEN��O: erro no registro do log.\n' . $w_resultado . '\');');
    }
    ShowHTML('  document.forms["RetornaDados"].submit();');
    ScriptClose();
  }
  // =========================================================================
  // Exibe o conte�do da querystring, do formul�rio e das vari�veis de sess�o
  // -------------------------------------------------------------------------
  function ExibeVariaveis() {
    extract($GLOBALS);
    ShowHTML('<DT><font face="Verdana" size=1><b>Dados da querystring:</b></font>');
    foreach ($_GET as $chv => $vlr) {
      if (strpos(strtoupper($chv), 'W_ASSINATURA') === false) {
        ShowHTML('<DD><FONT FACE="courier" size=1>' . $chv . ' => [' . $vlr . ']</font><br>');
      }
    }
    ShowHTML('</DT>');
    ShowHTML('<DT><font face="Verdana" size=1><b>Dados do formul�rio:</b></font>');
    foreach ($_POST as $chv => $vlr) {
      if (strpos(strtoupper($chv), 'W_ASSINATURA') === false) {
        ShowHTML('<DD><FONT FACE="courier" size=1>' . $chv . ' => [' . $vlr . ']</font><br>');
      }
    }
    ShowHTML('</DT>');
    ShowHTML('<DT><font face="Verdana" size=1><b>Vari�veis de sess�o</b></font>:');
    foreach ($_SESSION as $chv => $vlr) {
      if (strpos(strtoupper($chv), 'W_ASSINATURA') === false) {
        ShowHTML('<DD><FONT FACE="courier" size=1>' . $chv . ' => [' . $vlr . ']</font><br>');
      }
    }
    ShowHTML('</DT>');
    ShowHTML('<DT><font face="Verdana" size=1><b>Vari�veis do servidor</b></font>:');
    foreach ($_SERVER as $chv => $vlr) {
      ShowHTML('<DD><FONT FACE="courier" size=1>' . $chv . ' => [' . $vlr . ']</font><br>');
    }
    ShowHTML('</DT>');
    $w_item = null;
    exit ();
  }

  // =========================================================================
  // Montagem da URL para consulta ao m�dulo de telefonia
  // -------------------------------------------------------------------------
  function consultaTelefone($p_cliente) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');

    // Verifica se o cliente tem o m�dulo de telefonia contratado
    include_once ($w_dir_volta . 'classes/sp/db_getSiwCliModLis.php');
    $l_rs_ac = db_getSiwCliModLis :: getInstanceOf($dbms, $p_cliente, null, 'TT');
    $l_mod_tt = false;
    $l_string = '';
    foreach ($l_rs_ac as $row) {
      $l_mod_tt = true;
      $l_string = '<a href="' . $conRootSIW . montaUrl('LIGACAO') . '" target="telefone" title="Procurar na base de liga��es telef�nicas."><img src="' . $conRootSIW . '/images/icone/fone_1.gif" border=0></a>';
      break;
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL para visualiza��o de uma solicita��o
  // -------------------------------------------------------------------------
  function ExibeSolic($l_dir, $l_chave, $l_texto = null, $l_exibe_titulo = null, $l_word = null) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if ($_REQUEST['p_tipo'] == 'PDF') {
      $w_embed = 'WORD';
    }
    if (strpos($l_texto, '|@|') !== false) {
      $l_array = explode('|@|', $l_texto);
      if (nvl($l_word, '') == '') {
        $l_hint = $l_array[4];
        if ($w_embed != 'WORD') {
          $l_string = '<A class="hl" HREF="' . $conRootSIW . $l_array[10] . '&O=L&w_chave=' . $l_chave . '&P1=' . $l_array[6] . '&P2=' . $l_array[7] . '&P3=' . $l_array[8] . '&P4=' . $l_array[9] . '&TP=' . $TP . '&SG=' . $l_array[5] . '" target="_blank" title="' . $l_hint . '">' . $l_array[1] . (($l_exibe_titulo == 'S') ? ' - ' . $l_array[2] : '') . '</a>';
        } else {
          $l_string = $l_array[1] . (($l_exibe_titulo == 'S') ? ' - ' . $l_array[2] : '');
        }
      } else {
        $l_string = $l_array[1] . (($l_exibe_titulo == 'S') ? ' - ' . $l_array[2] : '');
      }
    }
    elseif (nvl($l_chave, '') != '') {
      include_once ($w_dir_volta . 'classes/sp/db_getSolicData.php');
      $RS = db_getSolicData :: getInstanceOf($dbms, $l_chave);
      $l_hint = $l_array[4];
      $l_array = explode('|@|', f($RS, 'dados_solic'));
      if (nvl($l_word, '') == '') {
        $l_hint = 'Exibe as informa��es deste registro.';
        $l_string = '<A class="hl" HREF="' . $conRootSIW . $l_array[10] . '&O=L&w_chave=' . $l_chave . '&P1=' . $l_array[6] . '&P2=' . $l_array[7] . '&P3=' . $l_array[8] . '&P4=' . $l_array[9] . '&TP=' . $TP . '&SG=' . $l_array[5] . '" target="_blank" title="' . $l_hint . '">' . $l_array[1] . (($l_exibe_titulo == 'S') ? ' - ' . $l_array[2] : '') . '</a>';
      } else {
        $l_string = $l_array[1] . (($l_exibe_titulo == 'S') ? ' - ' . $l_array[2] : '');
      }
    } else {
      $l_string = $l_texto;
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de uma pessoa
  // -------------------------------------------------------------------------
  function ExibePessoa($p_dir, $p_cliente, $p_pessoa, $p_tp, $p_nome) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_nome, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . montaURL_JS(null, $conRootSIW . 'seguranca.php?par=TELAUSUARIO&w_cliente=' . $p_cliente . '&w_sq_pessoa=' . $p_pessoa . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=') . '\',\'Pessoa\',\'width=780,height=300,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados desta pessoa!">' . $p_nome . '</A>';
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de uma pessoa no relat�rio de permiss�es
  // -------------------------------------------------------------------------
  function ExibePessoaRel($p_dir, $p_cliente, $p_pessoa, $p_nome, $p_nome_resumido, $p_tipo) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_nome, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="' . $conRootSIW . $p_dir . 'relatorios.php?par=TELAUSUARIOREL&w_cliente=' . $p_cliente . '&w_sq_pessoa=' . $p_pessoa . '&w_tipo=' . $p_tipo . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&SG=" title="' . $p_nome . '">' . $p_nome_resumido . '</a>';
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de um fornecedor
  // -------------------------------------------------------------------------
  function ExibeFornecedor($p_dir, $p_cliente, $p_pessoa, $p_tp, $p_nome) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_nome, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . montaURL_JS(null, $conRootSIW . 'mod_eo/fornecedor.php?par=Visual&w_sq_pessoa=' . $p_pessoa . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=') . '\',\'Fornecedor\',\'width=780,height=300,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados deste fornecedor!">' . $p_nome . '</A>';
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de um plano estrat�gico
  // -------------------------------------------------------------------------
  function ExibePlano($p_dir, $p_cliente, $p_plano, $p_tp, $p_nome, $p_pitce = null) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_nome, '') == '') {
      $l_string = '---';
    } else {
      if (nvl($p_pitce, '') == '') {
        $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . montaURL_JS(null, $conRootSIW . 'mod_pe/tabelas.php?par=TELAPLANO&w_cliente=' . $p_cliente . '&w_sq_plano=' . $p_plano . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=') . '\',\'plano\',\'width=780,height=500,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados deste plano!">' . $p_nome . '</A>';
      } else {
        $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . montaURL_JS(null, $conRootSIW . 'cl_pitce/tabelas.php?par=TELAPLANO&w_cliente=' . $p_cliente . '&w_sq_plano=' . $p_plano . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=') . '\',\'plano\',\'width=780,height=500,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados deste plano!">' . $p_nome . '</A>';
      }
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de uma pessoa com os pacontes vinculados
  // -------------------------------------------------------------------------
  function ExibeUnidadePacote($O, $p_cliente, $p_chave, $p_chave_aux, $p_unidade, $p_tp, $p_nome) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_nome, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'projeto.php?par=InteressadoPacote&w_chave=' . $p_chave . '&O=' . $O . '&w_chave_aux=' . $p_chave_aux . '&w_sq_unidade=' . $p_unidade . '&P1=' . $p_P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=' . $p_sg . '\',\'Interessados\',\'width=780,height=550,top=50,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados!">' . $p_nome . '</A>';
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de uma pessoa
  // -------------------------------------------------------------------------
  function VisualIndicador($p_dir, $p_cliente, $p_sigla, $p_tp, $p_nome) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    $l_RS = db_getIndicador :: getInstanceOf($dbms, $w_cliente, $w_usuario, null, null, null, $p_sigla, null, null, null, null, null, null, null, null, null, null, null, "EXISTE");
    if (count($l_RS) > 0) {
      if (Nvl($p_nome, '') == '') {
        $l_string = '---';
      } else {
        $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'mod_pe/indicador.php?par=TELAINDICADOR&w_cliente=' . $p_cliente . '&w_sigla=' . $p_sigla . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=' . '\',\'Indicador\',\'width=780,height=300,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados deste de indicador!">' . $p_nome . '</A>';
      }
    } else {
      $l_string = $p_sigla;
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de uma unidade
  // -------------------------------------------------------------------------
  function ExibeUnidade($p_dir, $p_cliente, $p_unidade, $p_sq_unidade, $p_tp) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_unidade, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . montaURL_JS(null, $conRootSIW . 'seguranca.php?par=TELAUNIDADE&w_cliente=' . $p_cliente . '&w_sq_unidade=' . $p_sq_unidade . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=') . '\',\'Unidade\',\'width=780,height=300,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados desta unidade!">' . $p_unidade . '</A>';
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de um recurso
  // -------------------------------------------------------------------------
  function ExibeRecurso($p_dir, $p_cliente, $p_nome, $p_chave, $p_tp, $p_solic) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_chave, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'mod_pe/recurso.php?par=TELARECURSO&w_cliente=' . $p_cliente . '&w_chave=' . $p_chave . '&w_solic=' . $p_solic . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=' . '\',\'Telarecurso\',\'width=785,height=570,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados deste recurso!">' . $p_nome . '</A>';
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de um material ou servi�o
  // -------------------------------------------------------------------------
  function ExibeMaterial($p_dir, $p_cliente, $p_nome, $p_chave, $p_tp, $p_solic) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_chave, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'mod_cl/catalogo.php?par=TELAMATERIAL&w_cliente=' . $p_cliente . '&w_chave=' . $p_chave . '&w_solic=' . $p_solic . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=' . '\',\'Telarecurso\',\'width=785,height=570,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados deste material ou servi�o!">' . $p_nome . '</A>';
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de uma restricao
  // -------------------------------------------------------------------------
  function ExibeRestricao($O, $p_dir, $p_cliente, $p_tipo, $p_chave, $p_chave_aux, $p_tp, $p_solic) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_tipo, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'mod_pr/restricao.php?par=VisualRestricao&w_cliente=' . $p_cliente . '&w_chave=' . $p_chave . '&w_chave_aux=' . $p_chave_aux . '&O=' . $O . '&w_solic=' . $p_solic . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=' . '\',\'VisualRestriao\',\'width=785,height=570,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados desta restricao!">' . $p_tipo . '</A>';
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de uma meta
  // -------------------------------------------------------------------------
  function ExibeMeta($O, $p_dir, $p_cliente, $p_nome, $p_chave, $p_chave_aux, $p_tp, $p_solic, $p_plano = null) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_plano, '') != '') {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'mod_pe/indicador.php?par=VisualMeta&w_cliente=' . $p_cliente . '&w_chave=&w_chave_aux=' . $p_chave_aux . '&w_plano=' . $p_plano . '&O=' . $O . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=' . '\',\'VisualMeta\',\'width=785,height=570,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados desta meta!">' . $p_nome . '</A>';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'mod_pe/indicador.php?par=VisualMeta&w_cliente=' . $p_cliente . '&w_chave=' . $p_chave . '&w_chave_aux=' . $p_chave_aux . '&w_solic=' . $p_solic . '&O=' . $O . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=' . '\',\'VisualMeta\',\'width=785,height=570,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados desta meta!">' . $p_nome . '</A>';
    }
    return $l_string;
  }

  // =========================================================================
  // Montagem da URL com os dados de um recurso
  // -------------------------------------------------------------------------
  function ExibeIndicador($p_dir, $p_cliente, $p_nome, $p_dados, $p_tp) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_dados, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'mod_pe/indicador.php?par=FramesAfericao&R=' . $w_pagina . $par . '&O=L' . $p_dados . '&P1=' . $l_p1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $TP . '&SG=' . $SG . '\',\'TelaIndicador\',\'width=785,height=570,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados deste indicador!">' . $p_nome . '</A>';
    }
    return $l_string;
  }

  function SelecaoOperacao($label, $accesskey, $hint, $chave, $chaveAux, $campo, $restricao, $atributo) {
    extract($GLOBALS);
    If (is_null($hint)) {
      ShowHTML('          <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_Disabled . ' ' . $atributo . '>');
    } Else {
      ShowHTML('          <td valign="top" title="' . $hint . '"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_Disabled . ' ' . $atributo . '>');
    }
    ShowHTML("          <option value=''>Todas");
    If (Nvl($chave, "") == "0")
      ShowHTML('          <option value="0" SELECTED>Consulta');
    Else
      ShowHTML('          <option value="0">Consulta');
    If (Nvl($chave, "") == "1")
      ShowHTML('          <option value="1" SELECTED>Inclus�o');
    Else
      ShowHTML('          <option value="1">Inclus�o');
    If (Nvl($chave, "") == "2")
      ShowHTML('          <option value="2" SELECTED>Altera��o');
    Else
      ShowHTML('          <option value="2">Altera��o');
    If (Nvl($chave, "") == "3")
      ShowHTML('          <option value="3" SELECTED>Exclus�o');
    Else
      ShowHTML('          <option value="3">Exclus�o');
    ShowHTML('          </select>');
  }

  function addDayIntoDate($date, $days) {
    $thisyear = substr($date, 0, 4);
    $thismonth = substr($date, 4, 2);
    $thisday = substr($date, 6, 2);
    $nextdate = mktime(0, 0, 0, $thismonth, $thisday + $days, $thisyear);
    return strftime("%Y%m%d", $nextdate);
  }

  function subDayIntoDate($date, $days, $p = 1) {
    $thisyear = substr($date, 0, 4);
    $thismonth = substr($date, 4, 2);
    $thisday = substr($date, 6, 2);
    $nextdate = mktime(0, 0, 0, $thismonth, $thisday - $days, $thisyear);
    if ($p == 1)
      return strftime("%Y%m%d", $nextdate);
    else
      return strftime("%d/%m/%Y", $nextdate);
  }

  //  =========================================================================
  //  Montagem da sele��o de bloco de dados
  //  -------------------------------------------------------------------------
  function SelecaoBlocoDados($label, $accesskey, $hint, $chave, $chaveAux, $campo, $restricao, $atributo) {
    extract($GLOBALS);
    If (is_null($hint)) {
      ShowHTML('          <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_Disabled . ' ' . $atributo . '>');
    } Else {
      ShowHTML('          <td valign="top" title="' . $hint . '"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_Disabled . ' ' . $atributo . '>');
    }
    ShowHTML('          <option value="">Todos');
    If (Nvl($chave, "") == "223")
      ShowHTML('          <option value="223" SELECTED>Administrativo');
    Else
      ShowHTML('          <option value="223">Administrativo');
    If (Nvl($chave, "") == "135")
      ShowHTML('          <option value="135" SELECTED>Arquivos');
    Else
      ShowHTML('          <option value="135">Arquivos');
    If (Nvl($chave, "") == "133")
      ShowHTML('          <option value="133" SELECTED>Calend�rio');
    Else
      ShowHTML('          <option value="133">Calend�rio');
    If (Nvl($chave, "") == "220")
      ShowHTML('          <option value="220" SELECTED>Dados adicionais');
    Else
      ShowHTML('          <option value="220">Dados adicionais');
    If (Nvl($chave, "") == "123")
      ShowHTML('          <option value="123" SELECTED>Dados b�sicos');
    Else
      ShowHTML('          <option value="123">Dados b�sicos');
    If (Nvl($chave, "") == "144")
      ShowHTML('          <option value="144" SELECTED>Dados do site');
    Else
      ShowHTML('          <option value="144">Dados do site');
    If (Nvl($chave, "") == "221")
      ShowHTML('          <option value="221" SELECTED>Mensagens');
    Else
      ShowHTML('          <option value="221">Mensagens');
    If (Nvl($chave, "") == "127")
      ShowHTML('          <option value="127" SELECTED>Modalidades');
    Else
      ShowHTML('          <option value="127">Modalidades');
    If (Nvl($chave, "") == "222")
      ShowHTML('          <option value="222" SELECTED>Not�cias');
    Else
      ShowHTML('          <option value="222">Not�cias');
    ShowHTML("          </select>");
  }

  function ImprimeCabecalho() {
    extract($GLOBALS);
    ShowHTML('<tr><td align="center">');
    ShowHTML('    <TABLE WIDTH="100%" bgcolor="" & conTableBgColor & "" BORDER="" & conTableBorder & "" CELLSPACING="" & conTableCellSpacing & "" CELLPADDING="" & conTableCellPadding & "" BorderColorDark="" & conTableBorderColorDark & "" BorderColorLight="" & conTableBorderColorLight & "">');
    ShowHTML('        <tr bgcolor="#DCDCDC" align="center">');
    //var_dump($p_agrega);
    switch ($_REQUEST['p_agrega']) {
      Case "SECRETARIA" :
        ShowHTML(' <td><font size="1"><b>Secretaria</font></td>');
        break;
      Case "REGIONAL" :
        ShowHTML(' <td><font size="1"><b>Regional de Ensino</font></td>');
        break;
      Case "ESCOLA" :
        ShowHTML(' <td><font size="1"><b>Escola</font></td>');
        break;
      Case "BLOCODADOS" :
        ShowHTML(' <td><font size="1"><b>Bloco de Dados</font></td>');
        break;
    }
    ShowHTML('          <td><font size="1"><b>Total</font></td>');
    ShowHTML('          <td><font size="1"><b>Consulta</font></td>');
    ShowHTML('          <td><font size="1"><b>Inclus�es</font></td>');
    ShowHTML('          <td><font size="1"><b>Altera��es</font></td>');
    ShowHTML('          <td><font size="1"><b>Exclus�es</font></td>');
    ShowHTML('        </tr>');
  }

  //  =========================================================================
  //  Final da rotina
  //  ---------------------

  // =========================================================================
  // Montagem da URL com os dados da etapa
  // -------------------------------------------------------------------------
  function ExibeEtapa($O, $p_chave, $p_chave_aux, $p_tipo, $p_P1, $p_etapa, $p_tp, $p_sg) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if ($p_tipo == 'PDF') {
      $w_embed = 'WORD';
    }

    if (Nvl($p_etapa, '') == '') {
      $l_string = "---";
    } else {
      if ($w_embed == 'WORD') {
        $l_string .= $p_etapa;
      } else {
        $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . (($w_dir == 'mod_pr/') ? '' : $w_dir) . 'projeto.php?par=AtualizaEtapa&w_chave=' . $p_chave . '&O=' . $O . '&w_chave_aux=' . $p_chave_aux . '&w_tipo=' . $p_tipo . '&P1=' . $p_P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=' . $p_sg . '\',\'Etapa\',\'width=780,height=550,top=50,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;" title="Clique para exibir os dados!">' . $p_etapa . '</A>';
      }
    }
    return $l_string;
  }

  // =========================================================================
  // Exibe imagem da restri��o conforme tipo e criticidade
  // -------------------------------------------------------------------------
  function ExibeImagemRestricao($l_tipo, $l_imagem = null, $l_legenda = 0) {
    extract($GLOBALS);
    $l_string = '';
    if ($l_legenda) {
      $l_string .= '<tr valign="top">';
      $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgProblem . '" border=0 width=10 height=10 align="center"><td>Problema.';
      $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgRiskHig . '" border=0 width=10 height=10 align="center"><td>Risco de alta criticidade.';
      $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgRiskMed . '" border=0 width=10 height=10 align="center"><td>Risco de moderada ou baixa criticidade. ';
    } else {
      if ($l_imagem == 'P') {
        if (Nvl($l_tipo, 'N') != 'N') {
          switch ($l_tipo) {
            case 'S1' :
              $l_string .= '<img title="Problema de baixa criticidade" src="' . $conRootSIW . $conImgProblem . '" width=10 height=10 border=0 align="center">';
              break;
            case 'S2' :
              $l_string .= '<img title="Problema de moderada criticidade" src="' . $conRootSIW . $conImgProblem . '" width=10 height=10 border=0 align="center">';
              break;
            case 'S3' :
              $l_string .= '<img title="Problema de alta criticidade" src="' . $conRootSIW . $conImgProblem . '" width=10 height=10 border=0 align="center">';
              break;
          }
        }
      } else {
        if (Nvl($l_tipo, 'N') != 'N') {
          switch ($l_tipo) {
            case 'S1' :
              $l_string .= '<img title="Problema de baixa criticidade" src="' . $conRootSIW . $conImgProblem . '" width=10 height=10 border=0 align="center">';
              break;
            case 'S2' :
              $l_string .= '<img title="Problema de moderada criticidade" src="' . $conRootSIW . $conImgProblem . '" width=10 height=10 border=0 align="center">';
              break;
            case 'S3' :
              $l_string .= '<img title="Problema de alta criticidade" src="' . $conRootSIW . $conImgProblem . '" width=10 height=10 border=0 align="center">';
              break;
            case 'N1' :
              $l_string .= '<img title="Risco de baixa criticidade" src="' . $conRootSIW . $conImgRiskLow . '" width=10 height=10 border=0 align="center">';
              break;
            case 'N2' :
              $l_string .= '<img title="Risco de moderada criticidade" src="' . $conRootSIW . $conImgRiskMed . '" width=10 height=10 border=0 align="center">';
              break;
            case 'N3' :
              $l_string .= '<img title="Risco de alta criticidade" src="' . $conRootSIW . $conImgRiskHig . '" width=10 height=10 border=0 align="center">';
              break;
          }
        }
      }
    }
    return $l_string;
  }

  // =========================================================================
  // Exibe imagem do �cone smile
  // -------------------------------------------------------------------------
  function ExibeSmile($l_tipo, $l_andamento, $l_legenda = 0) {
    extract($GLOBALS);
    $l_tipo = trim(strtoupper($l_tipo));
    $l_andamento = nvl($l_andamento, 0);
    if ($l_legenda) {
      if ($l_tipo == 'IDE') {
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgSmAtraso . '" border=0 width=10 height=10 align="center"><td>Fora da faixa desej�vel (abaixo de 70%).';
        $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgSmAviso . '" border=0 width=10 height=10 align="center"><td>Pr�ximo da faixa desej�vel (de 70% a 89,99% ou acima de 120%).';
        $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgSmNormal . '" border=0 width=10 height=10 align="center"><td>Na faixa desej�vel (de 90% a 120%). ';
      }
      elseif ($l_tipo == 'IDC') {
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgSmAtraso . '" border=0 width=10 height=10 align="center"><td>Fora da faixa desej�vel (abaixo de 70%).';
        $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgSmAviso . '" border=0 width=10 height=10 align="center"><td>Pr�ximo da faixa desej�vel (de 70% a 89,99% ou acima de 120%).';
        $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgSmNormal . '" border=0 width=10 height=10 align="center"><td>Na faixa desej�vel (de 90% a 120%). ';
      }
    } else {
      if ($l_tipo == 'IDE') {
        if ($l_andamento < 70)
          $l_string .= '<img title="IDE fora da faixa desej�vel." src="' . $conRootSIW . $conImgSmAtraso . '" border=0 width="10" height="10">';
        elseif ($l_andamento < 90 || $l_andamento > 120) $l_string .= '<img title="IDE pr�ximo da faixa desej�vel." src="' . $conRootSIW . $conImgSmAviso . '" border=0 width="10" height="10">';
        else
          $l_string .= '<img title="IDE na faixa desej�vel." src="' . $conRootSIW . $conImgSmNormal . '" border=0 width="10" height="10">';
      }
      elseif ($l_tipo == 'IDC') {
        if ($l_andamento < 70)
          $l_string .= '<img title="IDC fora da faixa desej�vel." src="' . $conRootSIW . $conImgSmAtraso . '" border=0 width="10" height="10">';
        elseif ($l_andamento < 90 || $l_andamento > 120) $l_string .= '<img title="IDC pr�ximo da faixa desej�vel." src="' . $conRootSIW . $conImgSmAviso . '" border=0 width="10" height="10">';
        else
          $l_string .= '<img title="IDC na faixa desej�vel." src="' . $conRootSIW . $conImgSmNormal . '" border=0 width="10" height="10">';
      }
    }
    return $l_string;
  }

  // =========================================================================
  // Exibe sinalizador para pesquisa de pre�o
  // -------------------------------------------------------------------------
  function ExibeSinalPesquisa($l_legenda, $l_inicio, $l_fim, $l_dias_aviso = 0) {
    extract($GLOBALS);
    include_once ($w_dir_volta . 'classes/sp/db_getCLParametro.php');
    $RS_Parametro = db_getCLParametro :: getInstanceOf($dbms, $w_cliente, null, null);
    foreach ($RS_Parametro as $row_parametro) {
      $RS_Parametro = $row_parametro;
      break;
    }
    if ($l_legenda) {
      $l_string .= '<tr valign="top">';
      $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgSmAtraso . '" border=0 width=10 height=10 align="center"><td>Pesquisa com mais de ' . f($RS_Parametro, 'dias_validade_pesquisa') . ' dias.';
      $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgSmAviso . '" border=0 width=10 height=10 align="center"><td>Pesquisa com ' . f($RS_Parametro, 'dias_aviso_pesquisa') . ' dias ou menos de validade.';
      $l_string .= '<td width="1%" nowrap><img src="' . $conRootSIW . $conImgSmNormal . '" border=0 width=10 height=10 align="center"><td>Pesquisa v�lida. ';
    } else {
      if ($l_fim < addDays(time(), -1))
        $l_string .= '<img title="Pesquisa com mais de ' . f($RS_Parametro, 'dias_validade_pesquisa') . ' dias." src="' . $conRootSIW . $conImgSmAtraso . '" border=0 width="10" height="10" align="center">';
      elseif ($l_dias_aviso <= time()) $l_string .= '<img title="Pesquisa com ' . f($RS_Parametro, 'dias_aviso_pesquisa') . ' dias ou menos de validade." src="' . $conRootSIW . $conImgSmAviso . '" border=0 width="10" height="10" align="center">';
      else
        $l_string .= '<img title="Pesquisa v�lida." src="' . $conRootSIW . $conImgSmNormal . '" border=0 width="10" height="10" align="center">';
    }
    return $l_string;
  }

  // =========================================================================
  // Exibe sinalizador para pesquisa de pre�o
  // -------------------------------------------------------------------------
  function exibeImagemAnexo($l_exibe = 0) {
    extract($GLOBALS);
    $l_string = '';
    if ($l_exibe > 0)
      $l_string .= '<img title="H� arquivos dispon�veis para download." src="' . $conRootSIW . $conImgDownload . '" border=0 width="14" height="14" align="center">';
    return $l_string;
  }

  // =========================================================================
  // Exibe imagem da solicita��o informada
  // -------------------------------------------------------------------------
  function ExibeImagemSolic($l_tipo, $l_inicio, $l_fim, $l_inicio_real, $l_fim_real, $l_aviso, $l_dias_aviso, $l_tramite, $l_perc, $l_legenda = 0, $l_restricao = null) {
    extract($GLOBALS);
    $l_string = '';
    $l_imagem = '';
    $l_title = '';
    $l_tipo = strtoupper($l_tipo);
    if ($l_legenda) {
      if ($l_tipo == 'ETAPA') {
        // Etapas de projeto
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgAtraso . '" border=0 width=10 align="center"><td>Execu��o n�o iniciada. Fim previsto superado.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgAviso . '" border=0 width=10 heigth=10 align="center"><td>Execu��o n�o iniciada. Percentual de conclus�o incompat�vel com os dias transcorridos.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgNormal . '" border=0 width=10 heigth=10 align="center"><td>Execu��o n�o iniciada. Prazo final dentro do previsto.';
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgStAtraso . '" border=0 width=10 heigth=10 align="center"><td>Em execu��o. Fim previsto superado.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgStAviso . '" border=0 width=10 heigth=10 align="center"><td>Em execu��o. Percentual de conclus�o incompat�vel com os dias transcorridos.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgStNormal . '" border=0 width=10 heigth=10 align="center"><td>Em execu��o. Prazo final dentro do previsto.';
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgOkAtraso . '" border=0 width=10 heigth=10 align="center"><td>Execu��o conclu�da ap�s a data prevista.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgOkAcima . '" border=0 width=10 heigth=10 align="center"><td>Execu��o conclu�da antes da data prevista.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgOkNormal . '" border=0 width=10 heigth=10 align="center"><td>Execu��o conclu�da na data prevista.';
      }
      elseif (substr($l_tipo, 0, 2) == 'GD' || substr($l_tipo, 0, 2) == 'SR' || substr($l_tipo, 0, 2) == 'PJ') {
        // Tarefas, demandas eventuais e recursos log�sticos
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgCancel . '" border=0 width=10 heigth=10 align="center"><td>Registro cancelado.';
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgAtraso . '" border=0 width=10 heigth=10 align="center"><td>Execu��o n�o iniciada. Fim previsto superado.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgAviso . '" border=0 width=10 heigth=10 align="center"><td>Execu��o n�o iniciada. Fim previsto pr�ximo.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgNormal . '" border=0 width=10 heigth=10 align="center"><td>Execu��o n�o iniciada. Prazo final dentro do previsto.';
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgStAtraso . '" border=0 width=12 align="center"><td>Em execu��o. Fim previsto superado.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgStAviso . '" border=0 width=10 heigth=10 align="center"><td>Em execu��o. Fim previsto pr�ximo.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgStNormal . '" border=0 width=10 heigth=10 align="center"><td>Em execu��o. Prazo final dentro do previsto.';
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgOkAtraso . '" border=0 width=12 align="center"><td>Execu��o conclu�da ap�s a data prevista.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgOkAcima . '" border=0 width=10 heigth=10 align="center"><td>Execu��o conclu�da antes da data prevista.';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgOkNormal . '" border=0 width=10 heigth=10 align="center"><td>Execu��o conclu�da na data prevista.';
      }
      elseif (substr($l_tipo, 0, 2) == 'PD') {
        // Viagens
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgCancel . '" border=0 width=10 heigth=10 align="center"><td>Registro cancelado';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgAviso . '" border=0 width=10 heigth=10 align="center"><td>In�cio pr�ximo';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgNormal . '" border=0 width=10 heigth=10 align="center"><td>N�o iniciada';
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgStAtraso . '" border=0 width=10 heigth=10 align="center"><td>Tramita��o em atraso';
        $l_string .= '<td><td>';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgStNormal . '" border=0 width=10 heigth=10 align="center"><td>Em andamento';
        $l_string .= '<tr valign="top">';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgOkAtraso . '" border=0 width=10 heigth=10 align="center"><td>Tramita��o em atraso';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgOkAcima . '" border=0 width=10 heigth=10 align="center"><td>Pendente presta��o de contas';
        $l_string .= '<td width="1%" nowrap><img src="' . $conImgOkNormal . '" border=0 width=10 heigth=10 align="center"><td>Encerrada';
      }
    } else {
      if ($l_tipo == 'ETAPA') {
        // Etapas de projeto
        if ($l_perc < 100) {
          if (nvl($l_inicio_real, '') == '') {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgAtraso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto superado.';
            }
            elseif (((time() - $l_inicio) / ($l_fim - $l_inicio +1)) * 100 > $l_perc) {
              $l_imagem = $conImgAviso;
              $l_title = 'Execu��o n�o iniciada. Percentual de conclus�o incompat�vel com os dias transcorridos.';
            } else {
              $l_imagem = $conImgNormal;
              $l_title = 'Execu��o n�o iniciada. Prazo final dentro do previsto.';
            }
          } else {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgStAtraso;
              $l_title = 'Em execu��o. Fim previsto superado.';
            }
            elseif (((time() - $l_inicio) / ($l_fim - $l_inicio +1)) * 100 > $l_perc) {
              $l_imagem = $conImgStAviso;
              $l_title = 'Em execu��o. Percentual de conclus�o incompat�vel com os dias transcorridos.';
            } else {
              $l_imagem = $conImgStNormal;
              $l_title = 'Em execu��o. Prazo final dentro do previsto.';
            }
          }
        } else {
          if ($l_fim < Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Execu��o conclu�da ap�s a data prevista.';
          }
          elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAcima;
            $l_title = 'Execu��o conclu�da antes da data prevista.';
          } else {
            $l_imagem = $conImgOkNormal;
            $l_title = 'Execu��o conclu�da na data prevista.';
          }
        }
      }
      elseif (substr($l_tipo, 0, 2) == 'GC') {
        // Contratos e conv�nios
        if ($l_tramite != 'AT' && $l_tramite != 'CR') {
          if ($l_tramite == 'CA') {
            $l_imagem = $conImgCancel;
            $l_title = 'Registro cancelado.';
          }
          elseif ($l_tramite == 'CI') {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgAtraso;
              $l_title = 'Execu��o n�o iniciada. Vig�ncia prevista ultrapassada.';
            }
            elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
              $l_imagem = $conImgAviso;
              $l_title = 'Execu��o n�o iniciada. Vig�ncia prevista pr�xima do t�rmino.';
            } else {
              $l_imagem = $conImgNormal;
              $l_title = 'Execu��o n�o iniciada.';
            }
          }
          elseif ($l_tramite == 'ER') {
            $l_imagem = $conImgStAcima;
            $l_title = 'Vig�ncia encerrada, com restos a pagar.';
          } else {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgStAtraso;
              $l_title = 'Em execu��o. Vig�ncia prevista ultrapassada.';
            }
            elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
              $l_imagem = $conImgStAviso;
              $l_title = 'Em execu��o. Vig�ncia prevista pr�xima do t�rmino.';
            } else {
              $l_imagem = $conImgStNormal;
              $l_title = 'Em execu��o.';
            }
          }
        } else {
          if ($l_fim < Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Vig�ncia superior � prevista.';
          }
          elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAcima;
            $l_title = 'Vig�ncia encerrada antes do previsto.';
          } else {
            $l_imagem = $conImgOkNormal;
            $l_title = 'Vig�ncia encerrada conforme previs�o.';
          }
        }
      }
      elseif (substr($l_tipo, 0, 2) == 'GD') {
        // Tarefas, demandas eventuais e demandas de triagem
        if ($l_tramite != 'AT') {
          if ($l_tramite == 'CA') {
            $l_imagem = $conImgCancel;
            $l_title = 'Registro cancelado.';
          }
          elseif ($l_tramite == 'CI' || $l_restricao == 'SEMEXECUCAO') {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgAtraso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
              $l_imagem = $conImgAviso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgNormal;
              $l_title = 'Execu��o n�o iniciada. Prazo final dentro do previsto.';
            }
          } else {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgStAtraso;
              $l_title = 'Em execu��o. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
              $l_imagem = $conImgStAviso;
              $l_title = 'Em execu��o. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgStNormal;
              $l_title = 'Em execu��o. Prazo final dentro do previsto.';
            }
          }
        } else {
          if ($l_fim < Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Execu��o conclu�da ap�s a data prevista.';
          }
          elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAcima;
            $l_title = 'Execu��o conclu�da antes da data prevista.';
          } else {
            $l_imagem = $conImgOkNormal;
            $l_title = 'Execu��o conclu�da na data prevista.';
          }
        }
      }
      elseif (substr($l_tipo, 0, 2) == 'FN') {
        // Tarefas e demandas eventuais
        if ($l_tramite != 'AT') {
          if ($l_tramite == 'CA') {
            $l_imagem = $conImgCancel;
            $l_title = 'Registro cancelado.';
          }
          elseif ($l_tramite == 'CI') {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgAtraso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
              $l_imagem = $conImgAviso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgNormal;
              $l_title = 'Execu��o n�o iniciada. Prazo final dentro do previsto.';
            }
          } else {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgStAtraso;
              $l_title = 'Em execu��o. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
              $l_imagem = $conImgStAviso;
              $l_title = 'Em execu��o. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgStNormal;
              $l_title = 'Em execu��o. Prazo final dentro do previsto.';
            }
          }
        } else {
          if ($l_fim < Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Execu��o conclu�da ap�s a data prevista.';
          }
          elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAcima;
            $l_title = 'Execu��o conclu�da antes da data prevista.';
          } else {
            $l_imagem = $conImgOkNormal;
            $l_title = 'Execu��o conclu�da na data prevista.';
          }
        }
      }
      elseif (substr($l_tipo, 0, 2) == 'SR') {
        // Tarefas
        if ($l_tramite != 'AT') {
          if ($l_tramite == 'CA') {
            $l_imagem = $conImgCancel;
            $l_title = 'Registro cancelado.';
          }
          elseif ($l_tramite == 'CI') {
            if ($l_fim < time()) {
              $l_imagem = $conImgAtraso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
              $l_imagem = $conImgAviso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgNormal;
              $l_title = 'Execu��o n�o iniciada. Prazo final dentro do previsto.';
            }
          } else {
            if ($l_fim < time()) {
              $l_imagem = $conImgStAtraso;
              $l_title = 'Em execu��o. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
              $l_imagem = $conImgStAviso;
              $l_title = 'Em execu��o. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgStNormal;
              $l_title = 'Em execu��o. Prazo final dentro do previsto.';
            }
          }
        } else {
          if ($l_fim < Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Execu��o conclu�da ap�s a data prevista.';
          }
          elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAcima;
            $l_title = 'Execu��o conclu�da antes da data prevista.';
          } else {
            $l_imagem = $conImgOkNormal;
            $l_title = 'Execu��o conclu�da na data prevista.';
          }
        }
      }
      elseif (substr($l_tipo, 0, 2) == 'PJ') {
        // Projetos
        if ($l_tramite != 'AT') {
          if ($l_tramite == 'CA') {
            $l_imagem = $conImgCancel;
            $l_title = 'Registro cancelado.';
          }
          elseif ($l_tramite == 'CI') {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgAtraso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
              $l_imagem = $conImgAviso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgNormal;
              $l_title = 'Execu��o n�o iniciada. Prazo final dentro do previsto.';
            }
          } else {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgStAtraso;
              $l_title = 'Em execu��o. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
              $l_imagem = $conImgStAviso;
              $l_title = 'Em execu��o. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgStNormal;
              $l_title = 'Em execu��o. Prazo final dentro do previsto.';
            }
          }
        } else {
          if ($l_fim < Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Execu��o conclu�da ap�s a data prevista.';
          }
          elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAcima;
            $l_title = 'Execu��o conclu�da antes da data prevista.';
          } else {
            $l_imagem = $conImgOkNormal;
            $l_title = 'Execu��o conclu�da na data prevista.';
          }
        }
      }
      elseif (substr($l_tipo, 0, 2) == 'CL') {
        // Projetos
        if ($l_tramite != 'AT') {
          if ($l_tramite == 'CA') {
            $l_imagem = $conImgCancel;
            $l_title = 'Registro cancelado.';
          }
          elseif ($l_tramite == 'CI') {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgAtraso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
              $l_imagem = $conImgAviso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgNormal;
              $l_title = 'Execu��o n�o iniciada. Prazo final dentro do previsto.';
            }
          } else {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgStAtraso;
              $l_title = 'Em execu��o. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
              $l_imagem = $conImgStAviso;
              $l_title = 'Em execu��o. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgStNormal;
              $l_title = 'Em execu��o. Prazo final dentro do previsto.';
            }
          }
        } else {
          if ($l_fim < Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Execu��o conclu�da ap�s a data prevista.';
          }
          elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAcima;
            $l_title = 'Execu��o conclu�da antes da data prevista.';
          } else {
            $l_imagem = $conImgOkNormal;
            $l_title = 'Execu��o conclu�da na data prevista.';
          }
        }
      }
      elseif (substr($l_tipo, 0, 2) == 'PD') {
        // Viagens
        if ($l_tramite == 'CA') {
          $l_imagem = $conImgCancel;
          $l_title = 'Registro cancelado.';
        }
        elseif ($l_fim < addDays(time(), -1)) {
          if ($l_tramite == 'AT') {
            $l_imagem = $conImgOkNormal;
            $l_title = 'Miss�o encerrada.';
          }
          elseif ($l_tramite != 'EE') {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Miss�o com tramita��o em atraso.';
          } else {
            $l_imagem = $conImgOkAcima;
            $l_title = 'Miss�o pendente de presta��o de contas.';
          }
        }
        elseif ($l_inicio > time()) {
          if ($l_dias_aviso <= time()) {
            $l_imagem = $conImgAviso;
            $l_title = 'Miss�o com in�cio pr�ximo.';
          } else {
            $l_imagem = $conImgNormal;
            $l_title = 'Miss�o n�o iniciada.';
          }
        } else {
          if ($l_tramite != 'EE') {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Miss�o em andamento com tramita��o em atraso.';
          } else {
            $l_imagem = $conImgStNormal;
            $l_title = 'Miss�o em andamento.';
          }
        }
      }
      elseif (substr($l_tipo, 0, 2) == 'PE') {
        // Projetos
        if ($l_tramite != 'AT') {
          if ($l_tramite == 'CA') {
            $l_imagem = $conImgCancel;
            $l_title = 'Registro cancelado.';
          }
          elseif ($l_tramite == 'CI') {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgAtraso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
              $l_imagem = $conImgAviso;
              $l_title = 'Execu��o n�o iniciada. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgNormal;
              $l_title = 'Execu��o n�o iniciada. Prazo final dentro do previsto.';
            }
          } else {
            if ($l_fim < addDays(time(), -1)) {
              $l_imagem = $conImgStAtraso;
              $l_title = 'Em execu��o. Fim previsto superado.';
            }
            elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
              $l_imagem = $conImgStAviso;
              $l_title = 'Em execu��o. Fim previsto pr�ximo.';
            } else {
              $l_imagem = $conImgStNormal;
              $l_title = 'Em execu��o. Prazo final dentro do previsto.';
            }
          }
        } else {
          if ($l_fim < Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAtraso;
            $l_title = 'Execu��o conclu�da ap�s a data prevista.';
          }
          elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
            $l_imagem = $conImgOkAcima;
            $l_title = 'Execu��o conclu�da antes da data prevista.';
          } else {
            $l_imagem = $conImgOkNormal;
            $l_title = 'Execu��o conclu�da na data prevista.';
          }
        }
      }

      if ($l_imagem != '') {
        $l_string = '           <img src="' . $l_imagem . '" title="' . $l_title . '" border=0 width=10 heigth=10>';
      }
    }

    return $l_string;
  }

  // =========================================================================
  // Exibe �cone da solicita��o para geo-referenciamento
  // -------------------------------------------------------------------------
  function ExibeIconeSolic($l_tipo, $l_inicio, $l_fim, $l_inicio_real, $l_fim_real, $l_aviso, $l_dias_aviso, $l_tramite, $l_perc, $l_legenda = 0, $l_restricao = null) {
    extract($GLOBALS);
    $l_imagem = '';
    $l_tipo = strtoupper($l_tipo);
    if ($l_tipo == 'ETAPA') {
      // Etapas de projeto
      if ($l_perc < 100) {
        if (nvl($l_inicio_real, '') == '') {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoAtraso;
          }
          elseif (((time() - $l_inicio) / ($l_fim - $l_inicio +1)) * 100 > $l_perc) {
            $l_imagem = $conIcoAviso;
          } else {
            $l_imagem = $conIcoNormal;
          }
        } else {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoStAtraso;
          }
          elseif (((time() - $l_inicio) / ($l_fim - $l_inicio +1)) * 100 > $l_perc) {
            $l_imagem = $conIcoStAviso;
          } else {
            $l_imagem = $conIcoStNormal;
          }
        }
      } else {
        if ($l_fim < Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAtraso;
        }
        elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAcima;
        } else {
          $l_imagem = $conIcoOkNormal;
        }
      }
    }
    elseif (substr($l_tipo, 0, 2) == 'GC') {
      // Contratos e conv�nios
      if ($l_tramite != 'AT' && $l_tramite != 'CR') {
        if ($l_tramite == 'CA') {
          $l_imagem = $conIcoCancel;
        }
        elseif ($l_tramite == 'CI') {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoAtraso;
          }
          elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
            $l_imagem = $conIcoAviso;
          } else {
            $l_imagem = $conIcoNormal;
          }
        }
        elseif ($l_tramite == 'ER') {
          $l_imagem = $conIcoStAcima;
        } else {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoStAtraso;
          }
          elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
            $l_imagem = $conIcoStAviso;
          } else {
            $l_imagem = $conIcoStNormal;
          }
        }
      } else {
        if ($l_fim < Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAtraso;
        }
        elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAcima;
        } else {
          $l_imagem = $conIcoOkNormal;
        }
      }
    }
    elseif (substr($l_tipo, 0, 2) == 'GD') {
      // Tarefas, demandas eventuais e demandas de triagem
      if ($l_tramite != 'AT') {
        if ($l_tramite == 'CA') {
          $l_imagem = $conIcoCancel;
        }
        elseif ($l_tramite == 'CI' || $l_restricao == 'SEMEXECUCAO') {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoAtraso;
          }
          elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
            $l_imagem = $conIcoAviso;
          } else {
            $l_imagem = $conIcoNormal;
          }
        } else {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoStAtraso;
          }
          elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
            $l_imagem = $conIcoStAviso;
          } else {
            $l_imagem = $conIcoStNormal;
          }
        }
      } else {
        if ($l_fim < Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAtraso;
        }
        elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAcima;
        } else {
          $l_imagem = $conIcoOkNormal;
        }
      }
    }
    elseif (substr($l_tipo, 0, 2) == 'FN') {
      // Tarefas e demandas eventuais
      if ($l_tramite != 'AT') {
        if ($l_tramite == 'CA') {
          $l_imagem = $conIcoCancel;
        }
        elseif ($l_tramite == 'CI') {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoAtraso;
          }
          elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
            $l_imagem = $conIcoAviso;
          } else {
            $l_imagem = $conIcoNormal;
          }
        } else {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoStAtraso;
          }
          elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
            $l_imagem = $conIcoStAviso;
          } else {
            $l_imagem = $conIcoStNormal;
          }
        }
      } else {
        if ($l_fim < Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAtraso;
        }
        elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAcima;
        } else {
          $l_imagem = $conIcoOkNormal;
        }
      }
    }
    elseif (substr($l_tipo, 0, 2) == 'SR') {
      // Tarefas
      if ($l_tramite != 'AT') {
        if ($l_tramite == 'CA') {
          $l_imagem = $conIcoCancel;
        }
        elseif ($l_tramite == 'CI') {
          if ($l_fim < time()) {
            $l_imagem = $conIcoAtraso;
          }
          elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
            $l_imagem = $conIcoAviso;
          } else {
            $l_imagem = $conIcoNormal;
          }
        } else {
          if ($l_fim < time()) {
            $l_imagem = $conIcoStAtraso;
          }
          elseif ($l_aviso == 'S' && $l_dias_aviso <= time()) {
            $l_imagem = $conIcoStAviso;
          } else {
            $l_imagem = $conIcoStNormal;
          }
        }
      } else {
        if ($l_fim < Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAtraso;
        }
        elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAcima;
        } else {
          $l_imagem = $conIcoOkNormal;
        }
      }
    }
    elseif (substr($l_tipo, 0, 2) == 'PJ') {
      // Projetos
      if ($l_tramite != 'AT') {
        if ($l_tramite == 'CA') {
          $l_imagem = 'project_red';
        }
        elseif ($l_tramite == 'CI') {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = 'project_red';
          }
          elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
            $l_imagem = 'project_yellow';
          } else {
            $l_imagem = 'project_green';
          }
        } else {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = 'project_red';
          }
          elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
            $l_imagem = 'project_yellow';
          } else {
            $l_imagem = 'project_green';
          }
        }
      } else {
        if ($l_fim < Nvl($l_fim_real, $l_fim)) {
          $l_imagem = 'project_red';
        }
        elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
          $l_imagem = 'project_yellow';
        } else {
          $l_imagem = 'project_green';
        }
      }
    }
    elseif (substr($l_tipo, 0, 2) == 'CL') {
      // Projetos
      if ($l_tramite != 'AT') {
        if ($l_tramite == 'CA') {
          $l_imagem = $conIcoCancel;
        }
        elseif ($l_tramite == 'CI') {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoAtraso;
          }
          elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
            $l_imagem = $conIcoAviso;
          } else {
            $l_imagem = $conIcoNormal;
          }
        } else {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoStAtraso;
          }
          elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
            $l_imagem = $conIcoStAviso;
          } else {
            $l_imagem = $conIcoStNormal;
          }
        }
      } else {
        if ($l_fim < Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAtraso;
        }
        elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAcima;
        } else {
          $l_imagem = $conIcoOkNormal;
        }
      }
    }
    elseif (substr($l_tipo, 0, 2) == 'PD') {
      // Viagens
      if ($l_tramite == 'CA') {
        $l_imagem = $conIcoCancel;
      }
      elseif ($l_fim < addDays(time(), -1)) {
        if ($l_tramite == 'AT') {
          $l_imagem = $conIcoOkNormal;
        }
        elseif ($l_tramite != 'EE') {
          $l_imagem = $conIcoOkAtraso;
        } else {
          $l_imagem = $conIcoOkAcima;
        }
      }
      elseif ($l_inicio > time()) {
        if ($l_dias_aviso <= time()) {
          $l_imagem = $conIcoAviso;
        } else {
          $l_imagem = $conIcoNormal;
        }
      } else {
        if ($l_tramite != 'EE') {
          $l_imagem = $conIcoOkAtraso;
        } else {
          $l_imagem = $conIcoStNormal;
        }
      }
    }
    elseif (substr($l_tipo, 0, 2) == 'PE') {
      // Projetos
      if ($l_tramite != 'AT') {
        if ($l_tramite == 'CA') {
          $l_imagem = $conIcoCancel;
        }
        elseif ($l_tramite == 'CI') {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoAtraso;
          }
          elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
            $l_imagem = $conIcoAviso;
          } else {
            $l_imagem = $conIcoNormal;
          }
        } else {
          if ($l_fim < addDays(time(), -1)) {
            $l_imagem = $conIcoStAtraso;
          }
          elseif ($l_aviso == 'S' && ($l_dias_aviso <= addDays(time(), -1))) {
            $l_imagem = $conIcoStAviso;
          } else {
            $l_imagem = $conIcoStNormal;
          }
        }
      } else {
        if ($l_fim < Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAtraso;
        }
        elseif ($l_fim > Nvl($l_fim_real, $l_fim)) {
          $l_imagem = $conIcoOkAcima;
        } else {
          $l_imagem = $conIcoOkNormal;
        }
      }
    }

    return $l_imagem;
  }
  // =========================================================================
  // Montagem da URL com os par�metros de filtragem quando o for UPLOAD
  // -------------------------------------------------------------------------
  function MontaFiltroUpload($p_Form) {
    extract($GLOBALS);
    $l_string = '';
    foreach ($p_Form as $l_Item) {
      if (substr($l_item, 0, 2) == "p_" && $l_item->value > '') {
        $l_string .= "&" . $l_Item . "=" . $l_item->value;
      }
    }
    return $l_string;
  }

  // =========================================================================
  // Rotina que monta n�mero de ordem da etapa do projeto
  // -------------------------------------------------------------------------
  function MontaOrdemEtapa($l_chave) {
    extract($GLOBALS);
    if (nvl($l_chave, '') == '')
      return null;
    include_once ($w_dir_volta . 'classes/sp/db_getEtapaDataParents.php');
    $l_rs = db_getEtapaDataParents :: getInstanceOf($dbms, $l_chave);
    foreach ($l_rs as $row) {
      $w_texto = f($row, 'ordem');
      break;
    }
    return $w_texto;
  }

  // =========================================================================
  // Rotina que monta o c�digo da especificacao
  // -------------------------------------------------------------------------
  function MontaOrdemEspec($l_chave) {
    extract($GLOBALS);
    $RSQuery = db_getEspecOrdem :: getInstanceOf($dbms, $l_chave);
    $w_texto = '';
    $w_contaux = 0;
    foreach ($RSQuery as $row) {
      $w_contaux = $w_contaux +1;
      if ($w_contaux == 1) {
        $w_texto = f($row, 'codigo') . '.' . $w_texto;
      } else {
        $w_texto = f($row, 'codigo') . '.' . $w_texto;
      }
    }
    return substr($w_texto, 0, strlen($w_texto) - 1);
  }

  // =========================================================================
  // Converte CFLF para <BR>
  // -------------------------------------------------------------------------
  function CRLF2BR($expressao) {
    $result = '';
    if (Nvl($expressao, '') == '') {
      return '';
    } else {
      $result = $expressao;
      if (false !== strpos($result, chr(10) . chr(13)))
        $result = str_replace(chr(10) . chr(13), '<br>', $result);
      if (false !== strpos($result, chr(13) . chr(10)))
        $result = str_replace(chr(13) . chr(10), '<br>', $result);
      if (false !== strpos($result, chr(13)))
        $result = str_replace(chr(13), '<br>', $result);
      if (false !== strpos($result, chr(10)))
        $result = str_replace(chr(10), '<br>', $result);
      //return str_replace('<br><br>','<br>',htmlentities($result)); 
      return str_replace('<br><br>', '<br>', $result);
    }
  }

  // =========================================================================
  // Trata valores nulos
  // -------------------------------------------------------------------------
  function Nvl($expressao, $valor) {
    if ((!isset ($expressao)) || $expressao === '') {
      return $valor;
    } else {
      return $expressao;
    }
  }

  // =========================================================================
  // Trata valores nulos e cadeias vazias
  // -------------------------------------------------------------------------
  function trocaNulo($expressao, $valor = '-') {
    return nvl(trim($expressao), $valor);
  }

  // =========================================================================
  // Retorna valores nulos se chegar cadeia vazia
  // -------------------------------------------------------------------------
  function Tvl($expressao) {
    if (!isset ($expressao) || $expressao === '' || $expressao === false) {
      return null;
    } else {
      return $expressao;
    }
  }

  // =========================================================================
  // Retorna valores nulos se chegar cadeia vazia
  // -------------------------------------------------------------------------
  function Cvl($expressao) {
    if (!isset ($expressao) || $expressao == '') {
      return 0;
    } else {
      return $expressao;
    }
  }

  // =========================================================================
  // Informar os dados de uma variavel mostrando a linha do arquivo em que a fun��o foi chamada.
  // 
  function dbg($var = NULL, $morte = FALSE) {
    print "/*<center>-- ----- ----- ----- ----- ----- ----- ----- ----- ----- \n ";
    print "<strong>DEBUG INICIO</strong>";
    print " ----- ----- ----- ----- ----- ----- ----- ----- ----- --</center> \n ";
    $array = debug_backtrace();
    print ("<center>linha '" . $array[0]['line'] . "' do arquivo '" .
    $array[0]['file'] . "'</center> \n ");
    print ("<pre><font size='3'> \n ");
    var_dump($var);
    print ("</font></pre> \n ");
    print "<center>---- ----- ----- ----- ----- ----- ----- ----- ----- ----- \n ";
    print "<strong>DEBUG FIM</strong>";
    print " ----- ----- ----- ----- ----- ----- ----- ----- ----- ----</center> \n ";

    if ($morte) {
      die("<center><font color='#ff0000'><strong>D I E</strong></font></center>  */ \n ");
    }
  }

  // =========================================================================
  // Retorna o caminho f�sico para o diret�rio  do cliente informado
  // -------------------------------------------------------------------------
  function DiretorioCliente($p_cliente) {
    extract($GLOBALS);
    return $conFilePhysical . $p_cliente;
  }

  // =========================================================================
  // Verifica se um arquivo ou diret�rio existe, se � poss�vel a leitura 
  // e se � poss�vel a escrita
  // -------------------------------------------------------------------------
  function testFile($l_erro, $l_raiz, $l_leitura = false, $l_escrita = false) {
    if (!file_exists($l_raiz)) {
      $l_erro = 'inexistente';
      return false;
    }
    elseif (!is_readable($l_raiz)) {
      $l_erro = 'sem permiss�o de leitura';
      return false;
    }
    elseif (!is_writable($l_raiz)) {
      $l_erro = 'sem permiss�o de escrita';
      return false;
    }
    return true;
  }

  // =========================================================================
  // Montagem de URL a partir da sigla da op��o do menu
  // -------------------------------------------------------------------------
  function MontaURL($p_sigla) {
    extract($GLOBALS);
    include_once ($w_dir_volta . 'classes/sp/db_getLinkData.php');
    $RS_MontaURL = db_getLinkData :: getInstanceOf($dbms, $_SESSION['P_CLIENTE'], $p_sigla);
    $l_ImagemPadrao = 'images/Folder/SheetLittle.gif';
    if (count($RS_MontaURL) <= 0)
      return '';
    else {
      if (nvl(f($RS_MontaURL, 'imagem'), '-') != '-') {
        $l_Imagem = f($RS_MontaURL, 'imagem');
      } else {
        $l_Imagem = $l_ImagemPadrao;
      }
      return f($RS_MontaURL, 'link') . "&P1=" . f($RS_MontaURL, 'p1') . "&P2=" . f($RS_MontaURL, 'p2') . "&P3=" . f($RS_MontaURL, 'p3') . "&P4=" . f($RS_MontaURL, 'p4') . "&TP=<img src=" . $l_Imagem . " BORDER=0>" . f($RS_MontaURL, 'nome') . "&SG=" . f($RS_MontaURL, 'sigla');
    }
  }

  // =========================================================================
  // Montagem de cabe�alho padr�o de formul�rio
  // -------------------------------------------------------------------------
  function AbreForm($p_Name, $p_Action, $p_Method, $p_onSubmit, $p_Target) {
    $l_html = '';
    if (!isset ($p_Target)) {
      $l_html .= '<FORM action="' . $p_Action . '" method="' . $p_Method . '" NAME="' . $p_Name . '" onSubmit="' . $p_onSubmit . '">';
    } else {
      $l_html .= '<FORM action="' . $p_Action . '" method="' . $p_Method . '" NAME="' . $p_Name . '" onSubmit="' . $p_onSubmit . '" target="' . $p_Target . '">';
    }

    ShowHtml($l_html);
  }

  // =========================================================================
  // Montagem de campo do tipo radio com padr�o N�o
  // -------------------------------------------------------------------------
  function MontaRadioNS($label, $chave, $campo, $hint = null, $restricao = null, $atributo = null) {
    extract($GLOBALS);
    ShowHTML('          <td' . ((nvl($hint, '') != '') ? ' TITLE="' . $hint . '"' : '') . '>');
    if (Nvl($label, '') > '') {
      ShowHTML($label . '</b><br>');
    }
    if ($chave == 'S') {
      ShowHTML('              <input ' . $w_Disabled . ' type="radio" name="' . $campo . '" value="S" checked ' . $atributo . '> Sim <input ' . $w_Disabled . ' type="radio" name="' . $campo . '" value="N" ' . $atributo . '> N�o');
    } else {
      ShowHTML('              <input ' . $w_Disabled . ' type="radio" name="' . $campo . '" value="S" ' . $atributo . '> Sim <input ' . $w_Disabled . ' type="radio" name="' . $campo . '" value="N" checked ' . $atributo . '> N�o');
    }
  }

  // =========================================================================
  // Montagem de campo do tipo radio com padr�o Sim
  // -------------------------------------------------------------------------
  function MontaRadioSN($label, $chave, $campo, $hint = null, $restricao = null, $atributo = null) {
    extract($GLOBALS);
    ShowHTML('          <td' . ((nvl($hint, '') != '') ? ' TITLE="' . $hint . '"' : '') . '>');
    if (Nvl($label, '') > '') {
      ShowHTML($label . '</b><br>');
    }
    if ($chave == 'N') {
      ShowHTML('              <input ' . $w_Disabled . ' type="radio" name="' . $campo . '" value="S" ' . $atributo . '> Sim <input ' . $w_Disabled . ' type="radio" name="' . $campo . '" value="N" checked ' . $atributo . '> N�o');
    } else {
      ShowHTML('              <input ' . $w_Disabled . ' type="radio" name="' . $campo . '" value="S" checked ' . $atributo . '> Sim <input ' . $w_Disabled . ' type="radio" name="' . $campo . '" value="N" ' . $atributo . '> N�o');
    }
  }

  // =========================================================================
  // Retorna a prioridade a partir do c�digo
  // -------------------------------------------------------------------------
  function RetornaPrioridade($p_chave) {
    switch (Nvl($p_chave, 999)) {
      case 0 :
        return 'Alta';
        break;
      case 1 :
        return 'M�dia';
        break;
      case 2 :
        return 'Normal';
        break;
      default :
        return '---';
        break;
    }
  }

  // =========================================================================
  // Retorna o tipo de visao a partir do c�digo
  // -------------------------------------------------------------------------
  function RetornaTipoVisao($p_chave) {
    switch ($p_chave) {
      case 0 :
        return 'Completa';
        break;
      case 1 :
        return 'Parcial';
        break;
      case 2 :
        return 'Resumida';
        break;
      default :
        return 'Erro';
        break;
    }
  }

  // =========================================================================
  // Fun��o que formata dias, horas, minutos e segundos a partir dos segundos
  // -------------------------------------------------------------------------
  function FormataTempo($p_segundos) {
    $l_horas = intval($p_segundos / 3600);
    $l_minutos = intval(($p_segundos - ($l_horas * 3600)) / 60);
    $l_segundos = $p_segundos - ($l_horas * 3600) - ($l_minutos * 60);
    return substr(1000 + $l_horas, 1, 3) . ":" . substr(100 + $l_minutos, 1, 2) . ":" . substr(100 + $l_segundos, 1, 2);
  }

  // =========================================================================
  // Fun��o que formata valores com separadores de milhar e decimais
  // -------------------------------------------------------------------------
  function FormatNumber($p_valor, $p_decimais = 2) {
    return number_format($p_valor, $p_decimais, ',', '.');
  }

  // =========================================================================
  // Fun��o que retorna o c�digo de tarifa��o telef�nica do usu�rio logado
  // -------------------------------------------------------------------------
  function RetornaUsuarioCentral() {
    extract($GLOBALS);
    // Se receber o c�digo do usuario do SIW, o usu�rio ser� determinado por par�metro;
    // caso contr�rio, retornar� o c�digo do usu�rio logado.
    if ($_REQUEST['w_sq_usuario_central'] > '') {
      return $_REQUEST['w_sq_usuario_central'];
    } else {
      $RS = db_getPersonData :: getInstanceOf($dbms, $w_cliente, $w_usuario, null, null);
      return f($RS, 'sq_usuario_central');
    }
  }

  // =========================================================================
  // Fun��o que retorna o c�digo do usu�rio logado
  // -------------------------------------------------------------------------
  function RetornaUsuario() {
    extract($GLOBALS);
    // Se receber o c�digo do usuario do SIW, o usu�rio ser� determinado por par�metro;
    // caso contr�rio, retornar� o c�digo do usu�rio logado.
    if ($_REQUEST['w_usuario'] > '') {
      return $_REQUEST['w_usuario'];
    } else {
      return $_SESSION['SQ_PESSOA'];
    }
  }

  // =========================================================================
  // Fun��o que retorna o ano a ser utilizado para recupera��o de dados
  // -------------------------------------------------------------------------
  function RetornaAno() {
    extract($GLOBALS);
    if ($_REQUEST['w_ano'] > '')
      return $_REQUEST['w_ano'];
    elseif (nvl($AL, '') != '') {
      $q = "SELECT coalesce(max(ano_letivo),to_number(to_char(sysdate,'yyyy'))) ano_letivo from sbpi.Aluno_Turma WHERE sq_aluno = " . $AL;
      $l_rs = db_exec :: getInstanceOf($dbms, $q, & $numRows);
      if (count($l_rs) > 0)
        foreach ($l_rs as $row)
          return f($row, 'ano_letivo');
      else
        return Date('Y');
    } else
      return Date('Y');
  }

  // =========================================================================
  // Fun��o que retorna o c�digo do menu
  // -------------------------------------------------------------------------
  function RetornaMenu($p_cliente, $p_sigla) {
    extract($GLOBALS);
    // Se receber o c�digo do menu do SIW, o c�digo ser� determinado por par�metro;
    // caso contr�rio, retornar� o c�digo retornado a partir da sigla.
    if ($_REQUEST['w_menu'] > '') {
      return $_REQUEST['w_menu'];
    } else {
      include_once ($w_dir_volta . 'classes/sp/db_getMenuCode.php');
      $l_RS = db_getMenuCode :: getInstanceOf($dbms, $p_cliente, $p_sigla);
      foreach ($l_RS as $l_row) {
        if (f($l_row, 'ativo') == 'S') {
          return f($l_row, 'sq_menu');
          break;
        }
      }
      return null;
    }
  }

  // =========================================================================
  // Fun��o que retorna o c�digo do cliente
  // -------------------------------------------------------------------------
  function RetornaCliente() {
    extract($GLOBALS);
    include_once ($w_dir_volta . 'classes/sp/db_getCompanyData.php');
    // Se receber o c�digo do cliente do SIW, o cliente ser� determinado por par�metro;
    // caso contr�rio, o cliente ser� a empresa ao qual o usu�rio logado est� vinculado.
    if ($_REQUEST['w_cgccpf'] > '' && strlen($_REQUEST['w_cgccpf']) > 11) {
      $RS = db_getCompanyData :: getInstanceOf($dbms, $_SESSION['P_CLIENTE'], $_REQUEST['w_cgccpf']);
      if (count($RS) > 0) {
        return f($RS, 'sq_pessoa');
      } else {
        return $_SESSION['P_CLIENTE'];
      }
    }
    elseif ($_REQUEST['w_cliente'] > '') {
      return $_REQUEST['w_cliente'];
    } else {
      return $_SESSION['P_CLIENTE'];
    }
  }

  // =========================================================================
  // Fun��o que retorna S/N indicando se o usu�rio informado � gestor do sistema
  // ou do m�dulo ao qual a solicita��o pertence
  // -------------------------------------------------------------------------
  function RetornaGestor($p_solicitacao, $p_usuario) {
    extract($GLOBALS);
    include_once ($w_dir_volta . 'classes/sp/db_getGestor.php');
    $l_acesso = db_getGestor :: getInstanceOf($dbms, $p_solicitacao, $p_usuario);
    return $l_acesso;
  }

  // =========================================================================
  // Fun��o que retorna valor maior que 0 se o usu�rio informado tem acesso �
  // op��o e tr�mite indicados
  // -------------------------------------------------------------------------
  function RetornaMarcado($p_menu, $p_usuario, $p_endereco, $p_tramite) {
    extract($GLOBALS);
    include_once ($w_dir_volta . 'classes/sp/db_getMarcado.php');
    $l_acesso = db_getMarcado :: getInstanceOf($dbms, $p_menu, $p_usuario, $p_endereco, $p_tramite);
    return $l_acesso;
  }

  // =========================================================================
  // Fun��o que retorna S/N indicando se o usu�rio informado � gestor do m�dulo 
  // a op��o do menu pertence
  // -------------------------------------------------------------------------
  function RetornaModMaster($p_cliente, $p_usuario, $p_menu) {
    include_once ($w_dir_volta . 'classes/sp/db_getModMaster.php');
    extract($GLOBALS);
    $l_RS = db_getModMaster :: getInstanceOf($dbms, $p_cliente, $p_usuario, $p_menu);
    if (count($l_RS) == 0) {
      return 'N';
    } else {
      foreach ($l_RS as $l_row) {
        $l_RS = $l_row;
        break;
      }
      return f($l_RS, 'gestor_modulo');
    }
  }

  // =========================================================================
  // Rotina que encerra a sess�o e fecha a janela do SIW
  // -------------------------------------------------------------------------
  function EncerraSessao() {
    extract($GLOBALS);
    ScriptOpen('JavaScript');
    ShowHTML(' alert("Tempo m�ximo de inatividade atingido! Autentique-se novamente."); ');
    ShowHTML(' top.location.href=\'' . $conDefaultPath . '\';');
    ScriptClose();
    exit ();
  }

  // =========================================================================
  // Montagem da URL com os dados de um assunto da tabela de temporalidade
  // -------------------------------------------------------------------------
  function ExibeAssunto($p_dir, $p_cliente, $p_nome, $p_chave, $p_tp) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'l_');
    if (Nvl($p_chave, '') == '') {
      $l_string = '---';
    } else {
      $l_string .= '<A class="hl" HREF="javascript:this.status.value;" onClick="window.open(\'' . $conRootSIW . 'mod_pa/tabelas.php?par=TELAASSUNTO&w_cliente=' . $p_cliente . '&w_chave=' . $p_chave . '&P1=' . $P1 . '&P2=' . $P2 . '&P3=' . $P3 . '&P4=' . $P4 . '&TP=' . $p_tp . '&SG=' . '\',\'TelaAssunto\',\'width=785,height=570,top=10,left=10,toolbar=no,scrollbars=yes,resizable=yes,status=no\'); return false;">' . $p_nome . '</A>';
    }
    return $l_string;
  }

  // =========================================================================
  // Fun��o que formata um texto para exibi��o em HTML
  // -------------------------------------------------------------------------
  function ExibeTexto($p_texto) {
    return str_replace('  ', '&nbsp;&nbsp;', str_replace('\r\n', '<br>', $p_texto));
  }

  // =========================================================================
  // Fun��o que retorna a data/hora do banco
  // -------------------------------------------------------------------------
  function DataHora() {
    return diaSemana(date('l, d/m/Y, H:i:s'));
  }

  // =========================================================================
  // Fun��o que retorna um timestamp da string informada
  // date: string contendo a data no formato DD/MM/YYYY, HH24:MI:SS
  // -------------------------------------------------------------------------
  function toDate($date) {
    if (strlen($date) != 20 && strlen($date) != 10)
      return nil;
    else {
      $l_date = $date;
      $l_temp = substr($date, 6, 4) . substr($date, 3, 2) . substr($date, 0, 2) . '000000';
      if ($l_temp < 19011213204554)
        $l_date = '13/12/1901, 20:45:54';
      elseif ($l_temp > 20381901031407) $l_date = '01/19/2038, 03:14:07';
      if (strlen($l_date) == 10)
        return mktime(0, 0, 0, substr($l_date, 3, 2), substr($l_date, 0, 2), substr($l_date, 6, 4));
      else {
        return mktime(substr($l_date, 12, 2), substr($l_date, 15, 2), substr($l_date, 18, 2), substr($l_date, 3, 2), substr($l_date, 0, 2), substr($l_date, 6, 4));
      }
    }
  }

  // =========================================================================
  // Fun��o que retorna um timestamp da string informada para o formato SQL Server
  // date: string contendo a data no formato DD/MM/YYYY, HH24:MI:SS
  // -------------------------------------------------------------------------
  function toSQLDate($date) {
    if (nvl($date, '') == '')
      return '';
    if (strlen($date) == 20)
      return substr($date, 6, 4) . '-' . substr($date, 3, 2) . '-' . substr($date, 0, 2) . ' ' . substr($l_date, 12);
    else
      return substr($date, 6, 4) . '-' . substr($date, 3, 2) . '-' . substr($date, 0, 2);
  }

  // =========================================================================
  // Fun��o que retorna um valor da string informada
  // valor: string contendo o valor
  // -------------------------------------------------------------------------
  function toNumber($valor) {
    if ($_SESSION['DBMS'] == 2)
      return str_replace(',', '.', str_replace('.', '', $valor));
    else
      return str_replace('.', '', $valor);
  }

  // =========================================================================
  // Fun��o que retorna uma data como string para manipula��o em formul�rios
  // -------------------------------------------------------------------------
  function FormatDateTime($date) {
    if (strlen($date) != 8 && strlen($date) != 10 && strlen($date) != 19)
      return null;
    else {
      if (strlen($date) == 10)
        return $date;
      elseif (strlen($date) == 19) {
        $temp = substr($date, 0, 10);
        $l_date = implode('/', array_reverse(explode('-', $temp)));
      } else {
        $l_date = substr($date, 0, 2) . '/' . substr($date, 3, 2) . '/';
        if (substr($date, 6, 2) < 30)
          $l_date = $l_date . '20' . substr($date, 6, 2);
        else
          $l_date = $l_date . '19' . substr($date, 6, 2);
      }
      return $l_date;
    }
  }

  // =========================================================================
  // Fun��o que adiciona dias a uma data
  // date: timestamp gerado a partir da fun�ao toDate()
  // inc:  inteiro precedido do sinal de adi��o ou subtra��o de dias (+1, -3 etc.)
  // -------------------------------------------------------------------------
  function addDays($date, $inc) {
    return mktime(date(H, $date), date(i, $date), date(s, $date), date(m, $date), date(d, $date) + $inc, date(Y, $date));
  }

  // =========================================================================
  // Fun��o que transforma um array numa lista separada por v�rgulas
  // array: array de entrada
  // -------------------------------------------------------------------------
  function explodeArray($array) {
    if (is_array($array)) {
      $lista = '';
      foreach ($array as $key => $val)
        $lista = $lista . ',' . trim($val);
      return substr($lista, 1, strlen($lista) + 1);
    } else {
      return $array;
    }
  }

  // =========================================================================
  // Rotina que monta a m�scara do benefici�rio
  // -------------------------------------------------------------------------
  function MascaraBeneficiario($cgccpf) {
    // Se o campo tiver m�scara, retira
    if ((strpos($cgccpf, '.') ? strpos($cgccpf, '.') + 1 : 0) > 0) {
      return str_replace('/', '', str_replace('-', '', str_replace('.', '', $cgccpf)));
    } // Caso contr�rio, aplica a m�scara, dependendo do tamanho do par�metro
    elseif (strlen($cgccpf) == 11) {
      return substr($cgccpf, 0, 3) . '.' . substr($cgccpf, 3, 3) . '.' . substr($cgccpf, 6, 3) . '-' . substr($cgccpf, 9, 2);
    }
    elseif (strlen($cgccpf) == 14) {
      return substr($cgccpf, 0, 2) . '.' . substr($cgccpf, 2, 3) . '.' . substr($cgccpf, 5, 3) . '/' . substr($cgccpf, 8, 4) . '-' . substr($cgccpf, 12, 2);
    }
  }

  // =========================================================================
  // Rotina de envio de e-mail
  // -------------------------------------------------------------------------
  function EnviaMail($w_subject, $w_mensagem, $w_recipients, $w_attachments = null) {
    extract($GLOBALS);

    include_once ($w_dir_volta . 'classes/mail/email_message.php');
    include_once ($w_dir_volta . 'classes/mail/smtp_message.php');
    include_once ($w_dir_volta . 'classes/mail/smtp.php');
    include_once ($w_dir_volta . 'classes/sp/db_getCustomerData.php');

    // Recupera informa��es para configurar o remetente da mensagem e o servi�o de entrega
    $RS_Cliente = db_getCustomerData :: getInstanceOf($dbms, $_SESSION['P_CLIENTE']);

    $subject = $w_subject;
    $from_name = f($RS_Cliente, 'siw_email_nome');
    $from_address = f($RS_Cliente, 'siw_email_conta');
    $reply_name = $from_name;
    $reply_address = $from_address;
    $error_delivery_name = $from_name;
    $error_delivery_address = $from_address;

    $email_message = new smtp_message_class;

    $email_message->localhost = $_SERVER['HOSTNAME'];
    $email_message->smtp_host = f($RS_Cliente, 'smtp_server');
    if (strpos($email_message->smtp_host, ':') !== false) {
      list ($email_message->smtp_host, $email_message->smtp_port) = explode(':', $email_message->smtp_host);
    }
    $email_message->smtp_ssl = ((strpos(f($RS_Cliente, 'smtp_server'), 'gmail') === false) ? 0 : 1); /* Use SSL to connect to the SMTP server. Gmail requires SSL */
    $email_message->smtp_direct_delivery = 0; /* Deliver directly to the recipients destination SMTP server */
    $email_message->smtp_user = ((nvl(f($RS_Cliente, 'siw_email_senha'), 'nulo') == 'nulo') ? '' : f($RS_Cliente, 'siw_email_conta')); /* authentication user name */
    $email_message->smtp_realm = ''; /* authentication realm or Windows domain when using NTLM authentication */
    $email_message->smtp_workstation = ''; /* authentication workstation name when using NTLM authentication */
    $email_message->smtp_password = ((nvl(f($RS_Cliente, 'siw_email_senha'), 'nulo') == 'nulo') ? '' : f($RS_Cliente, 'siw_email_senha')); /* authentication password */
    $email_message->smtp_debug = 0; /* Output dialog with SMTP server */
    $email_message->smtp_html_debug = 0; /* set this to 1 to make the debug output appear in HTML */

    /* if you need POP3 authetntication before SMTP delivery,
    * specify the host name here. The smtp_user and smtp_password above
    * should set to the POP3 user and password*/
    $email_message->smtp_pop3_auth_host = ((strpos(f($RS_Cliente, 'smtp_server'), 'gmail') === false) ? '' : $email_message->smtp_host);
    $email_message->pop3_auth_port = ((strpos(f($RS_Cliente, 'smtp_server'), 'gmail') === false) ? 110 : 995);

    /* In directly deliver mode, the DNS may return the IP of a sub-domain of
    * the default domain for domains that do not exist. If that is your
    * case, set this variable with that sub-domain address. */
    $email_message->smtp_exclude_address = "";

    /* If you use the direct delivery mode and the GetMXRR is not functional,
    * you need to use a replacement function. */
    /*
    $_NAMESERVERS=array();
    include("rrcompat.php");
    $email_message->smtp_getmxrr="_getmxrr";
    */

    if (strpos($w_recipients, ';') === false)
      $w_recipients .= ';';
    $l_recipients = explode(';', $w_recipients);
    $l_cont = 0;
    foreach ($l_recipients as $k => $v) {
      if (nvl($v, '') != '') {
        if (strpos($v, '|') !== false) {
          $rec = explode('|', $v);
          $rec_address = trim($rec[0]);
          $rec_name = trim($rec[1]);
        } else {
          $rec_address = trim($v);
          $rec_name = trim($v);
        }
        // Efvita a repeti��o de nomes
        if (count($l_dest[$rec_address]) == 0) {
          $l_dest[$rec_address] = $rec_name;
          $l_cont++;
        }
      }
    }
    $i = 0;
    if (is_array($l_dest)) {
      foreach ($l_dest as $k => $v) {
        if ($i == 0) {
          // O primeiro destinat�rio ser� colocado como "To"
          $email_message->SetEncodedEmailHeader("To", $k, $v);
          unset ($l_dest[$k]);
        }
        elseif ($i == 1 && $l_cont == 2) {
          // Se s� tiver mais um destinat�rio, coloca header �nico
          $email_message->SetEncodedEmailHeader("Cc", $k, $v);
          break;
        } else {
          // Se tiver mais um destinat�rio, al�m do principal, coloca headers m�ltiplos
          $email_message->SetMultipleEncodedEmailHeader("Cc", $l_dest);
          break;
        }
        $i++;
      }
    }
    if (is_array($w_attachments)) {
      foreach ($w_attachments as $l_attach)
        $email_message->AddFilePart($l_attach);
    }
    $email_message->SetEncodedEmailHeader('From', $from_address, $from_name);
    $email_message->SetEncodedEmailHeader("Reply-To", $reply_address, $reply_name);
    // Set the Return-Path header to define the envelope sender address to which bounced messages are delivered.
    // If you are using Windows, you need to use the smtp_message_class to set the return-path address.
    if (defined("PHP_OS") && strcmp(substr(PHP_OS, 0, 3), "WIN"))
      $email_message->SetHeader("Return-Path", $error_delivery_address);
    $email_message->SetEncodedEmailHeader('Errors-To', 'desenv@sbpi.com.br', 'SBPI Suporte');
    $email_message->SetEncodedHeader("Subject", $subject);
    $email_message->AddQuotedPrintableHTMLPart($w_mensagem, '', $html_part);

    /*
    // It is strongly recommended that when you send HTML messages,
    // also provide an alternative text version of HTML page,
    // even if it is just to say that the message is in HTML,
    // because more and more people tend to delete HTML only
    // messages assuming that HTML messages are spam.
    $text_message='Esta � uma mensagem no formato HTML. Favor usar um programa capaz de ler mensagens nesse formato';
    $email_message->CreateQuotedPrintableTextPart($email_message->WrapText($text_message),'',$text_part);
    
    // The complete HTML parts are gathered in a single multipart/related part.
    $related_parts=array(
    $html_part,
    $image_part,
    $background_image_part
    );
    $email_message->CreateRelatedMultipart($related_parts,$html_parts);
    
    // Multiple alternative parts are gathered in multipart/alternative parts.
    // It is important that the fanciest part, in this case the HTML part,
    // is specified as the last part because that is the way that HTML capable
    // mail programs will show that part and not the text version part.
    $alternative_parts=array(
    $text_part,
    $html_parts
    );
    $email_message->AddAlternativeMultipart($alternative_parts);
    */

    //send your e-mail
    if ($conEnviaMail) {
      $error = $email_message->Send();
      if (strcmp($error, '')) {
        // Solaris (SunOS) sempre retorna falso, mesmo enviando a mensagem.
        if (strtoupper(PHP_OS) != 'SUNOS') {
          enviaSyslog('RI', 'RECURSO INDISPON�VEL', 'SMTP [' . $email_message->smtp_host . '] Porta [' . $email_message->smtp_port . '] Conta [' . $email_message->smtp_user . '/' . $email_message->smtp_password . ']');
          return 'ERRO: ocorreu algum erro no envio da mensagem.\\SMTP [' . $email_message->smtp_host . ']\nPorta [' . $email_message->smtp_port . ']\nConta [' . $email_message->smtp_user . '/' . $email_message->smtp_password . ']\n' . $error;
        } else {
          return null;
        }
      } else {
        return null;
      }
    }
  }

  // =========================================================================
  // Rotina de registro de mensagem em servidor syslog
  // -------------------------------------------------------------------------
  function enviaSyslog($tipo, $objeto, $mensagem) {
    if (function_exists('fsockopen')) {
      extract($GLOBALS);

      include_once ($w_dir_volta . 'classes/sp/db_getCustomerData.php');
      include_once ($w_dir_volta . 'classes/syslog/syslog.php');

      // Recupera informa��es para configurar o remetente da mensagem e o servi�o de entrega
      $RS_Cliente = db_getCustomerData :: getInstanceOf($dbms, $_SESSION['P_CLIENTE']);

      if (nvl(f($RS_Cliente, 'syslog_server_name'), '') != '' && (($tipo == 'LV' && nvl(f($RS_Cliente, 'syslog_level_pass_ok'), '') != '') || ($tipo == 'LI' && nvl(f($RS_Cliente, 'syslog_level_pass_er'), '') != '') || ($tipo == 'AI' && nvl(f($RS_Cliente, 'syslog_level_sign_er'), '') != '') || ($tipo == 'RI' && nvl(f($RS_Cliente, 'syslog_level_res_er'), '') != '') || ($tipo == 'GV' && nvl(f($RS_Cliente, 'syslog_level_write_ok'), '') != '') || ($tipo == 'GI' && nvl(f($RS_Cliente, 'syslog_level_write_er'), '') != ''))) {
        switch ($tipo) {
          case 'LV' :
            $severity = f($RS_Cliente, 'syslog_level_pass_ok');
            break;
          case 'LI' :
            $severity = f($RS_Cliente, 'syslog_level_pass_er');
            break;
          case 'AI' :
            $severity = f($RS_Cliente, 'syslog_level_sign_er');
            break;
          case 'RI' :
            $severity = f($RS_Cliente, 'syslog_level_res_er');
            break;
          case 'GV' :
            $severity = f($RS_Cliente, 'syslog_level_write_ok');
            break;
          case 'GI' :
            $severity = f($RS_Cliente, 'syslog_level_write_er');
            break;
        }
        $syslog = new Syslog();
        $syslog->SetProtocol(f($RS_Cliente, 'syslog_server_protocol'));
        $syslog->SetPort(f($RS_Cliente, 'syslog_server_port'));
        $syslog->SetFacility(f($RS_Cliente, 'syslog_facility'));
        $syslog->SetFqdn(f($RS_Cliente, 'syslog_fqdn'));
        $syslog->SetSeverity($severity);
        $syslog->SetHostname($_SERVER["REMOTE_ADDR"]);
        $syslog->SetIpFrom($_SERVER["REMOTE_ADDR"]);
        $prefixo = explode(' ', f($RS_Cliente, 'siw_email_nome'));
        $syslog->SetProcess($prefixo[0] . ' - ' . $objeto);

        //envia o registro para o servidor
        $error = $syslog->Send(f($RS_Cliente, 'syslog_server_name'), $mensagem, f($RS_Cliente, 'syslog_timeout'));
        if (strcmp($error, '')) {
          return 'ERRO: ocorreu algum erro no envio da mensagem.\nSYSLOG [' . f($RS_Cliente, 'syslog_server_name') . ']\nProtocolo [' . f($RS_Cliente, 'syslog_server_protocol') . ']\nPorta [' . f($RS_Cliente, 'syslog_server_port') . ']\n' . $error;
        } else {
          return null;
        }
      }
    } else {
      return null;
    }
  }

  // =========================================================================
  // Rotina que extrai a �ltima parte da vari�vel TP
  // -------------------------------------------------------------------------
  function RemoveTP($TP) {
    $w_TP = $TP;
    while (!(strpos($w_TP, '-') === false)) {
      $w_TP = substr($w_TP, strpos($w_TP, '-') + 1, strlen($w_TP));
    }
    return str_replace(' -' . $w_TP, '', $TP);
  }

  // =========================================================================
  // Rotina que extrai o nome de um arquivo, removendo o caminho
  // -------------------------------------------------------------------------
  function ExtractFileName($arquivo) {
    extract($GLOBALS);
    $fsa = $arquivo;
    while ((strpos($fsa, "\\") ? strpos($fsa, "\\") + 1 : 0) > 0) {
      $fsa = substr($fsa, (strpos($fsa, "\\") ? strpos($fsa, "\\") + 1 : 0) + 1 - 1, strlen($fsa));
    }
    while ((strpos($fsa, "/") ? strpos($fsa, "/") + 1 : 0) > 0) {
      $fsa = substr($fsa, (strpos($fsa, "/") ? strpos($fsa, "/") + 1 : 0) + 1 - 1, strlen($fsa));
    }
    return $fsa;
  }

  // =========================================================================
  // Rotina de dele��o de arquivos em disco
  // -------------------------------------------------------------------------
  function DeleteAFile($filespec) {
    extract($GLOBALS);

    $fso = $CreateObject['Scripting.FileSystemObject'];
    $fso->DeleteFile($filespec);
    return $function_ret;
  }

  // =========================================================================
  // Rotina de tratamento de erros
  // -------------------------------------------------------------------------
  function TrataErro($sp, $Err, $params, $file, $line, $object) {
    extract($GLOBALS);
    if (!(strpos($Err['message'], 'ORA-02292') === false) || !(strpos($Err['message'], 'ORA-02292') === false)) {
      // REGISTRO TEM FILHOS
      ScriptOpen('JavaScript');
      ShowHTML(' alert("Existem registros vinculados ao que voc� est� excluindo. Exclua-os primeiro.\\n\\n' . substr($Err['message'], 0, (strpos($Err['message'], chr(10)) ? strpos($Err['message'], chr(10)) + 1 : 0) - 1) . '");');
      ShowHTML(' history.back(1);');
      ScriptClose();
    }
    //elseif (!(strpos($Err['message'],'ORA-02291')===false) || !(strpos($Err['message'],'ORA-02291')===false)) {
    // REGISTRO N�O ENCONTRADO
    //   ScriptOpen('JavaScript');
    //   ShowHTML(' alert("Registro n�o encontrado.");');
    //   ShowHTML(' history.back(1);');
    //   ScriptClose();
    // }
    elseif (!(strpos($Err['message'], 'ORA-0000x1') === false)) {
      // REGISTRO J� EXISTENTE
      ScriptOpen('JavaScript');
      ShowHTML(' alert("Um dos campos digitados j� existe no banco de dados e � �nico.\\n\\n' . substr($Err['message'], 0, (strpos($Err['message'], chr(10)) ? strpos($Err['message'], chr(10)) + 1 : 0) - 1) . '");');
      ShowHTML(' history.back(1);');
      ScriptClose();
    }
    elseif (!(strpos($Err['message'], 'ORA-03113') === false) || !(strpos($Err['message'], 'ORA-03113') === false) || !(strpos($Err['message'], 'ORA-03114') === false) || !(strpos($Err['message'], 'ORA-03114') === false) || !(strpos($Err['message'], 'ORA-12224') === false) || !(strpos($Err['message'], 'ORA-12224') === false) || !(strpos($Err['message'], 'ORA-12514') === false) || !(strpos($Err['message'], 'ORA-12514') === false) || !(strpos($Err['message'], 'ORA-12541') === false) || !(strpos($Err['message'], 'ORA-12541') === false) || !(strpos($Err['message'], 'ORA-12545') === false) || !(strpos($Err['message'], 'ORA-24327') === false) || !(strpos($Err['message'], 'ORA-12545') === false)) {

      ScriptOpen('JavaScript');
      ShowHTML(' alert("Banco de dados fora do ar. Aguarde alguns instantes e tente novamente!");');
      ShowHTML(' history.back(1);');
      ScriptClose();
    } else {
      $w_html = '<html>';
      $w_html .= chr(10) . '<head>';
      $w_html .= chr(10) . '  <BASEFONT FACE="Arial" SIZE="2">';
      $w_html .= chr(10) . '</head>';
      $w_html .= chr(10) . '<body BGCOLOR="#FF5555">';
      $w_html .= chr(10) . '<CENTER><H2>ATEN��O</H2></CENTER>';
      $w_html .= chr(10) . '<BLOCKQUOTE>';
      $w_html .= chr(10) . '<P ALIGN="JUSTIFY">Erro n�o previsto. <b>Uma c�pia desta tela foi enviada por e-mail para os respons�veis pela corre��o. Favor tentar novamente mais tarde.</P>';
      $w_html .= chr(10) . '<TABLE BORDER="2" BGCOLOR="#FFCCCC" CELLPADDING="5"><TR><TD><FONT COLOR="#000000">';
      $w_html .= chr(10) . '<DL><DT>Data e hora da ocorr�ncia: <FONT FACE="courier">' . date('d/m/Y, h:i:s') . '<br><br></font></DT>';
      $w_html .= chr(10) . '<DT>Descri��o:<DD><FONT FACE="courier">' . crlf2br($Err['message']) . '<br><br></font>';
      $w_html .= chr(10) . '<DT>Arquivo:<DD><FONT FACE="courier">' . $file . ', linha: ' . $line . '<br><br></font>';
      //$w_html .= chr(10).'<DT>Objeto:<DD><FONT FACE="courier">'.$object.'<br><br></font>';

      $w_html .= chr(10) . '<DT>Comando em execu��o:<blockquote> <FONT FACE="courier">' . nvl($Err['sqltext'], 'nenhum') . '</blockquote></font></DT>';
      if (is_array($params)) {
        $w_html .= "<DT>Valores dos par�metros:<DD><FONT FACE=\"courier\" size=1>";
        foreach ($params as $w_chave => $w_valor) {
          if ($_SESSION['DBMS'] == 2 && $w_valor[1] == B_DATE) {
            $w_html .= chr(10) . $w_chave . ' [' . toSQLDate($w_valor[0]) . ']<br>';
          } else {
            $w_html .= chr(10) . $w_chave . ' [' . $w_valor[0] . ']<br>';
          }
        }
      }
      $w_html .= "   <br><br></font>";

      $w_html .= chr(10) . '<DT>Vari�veis de servidor:<DD><FONT FACE="courier" size=1>';
      $w_html .= chr(10) . ' SCRIPT_NAME => [' . $_SERVER['SCRIPT_NAME'] . ']<br>';
      $w_html .= chr(10) . ' SERVER_NAME => [' . $_SERVER['SERVER_NAME'] . ']<br>';
      $w_html .= chr(10) . ' SERVER_PORT => [' . $_SERVER['SERVER_PORT'] . ']<br>';
      $w_html .= chr(10) . ' SERVER_PROTOCOL => [' . $_SERVER['SERVER_PROTOCOL'] . ']<br>';
      $w_html .= chr(10) . ' HTTP_ACCEPT_LANGUAGE => [' . $_SERVER['HTTP_ACCEPT_LANGUAGE'] . ']<br>';
      $w_html .= chr(10) . ' HTTP_USER_AGENT => [' . $_SERVER['HTTP_USER_AGENT'] . ']<br>';
      $w_html .= chr(10) . '</DT>';
      $w_html .= chr(10) . '   <br><br></font>';

      $w_html .= chr(10) . '<DT>Dados da querystring:';
      foreach ($_GET as $chv => $vlr) {
        $w_html .= chr(10) . '<DD><FONT FACE="courier" size=1>' . $chv . ' => [' . $vlr . ']<br>';
      }

      $w_html .= chr(10) . '</DT>';
      $w_html .= chr(10) . '<DT>Dados do formul�rio:';
      foreach ($_POST as $chv => $vlr) {
        if (strtolower($chv) != 'w_assinatura')
          $w_html .= chr(10) . '<DD><FONT FACE="courier" size=1>' . $chv . ' => [' . $vlr . ']<br>';
      }

      $w_html .= chr(10) . '</DT>';
      $w_html .= chr(10) . '   <br><br></font>';
      $w_html .= chr(10) . '</DT>';
      $w_html .= chr(10) . '<DT>Vari�veis de sess�o:<DD><FONT FACE="courier" size=1>';
      foreach ($_SESSION as $chv => $vlr) {
        if (strpos(strtoupper($chv), 'SENHA') !== true) {
          $w_html .= chr(10) . $chv . ' => [' . $vlr . ']<br>';
        }
      }
      $w_html .= chr(10) . '</DT>';
      $w_html .= chr(10) . '   <br><br></font>';
      $w_html .= chr(10) . '</FONT></TD></TR></TABLE><BLOCKQUOTE>';
      $w_html .= '</body></html>';

      ShowHTML($w_html);

      if ($conEnviaMail)
        $w_resultado = EnviaMail('ERRO ' . $conSgSistema, $w_html, 'desenv@sbpi.com.br');
      if ($w_resultado > '') {
        ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT">');
        ShowHTML('   alert("N�o foi poss�vel enviar o e-mail comunicando sobre o erro. Favor copiar esta p�gina e envi�-la por e-mail aos gestores do sistema.");');
        ShowHTML('</SCRIPT>');
      }
      die();
    }
    exit;
  }
  // =========================================================================
  // Fim da rotina de tratamento de erros
  // -------------------------------------------------------------------------

  // =========================================================================
  // Rotina de cabe�alho
  // -------------------------------------------------------------------------
  function Cabecalho() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">');
      ShowHTML('<HTML xmlns="http://www.w3.org/1999/xhtml">');
    } else {
      ShowHTML('<HTML>');
    }
  }

  // =========================================================================
  // Rotina de rodap�
  // -------------------------------------------------------------------------
  function Rodape() {
    extract($GLOBALS);
    ShowHTML('<script language="javascript" type="text/javascript" src="js/tooltip.js"></script>');
    ShowHTML('</BODY>');
    ShowHTML('</HTML>');

  }

  // =========================================================================
  // Rotina de rodap� para PDF
  // -------------------------------------------------------------------------
  function RodapePDF() {
    extract($GLOBALS);
    $p_orientation = $_GET["orientacao"];
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('</center>');
      ShowHTML('<center>');
      ShowHTML('<DIV id=rodape>');
      ShowHTML('  <DIV id=endereco>');
      ShowHTML('    <P>Setor Comercial Sul, Ed. Denasa - Salas 901/902 - Bras�lia-DF <BR>Tel : (61) 225 6302 (61) 321 8938 | Fax (61) 225 7599| email: <A href="mailto:pbf@cespe.unb.br">bresil2005@minc.gov.br</A>');
      ShowHTML('    </P>');
      ShowHTML('  </DIV>');
      ShowHTML('</DIV>');
      ShowHTML('</center>');
    } else {
      ShowHTML('<HR>');
    }
    ShowHTML('</BODY>');
    ShowHTML('</HTML>');

    $shtml = ob_get_contents();
    ob_end_clean();

    //  $shtml = str_replace("msword",'', $shtml);	
    $shtml = str_replace("'", '"', $shtml);
    $shtml = str_replace('"', "'", $shtml);

    //echo $shtml;
    //exit();

    list ($usec, $sec) = explode(" ", microtime());
    $w_microtime = ((float) $usec + (float) $sec);

    // Define o nome do arquivo
    $filename = md5($w_microtime . substr($shtml, 200)) . '.htm';
    // Verifica a necessidade de cria��o do diret�rio de arquivos tempor�rios
    $l_dir = DiretorioCliente($w_cliente) . '/' . tmp;
    if (!(file_exists($l_dir)))
      mkdir($l_dir);

    $handle = fopen($l_dir . '/' . $filename, 'a+');
    if (is_writable($l_dir . '/' . $filename)) {
      fwrite($handle, $shtml);
    }
    $w_protocolo = explode('/', $_SERVER["SERVER_PROTOCOL"]);
    $w_prot = $w_protocolo[0];
  ?>
    <html>
    <body>
    <form name="formpdf" action="<?php echo $w_dir_volta . 'classes/pd4ml/pd4ml.php'; ?>" method="post">
    <input type="hidden" value="<?php echo $w_prot.'://'.$_SERVER['SERVER_NAME'].':'.$_SERVER['SERVER_PORT'].$conFileVirtual . $w_cliente . '/tmp/' . $filename; ?>" name="url">
    <input type="hidden" value="<?php echo $conFilePhysical . $w_cliente . '/tmp/' . $filename; ?>" name="filename">
    <input type="hidden" value="<?php echo $p_orientation; ?>" name="orientation">
    </form>
    <?php echo('<script>document.formpdf.submit();</script>'); ?>
    </body>
    </html>
    <?php

  }

  // =========================================================================
  // Montagem da estrutura do documento
  // -------------------------------------------------------------------------
  function Estrutura_Topo() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('<DIV id=container>');
      ShowHTML('  <DIV id=cab>');
    }
  }

  // =========================================================================
  // Defini��o dos arquivos de CSS
  // -------------------------------------------------------------------------
  function Estrutura_CSS($l_cliente) {
    extract($GLOBALS);
    if ($l_cliente == 6761) {
      ShowHTML('<LINK  media=screen href="' . $conFileVirtual . $l_cliente . '/css/estilo.css" type=text/css rel=stylesheet>');
      ShowHTML('<LINK media=print href="' . $conFileVirtual . $l_cliente . '/css/print.css" type=text/css rel=stylesheet>');
      ShowHTML('<SCRIPT language=javascript src="' . $conFileVirtual . $l_cliente . '/js/scripts.js" type=text/javascript> ');
      ShowHTML('</SCRIPT>');
    }
  }

  // =========================================================================
  // Montagem da estrutura do documento
  // -------------------------------------------------------------------------

  function Estrutura_Topo_Limpo() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('<center>');
      ShowHTML('<DIV id=container_limpo>');
      ShowHTML('  <DIV id=cab>');
    }
  }

  // =========================================================================
  // Montagem do corpo do documento
  // -------------------------------------------------------------------------
  function Estrutura_Fecha() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('  </DIV>');
    }
  }

  // =========================================================================
  // Montagem do corpo do documento
  // -------------------------------------------------------------------------
  function Estrutura_Corpo_Abre() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('  <DIV id=corpo>');
    }
  }

  // =========================================================================
  // Montagem do texto do corpo
  // -------------------------------------------------------------------------
  function Estrutura_Texto_Abre() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('    <DIV id=texto>');
      ShowHTML('        <DIV class=retranca>' . $TP . '</DIV>');
    } else {
      ShowHTML('<B><FONT COLOR="#000000">' . $w_TP . '</FONT></B>');
      ShowHTML('<HR>');
      ShowHTML('<div align=center><center>');
    }
  }

  // =========================================================================
  // Encerramento do texto do corpo
  // -------------------------------------------------------------------------
  function Estrutura_Texto_Fecha() {
    ShowHTML('    </center>');
    ShowHTML('    </DIV>');
  }

  // =========================================================================
  // Montagem da estrutura do menu � esquerda
  // -------------------------------------------------------------------------
  function Estrutura_Menu_Esquerda() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('    <DIV id=menuesq>');
    }
  }

  // =========================================================================
  // Montagem da estrutura do menu � esquerda
  // -------------------------------------------------------------------------
  function Estrutura_Menu_Direita() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('    <DIV id=menudir>');
    }
  }

  // =========================================================================
  // Montagem do menu � esquerda
  // -------------------------------------------------------------------------
  function Estrutura_Menu_Separador() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('      <DIV id=menusep><HR></DIV>');
    }
  }

  // =========================================================================
  // Montagem do menu � esquerda
  // -------------------------------------------------------------------------
  function Estrutura_Menu_Gov_Abre() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      Estrutura_Menu_Separador();
      ShowHTML('      <UL id=menugov>');
    }
  }

  // =========================================================================
  // Montagem do menu � esquerda
  // -------------------------------------------------------------------------
  function Estrutura_Menu_Nav_Abre() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      Estrutura_Menu_Separador();
      ShowHTML('      <DIV id=menunav>');
      ShowHTML('        <UL>');
    }
  }

  // =========================================================================
  // Montagem do menu � esquerda
  // -------------------------------------------------------------------------
  function Estrutura_Menu_Fecha() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('      </UL>');
    }
  }

  // =========================================================================
  // Montagem do sub-menu � esquerda alternativo
  // -------------------------------------------------------------------------
  function Estrutura_Corpo_Menu_Esquerda() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('    <DIV id=menuesq>');
      ShowHTML('      <DIV id=logomenuesq><H3>BresilBresils</H3></DIV>');
    }
  }

  // =========================================================================
  // Montagem da estrutura do documento
  // -------------------------------------------------------------------------
  function Estrutura_Menu() {
    extract($GLOBALS);
    if ($_SESSION['P_CLIENTE'] == 6761) {
      ShowHTML('    <DIV id=cabtopo>');
      ShowHTML('      <DIV id=logoesq>');
      ShowHTML('        <H1>Minist�rio da Cultura</H1>');
      ShowHTML('        <br>');
      ShowHTML('        <select name="opcoes" onChange="if(options[selectedIndex].value) window.location.href= (options[selectedIndex].value)" class="pr">');
      ShowHTML('          <option>Destaques do governo</option>');
      ShowHTML('          <option value="javascript:nova_jan("http://www.brasil.gov.br")">Portal do Governo Federal</option>');
      ShowHTML('          <option value="javascript:nova_jan("http://www.e.gov.br")">Portal de Servi&ccedil;os do Governo</option>');
      ShowHTML('          <option value="javascript:nova_jan("http://www.radiobras.gov.br")">Portal da Ag&ecirc;ncia de Not&iacute;cias</option>');
      ShowHTML('          <option value="javascript:nova_jan("http://www.brasil.gov.br/emquestao")">Em Quest�o</option>');
      ShowHTML('          <option value="javascript:nova_jan("http://www.fomezero.gov.br")">Programa Fome Zero</option>');
      ShowHTML('        </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;');
      ShowHTML('      </DIV>');
      ShowHTML('      <DIV id=logodir><H2>Projeto Ano do Brasil na Fran�a</H2></DIV>');
      ShowHTML('    </DIV>');
      ShowHTML('');
      ShowHTML('    <DIV id=menutxt>');
      ShowHTML('      <SCRIPT src="' . $conFileVirtual . $w_cliente . '/js/newcssmenu.js" type=text/javascript></SCRIPT>');
      ShowHTML('      ');
      ShowHTML('      <DIV id=menutexto>');
      ShowHTML('        <DIV id=mainMenu>');
      ShowHTML('          <UL id=menuList>');
      $l_cont = 0;
      $l_RS = db_getLinkDataUser :: getInstanceOf($dbms, $_SESSION['P_CLIENTE'], $_SESSION['SQ_PESSOA'], null);
      $l_cont = 0;
      foreach ($l_RS as $row) {
        $l_titulo = f(row, 'nome');
        if (f(row, 'filho') > 0) {
          $l_cont = $l_cont +1;
          ShowHTML('            <LI class=menubar>::<A class=starter HREF="javascript:this.status.value;"> ' . f(row, 'nome') . '</A>');
          ShowHTML('            <UL class=menu id=menu' . $l_cont . '>');
          $l_cont1 = 0;
          $l_RS1 = db_getLinkDataUser :: getInstanceOf($dbms, $_SESSION['P_CLIENTE'], $_SESSION['SQ_PESSOA'], f(row, 'sq_menu'));
          foreach ($l_RS1 as $row1) {
            $l_titulo = $l_titulo . ' - ' . f(row1, 'nome');
            if (f(row1, 'filho') > 0) {
              $l_cont1 = $l_cont1 +1;
              ShowHTML('              <LI><A HREF="javascript:this.status.value;"><IMG height=12 alt=">" src="' . $conFileVirtual . $w_cliente . '/img/arrows.gif" width=8> ' . f(row1, 'nome') . '</A> ');
              ShowHTML('              <UL class=menu id=menu' . $l_cont . '_' . $l_cont1 . '>');
              $l_cont2 = 0;
              $l_RS2 = db_getLinkDataUser :: getInstanceOf($dbms, $_SESSION['P_CLIENTE'], $_SESSION['SQ_PESSOA'], f(row1, 'sq_menu'));
              foreach ($l_RS2 as $row2) {
                $l_titulo = $l_titulo . ' - ' . f(row2, 'nome');
                if (f(row2, 'filho') > 0) {
                  $l_cont2 = $l_cont2 +1;
                  ShowHTML('                <LI><A HREF="javascript:this.status.value;"><IMG height=12 alt=">" src="' . $conFileVirtual . $w_cliente . '/img/arrows.gif" width=8> ' . f(row2, 'nome') . '</A> ');
                  ShowHTML('                <UL class=menu id=menu' . $l_cont . '_' . $l_cont1 . '_' . $l_cont2 . '>');
                  $l_RS3 = db_getLinkDataUser :: getInstanceOf($dbms, $_SESSION['P_CLIENTE'], $_SESSION['SQ_PESSOA'], f(row2, 'sq_menu'));
                  foreach ($l_RS3 as $row3) {
                    $l_titulo = $l_titulo . ' - ' . f(row3, 'nome');
                    if (f(row3, 'externo') == 'S') {
                      ShowHTML('                  <LI><A href="' . str_replace('@files', $conFileVirtual . $_SESSION['P_CLIENTE'], f(row3, 'link')) . '" TARGET="' . f(row3, 'target') . '">' . f(row3, 'nome') . '</A> ');
                    } else {
                      ShowHTML('                  <LI><A href="' . f(row3, 'link') . '&P1=' . f(row3, 'p1') . '&P2=' . f(row3, 'p2') . '&P3=' . f(row3, 'p3') . '&P4=' . f(row3, 'p4') . '&TP=' . $l_titulo . '&SG=' . f(row3, 'sigla') . '">' . f(row3, 'nome') . '</A> ');
                    }
                    $l_titulo = str_replace(' - ' . f(row3, 'nome'), '', $l_titulo);
                  }
                  ShowHTML('            </UL>');
                } else {
                  if (f(row2, 'externo') == 'S') {
                    ShowHTML('                <LI><A href="' . str_replace('@files', $conFileVirtual . $_SESSION['P_CLIENTE'], f(row2, 'link')) . '" TARGET="' . f(row2, 'target') . '">' . f(row2, 'nome') . '</A> ');
                  } else {
                    ShowHTML('                <LI><A href="' . f(row2, 'link') . '&P1=' . f(row2, 'p1') . '&P2=' . f(row2, 'p2') . '&P3=' . f(row2, 'p3') . '&P4=' . f(row2, 'p4') . '&TP=' . $l_titulo . '&SG=' . f(row2, 'sigla') . '">' . f(row2, 'nome') . '</A> ');
                  }
                }
                $l_titulo = str_replace(' - ' . f(row2, 'nome'), '', $l_titulo);
              }
              ShowHTML('            </UL>');
            } else {
              if (f(row1, 'externo') == 'S') {
                if (f(row1, 'link') > '') {
                  ShowHTML('              <LI><A href="' . str_replace('@files', $conFileVirtual . $_SESSION['P_CLIENTE'], f(row1, 'link')) . '" TARGET="' . f(row1, 'target') . '">' . f(row1, 'nome') . '</A> ');
                } else {
                  ShowHTML('              <LI>' . f(row1, 'nome') . ' ');
                }
              } else {
                ShowHTML('              <LI><A href="' . f(row1, 'link') . '&P1=' . f(row1, 'p1') . '&P2=' . f(row1, 'p2') . '&P3=' . f(row1, 'p3') . '&P4=' . f(row1, 'p4') . '&TP=' . $l_titulo . '&SG=' . f(row1, 'sigla') . '">' . f(row1, 'nome') . '</A> ');
              }
            }
            $l_titulo = str_replace(' - ' . f(row1, 'nome'), '', $l_titulo);
          }
          ShowHTML('            </UL>');
        } else {
          if (f(row, 'externo') == 'S') {
            ShowHTML('            <LI class=menubar>::<A class=starter href="' . str_replace('@files', $conFileVirtual . $_SESSION['P_CLIENTE'], f(row, 'link')) . '" TARGET="' . f(row, 'target') . '"> ' . f(row, 'nome') . '</A>');
          } else {
            ShowHTML('            <LI class=menubar>::<A class=starter href="' . f(row, 'link') . '&P1=' . f(row, 'p1') . '&P2=' . f(row, 'p2') . '&P3=' . f(row, 'p3') . '&P4=' . f(row, 'p4') . '&TP=' . $l_titulo . '&SG=' . f(row, 'sigla') . '"> ' . f(row, 'nome') . '</A>');
          }
        }
      }
      ShowHTML('            <LI class=menubar>::<A class=starter href="' . $w_dir . 'menu.php?par=Sair" & " onClick="return(confirm("Confirma sa�da do sistema?"));"> Sair</A>');
      ShowHTML('          </UL>');
      ShowHTML('        </DIV>');
      ShowHTML('      </DIV>');
      ShowHTML('    </DIV>');
    }
  }

  // =========================================================================
  // Abre conex�o com o banco de dados
  // -------------------------------------------------------------------------
  function abreSessao() {
    return abreSessao :: getInstanceOf($_SESSION["DBMS"]);
  }

  // =========================================================================
  // Fecha conex�o com o banco de dados
  // -------------------------------------------------------------------------
  function FechaSessao($dbms) {
    extract($GLOBALS);
    unset ($objConnectionClass);
  }

  // =========================================================================
  // Rotina de Fechamento do BD
  // -------------------------------------------------------------------------
  function DesConectaBD() {
    return true;
  }

  // -------------------------------------------------------------------------
  // Torna a chamada a um campo de recordset case insensitive
  // =========================================================================
  function f($rs, $fld) {
    if (is_array($rs)) {
      $rs = array_change_key_case($rs, CASE_LOWER);
    }
    if (isset ($rs[Strtolower($fld)])) {
      if (strpos(strtoupper($rs[Strtoupper($fld)]), '/SCRIPT') !== false) {
        $val = str_replace('/SCRIPT', '/ SCRIPT', strtoupper($rs[Strtolower($fld)]));
      } else
        $val = str_replace('.asp', '.php', $rs[Strtolower($fld)]);
    }
    elseif (isset ($rs[Strtoupper($fld)])) {
      if (strpos(strtoupper($rs[Strtoupper($fld)]), '/SCRIPT') !== false) {
        $val = str_replace('/SCRIPT', '/ SCRIPT', strtoupper($rs[Strtoupper($fld)]));
      } else
        $val = str_replace('.asp', '.php', $rs[Strtoupper($fld)]);
    }
    elseif (isset ($rs[$fld])) {
      if (strpos(strtoupper($rs[Strtoupper($fld)]), '/SCRIPT') !== false) {
        $val = str_replace('/SCRIPT', '/ SCRIPT', strtoupper($rs[$fld]));
      } else
        $val = str_replace('.asp', '.php', $rs[$fld]);
    } else
      return null;
    return str_replace('"','&quot;',$val); //str_replace(chr(166),'�',str_replace(chr(161),'�',str_replace(chr(144),'�',str_replace('�','�',str_replace('�','�',str_replace('�','�',str_replace('�','�',str_replace('�','�',str_replace('�','�',str_replace('�','�',str_replace('��','��',str_replace('�','�',str_replace('�','�',str_replace('�','�',str_replace('�','�',$val)))))))))))))));
  }

  // -------------------------------------------------------------------------
  // Verifica se a senha de acesso do usu�rio est� correta
  // =========================================================================
  function VerificaSenhaAcesso($Usuario, $Senha) {
    extract($GLOBALS);
    if (db_verificaSenha :: getInstanceOf($dbms, $_SESSION["P_CLIENTE"], $Usuario, $Senha) == 0)
      return true;
    else
      return false;
  }

  // -------------------------------------------------------------------------
  // Verifica se a Assinatura Eletronica do usu�rio est� correta
  // =========================================================================
  function VerificaAssinaturaEletronica($Usuario, $Senha) {
    extract($GLOBALS);
    if ($Senha > '') {
      if (db_verificaAssinatura :: getInstanceOf($dbms, $_SESSION["P_CLIENTE"], $Usuario, $Senha) == 0)
        return true;
      else
        return false;
    } else {
      return true;
    }
  }

  // =========================================================================
  // Fun��o que formata dias, horas, minutos e segundos a partir dos segundos
  // -------------------------------------------------------------------------
  function FormataDataEdicao($w_dt_grade, $w_formato = 1) {
    if (nvl($w_dt_grade, '') > '') {
      if (is_numeric($w_dt_grade)) {
        switch ($w_formato) {
          case 1 :
            return date('d/m/Y', $w_dt_grade);
            break;
          case 2 :
            return date('H:i:s', $w_dt_grade);
            break;
          case 3 :
            return date('d/m/Y, H:i:s', $w_dt_grade);
            break;
          case 4 :
            return diaSemana(date('l, d/m/y, H:i:s', $w_dt_grade));
            break;
          case 5 :
            return date('d/m/y', $w_dt_grade);
            break;
          case 6 :
            return date('d/m/y, H:i:s', $w_dt_grade);
            break;
          case 7 :
            return date('Y-m-d', $w_dt_grade);
            break;
          case 8 :
            return date('Y-m-d H:i:s', $w_dt_grade);
            break;
          case 9 :
            return diaSemana(date('l, ')) . date('d \d\e ') . mesAno(date('F')) . date(' \d\e Y', $w_dt_grade);
            break;
        }
      } else {
        return $w_dt_grade;
      }
    } else {
      return null;
    }
  }

  // =========================================================================
  // Fun��o que retorna datetime com o primeiro dia da data informada no formato datetime
  // -------------------------------------------------------------------------
  function first_day($w_valor) {
    extract($GLOBALS);

    $l_valor = FormataDataEdicao($w_valor);
    $l_mes = substr($l_valor, 3, 2);
    $l_ano = substr($l_valor, 6, 4);
    return mktime(0, 0, 0, $l_mes, 1, $l_ano);
  }

  // =========================================================================
  // Fun��o que retorna datetime com o �ltimo dia da data informada no formato datetime
  // -------------------------------------------------------------------------
  function last_day($w_valor) {
    extract($GLOBALS);

    $l_valor = FormataDataEdicao($w_valor);
    $l_dia = substr($l_valor, 0, 2);
    $l_mes = substr($l_valor, 3, 2);
    $l_ano = substr($l_valor, 6, 4);

    $l_result = mktime(0, 0, 0, ($l_mes +1), 0, $l_ano);

    return $l_result;
  }

  // =========================================================================
  // Fun��o que retorna data indicando o domingo de p�scoa de um ano
  // -------------------------------------------------------------------------
  function DomingoPascoa($p_ano) {
    extract($GLOBALS);

    $a = intval($p_ano % 19);
    $b = intval($p_ano / 100);
    $c = intval($p_ano % 100);
    $d = intval($b / 4);
    $e = intval($b % 4);
    $f = intval(($b +8) / 25);
    $g = intval(($b - $f +1) / 3);
    $h = intval(((19 * $a) + $b - $d - $g +15) % 30);
    $i = intval($c / 4);
    $k = intval($c % 4);
    $l = intval((32 + (2 * $e) + (2 * $i) - $h - $k) % 7);
    $m = intval(($a + (11 * $h) + (22 * $l)) / 451);
    $p = intval(($h + $l - (7 * $m) + 114) / 31);
    $q = intval(($h + $l - (7 * $m) + 114) % 31);
    return mktime(0, 0, 0, $p, $q +1, $p_ano);
  }

  // =========================================================================
  // Fun��o que retorna data indicando a sexta-feira santa de um ano
  // Sexta-feira Santa � 2 dias antes do Domingo de P�scoa
  // -------------------------------------------------------------------------
  function SextaSanta($p_ano) {
    extract($GLOBALS);

    return addDays(DomingoPascoa($p_ano), -2);
  }

  // =========================================================================
  // Fun��o que retorna data de Corpus Christi de um ano
  // Corpus Chirsti � 60 dias depois do Domingo de P�scoa
  // -------------------------------------------------------------------------
  function CorpusChristi($p_ano) {
    extract($GLOBALS);

    return addDays(DomingoPascoa($p_ano), 60);
  }

  // =========================================================================
  // Fun��o que retorna data indicando a ter�a-feira de carnaval de um ano
  // Ter�a-feira de carnaval � a primeira ter�a-feira 42 dias antes do domingo
  // de p�scoa
  // -------------------------------------------------------------------------
  function TercaCarnaval($p_ano) {
    extract($GLOBALS);

    $l_dia = addDays(DomingoPascoa($p_ano), -42);
    if (date('w', $l_dia) > 2) {
      return addDays($l_dia, (-1 * date('w', addDays($l_dia, -2))));
    } else {
      return addDays($l_dia, (-1 * date('w', addDays($l_dia, -4))));
    }
  }

  // =========================================================================
  // Fun��o que traduz os dias da semana de ingl�s para portugu�s
  // -------------------------------------------------------------------------
  function diaSemana($l_data) {
    if (nvl($l_data, '') > '') {
      $l_texto = substr($l_data, strpos($l_data, ','));
      switch (strtoupper(substr($l_data, 0, strpos($l_data, ',')))) {
        case 'SUNDAY' :
          return 'Domingo' . $l_texto;
          break;
        case 'MONDAY' :
          return 'Segunda-feira' . $l_texto;
          break;
        case 'TUESDAY' :
          return 'Ter�a-feira' . $l_texto;
          break;
        case 'WEDNESDAY' :
          return 'Quarta-feira' . $l_texto;
          break;
        case 'THURSDAY' :
          return 'Quinta-feira' . $l_texto;
          break;
        case 'FRIDAY' :
          return 'Sexta-feira' . $l_texto;
          break;
        case 'SATURDAY' :
          return 'S�bado' . $l_texto;
          break;
      }
    } else {
      return null;
    }
  }
  // =========================================================================
  // Fun��o que traduz os meses do ano de ingl�s para portugu�s
  // -------------------------------------------------------------------------
  function mesAno($l_data, $l_formato = null) {
    if (nvl($l_data, '') > '') {
      if (nvl($l_formato, 'nulo') == 'nulo') {
        switch (strtoupper($l_data)) {
          case 'JANUARY' :
            return 'Janeiro';
            break;
          case 'FEBRUARY' :
            return 'Fevereiro';
            break;
          case 'MARCH' :
            return 'Mar�o';
            break;
          case 'APRIL' :
            return 'Abril';
            break;
          case 'MAY' :
            return 'Maio';
            break;
          case 'JUNE' :
            return 'Junho';
            break;
          case 'JULY' :
            return 'Julho';
            break;
          case 'AUGUST' :
            return 'Agosto';
            break;
          case 'SEPTEMBER' :
            return 'Setembro';
            break;
          case 'OCTOBER' :
            return 'Outubro';
            break;
          case 'NOVEMBER' :
            return 'Novembro';
            break;
          case 'DECEMBER' :
            return 'Dezembro';
            break;
          default :
            return $l_data;
        }
      } else {
        switch (strtoupper($l_data)) {
          case 'JANUARY' :
            return 'Jan';
            break;
          case 'FEBRUARY' :
            return 'Fev';
            break;
          case 'MARCH' :
            return 'Mar';
            break;
          case 'APRIL' :
            return 'Abr';
            break;
          case 'MAY' :
            return 'Mai';
            break;
          case 'JUNE' :
            return 'Jun';
            break;
          case 'JULY' :
            return 'Jul';
            break;
          case 'AUGUST' :
            return 'Ago';
            break;
          case 'SEPTEMBER' :
            return 'Set';
            break;
          case 'OCTOBER' :
            return 'Out';
            break;
          case 'NOVEMBER' :
            return 'Nov';
            break;
          case 'DECEMBER' :
            return 'Dez';
            break;
        }
      }
    } else {
      return null;
    }
  }

  function montaCalendario($p_mes, $p_datas, $p_cores, $p_imagem) {
    extract($GLOBALS, EXTR_PREFIX_SAME, 'ex_');
    $p_detalhe = false;

    // Atribui nomes dos meses
    $l_meses[1] = 'Janeiro';
    $l_meses[2] = 'Fevereiro';
    $l_meses[3] = 'Mar�o';
    $l_meses[4] = 'Abril';
    $l_meses[5] = 'Maio';
    $l_meses[6] = 'Junho';
    $l_meses[7] = 'Julho';
    $l_meses[8] = 'Agosto';
    $l_meses[9] = 'Setembro';
    $l_meses[10] = 'Outubro';
    $l_meses[11] = 'Novembro';
    $l_meses[12] = 'Dezembro';

    // Atribui quantidade de dias em cada m�s
    $l_qtd[1] = 31;
    $l_qtd[2] = 28;
    $l_qtd[3] = 31;
    $l_qtd[4] = 30;
    $l_qtd[5] = 31;
    $l_qtd[6] = 30;
    $l_qtd[7] = 31;
    $l_qtd[8] = 31;
    $l_qtd[9] = 30;
    $l_qtd[10] = 31;
    $l_qtd[11] = 30;
    $l_qtd[12] = 31;

    // Atribui sigla para cada dia da semana
    $l_dias[1] = 'D';
    $l_dias[2] = 'S';
    $l_dias[3] = 'T';
    $l_dias[4] = 'Q';
    $l_dias[5] = 'Q';
    $l_dias[6] = 'S';
    $l_dias[7] = 'S';

    // Recupera o m�s e o ano desejado para montagem do calend�rio
    $l_mes = substr($p_mes, 0, 2);
    $l_ano = substr($p_mes, 2, 4);

    // Define cor de fundo padr�o para as c�lulas de s�bado e domingo
    $l_cor_padrao = ' bgcolor=#DAEABD';

    // Trata o m�s de fevereiro anos bissextos
    if (fMod($l_ano, 4) == 0)
      $l_qtd[2] = 29;

    $l_html = '<table border=0 cellspacing=1 cellpadding=1>' . $crlf;
    $l_html .= '  <tr><td colspan=7 align="center" ' . $l_cor_padrao . '><b>' . $l_meses[intVal($l_mes)] . '/' . $l_ano . '</td></tr>' . $crlf;
    $l_html .= '  <tr align="center">' . $crlf;

    // Monta a linha com a sigla para os dias das semanas
    for ($i = 1; $i <= 7; $i++)
      $l_html .= '    <td ' . $l_cor_padrao . '><b>' . $l_dias[$i] . '</td>' . $crlf;
    $l_html .= '  </tr>' . $crlf;

    // Define em que dia da semana o m�s inicia
    $l_inicio = date('w', toDate('01/' . $l_mes . '/' . $l_ano));

    // Carrega os dias do m�s num array que ser� usado para montagem do calend�rio, colocando
    // o dia ou um espa�o em branco, dependendo do in�cio e do fim do m�s
    for ($i = 1; $i <= ($l_inicio); $i++)
      $l_celulas[$i] = '&nbsp;';
    for ($i = ($l_qtd[intVal($l_mes)] + 1); $i <= 42; $i++)
      $l_celulas[$i] = '&nbsp;';
    for ($i = 1; $i <= ($l_qtd[intVal($l_mes)]); $i++)
      $l_celulas[($i + $l_inicio)] = $i;

    // Monta o calend�rio, usando o array $l_celulas
    $l_html .= '  <tr align="center">' . $crlf;
    for ($i = 1; $i <= 42; $i++) {
      // Trata a borda da c�lula para datas especiais
      $l_borda = '';
      $l_ocorrencia = '';
      if ($l_celulas[$i] != "&nbsp;") {
        If (isset ($p_datas[$l_celulas[$i]][$l_mes][SubStr($l_ano, 2, 2)])) {
          $l_borda = ' style="border: 1px solid rgb(0,0,0);"';
          $l_ocorrencia = 'title="' . $p_datas[$l_celulas[$i]][$l_mes][SubStr($l_ano, 2, 2)] . '"';
        }
      }

      // Trata a cor de fundo da c�lula
      $l_cor = '';
      $l_img = '';
      if ($i == 1 or ($l_celulas[$i] != '&nbsp;' and ((fMod($i, 7) == 0) or (fMod($i -1, 7) == 0)))) {
        $l_img = ' style="background-color:#DAEABD;"';
        if (isset ($p_imagem[$l_celulas[$i]][$l_mes][Substr($l_ano, 2, 2)])) {
          $l_img = ' style="background: url(img/' . $p_imagem[$l_celulas[$i]][$l_mes][Substr($l_ano, 2, 2)] . ') no-repeat center; background-color: #DAEABD;"';
        }
      } else
        if ($l_celulas[$i] != '&nbsp;') {
          if (isset ($p_cores[$l_celulas[$i]][$l_mes][SubStr($l_ano, 2, 2)])) {
            $l_cor = ' background-color: ' . $p_cores[$l_celulas[$i]][$l_mes][SubStr($l_ano, 2, 2)];
          }
          if (isset ($p_imagem[$l_celulas[$i]][$l_mes][SubStr($l_ano, 2, 2)])) {
            //                $l_img = ' style="background: url(/img/' . $p_imagem[$l_celulas[$i]] [$l_mes][SubStr($l_ano,2,2)] . ') no-repeat center;' . $l_cor. ' "';
            $l_img = ' style="background: url(img/' . $p_imagem[$l_celulas[$i]][$l_mes][Substr($l_ano, 2, 2)] . ') no-repeat center; ' . $l_cor . '"';

          }
        }

      //Coloca uma c�lula do calend�rio
      $l_html .= '    <td ' . $l_img . $l_ocorrencia . '>' . $l_celulas[$i] . '</td>' . $crlf;

      // Trata a data de hoje
      if ($l_data == formataDataEdicao(time())) {
        if ($l_data == formataDataEdicao(time()))
          $l_ocorrencia = 'HOJE\r\n' . $l_ocorrencia;
        $l_borda = ' style="border: 2px solid rgb(0,0,0);"';
      }

      if ($l_ocorrencia != '')
        $l_ocorrencia = ' onClick="javascript:alert(\'' . $l_ocorrencia . '\')"';

      // Trata a quebra de linha ao final de cada semana
      if (fMod($i, 7) == 0) {
        $l_html .= '  </tr>' . $crlf;
        // Interrompe a montagem do calend�rio na �ltima linha que cont�m datas
        if ($i > $l_qtd[intVal($l_mes)] && $l_celulas[$i +1] == '&nbsp;') {
          break;
        } else {
          $l_html .= '  <tr align="center">' . $crlf;
        }
      }
    }
    $l_html .= '</table>' . $crlf;

    // Devolve o calend�rio montado
    return $l_html;
  }

  // =========================================================================
  // Fun��o para retornar um array com todos os dias de um per�odo
  // Recebe o in�cio e o fim do per�odo no formato data
  // Todos os elementos do array recebem o valor definido em p_valor
  // -------------------------------------------------------------------------
  function retornaArrayDias($p_inicio, $p_fim, $p_array, $p_valor, $p_dia_util = null) {
    $l_inicio = date(Ymd, $p_inicio);
    $l_fim = date(Ymd, $p_fim);
    // Atribui quantidade de dias em cada m�s
    $l_qtd[1] = 31;
    $l_qtd[2] = 28;
    $l_qtd[3] = 31;
    $l_qtd[4] = 30;
    $l_qtd[5] = 31;
    $l_qtd[6] = 30;
    $l_qtd[7] = 31;
    $l_qtd[8] = 31;
    $l_qtd[9] = 30;
    $l_qtd[10] = 31;
    $l_qtd[11] = 30;
    $l_qtd[12] = 31;

    // Trata o m�s de fevereiro anos bissextos
    if (fMod($l_ano, 4) == 0)
      $l_qtd[2] = 29;

    for ($i = $l_inicio; $i <= $l_fim; $i++) {
      $l_ano = substr($i, 0, 4);
      $l_mes = substr($i, 4, 2);
      $l_dia = substr($i, 6, 2);
      if (intVal($l_dia) > $l_qtd[intVal($l_mes)]) {
        if (intVal($l_mes) == 12) {
          $i = ($l_ano +1) . '0101';
        } else {
          $i = $l_ano . substr((100 + intVal($l_mes) + 1), 1, 2) . '01';
        }
      }
      $p_array[substr($i, 6, 2) . '/' . substr($i, 4, 2) . '/' . substr($i, 0, 4)]['valor'] = $p_valor;
      $p_array[substr($i, 6, 2) . '/' . substr($i, 4, 2) . '/' . substr($i, 0, 4)]['dia_util'] = $p_dia_util;
    }

    return true;
  }

  // =========================================================================
  // Fun��o para retornar array com o tipo do nome e o nome mais adequado para um per�odo de datas
  // Recebe o in�cio e o fim do per�odo no formato data
  // Devolve array com dois �ndices: 
  //    [TIPO] pode valer ANO, MES_ANO, DIA, OUT
  //    [VALOR] tipo=ANO retorna o ano do in�cio informado
  //            tipo=MES retorna o nome_mes/ano
  //            tipo=DIA retorna datetime do in�cio informado
  //            tipo=OUT retorna nulo
  // -------------------------------------------------------------------------
  function retornaNomePeriodo($p_inicio, $p_fim) {
    if (date(dm, $p_inicio) == '0101' && date(dm, $p_fim) == '3112' && date(Y, $p_inicio) == date(Y, $p_fim)) {
      // se o per�odo compreende totalmente um �nico ano, devolve o ano
      $p_array['TIPO'] = 'ANO';
      $p_array['VALOR'] = date(Y, $p_inicio);
    }
    elseif (date(d, $p_inicio) == '01' && last_day($p_inicio) == $p_fim) {
      // se o per�odo compreende um �nico dia, devolve o dia
      $p_array['TIPO'] = 'MES';
      $p_array['VALOR'] = mesAno(date(F, $p_inicio), 'resumido') . '/' . date(y, $p_inicio);
    }
    elseif ($p_inicio == $p_fim) {
      // se o per�odo compreende um �nico dia, devolve o dia
      $p_array['TIPO'] = 'DIA';
      $p_array['VALOR'] = $p_inicio;
    } else {
      $p_array['TIPO'] = 'OUT';
      $p_array['VALOR'] = null;
    }
    return $p_array;
  }

  // =========================================================================
  // Fun��o para retornar a strig 'Sim' ou 'N�o'
  // -------------------------------------------------------------------------
  function retornaSimNao($chave, $formato = null) {
    extract($GLOBALS);
    if (strtoupper($formato) == 'IMAGEM') {
      switch ($chave) {
        case 'S' :
          return '<img src="' . $conImgOkNormal . '" border=0 width=15 height=15 align="center">';
          break;
        default :
          return '&nbsp;';
      }
    } else {
      switch ($chave) {
        case 'S' :
          return 'Sim';
          break;
        case 'N' :
          return 'N�o';
          break;
        default :
          return 'N�o';
      }
    }
  }

  //Limpa Mascara para gravar os dados no banco de dados
  function LimpaMascara($Campo) {
    return str_replace(str_replace(str_replace(str_replace(str_replace(str_replace($campo, ',', ''), ';', ''), '.', ''), '-', ''), '/', ''), '"', '');
  }

  // Cria a tag Body
  function BodyOpen($cProperties) {
    extract($GLOBALS);
    Required();
    $wProperties = $cProperties;
    if (nvl($wProperties, '') != '') {
      if (strpos($wProperties, '"init') !== false) {
        $wProperties = str_replace('init', 'required(); init', $wProperties);
      }
      elseif (strpos($wProperties, 'this.') !== false) {
        $wProperties = str_replace('this.', 'required(); this.', $wProperties);
      }
      elseif (strpos($wProperties, '"document.') !== false || strpos($wProperties, '=document.') !== false) {
        $wProperties = str_replace('document.', 'required(); document.', $wProperties);
        if (strpos($wProperties, 'required()') === false)
          $wProperties = str_replace('this.focus', 'required(); this.focus', $wProperties);
      } else {
        $wProperties = str_replace('document.', 'required(); document.', $wProperties);
        //       $wProperties = 'required(); '.$wProperties;
      }
    }

    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<link rel="stylesheet" type="text/css" href="' . $conRootSIW . 'classes/menu/xPandMenu.css">');

    if ($_SESSION['P_CLIENTE'] == '6761') {
      ShowHTML('<body Text="' . $conBodyText . '" ' . $wProperties . '> ');
    } else {
      ShowHTML('<body Text="' . $conBodyText . '" Link="' . $conBodyLink . '" Alink="' . $conBodyALink . '" ' .
      'Vlink="' . $conBodyVLink . '" Bgcolor="' . $conBodyBgColor . '" Background="' . $conBodyBackground . '" ' .
      'Bgproperties="' . $conBodyBgproperties . '" Topmargin="' . $conBodyTopmargin . '" ' .
      'Leftmargin="' . $conBodyLeftmargin . '" ' . $wProperties . '> ');
    }
    flush();
  }

  function BodyOpenImage($cProperties, $cImage, $cFixed) {
    extract($GLOBALS);

    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');

    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<link rel="stylesheet" type="text/css" href="' . $conRootSIW . 'classes/menu/xPandMenu.css">');
    if ($_SESSION['P_CLIENTE'] == '6761') {
      ShowHTML('<body Text="' . $conBodyText . '" ' . $cProperties . '> ');
    } else {
      ShowHTML('<body Text="' . $conBodyText . '" Link="' . $conBodyLink . '" Alink="' . $conBodyALink . '" ' .
      'Vlink="' . $conBodyVLink . '" Bgcolor="' . $conBodyBgcolor . '" Background="' . $cImage . '" ' .
      'Bgproperties="' . $cFixed . '" Topmargin="' . $conBodyTopmargin . '" ' .
      'Leftmargin="' . $conBodyLeftmargin . '" ' . $cProperties . '> ');
    }
  }

  // Imprime uma linha HTML
  function ShowHtml($Line) {
    print $Line . chr(13) . chr(10);
  }

  // Cria a tag Body
  function BodyOpenClean($cProperties) {
    extract($GLOBALS);
    Required();
    $wProperties = $cProperties;
    if (nvl($wProperties, '') != '') {
      $wProperties = str_replace('document.', 'required(); document.', $wProperties);
      if (strpos($wProperties, 'required()') === false)
        $wProperties = str_replace('this.focus', 'required(); this.focus', $wProperties);
    } else {
      $wProperties = ' onLoad=\'required();\' ';
    }

    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax.js"></script>');
    ShowHTML('<script type="text/javascript" src="js/modal/js/ajax-dynamic-content.js"></script> ');
    ShowHTML('<script type="text/javascript" src="js/modal/js/modal-message.js"></script> ');
    ShowHTML('<link rel="stylesheet" href="js/modal/css/modal-message.css" type="text/css" media="screen" />');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/funcoes.js"></script>');
    ShowHTML('<script language="javascript" type="text/javascript" src="js/jquery.js"></script>');
    ShowHTML('<link rel="stylesheet" type="text/css" href="' . $conRootSIW . 'classes/menu/xPandMenu.css">');
    ShowHTML('<body Text="' . $conBodyText . '" Link="' . $conBodyLink . '" Alink="' . $conBodyALink . '" ' .
    'Vlink="' . $conBodyVLink . '" Background="' . $conBodyBackground . '" ' .
    'Bgproperties="' . $conBodyBgproperties . '" Topmargin="' . $conBodyTopmargin . '" ' .
    'Leftmargin="' . $conBodyLeftmargin . '" ' . $wProperties . '> ');
    flush();
  }

  // Cria a tag Body
  function BodyOpenMail($cProperties = null) {
    extract($GLOBALS);
    $l_html = '';
    $l_html = $l_html . '<link rel="stylesheet" type="text/css" href="' . $conRootSIW . 'classes/menu/xPandMenu.css">' . chr(13);
    $l_html = $l_html . '<body Text="' . $conBodyText . '" Link="' . $conBodyLink . '" Alink="' . $conBodyALink . '" ' .
    'Vlink="' . $conBodyVLink . '" Bgcolor="' . $conBodyBgcolor . '" Background="' . $conBodyBackground . '" ' .
    'Bgproperties="' . $conBodyBgproperties . '" Topmargin="' . $conBodyTopmargin . '" ' .
    'Leftmargin="' . $conBodyLeftmargin . '" ' . $cProperties . '> ' . chr(13);
    return $l_html;
  }

  // Cria a tag Body
  function BodyOpenWord($cProperties = null) {
    extract($GLOBALS);
    $l_html = '';
    $l_html = $l_html . '<link rel="stylesheet" type="text/css" href="' . $conRootSIW . 'classes/menu/xPandPrint.css">' . chr(13);
    $l_html = $l_html . '<body Text="' . $conBodyText . '" Link="' . $conBodyLink . '" Alink="' . $conBodyALink . '" ' .
    'Vlink="' . $conBodyVLink . '" Bgcolor="' . $conBodyBgcolor . '" Background="' . $conBodyBackground . '" ' .
    'Bgproperties="' . $conBodyBgproperties . '" Topmargin="' . $conBodyTopmargin . '" ' .
    'Leftmargin="' . $conBodyLeftmargin . '" ' . $cProperties . '> ' . chr(13);
    return $l_html;
  }

  //Fun��o que captura o conte�do HTML, viabilizando a transforma��o HTML-PDF.
  function curPageURL() {
    $pageURL = 'http';
    if ($_SERVER["HTTPS"] == "on") {
      $pageURL .= "s";
    }
    $pageURL .= "://";
    if ($_SERVER["SERVER_PORT"] != "80") {
      $pageURL .= "localhost:" . $_SERVER["SERVER_PORT"] . $_SERVER["REQUEST_URI"];
    } else {
      $pageURL .= "localhost" . $_SERVER["REQUEST_URI"];
    }
    return $pageURL;
  }

  function csv($arq) {
    $var = file($arq);
    $saida = array ();
    $qtd = count($var);
    $table = array ();
    $ind = null;

    for ($i = 0; $i < $qtd; $i++) {

      if (substr(trim($var[$i]), 0, 1) == '[' && substr(trim($var[$i]), -1, 1) == ']') {
        if (!is_null($ind)) {
          $saida[$ind] = $table;
        }

        $table = array ();
        $ind = trim($var[$i]);
        ++ $i;
        continue;
      }

      $linha = "";
      $linha = $var[$i];

      while (isset ($var[$i +1]) && substr(trim($var[$i +1]), 0, 1) != '"' && substr(trim($var[$i +1]), 0, 1) != '[' && substr(trim($var[$i]), -1, 1) != ']') {
        $linha .= $var[++ $i];
      }

      $linha = corrigeCar($linha);
      $linha = explode('","', $linha);
      $linha[0] = substr($linha[0], 1);
      $linha[count($linha) - 1] = substr($linha[count($linha) - 1], 0, -1);
      $linha = array_map('colocaPlique', $linha);
      array_push($table, $linha);
    }
    //add a ultima tabela
    $saida[$ind] = $table;

    return $saida;
  }

  function colocaPlique($str) {
    if (trim(strtoupper($str)) != 'NULL') {
      $str = "'" . $str . "'";
    }
    return $str;
  }

  function corrigeCar($str) {
    $str = trim($str);
    $str = str_replace('NULL', '"NULL"', $str);
    $str = str_replace(',,', ',"",', $str);
    $str = str_replace(',,', ',"",', $str);
    $str = str_replace("'", "�", $str);
    return $str;
  }

  function exibeTurno($l_chave) {
    extract($GLOBALS);
    switch ($l_chave) {
      case 1 :
        return 'Matutino';
        break;
      case 2 :
        return 'Vespertino';
        break;
      case 3 :
        return 'Noturno';
        break;

    }
  }

  function day($data) {
    extract($GLOBALS);
    if (nvl($data, '') != '') {
      $day = Date('d', $data);
    }
    return $day;

  }
  function month($data) {
    extract($GLOBALS);
    if (nvl($data, '') != '') {
      $month = Date('m', $data);
    }
    return $month;
  }
  function year($data) {
    extract($GLOBALS);
    if (nvl($data, '') != '') {
      $year = Date('Y', $data);
    }
    return $year;
  }

  Function nomeMes($p_mes,$p_retorno=null) {
  if($p_retorno != ''){
    switch ($p_mes) {
      Case 1 : return 'JANEIRO'; break;
      Case 2 : return 'FEVEREIRO'; break;
      Case 3 : return 'MAR�O'; break;
      Case 4 : return 'ABRIL'; break;
      Case 5 : return 'MAIO'; break;
      Case 6 : return 'JUNHO'; break;
      Case 7 : return 'JULHO'; break;
      Case 8 : return 'AGOSTO'; break;
      Case 9 : return 'SETEMBRO'; break;
      Case 10 : return 'OUTUBRO'; break;
      Case 11 : return 'NOVEMBRO'; break;
      Case 12 : return 'DEZEMBRO'; break;
    }
  }
    switch ($p_mes) {
      Case 1 : return 'JAN'; break;
      Case 2 : return 'FEV'; break;
      Case 3 : return 'MAR'; break;
      Case 4 : return 'ABR'; break;
      Case 5 : return 'MAI'; break;
      Case 6 : return 'JUN'; break;
      Case 7 : return 'JUL'; break;
      Case 8 : return 'AGO'; break;
      Case 9 : return 'SET'; break;
      Case 10 : return 'OUT'; break;
      Case 11 : return 'NOV'; break;
      Case 12 : return 'DEZ'; break;
    }
  }

  function SelecaoTipoEscola($label, $accesskey, $hint = null, $chave, $chaveaux, $campo, $restricao = null, $atributo = null) {
    extract($GLOBALS);

    If (Nvl($chaveAux, "nulo") <> "nulo") {

      $SQL = "select sq_cliente, sq_tipo_cliente cliente " .
      "  from sbpi.Cliente " .
      "  where sq_cliente = " . $chaveAux;

      $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
      $w_sq_tipo_cliente = $RS[0]["cliente"];
    }

    $SQL = "select sq_tipo_cliente , ds_tipo_cliente, ds_registro, tipo, " .
    "       case tipo when '1' then 'Secretaria' " .
    "                 when '2' then 'Regional'   " .
    "                 when '3' then 'Escola'     " .
    "       end nm_tipo                        " .
    "  from sbpi.Tipo_Cliente                  " .
    "  where tipo = 3                          ";

    // if( $chaveaux > "" ){
    //  $SQL .= '   and IsNull(sq_cliente_pai,0) = ' . $chaveaux . $crlf;
    // }
    //  $SQL .= 'ORDER BY b.tipo, ds_cliente' . $crlf;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    if (is_null($hint)) {
      ShowHTML('         <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_disabled . " " . $atributo . '>');
    } else {
      ShowHTML('          <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_disabled . " " . $atributo . '>');
    }
    ShowHTML('          <option value="">---');
    foreach ($RS as $row) {
      if ((nvl(f($row, "sq_tipo_cliente"), -1)) == (nvl($chave, -1))) {
        ShowHTML('          <option value="' . f($row, "sq_tipo_cliente") . '" SELECTED>' . f($row, "ds_tipo_cliente"));
      } else {
        ShowHTML('          <option value="' . f($row, "sq_tipo_cliente") . '">' . f($row, "ds_tipo_cliente"));
      }
    }
    ShowHTML('          </select>');
  }

  function SelecaoRegionalEscola($label, $accesskey, $hint = null, $chave, $chaveaux, $campo, $restricao = null, $atributo = null) {
    extract($GLOBALS);

    $SQL = "SELECT b.tipo, a.sq_cliente cliente, a.sq_tipo_cliente, a.ds_cliente " .
    "  from sbpi.CLIENTE                      a " .
    "       inner      join sbpi.Tipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente and b.tipo = 2) " .
    "       left outer join sbpi.Cliente      c on (a.sq_cliente = c.sq_cliente_pai and c.sq_cliente = " . nvl($chaveAux, -1) . ") ";

    if ($chaveaux > "") {
      $SQL .= '   and nvl(a.sq_cliente_pai,0) = ' . $chaveaux . $crlf;
    }
    $SQL .= 'ORDER BY b.tipo, a.ds_cliente' . $crlf;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    if (is_null($hint)) {
      ShowHTML('         <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS"  NAME="' . $campo . '" ' . $w_disabled . '  ' . $atributo . '>');
    } else {
      ShowHTML('          <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_disabled . '  ' . $atributo . '>');
    }
    ShowHTML('          <option value="">---');

    foreach ($RS as $row) {
      if ((nvl(f($row, "cliente"), -1)) == (nvl($chave, -1))) {
        ShowHTML('          <option value="' . f($row, "cliente") . '" SELECTED>' . f($row, "ds_cliente"));
      } else {
        ShowHTML('          <option value="' . f($row, "cliente") . '">' . f($row, "ds_cliente"));
      }
    }
    ShowHTML('          </select>');
  }

  function SelecaoEscola($label, $accesskey, $hint = null, $chave, $chaveaux, $campo, $restricao = null, $atributo = null) {
    extract($GLOBALS);
    $SQL = "SELECT a.sq_cliente cliente, a.sq_tipo_cliente, case b.tipo when '1' then upper(a.ds_username) else a.ds_cliente end ds_cliente, b.tipo " .
    "  from sbpi.CLIENTE            a" .
    "      inner join sbpi.Tipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente) " .
    " WHERE PUBLICA = 'S' and upper(a.ds_username) <> 'SBPI' ";

    if ($chaveaux > "") {
      $SQL .= '   and nvl(sq_cliente_pai,0) = ' . $chaveaux . $crlf;
    }
    if ($restricao=='ESCOLA') {
      $SQL .= '   and b.tipo = 3' . $crlf;
    }
    $SQL .= 'ORDER BY b.tipo, ds_cliente' . $crlf;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    if (is_null($hint)) {
      ShowHTML('         <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_disabled . " " . $atributo . '>');
    } else {
      ShowHTML('          <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_disabled . " " . $atributo . '>');
    }
    ShowHTML('          <option value="">---');
    foreach ($RS as $row) {
      if ((nvl(f($row, "cliente"), -1)) == (nvl($chave, -1))) {
        ShowHTML('          <option value="' . f($row, "cliente") . '" SELECTED>' . f($row, "ds_cliente"));
      } else {
        ShowHTML('          <option value="' . f($row, "cliente") . '">' . f($row, "ds_cliente"));
      }
    }
    ShowHTML('          </select>');
  }
  // =========================================================================
  // Final da rotina
  // -----------------

  // =========================================================================
  // Montagem da sele��o de escolas
  // -------------------------------------------------------------------------
  function SelecaoEscolaParticular($label, $accesskey, $hint = null, $chave, $chaveaux, $campo, $restricao = null, $atributo = null) {
    extract($GLOBALS);
    $SQL = "SELECT distinct a.sq_cliente cliente, a.sq_tipo_cliente, case b.tipo when '1' then upper(a.ds_username) else a.ds_cliente end ds_cliente, b.tipo " . $crlf .
    "  FROM sbpi.CLIENTE            a" . $crlf .
    "      inner join sbpi.Tipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente) " . $crlf .
    "      inner join sbpi.Particular_Calendario c on (a.sq_cliente = c.sq_cliente) " . $crlf .
    " WHERE PUBLICA = 'N' and upper(a.ds_username) <> 'SBPI' ";

    if ($chaveaux > "") {
      $SQL .= '   and IsNull(sq_cliente_pai,0) = ' . $chaveaux . $crlf;
    }
    $SQL .= 'ORDER BY b.tipo, ds_cliente' . $crlf;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);
    if (is_null($hint)) {
      ShowHTML('         <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_disabled . " " . $atributo . '>');
    } else {
      ShowHTML('          <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_disabled . " " . $atributo . '>');
    }
    ShowHTML('          <option value="">---');
    foreach ($RS as $row) {
      if (doubleval(nvl(f($row, "cliente"), -1)) == doubleval(nvl($chave, -1))) {
        ShowHTML('          <option value="' . f($row, "cliente") . '" SELECTED>' . f($row, "ds_cliente"));
      } else {
        ShowHTML('          <option value="' . f($row, "cliente") . '">' . f($row, "ds_cliente"));
      }
    }
    ShowHTML('          </select>');
  }
  // =========================================================================
  // Final da rotina
  // -------------------------------------------------------------------------

  function ExibeOperacao($chave) {
    switch ($chave) {
      Case 0 :
        $ope = "Consulta";
        break;
      Case 1 :
        $ope = "Inclus�o";
        break;
      Case 2 :
        $ope = "Altera��o";
        break;
      Case 3 :
        $ope = "Exclus�o";
        break;
      default :
        $ope = "Erro";
    }
    return $ope;
  }

  function ExibeBlocoDados($chave) {
    switch ($chave) {
      Case 135 :
        $ope = "Arquivos";
        break;
      Case 133 :
        $ope = "Calend�rio";
        break;
      Case 220 :
        $ope = "Dados adicionais";
        break;
      Case 123 :
        $ope = "Dados b�sicos";
        break;
      Case 144 :
        $ope = "Dados do site";
        break;
      Case 221 :
        $ope = "Mensagens";
        break;
      Case 127 :
        $ope = "Modalidades";
        break;
      Case 222 :
        $ope = "Not�cias";
        break;
      default :
        $ope = "Consulta";
    }
    return $ope;
  }

  function ImprimeLinha($p_acesso, $p_consulta, $p_inc, $p_alt, $p_exc, $p_chave) {
    extract($GLOBALS);
    $O = $_REQUEST['O'];
    If ($O == "L")
      ShowHTML('          <td align="right"><font size="1"><a class="hl" href="javascript:lista(\'' . $p_chave . '\',-1);" onMouseOver="window.status=\'Exibe os registros de log.\'; return true;" onMouseOut="window.status=\'\'; return true;">' . FormatNumber($p_acesso, 0) . '</a>&nbsp;</font></td>');
    Else
      ShowHTML('         <td align="right"><font size="1"> ' . FormatNumber($p_acesso, 0) . '&nbsp;</font></td>');
    If ($p_consulta > 0 and $O == "L")
      ShowHTML('          <td align="right"><font size="1"><a class="hl" href="javascript:lista(\'' . $p_chave . '\', 0);" onMouseOver="window.status=\'Exibe os registros de log.\'; return true;" onMouseOut="window.status=\'\'; return true;">' . FormatNumber($p_consulta, 0) . '</a>&nbsp;</font></td>');
    Else
      ShowHTML('         <td align="right"><font size="1"> ' . FormatNumber($p_consulta, 0) . '&nbsp;</font></td>');
    If ($p_inc > 0 and $O == "L")
      ShowHTML('          <td align="right"><font size="1"><a class="hl" href="javascript:lista(\'' . $p_chave . '\', 1);" onMouseOver="window.status=\'Exibe os registros de log.\'; return true;" onMouseOut="window.status=\'\'; return true;">' . FormatNumber($p_inc, 0) . '</a>&nbsp;</font></td>');
    Else
      ShowHTML('         <td align="right"><font size="1"> ' . FormatNumber($p_inc, 0) . '&nbsp;</font></td>');
    If ($p_alt > 0 and $O == "L")
      ShowHTML('          <td align="right"><font size="1"><a class="hl" href="javascript:lista(\'' . $p_chave . '\', 2);" onMouseOver="window.status=\'Exibe os registros de log.\'; return true;" onMouseOut="window.status=\'\'; return true;">' . FormatNumber($p_alt, 0) . '</a>&nbsp;</font></td>');
    Else
      ShowHTML('         <td align="right"><font size="1"> ' . FormatNumber($p_alt, 0) . '&nbsp;</font></td>');
    If ($p_exc > 0 and $O == "L")
      ShowHTML('          <td align="right"><font size="1"><a class="hl" href="javascript:lista(\'' . $p_chave . '\', 3);" onMouseOver="window.status=\'Exibe os registros de log.\'; return true;" onMouseOut="window.status=\'\'; return true;">' . FormatNumber($p_exc, 0) . '</a>&nbsp;</font></td>');
    Else
      ShowHTML('         <td align="right"><font size="1"> ' . FormatNumber($p_exc, 0) . '&nbsp;</font></td>');

    ShowHTML("        </tr>");
  }

  // =========================================================================
  // Montagem da sele��o de regionais de ensino
  // -------------------------------------------------------------------------
  function SelecaoRegional($label, $accesskey, $hint = null, $chave, $chaveAux, $campo, $restricao = null, $atributo = null) {
    extract($GLOBALS);
    if (substr($_SESSION["username"], 0, 2) == 'RE') {
      $SQL = "SELECT  a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " . $crlf .
             "  FROM  sbpi.CLIENTE a " . $crlf .
             " WHERE  a.sq_cliente = " . $CL . $crlf .
             "ORDER BY a.ds_cliente ";
    } else {
      $SQL = "SELECT a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente " . $crlf .
             "  from sbpi.CLIENTE a " . $crlf .
             "       inner join sbpi.Tipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente) " . $crlf .
             " WHERE b.tipo = 2 " . $crlf .
             "    OR (b.tipo = 1 and a.sq_cliente_pai is null) " . $crlf .
             "ORDER BY b.tipo, a.ds_cliente" . $crlf;
    }
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    if (is_null($hint)) {
      ShowHTML('          <td valign="top"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_disabled . ' ' . $atributo . '>');
    } else {
      ShowHTML('          <td valign="top" title="' . $hint . '"><font size="1"><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_disabled . '  ' . $atributo . '>');
    }

    if ($restricao == 'CADASTRO') {
      ShowHTML('          <option value="">---');
    } else {
      if (Count($RS) > 0) {
        ShowHTML('          <option value="" SELECTED >Tanto faz');
      }
    }
    foreach ($RS as $row) {
      if ((nvl(f($row, "sq_cliente"), -1)) == (nvl($chave, -1))) {
        ShowHTML('          <option value="' . f($row, "sq_cliente") . '" SELECTED >' . f($row, "ds_cliente"));
      } else {
        ShowHTML('          <option value="' . f($row, "sq_cliente") . '">' . f($row, "ds_cliente"));
      }
    }
    ShowHTML('          </select>');
  }
  // =========================================================================
  // Final da rotina
  // -------------------------------------------------------------------------

  function montaCheckBoxCal($cliente) {
    extract($GLOBALS);

    ShowHTML('    <script language=""JavaScript"" type=""text/JavaScript"">' . $crlf);
    ShowHTML(' ok=false;' . $crlf);
    ShowHTML('function CheckAll() {' . $crlf);
    ShowHTML('   if(!ok){' . $crlf);
    ShowHTML('     for (var i=0;i<document.Form.elements.length;i++) {' . $crlf);
    ShowHTML('       var x = document.Form.elements[i];' . $crlf);
    ShowHTML('       if (x.name == \'w_atribuicao[]\') {        ' . $crlf);
    ShowHTML('               x.checked = true;' . $crlf);
    ShowHTML('               ok=true;' . $crlf);
    ShowHTML('           }' . $crlf);
    ShowHTML('       }' . $crlf);
    ShowHTML('   }' . $crlf);
    ShowHTML('   else{' . $crlf);
    ShowHTML('   for (var i=0;i<document.Form.elements.length;i++) {' . $crlf);
    ShowHTML('       var x = document.Form.elements[i];' . $crlf);
    ShowHTML('       if (x.name == \'w_atribuicao[]\') {        ' . $crlf);
    ShowHTML('               x.checked = false;' . $crlf);
    ShowHTML('               ok=false;' . $crlf);
    ShowHTML('           }' . $crlf);
    ShowHTML('       }    ' . $crlf);
    ShowHTML('   }' . $crlf);
    ShowHTML(' }' . $crlf);
    ShowHTML(' </script>' . $crlf);

    $SQL = "select homologado, sq_particular_calendario, nome from sbpi.Particular_Calendario " . $crlf .
    " where ativo <> 'N' and sq_cliente = " . $cliente;
    $RS = db_exec :: getInstanceOf($dbms, $SQL, & $numRows);

    if (count($RS) > 0) {
      ShowHTML('<label><font size="1"><b>Atribui��o:</b></font></label>');
      ShowHTML('<br><a href="javascript:void(null)" onClick="CheckAll();"><font size="1">Marcar/Desmarcar Todos</font></a><br/>');
      foreach ($RS as $row) {
        If (strpos(f($row, "homologado"), "S")) {
          ShowHTML('<br style=""><font size="1"><input disabled type="checkbox" name="homologado" value="' . f($row, "sq_particular_calendario") . '">' . f($row, "nome") . '<font color="#ff3737"> (Homologado pela SEDF)</font></label>');
        } else {
          ShowHTML('<br style=""><font size="1"><input type="checkbox" name="w_atribuicao[]" value="' . f($row, "sq_particular_calendario") . '">' . f($row, "nome") . '</label>');
        }
      }
    }

  }
?>
