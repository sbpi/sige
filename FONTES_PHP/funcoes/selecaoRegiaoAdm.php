  <?

  // =========================================================================
  // Montagem da seleção de região administrativa
  // -------------------------------------------------------------------------
  function selecaoRegiaoAdm($label, $accesskey, $hint, $chave, $chaveAux, $campo, $restricao, $atributo, $colspan = 1) {
    extract($GLOBALS);
    $sql = 'SELECT distinct b.sq_regiao_adm, b.no_regiao nome ' . $crlf .
    '  from sbpi.CLIENTE a ' . $crlf .
    '       inner join sbpi.Regiao_Administrativa b on (a.sq_regiao_adm = b.sq_regiao_adm) ' . $crlf .
    'ORDER BY b.no_regiao' . $crlf;
    $rs = db_exec :: getInstanceOf($dbms, $sql, & $numRows);

    if (isset ($label)) {
      ShowHTML('          <td colspan="' . $colspan . '" ' . ((!isset ($hint)) ? '' : 'title="' . $hint . '"') . '><b>' . $label . '</b><br><SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_Disabled . ' ' . $atributo . '>');
    } else {
      echo '<SELECT ACCESSKEY="' . $accesskey . '" CLASS="STS" NAME="' . $campo . '" ' . $w_Disabled . ' ' . $atributo . '>';
    }

    ShowHTML('          <option value="">Todas');
    foreach ($rs as $row) {
      if (nvl(f($row, 'sq_regiao_adm'), 0) == nvl($chave, 0)) {
        ShowHTML('          <option value="' . f($row, 'sq_regiao_adm') . '" SELECTED>' . f($row, 'nome'));
      } else {
        ShowHTML('          <option value="' . f($row, 'sq_regiao_adm') . '">' . f($row, 'nome'));
      }
    }
    ShowHTML('          </select>');
  }
  ?>
