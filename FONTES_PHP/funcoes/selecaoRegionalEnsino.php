<?
// =========================================================================
// Montagem da seleção de regionais de ensino
// -------------------------------------------------------------------------
function selecaoRegionalEnsino($label,$accesskey,$hint,$chave,$chaveAux,$campo,$restricao,$atributo,$colspan=1) {
  extract($GLOBALS);
  $sql= 'SELECT a.sq_cliente, a.sq_tipo_cliente, a.ds_cliente as nome ' . $crlf .
        '  from sbpi.CLIENTE a ' . $crlf .
        '       inner join sbpi.Tipo_Cliente b on (a.sq_tipo_cliente = b.sq_tipo_cliente and b.tipo = 2) ' . $crlf .
        'ORDER BY a.ds_cliente' . $crlf;
  $rs = db_exec::getInstanceOf($dbms, $sql, &$numRows);

  if (isset($label)) {
    ShowHTML('          <td colspan="'.$colspan.'" '.((!isset($hint)) ? '' : 'title="'.$hint.'"').'>'.((!isset($hint)) ? '' : 'title="'.$hint.'"').'<b>'.$label.'</b><br><SELECT ACCESSKEY="'.$accesskey.'" CLASS="texto" NAME="'.$campo.'" '.$w_Disabled.' '.$atributo.'>');
  } else {
    echo '<SELECT ACCESSKEY="'.$accesskey.'" CLASS="texto" NAME="'.$campo.'" '.$w_Disabled.' '.$atributo.'>';
  }

  ShowHTML('          <option value="">Todas');
  foreach($rs as $row) {
    if (nvl(f($row,'sq_cliente'),0)==nvl($chave,0)) {
       ShowHTML('          <option value="'.f($row,'sq_cliente').'" SELECTED>'.f($row,'nome'));
    } else {
       ShowHTML('          <option value="'.f($row,'sq_cliente').'">'.f($row,'nome'));
    }
  }
  ShowHTML('          </select>');
}
?>
