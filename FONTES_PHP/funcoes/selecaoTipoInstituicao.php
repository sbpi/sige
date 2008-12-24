<?
// =========================================================================
// Montagem da seleção de tipos de instituição
// -------------------------------------------------------------------------
function selecaoTipoInstituicao($label,$accesskey,$hint,$chave,$chaveAux,$campo,$restricao,$atributo,$colspan=1) {
  extract($GLOBALS);
  $sql= 'SELECT * from sbpi.Tipo_Cliente a WHERE a.tipo = 3 ORDER BY a.ds_tipo_cliente' . $crlf;
  $rs = db_exec::getInstanceOf($dbms, $sql, &$numRows);

  if (isset($label)) {
    ShowHTML('          <td colspan="'.$colspan.'" '.((!isset($hint)) ? '' : 'title="'.$hint.'"').'>'.((!isset($hint)) ? '' : 'title="'.$hint.'"').'<b>'.$label.'</b><br><SELECT ACCESSKEY="'.$accesskey.'" CLASS="texto" NAME="'.$campo.'" '.$w_Disabled.' '.$atributo.'>');
  } else {
    echo '<SELECT ACCESSKEY="'.$accesskey.'" CLASS="texto" NAME="'.$campo.'" '.$w_Disabled.' '.$atributo.'>';
  }

  ShowHTML('          <option value="">Todos');
  foreach($rs as $row) {
    if (nvl(f($row,'sq_tipo_cliente'),0)==nvl($chave,0)) {
       ShowHTML('          <option value="'.f($row,'sq_tipo_cliente').'" SELECTED>'.f($row,'ds_tipo_cliente'));
    } else {
       ShowHTML('          <option value="'.f($row,'sq_tipo_cliente').'">'.f($row,'ds_tipo_cliente'));
    }
  }
  ShowHTML('          </select>');
}
?>
