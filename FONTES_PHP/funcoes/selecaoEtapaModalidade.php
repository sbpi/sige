<?
// =========================================================================
// Montagem da seleção da etapas/modalidades de ensino
// -------------------------------------------------------------------------
function selecaoEtapaModalidade($label,$accesskey,$hint,$chave,$chaveAux,$campo,$restricao,$atributo,$colspan=1) {
  extract($GLOBALS);
  $sql = "SELECT DISTINCT a.curso sq_especialidade, a.curso ds_especialidade, '1' nr_ordem, 'M' tp_especialidade " . $crlf . 
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
  $rs = db_exec::getInstanceOf($dbms, $sql, &$numRows);
  if (count($rs)>0) {
    $wCont  = 0;
    $wAtual = '';
    //ShowHTML('          <td colspan="'.$colspan.'"><b>'.$label.'</b>');
    foreach($rs as $row)  {
      if ($wAtual=='' || $wAtual <> f($row,'tp_especialidade')) {
        $wAtual = f($row,'tp_especialidade');
        if ($wAtual=='M') {
           if (Nvl($chaveAux,'S')=='P') {
              ShowHTML('<dl>');
              ShowHTML('          <dt><b>Modalidades de ensino</b>: </dt>');
           } else {
              ShowHTML('          <dt><b>Etapas / Modalidades de ensino</b>: </dt>');
           }
        } elseIf ($wAtual=='R') {
           ShowHTML('          <dt><b>Em Regime de Intercomplementaridade</b>: </dt>');
        } else {
           ShowHTML('          <dt><b>Outras</b>: </dt>');
        }
      }
      $l_marcado = 'N';
      $l_chave   = $chave.',';
      while (strpos($l_chave,',')!==false) {
        $l_item  = trim(substr($l_chave,0,strpos($l_chave,',')));
        $l_chave = trim(substr($l_chave,(strpos($l_chave,',')+1),100));
        if ($l_item > '') {if ("'".f($row,'sq_especialidade')."'"==$l_item) $l_marcado = 'S'; }
      }
      if ($l_marcado=='S') { 
        ShowHTML('          <dd><input type="CHECKBOX" name="'.$campo.'" value="\''.f($row,'sq_especialidade').'\'" CHECKED> '.f($row,'ds_especialidade')); 
      } else { 
        ShowHTML('          <dd><input type="CHECKBOX" name="'.$campo.'" value="\''.f($row,'sq_especialidade').'\'" > '.f($row,'ds_especialidade')); 
      }
    }
  }
}
?>
