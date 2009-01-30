<?php
// Define o banco de dados a ser utilizado
$_SESSION['DBMS']=1;

$w_dir_volta = '';
$w_dir = '';
$w_pagina = 'default.php?';

include_once('constants.inc');
include_once('jscript.php');
include_once('funcoes.php');
include_once('classes/db/abreSessao.php');
include_once('classes/sp/db_exec.php');
include_once('funcoes/selecaoRegiaoAdm.php');
include_once('funcoes/selecaoRegionalEnsino.php');
include_once('funcoes/selecaoTipoInstituicao.php');
include_once('funcoes/selecaoEtapaModalidade.php');

// =========================================================================
//  /default.php
// ------------------------------------------------------------------------
// Nome     : Alexandre Vinhadelli Papadópolis
// Descricao: Mecanismo de busca
// Mail     : alex@sbpi.com.br
// Criacao  : 18/12/2008 13:02
// Versao   : 1.0.0.0
// Local    : Brasília - DF
// SBPI Consultoria Ltda - Todos os direitos reservados
// -------------------------------------------------------------------------
//
// Declaração de variáveis
$RS=null;

// Abre conexão com o banco de dados
if (isset($_SESSION['DBMS'])) $dbms = abreSessao::getInstanceOf($_SESSION['DBMS']);

// Recupera o filtro selecionado
$P3           = nvl($_REQUEST['P3'],1);
$P4           = nvl($_REQUEST['P4'],$conPageSize);

$p_rede       = nvl($_REQUEST['p_rede'],'S');
$p_regiao     = $_REQUEST['p_rede'];
$p_regional   = $_REQUEST['p_regional'];
$p_local      = $_REQUEST['p_local'];
$p_modalidade = explodeArray($_REQUEST['p_modalidade']);
$p_tipo_inst  = intVal($_REQUEST['p_tipo_inst']);

Main();

// Fecha conexão com o banco de dados
if (isset($_SESSION['DBMS'])) FechaSessao($dbms);

exit;

// =========================================================================
// Rotina de apresentação do mecanismo de busca
// -------------------------------------------------------------------------
function Inicial() {
  extract($GLOBALS);
  ShowHTML('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">');
  ShowHTML('<html xmlns="http://www.w3.org/1999/xhtml">');
  ShowHTML('<head>');
  ShowHTML('   <title>Secretaria de Estado de Educa&ccedil;&atilde;o</title>');
  ShowHTML('   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /> ');
  ShowHTML('   <link href="css/estilo.css" media="screen" rel="stylesheet" type="text/css" />');
  ShowHTML('   <link href="css/print.css"  media="print"  rel="stylesheet" type="text/css" />');
  ShowHTML('   <script language="javascript" src="js/scripts.js"> </script>');
  ShowHTML('   <script src="js/jquery.js"></script>');
  ShowHTML('</head>');
  if ($p_regiao>''||$p_regional>''||$p_local>''||$p_modalidade>'') {
    ShowHTML('<body onLoad="location.href=\'#resultado\'">');
  } else {
    ShowHTML('<body>');
  }
  ShowHTML('<div id="pagina">');
  ShowHTML('  <div id="topo"> </div>');
  ShowHTML('  <div id="busca">');
  ShowHTML('    <div class="data"> '.formataDataEdicao(time(),9).'</div>');
  ShowHTML('    <div class="clear"></div>');
  ShowHTML('  </div>');
  ShowHTML('  <div id="menu">');
  ShowHTML('    <div id="menuTop">');
  ShowHTML('      <div class="esquerda">');
  ShowHTML('        <ul>');
  ShowHTML('          <li><a href="http://www.se.df.gov.br/300/30001001.asp">Secretaria de Educação</a></li>');
  ShowHTML('          <li><a href="http://www.se.df.gov.br/300/30001009.asp">mapa do site</a></li>');
  ShowHTML('          <li class="ultimo"><a class="ultimo" href="http://www.se.df.gov.br/300/30001005.asp">fale conosco</a></li>');
  ShowHTML('        </ul>');
  ShowHTML('      </div>');
  ShowHTML('      <div class="direita">');
  ShowHTML('        <ul>');
  ShowHTML('          <li><a class="aluno" href="http://www.se.df.gov.br/300/30002001.asp">Aluno</a></li>');
  ShowHTML('          <li><a class="educador" href="http://www.se.df.gov.br/300/30003001.asp">Educador</a></li>');
  ShowHTML('          <li><a class="comunidade" href="http://www.se.df.gov.br/300/30004001.asp">Comunidade</a></li>');
  ShowHTML('        </ul>');
  ShowHTML('      </div>');
  ShowHTML('    </div>');
  ShowHTML('    <div id="menuMiddle"> </div>');
  ShowHTML('  </div>');
  ShowHTML('  <div class="clear"></div>');
  ShowHTML('<div id="menuBottom">');
  ShowHTML('  <ul>');
  ShowHTML('    <li><a href="#"><span>Inicial</span></a></li>');
  ShowHTML('    <li><a href="newsletter.php?par=inicial"><span>Assine nosso boletim</span></a></li>');
  ShowHTML('  </ul>');
  ShowHTML('  <div class="clear"></div>');
  ShowHTML('</div>');
  ShowHTML('<div id="conteudo"><h2>Solução Integrada de Gestão Educacional - SIGE</h2>');
  ShowHTML(' <h3>Busca avançada</h3>');
  ShowHTML(' <!-- <form action=" id="formConteudo">-->');
  ShowHTML(' <FORM ACTION="'.$w_dir.$w_pagina.'?EW=121" id="form1" name="form1" METHOD="POST">');
  ShowHTML(' <div id="regional">');
  ShowHTML(' <div class="topo"> </div>');
 
  ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT"><!--  ');
  ShowHTML('function marcaTipo() {  ');
  ShowHTML('  document.form1.submit();');
  ShowHTML('}  ');
  ShowHTML('--></SCRIPT>');

  ShowHTML('        <FORM ACTION="'.$w_dir.$w_pagina.'?EW=121" id="form1" name="form1" METHOD="POST">');

  // Verifica se há alguma escola particular no cadastro. Se não, marca para tratar apenas escolas públicas
  $SQL = "select count(sq_cliente) cadprivada from sbpi.Cliente_Particular";
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  foreach($RS as $row) { $RS = $row; break; }

  //If (f($RS,'cadprivada') > 0) {
  //    If (Nvl($p_rede,'S')=='S') ShowHTML('<input type="radio" name="p_rede" value="S" checked="checked" onclick="marcaTipo();"> Rede pública');  else ShowHTML('<input type="radio" name="p_rede" value="S" onclick="marcaTipo();"> Rede pública');
  //    If ($p_rede=='P')          ShowHTML('<input type="radio" name="p_rede" value="P" checked="checked" onclick="marcaTipo();"> Rede privada');  else ShowHTML('<input type="radio" name="p_rede" value="P" onclick="marcaTipo();"> Rede privada');        
  //    ShowHTML('<br/><br/>');
  //} else {
    ShowHTML('        <input type="Hidden" name="p_rede" value="S">');
  //}

  if ($p_rede=='P') {
    ShowHTML(' <h4>Região administrativa:</h4>');
    selecaoRegiaoAdm(null,null,null,$p_regiao,null,'p_regiao',null,null);
  } else {
    ShowHTML(' <h4>Regional de ensino:</h4>');
    selecaoRegionalEnsino(null,null,null,$p_regional,null,'p_regional',null,null);
  }
  ShowHTML(' <div class="base"></div>');
  ShowHTML(' </div>');
  ShowHTML(' <div id="instituicao">');
  ShowHTML(' <div class="topo"> </div>');
  ShowHTML(' <h4>Tipo de instituição:</h4>');
  selecaoTipoInstituicao(null,null,null,$p_tipo_inst,null,'p_tipo_inst',null,null);
  ShowHTML(' <div class="base"> </div>');
  ShowHTML(' </div>');
  ShowHTML(' <div id="urbana">');
  ShowHTML(' <div class="topo"> </div>');
  
  ShowHTML('         <h4>Localização:</h4><div class="radio_option">');
  if ($p_local=='2') ShowHTML('<input type="radio" name="p_local" value="2" checked="checked"> Rural');        else ShowHTML('<input type="radio" name="p_local" value="2"> Rural');
  if ($p_local=='1') ShowHTML('&nbsp;<input type="radio" name="p_local" value="1" checked="checked"> Urbana'); else ShowHTML('&nbsp;<input type="radio" name="p_local" value="1"> Urbana');
  if ($p_local=='')  ShowHTML('&nbsp;<input type="radio" name="p_local" value=""  checked="checked"> Ambas');  else ShowHTML('&nbsp;<input type="radio" name="p_local" value="" > Ambas');
  ShowHTML(' </div><div class="base"> </div>');
  ShowHTML(' </div>');

  ShowHTML(' <div id="opcoes">');
  ShowHTML(' <div class="topo"> </div>');
  selecaoEtapaModalidade(null,null,null,$p_modalidade,$p_rede,'p_modalidade[]',null,null);
  ShowHTML(' </dd>');
  ShowHTML(' <div class="base"> </div>');
  ShowHTML(' </div>');
  ShowHTML(' <div id="botao">');
  ShowHTML(' <input id="pesquisar" type="button" name="Botao" value="Pesquisar" class="botao" onClick="javascript: document.form1.Botao.disabled=true; document.form1.submit();">');
  ShowHTML(' </form>');
  ShowHTML(' </div> ');
  ShowHTML(' <div id="resultado">');
  if ($p_regiao>''||$p_regional>''||$p_local>''||$p_modalidade>'') {
    $sql= "SELECT DISTINCT 0 Tipo, sbpi.acentos(d.ds_cliente) as ordena, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " . $crlf .
          "       e.ds_logradouro,e.no_bairro,e.nr_cep, e.nr_fone_contato, d.localizacao, 'S' publica, null ds_username, " . $crlf .
          "       r.no_regiao " . $crlf .
          "  from sbpi.Cliente                               d " . $crlf .
          "       INNER      join sbpi.Tipo_Cliente          b ON (d.sq_tipo_cliente  = b.sq_tipo_cliente and " . $crlf .
          "                                                        b.tipo             = '2' " . $crlf .
          "                                                       ) " . $crlf .
          "       LEFT OUTER join sbpi.Cliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " . $crlf .
          "       INNER      join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " . $crlf .
          " where d.ativo <> 'N' and d.publica = 'S' and 'P' <> '".$p_rede."' and d.sq_cliente = ".Nvl($p_regional,0).$crlf;
    if ($p_local>"")     $sql.="    and 0 < (select count(*) from sbpi.Cliente where sq_cliente_pai = d.sq_cliente and localizacao = ".$p_local.") " . $crlf;
    if ($p_tipo_inst>"") $sql.="    and 0 < (select count(*) from sbpi.Cliente where sq_cliente_pai = d.sq_cliente and d.sq_tipo_cliente= ".$p_tipo_inst.") " . $crlf;
    if ($p_modalidade>"") {
       $sql.="    and (0 < (select count(sq_cliente) from sbpi.Especialidade_Cliente where sq_cliente_pai = d.sq_cliente and to_char(sq_especialidade) in (".$p_modalidade.")) or " . $crlf .
             "         0 < (select count(*) from sbpi.Turma_Modalidade  w INNER join sbpi.Turma x ON (w.serie = x.ds_serie) INNER join sbpi.Cliente y ON (x.sq_cliente = y.sq_cliente) where y.sq_cliente_pai = d.sq_cliente and w.curso in (".$p_modalidade.")) " . $crlf .
             "        ) " . $crlf;
    }
    $sql .=" UNION ".$crlf.
          "SELECT DISTINCT 0 Tipo, sbpi.acentos(d.ds_cliente) as ordena, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " . $crlf .
          "       e.ds_logradouro,e.no_bairro,e.nr_cep, e.nr_fone_contato, d.localizacao, 'S' publica, null ds_username, " . $crlf .
          "       r.no_regiao " . $crlf .
          "  from sbpi.Cliente                               d " . $crlf .
          "       LEFT OUTER join sbpi.Cliente_Dados         e ON (d.sq_cliente       = e.sq_cliente) " . $crlf .
          "       INNER      join sbpi.Cliente_Site          a ON (d.sq_cliente       = a.sq_cliente) " . $crlf .
          "       INNER      join sbpi.Tipo_Cliente          b ON (d.sq_tipo_cliente  = b.sq_tipo_cliente and " . $crlf .
          "                                                        b.tipo             = '2' " . $crlf .
          "                                                       ) " . $crlf .
          "       INNER      join sbpi.Cliente               c ON (d.sq_cliente       = c.sq_cliente_pai) " . $crlf .
          "       INNER      join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " . $crlf .
          " where d.ativo <> 'N' and 'P' <> '".$p_rede."' " . $crlf;
    if ($p_local>"")      $sql .="    and d.localizacao    = ".$p_local.$crlf;
    if ($p_regional>"")   $sql .="    and c.sq_cliente_pai = ".$p_regional.$crlf;
    if ($p_tipo_inst >"") $sql .="    and c.sq_tipo_cliente= ".$p_tipo_inst.$crlf;
    if ($p_modalidade>"") {
      $sql.="    and (0 < (select count(*) from sbpi.Especialidade_Cliente w inner join sbpi.Cliente x on (w.sq_cliente = x.sq_cliente) where x.sq_cliente_pai = d.sq_cliente and to_char(w.sq_especialidade) in (".$p_modalidade.")) or " . $crlf .
             "         0 < (select count(*) from sbpi.Turma_Modalidade  w INNER join sbpi.Turma x ON (w.serie = x.ds_serie) INNER join sbpi.Cliente y ON (x.sq_cliente = y.sq_cliente) where y.sq_cliente_pai = d.sq_cliente and w.curso in (".$p_modalidade.")) " . $crlf .
             "        ) " . $crlf;
    }
    $sql .=" UNION ". $crlf .
          "SELECT DISTINCT 1 Tipo, sbpi.acentos(d.ds_cliente) as ordena, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " . $crlf . 
          "       f.endereco logradouro, null no_bairro, null nr_cep, null nr_fone_contato, d.localizacao, d.publica, " . $crlf . 
          "       d.ds_username, r.no_regiao " . $crlf .
          "  from sbpi.Cliente                          d " . $crlf .
          "       INNER join sbpi.Cliente_Particular    f ON (d.sq_cliente       = f.sq_cliente) " . $crlf .
          "       LEFT  join sbpi.Especialidade_cliente g ON (d.sq_cliente       = g.sq_cliente) " . $crlf .
          "       INNER join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " . $crlf .
          " where d.ativo <> 'N' and f.situacao = 1 and d.publica = 'N' and 'S' <> '" . $p_rede . "' " . $crlf;
    if ($p_local>"")      $sql .="    and d.localizacao    = ".$p_local.$crlf;
    if ($p_regional>"")   $sql .="    and d.sq_regiao_adm  = ".$p_regional.$crlf;
    if ($p_tipo_inst >"") $sql .="    and d.sq_tipo_cliente= ".$p_tipo_inst.$crlf;
    if ($p_modalidade >"") {
      $sql.="  and (to_char(g.sq_especialidade) in (".$p_modalidade.")) ".$crlf;
    }
    $sql.=" UNION " . $crlf . 
          "SELECT DISTINCT 1 Tipo, sbpi.acentos(d.ds_cliente) as ordena, d.sq_cliente, d.ds_cliente, d.ds_apelido,d.ln_internet,d.no_municipio,d.sg_uf, " . $crlf . 
          "       e.ds_logradouro, e.no_bairro,e.nr_cep, e.nr_fone_contato, d.localizacao, d.publica, " . $crlf . 
          "       d.ds_username, r.no_regiao " . $crlf . 
          "  from sbpi.Cliente                            d " . $crlf .
          "       INNER join sbpi.Cliente_Dados           e ON (d.sq_cliente       = e.sq_cliente) " . $crlf .
          "       INNER   join sbpi.Regiao_Administrativa r ON (d.sq_regiao_adm    = r.sq_regiao_adm) " . $crlf .
          " where d.ativo <> 'N' and d.publica = 'S' and 'P' <> '".$p_rede."' " . $crlf;
          
    if ($p_local >"")     $sql .="    and d.localizacao    = ".$p_local.$crlf;
    if ($p_regional >"")  $sql .="    and d.sq_cliente_pai = ".$p_regional.$crlf;
    if ($p_tipo_inst >"") $sql .="    and d.sq_tipo_cliente= ".$p_tipo_inst.$crlf;

    if ($p_modalidade>"") { 
       $sql.="    and (0 < (select count(*) from sbpi.Especialidade_Cliente where sq_cliente = d.sq_cliente and to_char(sq_especialidade) in (".$p_modalidade.")) or " . $crlf .
             "         0 < (select count(*) from sbpi.Turma_Modalidade  w INNER join sbpi.Turma x ON (w.serie = x.ds_serie) INNER join sbpi.Cliente y ON (x.sq_cliente = y.sq_cliente) where y.sq_cliente = d.sq_cliente and w.curso in (".$p_modalidade.")) " . $crlf .
             "        ) " . $crlf;
    }
    $sql .=" ORDER BY 1, 2 ".$crlf;
    $RS = db_exec::getInstanceOf($dbms, $sql, &$numRows);
    
    ShowHTML(' <h3>Resultado da Pesquisa</h3>');
    if (count($RS)>0) {
      ShowHTML('          <UL> ');

      $RS1 = array_slice($RS, (($P3-1)*$P4), $P4);
      foreach ($RS1 as $row) {
        if (f($row,'PUBLICA')=='S') {
           ShowHTML('                <font size="2"><b>');
           if (f($row,'ln_internet')>'') {
              if (strpos(strtolower(f($row,'ln_internet')),'http://')!==false) {
                 ShowHTML('                <li><a href="http://'.str_replace('http://','',f($row,'ln_internet')).'" target="_blank">'.f($row,'ds_cliente').'</a></b>');
              } else {
                ShowHTML('                 <li><a href="'.f($row,'ln_internet').'" target="_blank">'.f($row,'ds_cliente').'</a></b>');
              }
           } else {
              ShowHTML('<li>http://'.f($row,'ds_cliente').'</b><li>');
           }
        } else {
           ShowHTML('                <li><a href="modelos/ModPart/default.asp?EW=110&EF=sq_cliente='.f($row,'sq_cliente').'&CL='.f($row,'sq_cliente').'" target="'.f($row,'ds_username').'">'.f($row,'ds_cliente').'</a></b>');
        }
        if (f($row,'tipo')!='0') {
          ShowHTML('<ul>');
          ShowHTML('<li>Telefone: '.f($row,'nr_fone_contato').'</li>');
          ShowHTML('<li>Endereço: '.f($row,'ds_logradouro'));
          if (Nvl(f($row,'no_bairro'),'nulo') <> 'nulo') ShowHTML(' - Bairro: '.f($row,'no_bairro'));
          ShowHTML(' - '.f($row,'no_regiao'));
          $sql2 = "SELECT f.ds_especialidade " . $crlf . 
                 "  from sbpi.Especialidade_cliente e " . $crlf .
                 "       INNER join sbpi.Especialidade f ON (e.sq_especialidade = f.sq_especialidade) " . $crlf .
                 " WHERE e.sq_cliente = ".f($row,'sq_cliente')." " . $crlf .
                 "UNION " . $crlf .
                 "SELECT f.ds_especialidade " . $crlf . 
                 "  from sbpi.Especialidade f " . $crlf .
                 " WHERE upper(ds_especialidade) like '%PROFISSIONAL%' " . $crlf .
                 "   and 0 < (select count(sq_cliente) from sbpi.Particular_curso where sq_cliente = ".f($row,'sq_cliente').") " . $crlf .
                 "UNION " . $crlf .
                 "select distinct ltrim(rtrim(w.curso)) ds_especialidade " . $crlf . 
                 "  from sbpi.Turma_Modalidade     w " . $crlf .
                 "       INNER   join sbpi.Turma   x ON (w.serie = x.ds_serie) " . $crlf .
                 "         INNER join sbpi.Cliente y ON (x.sq_cliente = y.sq_cliente) " . $crlf .
                 " where y.sq_cliente = ".f($row,'sq_cliente') . " ". $crlf .
                 "ORDER BY ds_especialidade ";
          $RS2 = db_exec::getInstanceOf($dbms, $sql2, &$numRows);
          $wCont = 0;
          ShowHTML('<li>Oferta: ');
          foreach($RS2 as $row2) {
             if ($wCont==0) {
               echo strtoupper(f($row2,'ds_especialidade'));
               $wCont = 1;
             } else {
               echo ', '.strtoupper(f($row2,'ds_especialidade'));
             }
          }
          ShowHTML('</li>');
          ShowHTML('</ul></li>');
        }
      }  
      ShowHTML('            </li>');
      ShowHTML('            </ul>');
      ShowHTML('            </div>');
      barra($w_dir.$w_pagina.$par.'&R='.$w_pagina.$par.'&O='.$O.'&P1='.$P1.'&P2='.$P2.'&TP='.$TP.'&SG='.$SG.MontaFiltro('GET'),ceil(count($RS)/$P4),$P3,$P4,count($RS));
    } else {
      ShowHTML('            <p align="justify"><img src="img/ico_educacao.gif" width="16" height="16" border="0" align="center">&nbsp;<b>Nenhuma ocorrência encontrada para opções acima.');
    }
    
  }
  ShowHTML('    </div>');
  ShowHTML('    <div class="clear"/>');
  ShowHTML('    </div>');
  ShowHTML('  </div>');
  ShowHTML('</div>');
  ShowHTML('<div id="rodape">');
  ShowHTML('<!-- Menu Rodape -->');
  ShowHTML('<ul>');
  ShowHTML('<li><a href="http://www.se.df.gov.br">Página Inicial</a></li>');
  ShowHTML('<li><a href="http://www.se.df.gov.br/300/30001005.asp">Fale conosco</a></li>');
  ShowHTML('<li class="ultimo"><a href="http://www.se.df.gov.br/300/30001009.asp">Mapa do site</a></li>');
  ShowHTML('</ul>');
  ShowHTML('  <!-- Fim Menu Rodape -->');
  ShowHTML('  <div class="clear"/>');
  ShowHTML('  <!-- Endereço -->');
  ShowHTML('  <div class="endereco">');
  ShowHTML('  <div class="buriti">');
  ShowHTML('  <p>Anexo do Palácio do Buriti - 9º andar  - Brasília </p>');
  ShowHTML('  <p>Telefone (61) 3224 0016 (61) 3225 1266 | Fax (61) 3901 3171</p>');
  ShowHTML('  </div>');
  ShowHTML('  <div class="buritinga">');
  ShowHTML('  <p>QNG AE Lote 22 bl 05 sala 03 - Taguatinga   Norte</p>');
  ShowHTML('  <p>Telefone (61) 3355 86 30 | Fax 3355 86 94</p>');
  ShowHTML('  </div>');
  ShowHTML('  </div>');
    
  ShowHTML('  <p class="copy">Copyright ® 2000/2008 - SE/GDF - Todos os Direitos Reservados</p>');
  ShowHTML('  <!-- Fim Endereço -->');
  ShowHTML('</div>');
  Rodape();
}

// =========================================================================
// Monta barra de navegação
// -------------------------------------------------------------------------
function barra($p_link,$p_PageCount,$p_AbsolutePage,$p_PageSize,$p_RecordCount) {
  extract($GLOBALS);
  ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT">');
  ShowHTML('  function pagina (pag) {');
  ShowHTML('    document.Barra.P3.value = pag;');
  ShowHTML('    document.Barra.submit();');
  ShowHTML('  }');
  ShowHTML('</SCRIPT>');
  ShowHTML('<left><FORM ACTION="'.$p_link.'" METHOD="POST" name="Barra">');
  ShowHTML('<input type="Hidden" name="P4" value="'.$p_PageSize.'">');
  ShowHTML('<input type="Hidden" name="p_modalidade" value="'.$p_modalidade.'">');
  ShowHTML('<input type="Hidden" name="w_ew" value="'.$w_ew.'">');
  ShowHTML(MontaFiltro('POST'));
  if ($p_PageCount==$p_AbsolutePage) {
    ShowHTML('<br><b>'.($p_RecordCount-(($p_PageCount-1)*$p_PageSize)).'</b> resultados de <b>'.$p_RecordCount.'</b>');
  } else {
    ShowHTML('<br><b>'.$p_PageSize.'</b> resultados de <b>'.$p_RecordCount.'</b>');
  }
  if ($p_PageSize<$p_RecordCount) {
    ShowHTML('<br/>Página ');
    echo '<SELECT CLASS="texto_livre" NAME="P3" SIZE=1 onChange="document.Barra.submit();">';
    for ($i=1; $i<=$p_PageCount; $i++) {
      echo '<option value="'.$i.'" '.(($i==$p_AbsolutePage) ? 'SELECTED' : '').'>'.$i;
    }
    echo '</SELECT>';
    ShowHTML(' de <span class="texto_livre">&nbsp;'.$p_PageCount.'&nbsp;</span>');
    ShowHTML('<p class="pag">');
    if ($p_AbsolutePage>1) {
      ShowHTML('<A TITLE="Primeira página" HREF="javascript:pagina(1)"><span class="primeira">Primeira</span></A>');
      ShowHTML('<A TITLE="Página anterior" HREF="javascript:pagina('.($p_AbsolutePage-1).')"><span class="anterior">Anterior</span></A>');
    } else {
      ShowHTML('<span class="primeiraOff">Primeira</span>');
      ShowHTML('<span class="anteriorOff">Anterior</span>');
    }
    if ($p_PageCount==$p_AbsolutePage) {
      ShowHTML('<span class="proximaOff">Próxima</span>');
      ShowHTML('<span class="ultimaOff">Última</span>');
    } else {
      ShowHTML('<A TITLE="Página seguinte" HREF="javascript:pagina('.($p_AbsolutePage+1).')"><span class="proxima">Próxima</span></A>');
      ShowHTML('<A TITLE="Última página" HREF="javascript:pagina('.$p_PageCount.')"><span class="ultima">Última</span></A>');
    }
    ShowHTML('</p>');
  } else {
    ShowHTML('<br/>Página <b>'.$p_AbsolutePage.'</b> de <b>'.$p_PageCount.'</b>');
  }
  ShowHtml('</FORM></center>');
}

// =========================================================================
// Rotina principal
// -------------------------------------------------------------------------
function Main() {
  extract($GLOBALS);
  // Monta o formulário de autenticação apenas para a SBPI
  Inicial();
}
?>