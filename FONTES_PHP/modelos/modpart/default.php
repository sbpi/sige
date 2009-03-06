<?php
// Define o banco de dados a ser utilizado
$_SESSION['DBMS']=1;

$w_dir_volta = '../../';
$w_dir = 'modelos/modpart/';
$w_pagina = 'default.php?par=';


include_once($w_dir_volta.'constants.inc');
include_once($w_dir_volta.'jscript.php');
include_once($w_dir_volta.'funcoes.php');
include_once($w_dir_volta.'classes/db/abreSessao.php');
include_once($w_dir_volta.'classes/sp/db_exec.php');

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
$RS=null;

// Abre conexão com o banco de dados
if (isset($_SESSION['DBMS'])) $dbms = abreSessao::getInstanceOf($_SESSION['DBMS']);

// Carrega variáveis de controle
$CL           = intVal($_REQUEST['CL']);
$par          = substr(strtoupper($_REQUEST['par']),0,30);
$O            = substr(strtoupper($_REQUEST['O']),0,1);
$R            = substr(strtoupper($_REQUEST['R']),0,30);
$SG           = substr(strtoupper($_REQUEST['SG']),0,30);
$P3           = intVal(nvl($_REQUEST['P3'],1));
$P4           = intVal(nvl($_REQUEST['P4'],$conPageSize));

  ShowHTML('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">');
  ShowHTML('<html xmlns="http://www.w3.org/1999/xhtml">');
  ShowHTML('<head>');
  ShowHTML('<BASE HREF="'.$conRootSIW.'">');
  ShowHTML('   <title>Secretaria de Estado de Educa&ccedil;&atilde;o</title>');
  ShowHTML('   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /> ');
  ShowHTML('   <link href="css/estilo.css" media="screen" rel="stylesheet" type="text/css" />');
  ShowHTML('   <link href="css/print.css"  media="print"  rel="stylesheet" type="text/css" />');
  ShowHTML('   <script language="javascript" src="js/scripts.js"> </script>');
  ShowHTML('   <script src="js/jquery.js"></script>');
  ShowHTML('</head>');
  ShowHTML('<body>');
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
  ShowHTML('      <li><a '.(($par=='INICIAL') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'inicial&CL='.$CL.'" >Inicial</a> </li>');
  // ShowHTML('      <li><a '.(($par=='QUEM') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'quem&CL='.$CL.'" id="link1">Quem somos</a> </li>');
  // ShowHTML('      <li><a '.(($par=='FALE') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'fale&CL='.$CL.'" >Fale conosco </a></li>');
  // ShowHTML('      <li><a '.(($par=='PROJETO') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'projeto&CL='.$CL.'&O=1" >Projeto</a></li>');
  // ShowHTML('      <li><a '.(($par=='NOTICIA') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'noticia&CL='.$CL.'" id="link2">Notícias</a> </li>');
  ShowHTML('      <li><a '.(($par=='CALENDARIO') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'calendario&CL='.$CL.'" id="link5">Calendário</a>');
  ShowHTML('      <li><a '.(($par=='ARQUIVO') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'arquivo&CL='.$CL.'" >Arquivos (<i>download</i>)</a></li>');
  // ShowHTML('      <li><a '.(($par=='ALUNO') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'aluno&O=0&CL='.$CL.'" >Alunos</a></li>');
  ShowHTML('      <li><a '.(($par=='OFERTA') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'oferta&CL='.$CL.'&O=1" >Oferta</a></li>');
  //ShowHTML '      <li><a target="_blank" href="http://siade.cesgranrio.org.br" >Questionário SIADE</a></li>');
  // ShowHTML ('      <li><a target="_blank" href="http://www.prodatadf.com.br/diarioeletronico/" >Diário de Classe</a></li>');
  ShowHTML('  </ul>');
  ShowHTML('  <div class="clear"></div>');
  ShowHTML('</div>');
  $SQL = "SELECT a.sq_cliente, a.ds_cliente, b.ds_mensagem from sbpi.Cliente a inner join sbpi.cliente_site b on (a.sq_cliente = b.sq_cliente)WHERE a.sq_cliente = ".$CL;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  foreach($RS as $row) { $RS = $row; break; }
  ShowHTML('  <div id="conteudo"><h2>'.f($RS,'ds_cliente').'</h2>');

  Main();
  ShowHTML('  </div>');
  ShowHTML('  <div class="clear"/></div>');
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
  
// Fecha conexão com o banco de dados
if (isset($_SESSION['DBMS'])) FechaSessao($dbms);

exit;

// =========================================================================
// Tela de abertura
// -------------------------------------------------------------------------
function Inicial() {
  extract($GLOBALS);
         
 $SQL = "SELECT b.ds_cliente, a.cnpj_escola, a.mantenedora, a.cnpj_executora, a.codinep, a.diretor, a.secretario, " . $crlf . 
        "       coalesce(to_char(a.vencimento,'dd/mm/yyyy'),'Sem informação') vencimento, coalesce(to_char(a.primeiro_credenc,'dd/mm/yyyy'),'Sem informação') primeiro_credenc, " . $crlf . 
        "       a.endereco, b.no_municipio, b.sg_uf, " . $crlf . 
        "       telefone_1, " . $crlf . 
        "       telefone_2, " . $crlf . 
        "       a.email_1, a.email_2, a.cep, " . $crlf . 
        "       fax, " . $crlf .  
        "       c.no_regiao " . $crlf .  
        "from sbpi.cliente_particular                 a " . $crlf . 
        "     INNER   join sbpi.Cliente               b on (a.sq_cliente   = b.sq_cliente) " . $crlf . 
        "       INNER join sbpi.Regiao_Administrativa c on (b.sq_regiao_adm = c.sq_regiao_adm) " . $crlf . 
        " where a.sq_cliente = " . $CL;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);

  ShowHTML ('IDENTIFICAÇÃO:<hr height="1">');

  if (count($RS)>0) {
  foreach($RS as $row){

  ShowHTML ('<table width="100%" border="0" cellspacing="1" cellpadding="1">');
  ShowHTML ('  <tr valign="top">');
  ShowHTML ('    <td align="right" width="30%" nowrap="nowrap"><b>Nome da Instiuição:');
  ShowHTML ('    <td align="left" width="70%" >' . f($row, "ds_cliente")); 
  ShowHTML ('  </tr>');
  ShowHTML ('  <tr valign="top">');
  ShowHTML ('    <td align="right" width="30%" nowrap="nowrap"><b>Mantenedora:');
  ShowHTML ('    <td align="left" width="70%" >' . Nvl(f($row, "mantenedora"),"Sem informação")); 
  ShowHTML ('  </tr>');
  ShowHTML ('  <tr valign="top">');
  ShowHTML ('    <td align="right" width="30%"><b>Diretor(a):');
  ShowHTML ('    <td align="left" width="70%">' . Nvl(f($row, "diretor"),"Sem informação")); 
  ShowHTML ('  </tr>');
  ShowHTML ('  <tr valign="top">');
  ShowHTML ('    <td align="right" width="30%"><b>Secretário(a):');
  ShowHTML ('    <td align="left" width="70%">' . Nvl(f($row, "secretario"),"Sem informação"));
  ShowHTML ('  </tr>');
  ShowHTML ('  <tr>');
  ShowHTML ('    <td align="right" nowrap="nowrap" width="30%" ><b>Data do primeiro credenciamento:</b>');
  ShowHTML ('    <td><table width="1%" border=0 cellpadding=0 cellspacing=0 bgcolor="#DFDFDF"><tr><td width="1%" nowrap style="border: solid 1px;"><font size="2"><b>' . f($row, "primeiro_credenc") . '</b></td></tr></table>');
  ShowHTML ('  </tr>');
  ShowHTML ('  <tr>');
  ShowHTML ('    <td align="right" nowrap="nowrap" width="30%" ><b>Validade credenciamento:</b>');
  ShowHTML ('    <td><table width="1%" border=0 cellpadding=0 cellspacing=0 bgcolor="#DFDFDF"><tr><td width="1%" nowrap style="border: solid 1px;"><font size="2"><b>' . f($row, "vencimento") . '</b></td></tr></table>');
  ShowHTML ('  </tr>');
  ShowHTML ('  </TABLE>');

  ShowHTML ('<br>LOCALIZAÇÃO E CONTATOS:<hr height="1">');
  ShowHTML ('<table width="100%" border="0" cellspacing="1" cellpadding="1">');
    ShowHTML ('  <tr valign="top">');
    ShowHTML ('    <td width="30%" align="right"><b>Endereço:');
    ShowHTML ('    <td width="70%" align="left">' . f($row, "endereco"));
    ShowHTML ('  </tr>');
    ShowHTML ('  <tr valign="top">');
    ShowHTML ('    <td width="30%" align="right"><b>Região Administrativa:');
    ShowHTML ('    <td width="70%" align="left">' . f($row, "no_regiao"));
    ShowHTML ('  </tr>');
    ShowHTML ('  <tr>');
    ShowHTML ('    <td width="30%" align="right" valign="top"><b>Telefones:');
    If(Nvl(f($row, "telefone_1"),'') != '' && Nvl(f($row, "telefone_2"),'') != '') {
      ShowHTML ('    <td width="70%" align="left"> ' . f($row, "telefone_1") . " / " . f($row, "telefone_2"));
    } elseif( Nvl(f($row, "telefone_1"),'') != '' && Nvl(f($row, "telefone_2"),'')  == '' ){
      ShowHTML ('    <td width="70%" align="left"> ' . f($row, "telefone_1"));
    } elseif(Nvl(f($row, "telefone_2"),'') != '' and Nvl(f($row, "telefone_1"),'')  == '') {
      ShowHTML ('    <td width="70%" align="left"> ' . f($row, "telefone_2"));
    } else {
      ShowHTML ('    <td width="70%" align="left">Sem informação');
    } 
    ShowHTML ('  </tr>');
    ShowHTML ('  <tr valign="top">');
    ShowHTML ('    <td width="30%" align="right"><b>Fax:');
    ShowHTML ('    <td width="70%" align="left">' . nvl(f($row, "fax"),"Sem informação"));
    ShowHTML ('  </tr>');
    ShowHTML ('  <tr valign="top">');
    ShowHTML ('    <td width="30%" align="right"><b>E-mails:');
    if( Nvl(f($row, "email_1"),'')  != '' and Nvl(f($row, "email_2"),'') != '' ){
      ShowHTML ('    <td width="70%" align="left"> ' . f($row, "email_1") . " / " . f($row, "email_2"));
    } elseif(Nvl(f($row, "email_1"),'')  != '' and Nvl(f($row, "email_2"),'')  == '') {
      ShowHTML ('    <td width="70%" align="left"> ' . f($row, "email_1"));
    } elseif(Nvl(f($row, "email_2"),'')  != '' and Nvl(f($row, "email_1"),'')  == '') {
      ShowHTML ('    <td width="70%" align="left"> ' . f($row, "email_2"));
    } else {
      ShowHTML ('    <td width="70%" align="left">Sem informação');
    } 
    ShowHTML ('  </tr>');

  ShowHTML ('  </TABLE>');
  
  }

  }


}

// =========================================================================
// Tela de Quem somos
// -------------------------------------------------------------------------
function quem() {
  extract($GLOBALS);
  $SQL = "SELECT b.ds_institucional, b.ds_diretorio, e.sq_cliente_foto, e.ln_foto, e.ds_foto ".$crlf.
         "  from sbpi.Cliente                    a ".$crlf.
         "       INNER   join sbpi.CLIENTE_SITE  b on (a.sq_cliente = b.sq_cliente) ".$crlf.
         "         INNER join sbpi.Modelo        c on (b.sq_modelo  = c.sq_modelo) ".$crlf.
         "       INNER   join sbpi.Cliente_Dados d on (a.sq_cliente = d.sq_cliente) ".$crlf.
         "       LEFT    join sbpi.Cliente_Foto  e on (a.sq_cliente = e.sq_cliente and e.tp_foto = 'Q') ".$crlf.
         " WHERE a.sq_cliente = ".$CL." ".$crlf.
         "ORDER BY e.nr_ordem";
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  
  $i = 0;
  if (count($RS)>0) {
    foreach($RS as $row) {
      if ($i==0) {
        if (nvl(f($row,'ds_institucional'),'')!='') {
          ShowHTML('        <p class="chamada">'.crlf2br(f($row,'ds_institucional')).'</p>');
        } else {
          ShowHTML('        <p class="chamada">');
        }
      
        if (nvl(f($row,'sq_cliente_foto'),'')!='') {
          ShowHTML('        <ul class="fotos">');
        }
        $i++;
      }
      if (nvl(f($row,'sq_cliente_foto'),'')!='') {
        ShowHTML('<li><a href="'.f($row,'ds_diretorio').'/'.f($row,'ln_foto').'" target="_blank" title="Clique sobre a imagem para ampliar"><img align="top" class="foto" src="'.f($row,'ds_diretorio').'/'.f($row,'ln_foto').'" width="302" height="206"><br>'.f($row,'ds_foto').'</a></li>');
      }
    }
  }
}

// =========================================================================
// Tela de Fale conosco
// -------------------------------------------------------------------------
function fale() {
  extract($GLOBALS);
  $SQL = "SELECT c.no_diretor, c.no_secretario, c.ds_logradouro, c.no_bairro, c.nr_cep, b.no_contato_internet, b.ds_email_internet, b.nr_fone_internet, b.nr_fax_internet, a.no_municipio, a.sg_uf ".$crlf.
         "  from sbpi.Cliente                    a ".$crlf.
         "       INNER   join sbpi.CLIENTE_SITE  b on (a.sq_cliente = b.sq_cliente) ".$crlf.
         "       INNER   join sbpi.Cliente_Dados c on (a.sq_cliente = c.sq_cliente) ".$crlf.
         " WHERE a.sq_cliente = ".$CL." ".$crlf;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  
  ShowHTML('<tr><td>');

  ShowHTML('<p align="justify">Informações, sugestões e reclamações podem ser feitas utilizando os dados abaixo:</p>');
  
  foreach($RS as $row) {
    ShowHTML('<table border="0" cellspacing="3">');
    ShowHTML('  <tr>');
    ShowHTML('    <td width="10%">');
    ShowHTML('    <td width="20%" align="right"><b>Diretor(a):');
    ShowHTML('    <td width="70%" align="left">'.f($row,'no_diretor') );
    ShowHTML('  </tr>');
    ShowHTML('  <tr>');
    ShowHTML('    <td width="10%">');
    ShowHTML('    <td width="20%" align="right"><b>Secretário(a):');
    ShowHTML('    <td width="70%" align="left">'.f($row,'no_secretario') );
    ShowHTML('  </tr>');
    ShowHTML('  <tr>');
    ShowHTML('    <td width="10%">');
    ShowHTML('    <td width="20%" align="right"><b>Contato:</b>');
    ShowHTML('    <td width="70%" align="left">'.f($row,'no_contato_internet'));
    if (nvl(f($row,'ds_email_internet') ,'')!='') {
       ShowHTML('  </tr>');
       ShowHTML('  <tr>');
       ShowHTML('    <td width="10%">');
       ShowHTML('    <td width="20%" align="right">');
       ShowHTML('    <td width="70%" align="left">(<a href="mailto:'.f($row,'ds_email_internet').'">'.f($row,'ds_email_internet').'</a>)');
    }
    ShowHTML('  </tr>');
    ShowHTML('  <tr>');
    ShowHTML('    <td width="10%">');
    ShowHTML('    <td width="20%" align="right"><b>Telefone:');
    ShowHTML('    <td width="70%" align="left">'.f($row,'nr_fone_internet'));
    if (nvl(f($row,'nr_fax_internet') ,'')!='') {
       ShowHTML('    <b> Fax: </b>'.f($row,'nr_fax_internet'));
    }
    ShowHTML('  </tr>');
    ShowHTML('  <tr>');
    ShowHTML('    <td width="10%">');
    ShowHTML('    <td width="20%" align="right"><b>Endereço:');
    ShowHTML('    <td width="70%" align="left">'.f($row,'ds_logradouro'));
    ShowHTML('  </tr>');
    if (nvl(f($row,'no_bairro') ,'')!='') {
       ShowHTML('  <tr>');
       ShowHTML('    <td width="10%">');
       ShowHTML('    <td width="20%" align="right"><b>Bairro:');
       ShowHTML('    <td width="70%" align="left">'.f($row,'no_bairro'));
       ShowHTML('  </tr>');
    }
    ShowHTML('  <tr>');
    ShowHTML('    <td width="10%">');
    ShowHTML('    <td width="20%" align="right"><b>Município:');
    ShowHTML('    <td width="70%" align="left">'.f($row,'no_municipio').'-'.f($row,'sg_uf').'&nbsp;&nbsp;<b>CEP:</b> '.f($row,'nr_cep'));
    ShowHTML('  </tr>');
    ShowHTML('</table>');
  }
  ShowHTML('</tr>');
}

// =========================================================================
// Tela de Projeto
// -------------------------------------------------------------------------
function projeto() {
  extract($GLOBALS);
  $SQL = "SELECT b.ds_diretorio, b.ln_prop_pedagogica ".$crlf.
         "  from sbpi.Cliente                    a ".$crlf.
         "       INNER   join sbpi.CLIENTE_SITE  b on (a.sq_cliente = b.sq_cliente) ".$crlf.
         "         INNER join sbpi.Modelo        c on (b.sq_modelo  = c.sq_modelo) ".$crlf.
         "       INNER   join sbpi.Cliente_Dados d on (a.sq_cliente = d.sq_cliente) ".$crlf.
         " WHERE a.sq_cliente = ".$CL." ".$crlf;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  
  ShowHTML('<tr><td>');

  ShowHTML('<p align="justify">Informações, sugestões e reclamações podem ser feitas utilizando os dados abaixo:</p>');
  
  foreach($RS as $row) {
    if (nvl(f($row,'ln_prop_pedagogica') ,'')!='') {
       ShowHTML('    <font size=1><b>');
       if (strpos(strtolower(f($row,'ln_prop_pedagogica')),'.pdf')!==false) {
         ShowHTML('        A exibição do arquivo exige que o Acrobat Reader tenha sido instalado em seu computador.');
         ShowHTML('        <br>Se o arquivo não for exibido no quadro abaixo, clique <a href="http://www.adobe.com.br/products/acrobat/readstep2.html" target="_blank">aqui</a> para instalar ou atualizar o Acrobat Reader.');
       } else {
         ShowHTML('        A exibição do arquivo exige o editor de textos Word ou equivalente.');
         ShowHTML('        <br>Se o arquivo não for exibido no quadro abaixo, verifique se o Word foi corretamente instalado em seu computador.');
       }
       ShowHTML('<table align="center" width="100%" cellspacing=0 style="border: 1px solid rgb(0,0,0);"><tr><td>');
       ShowHTML('    <iframe src="'.f($row,'ds_diretorio').'/'.f($row,'ln_prop_pedagogica').'" width="100%" height="510">');
       ShowHTML('    </iframe>');
       ShowHTML('</table>');
    } else {
      ShowHTML('    <font size=1><b>Projeto não informado.');
    }
  }
}

// =========================================================================
// Tela de Notícias
// -------------------------------------------------------------------------
function noticia() {
  extract($GLOBALS);
  
  
  $wAno      = intVal($_REQUEST['wAno']);  
  $w_noticia = intVal($_REQUEST['w_noticia']);
  
  If ($wAno == ""){
     $wAno = Date("Y");
  }  
  
  If (nvl($w_noticia,'') == ''){
     $SQL = "SELECT a.sq_noticia, a.dt_noticia data, a.ds_titulo ocorrencia, 'Escola' origem " .$crlf.
           "  from sbpi.Noticia_Cliente  a" .$crlf.
           " WHERE sq_cliente   = " . $CL .$crlf.
           "  and a.ativo         = 'S'" .$crlf.
           "  and a.in_exibe         = 'S'" .$crlf.
           "  and sbpi.year(a.dt_noticia) = " . $wAno .$crlf.
           "UNION " .$crlf.
           "SELECT a.sq_noticia, dt_noticia data, ds_titulo ocorrencia, 'Regional' origem " .$crlf.
           "from sbpi.Noticia_Cliente    a" .$crlf.
           "     INNER join sbpi.Cliente b ON (a.sq_cliente = b.sq_cliente) " .$crlf.
           "     INNER join sbpi.Cliente c ON (b.sq_cliente      = c.sq_cliente_pai) " .$crlf.
           "WHERE c.sq_cliente       = " . $CL .$crlf.
           "  and a.ativo         = 'S'" .$crlf.
           "  and a.in_exibe         = 'S'" .$crlf.
           "  and sbpi.year(a.dt_noticia) = " . $wAno .$crlf.
           "UNION " .$crlf.
           "SELECT a.sq_noticia, dt_noticia data, ds_titulo ocorrencia, 'SEDF' origem " .$crlf.
           "from sbpi.Noticia_Cliente    a" .$crlf.
           "     INNER join sbpi.Cliente b ON (a.sq_cliente = b.sq_cliente) " .$crlf.
           "     INNER join sbpi.Cliente c ON (b.sq_cliente      = c.sq_cliente_pai) " .$crlf.
           "     INNER join sbpi.Cliente d ON (c.sq_cliente      = d.sq_cliente_pai) " .$crlf.
           "WHERE d.sq_cliente       = " . $CL .$crlf.
           "  and a.ativo         = 'S'" .$crlf.
           "  and a.in_exibe         = 'S'" .$crlf.
           "  and sbpi.year(a.dt_noticia) = " . $wAno .$crlf.
           "ORDER BY data desc, ocorrencia" . $crlf;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
	 
     ShowHTML ('<tr><td><TABLE align="center" width="95%"border=0 cellSpacing=1> ');
     ShowHTML ('  <TR valign="top" align="center"> ');
     ShowHTML ('    <TD width="2%">&nbsp; ');
     ShowHTML ('    <TD width="15%"><b>Data ');
     ShowHTML ('    <TD width="10%"><b>Origem ');
     ShowHTML ('    <TD width="15%"><b>Ocorrência ');
     ShowHTML ('    <TD width="48%"> ');
     ShowHTML ('  </TR> ');
     ShowHTML ('  <TR> ');
     ShowHTML ('    <TD COLSPAN="6" HEIGHT="1" BGCOLOR="#DAEABD"> ');
     ShowHTML ('  </TR> ');
     $wCont = 0;
  $i = 0;
  if (count($RS)>0) {
    foreach($RS as $row) {
           ShowHTML ('  <TR valign="top"> ');
           ShowHTML ('    <TD>&nbsp; ');
		   ShowHTML ('    <TD align="center">' . formatadataedicao(f($row, 'data')));
           ShowHTML ('    <TD align="center">' . f($row, 'origem'));
           ShowHTML ('    <TD colspan=2>' . f($row, 'ocorrencia') . '&nbsp;&nbsp;&nbsp;<a href="' . $w_dir . $w_pagina . "noticia&CL=" . $CL . "&w_noticia=" . f($row, "sq_noticia") . '"><b>[Ler]</a> ');
           ShowHTML ('  </TR> ');
           $i++;
    } 
  }else{
    ShowHTML ('  <TR valign="top"> ');
    ShowHTML ('    <TD colspan="5"><BR>Nenhuma notícia foi encontrada para o ano ' . $wAno . '</TD> ');
    ShowHTML ('  </TR> ');	
  }
  
  $SQL   = "SELECT sbpi.year(a.dt_noticia) ano " . $crlf . 
           "  from sbpi.Noticia_Cliente  a" . $crlf . 
           " WHERE sq_cliente   = " . $CL . $crlf . 
           "  and a.ativo         = 'S'" . $crlf . 
           "  and a.in_exibe         = 'S'" . $crlf . 
           "  and sbpi.year(a.dt_noticia) <> " . $wAno . $crlf . 
           "UNION " . $crlf . 
           "SELECT sbpi.year(a.dt_noticia) ano " . $crlf . 
           "from sbpi.Noticia_Cliente    a" . $crlf . 
           "     INNER join sbpi.Cliente b ON (a.sq_cliente = b.sq_cliente) " . $crlf . 
           "     INNER join sbpi.Cliente c ON (b.sq_cliente      = c.sq_cliente_pai) " . $crlf . 
           "WHERE c.sq_cliente       = " . $CL . $crlf . 
           "  and a.ativo         = 'S'" . $crlf . 
           "  and a.in_exibe         = 'S'" . $crlf . 
           "  and sbpi.year(a.dt_noticia) <> " . $wAno . $crlf . 
           "UNION " . $crlf . 
           "SELECT sbpi.year(a.dt_noticia) ano " . $crlf . 
           "from sbpi.Noticia_Cliente    a" . $crlf . 
           "     INNER join sbpi.Cliente b ON (a.sq_cliente = b.sq_cliente) " . $crlf . 
           "     INNER join sbpi.Cliente c ON (b.sq_cliente      = c.sq_cliente_pai) " . $crlf . 
           "     INNER join sbpi.Cliente d ON (c.sq_cliente      = d.sq_cliente_pai) " . $crlf . 
           "WHERE d.sq_cliente       = " . $CL . $crlf . 
           "  and a.ativo         = 'S'" . $crlf . 
           "  and a.in_exibe         = 'S'" . $crlf . 
           "  and sbpi.year(a.dt_noticia) <> " . $wAno . $crlf . 
           "ORDER BY 1 desc" . $crlf; 
		   
		   $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
		   
	if (count($RS)>0) {
		ShowHTML ('  <TR valign="top"><TD colspan="5"><br><b>Notícias de outros anos</b><br> ');
		ShowHTML ('  <ul>');  
		foreach($RS as $row) {
			ShowHTML ('    <li><a href="' . $w_dir . $w_pagina . $par ."&CL=" . str_replace('sq_cliente=',$CL,'sq_cliente=') . "&wAno=" . f($row,"ano") . '" id="link2">Notícias de ' . f($row,"ano") . '</a> </li>');    
		}
		ShowHTML ('  </ul>');
        ShowHTML ('  </TR>');
	}
  ShowHTML ('    </TABLE>'); 
 
  }else{
     $SQL = "SELECT * from sbpi.Noticia_Cliente WHERE sq_noticia = " . $w_noticia;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
     foreach($RS as $row) { $RS = $row; break;}
     ShowHTML ('<tr><td><p align="center"><font size="4"><b>' . f($RS,'ds_titulo') . '</b></font><p align="left"> '.crlf2br(f($RS,'ds_noticia')));
     ShowHTML ('<p><center><img src="img/bt_voltar.gif" border=0 valign="center" onClick="history.go(-1)" alt="Volta">');
  }
}

// =========================================================================
// Tela de Calendário
// -------------------------------------------------------------------------
function calendario() {
  extract($GLOBALS);
	
  $wAno         = intVal($_REQUEST['wAno']);  
  $w_noticia    = intVal($_REQUEST['w_noticia']);
  $w_calendario = intVal($_REQUEST['w_calendario']);  
  
  If ($wAno == ""){
      $wAno = Date("Y");
  }
	
  $SQL = "select distinct b.nome, a.sq_particular_calendario, b.ordem " . $crlf . 
         "  from sbpi.Calendario_Cliente a "  . $crlf .
         "       left join sbpi.Particular_Calendario b on (a.sq_particular_calendario = b.sq_particular_calendario) " . $crlf .
         " where a.sq_cliente = " . $CL . " " . $crlf . 
         "   and sbpi.year(dt_ocorrencia) = " . $wAno . " " . $crlf . 
         "   and ativo = 'S' "  . $crlf . 
         "   and homologado = 'S' "  . $crlf . 
         " order by ordem"  . $crlf;
  
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  if (count($RS) > 0){
    ShowHTML ('  <ul>');
    foreach($RS as $row) {
      ShowHTML ('<li title="' . f($row, "nome") . '"><a href="' . $w_dir . $w_pagina . 'calendario' . $strAction . '&w_calendario=' .f($row, 'sq_particular_calendario') .'&CL='. $CL .'#calendario'.'">' . f($row, "nome"));
    }
    ShowHTML ('  </ul>');
    ShowHTML ('  <br/>');
    ShowHTML ('  <br/>');       
  } else {
    ShowHTML ('<p>A instituição não possui calendário(s) homologado(s).</p>');
  }
  if (!nvl(trim($_REQUEST["w_calendario"]),0) == 0){
    ShowHTML ('<a name="calendario">');
          
    $SQL = " SELECT '' cor, b.imagem, a.dt_ocorrencia data, " . $crlf . 
           " a.ds_ocorrencia ocorrencia, 'B' origem from sbpi.Calendario_base a " . $crlf . 
           " left join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) " . $crlf . 
           " WHERE sbpi.year(dt_ocorrencia)=" . $wAno . " " . $crlf . 
           " UNION " . $crlf . 
           " SELECT '#99CCFF' cor, b.imagem, a.dt_ocorrencia data, a.ds_ocorrencia ocorrencia, 'E' origem"  . $crlf . 
           " from sbpi.Calendario_Cliente a left join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data)"  . $crlf . 
           " WHERE a.sq_cliente = " . $CL . $crlf . 
           "   AND sbpi.year(dt_ocorrencia)= " . $wAno . "  ". $crlf . 
           "   AND sq_particular_calendario = " . $w_calendario . $crlf . 
           "ORDER BY data, origem desc, ocorrencia" . $crlf; 
    $RS1 = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    
    if (count($RS1)>0) {
      foreach($RS1 as $row) {
        if(f($row, 'origem') == 'E'){
          $data = f($row, 'data');
          $wDatas  [IntVal(Day(f($row, "data")))] [Month(f($row, "data"))] [ Substr(year(f($row, "data")),2,2)] = f($row, "ocorrencia") . " (Origem: Escola)";
          $wImagem [IntVal(Day(f($row, "data")))] [Month(f($row, "data"))] [ Substr(year(f($row, "data")),2,2)] = f($row, "imagem");
          $wCores  [IntVal(Day(f($row, "data")))] [Month(f($row, "data"))] [ Substr(year(f($row, "data")),2,2)] = f($row, "cor");
        }else if(f($row, 'origem') == 'ER'){
          $wDatas  [IntVal(Day(f($row, "data")))] [Month(f($row, "data"))] [ Substr(year(f($row, "data")),2,2)] = f($row, "ocorrencia") . " (Origem: SEDF)";
          $wImagem [IntVal(Day(f($row, "data")))] [Month(f($row, "data"))] [ Substr(year(f($row, "data")),2,2)] = f($row, "imagem");
          $wCores  [IntVal(Day(f($row, "data")))] [Month(f($row, "data"))] [ Substr(year(f($row, "data")),2,2)] = f($row, "cor");            
        }else{
          $wDatas  [IntVal(Day(f($row, "data")))] [Month(f($row, "data"))] [ Substr(year(f($row, "data")),2,2)] = f($row, "ocorrencia") . " (Origem: Oficial)";
          $wImagem [IntVal(Day(f($row, "data")))] [Month(f($row, "data"))] [ Substr(year(f($row, "data")),2,2)] = f($row, "imagem");            
        }
      }
    }
      
    ShowHTML(' <tr><td><TABLE align="center" width="100%" border=0 cellSpacing=0> ');
    ShowHTML ('<tr align=""center""><td><strong>'. $w_nome .'</strong><br/></td></tr>');
    ShowHTML(' <tr valign="top"> ');
    ShowHTML('   <td>' . MontaCalendario("01" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML('   <td>' . MontaCalendario("02" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML('   <td>' . MontaCalendario("03" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML('   <td>' . MontaCalendario("04" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML(' <tr valign="top"> ');
    ShowHTML('   <td>' . MontaCalendario("05" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML('   <td>' . MontaCalendario("06" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML('   <td>' . MontaCalendario("07" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML('   <td>' . MontaCalendario("08" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML(' <tr valign="top"> ');
    ShowHTML('   <td>' . MontaCalendario("09" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML('   <td>' . MontaCalendario("10" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML('   <td>' . MontaCalendario("11" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML('   <td>' . MontaCalendario("12" . $wAno, $wDatas, $wCores, $wImagem));
    ShowHTML(' </table> ');
      
    ShowHTML(' <tr><td colspan="2"><TABLE width="100%" align="center" border=0 cellSpacing=1>');
    ShowHTML(' <tr valign="middle" align="center">');
    ShowHTML('   <td><font size=1><b>LEGENDA</b></font>');
    ShowHTML('   <td><font size=1><b>FERIADOS</b></font>');
    ShowHTML('   <td><font size=1><b>RECESSOS</b></font>');
    ShowHTML('   <td><font size=1><b>DATAS DA ESCOLA</b></font>');
    ShowHTML('   <td><font size=1><b>DIAS LETIVOS</b></font>');
    ShowHTML(' </tr>');
    ShowHTML('   <TR>');
    ShowHTML('     <TD CHEIGHT="1" BGCOLOR="#DAEABD">');
    ShowHTML('   </TR>');
     
    ShowHTML('<tr valign="top">');

    //Legenda
    $SQL = "SELECT * from sbpi.Tipo_Data where abrangencia <> 'U' order by nome " . $crlf;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);

    if (count($RS)>0) {
        ShowHTML (' <td><TABLE width="90%" align="center" border=0 cellSpacing=1>');
        foreach($RS as $row) { 
            ShowHTML('   <TR VALIGN="TOP"> ');
            ShowHTML('     <TD>&nbsp; ');
            ShowHTML('     <TD><img src="img/' . f($row, "imagem") . '" align="center"> ');
            ShowHTML('     <TD>' . f($row, "nome"));
            ShowHTML('   </TR> ');
        }
        ShowHTML('     </TABLE>');
    }else{
       ShowHTML('  <td>Sem informação');
    }    
      
    //Feriados
    $SQL = "SELECT '' cor, b.imagem, a.dt_ocorrencia data, a.ds_ocorrencia ocorrencia, 'B' origem from sbpi.Calendario_base a left join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data) WHERE sbpi.year(dt_ocorrencia)=" . $wAno . " AND b.sigla <> 'CN' " . $crlf . 
           "ORDER BY data, origem desc, ocorrencia" . $crlf;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
      
    if (count($RS)>0) {
      ShowHTML (' <td><TABLE width="90%" align="center" border=0 cellSpacing=1>');
      foreach($RS as $row) {
        ShowHTML('   <TR VALIGN="TOP">');
        ShowHTML('     <TD>&nbsp;');
        ShowHTML('     <TD>' . Substr(100+Day(f($row, "data")),1,2) . '/' . Substr(100+Month(f($row, "data")),1,2));
        ShowHTML('     <TD>' . f($row, "ocorrencia"));
        ShowHTML('   </TR>');
      }
      ShowHTML('</TABLE>');
    }else{
      ShowHTML('<td>Sem informação');
    }
        
    //Exibe recessos
    $SQL  = "SELECT '#99CCFF' cor, b.imagem, a.dt_ocorrencia data, a.ds_ocorrencia ocorrencia, 'E' origem from sbpi.Calendario_Cliente a inner join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla in ('RE','RA')) WHERE sq_cliente = " . $CL . "  AND sbpi.year(dt_ocorrencia)= " . year(Time()) . " " . $crlf . 
            "UNION " . $crlf . 
            "SELECT '#FFFF99' cor, e.imagem, a.dt_ocorrencia data, a.ds_ocorrencia ocorrencia, 'R' origem " . $crlf . 
            "from sbpi.Calendario_Cliente   a" . $crlf . 
            "     INNER join sbpi.Cliente   b ON (a.sq_cliente = b.sq_cliente) " . $crlf . 
            "     INNER join sbpi.Cliente   c ON (b.sq_cliente      = c.sq_cliente_pai) " . $crlf . 
            "     INNER join sbpi.Cliente   d ON (c.sq_cliente      = d.sq_cliente_pai) " . $crlf . 
            "     INNER  join sbpi.Tipo_Data e ON (a.sq_tipo_data    = e.sq_tipo_data and e.sigla in ('RE','RA')) " . $crlf . 
            "WHERE d.sq_cliente = " . $CL . $crlf . 
            "  AND sbpi.year(dt_ocorrencia) = " . $wAno . " " . $crlf .  
           " AND a.sq_particular_calendario = " . $w_calendario . $crlf .  
           "ORDER BY data, origem desc, ocorrencia" . $crlf;

      $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
              
    if (count($RS)>0) {
      ShowHTML('  <td><TABLE width="90%" align="center" border=0 cellSpacing=1>');
      foreach($RS as $row) {
        ShowHTML('  <TR VALIGN="TOP">');
        ShowHTML('    <TD>&nbsp;');
        ShowHTML('    <TD>' . Substr(100+Day(f($row, "data")),1,2) . '/' . Substr(100+Month(f($row, "data")),1,2));
        ShowHTML('    <TD><font color="#0000FF">' . f($row, "ocorrencia"));
      ShowHTML('  </TR>');
      }
      ShowHTML('    </TABLE></td>');
    }else{
      ShowHTML('  <td>Sem informação');
    }
       
    $SQL = " SELECT '#99CCFF' cor, b.imagem, a.dt_ocorrencia data, 'IE: ' " . $crlf .
           " || a.ds_ocorrencia ocorrencia, 'E' origem from sbpi.Calendario_Cliente a " . $crlf .
           " inner join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla not in ('RE','RA'))" . $crlf .
           " WHERE sq_cliente = " . $CL . "  AND sbpi.year(dt_ocorrencia)= " . year(Time()) . " " . $crlf .
           " AND a.sq_particular_calendario = " . $w_calendario . $crlf . 
           "ORDER BY data, origem desc, ocorrencia" . $crlf;
             
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    
    if (count($RS)>0) {
      ShowHTML('  <td><TABLE width="90%" align="center" border=0 cellSpacing=1>');
      foreach($RS as $row) {
        ShowHTML('  <TR VALIGN="TOP">');
        ShowHTML('    <TD>&nbsp;');
        ShowHTML('    <TD>' . SubStr(100+Day(f($row, "data")),1,2) . '/' . SubStr(100+Month(f($row, "data")),1,2));
        ShowHTML('    <TD><font color="#0000FF">' . f($row, "ocorrencia"));
        ShowHTML('  </TR>');
      }
      ShowHTML('    </TABLE></td>');
    }else{
      ShowHTML('  <td>Sem informação');
    }
    ShowHTML ('  <td valign="top"><TABLE align="center" border=0 cellSpacing=1>');
       
    //Recupera o ano letivo e o período de recesso
    $SQL =  "  select * from " . $crlf . 
            "        (select dt_ocorrencia w_let_ini " . $crlf . 
            "           from sbpi.Calendario_Cliente a inner join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and sbpi.year(a.dt_ocorrencia) = " .  $wAno .  "and b.sigla = 'IA') " . $crlf . 
            "          where sq_cliente = " . $CL . $crlf . 
            "            and a.sq_particular_calendario = " . $_REQUEST["w_calendario"] . $crlf . 
            "        ) a, " . $crlf . 
            "        (select dt_ocorrencia w_let_fim " . $crlf . 
            "           from sbpi.Calendario_Cliente a inner join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and sbpi.year(a.dt_ocorrencia) = " .  $wAno .  " and b.sigla = 'TA') " . $crlf . 
            "          where sq_cliente = " . $CL . $crlf . 
            "            and a.sq_particular_calendario = " . $_REQUEST["w_calendario"] . $crlf . 
            "        ) b, " . $crlf . 
            "        (select dt_ocorrencia w_let2_ini " . $crlf . 
            "           from sbpi.Calendario_Cliente a inner join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and sbpi.year(a.dt_ocorrencia) = " .  $wAno .  " and b.sigla = 'I2') " . $crlf . 
            "          where sq_cliente = " . $CL . $crlf . 
            "            and a.sq_particular_calendario = " . $_REQUEST["w_calendario"] . $crlf . 
            "        ) c, " . $crlf . 
            "        (select dt_ocorrencia w_let1_fim " . $crlf . 
            "           from sbpi.Calendario_Cliente a inner join sbpi.Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and sbpi.year(a.dt_ocorrencia) = " .  $wAno .  " and b.sigla = 'T1') " . $crlf . 
            "          where sq_cliente = " . $CL . $crlf . 
            "            and a.sq_particular_calendario = " . $_REQUEST["w_calendario"] . $crlf . 
            "        ) d " . $crlf;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    
    
    foreach($RS as $row) {
      $w_fim1 = f($row,"w_let1_fim");
      $w_ini2 = f($row,"w_let2_ini");
    }
          
    for ($w_cont=1; $w_cont<=2; $w_cont++){
      If ($w_cont == 1){
        $w_ini = 1;
        $w_fim = 7;
      }else{
        $w_ini = 7;
        $w_fim = 12;
        ShowHTML('   <TR valign="top" ALIGN="CENTER"><TD COLSPAN="4"><br/><br/></td></tr>');
      }
      $w_dias = 0;

      ShowHTML('   <TR valign="top" ALIGN="CENTER"><TD COLSPAN="4"><b>' . $w_cont . 'º Semestre</b></td></tr>');
      ShowHTML('   <TR valign="top">');
      ShowHTML('     <TD><b>MÊS');
      ShowHTML('     <TD><b>D.L.');
      ShowHTML('   </TR>');
      ShowHTML('   <TR>');
      ShowHTML('     <TD COLSPAN="2" HEIGHT="1" BGCOLOR="#DAEABD">');
      ShowHTML('   </TR>');
      for ($wCont = $w_ini; $wCont <= $w_fim; $wCont++){
        if(intVal(month($w_ini2)) == intVal($wCont) and intVal($w_cont) == 2){
          $w_inicial = formataDataEdicao($w_ini2);
        }else{
          $w_inicial = "01/" . SubStr(100+$wCont,1,2) . '/' . year(time());
        }

        if(intVal(month($w_fim1)) == intVal($wCont) and intVal($w_cont) == 1){
          $w_final = formataDataEdicao($w_fim1);
        }else{
          $SQL = "SELECT last_Day(to_date('01/" . SubStr(100+$wCont,1,2) . "/'||sbpi.year(sysdate),'dd/mm/yyyy')) fim from dual" . $crlf; 
          $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
          foreach($RS as $row) { $RS = $row; break; }
          $w_final = formataDataEdicao(f($RS, "fim"));
        }         
        $SQL = "SELECT sbpi.diasLetivos('" . $w_inicial . "', '" . $w_final . "'," . $CL . ",". $_REQUEST["w_calendario"] .") qtd from dual" . $crlf; 
        $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
        foreach($RS as $row) { $RS = $row; break; }      
        If(IntVal(f($RS, "qtd")) > 0) {
          ShowHTML('   <TR>');
          ShowHTML('     <TD>' . nomeMes($wCont));
          ShowHTML('     <TD ALIGN="CENTER">' . intVal(f($row, "qtd")));
          $w_mes = IntVal(f($RS,"qtd"));
          ShowHTML('   </TR>');
          $w_dias = $w_dias + $w_mes;
        }
      }
      ShowHTML('   <TR>');
      ShowHTML('     <TD nowrap>Dias Letivos');
      ShowHTML('     <TD ALIGN="CENTER">' . $w_dias);
      ShowHTML('   </TR>');        
    }
    ShowHTML('     </TABLE></td>');
    ShowHTML('     </TABLE>'); 
  
  }
}

// =========================================================================
// Tela de Arquivos
// -------------------------------------------------------------------------
function arquivo() {
  extract($GLOBALS);
  
  $wAno      = intVal($_REQUEST['wAno']);  
  
  If ($wAno == ""){
     $wAno = Date("Y");
  }  

  $SQL = "SELECT case in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end in_destinatario, " . $crlf . 
         "       dt_arquivo, ds_titulo, nr_ordem, ds_arquivo, ln_arquivo, 'Escola' Origem, x.ln_internet diretorio " . $crlf . 
         "from sbpi.Cliente_Arquivo a, sbpi.cliente x " . $crlf . 
         "WHERE a.ativo = 'S'" . $crlf . 
         "  AND x.sq_cliente = " . $CL . " " . $crlf . 
         "  AND a.sq_cliente = " . $CL . " " . $crlf . 
         "  and in_destinatario <> 'E'" . $crlf . 
         "  and sbpi.year(a.dt_arquivo) = " . $wAno . $crlf . 
         "UNION " . $crlf . 
         "SELECT case a.in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end in_destinatario, " . $crlf . 
         "       a.dt_arquivo, a.ds_titulo, a.nr_ordem, ds_arquivo, ln_arquivo, 'SEDF' Origem, d.ds_diretorio diretorio " . $crlf . 
         "from sbpi.Cliente_Arquivo a INNER join sbpi.Cliente b ON (a.sq_cliente = b.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente c ON (c.sq_cliente_pai  = b.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente e ON (e.sq_cliente_pai  = c.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente_Site d ON (b.sq_cliente = d.sq_cliente) " . $crlf . 
         "WHERE a.ativo = 'S'" . $crlf . 
         "  and e.sq_cliente = " . $CL . " " . $crlf .
         "  and in_destinatario <> 'E'" . $crlf . 
         "  and sbpi.year(a.dt_arquivo) = " . $wAno . $crlf . 
         "UNION " . $crlf . 
         "SELECT case a.in_destinatario when 'A' then 'Aluno' when 'P' then 'Professor' when 'E' then 'Escola' else 'Todos' end in_destinatario, " . $crlf . 
         "       a.dt_arquivo, a.ds_titulo, a.nr_ordem, ds_arquivo, ln_arquivo, 'Regional' Origem, d.ds_diretorio diretorio " . $crlf . 
         "from sbpi.Cliente_Arquivo a INNER join sbpi.Cliente c ON (a.sq_cliente = c.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente e ON (e.sq_cliente_pai  = c.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente_Site d ON (c.sq_cliente = d.sq_cliente) " . $crlf . 
         "WHERE a.ativo = 'S'" . $crlf . 
         "  and e.sq_cliente = " . $CL . " " . $crlf .
         "  and in_destinatario <> 'E'" . $crlf . 
         "  and sbpi.year(a.dt_arquivo) = " . $wAno . $crlf . 
         "ORDER BY dt_arquivo desc, origem, nr_ordem, in_destinatario " . $crlf; 
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);	 

  $SQL = "SELECT idescola from sbpi.Cliente_Particular where sq_cliente = " . $CL;
  $RS1 = db_exec::getInstanceOf($dbms, $SQL, &$numRows);	 
  $idEscola = $RS1[0]["idescola"]  ;

  $url = "http://www.prodatadf.com.br/recadastramento/uploads/esc_id".$idEscola;

  ShowHTML ('<tr><td><TABLE border=0 cellSpacing=5 width="95%"> ');
 
  ShowHTML ('  <TR> ');
  ShowHTML ('    <TD HEIGHT="1" BGCOLOR="#DAEABD"> ');
  ShowHTML ('  </TR>');


  if (count($RS)>0) {
	    ShowHTML('<ul>');
    foreach($RS as $row) {
           ShowHTML ('    <li><a href="' . $url. '/' . f($row, "ln_arquivo") . '" target="_blank">' . f($row, "ds_titulo") . '</a><br><div align="justify"><font size=1>.:. ' . f($row, "ds_arquivo") . '</font></div></li>');
           $i++;
    } 
	  ShowHTML('</ul>');
  }else{
    ShowHTML ('  <TR valign="top"> ');
    ShowHTML ('    <TD colspan="5"><BR>Não há arquivos disponíveis no momento para o ano de ' . $wAno . '</TD> ');
    ShowHTML ('  </TR> ');	
  }
  	ShowHTML ('  </TABLE> ');	

  $SQL = "SELECT sbpi.year(dt_arquivo) ano " . $crlf . 
         "from sbpi.Cliente_Arquivo a, sbpi.cliente x " . $crlf . 
         "WHERE a.ativo = 'S'" . $crlf . 
         "  AND x.sq_cliente = " . $CL . " " . $crlf . 
         "  AND a.sq_cliente = " . $CL . " " . $crlf . 
         "  and in_destinatario <> 'E'" . $crlf . 
         "  and sbpi.year(a.dt_arquivo) <> " . $wAno . $crlf . 
         "UNION " . $crlf . 
         "SELECT sbpi.year(dt_arquivo) ano " . $crlf . 
         "from sbpi.Cliente_Arquivo a INNER join sbpi.Cliente b ON (a.sq_cliente = b.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente c ON (c.sq_cliente_pai  = b.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente e ON (e.sq_cliente_pai  = c.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente_Site d ON (b.sq_cliente = d.sq_cliente) " . $crlf . 
         "WHERE a.ativo = 'S'" . $crlf . 
         "  AND e.sq_cliente = " . $CL . " " . $crlf . 
         "  and in_destinatario <> 'E'" . $crlf . 
         "  and sbpi.year(a.dt_arquivo) <> " . $wAno . $crlf . 
         "UNION " . $crlf . 
         "SELECT sbpi.year(dt_arquivo) ano " . $crlf . 
         "from sbpi.Cliente_Arquivo a INNER join sbpi.Cliente c ON (a.sq_cliente = c.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente e ON (e.sq_cliente_pai  = c.sq_cliente) " . $crlf . 
         "                             INNER join sbpi.Cliente_Site d ON (c.sq_cliente = d.sq_cliente) " . $crlf . 
         "WHERE a.ativo = 'S'" . $crlf . 
         "  AND e.sq_cliente = " . $CL . " " . $crlf . 
         "  and in_destinatario <> 'E'" . $crlf . 
         "  and sbpi.year(a.dt_arquivo) <> " . $wAno . $crlf . 
         "ORDER BY 1 " . $crlf;
		 
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);	

	if (count($RS)>0) {
		ShowHTML ('  <TR valign="top"><TD colspan="5"><br><b>Arquivos de outros anos</b><br> ');
		ShowHTML ('  <ul>');  
		foreach($RS as $row) {
			ShowHTML ('     <li><a href="' . $w_dir . $w_pagina . $par . '&CL=' . str_replace("sq_cliente=", $CL,"sq_cliente=") . '&wAno=' . f($row, "ano") . '" >Arquivos de ' . f($row, "ano") . '</a></li>');
		}
		ShowHTML ('  </ul>');
		ShowHTML ('  </TR>');
	}
  ShowHTML ('    </TABLE></CENTER>'); 
}

// =========================================================================
// Tela de Alunos
// -------------------------------------------------------------------------
function aluno() {
  extract($GLOBALS);
  $strTexto = 'Informe nos campos abaixo sua Matrícula e Senha de Acesso, conforme informado pela escola. Se você não recebeu esses dados, clique no botão <i>Fale Conosco</I>, acima, para ver como entrar em contato com a escola e consegui-los.';
  $strAction = $w_dir.$w_pagina.'Valida&CL='.$CL.'&O=0';
  $strCampo = 'Matrícula';
  
  ShowHTML('<script Language="JavaScript">');
  ShowHTML('<!--');
  ShowHTML('function Validacao(theForm)');
  ShowHTML('{');
  ShowHTML('  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-";');
  ShowHTML('  var checkStr = theForm.UID.value;');
  ShowHTML('  var allValid = true;');
  ShowHTML('  for (i = 0;  i < checkStr.length;  i++)');
  ShowHTML('  {');
  ShowHTML('    ch = checkStr.charAt(i);');
  ShowHTML('    for (j = 0;  j < checkOK.length;  j++)');
  ShowHTML('      if (ch == checkOK.charAt(j))');
  ShowHTML('        break;');
  ShowHTML('    if (j == checkOK.length)');
  ShowHTML('    {');
  ShowHTML('      allValid = false;');
  ShowHTML('      break;');
  ShowHTML('    }');
  ShowHTML('  }');
  ShowHTML('  if (!allValid)');
  ShowHTML('  {');
  ShowHTML('    alert("Informe apenas letras e números no campo \"Nome de Usuário\".");');
  ShowHTML('    theForm.UID.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');
  ShowHTML('  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-";');
  ShowHTML('  var checkStr = theForm.PWD.value;');
  ShowHTML('  var allValid = true;');
  ShowHTML('  for (i = 0;  i < checkStr.length;  i++)');
  ShowHTML('  {');
  ShowHTML('    ch = checkStr.charAt(i);');
  ShowHTML('    for (j = 0;  j < checkOK.length;  j++)');
  ShowHTML('      if (ch == checkOK.charAt(j))');
  ShowHTML('        break;');
  ShowHTML('    if (j == checkOK.length)');
  ShowHTML('    {');
  ShowHTML('      allValid = false;');
  ShowHTML('      break;');
  ShowHTML('    }');
  ShowHTML('  }');
  ShowHTML('  if (!allValid)');
  ShowHTML('  {');
  ShowHTML('    alert("Informe apenas letras e números no campo \"Senha\".");');
  ShowHTML('    theForm.PWD.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');
  ShowHTML('  return (true);');
  ShowHTML('}');
  ShowHTML('//-->');
  ShowHTML('</script>');
  ShowHTML('<form method="POST" action="'.$strAction.'" onsubmit="return Validacao(this)" name="Form1">');
  
  ShowHTML('<tr><td><table align="center" width="95%">');
  ShowHTML('<tr><td colspan="3"><p align="justify">'.$strTexto.'<p>&nbsp');
  ShowHTML('<tr><td align="right"><b>'.$strCampo.':</b><td>&nbsp;<td><input type="text" class="texto" lenght="14" maxsize="14" name="UID" value="">');
  ShowHTML('<tr><td align="right"><b>Senha de acesso:</b><td>&nbsp;<td><input type="password" class="texto" lenght="14" maxsize="14" name="PWD" value="">');
  ShowHTML('<tr><td align="right">&nbsp;<td>&nbsp;<td><input type="submit" name="BTN" class="botao" value="Encontrar">&nbsp;<input type="reset" class="botao" name="CLR" value="Limpar">');
  ShowHTML('</table>');
  
  ShowHTML('</form>');
}

// =========================================================================
// Tela de validação do aluno
// -------------------------------------------------------------------------
function valida() {
  extract($GLOBALS);
  $w_uid = str_replace('"', '',str_replace("'", "",Trim(strtoUpper($_REQUEST["UID"]))));
  $w_pwd = str_replace('"', '',str_replace("'", "",Trim(strtoUpper($_REQUEST["PWD"]))));

  if ($O==0) {
     $SQL = "SELECT * from sbpi.Aluno ".$crlf.
            "WHERE sq_cliente = ".$CL.$crlf.
            "  AND NR_MATRICULA   ='".$w_uid."'".$crlf.
            "  AND DS_SENHA_ACESSO='".$w_pwd."'".$crlf;
  } else {
     $SQL = "SELECT * from sbpi.Cliente ".$crlf.
            "WHERE sq_cliente     = ".$CL.$crlf.
            "  AND DS_USERNAME    ='".$w_uid."'".$crlf.
            "  AND DS_SENHA_ACESSO='".$w_pwd."'".$crlf;
  }
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);

  if (count($RS)==0) {
    ShowHTML('<tr><td align="right"><b><font size=2>Validação</font></b></td></tr>');

    ShowHTML('<tr><td><p align="justify">Nome de usuário ou senha de acesso inválida. Volte à página anterior para tentar novamente.</p>');
    ShowHTML('<p><center><img src="img/bt_voltar.gif" border=0 valign="center" onClick="history.go(-1)" alt="Volta">');
    ShowHTML('</tr>');

  } else {
    foreach($RS as $row) { $RS = $row; break; }
    if ($O==0) {
       // Grava o acesso na tabela de log
       $w_chave = f($RS,"sq_aluno");
       $SQL= "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql) ".$crlf.
             "( select sbpi.sq_cliente_log.nextval, ".$crlf.
             "         ".$CL.", ".$crlf.
             "         sysdate, ".$crlf.
             "         '".$_SERVER['REMOTE_ADDR']."', ".$crlf.
             "         0, ".$crlf.
             "         'Acesso à tela do aluno ".f($RS,'no_aluno')." (".f($RS,'nr_matricula').").', ".$crlf.
             "         null ".$crlf.
             "   from dual) ".$crlf;
       $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
       ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT">');
       ShowHTML('   window.open("'.montaURL_JS(null, $conRootSiw . $w_dir . 'aluno.php?par=inicial&AL='.$w_chave.'&CL='.$CL.'&O=10').'", "aluno", "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=780,height=500,left=10,top=10");');
       ShowHTML('   history.go(-1);');
       ShowHTML('</SCRIPT>');
    } else {
       // Grava o acesso na tabela de log
       $SQL= "insert into sbpi.Cliente_Log (sq_cliente_log, sq_cliente, data, ip_origem, tipo, abrangencia, sql) ".$crlf.
             "( select sbpi.sq_cliente_log.nextval, ".$crlf.
             "         ".$CL.", ".$crlf.
             "         sysdate, ".$crlf.
             "         '".$_SERVER['REMOTE_ADDR']."', ".$crlf.
             "         0, ".$crlf.
             "         'Acesso à tela de atualização da escola.', ".$crlf.
             "         null ".$crlf.
             "   from dual) ".$crlf;
       $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
       
       ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT">');
       ShowHTML('   window.open("manut.php?par=inicial?CL='.$CL.'&O='.$O.'", "cliente", "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=780,height=580,left=10,top=10");');
       ShowHTML('   history.go(-1);');
       ShowHTML('</SCRIPT>');
    }
  }
  ShowHTML('</tr></center>');
}

// =========================================================================
// Tela de Oferta
// -------------------------------------------------------------------------
function oferta() {
  extract($GLOBALS);

  $wAno =  $_REQUEST["wAno"];
  
  If ($wAno == ''){
     $wAno = Year(Time());
  }

     $SQL = "SELECT INFANTIL, FUNDAMENTAL, EJA, MEDIO, DISTANCIA, PROFISSIONAL " . $crlf .
           "FROM sbpi.cliente_particular a " . $crlf .
           "WHERE a.sq_cliente = " . $CL;
     
     $RS1 = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
     foreach($RS1 as $row) { $RS = $row; break; }
     $wCont = 0;
     
     
     If(Count($RS) > 0){
     
    // case 'INICIAL':   ShowHTML('      <h3>Inicial</h3>');       Inicial();    break;
        switch ( intval(f($RS,"INFANTIL") ) ) {
            
            Case 0 : $modalidadeinf = "Não oferece";         break;
            Case 1 : $modalidadeinf = "Creche";              break;
            Case 2 : $modalidadeinf = "Pré-Escola";          break;
            Case 4 : $modalidadeinf = "Creche e Pré-Escola"; break;
        }
        switch  (intval(f($RS,"FUNDAMENTAL"))){
            Case 0:  $modalidadefund = "Não oferece";                          break;
            Case 1:  $modalidadefund = "1º ao 9º ano";                         break;
            Case 2:  $modalidadefund = "1º ao 9º ano e 1ª a 4ª	";             break;
            Case 3:  $modalidadefund = "1º ao 9º ano e 1ª a 8ª"   ;            break;
            Case 4:  $modalidadefund = "1º ao 5º ano e 1ª a 4ª"    ;           break;
            Case 5:  $modalidadefund = "1º ao 5º ano e 1ª a 8ª"     ;          break;
            Case 6:  $modalidadefund = "1º ao 5º ano"            ;             break;
            Case 7:  $modalidadefund = "3ª e 4ª c/ext. prog. e 1º ao 9º ano";  break;          
            Case 8:  $modalidadefund = "3ª e 4ª c/ext. prog. e 1º ao 5º ano" ; break;          
            Case 9:  $modalidadefund = "1ª a 4ª Séries"            ;           break;
            Case 10: $modalidadefund = "1ª a 8ª Séries"            ;           break;
            Case 11: $modalidadefund = "5ª a 8ª Séries"            ;           break;
        }
        
        switch (intval(f($RS,"MEDIO"))) {
            Case 0: $modalidademed = "Não oferece"; break;            
            Case 1:  $modalidademed = "Oferece";    break;
        }
        
        switch (intval(f($RS,"EJA"))){
            Case 0: $modalidadeeja = "Não oferece"       ;  break;     
            Case 1: $modalidadeeja = "1º e 2º Segmento" ;   break;        
            Case 2: $modalidadeeja = "1º e 3º Segmento";    break;
            Case 3: $modalidadeeja = "2º e 3º Segmento"   ; break;             
            Case 4: $modalidadeeja = "1º,2º e 3º Segmento"; break;
            Case 5: $modalidadeeja = "2º Segmento";         break;
            Case 6: $modalidadeeja = "3º Segmento";         break;
            Case 7: $modalidadeeja = "1º Segmento";         break;
        }
        
        switch (intval(f($RS, "DISTANCIA"))){
            Case 0: $modalidadedist = "Não oferece"; break;            
            Case 1: $modalidadedist = "Oferece";     break;
       }
        switch (intval(f($RS,"PROFISSIONAL"))){
            Case 0: $modalidadeprof = "Não oferece" ; break;          
            Case 1: $modalidadeprof = "Oferece";      break;
        }
     
        
        $wCont = 1;
        foreach($RS1 as $row){

        ShowHTML ('<p align="justify">Informações relativas às ofertas da instituição de ensino:</p>');
        ShowHTML ('<table width="100%" border="0" cellspacing="5" cellpadding="2">');
        ShowHTML ('  <TR valign="top">');
        ShowHTML ('    <TD align="right" width="30%"><b>Ensino Infantil: </b></TD>');
        ShowHTML ('    <TD align "left" width="70%">' . $modalidadeinf . '</TD>');
        ShowHTML ('  </TR>');
        ShowHTML ('  <TR valign="top">');
        ShowHTML ('    <TD align="right"><b>Ensino Fundamental: </b></TD>');
        ShowHTML ('    <TD align "left">' . $modalidadefund . '</TD>');
        ShowHTML ('  </TR>');
        ShowHTML ('  <TR valign="top">');
        ShowHTML ('    <TD align="right"><b>Ensino Médio: </b></TD>');
        ShowHTML ('    <TD align "left">' . $modalidademed . '</TD>');
        ShowHTML ('  </TR>');
        ShowHTML ('  <TR valign="top">');
        ShowHTML ('    <TD align="right"><b>Ensino de Jovens e Adultos: </b></TD>');
        ShowHTML ('    <TD align "left">' . $modalidadeeja . '</TD>');
        ShowHTML ('  </TR>');
        ShowHTML ('  <TR valign="top">');
        ShowHTML ('    <TD align="right"><b>Ensino à Distância: </b></TD>');
        ShowHTML ('    <TD align "left">' . $modalidadedist . '</TD>');
        ShowHTML ('  </TR>');
        ShowHTML ('  <TR valign="top">');
        ShowHTML ('    <TD align="right"><b>Educação Profissional: </b></TD>');
        ShowHTML ('    <TD align "left">'); 
        
         If (trim($modalidadeprof) == 'Oferece'){
          $SQL = "select sq_cliente, ds_curso ". 
                "       from sbpi.Particular_Curso a " .
                " inner join sbpi.Curso            b on (a.sq_curso = b.sq_curso)" .
                " where sq_cliente = " . $CL;
         
          $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
          
		  if(count($RS)>0){
			  ShowHTML ('  <ul id=""cursos_ul"">');
			  foreach($RS as $row){
				ShowHTML ('    <li id="cursos">' . $row["ds_curso"] . "</li>");  
			  }
			  ShowHTML ("  </ul>");
			  }else{
				  showHTML("Oferece");
			  }
         }Else{
         	 showHTML($modalidadeprof);
         }       
        
        ShowHTML('</TD>');
        ShowHTML ('  </TR>');
        ShowHTML ('</TABLE>');
        $wCont++;
        
        }
        
     }Else{
        ShowHTML ('    <p><BR>Até o presente momento, a Instituição não oferece cursos profissionalizantes. </p>');
     }     
}

// =========================================================================
// Rotina principal
// -------------------------------------------------------------------------
function Main() {
  extract($GLOBALS);
  
  switch ($par) {
  case 'INICIAL':   ShowHTML('      <h3>Inicial</h3>');       Inicial();    break;
  case 'CALENDARIO':ShowHTML('      <h3>Calendário</h3>');    Calendario(); break;
  case 'ARQUIVO':   ShowHTML('      <h3>Arquivos</h3>');      Arquivo();    break;
  case 'OFERTA':    ShowHTML('      <h3>Oferta</h3>');        Oferta();     break;
  case 'VALIDA':    Valida(); break;
  default:
    Cabecalho();
    ShowHTML('<BASE HREF="'.$conRootSIW.'">');
    BodyOpen('onLoad=this.focus();');
    Estrutura_Topo_Limpo();
    Estrutura_Menu();
    Estrutura_Corpo_Abre();
    Estrutura_Texto_Abre();
    ShowHTML('<center><br><br><br><br><br><br><br><br><br><br><img src="images/icone/underc.gif" align="center"> <b>Esta opção está sendo desenvolvida.</b><br><br><br><br><br><br><br><br><br><br></center>');
    Estrutura_Texto_Fecha();
    Estrutura_Fecha();
    Estrutura_Fecha();
    Estrutura_Fecha();
    Rodape();
  }
  
}
?>