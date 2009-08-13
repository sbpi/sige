<?php
// Define o banco de dados a ser utilizado
$_SESSION['DBMS']=1;

$w_dir_volta = '../../';
$w_dir = 'modelos/mod16/';
$w_pagina = 'aluno.php?par=';


include_once($w_dir_volta.'constants.inc');
include_once($w_dir_volta.'jscript.php');
include_once($w_dir_volta.'funcoes.php');
include_once($w_dir_volta.'classes/db/abreSessao.php');
include_once($w_dir_volta.'classes/sp/db_exec.php');

// =========================================================================
//  /aluno.php
// ------------------------------------------------------------------------
// Nome     : Alexandre Vinhadelli Papadópolis
// Descricao: Página de exibição dos dados do aluno
// Mail     : alex@sbpi.com.br
// Criacao  : 22/01/2009 11:31
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
$AL           = intVal($_REQUEST['AL']);
$par          = substr(strtoupper($_REQUEST['par']),0,30);
$O            = substr(strtoupper($_REQUEST['O']),0,1);
$R            = substr(strtoupper($_REQUEST['R']),0,30);
$SG           = substr(strtoupper($_REQUEST['SG']),0,30);
$P3           = intVal(nvl($_REQUEST['P3'],1));
$P4           = intVal(nvl($_REQUEST['P4'],$conPageSize));

$w_ano        = retornaAno();

  ShowHTML('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">');
  ShowHTML('<html xmlns="http://www.w3.org/1999/xhtml">');
  ShowHTML('<head>');
  ShowHTML('   <BASE HREF="'.$conRootSIW.'">');
  ShowHTML('   <title>Secretaria de Estado de Educa&ccedil;&atilde;o</title>');
  ShowHTML('   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /> ');
  if ($par!='BOLETIMIMP') {
    ShowHTML('   <link href="css/estilo.css" media="screen" rel="stylesheet" type="text/css" />');
    ShowHTML('   <link href="css/print.css"  media="print"  rel="stylesheet" type="text/css" />');
    ShowHTML('   <script language="javascript" src="js/scripts.js"> </script>');
  }
  ShowHTML('   <script src="js/jquery.js"></script>');
  ShowHTML('</head>');
  ShowHTML('<body>');
  if ($par!='BOLETIMIMP') {
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
    ShowHTML('      <li><a '.(($par=='INICIAL') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'inicial&CL='.$CL.'&AL='.$AL.'" >Inicial</a></li>');
    ShowHTML('      <li><a '.(($par=='SENHA') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'senha&CL='.$CL.'&AL='.$AL.'" id="link1">Troca de senha</a></li>');
    ShowHTML('      <li><a '.(($par=='BOLETIM') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'boletim&CL='.$CL.'&AL='.$AL.'" >Boletim</a></li>');
    ShowHTML('      <li><a '.(($par=='GRADE') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'grade&CL='.$CL.'&AL='.$AL.'" >Grade horária</a></li>');
    ShowHTML('      <li><a '.(($par=='MENSAGEM') ? 'class="selected"' : '').' href="'.$w_dir.$w_pagina.'mensagem&CL='.$CL.'&AL='.$AL.'" id="link2">Mensagens</a></li>');
    ShowHTML('  </ul>');
    ShowHTML('</div>');
  }
  $SQL = "SELECT a.sq_cliente, a.ds_cliente, b.ds_mensagem, c.no_aluno, c.nr_matricula ".$crlf.
         "  FROM sbpi.Cliente a ".$crlf.
         "       INNER JOIN sbpi.cliente_site b on (a.sq_cliente = b.sq_cliente) ".$crlf.
         "       INNER JOIN sbpi.aluno        c on (b.sq_cliente = c.sq_cliente) ".$crlf.
  "WHERE c.sq_aluno = ".$AL;
  
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  foreach($RS as $row) { $RS = $row; break; }
  if ($par!='BOLETIMIMP') {
    ShowHTML('  <div id="conteudo"><h2>'.f($RS,'ds_cliente').'</h2>');
    ShowHTML('  <p><b>'.f($RS,'no_aluno').' ('.f($RS,'nr_matricula').')</b> »');

    switch ($par) {
      case 'INICIAL':       ShowHTML('      <b>Inicial</b></p>');         break;
      case 'SENHA':         ShowHTML('      <b>Troca de senha</b></p>');  break;
      case 'BOLETIM':       ShowHTML('      <b>Boletim</b></p>');         break;
      case 'GRADE':         ShowHTML('      <b>Grade horária</b></p>');   break;
      case 'MENSAGEM':      ShowHTML('      <b>Mensagens</b></p>');       break;
    }
    ShowHTML('  </ul>');
    ShowHTML('</div>');
  }
  ShowHTML('    <div id="texto"><!-- Conteúdo -->');
  ShowHTML('        <table width="750" border="0">');

  Main();

  ShowHTML('        </table>');
  if ($par!='BOLETIMIMP') ShowHTML('  </div>');
  ShowHTML('  <br clear="all" />');
  if ($par!='BOLETIMIMP') {
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
  }
  Rodape();
  
// Fecha conexão com o banco de dados
if (isset($_SESSION['DBMS'])) FechaSessao($dbms);

exit;

// =========================================================================
// Tela de abertura
// -------------------------------------------------------------------------
function Inicial() {
  extract($GLOBALS);
  ShowHTML('<tr><td>');

  ShowHTML('<p align="justify"><b>Instruções:</b></p>');
  ShowHTML('<ul class="texto">');
  ShowHTML('<li>Selecione a opção desejada no menu acima para ver o boletim do aluno, sua grade horária ou histórico de mensagens.');
  ShowHTML('<li>Abaixo poderão aparecer mensagens dirigidas pela escola ou pela rede de ensino. Leia-a atentamente e clique sobre <i>[Arquivar]</i> para enviá-la ao histórico de mensagens.');
  ShowHTML('<li>Suas informações cadastrais também serão exibidas. Qualquer informação incompleta ou incorreta pode ser alterada junto à secretaria de sua escola.');
  ShowHTML('<li>Se o conteúdo da página for maior que sua altura máxima, use a barra de rolagem, à direita, para subir ou descer o conteúdo.');
  ShowHTML('<li>Se você tiver uma impressora disponível e quiser imprimir o conteúdo de qualquer página, clique com o botão direito do <i>mouse</i> e, em seguida, sobre a opção <i>Imprimir</i>.');
  ShowHTML('</ul>');

  // Exibe eventuais mensagens dirigidas ao aluno
  $SQL = "SELECT a.no_aluno,b.* ".$crlf.
        "  from sbpi.Aluno a INNER JOIN sbpi.Mensagem_aluno b ON (a.sq_aluno = b.sq_aluno) ".$crlf.
        " WHERE b.in_lida = 'N' ".$crlf.
        "   AND a.sq_aluno = ".$AL." ".$crlf.
        "ORDER BY b.DT_MENSAGEM";
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);

  if (count($RS)>0) {
    ShowHTML('<br><div align="justify"><b>Novas mensagens:</b><hr></div>');
    ShowHTML('<TABLE BORDER=0 WIDTH="100%">');
    foreach ($RS as $row) {
      if (f($row,'sq_mensagem') > '') ShowHTML('<TR valign="top"><TD width="2%">::<TD width="98%"><div align=justify>'.f($row,'ds_mensagem').' ('.formataDataEdicao(f($row,'dt_mensagem')).')'.'&nbsp;&nbsp;&nbsp;<a href="'.w_dir.sstrCL.'?EW='.conWhatGravaMens.'&EF='.sstrEF.'&CL='.CL.'&EA='.sstrEA.'&IN=sq_mensagem='.f($row,'sq_mensagem').'"><b><font size="2">[Arquivar]</font></b></a></div></font></TD></TR>');
    }
    ShowHTML('</TABLE>');
  }

  // Recupera os dados cadastrais do aluno
  $SQL = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, trim(a.dt_nascimento) dt_nascimento,  ".$crlf.
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end in_sexo, ".$crlf.
        "       trim(a.ds_naturalidade), trim(a.no_mae), trim(a.nr_fone_mae), ".$crlf.
        "       trim(a.no_pai) no_pai, trim(a.nr_fone_pai) nr_fone_pai, ".$crlf.
        "       trim(a.no_resposavel) no_resposavel, trim(a.nr_fone_responsavel) nr_fone_responsavel, ".$crlf.
        "       trim(a.ds_email_responsavel) ds_email_responsavel, trim(nr_fone_1) nr_fone_1, trim(nr_fone_2) nr_fone_2 ".$crlf.
        "from sbpi.Aluno a ".$crlf.
        "WHERE a.sq_aluno = ".$AL;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  
  if (count($RS)>0) {
    foreach ($RS as $row) { $RS = $row; break; }
    ShowHTML('<b>Informações cadastrais:</b><hr>');
    ShowHTML('<TABLE align="center" WIDTH="100%" BORDER=0>');
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><B>Nome:</b><BR>'.f($RS,'NO_ALUNO').'</TD>');
    ShowHTML('    <TD><B>Matrícula:</b><BR>'.f($RS,'NR_MATRICULA').'</TD>');
    ShowHTML('  </TR>');

    // Recupera os dados cadastrais do aluno
    $SQL = "SELECT c.ano_letivo, c.st_aluno, ".$crlf.
          "       b.ds_grau, b.ds_serie, b.ds_turma, replace(b.ds_curso,'¬','ª') ds_curso, ".$crlf.
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end ds_turno ".$crlf.
          "from sbpi.Aluno_Turma      c ".$crlf.
          "     INNER join sbpi.Turma b ON (b.sq_turma = c.sq_turma and b.sq_cliente = c.sq_cliente and c.ano_letivo = ".$w_ano.") ".$crlf.
          "WHERE c.sq_aluno = ".f($RS,'SQ_ALUNO');
    $RS1 = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    ShowHTML('  <TR valign="TOP"><TD COLSPAN="3"><TABLE WIDTH="100%" BORDER="1">');
    ShowHTML('    <TR valign="TOP"><TD COLSPAN="6"><B>Turma(s):');
    ShowHTML('    <TR valign="TOP" align="CENTER">');
    ShowHTML('    <TD><B>Ano Letivo</b></TD>');
    ShowHTML('    <TD><B>Série</b></TD>');
    ShowHTML('    <TD><B>Turno</b></TD>');
    ShowHTML('    <TD><B>Turma</b></TD>');
    ShowHTML('    <TD><B>Etapa/Curso/Modalidade</b></TD>');
    ShowHTML('    <TD><B>Situação</b></TD>');
    if (count($RS1)==0) {
      ShowHTML('  <TR><TD colspan="6" ALIGN="CENTER">Turmas não encontradas.</TD></TR>');
    } else {
      foreach($RS1 as $row) {
        ShowHTML('  <TR valign="TOP">');
        //ShowHTML('    <TD><B>Grau:</b><BR>'.f($row,'DS_GRAU').'</TD>');
        ShowHTML('    <TD ALIGN="CENTER">'.f($row,'ANO_LETIVO').'</TD>');
        ShowHTML('    <TD ALIGN="CENTER">'.f($row,'DS_SERIE').'</TD>');
        ShowHTML('    <TD>'.f($row,'DS_TURNO').'</TD>');
        ShowHTML('    <TD>'.f($row,'DS_TURMA').'</TD>');
        ShowHTML('    <TD>'.nvl(f($row,'DS_CURSO'),'-'));
        ShowHTML('    <TD>'.f($row,'ST_ALUNO').'</TD>');
      }
    }
    ShowHTML('  </TR></TABLE>');
    
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><B>Data de nascimento:</b><BR>'.nvl(f($RS,'DT_NASCIMENTO'),'-'));
    ShowHTML('    <TD><B>Sexo:</b><BR>'.f($RS,'IN_SEXO').'</TD>');
    ShowHTML('    <TD><B>Naturalidade:</b><BR>'.nvl(f($RS,'DS_NATURALIDADE'),'-'));
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><B>Mãe:</b><BR>'.nvl(f($RS,'NO_MAE'),'-'));
    ShowHTML('    <TD><B>Telefone:</b><BR>'.nvl(f($RS,'NR_FONE_MAE'),'-'));
    ShowHTML('  </TR>');
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><B>Pai:</b><BR>'.nvl(f($RS,'NO_PAI'),'-'));
    ShowHTML('    <TD><B>Telefone:</b><BR>'.nvl(f($RS,'NR_FONE_PAI'),'-'));
    ShowHTML('  </TR>');
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><B>Responsável pelo aluno:</b><BR>'.nvl(f($RS,'NO_RESPOSAVEL'),'-'));
    ShowHTML('    <TD><B>Telefone:</b><BR>'.f($RS,'NR_FONE_RESPONSAVEL'),'-');
    ShowHTML('    <TD><B>e-Mail:</b><BR>'.nvl(f($RS,'DS_EMAIL_RESPONSAVEL'),'-'));
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><B>Outros telefones do aluno:</b></TD>');
    ShowHTML('    <TD><B>Telefone 1:</b><BR>'.nvl(f($RS,'NR_FONE_1'),'-'));
    ShowHTML('    <TD><B>Telefone 2:</b><BR>'.nvl(f($RS,'NR_FONE_2'),'-'));
    ShowHTML('</TABLE></CENTER>');
  }
}

// =========================================================================
// Tela de Boletim
// -------------------------------------------------------------------------
function boletim() {
  extract($GLOBALS);
  ShowHTML('<tr><td>');

  // Recupera os dados cadastrais do aluno
  $SQL = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, a.dt_nascimento,  ".$crlf.
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end in_sexo, ".$crlf.
        "       a.ds_naturalidade, a.no_mae, a.nr_fone_mae, a.no_pai, a.nr_fone_pai, ".$crlf.
        "       a.no_resposavel, a.nr_fone_responsavel, a.ds_email_responsavel, nr_fone_1, nr_fone_2 ".$crlf.
        "from sbpi.Aluno a ".$crlf.
        "WHERE a.sq_aluno = ".$AL; 
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  if (count($RS)>0) {
    foreach ($RS as $row) { $RS = $row; break; }
    ShowHTML('<TABLE align="center" WIDTH="100%" BORDER=0>');
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD COLSPAN="2" ALIGN="RIGHT"><FONT FACE=VERDANA SIZE=1><A HREF="#" onClick="window.open(\''.montaUrl_JS($w_dir,$w_pagina.'BoletimImp&CL='.$CL.'&AL='.$AL).'\',\'Boletim\',\'width=600, height=350, top=50, left=50, toolbar=no, scrollbars=yes, resizable=yes, status=no\'); return false;" title="Clique para visualizar a versão de impressão do boletim!"><B>Versão p/ Impressão</B><IMG ALIGN="CENTER" BORDER=0   TITLE="Imprimir" SRC="img/bt_imprimir.gif"></A></TD>');
    ShowHTML('  </TR>');
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Nome:</b> '.f($RS,'NO_ALUNO').'</TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Matrícula:</b> '.f($RS,'NR_MATRICULA').'</TD>');
    ShowHTML('  </TR>');
    ShowHTML('  <TR valign="TOP"><TD COLSPAN="3"><TABLE WIDTH="100%" BORDER="1">');
    ShowHTML('    <TR valign="TOP" align="CENTER">');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Ano Letivo</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Série</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Turno</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Turma</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Etapa/Curso/Modalidade</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Situação</b></TD>');
    $SQL = "SELECT c.ano_letivo, c.st_aluno, ".$crlf.
          "       b.ds_grau, b.ds_serie, b.ds_turma, trim(replace(b.ds_curso,'¬','ª')) ds_curso, ".$crlf.
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end ds_turno ".$crlf.
          "from sbpi.Aluno_Turma      c ".$crlf.
          "     INNER join sbpi.Turma b ON (b.sq_turma = c.sq_turma and b.sq_cliente = c.sq_cliente and c.ano_letivo = ".$w_ano.") ".$crlf.
          "WHERE c.sq_aluno = ".f($RS,"SQ_ALUNO"); 
    $RS1 = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    
    if (count($RS1)==0) {
      ShowHTML('  <TR><TD colspan="6" ALIGN="CENTER">Turmas não encontradas.</TD></TR>');
    } else {
      foreach($RS1 as $row) {
        ShowHTML('  <TR valign="TOP">');
        ShowHTML('    <TD ALIGN="CENTER"><FONT FACE=VERDANA SIZE=1>'.f($row,'ANO_LETIVO').'</TD>');
        ShowHTML('    <TD ALIGN="CENTER"><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_SERIE').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_TURNO').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_TURMA').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>');
        ShowHTML(trocanulo(f($row,'DS_CURSO')));
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'ST_ALUNO').'</TD>');
      }
      ShowHTML('     </TR>');
    }
    ShowHTML('  </TABLE>');
    ShowHTML('</TABLE><br><br>');
  }

  ShowHTML('<tr><td><TABLE border=1 cellpadding=5 cellspacing=0 width="100%" BGCOLOR="#F7F7F7">');
  ShowHTML('  <TR ALIGN="CENTER">');
  ShowHTML('    <TD rowspan=2 nowrap>Comp.Curric.');
  ShowHTML('    <TD colspan=2 nowrap>1º Bim');
  ShowHTML('    <TD colspan=2 nowrap>2º Bim');
  ShowHTML('    <TD colspan=2 nowrap>3º Bim');
  ShowHTML('    <TD colspan=2 nowrap>4º Bim');
  ShowHTML('    <TD rowspan=2>Média Anual');
  ShowHTML('    <TD rowspan=2>Recup. Final');
  ShowHTML('    <TD rowspan=2>Faltas');
  ShowHTML('    <TD rowspan=2>Média Final');
  ShowHTML('    <TD rowspan=2>Result.');
  ShowHTML('  </TR>');
  ShowHTML('  <TR>');
  ShowHTML('    <TD>Nt');
  ShowHTML('    <TD>Flt');
  ShowHTML('    <TD>Nt');
  ShowHTML('    <TD>Flt');
  ShowHTML('    <TD>Nt');
  ShowHTML('    <TD>Flt');
  ShowHTML('    <TD>Nt');
  ShowHTML('    <TD>Flt');
  ShowHTML('  <TR>');
  ShowHTML('    <TD COLSPAN="14" HEIGHT="1" BGCOLOR="##DAEABD">');
  ShowHTML('  </TR>');

  $SQL = "SELECT a.*, b.sg_disciplina, b.ds_disciplina, c.ds_mensagem_boletim ".$crlf.
        "from sbpi.Boletim a ".$crlf.
        "     INNER join sbpi.Disciplina b on (a.sq_disciplina = b.sq_disciplina) ".$crlf.
        "     INNER join sbpi.Aluno c on (a.sq_aluno = c.sq_aluno) ".$crlf.
        "WHERE a.sq_aluno = ".$AL.$crlf.
        //"  AND a.ano_letivo = ".$w_ano." ".$crlf.
        "ORDER BY b.sg_disciplina".$crlf;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);

  if (count($RS)>0) {
    foreach($RS as $row) {
      $w_ds_mensagem_boletim = f($RS,'DS_MENSAGEM_BOLETIM');
      ShowHTML('    <TR align="center">');
      ShowHTML('    <TD title="'.f($row,'ds_disciplina').'"><center><b>'.f($row,'sg_disciplina'));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b1_nota')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b1_falta')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b2_nota')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b2_falta')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b3_nota')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b3_falta')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b4_nota')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b4_falta')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'ds_media_anual')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'ds_recup_final')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'ds_falta_anual')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'ds_media_final')));
      ShowHTML('    <TD><b>' );
      if (strpos(strtolower(trocaNulo(f($row,'ds_resultado'))),'rep')!==false) {
        ShowHTML('<font color="#FF0000">'.trocaNulo(f($row,'ds_resultado')));
      } else {
        ShowHTML(trocaNulo(f($row,'ds_resultado')));
      }
    }
    if (nvl($w_ds_mensagem_boletim,'')!='') {
       ShowHTML('    <TR><TD colspan=14 align="left"><b>Mensagem:<BR><font color="#FF0000"><b>'.$w_ds_mensagem_boletim);
    }
    ShowHTML('<tr><td colspan="14"><TABLE border=0 cellpadding=1>');
    ShowHTML('  <TR><TD colspan="2"><FONT FACE=VERDANA SIZE=1><B>Legenda das disciplinas:</B>');
    $RS.reset;
    foreach($RS as $row) {
      ShowHTML('  <TR><TD><FONT FACE=VERDANA SIZE=1>'.f($row,'sg_disciplina').':');
      ShowHTML('      <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'ds_disciplina'));
    }
    ShowHTML('    </TABLE>');
  } else {
     ShowHTML('    <TR><TD colspan=14><b>Não há notas informadas.</b>');
  }
  ShowHTML('    </TABLE>');
}

// =========================================================================
// Tela de impressão do Boletim
// -------------------------------------------------------------------------
function boletimImp() {
  extract($GLOBALS);
  // Recupera o nome da escola
  $sql = "SELECT b.ds_cliente ".
        " from sbpi.Aluno a ".
        "      LEFT OUTER join sbpi.Cliente b on (a.sq_cliente = b.sq_cliente) ".
        " WHERE a.sq_aluno = ".$AL; 
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  foreach($RS as $row) { $RS = $row; break; }
  ShowHTML('<TABLE align="center" WIDTH="100%" BORDER=0>');
  ShowHTML('  <TR valign="TOP">');
  ShowHTML('    <TD align="right"><IMG ALIGN="CENTER" TITLE="Imprimir" SRC="img/bt_imprimir.gif" onClick="window.print();window.close();"></TD>');
  ShowHTML('  </TR>');
  ShowHTML('  <TR valign="TOP">');
  ShowHTML('    <TD ALIGN="center"><FONT FACE=VERDANA SIZE=2><B>'.f($RS,'ds_cliente').'</TD>');
  ShowHTML('   </TR>');
  ShowHTML('</TABLE>');
  
  ShowHTML('<tr><td>');

  // Recupera os dados cadastrais do aluno
  $SQL = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, a.dt_nascimento,  ".$crlf.
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end in_sexo, ".$crlf.
        "       a.ds_naturalidade, a.no_mae, a.nr_fone_mae, a.no_pai, a.nr_fone_pai, ".$crlf.
        "       a.no_resposavel, a.nr_fone_responsavel, a.ds_email_responsavel, nr_fone_1, nr_fone_2 ".$crlf.
        "from sbpi.Aluno a ".$crlf.
        "WHERE a.sq_aluno = ".$AL; 
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  if (count($RS)>0) {
    foreach ($RS as $row) { $RS = $row; break; }
    ShowHTML('<TABLE align="center" WIDTH="100%" BORDER=0>');
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Nome:</b> '.f($RS,'NO_ALUNO').'</TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Matrícula:</b> '.f($RS,'NR_MATRICULA').'</TD>');
    ShowHTML('  </TR>');
    ShowHTML('  <TR valign="TOP"><TD COLSPAN="3"><TABLE WIDTH="100%" BORDER="1">');
    ShowHTML('    <TR valign="TOP" align="CENTER">');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Ano Letivo</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Série</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Turno</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Turma</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Etapa/Curso/Modalidade</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Situação</b></TD>');
    $SQL = "SELECT c.ano_letivo, c.st_aluno, ".$crlf.
          "       b.ds_grau, b.ds_serie, b.ds_turma, trim(replace(b.ds_curso,'¬','ª')) ds_curso, ".$crlf.
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end ds_turno ".$crlf.
          "from sbpi.Aluno_Turma      c ".$crlf.
          "     INNER join sbpi.Turma b ON (b.sq_turma = c.sq_turma and b.sq_cliente = c.sq_cliente and c.ano_letivo = ".$w_ano.") ".$crlf.
          "WHERE c.sq_aluno = ".f($RS,"SQ_ALUNO"); 
    $RS1 = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    
    if (count($RS1)==0) {
      ShowHTML('  <TR><TD colspan="6" ALIGN="CENTER">Turmas não encontradas.</TD></TR>');
    } else {
      foreach($RS1 as $row) {
        ShowHTML('  <TR valign="TOP">');
        ShowHTML('    <TD ALIGN="CENTER"><FONT FACE=VERDANA SIZE=1>'.f($row,'ANO_LETIVO').'</TD>');
        ShowHTML('    <TD ALIGN="CENTER"><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_SERIE').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_TURNO').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_TURMA').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>');
        ShowHTML(trocanulo(f($row,'DS_CURSO')));
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'ST_ALUNO').'</TD>');
      }
    }
    ShowHTML('     </TR>');
    ShowHTML('  </TABLE>');
    ShowHTML('</TABLE>');
  }

  ShowHTML('<tr><td><TABLE border=1 cellpadding=5 cellspacing=0 width="100%" BGCOLOR="#F7F7F7">');
  ShowHTML('  <TR ALIGN="CENTER">');
  ShowHTML('    <TD rowspan=2 nowrap>Comp.Curric.');
  ShowHTML('    <TD colspan=2 nowrap>1º Bim');
  ShowHTML('    <TD colspan=2 nowrap>2º Bim');
  ShowHTML('    <TD colspan=2 nowrap>3º Bim');
  ShowHTML('    <TD colspan=2 nowrap>4º Bim');
  ShowHTML('    <TD rowspan=2>Média Anual');
  ShowHTML('    <TD rowspan=2>Recup. Final');
  ShowHTML('    <TD rowspan=2>Faltas');
  ShowHTML('    <TD rowspan=2>Média Final');
  ShowHTML('    <TD rowspan=2>Result.');
  ShowHTML('  </TR>');
  ShowHTML('  <TR>');
  ShowHTML('    <TD>Nt');
  ShowHTML('    <TD>Flt');
  ShowHTML('    <TD>Nt');
  ShowHTML('    <TD>Flt');
  ShowHTML('    <TD>Nt');
  ShowHTML('    <TD>Flt');
  ShowHTML('    <TD>Nt');
  ShowHTML('    <TD>Flt');
  ShowHTML('  <TR>');
  ShowHTML('    <TD COLSPAN="14" HEIGHT="1" BGCOLOR="##DAEABD">');
  ShowHTML('  </TR>');

  $SQL = "SELECT a.*, b.sg_disciplina, b.ds_disciplina, c.ds_mensagem_boletim ".$crlf.
        "  from sbpi.Boletim a ".$crlf.
        "       INNER join sbpi.Disciplina b on (a.sq_disciplina = b.sq_disciplina) ".$crlf.
        "       INNER join sbpi.Aluno c on (a.sq_aluno = c.sq_aluno) ".$crlf.
        " WHERE a.sq_aluno = ".$AL.$crlf.
        //"  AND a.ano_letivo = ".$w_ano." ".$crlf.
        " ORDER BY b.sg_disciplina".$crlf;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);

  if (count($RS)>0) {
    foreach($RS as $row) {
      $w_ds_mensagem_boletim = f($RS,'DS_MENSAGEM_BOLETIM');
      ShowHTML('    <TR align="center">');
      ShowHTML('    <TD title="'.f($row,'ds_disciplina').'"><center><b>'.f($row,'sg_disciplina'));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b1_nota')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b1_falta')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b2_nota')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b2_falta')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b3_nota')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b3_falta')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b4_nota')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'b4_falta')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'ds_media_anual')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'ds_recup_final')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'ds_falta_anual')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'ds_media_final')));
      ShowHTML('    <TD><b>' );
      if (strpos(strtolower(trocaNulo(f($row,'ds_resultado'))),'rep')!==false) {
        ShowHTML('<font color="#FF0000">'.trocaNulo(f($row,'ds_resultado')));
      } else {
        ShowHTML(trocaNulo(f($row,'ds_resultado')));
      }
    }
    if (nvl($w_ds_mensagem_boletim,'')!='') {
       ShowHTML('    <TR><TD colspan=14 align="left"><b>Mensagem:<BR><font color="#FF0000"><b>'.$w_ds_mensagem_boletim);
    }
    ShowHTML('<tr><td colspan="14"><TABLE border=0 cellpadding=1>');
    ShowHTML('  <TR><TD colspan="2"><FONT FACE=VERDANA SIZE=1><B>Legenda das disciplinas:</B>');
    $RS.reset;
    foreach($RS as $row) {
      ShowHTML('  <TR><TD><FONT FACE=VERDANA SIZE=1>'.f($row,'sg_disciplina').':');
      ShowHTML('      <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'ds_disciplina'));
    }
    ShowHTML('    </TABLE>');
  } else {
     ShowHTML('    <TR><TD colspan=14><b>Não há notas informadas.</b>');
  }
  ShowHTML('    </TABLE>');
}

// =========================================================================
// Tela de troca da senha
// -------------------------------------------------------------------------
function trocaSenha() {
  extract($GLOBALS);
  ShowHTML('<tr><td><p align="justify"><b>Instruções:</b></p>');
  ShowHTML('<ul>');
  ShowHTML('<li><div align="justify">Informe abaixo sua senha atual e a nova senha.</div>');
  ShowHTML('<li><div align="justify">A nova senha deve ser digitada nos campos "Informe a nova senha" e "Redigite a nova senha", para evitar erros de digitação. Deve ter pelo menos 5 letras ou números, não sendo aceitos espaços em branco, ponto, traço etc.</div>');
  ShowHTML('<li><div align="justify">Se ocorrer algum erro, tal como senha atual inválida ou valores diferentes nos campos relativos à nova senha, você receberá uma mensagem de alerta.</div>');
  ShowHTML('<li><div align="justify">Se os dados informados estiverem corretos, a senha atual será substituída pela nova senha e você receberá uma mensagem de confirmação.</div>');
  ShowHTML('<li><div align="justify">A nova senha entra em vigor assim que a troca for concluída. Assim, novos acessos à sua página pessoal devem ser feitos usando a nova senha.</div>');
  ShowHTML('</ul>');

  ShowHTML('<div align="justify"><b>Troca da senha de acesso:</b><hr></div>');

  ShowHTML('<script Language="JavaScript">');
  ShowHTML('<!--');
  ShowHTML('function Validacao(theForm)');
  ShowHTML('{');
  ShowHTML('  if (theForm.senha_atual.value == "")');
  ShowHTML('  {');
  ShowHTML('    alert("ERRO!!!\nFavor informar um valor para o campo \"Senha atual\".");');
  ShowHTML('    theForm.senha_atual.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');
  ShowHTML('  if (theForm.senha_atual.value.length < 5)');
  ShowHTML('  {');
  ShowHTML('    alert("ERRO!!!\nFavor informar pelo menos 5 letras ou números no campo \"Senha atual\".");');
  ShowHTML('    theForm.senha_atual.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');
  ShowHTML('  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-";');
  ShowHTML('  var checkStr = theForm.senha_atual.value;');
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
  ShowHTML('    alert("ERRO!!!\nInforme apenas letras e números no campo \"Senha atual\", sem espaços em branco.");');
  ShowHTML('    theForm.senha_atual.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');

  ShowHTML('  if (theForm.senha_nova.value == "")');
  ShowHTML('  {');
  ShowHTML('    alert("ERRO!!!\nFavor informar um valor para o campo \"Nova senha\".");');
  ShowHTML('    theForm.senha_nova.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');
  ShowHTML('  if (theForm.senha_nova.value.length < 5)');
  ShowHTML('  {');
  ShowHTML('    alert("ERRO!!!\nFavor informar pelo menos 5 letras ou números no campo \"Nova senha\".");');
  ShowHTML('    theForm.senha_nova.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');
  ShowHTML('  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";');
  ShowHTML('  var checkStr = theForm.senha_nova.value;');
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
  ShowHTML('    alert("ERRO!!!\nInforme apenas letras e números no campo \"Nova senha\", sem espaços em branco.");');
  ShowHTML('    theForm.senha_nova.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');

  ShowHTML('  if (theForm.senha_nova1.value == "")');
  ShowHTML('  {');
  ShowHTML('    alert("ERRO!!!\nFavor informar um valor para o campo \"Redigite a nova senha\".");');
  ShowHTML('    theForm.senha_nova1.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');
  ShowHTML('  if (theForm.senha_nova1.value.length < 5)');
  ShowHTML('  {');
  ShowHTML('    alert("ERRO!!!\nFavor informar pelo menos 5 letras ou números no campo \"Redigite a nova senha\".");');
  ShowHTML('    theForm.senha_nova1.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');
  ShowHTML('  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";');
  ShowHTML('  var checkStr = theForm.senha_nova1.value;');
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
  ShowHTML('    alert("ERRO!!!\nInforme apenas letras e números no campo \"Redigite a nova senha\", sem espaços em branco.");');
  ShowHTML('    theForm.senha_nova1.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');

  ShowHTML('  if (theForm.senha_nova.value != theForm.senha_nova1.value)');
  ShowHTML('  {');
  ShowHTML('    alert("ERRO!!!\nOs campos \"Nova senha\" e \"Redigite a nova senha\" devem ter o mesmo valor.");');
  ShowHTML('    theForm.senha_nova1.value="";');
  ShowHTML('    theForm.senha_nova.value="";');
  ShowHTML('    theForm.senha_nova.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');

  ShowHTML('  if (theForm.senha_atual.value == theForm.senha_nova.value)');
  ShowHTML('  {');
  ShowHTML('    alert("ERRO!!!\nA nova senha deve ser diferente da senha atual.");');
  ShowHTML('    theForm.senha_nova1.value="";');
  ShowHTML('    theForm.senha_nova.value="";');
  ShowHTML('    theForm.senha_nova.focus();');
  ShowHTML('    return (false);');
  ShowHTML('  }');
  ShowHTML('  return (true);');
  ShowHTML('}');
  ShowHTML('//-->');
  ShowHTML('</script>');

  ShowHTML('<FORM ACTION="'.$w_dir.$w_pagina.'senha&CL='.$CL.'&AL='.$AL.'" METHOD="POST" onSubmit="return(Validacao(this))">');
  ShowHTML('<TABLE align="center" WIDTH="100%" BORDER=0>');
  ShowHTML('  <TR valign="TOP"><TD ALIGN="RIGHT"><B>Senha atual:</b><TD><INPUT TYPE="PASSWORD" NAME="senha_atual" SIZE="14" MAXLENTH="14" VALUE=""></TD></TR>');
  ShowHTML('  <TR valign="TOP"><TD ALIGN="RIGHT"><B>Informe a nova senha:</b><TD><INPUT TYPE="PASSWORD" NAME="senha_nova" SIZE="14" MAXLENTH="14" VALUE=""></TD></TR>');
  ShowHTML('  <TR valign="TOP"><TD ALIGN="RIGHT"><B>Redigite a nova senha:</b><TD><INPUT TYPE="PASSWORD" NAME="senha_nova1" SIZE="14" MAXLENTH="14" VALUE=""></TD></TR>');
  ShowHTML('  <TR valign="TOP"><TD colspan="2" align="center"><INPUT TYPE="SUBMIT" class="botao" NAME="BOTAO" VALUE="Confirmar troca da senha de acesso"></TD></TR>');
  ShowHTML('</TABLE>');
  ShowHTML('</FORM>');

  if (nvl($_REQUEST['senha_nova'],'')!='') {
     
    ShowHTML('<SCRIPT LANGUAGE="JAVASCRIPT"><!--');
    $SQL = "select count(*) existe from sbpi.Aluno where ds_senha_acesso = '".Trim(strtoUpper($_REQUEST['senha_atual']))."' and sq_aluno = ".$AL;
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    foreach($RS as $row) { $RS = $row; break; }
    if (f($RS,'existe')==0) {
      ShowHTML('  alert("ERRO!!!\nSenha atual inválida.");');
    } else {
      $SQL = "update sbpi.Aluno set ds_senha_acesso = '".Trim(strtoUpper($_REQUEST['senha_nova']))."' where sq_aluno = ".$AL;
      $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
      ShowHTML('  alert("SUCESSO!!!\nSenha atualizada! A partir de agora, use sempre a nova senha.");');
    }
    ShowHTML('--></SCRIPT>');
  }
}

// =========================================================================
// Tela de exibição da grade horária
// -------------------------------------------------------------------------
function gradeHoraria() {
  extract($GLOBALS);

 // $Linha  = 11;
 // $Coluna = 8;

  $title = array();
  $grade = array();

  for( $i=0 ; $i < $Linha ; $i++){
      for( $j = 0;$j < $Coluna ; $j++){
          $grade[$i][$j]="-";
      }
  }

  $grade[0][0] = "Horário";
  $grade[0][1] = "Segunda";
  $grade[0][2] = "Terça";
  $grade[0][3] = "Quarta";
  $grade[0][4] = "Quinta";
  $grade[0][5] = "Sexta";
  $grade[0][6] = "Sábado";
  $grade[0][7] = "Domingo";
  $grade[1][0] = "Primeiro";
  $grade[2][0] = "Segundo";
  $grade[3][0] = "Terceiro";
  $grade[4][0] = "Quarto";
  $grade[5][0] = "Quinto";
  $grade[6][0] = "Sexto";
  $grade[7][0] = "Sétimo";
  $grade[8][0] = "Oitavo";
  $grade[9][0] = "Nono";

  

  ShowHTML('<tr><td>');

  // Recupera os dados cadastrais do aluno
  $SQL = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, a.dt_nascimento,  ".$crlf.
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end in_sexo, ".$crlf.
        "       a.ds_naturalidade, a.no_mae, a.nr_fone_mae, a.no_pai, a.nr_fone_pai, ".$crlf.
        "       a.no_resposavel, a.nr_fone_responsavel, a.ds_email_responsavel, nr_fone_1, nr_fone_2 ".$crlf.
        "from sbpi.Aluno a ".$crlf.
        "WHERE a.sq_aluno = ".$AL; 
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  if (count($RS)>0) {
    foreach ($RS as $row) { $RS = $row; break; }
    ShowHTML('<TABLE align="center" WIDTH="100%" BORDER=0>');
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Nome:</b> '.f($RS,'NO_ALUNO').'</TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Matrícula:</b> '.f($RS,'NR_MATRICULA').'</TD>');
    ShowHTML('  </TR>');
    ShowHTML('  <TR valign="TOP"><TD COLSPAN="3"><TABLE WIDTH="100%" BORDER="1">');
    ShowHTML('    <TR valign="TOP" align="CENTER">');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Ano Letivo</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Série</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Turno</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Turma</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Etapa/Curso/Modalidade</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Situação</b></TD>');
    $SQL = "SELECT c.ano_letivo, c.st_aluno, ".$crlf.
          "       b.ds_grau, b.ds_serie, b.ds_turma, trim(replace(b.ds_curso,'¬','ª')) ds_curso, ".$crlf.
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end ds_turno ".$crlf.
          "from sbpi.Aluno_Turma      c ".$crlf.
          "     INNER join sbpi.Turma b ON (b.sq_turma = c.sq_turma and b.sq_cliente = c.sq_cliente and c.ano_letivo = ".$w_ano.") ".$crlf.
          "WHERE c.sq_aluno = ".f($RS,"SQ_ALUNO"); 
    $RS1 = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    
    if (count($RS1)==0) {
      ShowHTML('  <TR><TD colspan="6" ALIGN="CENTER">Turmas não encontradas.</TD></TR>');
    } else {
      foreach($RS1 as $row) {
        ShowHTML('  <TR valign="TOP">');
        ShowHTML('    <TD ALIGN="CENTER"><FONT FACE=VERDANA SIZE=1>'.f($row,'ANO_LETIVO').'</TD>');
        ShowHTML('    <TD ALIGN="CENTER"><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_SERIE').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_TURNO').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_TURMA').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>');
        ShowHTML(trocanulo(f($row,'DS_CURSO')));
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'ST_ALUNO').'</TD>');
      }
    }
    ShowHTML('     </TR>');
    ShowHTML('  </TABLE>');
    ShowHTML('</TABLE><br><br>');
  }
	
 ShowHTML ('<tr><td align="center"><TABLE border=1 cellpadding=5 cellspacing=0 BGCOLOR="#F7F7F7" width="100%">');
 //ShowHTML( ' <TR><TD COLSPAN=14 ALIGN="LEFT"><FONT FACE=VERDANA SIZE=1>');
 $sql = "SELECT b.*, c.sg_disciplina, c.ds_disciplina " . $CrLf .
        "  from sbpi.Aluno_Turma              a " . $CrLf .
        "       INNER join sbpi.Grade_horaria b ON (a.sq_turma        = b.sq_turma and  " . $CrLf .
        "                                         a.sq_cliente = b.sq_cliente and " . $CrLf .
        //"                                         a.ano_letivo      = b.ano_letivo and " . $CrLf .
        "                                         a.ano_letivo      = " . $w_ano . $CrLf . 
        "                                        ) " . $CrLf . 
        "       INNER join sbpi.Disciplina    c ON (b.sq_disciplina = c.sq_disciplina) " . $CrLf .
        " WHERE a.sq_aluno = " . $AL . " " . $CrLf . 
        "ORDER BY b.ds_horario,b.ds_dia_semana " . $CrLf  ;
	  


 $RS = db_exec::getInstanceOf($dbms, $sql, &$numRows);


 if (count($RS)>0) {
	
    foreach ($RS as $row) {
        $grade[f($row,"DS_HORARIO")] [f($row,"DS_DIA_SEMANA")] = f($row,"SG_DISCIPLINA");
        $title[f($row,"DS_HORARIO")] [f($row,"DS_DIA_SEMANA")] = f($row,"DS_DISCIPLINA");
        If (f($row,"DS_DIA_SEMANA") > $Coluna ) $Coluna = f($row,"DS_DIA_SEMANA");
        If (f($row,"DS_HORARIO")    > $Linha  ) $Linha  = f($row,"DS_HORARIO");
	}



     ShowHTML ('  <TR valign="top" align="center">');
     For( $i = 0 ; $i <=$Coluna; $i++){
         ShowHTML ("    <TD><b>" . $grade[0][$i]);
	 }

     ShowHTML ("  </TR>");
     for($i=1;$i< $Linha;$i++){
        ShowHTML ('  <TR valign="top" align="center">');
        for($j=0; $j<=$Coluna; $j++){
            ShowHTML ('  <TD TITLE="' . $title[$i][$j] . '">' . $grade[$i][$j] );
        }
        ShowHTML ("  </TR>");
	 }
  	 


     ShowHTML ('<tr><td colspan=6 align="left"><TABLE border=0 cellpadding=1>');
     ShowHTML ('  <TR><TD colspan="2"><FONT FACE=VERDANA SIZE=1><B>Legenda das disciplinas:</B>');
     $sql = "SELECT distinct c.sg_disciplina, c.ds_disciplina " . $CrLf .
          "  from sbpi.Aluno_Turma              a " . $CrLf .
          "       INNER join sbpi.Grade_horaria b ON (a.sq_turma        = b.sq_turma and  " . $CrLf .
          "                                         a.sq_cliente = b.sq_cliente and " . $CrLf .
          "                                         a.ano_letivo      = b.ano_letivo and " . $CrLf .
          //"                                         a.ano_letivo      = " . $w_ano . $CrLf .
          "                                        ) " . $CrLf .
          "       INNER join sbpi.Disciplina    c ON (b.sq_disciplina = c.sq_disciplina) " . $CrLf .
          " WHERE a.sq_aluno =" . $AL . " " . $CrLf .
          "ORDER BY c.sg_disciplina" . $CrLf ;


     $RS = db_exec::getInstanceOf($dbms, $sql, &$numRows);


	 foreach ($RS as $row) {
       ShowHTML ("<TR><TD><FONT FACE=VERDANA SIZE=1>" . f($row,"SG_DISCIPLINA") . ":" ) ;
       ShowHTML ("<TD><FONT FACE=VERDANA SIZE=1>" . f($row,"DS_DISCIPLINA") );
       
     }
     ShowHTML ("    </TABLE>");

	}Else{
		 ShowHTML("    <TR><TD colspan=14><b>A grade horária não foi informada.</b>");
	}
  ShowHTML ("    </TABLE>");
}

// =========================================================================
// Tela de mensagens dirigidas ao aluno
// -------------------------------------------------------------------------
function mensagem() {
  extract($GLOBALS);
  ShowHTML('<tr><td>');

  // Recupera os dados cadastrais do aluno
  $SQL = "SELECT a.sq_aluno, a.no_aluno, a.nr_matricula, a.dt_nascimento,  ".$crlf.
        "       case in_sexo when 'M' then 'Masculino' when 'F' then 'Feminino' else '-' end in_sexo, ".$crlf.
        "       a.ds_naturalidade, a.no_mae, a.nr_fone_mae, a.no_pai, a.nr_fone_pai, ".$crlf.
        "       a.no_resposavel, a.nr_fone_responsavel, a.ds_email_responsavel, nr_fone_1, nr_fone_2 ".$crlf.
        "from sbpi.Aluno a ".$crlf.
        "WHERE a.sq_aluno = ".$AL; 
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
  if (count($RS)>0) {
    foreach ($RS as $row) { $RS = $row; break; }
    ShowHTML('<TABLE align="center" WIDTH="100%" BORDER=0>');
    ShowHTML('  <TR valign="TOP">');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Nome:</b> '.f($RS,'NO_ALUNO').'</TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Matrícula:</b> '.f($RS,'NR_MATRICULA').'</TD>');
    ShowHTML('  </TR>');
    ShowHTML('  <TR valign="TOP"><TD COLSPAN="3"><TABLE WIDTH="100%" BORDER="1">');
    ShowHTML('    <TR valign="TOP" align="CENTER">');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Ano Letivo</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Série</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Turno</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Turma</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Etapa/Curso/Modalidade</b></TD>');
    ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1><B>Situação</b></TD>');
    $SQL = "SELECT c.ano_letivo, c.st_aluno, ".$crlf.
          "       b.ds_grau, b.ds_serie, b.ds_turma, trim(replace(b.ds_curso,'¬','ª')) ds_curso, ".$crlf.
          "       case b.ds_turno when 'M' then 'Matutino' when 'V' then 'Vespertino' when 'N' then 'Noturno' when 'I' then 'Integral' else '-' end ds_turno ".$crlf.
          "from sbpi.Aluno_Turma      c ".$crlf.
          "     INNER join sbpi.Turma b ON (b.sq_turma = c.sq_turma and b.sq_cliente = c.sq_cliente and c.ano_letivo = ".$w_ano.") ".$crlf.
          "WHERE c.sq_aluno = ".f($RS,"SQ_ALUNO"); 
    $RS1 = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    
    if (count($RS1)==0) {
      ShowHTML('  <TR><TD colspan="6" ALIGN="CENTER">Turmas não encontradas.</TD></TR>');
    } else {
      foreach($RS1 as $row) {
        ShowHTML('  <TR valign="TOP">');
        ShowHTML('    <TD ALIGN="CENTER"><FONT FACE=VERDANA SIZE=1>'.f($row,'ANO_LETIVO').'</TD>');
        ShowHTML('    <TD ALIGN="CENTER"><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_SERIE').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_TURNO').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'DS_TURMA').'</TD>');
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>');
        ShowHTML(trocanulo(f($row,'DS_CURSO')));
        ShowHTML('    <TD><FONT FACE=VERDANA SIZE=1>'.f($row,'ST_ALUNO').'</TD>');
      }
    }
    ShowHTML('     </TR>');
    ShowHTML('  </TABLE>');
    ShowHTML('</TABLE><br><br>');
  }

  ShowHTML('<tr><td><TABLE border=1 cellpadding=5 cellspacing=0 width="100%" BGCOLOR="#F7F7F7">');
  ShowHTML('  <TR VALIGN="TOP">');
  ShowHTML('    <TD width=""15%""><b>Data');
  ShowHTML('    <TD width=""75%"" align=""left""><b>Texto');
  ShowHTML('    <TD width=""10%""><b>Situação');
  ShowHTML('  </TR>');

  $SQL = "SELECT b.dt_mensagem, b.ds_mensagem, b.in_lida ".$crlf.
         "  from sbpi.Aluno a INNER join sbpi.Mensagem_aluno b ON (a.sq_aluno = b.sq_aluno) ".$crlf.
         " WHERE a.sq_aluno = ".$AL." ".$crlf.
         "ORDER BY dt_mensagem DESC".$crlf;
  $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);

  if (count($RS)>0) {
    foreach($RS as $row) {
      ShowHTML('    <TR valign=""top"" align=""left"">');
      ShowHTML('    <TD>'.formataDataEdicao(f($row,'dt_mensagem')));
      ShowHTML('    <TD>'.trocaNulo(f($row,'ds_mensagem')));
      if (f($row,'in_lida')=='S') ShowHTML('    <TD align="center">Lida');
      elseif (f($row,'in_lida')=='-') ShowHTML('    <TD align="center">-');
      else ShowHTML('    <TD align="center">Nova');
      ShowHTML('    </TR>' );
    }
  } else {
     ShowHTML('    <TR><TD colspan=4><b>Não há mensagens.</b>');
  }
  ShowHTML('    </TABLE>');
  
}

// =========================================================================
// Rotina principal
// -------------------------------------------------------------------------
function Main() {
  extract($GLOBALS);
  
  // Verifica se a senha inicial foi alterada
  if ($par<>'SENHA') {
    $SQL = "select count(*) existe from sbpi.Aluno where nr_matricula = ds_senha_acesso and sq_aluno = ".nvl($AL,0);
    $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    foreach($RS as $row) { $RS = $row; break; }
    if (f($RS,'existe')!='0') {
        ShowHTML('<SCRIPT LANGUAGE=""JAVASCRIPT""><!--');
        ShowHTML('  alert("ALERTA!!!\nUse a opção Troca Senha para alterar sua senha de acesso.");');
        ShowHTML('--></SCRIPT>');
    }
  }

  switch ($par) {
  case 'INICIAL':       Inicial();      break;
  case 'SENHA':         TrocaSenha();   break;
  case 'BOLETIM':       Boletim();      break;
  case 'GRADE':         GradeHoraria(); break;
  case 'MENSAGEM':      Mensagem();     break;
  case 'BOLETIMIMP':    BoletimImp();   break;
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