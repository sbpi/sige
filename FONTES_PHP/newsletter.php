<?php
// Define o banco de dados a ser utilizado
$_SESSION['DBMS']=1;

$w_dir_volta = '';
$w_dir = '';
$w_pagina = 'newsletter.php?';

include_once('constants.inc');
include_once('jscript.php');
include_once('funcoes.php');
include_once('classes/db/abreSessao.php');
include_once('classes/sp/db_exec.php');

// =========================================================================
//  /newsletter.php
// ------------------------------------------------------------------------
// Nome     : Alexandre Vinhadelli Papadópolis
// Descricao: Mecanismo de newslettler
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
  ShowHTML('   <title>Secretaria de Estado de Educa&ccedil;&atilde;o</title>');
  ShowHTML('   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /> ');
  ShowHTML('   <link href="css/estilo.css" media="screen" rel="stylesheet" type="text/css" />');
  ShowHTML('   <link href="css/print.css"  media="print"  rel="stylesheet" type="text/css" />');
  ShowHTML('   <script language="javascript" src="js/scripts.js"> </script>');
  ShowHTML('   <script src="js/jquery.js"></script>');
 
  ScriptOpen ("JavaScript");
  ValidateOpen ("Validacao");
  if($par == 'INICIAL'){
    Validate ("w_nome", "Nome", "", "1", "3", "60", "1", "1");
  }

  Validate ("w_email", "e-Mail", "", "1", "4", "60", "1", "1");
  if($par == 'INICIAL'){
    ShowHTML ("  if (theForm.w_tipo[0].checked==false && theForm.w_tipo[1].checked==false && theForm.w_tipo[2].checked==false) {");
    ShowHTML ("     alert('Você deve selecionar uma das opções apresentadas no formulário!');");
    ShowHTML ("     return false;");
    ShowHTML ("  }");
  
    ShowHTML ("  theForm.Botao[0].disabled=true;");
    ShowHTML ("  theForm.Botao[1].disabled=true;");
  }
  ValidateClose();
  ScriptClose();  
  
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
  ShowHTML('    <li><a href="default.php"><span>Inicial</span></a></li>');
  ShowHTML('    <li><a href="newsletter.php?par=inicial"><span>Assine nosso boletim</span></a></li>');
  ShowHTML('  </ul>');
  ShowHTML('  <div class="clear"></div>');
  ShowHTML('</div>');
  ShowHTML('<div id="conteudo"><h2>Solução Integrada de Gestão Educacional - SIGE</h2>');

  Main();

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

// Fecha conexão com o banco de dados
if (isset($_SESSION['DBMS'])) FechaSessao($dbms);

exit;

// =========================================================================
// Monta tela de entrada de dados para lista de distribuição
// -------------------------------------------------------------------------
function Inicial() {
  extract($GLOBALS);

  ShowHtml ('    <div id="texto"><!-- Conteúdo --></div>');
  ShowHtml ('        <table width="570" border="0">');
  ShowHtml ('        <FORM ACTION="'.$w_pagina.'par=grava" method="POST" name="Form" onSubmit="return(Validacao(this));">');
  ShowHtml ('        <input type="hidden" name="SG" value="Cadastro">');
  ShowHtml ('          <TR><TD align="left" valign="middle">Informe os dados solicitados no formulário abaixo e clique no botão "Registrar" para registrar-se em nossa lista de distribuição.');
  ShowHtml ('          <TR><TD>Informe seu nome completo:<br><input type="text" size="60" maxlength="60" name="w_nome" value="" CLASS="texto">');
  ShowHtml ('          <TR><TD>Informe o e-Mail que deseja receber a newsletter:<br><input type="text" size="60" maxlength="60" name="w_email" value="" CLASS="texto">');
  ShowHtml ('          <TR><TD>Selecione uma das opções abaixo: ');
  ShowHtml ('                <br><input type="Radio" name="w_tipo" value="1"> Pai, mãe ou responsável por aluno da rede de ensino ');
  ShowHtml ('                <br><input type="Radio" name="w_tipo" value="2"> Aluno da rede de ensino ');
  ShowHtml ('                <br><input type="Radio" name="w_tipo" value="3"> Outro ');
  ShowHtml ('          <TR><TD align="center"><font size="2" CLASS="BTM">');
  ShowHtml ('               <input type="submit" name="Botao" value="Registrar" class="botao">');
  ShowHtml ('               <input type="button" name="Botao" value="Voltar" class="botao" onClick="javascript: history.back(1);" >');
  ShowHtml ('        </FORM>' . $VbCrLf);
  ShowHtml ('<br/><br/> Se desejar remover seu e-mail, clique <a class="SS" href="'.$w_pagina.'par=remove" target="_self">aqui</a>');
  
}


// =========================================================================
// Remove e-mail da lista de distribuição
// -------------------------------------------------------------------------
function remove(){
  extract($GLOBALS);
  ShowHTML ('        <table width="570" border="0">');
  ShowHtml ('        <FORM ACTION="'.$w_pagina.'par=grava" method="POST" name="Form" onSubmit="return(Validacao(this));">');
  ShowHTML ('        <input type="hidden" name="SG" value="Remove">');
  ShowHTML ('          <TR><TD>Informe seu e-mail no campo abaixo e clique no botão "Remover" para ser removido da nossa lista de distribuição.');
  ShowHTML ('          <TR><TD>Informe o e-Mail que deseja remover:<br><input type="text" size="60" maxlength="60" name="w_email" value="" CLASS="texto">');
  ShowHTML ('          <TR><TD align="center">');
  ShowHTML ('               <input type="submit" name="Botao" value="Remover" class="botao">');
  ShowHTML ('               <input type="button" name="Botao" value="Voltar" class="botao" onClick="javascript: history.back(1);">');
  ShowHTML ('        </FORM>');

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


function grava(){
   extract($GLOBALS);
  Cabecalho();
  //ShowHTML ("</HEAD>");
  //BodyOpen( "onLoad=document.focus();");

  switch($SG){
    Case "CADASTRO":
       
       $SQL = "select envia_mail from sbpi.Newsletter where upper(email) = '" . strtoupper(trim($_REQUEST["w_email"])) . "' " ;
     $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
        foreach($RS as $row) { $RS = $row; break; }

    if(f($row,'envia_mail')== 'S'){
      ScriptOpen ("JavaScript");
          ShowHTML   ("  alert('O e-mail informado já existe em nossa lista de distribuição!');" );
          ShowHTML   ("  location.href='$pagina?par=inicial'");
          ScriptClose();
      die();
    }
       

    if (f($row,'envia_mail')=='N')
    {

       $SQL = "update sbpi.Newsletter set " . $VbCrLf .
                   "   envia_mail     = 'S', " . $VbCrLf .
                   "   nome           = '" . trim($_REQUEST["w_nome"]) . "', " . $VbCrLf .
                   "   email          = '" . trim($_REQUEST["w_email"]) . "', " . $VbCrLf .
                   "   tipo           = '" .  $_REQUEST["w_tipo"] . "', " . $VbCrLf .
                   "   data_alteracao = sysdate " . $VbCrLf .
                   " where upper(email) = '" . strtoupper(trim($_REQUEST["w_email"])) . "' ";

    }else{

          $SQL = "insert into sbpi.Newsletter (SQ_NEWSLETTER,sq_cliente, nome, email, tipo, envia_mail, data_inclusao, data_alteracao) " . $VbCrLf .
                "  values ( sbpi.SQ_NEWSLETTER.nextval,0," . $VbCrLf .
                "          '" . trim($_REQUEST["w_nome"]) . "', " . $VbCrLf .
                "          '" . trim($_REQUEST["w_email"]) . "', " . $VbCrLf .
                "          '" . $_REQUEST["w_tipo"] . "', " . $VbCrLf .
                "          'S', " . $VbCrLf .
                "          sysdate, " . $VbCrLf .
                "          sysdate " . $VbCrLf .
                "         )" . $VbCrLf ;
    }


    db_exec::getInstanceOf($dbms, $SQL, &$numRows);

      ScriptOpen ('JavaScript');
          ShowHTML   ('  alert("Seu e-mail foi gravado com sucesso!\n A partir de agora você faz parte da lista de distribuição de nossa newsletter.");' );
          ShowHTML   ('  location.href="default.php";');
          ScriptClose();
      die();
      
         //  Envia e-mail comunicando a inclusão
         PreparaMail(trim($_REQUEST["w_nome"]), trim($_REQUEST["w_email"]), $_REQUEST["w_tipo"],1);
  case 'REMOVE':
    $SQL = "select envia_mail from sbpi.Newsletter where upper(email) = '" . strtoupper(trim($_REQUEST["w_email"])) . "'" ;
      $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    if(count($RS) < 1){
      ScriptOpen ("JavaScript");
          ShowHTML   ("  alert('O e-mail informado não existe em nossa lista de distribuição!');" );
          ShowHTML   ("  location.href='$pagina?par=inicial'");
          ScriptClose();
      die();
    }

    $SQL = "select envia_mail from sbpi.Newsletter where upper(email) = '" . strtoupper(trim($_REQUEST["w_email"])) . "' and envia_mail = 'N' " ;
      $RS = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
    
    if(count($RS) > 0){
      ScriptOpen ("JavaScript");
          ShowHTML   ("  alert('O e-mail informado já foi removido da nossa lista de distribuição!');" );
          ShowHTML   ("  location.href='$pagina?par=inicial'");
          ScriptClose();
      die();
    }

    $SQL = "update sbpi.Newsletter set envia_mail = 'N' , data_alteracao = sysdate where upper(email) = '" . strtoupper(trim($_REQUEST["w_email"])) . "' and envia_mail = 'S' " ;
      $RS  = db_exec::getInstanceOf($dbms, $SQL, &$numRows);
      ScriptOpen ("JavaScript");
        ShowHTML   ("  alert('Seu e-mail foi removido com sucesso!\\nA partir de agora você não faz parte da lista de distribuição de nossa newsletter.');" );
        ShowHTML   ("  location.href='$pagina?par=inicial'");
        ScriptClose();
    
       
  }

}


function PreparaMail($p_nome, $p_mail, $p_tipo, $p_evento){

  
  
  $w_html = '<HTML>' . $VbCrLf;
  $w_html .=  BodyOpenMail(null) . $VbCrLf;
  $w_html .=  '<table border="0" cellpadding="0" cellspacing="0" width="100%">' . $VbCrLf;
  $w_html .=  '<tr bgcolor="" . conTrBgColor . ""><td align="center">' . $VbCrLf;
  $w_html .=  '    <table width="97%" border="0">' . $VbCrLf;
  If ($p_evento == 1 ){
     $w_html .=  '      <tr valign="top"><td align="center"><font size=2><b>INCLUSÃO NA LISTA DE DISTRIBUIÇÃO</b></font><br><br><td></tr>' . $VbCrLf;
     $w_html .=  '      <tr align="center"><td><font size=2><b><font color="#BC3131">ATENÇÃO</font>: Esta é uma mensagem de envio automático. Não responda.</b></font><br><br><td></tr>' . $VbCrLf;
     $w_html .=  $VbCrLf . '<tr bgcolor="" . conTrBgColor . ""><td align="center">';
     $w_html .=  $VbCrLf . '      <tr><td><font size=2>Seu e-mail foi incluído na lista de distribuição de informativos da Secretaria de Estado de Educação do DF. A partir de agora, sempre que algum informativo for enviado para a lista, você será um dos destinatários.</b></font></td></tr>';
     $w_html .=  '      <tr align="center"><td><br><br><br><font size=1>' . $VbCrLf;
     $w_html .=  '         Se desejar remover seu e-mail, clique <a class="SS" href="http://www.gdfsige.df.gov.br/newsletter.php?par=Remover" target="_blank">aqui</a>.' . $VbCrLf;
     $w_html .=  '      </font></td></tr>' . $VbCrLf;
  }ElseIf ($p_evento == 2 ){
     $w_html .=  '      <tr valign="top"><td align="center"><font size=2><b>SAÍDA DA LISTA DE DISTRIBUIÇÃO</b></font><br><br><td></tr>' . $VbCrLf;
     $w_html .=  '      <tr align="center"><td><font size=2><b><font color="#BC3131">ATENÇÃO</font>: Esta é uma mensagem de envio automático. Não responda.</b></font><br><br><td></tr>' . $VbCrLf;
     $w_html .=  $VbCrLf . '<tr bgcolor="" . conTrBgColor . ""><td align="center">';
     $w_html .=  $VbCrLf .  '     <tr><td><font size=2>Seu e-mail foi removido da lista de distribuição de informativos da Secretaria de Estado de Educação do DF.</b></font></td></tr>';
     $w_html .=  '      <tr align="center"><td><br><br><br><font size=1>' . $VbCrLf;
     $w_html .=  '         Se desejar cadastrar novamente seu e-mail, clique <a class="SS" href="http://www.gdfsige.df.gov.br/newsletter.php?par=Remover" target="_blank">aqui</a>.' . $VbCrLf;
  }
  $w_html .=  '      <tr valign="top"><td><br><br><br><font size=1>' . $VbCrLf;
  $w_html .=  '         Dados da ocorrência:<br>' . $VbCrLf;
  $w_html .=  '         <ul>' . $VbCrLf;
  $w_html .=  '         <li>Data do servidor: <b>' . date('d/m/Y , G:i:s') .   '</b></li>' . $VbCrLf;
  $w_html .=  '         <li>IP de origem: <b>' . $_SERVER["REMOTE_ADDR"] . '</b></li>' . $VbCrLf;
  $w_html .=  '         </ul>' . $VbCrLf;
  $w_html .=  '      </font></td></tr>' . $VbCrLf;
  $w_html .=  '    </table>' . $VbCrLf;
  $w_html .=  '</td></tr>' . $VbCrLf;
  $w_html .=  '</table>' . $VbCrLf;
  $w_html .=  '</BODY>' . $VbCrLf;
  $w_html .=  '</HTML>' . $VbCrLf;

  //Prepara os dados necessários ao envio
  If ($p_evento == 1){ // Inclusão ou Conclusão
     $w_assunto = "SEDF - Inclusão na lista de distribuição de informativos";
  }ElseIf ($p_evento == 2 ){ // Tramitação
     $w_assunto = "SEDF - Saída da lista de distribuição de informativos";
  }

  // Executa o envio do e-mail
  $w_resultado = EnviaMail($w_assunto, $w_html, $p_mail, null);
        
  // Se ocorreu algum erro, avisa da impossibilidade de envio
  If (w_resultado > ""){
     ScriptOpen ("JavaScript");
     ShowHTML ("  alert('ATENÇÃO: não foi possível proceder o envio do e-mail.\\n" . $w_resultado . "');"); 
     ScriptClose();
  }

}

// =========================================================================
// Rotina principal
// -------------------------------------------------------------------------
function Main() {
  extract($GLOBALS);

  switch ($par) {
    case 'INICIAL':    
         Inicial();
        break;
  case 'GRAVA':  
         grava();
        break;
    case 'REMOVE':  
         remove();
        break;
    default :
    Cabecalho();    
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
  exibevariaveis();
    break;
  }  
}
?>