<?
//---------------------------Defini��es do BODY
$conBodyText		= '#000000';
$conBodyLink		= '#0000FF';
$conBodyALink		= '#0000FF';
$conBodyVLink		= '#0000FF';
$conBodyBgColor		= '#F8F8FF';
$conBodyBackground	= '';
$conBodyBgproperties	= 'fixed';
$conBodyTopmargin	= '5';
$conBodyLeftmargin	= '5';

//---------------------------Defini��es de Tabela
$conTableHeader			= '#CCCCCC';
$conTableBgColor		= '#F7F7F7';
$conTableAlternateBgColor	= '#E1F2FF';
$conTableBorder			= '0';
$conTableBorderColorDark	= '#000000';
$conTableBorderColor		= '#000000';
$conTableBorderColorlight	= '#FFFFFF';
$conTableCellPadding		= '1';
$conTableCellSpacing		= '1';
$conTrBgColor			= '#EFEFEF';
$conTrTotalBgColor		= '#55BBFF';
$conTrAlternateBgColor		= '#E0E0E0';
$conTrRelatorioBgColor		= '#E1F2FF';
$conTrBgColorLightBlue1		= '#BBBBFF';
$conTrBgColorLightRed1		= '#FF9999';
$conTrBgColorLightBlue2		= '#AADDFF';
$conTrBgColorLightRed2		= '#FFCCCC';

//---------------------------Defini��es de fontes
$conFontSize		= '2';
$conFontFace		= 'Arial, Helvetica, sans-serif';
$conFontColor		= '#000000';
$conTituloFontSize	= '3';
$conTituloFontFace	= 'Arial';
$conTituloFontColor	= '#008000';

//---------------------------Defini��es de imagens
$conImgAlHigh	= 'images/icone/alert_animated.gif';
$conImgAlMed	= 'images/icone/alert_red.gif';
$conImgAlLow	= 'images/icone/alert_yellow.gif';

$conImgOkAcima	= 'images/tickazul.gif';
$conImgOkNormal	= 'images/tickverde.gif';
$conImgOkAtraso	= 'images/tickvermelho.gif';

$conImgStAcima	= 'images/icone/arrow_blue.gif';
$conImgStNormal	= 'images/icone/arrow_green.gif';
$conImgStAviso	= 'images/icone/arrow_yellow.gif';
$conImgStAtraso	= 'images/icone/arrow_red.gif';

$conImgNormal	= 'images/BallGreen.gif';
$conImgAviso	= 'images/BallYellow.gif';
$conImgAtraso	= 'images/BallRed.gif';
$conImgCancel 	='images/icone/REQUIRED.GIF';
$conImgProblem  ='images/icone/problem.gif';
$conImgRiskLow  ='images/icone/risk_low.gif';
$conImgRiskMed  ='images/icone/risk_low.gif';
$conImgRiskHig  ='images/icone/risk_medium.gif';
$conImgDownload ='images/icone/application_put.png';

//$conImgSmNormal = 'images/icone/smile_approve.gif';
//$conImgSmAviso  = 'images/icone/smile_sleepy.gif';
//$conImgSmAtraso = 'images/icone/smile_dissapprove.gif';

$conImgSmNormal = 'images/BallGreen.gif';
$conImgSmAviso  = 'images/BallYellow.gif';
$conImgSmAtraso = 'images/BallRed.gif';

$conImgGeo      = 'images/icone/earth_16.png';
$conImgSheet     = 'images/Folder/SheetImg.gif';

//---------------------------Defini��es de banco

if ($_SESSION['DBMS']!=2) {
  $strschema	= 'sbpi.';
} else {
  $strschema	= 'dbo.';
}
$strschema_is	= 'sbpi.';
$conPageSize	= 30;
$conRefreshSec  = 1000;

//---------------------------Configura��o do ambiente do sistema

// Define se o servi�o de envio autom�tico de e-mails deve ser ativado (true) fou n�o (false)
$conEnviaMail   = true;

$crlf           = chr(13).chr(10);
$conSgSistema   = 'SIGE-WEB';
$conApacheStat  = 'http://www2.sbpi.com.br/server-status-sbpi';
$conDefaultPath = 'http://www2.sbpi.com.br';
$conNmSistema   = 'SIGE-WEB';
$conRootSIW     = 'http://www2.sbpi.com.br/sigeweb/';
$conDiretorio   = '/var/www/html/sigeweb/';
$conFontPath    = $conDiretorio.'fonts/linux/';
$conFileVirtual = '/sigeweb/sedf/';
$conFilePhysical= '/var/www/html/sigeweb/sedf/'; // Formato Linux
//$conFilePhysical= 'c:/inetpub/wwwroot/sigeweb/sedf/'; // Formato Windows (IIS)

// Define se deve ser gravado arquivo de log e o seu caminho
$conLog         = true;
$conLogPath     = $conFilePhysical.'log/';
?>