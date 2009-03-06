<?
extract($GLOBALS);
include_once($w_dir_volta.'classes/db/DatabaseQueriesFactory.php');
include_once($w_dir_volta.'classes/db/DatabaseQueries.php');
/**
* class db_exec
*
* { Description :- 
*    Executa comandos SQL no banco de dados
* }
*/

class db_exec {
   function getInstanceOf($dbms, $p_sql, $numRows) {
     extract($GLOBALS,EXTR_PREFIX_SAME,'strchema');
     $l_rs = DatabaseQueriesFactory::getInstanceOf($p_sql, $dbms, null, DB_TYPE);
     if($l_rs->executeQuery()) { 
       $numRows  = $l_rs->getNumRows();
       if ($l_rs = $l_rs->getResultData()) {
         return array_stripslashes($l_rs); 
       } else {
         return array();
       }
     }
   }   
}

function array_stripslashes($array){
	$saida = array();
	foreach($array as $k=>$v){
		if(is_array($v)){
			$v = array_stripslashes($v);
		}else{
			$v = str_replace('""','"',stripslashes($v));			
		}
		$saida[$k] = $v;
	}
	return $saida;
  }
?>
