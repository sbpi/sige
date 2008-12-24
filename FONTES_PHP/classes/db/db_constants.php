<?php
// Database Constants
define("ORA9_SERVER_NAME", "XE.localdomain");
define("ORA9_VERSION_TEXT", "Oracle Server 10g");
define("ORA9_DATABASE_NAME", "sbpi");

define("ORA9_DB_USERID_PORTAL", "sigecon");
define("ORA9_DB_PASSWORD_PORTAL", "sigecon");

define("ORA9_DB_USERID_DBO", "sbpi");
define("ORA9_DB_PASSWORD_DBO", "sbpi");

define("ORA9_DB_USERID_SYS", "sige_web");
define("ORA9_DB_PASSWORD_SYS", "sige_web");

define("DATABASE_NAME", ORA9_DATABASE_NAME);
define("DATABASE_VERSION", ORA9_VERSION_TEXT);
define("B_VARCHAR", null);
define("B_NUMERIC", null);
define("B_CURSOR", OCI_B_CURSOR);
define("B_REQUIRED", true);
define("B_OPTIONAL", false);

switch ($_SESSION["DBMS"]) {
   case 1 : {
     define("ORA9_DB_USERID", ORA9_DB_USERID_PORTAL);
     define("ORA9_DB_PASSWORD", ORA9_DB_PASSWORD_PORTAL);
     break;
   }
   case 2 : {
     define("ORA9_DB_USERID", ORA9_DB_USERID_SYS);
     define("ORA9_DB_PASSWORD", ORA9_DB_PASSWORD_SYS);
     break;
   }
   case 3 : {
     define("ORA9_DB_USERID", ORA9_DB_USERID_DBO);
     define("ORA9_DB_PASSWORD", ORA9_DB_PASSWORD_DBO);
     break;
   }
}
?>
