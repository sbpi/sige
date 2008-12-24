<%
REM =========================================================================
REM  /Constants.inc
REM -------------------------------------------------------------------------
REM Nome     : Constates do SIGE-WEB
REM Descricao: Menu para o SIGE-WEB
REM Home     : http://www2.sbpi.com.br/sige
REM Criacao  : 13/03/2000 10:00	
REM Autor    : SBPI - Sociedade Brasileira para a Pesquisa em Informática
REM Versao   : 1.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by SBPI
REM -------------------------------------------------------------------------

  Const conFilePhysical           = "c:\inetpub\wwwroot\sige_ora\"

  Const conVirtualPath = "/"
  Const conSMTPServer = "smtp.cabonet.com.br"
  Const conScriptTimeout    = 300
  Const conSessionTimeout   = 300
  Const conConnectionString = "Provider=OraOLEDB.Oracle.1;Persist Security Info=True;User ID=sigecon;Password=sigecon;Data Source=MERCURIO"
  Const conConnectionManut  = "Provider=OraOLEDB.Oracle.1;Persist Security Info=True;User ID=sige_web;Password=sige_web;Data Source=MERCURIO"
  Const conProduto			= "SIGE-WEB"
  conSite				= "http://" & Request.ServerVariables("HTTP_HOST")
  Const conMail				= "sbpir@sbpi.com.br"
  Const conHead				= "SIGE Web"
  Const conLogo				= "logoescola.jpg"
  Const conRelEscolas                   = "900"
  Const conWhatEstagio                  = "103"
  Const conWhatContrata5                = "104"
  Const conWhatContrata4                = "105"
  Const conWhatContrata3                = "106"
  Const conWhatContrata2                = "107"
  Const conWhatContrata1                = "108"
  Const conWhatCurso                    = "109"
  Const conWhatIntercambio              = "110"
  Const conWhatEmprego                  = "111"
  Const conWhatFormando                 = "112"
  Const conWhatTipoLink                 = "113"
  Const conWhatModelo                   = "114"
  Const conWhatSenha                    = "115"
  Const conWhatValida                   = "116"
  Const conWhatInstitucional            = "117"
  Const conWhatFale                     = "118"
  Const conWhatLink                     = "119"
  Const conWhatConsulta                 = "120"
  Const conWhatPrincipal                = "121"
  Const conWhatCabecalho                = "122"
  Const conWhatCliente                  = "123"
  Const conWhatTipoCliente              = "124"
  Const conWhatEspecialidade            = "125"
  Const conWhatEspecialidadeTipoCliente = "126"
  Const conWhatEspecialidadeCliente     = "127"
  Const conWhatUsuario                  = "128"
  Const conWhatPerfil                   = "129"
  Const conWhatNoticia                  = "130"
  Const conWhatArquivo                  = "131"
  Const conWhatCalendario               = "133"
  Const conWhatProcesso                 = "134"
  Const conWhatDocumento                = "135"
  Const conWhatTopo                     = "136"
  Const conWhatEsquerda                 = "137"
  Const conWhatDireita                  = "138"
  Const conWhatRodape                   = "139"
  Const conWhatEsquerda1                = "140"
  Const conWhatSite                     = "144"
  Const conWhatServico                  = "201"
  Const conWhatServicoCliente           = "202"
  Const conWhatClienteDados             = "203"
  Const conWhatAndamento                = "204"
  Const conWhatEstrutura                = "205"
  Const conWhatAcesso                   = "206"
  Const conWhatParceiro                 = "207"
  Const conWhatProduto                  = "208"
  Const conWhatIndice                   = "209"
  Const conWhatItemIndice               = "210"
  Const conWhatSistema                  = "211"
  Const conWhatSistemaVersao            = "212"
  Const conWhatRegiao                   = "213"
  Const conWhatParceiroServico          = "214"
  Const conWhatAgenda                   = "215"
  Const conWhatCargo                    = "216"
  Const conWhatInscreve2                = "217"
  Const conWhatInscreve3                = "218"
  Const conListEmpresaUF                = "101"
  Const conWhatCurriculo                = "219"
  Const conWhatDadosAdicionais		= "220"
  Const conWhatMensagem			= "221"
  Const conWhatNotCliente		= "222"
  Const conWhatAdmin   			= "223"
  Const conWhatSGE   			= "224"

  Const conListCampoTipo          = "141"
  Const conListCampoDelimitador   = "142"

  ' Constantes do sistema
  Const conAspas             = """"
  Const conConnectionTimeout = 300
  Const conCommandTimeout    = 300
  Const conSelectAll         = "SELECT * FROM "
  Const conTotalBlocks       = "/8 - "
  Const conBackTo            = "<BR>Voltar para o cadastro "
  Const conToMakeSystem      = "MS"
  Const conOptions           = "O"



REM -------------------------------------------------------------------------
REM Fim do Constants.inc
REM =========================================================================
%>