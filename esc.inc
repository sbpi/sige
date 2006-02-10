<%
InitSystem

Sub InitSiteDefault
REM -------------------------------------------------------------------------
  Session("Default") 				  = "SBPI"

  Session("MetaPragma")       = "NO-CACHE"
  Session("MetaContentType")  = "text/html; charset=iso-8859-1"
  Session("MetaReplyTo")      = "sbpi@sbpi.com.br"
  Session("MetaExpires")      = now + 90
  Session("MetaPicsLabel")    = "(PICS-1.1 ""http://www.rsac.org/ratingsv01.html""" &  _
                                " l gen true comment ""RSACi North America Server"" by" & _
                                " ""sbpi@sbpi.com.br"" for """ & _
                                "http://www2.spbi.com.br/sige""" & _
                                " on ""1997.06.26T21:24-0500"" r (n 0 s 0 v 0 l 0))"
  Session("MetaKeywords")     = "escola, advocacia, eletrônico, eletrônica, virtual, web, site, website, advogada, cliente, andamento, processo, jurídico, jurídica, notícia, legislação, clientela"

  Session("MetaDescription")  = "O SIGE-WEB é uma nova proposta de automação de escolas, baseada na Web, onde as escolas e redes de ensino poderão manter " & _
                                "sites dinâmicos alterando seu conteúdo pela própria Internet. Poderão ser cadastrados os professores das escolas, alunos, disciplinas e tudo mais que a escola desejar. " & _
                                "Assim, o texto de abertura, o quem somos, corpo docente, corpo dicente, disciplinas, bem como a estética do site podem ser atualizados dinamicamente, sem a intervenção de programadores."

  Session("MetaProduct")      = "hipertools v2.5 - Biblioteca ASP"
  Session("MetaLocale")       = "PO-BR"
  Session("MetaCharSet")      = "US-ASCII"
  Session("MetaCategory")     = "HOME PAGE"
  Session("MetaGenerator")    = "SBPI - Sociedade Brasileira para a Pesquisa em Informática"
  Session("MetaAuthor")       = "SBPI - Sociedade Brasileira para a Pesquisa em Informática - Brasilia - DF -" & _
                                " Brasil - webmaster@sbpi.com.br - http://www.sbpi.com.br"
  Session("MetaDistribution") = "global"
  Session("MetaFormatter")    = "HiperTools v3.0"
  Session("MetaTechnology")   = "SCRIPTING"

REM -------------------------------------------------------------------------
REM Constantes BASE FONT de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("BaseFont") = "Arial, Helvetica, Sans-Serif"

REM -------------------------------------------------------------------------
REM Constantes LINK de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("LinkHomeHRef")       = "/Default.asp"
  Session("LinkHomeTitle")      = "HiperTools v3.0"

  Session("LinkCopyrightHRef")  = "/copyright.htm"
  Session("LinkCopyrightTitle") = "Copyright da SBPI"

  Session("LinkMadeHRef")       = "mailto:webmaster@sbpi.com.br"
  Session("LinkMadeTitle")      = "WebMaster da SBPI"

  Session("LinkStyleSheetHRef")  = "/default.css"

REM -------------------------------------------------------------------------
REM Constantes BODY de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("BodyBGColor")    = "white"
  Session("BodyBackground") = "./img/fundo.jpg"
  Session("BodyStyle")      = ""
  Session("BodyText")       = ""
  Session("BodyLink")       = ""
  Session("BodyVLink")      = ""
  Session("BodyALink")      = ""
  Session("BodyTopMargin")  = 0
  Session("BodyLeftMargin") = 0
  Session("BodyWidth")      = 610

REM -------------------------------------------------------------------------
REM Constantes HEADER de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("HeaderIcon")       = ""
  Session("HeaderTitle")      = ""

REM -------------------------------------------------------------------------
REM Constantes FOOTER de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("FooterTitle")      = "" '"Um produto da <A HREF=http://www.sbpi.com.br> SBPI&reg;</A> Sociedade Brasileira para a Pesquisa em Informática."
  Session("FooterSugestions") = "" '"2000 <A HREF =http://www.sbpi.com.br>SBPI&reg;</a> todos os direitos reservados."
  Session("FooterCopyright")  = ""

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub InitSiteDefault

REM -------------------------------------------------------------------------
REM Inicializacao do sistema
REM -------------------------------------------------------------------------
Private Sub InitSystem

  If Not Session("Default") = "SBPI" Then
    InitSiteDefault

  End If

  REM -----------------------------------------------------------------------
  REM Inicializa os defaults do sistema
  REM -----------------------------------------------------------------------
  If Request.QueryString("Default") > "" Then
      InitDefault

  End If

  REM -----------------------------------------------------------------------
  REM Ativa o icone de ASP Source Code
  REM -----------------------------------------------------------------------
  REM   If Request.QueryString("Source") > "" Then
  REM     Select Case Request.QueryString("Source")
  REM       Case 0
         Session("SourceCode") = False
  REM 
  REM       Case 1
  REM         Session("SourceCode") = True
  REM 
  REM       Case Else
  REM     End Select
  REM   Else
  REM     If IsEmpty(Session("SourceCode")) Then Session("SourceCode") = False
  REM 
  REM   End If

  REM -----------------------------------------------------------------------
  REM Define a borda das tabelas
  REM -----------------------------------------------------------------------
  If Request.QueryString("Border") > "" Then
    Session("TableBorder") = Request.QueryString("Border")

  Else
    If IsEmpty(Session("TableBorder")) Then Session("TableBorder") = 0

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub InicioSistema
%>