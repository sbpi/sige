// ==================================================
function fechaImp(){
	$id('janela_impressao').style.display="none"; 
	$id('fImp').style.display="none"; 
}
function janelaImpressao(){
	var objBody = $tag("body").item(0);

	//=============================================
	// CRIA DIV "fImp" - Fundo pra area de impressao ter destaque
	$obj.create("div","fImp",objBody);
	$id("fImp").style.display="none";
	
	//=============================================
	// CRIA DIV "janela_impressao" - Janela onde mostrara a area a ser impressa
	$obj.create("div","janela_impressao",objBody);
	$id("janela_impressao").style.display = "none";	
	
	//=============================================
	// CRIA DIV "topo_janela" - Topo da Janela
	$obj.create("div","topo_janela",$id("janela_impressao"));

	//=============================================	
	// CRIA FORM
	$obj.create("form","form_impressao",$id("topo_janela"));

	//=============================================	
	// CRIA H1 "TITULO" - TITULO DA JANELA
	$obj.create("h1","titulo",$id("form_impressao"));
	$id("titulo").innerHTML = "Pré-visualização de Impressão";

	//=============================================	
	// CRIA INPUT "conf_imp" - CONFIRMA IMPRESSAO
	$obj.create("input","conf_imp",$id("form_impressao"),"button");
	
	$id("conf_imp").setAttribute('value','Confirma Impressão');

	if(/Internet Explorer/.test(navigator.appName)){ // IE sempre ele
		$id("conf_imp").setAttribute('onclick',confImpressao);
		$id("conf_imp").setAttribute('className','button');
	}else{ 
		$id("conf_imp").setAttribute('onclick',"confImpressao()");
		$id("conf_imp").setAttribute('class','button');
	}
	
	//=============================================	
	// CRIA INPUT "fechaJanela"
	$obj.create("input","fechaJanela",$id("form_impressao"),"button");
	
	$id("fechaJanela").setAttribute('value','Fechar');
	$id("fechaJanela").setAttribute('type','button');
	$id("fechaJanela").setAttribute('className','button');
	
	if(/Internet Explorer/.test(navigator.appName)){ // IE sempre ele
		$id("fechaJanela").setAttribute('onclick',fechaImp);
		$id("fechaJanela").setAttribute('className','button');
	}else{ 
		$id("fechaJanela").setAttribute('onclick',"fechaImp()");
		$id("fechaJanela").setAttribute('class','button');
	}
	
	//=============================================	
	// CRIA IFRAME "conteudo_impressao" - CONTENUTO
	$obj.create("iframe","conteudo_impressao",$id("janela_impressao"));
	$id("conteudo_impressao").setAttribute('src','imprime.htm');
	//=============================================	
}
function imprime(cont){
	// jogamos a tela pro topo
	scrollTo(0,0);
	
	telaH = document.body.offsetHeight;
	navH = document.documentElement.clientHeight;
	
	if(navH>telaH)
		telaH=navH;
	telaH+= "px";
	$id("fImp").style.height = telaH;
	
	$id("fImp").style.display='block';
	$id("janela_impressao").style.display='block';
	effects.fade(null,"janela_impressao", 0, 100)

	var conteudo = $id(cont).innerHTML;
	
	parent.conteudo_impressao.document.getElementById("visualiza_imp").innerHTML = conteudo;
}
function confImpressao(){
	parent.conteudo_impressao.stIE();
}
addEvent(window,"load",janelaImpressao);