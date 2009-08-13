//Atribuição de eventos
function inittree(){
	var uls=document.getElementsByTagName("ul")
	for(i=0;i<uls.length;i++)
		if(uls[i].className=="treelist"){
			var lis=uls[i].childNodes
			for(ii=0;ii<lis.length;ii++)
				if(lis[ii].nodeType==1)
					if(lis[ii].getElementsByTagName("ul").length>0){
						lis[ii].className="fechado"
						chi=lis[ii].childNodes
						addEvent(lis[ii].childNodes[0],"click",clicado)
					}
		}
}

//Abre/fecha quando clicado
function clicado(e){
	var source=getSource(e)
	var li=source.parentNode
	li.className=li.className=="fechado"?"aberto":"fechado"
	return false
}