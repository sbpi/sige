// FS Framework - http://www.fazedordesite.com

// ALIAS
function $id() {
  var elements = new Array();
  var totArgs = arguments.length;
  for (var i = 0; i < totArgs; i++) {
    if (typeof arguments[i] == 'string') arguments[i] = document.getElementById(arguments[i]);
    if (arguments.length == 1) return arguments[i];
    elements.push(arguments[i]);
  }
  return elements;
}
function fsClass(){
    return function() {
      this.initialize.apply(this, arguments);
	}
}
var Class = {
  create: function() {
    return function() {
      this.initialize.apply(this, arguments);
    }
  }
}
function $tag(t){
	return document.getElementsByTagName(t);
}
// EVENT
function addEvent(object, eventType, doIt){
	if(object.addEventListener){ // ALL BROWSERS
		object.addEventListener(eventType, doIt, false); 
		return true;
	} else if (object.attachEvent){ // IE	
		var r = object.attachEvent('on'+eventType, doIt);
		return r;
	} else 	return false;
}
// ================
var $obj = {
	create: function(type, id, parentObj,formType){
		this.newObj = document.createElement(type);
		if(id) this.newObj.setAttribute('id',id);
		if(formType) this.newObj.type=formType;
		if(type=="form" || type=="iframe" || formType) if(id) this.newObj.setAttribute('name',id);
		
		parentObj.appendChild(this.newObj);
	},
	clone: function(orObj,cloneObj,cloneType){
		this.dimension = Array();
		this.position = Array();
		this.dimension = this.getDimension(orObj);
		this.position = this.getPos(orObj);
		$obj.create(cloneType,cloneObj,$tag("body").item(0));
		$id(cloneObj).style.position='absolute';
		$id(cloneObj).style.left=this.position[1]+"px";
		$id(cloneObj).style.top=this.position[0]+"px";
		$id(cloneObj).style.width=this.dimension[1]+"px";
		$id(cloneObj).style.height=this.dimension[0]+"px";
		$id(cloneObj).style.zIndex=999;
	},
	opacity: function(obj,value){
		var bsv = value/100; 
		$id(obj).style.opacity=bsv;
		$id(obj).style.MozOpacity=bsv;
		$id(obj).style.filter="alpha(opacity="+value+")";
	},
	getPos: function(obj){
		if($id(obj)){
			obj = $id(obj);
			var top=obj.offsetTop;
			objY=obj;
			while((objY=objY.offsetParent) != null) top+=objY.offsetTop;
			var left=obj.offsetLeft;
			objX=obj;
			while((objX=objX.offsetParent)!=null) left+=objX.offsetLeft;
			var position=Array(top,left);
			return position;
		} else return Array("ERROR:","Object '"+obj+"' not found");
	},
	getDimension: function(obj){
		if($id(obj)){
			obj = $id(obj);
			var height=obj.offsetHeight;
			var width=obj.offsetWidth;
			var dimension=Array(height,width);
			return dimension;
		} else return Array("ERROR:","Object '"+obj+"' not found");
	}
}
// ==========================================
// AJAX
function ajax(url, div, info, method, formID, script, func){
	this.url = url; // content url
	this.method = (method) ? method : 'GET'; // method GET or POST, by standard  "GET"
	this.div = div; // content destination
	this.info = (info) ? info : 'loading'; // while loading
	this.formID = formID; // form ID is required for post requests only
	this.script = (script==0 || script==1) ? script : 1;
	this.func = func;
}
ajax.prototype = {
    connect: function(){
		if(!this.url) return;
        this.xmlHttp = null;
		if( window.XMLHttpRequest ) this.xmlHttp = new XMLHttpRequest(); // All browsers
		else if( window.ActiveXObject){ // IE
			try{
				this.xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
			} catch (e){
				try{
					this.xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
				} catch (e) {}
			}
		}
		if(this.xmlHttp != null && this.xmlHttp != undefined ){
			var object = this;
			this.xmlHttp.onreadystatechange = function(){ object.getState.call(object); }
			
			if(this.method=="GET") this.executeGET(); // get request
			else this.executePOST() // post request		
		}
	},
	getState: function(){
		if(this.xmlHttp.readyState == 4 ){
			this.result(this.xmlHttp.responseText);
			if(this.script) this.executeScripts(this.xmlHttp.responseText);
			if(this.func!=null) eval(this.func);
		} else this.loading();
	},
	executeGET: function(){
		this.xmlHttp.open(this.method,this.url, true);
		this.xmlHttp.send(null);
		
	},
	executePOST: function(){
		var fields = "";
		if(!this.formID) alert("Erro: falta indicar FORM")
		for(i=0;i<$id(this.formID).length;i++){
			fields+=$id(this.formID)[i].name+"="+$id(this.formID)[i].value+"&";
		}
		this.xmlHttp.open(this.method,this.url, true);
		this.xmlHttp.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
		this.xmlHttp.setRequestHeader('Content-Type',"application/x-www-form-urlencoded; charset=iso-8859-1");
		this.xmlHttp.send(fields);		
	},	
	executeScripts: function(text){
		var ini = 0;
		while (ini!=-1){
			ini = text.indexOf('<script', ini);
			if (ini >=0){
				ini = text.indexOf('>', ini) + 1;
				var fim = text.indexOf('</script>', ini);
				text = text.substring(ini,fim);
				eval(text);
			}
		}
	},
    loading: function(){ this.result(this.info) },           
	result: function(r){ $id(this.div).innerHTML= r; }
}
function fsAjax(url, div, info, method, formID, script,func){
	var requestContent = new ajax(url, div, info, method, formID, script,func);
	requestContent.connect();
}
/* END OF AJAX */ 

// =======================================================

// BELLOW HERE WILL BE IN SEPARETED FILES IN FUTURE

// FORM VALIDATION
var validation = {
	Email: function(mail){
		var er = new RegExp(/^[A-Za-z0-9_\-\.]+@[A-Za-z0-9_\-\.]{2,}\.[A-Za-z0-9]{2,}(\.[A-Za-z0-9])?/);
		if(typeof(mail) == "string"){
			if(er.test(mail)) return true;
		}else if(typeof(mail) == "object"){
			if(er.test(mail.value)){
				return true;
			}
		}else return false;
	}
}
// EFFECTS

var effects = {
	colorMask: function(color, obj, objOpacity){
		if(!$id("fsMask")) $obj.clone("imagem","fsMask","div");
		$id("fsMask").style.background=color;
		$obj.opacity("fsMask",objOpacity);		
	},
	fade: function(color, obj, from, to, func, nowOpacity){
		from = (!from)?0:from;
		to = (!to)?0:to;
		color = (!color)?null:color=="null"?null:color;
		
		nowOpacity = (!nowOpacity)?from:nowOpacity;
		
		if(!color) $obj.opacity(obj,nowOpacity);
		else this.colorMask(color,obj,nowOpacity);
		if(from>=to){			
			if(nowOpacity>to){
				nowOpacity-=10;
				setTimeout("effects.fade('"+color+"','"+obj+"',"+from+","+to+",'"+func+"','"+nowOpacity+"')",5);
			} else if(func) eval(func);
		} else {
			if(nowOpacity<to){
				nowOpacity+=10;
				setTimeout("effects.fade('"+color+"','"+obj+"',"+from+","+to+",'"+func+"',"+nowOpacity+")",5);
			}else if(func) eval(func);
		}
	}
}
