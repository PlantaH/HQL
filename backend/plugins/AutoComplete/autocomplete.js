var useBSNns = true;
if (useBSNns){
	if (typeof(bsn) == "undefined"){
		bsn = {}
		_bsn = bsn;
	}
}else{
	_bsn = this;
}
if (typeof(_bsn.Autosuggest) == "undefined"){
	_bsn.Autosuggest = {}
}
_bsn.AutoSuggest = function (fldID, param){
	if (!document.getElementById) return false;
	this.fld = $(fldID);
	if (!this.fld) return false;
	this.sInput 		= "";
	this.nInputChars 	= 0;
	this.aSuggestions 	= [];
	this.iHighlighted 	= 0;
	this.oP = (param) ? param : {};
	if (!this.oP.minchars) this.oP.minchars = 3;
	if (!this.oP.method) this.oP.meth = "get";
	if (!this.oP.varname) this.oP.varname = "input";
	if (!this.oP.className) this.oP.className = "autosuggest";
	if (!this.oP.timeout) this.oP.timeout = 5000;
	if (!this.oP.delay) this.oP.delay = 300;
	if (!this.oP.offsety) this.oP.offsety = -5;
	if (!this.oP.shownoresults) this.oP.shownoresults = true;
	if (!this.oP.noresults) this.oP.noresults = "No se encontraron resultados";
	if (!this.oP.maxheight && this.oP.maxheight !== 0) this.oP.maxheight = 250;
	if (!this.oP.cache && this.oP.cache != false) this.oP.cache = false;
	var pointer = this;
	this.fld.addEvents({
		'keyup' : function(ev){		
			return pointer.onKeyUp(ev);
		},
		'keypress' : function(ev){
			return pointer.onKeyPress(ev);
		}
	}).setProperty("autocomplete","off");
	if(document.all){
		this.fld["onpaste"] = new Function("return false");
		this.fld["oncontextmenu"] = new Function("return false");
	}else{
		with(this.fld){
			setAttribute("onpaste","return false");
			setAttribute("oncontextmenu","return false");
		}	
	}	
}
var _imgBgLoadingAC = 'js/autocomplete/img/bg_loading.gif';
var _imgBgLoadingACcache = new Image();
_imgBgLoadingACcache.src = _imgBgLoadingAC;
var _globalBgLoadingAC = '#ffffff url('+_imgBgLoadingAC+') no-repeat right center';
var _globalBgLoadedAC = '#ffffff';

_bsn.AutoSuggest.prototype.onKeyPress = function(ev){
	var bubble = true;
	switch(ev.key){
		case 'enter':
			this.setHighlightedValue();
			bubble = false;
			break;
		case 'esc':
			this.clearSuggestions();
			break;			
		case 'tab':
			this.clearSuggestions();
			break;
	}	
	if(!window.event){
		if(ev.ctrlKey) bubble = false;
	}
	return bubble;
}

_bsn.AutoSuggest.prototype.onKeyUp = function(ev){	
	var bubble = true;
	switch(ev.key){
		case 'up':
			this.changeHighlight(ev.key);
			bubble = false;
			break;

		case 'down':
			this.changeHighlight(ev.key);
			bubble = false;
			break;		
		
		default:
			this.getSuggestions(this.fld.value);
	}
	return bubble;
}

_bsn.AutoSuggest.prototype.getSuggestions = function (val){
	if(val.length < 1 && typeof(this.oP.callback) == "function") this.oP.callback('');
	if (val == this.sInput) return false;
	if (val.length < this.oP.minchars){	
		this.sInput = "";
		return false;
	}
	if (val.length>this.nInputChars && this.aSuggestions.length && this.oP.cache){
		var arr = [];
		for (var i=0;i<this.aSuggestions.length;i++){
			if (this.aSuggestions[i].value.substr(0,val.length).toLowerCase() == val.toLowerCase()) arr.push( this.aSuggestions[i] );
		}		
		this.sInput = val;
		this.nInputChars = val.length;
		this.aSuggestions = arr;		
		this.createList(this.aSuggestions);		
		return false;
	}else{
		this.sInput = val;
		this.nInputChars = val.length;
		var pointer = this;
		clearTimeout(this.ajID);
		this.ajID = setTimeout( function() { pointer.doAjaxRequest() }, this.oP.delay );
	}
	return false;
}

_bsn.AutoSuggest.prototype.doAjaxRequest = function (r){	
	var pointer = this;
	if(typeof(this.oP.script) == "function"){
		var _sc = this.oP.script();
	}else{
		var _sc = this.oP.script;
	}	
	var url = _sc+this.oP.varname+"="+encodeURIComponent(this.fld.value);
	if(r) url += "&zonas=1"	
	var meth = this.oP.meth;	
	var onSuccessFunc = function (req) { pointer.setSuggestions(req, r) };
	var onErrorFunc = function (status) { alert("error: "+status); };
	var myAjax = new _bsn.Ajax();
	myAjax.makeRequest( url, meth, onSuccessFunc, onErrorFunc, this.fld );
}
_bsn.AutoSuggest.prototype.setSuggestions = function (req, r){
	this.aSuggestions = [];		
	if (this.oP.json){
		var jsondata = eval('(' + req.responseText + ')');		
		for (var i=0;i<jsondata.results.length;i++){
			this.aSuggestions.push(  { 'id':jsondata.results[i].id, 'value':jsondata.results[i].value, 'info':jsondata.results[i].info, 'host':jsondata.results[i].host }  );
		}
	}else{
		var xml = req.responseXML;
		var results = xml.getElementsByTagName('results')[0].childNodes;
		for (var i=0;i<results.length;i++){
			if (results[i].hasChildNodes()) this.aSuggestions.push(  { 'id':results[i].getAttribute('id'), 'value':results[i].childNodes[0].nodeValue, 'info':results[i].getAttribute('info'), 'host':results[i].getAttribute('host') }  );
		}	
	}
	this.idAs = "as_"+this.fld.id;
	this.createList(this.aSuggestions, r);
}

_bsn.AutoSuggest.prototype.createList = function(arr, r){
	var pointer = this;
	_bsn.DOM.removeElement(this.idAs);
	this.killTimeout();
	var div = _bsn.DOM.createElement("div", {id:this.idAs, className:this.oP.className});	
	var divScroll = _bsn.DOM.createElement("div", {id:'divscroll', className:'divscroll'});
	var ul = _bsn.DOM.createElement("ul", {id:"as_ul"});
	for (var i=0;i<arr.length;i++){
		var val = arr[i].value;
		var st = val.toLowerCase().indexOf( this.sInput.toLowerCase() );
		var output = val.substring(0,st) + "<em>" + val.substring(st, st+this.sInput.length) + "</em>" + val.substring(st+this.sInput.length);
		var span = _bsn.DOM.createElement("span", {}, output, true);
		if (arr[i].info != ""){
			var br = _bsn.DOM.createElement("br", {});
			span.appendChild(br);
			var small = _bsn.DOM.createElement("small", {}, arr[i].info, true);
			span.appendChild(small);
		}		
		var a = _bsn.DOM.createElement("a", { href:"#" });
		a.appendChild(span);
		a.name = i+1;		
		a.rel = arr[i].id;		
		a.onclick = function () {
			if(this.rel == "#"){
				pointer.doAjaxRequest(true);				
				return false;
			}else{
				pointer.setHighlightedValue();
				return false;
			}			
		}
		a.onmouseover = function () { pointer.setHighlight(this.name); }				
		var li = _bsn.DOM.createElement(  "li", {}, a  );		
		ul.appendChild( li );
	}
	if (arr.length == 0){
		var li = _bsn.DOM.createElement(  "li", {className:"as_warning"}, this.oP.noresults  );		
		ul.appendChild( li );
	}	
	divScroll.appendChild( ul );
	div.appendChild( divScroll );
	div.style.left = this.fld.getCoordinates().left + "px";
	div.style.top =  this.fld.getCoordinates().bottom-5 + "px";
	div.style.width =  this.fld.offsetWidth + "px";
	document.getElementsByTagName("body")[0].appendChild(div);	
	if (ul.offsetHeight > this.oP.maxheight && this.oP.maxheight != 0){
		if(Browser.Engine.trident) ul.style.width = (this.fld.offsetWidth - 17) + "px";		
		divScroll.style['height'] = this.oP.maxheight + 'px';
		divScroll.style['overflow'] = 'auto';
	}
	this.iHighlighted = 0;
	var pointer = this;
	this.fld.addEvent('blur',function(){
		if(!r){
		 //pointer.killTimeout();
		 //pointer.toID = setTimeout(function(){pointer.clearSuggestions()}, 100)
		}
	});	
	if(Browser.Engine.trident4) this.togSelect(1);
}

_bsn.AutoSuggest.prototype.changeHighlight = function(key){		
	var list = _bsn.DOM.getElement("as_ul");
	if (!list) return false;	
	var n;
	if (key == 'down'){
		n = this.iHighlighted + 1;
	}else if (key == 'up'){
		n = this.iHighlighted - 1;
	}	
	if (n > list.childNodes.length)	n = list.childNodes.length;
	if (n < 1) n = 1;	
	this.setHighlight(n);
}

_bsn.AutoSuggest.prototype.setHighlight = function(n){
	var list = _bsn.DOM.getElement("as_ul");
	if (!list) return false;	
	if (this.iHighlighted > 0) this.clearHighlight();	
	this.iHighlighted = Number(n);	
	list.childNodes[this.iHighlighted-1].className = "as_highlight";
	this.killTimeout();
}

_bsn.AutoSuggest.prototype.clearHighlight = function(){
	var list = _bsn.DOM.getElement("as_ul");
	if (!list)	return false;
	
	if (this.iHighlighted > 0){
		list.childNodes[this.iHighlighted-1].className = "";
		this.iHighlighted = 0;
	}
}

_bsn.AutoSuggest.prototype.setHighlightedValue = function (){
	if (this.iHighlighted){
		this.sInput = this.fld.value = this.aSuggestions[ this.iHighlighted-1 ].value;
		this.fld.focus();
		if (this.fld.selectionStart) this.fld.setSelectionRange(this.sInput.length, this.sInput.length);
		this.clearSuggestions();
		if (typeof(this.oP.callback) == "function") this.oP.callback( this.aSuggestions[this.iHighlighted-1] );
	}
}

_bsn.AutoSuggest.prototype.killTimeout = function(){
	clearTimeout(this.toID);
}

_bsn.AutoSuggest.prototype.resetTimeout = function(){
	clearTimeout(this.toID);
	var pointer = this;
	this.toID = setTimeout(function () { pointer.clearSuggestions() }, 1000);
}
_bsn.AutoSuggest.prototype.clearSuggestions = function (){	
	this.killTimeout();	
	var ele = _bsn.DOM.getElement(this.idAs);
	var pointer = this;
	if (ele){
		var fade = new _bsn.Fader(ele,1,0,150,function () { _bsn.DOM.removeElement(pointer.idAs); if(Browser.Engine.trident4) pointer.togSelect(2); });		
	}
}

_bsn.AutoSuggest.cerrado = true;
_bsn.AutoSuggest.prototype.togSelect = function(t) {
	this.vEsO = $$('select');
	if(t == 1 && _bsn.AutoSuggest.cerrado){
		this.vLogEsS = [];
		this.vEsO.each(function(o,i){
			vLogEsS.push(o.getStyle('visibility'));		
			o.setStyle('visibility', 'hidden');
		});	
		_bsn.AutoSuggest.cerrado = false;
	}else if(t == 2){
		this.vEsO.each(function(o,i){		
			o.setStyle('visibility', this.vLogEsS[i]);
		});
		_bsn.AutoSuggest.cerrado = true;
	}		
}.bind(this);

if (typeof(_bsn.Ajax) == "undefined") _bsn.Ajax = {}

_bsn.Ajax = function (){
	this.req = {};
	this.isIE = false;
}

_bsn.Ajax.prototype.makeRequest = function (url, meth, onComp, onErr, obj){	
	if (meth != "POST") meth = "GET";	
	this.onComplete = onComp;
	this.onError = onErr;	
	var pointer = this;
	if (window.XMLHttpRequest){
		this.req = new XMLHttpRequest();
		this.req.onreadystatechange = function () { pointer.processReqChange(obj) };
		this.req.open("GET", url, true);
		this.req.send(null);
	}else if (window.ActiveXObject){
		this.req = new ActiveXObject("Microsoft.XMLHTTP");
		if (this.req){
			this.req.onreadystatechange = function () { pointer.processReqChange(obj) };
			this.req.open(meth, url, true);
			this.req.send();
		}
	}
}

_bsn.Ajax.prototype.processReqChange = function(o){
	o.style.background = _globalBgLoadingAC;
	if (this.req.readyState == 4) {
		o.style.background = _globalBgLoadedAC;
		if (this.req.status == 200)
		{
			this.onComplete( this.req );
		} else {
			this.onError( this.req.status );
		}
	}
}

if (typeof(_bsn.DOM) == "undefined") _bsn.DOM = {}

_bsn.DOM.createElement = function ( type, attr, cont, html ){
	var ne = document.createElement( type );
	if (!ne) return false;		
	for (var a in attr){
		ne[a] = attr[a];
	}		
	if (typeof(cont) == "string" && !html){
		ne.appendChild( document.createTextNode(cont) );
	}else if (typeof(cont) == "string" && html){
		ne.innerHTML = cont;
	}else if (typeof(cont) == "object"){
		ne.appendChild( cont );
	}
	return ne;
}

_bsn.DOM.clearElement = function ( id ){
	var ele = this.getElement( id );	
	if (!ele) return false;	
	while (ele.childNodes.length){
		ele.removeChild( ele.childNodes[0] );
	}	
	return true;
}

_bsn.DOM.removeElement = function ( ele ){
	var e = this.getElement(ele);	
	if (!e){
		return false;
	}else if (e.parentNode.removeChild(e)){
		return true;
	}else{
		return false;
	}
}

_bsn.DOM.replaceContent = function ( id, cont, html ){
	var ele = this.getElement( id );	
	if (!ele) return false;	
	this.clearElement( ele );	
	if (typeof(cont) == "string" && !html){
		ele.appendChild( document.createTextNode(cont) );
	}else if (typeof(cont) == "string" && html){
		ele.innerHTML = cont;
	}else if (typeof(cont) == "object"){
		ele.appendChild( cont );
	}
}

_bsn.DOM.getElement = function ( ele ){
	if (typeof(ele) == "undefined"){
		return false;
	}else if (typeof(ele) == "string")	{
		var re = document.getElementById( ele );
		if (!re){
			return false;
		}else if (typeof(re.appendChild) != "undefined" ) {
			return re;
		} else {
			return false;
		}
	}else if (typeof(ele.appendChild) != "undefined"){
		return ele;
	}else{
		return false;
	}
}

_bsn.DOM.appendChildren = function ( id, arr ){
	var ele = this.getElement( id );	
	if (!ele) return false;	
	if (typeof(arr) != "object"){
		return false;
	}
	for (var i=0;i<arr.length;i++){
		var cont = arr[i];
		if (typeof(cont) == "string")
			ele.appendChild( document.createTextNode(cont) );
		else if (typeof(cont) == "object")
			ele.appendChild( cont );
	}
}
if (typeof(_bsn.Fader) == "undefined") _bsn.Fader = {}

_bsn.Fader = function (ele, from, to, fadetime, callback){	
	if (!ele) return false;
	
	this.ele = ele;
	
	this.from = from;
	this.to = to;
	
	this.callback = callback;
	
	this.nDur = fadetime;
		
	this.nInt = 50;
	this.nTime = 0;
	
	var p = this;
	this.nID = setInterval(function() { p._fade() }, this.nInt);
}

_bsn.Fader.prototype._fade = function(){
	this.nTime += this.nInt;
	
	var ieop = Math.round( this._tween(this.nTime, this.from, this.to, this.nDur) * 100 );
	var op = ieop / 100;
	
	if (this.ele.filters) // internet explorer
	{
		try
		{
			this.ele.filters.item("DXImageTransform.Microsoft.Alpha").opacity = ieop;
		} catch (e) { 
			// If it is not set initially, the browser will throw an error.  This will set it if it is not set yet.
			this.ele.style.filter = 'progid:DXImageTransform.Microsoft.Alpha(opacity='+ieop+')';
		}
	}
	else // other browsers
	{
		this.ele.style.opacity = op;
	}
	
	
	if (this.nTime == this.nDur)
	{
		clearInterval( this.nID );
		if (this.callback != undefined)
			this.callback();
	}
}



_bsn.Fader.prototype._tween = function(t,b,c,d)
{
	return b + ( (c-b) * (t/d) );
}

var ac_cssDOM = document.createElement('link');
with(ac_cssDOM){
	setAttribute('rel', 'stylesheet');
	setAttribute('type', 'text/css');
	setAttribute('href', 'js/autocomplete/autocomplete.css');
	setAttribute('media', 'screen');
}
document.getElementsByTagName("head")[0].appendChild(ac_cssDOM);

//
var AutoComplete = function(c,h,t){
	var _options = {
		script: "js/autocomplete/getxml.asp?tipo="+t+"&",
		varname: "input",
		callback: function(obj){
			$(h).set('value',obj.id);					
		}
	};
	var as_xml = new bsn.AutoSuggest(c, _options);
}

var AutoCompleteAdicional = function(c,h,t,id){
	var _options = {
		script: "js/autocomplete/getxml.asp?tipo="+t+"&id="+id+"&",
		varname: "input",
		callback: function(obj){
			$(h).set('value',obj.id);					
		}
	};
	var as_xml = new bsn.AutoSuggest(c, _options);
}