document.write("<div id=PopUpMenu></div>")

var IsIE = ((UserAgent.indexOf("msie") != -1) && (UserAgent.indexOf("opera") == -1));

function MouseOverOpen(idName,object){
	obj  = $("PopUpMenu");
	obj2 = $(idName);
	obj3 = $(object);

	var left  = ie_x(obj3);
	var top   = ie_y(obj3);
		
	obj.innerHTML = obj2.innerHTML;
	obj.className = obj2.className;
	obj.style.filter = "alpha(opacity=96);opacity:0.96;";
	obj.style.display = "";
		
	var height = obj.offsetHeight;
	var width  = obj.offsetWidth;
		

	top +=  24;

	obj.style.top	= top + 'px';
	obj.style.left	= left + 'px';

	addEvent(document,"mouseout",DocumentMouseOut);
	return false;
}

function CloseMenu(){
	obj = $("PopUpMenu");
	obj.innerHTML = "";
	obj.className = "";
	obj.style.display = "none";
	removeEvent(document,"mousedown",DocumentMouseDown);
	return false;
}

function DocumentMouseDown(e){
	var e = IsIE ? event: e;
	obj	= $("PopUpMenu");
	_x	= IsIE ? e.x : e.pageX;
	_y	= IsIE ? e.y + IeTrueBody().scrollTop : e.pageY;
	_x1 = obj.offsetLeft;
	_x2 = obj.offsetLeft + obj.offsetWidth;
	_y1 = obj.offsetTop - 25;
	_y2 = obj.offsetTop + obj.offsetHeight;
	if(_x<_x1 || _x>_x2 || _y<_y1 || _y>_y2){
		CloseMenu();
	}
}

function DocumentMouseOut(e){
	var e = IsIE ? event: e;
	obj	= $("PopUpMenu");
	_x	= IsIE ? e.x : e.pageX;
	_y	= IsIE ? e.y + IeTrueBody().scrollTop : e.pageY;
	_x1 = obj.offsetLeft - 20;
	_x2 = obj.offsetLeft + obj.offsetWidth + 20;
	_y1 = obj.offsetTop - 35;
	_y2 = obj.offsetTop + obj.offsetHeight + 20;
	if(_x<_x1 || _x>_x2 || _y<_y1 || _y>_y2){
		CloseMenu('out');
	}
}

function IeTrueBody(){
	return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body;
}

function addEvent(el,evname,func){
	IsIE ? el.attachEvent("on" + evname,func) : el.addEventListener(evname,func,true);
};

function removeEvent(el,evname,func){
	IsIE ? el.detachEvent("on" + evname,func) : el.removeEventListener(evname,func,true);
};
