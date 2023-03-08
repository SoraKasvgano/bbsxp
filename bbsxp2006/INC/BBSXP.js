newDate=new Date()
newDate=""+newDate.getYear()+"-"+[newDate.getMonth()+1]+"-"+newDate.getDate()+""

//读取COOKIE
function getCookie (CookieName) { 
var CookieString = document.cookie; 
var CookieSet = CookieString.split (';'); 
var SetSize = CookieSet.length; 
var CookiePieces 
var ReturnValue = ""; 
var x = 0; 
for (x = 0; ((x < SetSize) && (ReturnValue == "")); x++) { 
CookiePieces = CookieSet[x].split ('='); 

if (CookiePieces[0].substring (0,1) == ' ') { 
CookiePieces[0] = CookiePieces[0].substring (1, CookiePieces[0].length); 
}

if (CookiePieces[0] == CookieName) {
ReturnValue = CookiePieces[1];
var value =ReturnValue
}


}
return value;
}


//跳出确认
function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}


//跳转页面显示
function ShowPage(TotalPage,PageIndex,url){
document.write("<table cellspacing=1 cellpadding=2 class=a2><tr class=a3><td class=a1><b>"+PageIndex+"/"+TotalPage+"</b></td>");
if (PageIndex<6){PageLong=11-PageIndex;}
else
if (TotalPage-PageIndex<6){PageLong=10-(TotalPage-PageIndex)}
else{PageLong=5;}
for (var i=1; i <= TotalPage; i++) {
if (i < PageIndex+PageLong && i > PageIndex-PageLong || i==1 || i==TotalPage){
if (PageIndex==i){document.write("<td>&nbsp;"+ i +"&nbsp;</td>");}else{document.write("<td>&nbsp;<a href=?PageIndex="+i+"&"+url+">"+ i +"</a>&nbsp;</td>");}
}
}
document.write("<td class=a4><input onkeydown=if((event.keyCode==13)&&(this.value!=''))window.location='?PageIndex='+this.value+'&"+url+"'; onkeyup=if(isNaN(this.value))this.value='' style='border:1px solid #698cc3;' size=2></td></tr></table>");
}

//全选复选框
function CheckAll(form){for (var i=0;i<form.elements.length;i++){var e = form.elements[i];if (e.name != 'chkall')e.checked = form.chkall.checked;}}

//全选主题ID复选框
function ThreadIDCheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.name == 'ThreadID')e.checked = form.chkall.checked;
}
}

//菜单
var menuOffX=0	//菜单距连接文字最左端距离
var menuOffY=20	//菜单距连接文字顶端距离

var ie4=document.all&&navigator.userAgent.indexOf("Opera")==-1
var ns6=document.getElementById&&!document.all
function showmenu(e,vmenu,mod){
	which=vmenu
	menuobj=document.getElementById("popmenu")
	menuobj.thestyle=menuobj.style
	menuobj.innerHTML=which
	menuobj.contentwidth=menuobj.offsetWidth
	eventX=e.clientX
	eventY=e.clientY
	var rightedge=document.body.clientWidth-eventX
	var bottomedge=document.body.clientHeight-eventY

		if (rightedge<menuobj.contentwidth)
			menuobj.thestyle.left=document.body.scrollLeft+eventX-menuobj.contentwidth+menuOffX
		else
			menuobj.thestyle.left=ie4? ie_x(event.srcElement)+menuOffX : ns6? window.pageXOffset+eventX : eventX
		
		if (bottomedge<menuobj.contentheight&&mod!=0)
			menuobj.thestyle.top=document.body.scrollTop+eventY-menuobj.contentheight-event.offsetY+menuOffY-23
		else
			menuobj.thestyle.top=ie4? ie_y(event.srcElement)+menuOffY : ns6? window.pageYOffset+eventY+10 : eventY

	menuobj.thestyle.visibility="visible"
}


function ie_y(e){  
	var t=e.offsetTop;  
	while(e=e.offsetParent){  
		t+=e.offsetTop;  
	}  
	return t;  
}  
function ie_x(e){  
	var l=e.offsetLeft;  
	while(e=e.offsetParent){  
		l+=e.offsetLeft;  
	}  
	return l;  
}

function highlightmenu(e,state){
	if (document.all)
		source_el=event.srcElement
		while(source_el.id!="popmenu"){
			source_el=document.getElementById? source_el.parentNode : source_el.parentElement
			if (source_el.className=="menuitems"){
				source_el.id=(state=="on")? "mouseoverstyle" : ""
		}
	}
}


function hidemenu(){if (window.menuobj)menuobj.thestyle.visibility="hidden"}
function dynamichide(e){if ((ie4||ns6)&&!menuobj.contains(e.toElement))hidemenu()}
document.onclick=hidemenu
document.write("<div class=menuskin id=popmenu onmouseover=highlightmenu(event,'on') onmouseout=highlightmenu(event,'off');dynamichide(event)></div>")
// 菜单END

// add area script
function focusEdit(editBox)
{
 if ( editBox.value == editBox.Helptext )
 {
 editBox.value = '';
 editBox.className = 'editbox';
 }
 return true;
}
function blurEdit(editBox)
{
 if ( editBox.value.length == 0 )
 {
 editBox.className = 'editbox Graytitle';
 editBox.value = editBox.Helptext;
 }
}
function ValidateTextboxAdd(box, button)
{
 var buttonCtrl = document.getElementById( button );
 if ( buttonCtrl != null )
 {
 if (box.value == "" || box.value == box.Helptext)
 {
 buttonCtrl.disabled = true;
 }
 else
 {
 buttonCtrl.disabled = false;
 }
 }
}
// add area script end


function loadtree(ino){
document.frames["hiddenframe"].location.replace("ForumTree.asp?id="+ino+"")
}


function loadThreadFollow(ino,Online){
var targetImg =document.getElementById("followImg" + ino);
var targetDiv =document.getElementById("follow" + ino);
if (targetDiv.style.display!='block'){
if(targetImg.loaded=="no"){document.frames["hiddenframe"].location.replace("loading.asp?id="+ino+"&ForumID="+Online+"");}
targetDiv.style.display="block";
targetImg.src="images/minus.gif";
}else{
targetDiv.style.display="none";
targetImg.src="images/plus.gif";
}
}


function ToggleMenuOnOff (menuName) {
    var menu = document.getElementById(menuName);

    if (menu.style.display == 'none') {
      menu.style.display = 'block';
    } else {
      menu.style.display = 'none';
    }

}