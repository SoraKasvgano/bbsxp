var sCurrMode
StyleEditorHeader = "<script>i=0;function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13 && i==0){i=1;parent.document.yuziform.content.value=document.body.innerHTML;parent.document.yuziform.submit();parent.document.yuziform.EditSubmit.disabled=true;}}<\/script><body onkeydown=ctlent()>" ;


//�������������ѡ��༭��
if (navigator.appName == "Microsoft Internet Explorer"){
HtmlEdit()
var IframeID=frames["HtmlEditor"];
if (navigator.appVersion.indexOf("MSIE 6.0",0)==-1){IframeID.document.designMode="On"}
IframeID.document.open();
IframeID.document.write (StyleEditorHeader);
IframeID.document.close();
IframeID.document.body.contentEditable = "True";
IframeID.document.body.innerHTML=document.getElementById("content").value;
}else{
document.write ('<iframe ID=HtmlEditor MARGINHEIGHT=5 MARGINWIDTH=5 width=100% height=100%></iframe>')
HtmlEditor=document.getElementById("HtmlEditor").contentWindow;
HtmlEditor.document.designMode="On"
HtmlEditor.document.open();
HtmlEditor.document.close();
HtmlEditor.document.body.contentEditable = "True";
}


function HtmlEdit(){

YBBCOLOR="<option style=COLOR:#ffffff;BACKGROUND-COLOR:black value=#000000>Black</option> <option style=BACKGROUND-COLOR:gray value=#808080>Gray</option> <option style=BACKGROUND-COLOR:darkgray value=#A9A9A9>DarkGray</option> <option style=BACKGROUND-COLOR:lightgrey value=#D3D3D3>LightGray</option> <option style=BACKGROUND-COLOR:white value=#FFFFFF>White</option> <option style=BACKGROUND-COLOR:aquamarine value=#7FFFD4>Aquamarine</option> <option style=BACKGROUND-COLOR:blue value=#0000FF>Blue</option> <option style=COLOR:#ffffff;BACKGROUND-COLOR:navy value=#000080>Navy</option> <option style=COLOR:#ffffff;BACKGROUND-COLOR:purple value=#800080>Purple</option> <option style=BACKGROUND-COLOR:deeppink value=#FF1493>DeepPink</option> <option style=BACKGROUND-COLOR:violet value=#EE82EE>Violet</option> <option style=BACKGROUND-COLOR:pink value=#FFC0CB>Pink</option> <option style=COLOR:#ffffff;BACKGROUND-COLOR:darkgreen value=#006400>DarkGreen</option> <option style=COLOR:#ffffff;BACKGROUND-COLOR:green value=#008000>Green</option> <option style=BACKGROUND-COLOR:yellowgreen value=#9ACD32>YellowGreen</option> <option style=BACKGROUND-COLOR:yellow value=#FFFF00>Yellow</option> <option style=BACKGROUND-COLOR:orange value=#FFA500>Orange</option> <option style=BACKGROUND-COLOR:red value=#FF0000>Red</option> <option style=BACKGROUND-COLOR:brown value=#A52A2A>Brown</option> <option style=BACKGROUND-COLOR:burlywood value=#DEB887>BurlyWood</option> <option style=BACKGROUND-COLOR:beige value=#F5F5DC>Beige</option>"
YBBFont="<option value='����'>����</option> <option value='����'>����</option> <option value='����_GB2312'>����_GB2312</option> <option value='����_GB2312'>����_GB2312</option> <option value='����'>����</option> <option value='��Բ'>��Բ</option> <option value='������'>������</option><option value='Arial'>Arial</option><option value='Courier New'>Courier New</option><option value='Garamond'>Garamond</option><option value='Georgia'>Georgia</option><option value='Tahoma'>Tahoma</option><option value='Times New Roman'>Times</option><option value='Verdana'>Verdana</option>"

document.writeln("<table width=100% height=100% cellpadding=2 cellspacing=0><tr id=EditButton style=DISPLAY:block><td align=left colspan=6>");
document.writeln("<select onChange=\"FormatText('FormatBlock',this[this.selectedIndex].value);\"><option>����</option><option value=\"&lt;P&gt;\">����</option><option value=\"&lt;H1&gt;\">����һ</option><option value=\"&lt;H2&gt;\">�����</option><option value=\"&lt;H3&gt;\">������</option><option value=\"&lt;H4&gt;\">������</option><option value=\"&lt;H5&gt;\">������</option><option value=\"&lt;H6&gt;\">������</option><option value=\"&lt;PRE&gt;\">Ԥ���ʽ</option></select>");
document.writeln("<select name=\"selectFont\" onChange=\"FormatText('fontname', selectFont.options[selectFont.selectedIndex].value);\"><option>����</option>"+YBBFont+"</select> <select onChange=\"FormatText('fontsize',this[this.selectedIndex].value);\" name=\"D2\"><option>��С</option><option value=1>1</option><option value=2>2</option><option value=3>3</option><option value=4>4</option><option value=5>5</option><option value=6>6</option><option value=7>7</option></select>");
document.writeln("<select onChange=\"FormatText('forecolor',this[this.selectedIndex].value);\"><option>������ɫ</option>"+YBBCOLOR+"</select> <select onChange=\"FormatText('BackColor',this[this.selectedIndex].value);\"><option>���ֱ�����ɫ</option>"+YBBCOLOR+"</select>");
document.writeln("<img src=images/ybb/replace.gif alt=\"�滻\" align=absmiddle style=cursor:hand onClick=replace()>");
document.writeln("<img src=images/ybb/img.gif alt=\"����ͼƬ\" align=absmiddle style=cursor:hand onClick=img()>");
document.writeln("<img alt=\"���� FLASH��MediaPlayer��RealPlayer �ļ�\" src=images/ybb/mp.gif align=absmiddle style=cursor:hand onclick=MediaPlayer()>");
document.writeln("<img alt=\"�������\" src=images/ybb/em.gif align=absmiddle style=cursor:hand onclick=em()>");
document.writeln("<img alt=\"����HTML����\" src=images/ybb/CleanCode.gif align=absmiddle style=cursor:hand onclick=CleanCode()><br>");

var FormatTextlist="���� bold|��б italic|�»��� underline|�ϱ� superscript|�±� subscript|ɾ���� strikethrough|ɾ�����ָ�ʽ RemoveFormat|����� Justifyleft|���� JustifyCenter|�Ҷ��� JustifyRight|���˶��� justifyfull|��� insertorderedlist|��Ŀ���� InsertUnorderedList|���������� Outdent|���������� indent|��ͨˮƽ�� InsertHorizontalRule|���� cut|���� copy|ճ�� paste|���� undo|�ָ� redo|ȫѡ selectAll|ȡ��ѡ�� unselect|ɾ����ǰѡ���� Delete|���볬���� createLink|ȥ�������� Unlink"
var list= FormatTextlist.split ('|'); 
for(i=0;i<list.length;i++) {
var TextName= list[i].split (' '); 
document.write(" <img align=absmiddle src=images/ybb/"+TextName[1]+".gif alt="+TextName[0]+" style=cursor:hand onClick=FormatText('"+TextName[1]+"')> ");
}
document.writeln("</td></tr><tr><td height=100% colspan=6><iframe ID=HtmlEditor MARGINHEIGHT=5 MARGINWIDTH=5 width=100% height=100%></iframe></td></tr>");
document.writeln("<tr><td></div><table border=0 cellpadding=0 cellspacing=0 height=20>");
document.writeln("<tr>");
document.writeln("<td width=5></td>");
document.writeln("<td class=StatusBarBtnON id=HtmlEditor_EDIT onclick=setMode('EDIT')><img border=0 src=images/ybb/modeedit.gif align=absmiddle></td>");
document.writeln("<td width=5></td>");
document.writeln("<td class=StatusBarBtnOff id=HtmlEditor_CODE onclick=setMode('CODE')><img border=0 src=images/ybb/modecode.gif align=absmiddle></td>");
document.writeln("<td width=5></td>");
document.writeln("<td class=StatusBarBtnOff id=HtmlEditor_VIEW onclick=setMode('VIEW')><img border=0 src=images/ybb/modepreview.gif width=50 height=15 align=absmiddle></td>");
document.writeln("<td width=5></td>");
document.writeln("<td align=right width=100%><img border=0 src=images/ybb/+.gif align=absmiddle onclick=sizeChange(100) style=cursor:pointer alt='����༭��'><img border=0 src=images/ybb/-.gif align=absmiddle onclick=sizeChange(-100) style=cursor:pointer alt='��С�༭��'></td>");
document.writeln("</tr></table>");
document.writeln("</td></tr></table>");

}


// �����༭���Ĵ�С
function sizeChange(size){
	var obj=document.getElementById("HtmlEditor");
	obj.height = (parseInt(obj.offsetHeight) + size);
	if (parseInt(obj.offsetHeight)< 200 ){obj.height ="100%"}
}

// ���ñ༭��������
function setHTML(html) {
	IframeID.document.body.innerHTML = html;
	switch (sCurrMode){
	case "CODE":
		HtmlEditor.document.open();
		HtmlEditor.document.write(html);
		HtmlEditor.document.close();
		HtmlEditor.document.body.innerText=html;
		HtmlEditor.document.body.contentEditable="true";
		EditButton.style.display = 'none';
                break;
	case "EDIT":
		HtmlEditor.document.open();
		HtmlEditor.document.write(StyleEditorHeader+html);
		HtmlEditor.document.close();
		HtmlEditor.document.body.contentEditable="true";
		EditButton.style.display = 'block';
                break;
	case "VIEW":
		HtmlEditor.document.open();
		HtmlEditor.document.write(StyleEditorHeader+html);
		HtmlEditor.document.close();
		HtmlEditor.document.body.contentEditable="false";
		EditButton.style.display = 'none';
		break;

	}
}

// �ı�ģʽ�����롢�༭��Ԥ��
function setMode(NewMode){
	if (NewMode!=sCurrMode){
		switch(sCurrMode){
		case "CODE":
			sBody = HtmlEditor.document.body.innerText;
			 break;
		case "EDIT":
		     sBody = HtmlEditor.document.body.innerHTML;
		      break;
		case "VIEW":
		     sBody = HtmlEditor.document.body.innerHTML;
		     break;
		default:
			sBody = IframeID.document.body.innerHTML;
			break;
		}
		// ��ͼƬ
		try{
			document.all["HtmlEditor_CODE"].className = "StatusBarBtnOff";
			document.all["HtmlEditor_EDIT"].className = "StatusBarBtnOff";
			document.all["HtmlEditor_VIEW"].className = "StatusBarBtnOff";
			document.all["HtmlEditor_"+NewMode].className = "StatusBarBtnOn";
			}
		catch(e){
			}
		sCurrMode = NewMode;
		setHTML(sBody);
		
	}
}
// ȡ�༭��������
function getHTML() {
	var html;
	if(sCurrMode=="CODE"){
		html = HtmlEditor.document.body.innerText;
	}else{
		html = HtmlEditor.document.body.innerHTML;
	}
	return html;
}


function em(){
var arr = showModalDialog("inc/Emotion.htm", "", "dialogWidth:20em; dialogHeight:9.5em; status:0;help:0");
if (arr != null){
IframeID.focus()
sel=IframeID.document.selection.createRange();
sel.pasteHTML(arr);
}
}

function FormatText(command,option){IframeID.focus();IframeID.document.execCommand(command,true,option);}

function CheckLength(){alert("����ַ�Ϊ "+60000+ " �ֽ�\n������������ "+IframeID.document.body.innerHTML.length+" �ֽ�");}

function emoticon(theSmilie){
IframeID.focus();
sel=IframeID.document.selection.createRange();
sel.pasteHTML("<img src=images/Emotions/"+theSmilie+".gif>");
}


function CheckForm(form){
if(document.yuziform.Subject.value.length<2){alert("���ⲻ��С��2���ַ���");return false;}
form.content.value = getHTML();
MessageLength=IframeID.document.body.innerHTML.length;
if(MessageLength<2){alert("���ݲ���С��2���ַ���");return false;}
if(MessageLength>60000){alert("���ݲ��ܳ���60000���ַ���");return false;}
document.yuziform.EditSubmit.disabled = true;
}


//////�滻����
function replace()
{
  var arr = showModalDialog("inc/Replace.htm", "", "dialogWidth:22em;dialogHeight:10em;status:0;help:0");
	if (arr != null){
		var ss;
		ss = arr.split("*")
		a = ss[0];
		b = ss[1];
		i = ss[2];
		con = IframeID.document.body.innerHTML;
		if (i == 1)
		{
			con = bbsxp_rCode(con,a,b,true);
		}else{
			con = bbsxp_rCode(con,a,b);
		}
		IframeID.document.body.innerHTML = con;
	}
	else IframeID.focus();
}
function bbsxp_rCode(s,a,b,i){
	a = a.replace("?","\\?");
	if (i==null)
	{
		var r = new RegExp(a,"gi");
	}else if (i) {
		var r = new RegExp(a,"g");
	}
	else{
		var r = new RegExp(a,"gi");
	}
	return s.replace(r,b); 
}
//////�滻���ݽ���

function img(){
url=prompt("������ͼƬ�ļ���ַ:","http://");
if(!url || url=="http://") return;
IframeID.focus();
sel=IframeID.document.selection.createRange();
sel.pasteHTML("<img src="+url+">");
}

function MediaPlayer(){
var arr = showModalDialog("inc/MediaPlayer.htm", "", "dialogWidth:22em; dialogHeight:10.5em; status:0;help:0");
if (arr != null){
IframeID.focus()
sel=IframeID.document.selection.createRange();
sel.pasteHTML(arr);
}
}

function CleanCode(){
var body = IframeID.document.body;
var html = IframeID.document.body.innerHTML;
html = html.replace(/\<p>/gi,"[$p]");
html = html.replace(/\<\/p>/gi,"[$\/p]");
html = html.replace(/\<br>/gi,"[$br]");
html = html.replace(/\<[^>]*>/g,"");
html = html.replace(/\[\$p\]/gi,"<p>");
html = html.replace(/\[\$\/p\]/gi,"<\/p>");
html = html.replace(/\[\$br\]/gi,"<br>");
IframeID.document.body.innerHTML = html;
}