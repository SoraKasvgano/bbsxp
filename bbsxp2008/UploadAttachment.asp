<!-- #include file="Setup.asp" -->
<!--#include FILE="Utility/UpFile_Class.asp"-->
<%
if CookieUserName=empty then Alert("����δ��¼��̳")


Set Rs=Execute("select * from ["&TablePrefix&"Roles] where RoleID="&CookieUserRoleID&"")
	RoleMaxFileSize=Rs("RoleMaxFileSize")								'�û���ɫ����ϴ��ļ���С
	RoleMaxPostAttachmentsSize=Rs("RoleMaxPostAttachmentsSize")			'�û���ɫ����������
Rs.Close

UpMaxFileSize=SiteConfig("MaxFileSize")*1024							'����ϴ��ļ���С
UpMaxPostAttachmentsSize=SiteConfig("MaxPostAttachmentsSize")*1024		'����������
UpFileTypes=LCase(SiteConfig("UpFileTypes"))							'�����ϴ��ļ�����

if RoleMaxFileSize>0 then UpMaxFileSize = RoleMaxFileSize*1024
if RoleMaxPostAttachmentsSize>0 then UpMaxPostAttachmentsSize = RoleMaxPostAttachmentsSize*1024


if Request_Method = "POST" then
	'����������浽Ӳ��������Ŀ¼
	if SiteConfig("AttachmentsSaveOption")=1 then
		If IsObjInstalled("Scripting.FileSystemObject") Then
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			strDir="UpFile/UpAttachment/"&year(now)&"-"&month(now)&""
			if not fso.folderexists(Server.MapPath(strDir)) then fso.CreateFolder(Server.MapPath(strDir))
			Set fso = nothing
			UpFolder=""&year(now)&"-"&month(now)&"/"
		end if

		UpFolder="UpFile/UpAttachment/"&UpFolder&""	'�ϴ�·��
	end if

	AttachmentIDList=""
	FileNameList=""
	FileMIMEList=""
	SaveFileList=""
%>
<!--#include file="Utility/UpFile.asp"-->
<%
end if

%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<style>
body{font-size:9pt;font-family:verdana;margin:0;padding:0;background-color:#FAFDFF;}
a.addfile{background:url(images/addfile.gif) no-repeat;display:block;float:left;height:20px;margin-top:-1px;position:relative;text-decoration:none;top:0pt;width:80px;cursor:pointer;}
a:hover.addfile{background:url(images/addfile.gif) no-repeat;display:block;float:left;height:20px;margin-top:-1px;position:relative;text-decoration:none;top:0pt;width:80px;cursor:pointer;}
input.addfile{cursor:pointer;height:20px;left:-10px;position:absolute;top:0px;width:1px;filter:alpha(opacity=0);opacity:0;}
</style>
<head>
<script language="javascript">
<!--
var BbsxpFileInput={
	count:0,
	realcount:0,
	uped:0,//�����Ѿ��ϴ�����
	max:1,//�������ϴ����ٸ�
	once:1,//�����ͬʱ�ϴ����ٸ�
	readme:'',
	MaxFileSize:0,
	FileTypes:'',
	$:function(d){return document.getElementById(d);},
	ae:function(o,t,h){
		if (o.addEventListener){
			o.addEventListener(t,h,false);
		}
		else if(o.attachEvent){
			o.attachEvent('on'+t,h);
		}
		else{
			try{o['on'+t]=h;}catch(e){;}
		}
	},
	add:function(){
		if (BbsxpFileInput.chkre()){
			BbsxpFileInput_OnEcho('<font color=red><b>���Ѿ���ӹ����ļ���!</b></font>');
		}
		else if (BbsxpFileInput.realcount>=BbsxpFileInput.max){
			BbsxpFileInput_OnEcho('<font color=red><b>�����ֻ���ϴ�'+BbsxpFileInput.max+'���ļ���</b></font>');
		}else if (BbsxpFileInput.realcount>=BbsxpFileInput.once){
			BbsxpFileInput_OnEcho('<font color=red><b>��һ�����ֻ���ϴ�'+BbsxpFileInput.once+'���ļ���</b></font>');
		}else{
			BbsxpFileInput_OnEcho('<font color=blue>���Լ�����Ӹ�����Ҳ���������ϴ���</font>');
			var o=BbsxpFileInput.$('Bbsxp_fileinput_'+BbsxpFileInput.count);
			++BbsxpFileInput.count;
			++BbsxpFileInput.realcount;
			BbsxpFileInput_OnResize();
			var oInput=document.createElement('input');
			oInput.type='file';
			oInput.id='Bbsxp_fileinput_'+BbsxpFileInput.count;
			oInput.name='Bbsxp_fileinput_'+BbsxpFileInput.count;
			oInput.size=1;
			oInput.className='addfile';
			BbsxpFileInput.ae(oInput,'change',function(){BbsxpFileInput.add();});
			o.parentNode.appendChild(oInput);
			o.blur();
			o.style.display='none';
			BbsxpFileInput.show();
		}
	},
	chkre:function(){
		var c=BbsxpFileInput.$('Bbsxp_fileinput_'+BbsxpFileInput.count).value;
		for (var i=BbsxpFileInput.count-1; i>=0; --i){
			var o=BbsxpFileInput.$('Bbsxp_fileinput_'+i);
			if (o&&o.value==c&&BbsxpFileInput.realcount>0){return true}
		}
		return false;
	},
	filename:function(u){	//nouse
		var p=u.lastIndexOf('\\');
		return (p==-1?u:u.substr(p+1));
	},
	show:function(){
		var oDiv=document.createElement('div');
		var oBtn=document.createElement('img');
		var i=BbsxpFileInput.count-1;
		oBtn.id='Bbsxp_fileinput_btn_'+i;
        oBtn.src='images/filedel.gif';
        oBtn.alt='ɾ��';
		oBtn.style.cursor='pointer';
		var o=BbsxpFileInput.$('Bbsxp_fileinput_'+i);
		BbsxpFileInput.ae(oBtn,'click',function(){
			BbsxpFileInput.remove(i);
        });
        oDiv.innerHTML='<img src="images/fileitem.gif" width="13" height="11" border="0" /> <font color=gray>'+o.value+'</font> ';
        oDiv.appendChild(oBtn);
        BbsxpFileInput.$('Bbsxp_fileinput_show').appendChild(oDiv);
	},
	remove:function(i){
		var oa=BbsxpFileInput.$('Bbsxp_fileinput_'+i);
		var ob=BbsxpFileInput.$('Bbsxp_fileinput_btn_'+i);
		if(oa&&i>=0){oa.parentNode.removeChild(oa);}
		if(ob){ob.parentNode.parentNode.removeChild(ob.parentNode);}
//		if(0==i){BbsxpFileInput.$('Bbsxp_fileinput_0').disabled=true;}
		if(0==BbsxpFileInput.realcount){BbsxpFileInput.clear();}else{--BbsxpFileInput.realcount;}
		BbsxpFileInput_OnResize();
	},
	init:function(){
		var a=document;
		a.writeln('<form id="Bbsxp_fileinput_form" name="Bbsxp_fileinput_form" action="?" target="_self" method="post" enctype="multipart/form-data" style="margin:0;padding:0;"><div id="Bbsxp_fileinput_formarea"><img src="images/fileitem.gif" title="���������Ӹ���" border="0" /> <a href="javascript:;" title="�����Ӹ���">��Ӹ���<input id="Bbsxp_fileinput_0" name="Bbsxp_fileinput_0" class="addfile" size="1" type="file" onchange="BbsxpFileInput.add();" /></a>��<span id="Bbsxp_fileinput_upbtn"><a href="javascript:BbsxpFileInput.send();" title="����ϴ�">�ϴ�����</a></span>��<span id="Bbsxp_fileinput_msg"></span>'+BbsxpFileInput.readme+'</div></form><div id="Bbsxp_fileinput_show"></div>');
	},
	send:function(){
		if (BbsxpFileInput.realcount>0){
			BbsxpFileInput.$('Bbsxp_fileinput_'+BbsxpFileInput.count).disabled=true;
			BbsxpFileInput.$('Bbsxp_fileinput_upbtn').innerHTML='�ϴ��У����Ե�..';
			BbsxpFileInput.$('Bbsxp_fileinput_form').submit();
		}else{
			alert('������Ӹ������ϴ���');
		}
	},
	clear:function(){
		for (var i=BbsxpFileInput.count; i>0; --i){
			BbsxpFileInput.remove(i);
		}
		BbsxpFileInput.$('Bbsxp_fileinput_form').reset();
		var o=BbsxpFileInput.$('Bbsxp_fileinput_btn_0');
		if(o){o.parentNode.parentNode.removeChild(o.parentNode);}
		BbsxpFileInput.$('Bbsxp_fileinput_0').disabled=false;
		BbsxpFileInput.$('Bbsxp_fileinput_0').style.display='';
		BbsxpFileInput.count=0;
		BbsxpFileInput.realcount=0;
	}
}
BbsxpFileInput_OnResize=function(){
	var o=parent.document.getElementById("UpLoadIframe");
	(o.style||o).height=(parseInt(BbsxpFileInput.realcount)*16+18)+'px';
}
BbsxpFileInput_OnEcho=function(str){
	BbsxpFileInput.$('Bbsxp_fileinput_msg').innerHTML=str;
}
Bbsxp_InsertIntoEdit=function(AttachmentID,FileName,FileMIME,SaveFile){
	parent.document.form.UpFileID.value+=AttachmentID+',';
	parent.$("UpFile").innerHTML+="<img src=images/affix.gif /><a target=_blank href='"+SaveFile+"' title='"+FileName+"'>"+FileName+"</a><br />";

	if (parent.YuZi_CurrentMode.toLowerCase() == 'design') {
		parent.BBSXPEditorForm.focus();
		if (FileMIME.split ('/')[0]=="image"){
			var element = parent.document.createElement("img");
			element.src = ""+SaveFile+"";
			element.border = 0;
		}else{
			var element = parent.document.createElement("span");
			element.innerHTML = "[url="+SaveFile+"][img]images/affix.gif[/img]"+FileName+"[/url]"
		}
		parent.BBSXPSelection();
		parent.BBSXPInsertItem(element);
		parent.BBSXPDisableMenu();
	}
	else {
		if (FileMIME.split ('/')[0]=="image"){
			parent.$("BBSXPCodeForm").value += "[img]"+SaveFile+"[/img]";
		}else{
			parent.$("BBSXPCodeForm").value += "[url="+SaveFile+"][img]images/affix.gif[/img]"+FileName+"[/url]"
		}
	}
}
//-->
</script>
</head>
<body>

<script language="javascript">
<!--
BbsxpFileInput.uped=parseInt('0');
BbsxpFileInput.max=parseInt('10');
BbsxpFileInput.once=parseInt('10');
BbsxpFileInput.MaxFileSize=parseInt('<%=UpMaxFileSize%>')/1024;
BbsxpFileInput.FileTypes='<%=UpFileTypes%>';

BbsxpFileInput.readme=' <a style="CURSOR: help" title="������С:'+BbsxpFileInput.MaxFileSize+'K\n�ϴ�����:'+BbsxpFileInput.FileTypes+'">(�鿴�ϴ�����:��С������)</a>';
BbsxpFileInput.init();	
BbsxpFileInput_OnResize();
if ('<%=FileMessage%>'!=""){
	BbsxpFileInput.$('Bbsxp_fileinput_msg').innerHTML='<font color=red><%=FileMessage%></font>';
}
//-->
</script>
</body>
</html>
<%
If ""&Script&""<>"" Then response.write "<script language=""javascript"">"&Script&"</script>"
%>