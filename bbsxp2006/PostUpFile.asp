<!-- #include file="Setup.asp" -->
<!--#include FILE="inc/UpFile_Class.asp"-->
<%
if CookieUserName=empty then error2("����δ��¼��̳")

sub check(typ)
if instr("|"&LCase(SiteSettings("UpFileTypes"))&"|","|"&LCase(typ)&"|") <= 0 then error2("�Բ��𣬹���Ա�趨����̳�������ϴ� "&typ&" ��ʽ���ļ�")
end sub

if Request.ServerVariables("request_method") = "POST" then

fname=""&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&""


if SiteSettings("UpFileOption")="" then
error2("�Բ��𣬹���Ա�ر��ļ��ϴ�����")

elseif SiteSettings("UpFileOption")="ADODB.Stream" or SiteSettings("AttachmentsSaveOption")=1 then

set upfile=new upfile_class				'�����ϴ�����
upfile.NoAllowExt="asp;asa;cdx;cer;"	'�����ϴ����͵ĺ�����
upfile.GetData ()						'ȡ���ϴ�����

FileName=upfile.file("file").FileName	'�ļ���
FileExt=upfile.file("file").FileExt		'��׺��
FileMIME=upfile.file("file").FileMIME	'����
FileSize=upfile.file("file").FileSize	'��С
FileData=upfile.FileData("file")		'�ļ�����

if FileSize < 1 then error2("��ǰ�ļ�Ϊ���ļ�")
if FileSize > int(SiteSettings("MaxFileSize")) then error2("�ļ���С���ó��� "&CheckSize(SiteSettings("MaxFileSize"))&"\n��ǰ���ļ���СΪ "&CheckSize(FileSize)&"")
check(FileExt)


if SiteSettings("AttachmentsSaveOption")=1 then
TotalUserPostAttachments=conn.execute("Select sum(ContentSize) from [BBSXP_PostAttachments] where UserName='"&CookieUserName&"'")(0)
if TotalUserPostAttachments>SiteSettings("MaxPostAttachmentsSize") then error2("���ĸ����ռ�������")
Rs.Open "[BBSXP_PostAttachments]",conn,1,3 
Rs.addnew 
Rs("UserName")=CookieUserName
Rs("FileName")=FileName
Rs("ContentType")=FileMIME
Rs("ContentSize")=FileSize
Rs("Content")=FileData
Rs.update
SaveFile="PostAttachment.asp?AttachmentID="&Rs("id")&""
AttachmentID=Rs("id")
Rs.close
else
SaveFile="UpFile/UpAttachment/"&fname&"."&FileExt&""
upfile.SaveToFile "file",Server.mappath(""&SaveFile&"")
end if


elseif SiteSettings("UpFileOption")="SoftArtisans.FileUp" then

Set FileUP = Server.CreateObject("SoftArtisans.FileUp")

FileName=FileUP.ShortFilename					'�ļ���
FileExt=mid(FileName,InStrRev(FileName, ".")+1)	'��׺��
FileMIME=FileUP.ContentType						'����
FileSize=FileUP.TotalBytes						'��С

if FileSize < 1 then error2("��ǰ�ļ�Ϊ���ļ�")
if FileSize > int(SiteSettings("MaxFileSize")) then error2("�ļ���С���ó��� "&CheckSize(SiteSettings("MaxFileSize"))&"\n��ǰ���ļ���СΪ "&CheckSize(FileSize)&"")
check(FileExt)

SaveFile="UpFile/UpAttachment/"&fname&"."&FileExt&""
FileUP.SaveAs Server.mappath(""&SaveFile&"")

end if

set FileUP=Nothing

%>
<body topmargin=0  rightmargin=0  Leftmargin=0>
<script>
if ("<%=FileExt%>"=="gif" || "<%=FileExt%>"=="jpg" || "<%=FileExt%>"=="png" || "<%=FileExt%>"=="bmp"){
img="[img]<%=SaveFile%>[/img]"
}else{
img="[url=<%=SaveFile%>][img]images/affix.gif[/img]<%=FileName%>[/url]"
}

<%if SiteSettings("AttachmentsSaveOption")=1 then%>
parent.yuziform.UpFileID.value+='<%=AttachmentID%>,';
<%end if%>

parent.UpFile.innerHTML+='<img src=images/affix.gif><a target=_blank href=<%=SaveFile%>><%=FileName%></a><br>'
parent.IframeID.document.body.innerHTML+=img;
document.oncontextmenu = new Function('return false')
</script>
<table cellpadding=0 cellspacing=0 width=100%  height=100% class=a4><tr><td><a target=_blank href=<%=SaveFile%>><%=SiteSettings("SiteURL")%><%=SaveFile%></a> [<a href=PostUpFile.asp>�����ϴ�</a>] </td></tr></table>
<%else%>
<script>if(top==self)document.location='';</script>
<body topmargin=0  rightmargin=0  Leftmargin=0>
<form enctype=multipart/form-data method=Post action=PostUpFile.asp?menu=up>
<table cellpadding=0 cellspacing=0 width=100% class=a4>
<tr><td>
<td><input type=file style=FONT-SIZE:9pt name=file size=60> <input style=FONT-SIZE:9pt type="submit" value=" �� �� "></td></tr></table>
<%
end if
CloseDatabase
%>