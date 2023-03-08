<!-- #include file="Setup.asp" -->
<!--#include FILE="inc/UpFile_Class.asp"-->
<%
if CookieUserName=empty then error2("您还未登录论坛")

sub check(typ)
if instr("|"&LCase(SiteSettings("UpFileTypes"))&"|","|"&LCase(typ)&"|") <= 0 then error2("对不起，管理员设定本论坛不允许上传 "&typ&" 格式的文件")
end sub

if Request.ServerVariables("request_method") = "POST" then

fname=""&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&""


if SiteSettings("UpFileOption")="" then
error2("对不起，管理员关闭文件上传功能")

elseif SiteSettings("UpFileOption")="ADODB.Stream" or SiteSettings("AttachmentsSaveOption")=1 then

set upfile=new upfile_class				'建立上传对象
upfile.NoAllowExt="asp;asa;cdx;cer;"	'设置上传类型的黑名单
upfile.GetData ()						'取得上传数据

FileName=upfile.file("file").FileName	'文件名
FileExt=upfile.file("file").FileExt		'后缀名
FileMIME=upfile.file("file").FileMIME	'类型
FileSize=upfile.file("file").FileSize	'大小
FileData=upfile.FileData("file")		'文件数据

if FileSize < 1 then error2("当前文件为空文件")
if FileSize > int(SiteSettings("MaxFileSize")) then error2("文件大小不得超过 "&CheckSize(SiteSettings("MaxFileSize"))&"\n当前的文件大小为 "&CheckSize(FileSize)&"")
check(FileExt)


if SiteSettings("AttachmentsSaveOption")=1 then
TotalUserPostAttachments=conn.execute("Select sum(ContentSize) from [BBSXP_PostAttachments] where UserName='"&CookieUserName&"'")(0)
if TotalUserPostAttachments>SiteSettings("MaxPostAttachmentsSize") then error2("您的附件空间已满！")
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

FileName=FileUP.ShortFilename					'文件名
FileExt=mid(FileName,InStrRev(FileName, ".")+1)	'后缀名
FileMIME=FileUP.ContentType						'类型
FileSize=FileUP.TotalBytes						'大小

if FileSize < 1 then error2("当前文件为空文件")
if FileSize > int(SiteSettings("MaxFileSize")) then error2("文件大小不得超过 "&CheckSize(SiteSettings("MaxFileSize"))&"\n当前的文件大小为 "&CheckSize(FileSize)&"")
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
<table cellpadding=0 cellspacing=0 width=100%  height=100% class=a4><tr><td><a target=_blank href=<%=SaveFile%>><%=SiteSettings("SiteURL")%><%=SaveFile%></a> [<a href=PostUpFile.asp>继续上传</a>] </td></tr></table>
<%else%>
<script>if(top==self)document.location='';</script>
<body topmargin=0  rightmargin=0  Leftmargin=0>
<form enctype=multipart/form-data method=Post action=PostUpFile.asp?menu=up>
<table cellpadding=0 cellspacing=0 width=100% class=a4>
<tr><td>
<td><input type=file style=FONT-SIZE:9pt name=file size=60> <input style=FONT-SIZE:9pt type="submit" value=" 上 传 "></td></tr></table>
<%
end if
CloseDatabase
%>