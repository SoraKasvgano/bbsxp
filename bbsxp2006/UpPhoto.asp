<!-- #include file="Setup.asp" -->
<!--#include FILE="inc/UpFile_Class.asp"-->
<%
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")

if Request("menu")="up" then

id=Conn.Execute("Select id From [BBSXP_Users] where UserName='"&CookieUserName&"'")(0)

select case ""&SiteSettings("UpFileOption")&""
case ""
error("<li>对不起，本系统关闭文件上传功能")
case "SoftArtisans.FileUp"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set FileUP = Server.CreateObject("SoftArtisans.FileUp")

FileMIME=FileUP.ContentType						'类型
FileSize=FileUP.TotalBytes						'大小

if FileMIME="image/gif" then
FileExt="gif"
elseif FileMIME="image/pjpeg" then
FileExt="jpg"
end if

if FileExt<>"gif" and FileExt<>"jpg" then error("<li>您上传的照片不是GIF、JPG格式的文件")
if FileSize < 0 then error2("当前文件为空文件")
if FileSize > int(SiteSettings("MaxPhotoSize")) then error("<li>图片大小不得超过 "&CheckSize(SiteSettings("MaxPhotoSize"))&"<br>当前的图片大小为 "&CheckSize(FileSize)&"")

FileUP.SaveAs Server.mappath("UpFile/UpPhoto/"&id&"."&FileExt&"")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "ADODB.Stream"
set upfile=new upfile_class				'建立上传对象
upfile.NoAllowExt="asp;asa;cdx;cer;"	'设置上传类型的黑名单
upfile.GetData ()						'取得上传数据

FileMIME=upfile.file("file").FileMIME	'类型
FileSize=upfile.file("file").FileSize	'大小

if FileMIME="image/gif" then
FileExt="gif"
elseif FileMIME="image/pjpeg" then
FileExt="jpg"
end if

if FileExt<>"gif" and FileExt<>"jpg" then error("<li>您上传的照片不是GIF、JPG格式的文件")
if FileSize < 0 then error2("当前文件为空文件")
if FileSize > int(SiteSettings("MaxPhotoSize")) then error2("文件大小不得超过 "&CheckSize(SiteSettings("MaxPhotoSize"))&"\n当前的文件大小为 "&CheckSize(FileSize)&"")

upfile.SaveToFile "file",Server.mappath("UpFile/UpPhoto/"&id&"."&FileExt&"")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
end select

set FileUP=Nothing


Conn.execute("update [BBSXP_Users] set Userphoto='UpFile/UpPhoto/"&id&"."&FileExt&"' where UserName='"&CookieUserName&"'")
Message=Message&"<li>照片上传成功<li><a href=UpPhoto.asp>返回上传照片</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=UpPhoto.asp>")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="custom" then

Userphoto=HTMLEncode(Request.Form("Userphoto"))
if Userphoto="" then error("<li>照片URL不能没有填写")
if instr(Userphoto,";")>0 or instr(Userphoto,"%")>0 or instr(Userphoto,"javascript:")>0 then error2("URL中不能含有特殊符号")

Conn.execute("update [BBSXP_Users] set Userphoto='"&Userphoto&"' where UserName='"&CookieUserName&"'")

End If

Userphoto=Conn.Execute("Select Userphoto From [BBSXP_Users] where UserName='"&CookieUserName&"'")(0)

%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <a href="UpPhoto.asp">上传照片</a></td>
</tr>
</table><br>

<form enctype=multipart/form-data method=Post action=UpPhoto.asp?menu=up>

<table width=100% border=0 cellpadding="0" cellspacing="1" class=a2 height="218" align="center"><tr>
	<td class=a1 height="25" width="70%"><p align=center>请点取下面的“浏览”按键选择您要上传的图片</p>
	</td><td class=a1><div align=center>您目前的照片</span></div></td>
</tr><tr><td class=a3 width="70%"><center>
<input type=file name="file" size=33><br><br><input type=submit value=" 确 认 "></form></td>
		<td rowspan=3 class=a3 align="center"><%if Userphoto<>"" then%><img src="<%=Userphoto%>" onload='javascript:if(this.width>200)this.width=200'><%end if%></td>
</tr><tr><td class=a1 height="25" width="70%"><center>自定义照片</td><tr>
<td class=a3 width="70%"><form method=Post action=UpPhoto.asp?menu=custom><center>
	照片URL：<input size="35" value="<%=Userphoto%>" name="Userphoto">
<br><br><input type=submit value=" 更 新 "></form></td></table>




<%
htmlend
%>