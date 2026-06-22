<!-- #include file="Setup.asp" -->
<!--#include FILE="inc/UpFile_Class.asp"-->
<%
top
if CookieUserName=empty then error("<li>您尚未<a href=Login.asp>登录</a>论坛")

if Request("menu")="up" then
id=Conn.Execute("Select id From [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'")(0)

select case ""&SiteSettings("UpFileOption")&""
case ""
error("<li>抱歉，本系统已关闭文件上传功能")
case "SoftArtisans.FileUp"
Set FileUP = Server.CreateObject("SoftArtisans.FileUp")

FileMIME=FileUP.ContentType
FileSize=FileUP.TotalBytes

if FileMIME="image/gif" then
FileExt="gif"
elseif FileMIME="image/pjpeg" then
FileExt="jpg"
end if

if FileExt<>"gif" and FileExt<>"jpg" then error("<li>上传图片只允许 GIF 或 JPG 格式")
if FileSize < 0 then error2("当前文件为空文件")
if FileSize > int(SiteSettings("MaxFaceSize")) then error("<li>图片大小超过限制 "&CheckSize(SiteSettings("MaxFaceSize"))&"<br>当前图片大小为 "&CheckSize(FileSize)&"")

FileUP.SaveAs Server.mappath("UpFile/UpFace/"&id&"."&FileExt&"")

case "ADODB.Stream"
set upfile=new upfile_class
upfile.NoAllowExt="asp;asa;cdx;cer;"
upfile.GetData ()

FileMIME=upfile.file("file").FileMIME
FileSize=upfile.file("file").FileSize

if FileMIME="image/gif" then
FileExt="gif"
elseif FileMIME="image/pjpeg" then
FileExt="jpg"
end if

if FileExt<>"gif" and FileExt<>"jpg" then error("<li>上传图片只允许 GIF 或 JPG 格式")
if FileSize < 0 then error2("当前文件为空文件")
if FileSize > int(SiteSettings("MaxFaceSize")) then error2("文件大小超过限制 "&CheckSize(SiteSettings("MaxFaceSize"))&"\n当前文件大小为 "&CheckSize(FileSize)&"")

upfile.SaveToFile "file",Server.mappath("UpFile/UpFace/"&id&"."&FileExt&"")
end select

set FileUP=Nothing

Conn.execute("update [BBSXP_Users] set Userface='UpFile/UpFace/"&id&"."&FileExt&"' where UserName='"&SqlString(CookieUserName)&"'")
Message=Message&"<li>头像上传成功<li><a href=UpFace.asp>继续上传头像</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=UpFace.asp>")

elseif Request("menu")="custom" then
RawUserface=Trim(Request.Form("Userface"))
Userface=SafeUrl(RawUserface)
if RawUserface="" then error("<li>头像URL地址没有填写")
if Userface="" then error2("URL地址非法")

Conn.execute("update [BBSXP_Users] set Userface='"&HTMLEncode(Userface)&"' where UserName='"&SqlString(CookieUserName)&"'")
End If

CurrentUserFace=SafeUrl(Conn.Execute("Select Userface From [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'")(0))
%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <a href="UpFace.asp">上传头像</a></td>
</tr>
</table><br>

<form enctype=multipart/form-data method=Post action=UpFace.asp?menu=up>
<table width=100% border=0 cellpadding="0" cellspacing="1" class=a2 height="218" align="center"><tr>
	<td class=a1 height="25" width="70%"><p align=center>点击浏览按钮选择要上传的头像图片</p>
	</td><td class=a1><div align=center>您目前的头像</div></td>
</tr><tr><td class=a3 width="70%"><center>
<input type=file name="file" size=33><br><br><input type=submit value=" 确 定 "></form></td>
		<td rowspan=3 class=a3 align="center"><%if CurrentUserFace<>"" then%><img src="<%=CurrentUserFace%>"><%end if%></td>
</tr><tr><td class=a1 height="25" width="70%"><center>自定义头像</td><tr>
<td class=a3 width="70%"><form method=Post action=UpFace.asp?menu=custom><center>头像URL：<input size="35" value="<%=CurrentUserFace%>" name="Userface">
<br><br><input type=submit value=" 确 定 "></form></td></table>

<%
htmlend
%>