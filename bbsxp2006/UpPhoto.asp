<!-- #include file="Setup.asp" -->
<!--#include FILE="inc/UpFile_Class.asp"-->
<%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")

if Request("menu")="up" then

id=Conn.Execute("Select id From [BBSXP_Users] where UserName='"&CookieUserName&"'")(0)

select case ""&SiteSettings("UpFileOption")&""
case ""
error("<li>�Բ��𣬱�ϵͳ�ر��ļ��ϴ�����")
case "SoftArtisans.FileUp"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set FileUP = Server.CreateObject("SoftArtisans.FileUp")

FileMIME=FileUP.ContentType						'����
FileSize=FileUP.TotalBytes						'��С

if FileMIME="image/gif" then
FileExt="gif"
elseif FileMIME="image/pjpeg" then
FileExt="jpg"
end if

if FileExt<>"gif" and FileExt<>"jpg" then error("<li>���ϴ�����Ƭ����GIF��JPG��ʽ���ļ�")
if FileSize < 0 then error2("��ǰ�ļ�Ϊ���ļ�")
if FileSize > int(SiteSettings("MaxPhotoSize")) then error("<li>ͼƬ��С���ó��� "&CheckSize(SiteSettings("MaxPhotoSize"))&"<br>��ǰ��ͼƬ��СΪ "&CheckSize(FileSize)&"")

FileUP.SaveAs Server.mappath("UpFile/UpPhoto/"&id&"."&FileExt&"")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "ADODB.Stream"
set upfile=new upfile_class				'�����ϴ�����
upfile.NoAllowExt="asp;asa;cdx;cer;"	'�����ϴ����͵ĺ�����
upfile.GetData ()						'ȡ���ϴ�����

FileMIME=upfile.file("file").FileMIME	'����
FileSize=upfile.file("file").FileSize	'��С

if FileMIME="image/gif" then
FileExt="gif"
elseif FileMIME="image/pjpeg" then
FileExt="jpg"
end if

if FileExt<>"gif" and FileExt<>"jpg" then error("<li>���ϴ�����Ƭ����GIF��JPG��ʽ���ļ�")
if FileSize < 0 then error2("��ǰ�ļ�Ϊ���ļ�")
if FileSize > int(SiteSettings("MaxPhotoSize")) then error2("�ļ���С���ó��� "&CheckSize(SiteSettings("MaxPhotoSize"))&"\n��ǰ���ļ���СΪ "&CheckSize(FileSize)&"")

upfile.SaveToFile "file",Server.mappath("UpFile/UpPhoto/"&id&"."&FileExt&"")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
end select

set FileUP=Nothing


Conn.execute("update [BBSXP_Users] set Userphoto='UpFile/UpPhoto/"&id&"."&FileExt&"' where UserName='"&CookieUserName&"'")
Message=Message&"<li>��Ƭ�ϴ��ɹ�<li><a href=UpPhoto.asp>�����ϴ���Ƭ</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=UpPhoto.asp>")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="custom" then

Userphoto=HTMLEncode(Request.Form("Userphoto"))
if Userphoto="" then error("<li>��ƬURL����û����д")
if instr(Userphoto,";")>0 or instr(Userphoto,"%")>0 or instr(Userphoto,"javascript:")>0 then error2("URL�в��ܺ����������")

Conn.execute("update [BBSXP_Users] set Userphoto='"&Userphoto&"' where UserName='"&CookieUserName&"'")

End If

Userphoto=Conn.Execute("Select Userphoto From [BBSXP_Users] where UserName='"&CookieUserName&"'")(0)

%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <a href="UpPhoto.asp">�ϴ���Ƭ</a></td>
</tr>
</table><br>

<form enctype=multipart/form-data method=Post action=UpPhoto.asp?menu=up>

<table width=100% border=0 cellpadding="0" cellspacing="1" class=a2 height="218" align="center"><tr>
	<td class=a1 height="25" width="70%"><p align=center>���ȡ����ġ����������ѡ����Ҫ�ϴ���ͼƬ</p>
	</td><td class=a1><div align=center>��Ŀǰ����Ƭ</span></div></td>
</tr><tr><td class=a3 width="70%"><center>
<input type=file name="file" size=33><br><br><input type=submit value=" ȷ �� "></form></td>
		<td rowspan=3 class=a3 align="center"><%if Userphoto<>"" then%><img src="<%=Userphoto%>" onload='javascript:if(this.width>200)this.width=200'><%end if%></td>
</tr><tr><td class=a1 height="25" width="70%"><center>�Զ�����Ƭ</td><tr>
<td class=a3 width="70%"><form method=Post action=UpPhoto.asp?menu=custom><center>
	��ƬURL��<input size="35" value="<%=Userphoto%>" name="Userphoto">
<br><br><input type=submit value=" �� �� "></form></td></table>




<%
htmlend
%>