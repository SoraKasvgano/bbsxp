<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
top

if Request("menu")="passwordok" then
Response.Cookies("password")=Request("password")
response.redirect Request("url")
end if

if Request.ServerVariables("request_method") = "POST" then

if sitesettings("EnableAntiSpamTextGenerateForLogin")=1 then
if Request.Form("VerifyCode")<>Session("VerifyCode") then error("<li>��֤�����")
end if

url=Trim(Request.Form("url"))
UserName=HTMLEncode(Request.Form("UserName"))
Userpass=md5(Trim(Request.Form("Userpass")))
if UserName=empty then error("<li>�û���û������")

sql="select * from [BBSXP_Users] where UserName='"&UserName&"'"
Set Rs1=Conn.Execute(SQL)
if rs1.eof then error("<li>���û�����δע��")
if rs1("membercode")=0 then error("<li>�����ʺ���δ����")

if Len(rs1("Userpass"))<16 then
if Request("Userpass")<>rs1("Userpass") then error("<li>��������������")
Conn.execute("update [BBSXP_Users] set Userpass='"&Userpass&"' where UserName='"&UserName&"'")
elseif Len(rs1("Userpass"))=16 then
mdfive=16
if md5(Request("Userpass"))<>rs1("Userpass") then error("<li>��������������")
Conn.execute("update [BBSXP_Users] set Userpass='"&Userpass&"' where UserName='"&UserName&"'")
else
if Userpass<>rs1("Userpass") then error("<li>��������������")
end if

Response.Cookies("UserName")=escape(rs1("UserName"))
Response.Cookies("Userpass")=Userpass
Response.Cookies("eremite")="0"
if Request("eremite")="1" then Response.Cookies("eremite")="1"
if Request("xuansave")=1 then
Response.Cookies("eremite").Expires=date+9999
Response.Cookies("UserName").Expires=date+9999
Response.Cookies("Userpass").Expires=date+9999
end if
Session("VerifyCode")=""
if ""&url&""<>"" and instr(url,"Login.asp")=0 and instr(url,"Left.asp")=0 then
response.redirect url
else
response.write "<SCRIPT>top.location='"&SiteSettings("SiteURL")&"';</SCRIPT>"
end if


end if



select case Request("menu")
case ""
%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> ��  ��¼��̳</td>
</tr>
</table><br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>

<form action="Login.asp" method="POST">
<input type="hidden" value="<%=Request.ServerVariables("HTTP_REFERER")%>" name="url">
<tr>
<td width="328" height="25" align="center" class=a1>
��¼��̳
</td></tr><tr>
<td height="19" width="328" valign="top" align="center" class=a3>
�û�����:
<input size="15" name="UserName" value="<%=CookieUserName%>">&nbsp; <a href="CreateUser.asp">û��ע��?</a><br>
�û�����: <input type="password" size="15" value name="Userpass">&nbsp;
<a href="RecoverPassword.asp">�һ�����?</a><br>
<%if sitesettings("EnableAntiSpamTextGenerateForLogin")=1 then%>
�� ֤ ��: <input size="10" name="VerifyCode">&nbsp;&nbsp;
<img src="VerifyCode.asp" alt="��֤��,�������?����ˢ����֤��" style=cursor:pointer onclick="this.src='VerifyCode.asp'"><br>
<%end if%>

<input type="checkbox" value="1" name="xuansave" id="xuansave"><label for="xuansave">�Զ���¼</label>

<input type="checkbox" value="1" name="eremite" id="eremite"><label for="eremite">�����¼</label><br>
<input type="submit" value=" ��¼ ">��<input type="reset" value=" ȡ�� ">
</td></tr> </form></table>

<br>
<center>
<a href=javascript:history.back()>BACK </a><br>

<%
case "password"
%>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� ��֤����</td>
</tr>
</table>
<br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>

<form action="Login.asp" method="POST">
<input type="hidden" value="passwordok" name="menu">
<input type="hidden" value="<%=Request("url")%>" name="url">
<tr>
<td width="328" height="25" align="center" class=a1>
��¼��̳
</td></tr><tr>
<td height="19" width="328" valign="top" align="center" class=a3>
ͨ������: <input type="password" size="15" name="password"><br>

<input type="submit" value=" ��¼ ">��<input type="reset" value=" ȡ�� ">
</td></tr> </form></table>

<br>
<center>
<a href=javascript:history.back()>BACK </a><br>

<%

case "out"
Conn.execute("Delete from [BBSXP_UsersOnline] where sessionid='"&session.sessionid&"'")
Response.Cookies("UserName")=""
Response.Cookies("Userpass")=""
succtitle="�Ѿ��ɹ��˳�"
Message=Message&"<li><a href=Default.asp>������ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Default.asp>")

end select



htmlend
%>