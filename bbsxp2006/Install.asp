<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
youip="127.0.0.1" '您本机的IP地址

if SiteSettings("AdminPassword")<>"" and Request.ServerVariables("REMOTE_ADDR")<>youip then error("<li>为了安全起见，请编辑 <font color=red>Install.asp</font> 内本机的IP地址<li>请把它设置成 <font color=red>"&Request.ServerVariables("REMOTE_ADDR")&"</font>")

top

if Request("menu")="ok" then
Administrators=Request.Form("Administrators")
if Administrators="" then error("<li>您没有设置管理员")
if Request("Adminpassword")="" then error("<li>您没有设置管理密码")
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&Administrators&"'" ).eof Then error("<li>"&Administrators&"的用户资料还未<a href=CreateUser.asp>注册</a>")

Conn.execute("update [BBSXP_SiteSettings] set Adminpassword='"&MD5(Request("Adminpassword"))&"'")
Conn.execute("update [BBSXP_Users] set membercode=5 where UserName='"&Administrators&"'")

Message=Message&"<li>安装成功<li><a href=Admin.asp target=_top>返回社区管理</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Default.asp>")
end if

%>
<title>安装BBSXP... - Powered By BBSXP</title><br><form method="POST"><input type="hidden" value="ok" name="menu">


<table cellpadding="1" cellspacing="1" border="0" class=a2 align=center width="380">
<tr class=a1>
<td align=center colspan="2" height="25"><p><b>设置社区管理密码</b></p></td>
</tr>


<tr class=a3><td align=right width="30%" height="20">管 理 员：</td>
<td height="20">
<input name=Administrators size="30"></td></tr>
<tr class=a4><td align=right width="30%" height="10">管理密码：</td>
<td height="10">
<input name=Adminpassword type=password size="30"></td></tr>
<tr class=a3><td align=center height="10" colspan="2"><input type="submit" value=" 下一步 "></td>
</tr>
</table>
</FORM>
<center>注：<a href="CreateUser.asp">如果管理员用户名还没有注册请点击这里进行注册</a></center>
<%htmlend%>