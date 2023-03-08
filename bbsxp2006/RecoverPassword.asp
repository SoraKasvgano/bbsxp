<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" --><%
top

Randomize
RandomCode=int(rnd*999999)+1

UserName=HTMLEncode(Request.Form("UserName"))
code=int(Request("code"))
Userpass=Trim(Request.Form("Userpass"))

if UserName<>"" then
sql="select * from [BBSXP_Users] where UserName='"&UserName&"'"
Rs.Open sql,Conn,1
if Rs.eof then error2(""&UserName&"的用户资料不存在")
UserMail=Rs("UserMail")
birthday=Rs("birthday")
PasswordAnswer=Rs("PasswordAnswer")
PasswordQuestion=Rs("PasswordQuestion")
Rs.close
end if

select case Request("menu")
case ""
index

case "MailRecover"
if UserName="" then Message=Message&"<li>用户名称没有填写"
if Request.Form("UserMail")<>UserMail then Message=Message&"<li>Email填写错误"
if Application(SiteSettings("CacheName")&UserName&"MailRecover")<>"" then Message=Message&"<li>2次取回密码的时间太短"
if Message<>"" then error(""&Message&"")

Application(SiteSettings("CacheName")&UserName&"MailRecover")=RandomCode
Mailaddress=UserMail
MailTopic="用户找回密码"
body=""&vbCrlf&"亲爱的"&UserName&", 您好!"&vbCrlf&""&vbCrlf&"　　请在20分钟内点击以下链接, 系统将自动生成新的密码!"&vbCrlf&""&vbCrlf&"　* "&SiteSettings("SiteURL")&"RecoverPassword.asp?menu=MailRecoverok&username="&UserName&"&code="&RandomCode&""&vbCrlf&""&vbCrlf&"　* "&SiteSettings("SiteName")&"("&SiteSettings("SiteURL")&"Default.asp)"&vbCrlf&""&vbCrlf&"　* 最后, 有几点注意事项请您牢记"&vbCrlf&"1、请遵守《计算机信息网络国际联网安全保护管理办法》里的一切规定。"&vbCrlf&"2、使用轻松而健康的话题，所以请不要涉及政治、宗教等敏感话题。"&vbCrlf&"3、承担一切因您的行为而直接或间接导致的民事或刑事法律责任。"&vbCrlf&""&vbCrlf&""&vbCrlf&"论坛服务由 "&SiteSettings("CompanyName")&"("&SiteSettings("CompanyURL")&") 提供　程序制作：YUZI工作室(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""
%>
<!-- #include file="inc/Mail.asp" -->
<%
error2("请在20分钟之内到邮箱中取回密码")
case "MailRecoverok"
if Application(SiteSettings("CacheName")&UserName&"MailRecover")="" then error2("取回密码时间已过期")
if code<>Application(SiteSettings("CacheName")&UserName&"MailRecover") then error2("邮件中的验证码与服务器最后一次发送的验证码不同")
Conn.execute("update [BBSXP_Users] set Userpass='"&md5(RandomCode)&"' where UserName='"&UserName&"'")
%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		→ 找回密码</td>
	</tr>
</table>
<br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class="a2">
		<tr>
			<td width="100%" align="center" height="25" class="a1"><%=UserName%>的新密码</td>
		</tr>
		<tr class="a3">
			<td width="100%" align="center" height="13"><font size="7"><%=RandomCode%></font></td>
		</tr>
		</table>
<%
Application(SiteSettings("CacheName")&UserName&"MailRecover")=""

case "rejigger"
rejigger
case "rejiggerok"
if UserName="" then error2("用户名称没有填写")
if ""&birthday&""="" then error2("您注册的时候没有填写出生日期，所以无法通过此功能找回密码")
if ""&PasswordAnswer&""="" or ""&PasswordQuestion&""="" then error2("您注册的时候没有填写密码提示问题或者密码提示答案，所以无法通过此功能找回密码")
if Request("birthday")<>birthday then error2("出生日期填写错误")
if md5(Request("PasswordAnswer"))<>PasswordAnswer then error2("答案错误")
if Request("Userpass")="" then error2("请输入新的密码")
if Request("Userpass")<>Request("Userpass2") then error2("您2次输入的密码不同")
Conn.execute("update [BBSXP_Users] set Userpass='"&md5(Userpass)&"' where UserName='"&UserName&"'")
Message=Message&"<li>更改密码成功<li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Default.asp>")
end select


sub rejigger
if UserName="" then error2("用户名称没有填写")
if ""&birthday&""="" then error2("您注册的时候没有填写出生日期，所以无法通过此功能找回密码")
if Request("birthday")<>birthday then error2("出生日期填写错误")
%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		→ 更改密码</td>
	</tr>
</table>
<br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class="a2">
	<form method="POST" action="RecoverPassword.asp?menu=rejiggerok">
		<input type="hidden" name="UserName" value="<%=Request("UserName")%>">
		<input type="hidden" name="birthyear" value="<%=Request("birthyear")%>">
		<input type="hidden" name="birthmonth" value="<%=Request("birthmonth")%>">
		<input type="hidden" name="birthday" value="<%=Request("birthday")%>">
		<tr>
			<td width="100%" align="center" height="20" class="a1" colspan="2">更改密码</td>
		</tr>
		<tr class="a3">
			<td width="50%" align="right" height="25">请回答问题：</td>
			<td width="50%" height="25"><%=PasswordQuestion%></td>
		</tr>
		<tr class="a4">
			<td width="50%" align="right" height="20">答案：</td>
			<td width="50%" height="20"><input size="15" value name="PasswordAnswer"></td>
		</tr>
		<tr class="a3">
			<td width="50%" align="right" height="20">请输入新的密码：</td>
			<td width="50%" height="20">
			<input type="password" size="15" name="Userpass"></td>
		</tr>
		<tr class="a4">
			<td width="50%" align="right" height="20">请再次输入密码：</td>
			<td width="50%" height="20">
			<input type="password" size="15" name="Userpass2"></td>
		</tr>
		<tr class=a3>
			<td width="100%" align="center" height="20" colspan="2">
			<input type="submit" value=" 确定 ">　<input type="reset" value=" 取消 "></td>
		</tr>
	</form>
</table>
<br>
<center><a href="javascript:history.back()">BACK </a><br>
<%
end sub

sub index
%><table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		→ 找回密码</td>
	</tr>
</table>
<br>
<script src="inc/birthday.js"></script>
<center><b><font size=5>请选择找回密码的方式</font></b></center>
<table border="0" width="100%">
	<tr>
		<td align="center">
	<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a2">
		<tr>
			<td class="a1" align="center" height="25">
			密码问题 &amp; 答案找回</td>
		</tr>
		<tr>
			<td valign="top" class="a3" align="center"><form action="RecoverPassword.asp?menu=rejigger" method="POST">

				用户名称：<input size="25" name="UserName" value="<%=CookieUserName%>"><br>
	
				出生日期：<input onfocus="calendar()" name="birthday" size="25"><br>
				<input type="submit" value=" 确定 ">　<input type="reset" value=" 取消 ">
		
			</td></form>
		</tr>
		</table>
		</td>
		<td align="center">
	<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a2" id="table2">
		<tr>
			<td class="a1" align="center" height="25">
			Email找回</td>
		</tr>
		<tr>
			<td valign="top" class="a3" align="center"><form action="RecoverPassword.asp?menu=MailRecover" method="POST">
				用户名称：<input size="25" name="UserName" value="<%=CookieUserName%>"><br>
&nbsp;&nbsp;
				Email：<input name="UserMail" size="25"><br>
				<input type="submit" value=" 确定 ">　<input type="reset" value=" 取消 ">
			</td></form>
		</tr>
		</table>
		</td>
	</tr>
</table>
<br>
<center><a href="javascript:history.back()">BACK </a><br>
<%

end sub

htmlend
%>