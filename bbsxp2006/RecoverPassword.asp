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
if Rs.eof then error2(""&UserName&"���û����ϲ�����")
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
if UserName="" then Message=Message&"<li>�û�����û����д"
if Request.Form("UserMail")<>UserMail then Message=Message&"<li>Email��д����"
if Application(SiteSettings("CacheName")&UserName&"MailRecover")<>"" then Message=Message&"<li>2��ȡ�������ʱ��̫��"
if Message<>"" then error(""&Message&"")

Application(SiteSettings("CacheName")&UserName&"MailRecover")=RandomCode
Mailaddress=UserMail
MailTopic="�û��һ�����"
body=""&vbCrlf&"�װ���"&UserName&", ����!"&vbCrlf&""&vbCrlf&"��������20�����ڵ����������, ϵͳ���Զ������µ�����!"&vbCrlf&""&vbCrlf&"��* "&SiteSettings("SiteURL")&"RecoverPassword.asp?menu=MailRecoverok&username="&UserName&"&code="&RandomCode&""&vbCrlf&""&vbCrlf&"��* "&SiteSettings("SiteName")&"("&SiteSettings("SiteURL")&"Default.asp)"&vbCrlf&""&vbCrlf&"��* ���, �м���ע�����������μ�"&vbCrlf&"1�������ء��������Ϣ�������������ȫ��������취�����һ�й涨��"&vbCrlf&"2��ʹ�����ɶ������Ļ��⣬�����벻Ҫ�漰���Ρ��ڽ̵����л��⡣"&vbCrlf&"3���е�һ����������Ϊ��ֱ�ӻ��ӵ��µ����»����·������Ρ�"&vbCrlf&""&vbCrlf&""&vbCrlf&"��̳������ "&SiteSettings("CompanyName")&"("&SiteSettings("CompanyURL")&") �ṩ������������YUZI������(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""
%>
<!-- #include file="inc/Mail.asp" -->
<%
error2("����20����֮�ڵ�������ȡ������")
case "MailRecoverok"
if Application(SiteSettings("CacheName")&UserName&"MailRecover")="" then error2("ȡ������ʱ���ѹ���")
if code<>Application(SiteSettings("CacheName")&UserName&"MailRecover") then error2("�ʼ��е���֤������������һ�η��͵���֤�벻ͬ")
Conn.execute("update [BBSXP_Users] set Userpass='"&md5(RandomCode)&"' where UserName='"&UserName&"'")
%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		�� �һ�����</td>
	</tr>
</table>
<br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class="a2">
		<tr>
			<td width="100%" align="center" height="25" class="a1"><%=UserName%>��������</td>
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
if UserName="" then error2("�û�����û����д")
if ""&birthday&""="" then error2("��ע���ʱ��û����д�������ڣ������޷�ͨ���˹����һ�����")
if ""&PasswordAnswer&""="" or ""&PasswordQuestion&""="" then error2("��ע���ʱ��û����д������ʾ�������������ʾ�𰸣������޷�ͨ���˹����һ�����")
if Request("birthday")<>birthday then error2("����������д����")
if md5(Request("PasswordAnswer"))<>PasswordAnswer then error2("�𰸴���")
if Request("Userpass")="" then error2("�������µ�����")
if Request("Userpass")<>Request("Userpass2") then error2("��2����������벻ͬ")
Conn.execute("update [BBSXP_Users] set Userpass='"&md5(Userpass)&"' where UserName='"&UserName&"'")
Message=Message&"<li>��������ɹ�<li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Default.asp>")
end select


sub rejigger
if UserName="" then error2("�û�����û����д")
if ""&birthday&""="" then error2("��ע���ʱ��û����д�������ڣ������޷�ͨ���˹����һ�����")
if Request("birthday")<>birthday then error2("����������д����")
%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		�� ��������</td>
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
			<td width="100%" align="center" height="20" class="a1" colspan="2">��������</td>
		</tr>
		<tr class="a3">
			<td width="50%" align="right" height="25">��ش����⣺</td>
			<td width="50%" height="25"><%=PasswordQuestion%></td>
		</tr>
		<tr class="a4">
			<td width="50%" align="right" height="20">�𰸣�</td>
			<td width="50%" height="20"><input size="15" value name="PasswordAnswer"></td>
		</tr>
		<tr class="a3">
			<td width="50%" align="right" height="20">�������µ����룺</td>
			<td width="50%" height="20">
			<input type="password" size="15" name="Userpass"></td>
		</tr>
		<tr class="a4">
			<td width="50%" align="right" height="20">���ٴ��������룺</td>
			<td width="50%" height="20">
			<input type="password" size="15" name="Userpass2"></td>
		</tr>
		<tr class=a3>
			<td width="100%" align="center" height="20" colspan="2">
			<input type="submit" value=" ȷ�� ">��<input type="reset" value=" ȡ�� "></td>
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
		�� �һ�����</td>
	</tr>
</table>
<br>
<script src="inc/birthday.js"></script>
<center><b><font size=5>��ѡ���һ�����ķ�ʽ</font></b></center>
<table border="0" width="100%">
	<tr>
		<td align="center">
	<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a2">
		<tr>
			<td class="a1" align="center" height="25">
			�������� &amp; ���һ�</td>
		</tr>
		<tr>
			<td valign="top" class="a3" align="center"><form action="RecoverPassword.asp?menu=rejigger" method="POST">

				�û����ƣ�<input size="25" name="UserName" value="<%=CookieUserName%>"><br>
	
				�������ڣ�<input onfocus="calendar()" name="birthday" size="25"><br>
				<input type="submit" value=" ȷ�� ">��<input type="reset" value=" ȡ�� ">
		
			</td></form>
		</tr>
		</table>
		</td>
		<td align="center">
	<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a2" id="table2">
		<tr>
			<td class="a1" align="center" height="25">
			Email�һ�</td>
		</tr>
		<tr>
			<td valign="top" class="a3" align="center"><form action="RecoverPassword.asp?menu=MailRecover" method="POST">
				�û����ƣ�<input size="25" name="UserName" value="<%=CookieUserName%>"><br>
&nbsp;&nbsp;
				Email��<input name="UserMail" size="25"><br>
				<input type="submit" value=" ȷ�� ">��<input type="reset" value=" ȡ�� ">
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