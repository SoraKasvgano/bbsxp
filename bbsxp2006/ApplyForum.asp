<!-- #include file="Setup.asp" -->
<%
ForumName=HTMLEncode(Request.Form("ForumName"))
ForumIcon=HTMLEncode(Request.Form("ForumIcon"))
ForumLogo=HTMLEncode(Request.Form("ForumLogo"))
ForumIntro=HTMLEncode(Request.Form("ForumIntro"))
ForumRules=HTMLEncode(Request.Form("ForumRules"))


if SiteSettings("ForumApply")=0 then error("<li>�Բ��𣬱�ϵͳ�ر���̳���빦��")

top

if Request("menu")="Apply" then

if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
if ForumName="" then error("<li>��������̳����")
if Len(ForumName)>30 then  error("<li>��̳���Ʋ��ܴ��� 30 ���ַ�")
if Len(ForumIntro)>255 then  error("<li>��̳��鲻�ܴ��� 255 ���ֽ�")
if Request.Cookies("ApplyForum")="1" then error("<li>���Ѿ��������̳�ˣ��벻Ҫ�ٴ����룡")


Rs.Open "[BBSXP_Forums]",Conn,1,3
Rs.addNew
Rs("ForumName")=ForumName
Rs("moderated")=CookieUserName
Rs("ForumIntro")=ForumIntro
Rs("ForumRules")=ForumRules
Rs("ForumIcon")=ForumIcon
Rs("ForumLogo")=ForumLogo
Rs("ForumHide")=1
Rs.update
id=Rs("id")
Rs.close

Conn.execute("insert into [BBSXP_Favorites](UserName,name,url)values('"&SqlString(CookieUserName)&"','"&id&"','forum')")


Mailaddress=Conn.Execute("Select UserMail From [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'")(0)
MailTopic="��̳ϵͳ��֪ͨͨ"
body=""&vbCrlf&"�װ���"&CookieUserName&", ����!"&vbCrlf&""&vbCrlf&"������ϲ! ���Ѿ��ɹ���������"&SiteSettings("CompanyName")&"("&SiteSettings("CompanyURL")&")����̳ϵͳ, �ǳ���л��ʹ��"&SiteSettings("CompanyName")&"�ķ���!"&vbCrlf&""&vbCrlf&"��* �������Ϊ������̳�ṩ��һ���ȽϺüǵĵ�ַ,��������"&vbCrlf&""&vbCrlf&"��* URL: "&Request("ForumName")&"("&SiteSettings("SiteURL")&"Default.asp?id="&id&")"&vbCrlf&""&vbCrlf&"��* ���, �м���ע�����������μ�"&vbCrlf&"1������ʹ�ñ���̳ϵͳ�����κΰ���ɫ�顢�Ƿ����Լ�Σ�����Ұ�ȫ�����ݵ���̳��"&vbCrlf&"2�������ڱ�ϵͳ�û���ӵ�е���̳�ڷ����κ�ɫ�顢�Ƿ�������Σ�����Ұ�ȫ�����ۡ�"&vbCrlf&"3�����Ϲ���Υ�������Ը�����վ��Ȩɾ�������û��������ݣ���׷���䷨�����Ρ�"&vbCrlf&""&vbCrlf&""&vbCrlf&"��̳������ "&SiteSettings("CompanyName")&"("&SiteSettings("CompanyURL")&") �ṩ������������YUZI������(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""
%>
<!-- #include file="inc/Mail.asp" -->
<%

Response.Cookies("ApplyForum")="1"


Message="<li>��ϲ��������ɹ���<li>����Ϊ��������̳����ṩ��һ���������Լ�������<li><a target=_blank href=Default.asp?id="&id&">"&SiteSettings("SiteURL")&"Default.asp?id="&id&"</a><li> ���� <a target=_blank href=Default.asp?id="&id&">"&Request("ForumName")&"</a> �ϵ���������Ҽ����������������������ǩ�����ղؼ���"
succeed(""&Message&"")


end if


%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� ������̳</td>
</tr>
</table><br>

<center>
<form name="form" method="POST" action="ApplyForum.asp">
<input type=hidden name=menu value="Apply">


<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2>
<tr>
<td height="20" align="center" colspan="2" class=a1><b>������̳</b></td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="20%">��̳���ƣ�</td>
<td class=a3 height="5" align="Left" valign="middle" width="78%">
&nbsp;
<input type="text" name="ForumName" size="30" maxlength="12">
</td>
</tr>

<tr class=a4>
<td height="2" align="right" width="20%">��̳���ܣ�</td>
<td height="2" align="Left" valign="middle" width="78%" class=a3>
&nbsp;
<textarea name="ForumIntro" rows="4" cols="50"></textarea>&nbsp;
</td>
</tr>
<tr class=a3>
<td height="2" align="right" width="20%">��̳����</td>
<td height="2" align="Left" valign="middle" width="78%" class=a3>
&nbsp;
<textarea name="ForumRules" rows="4" cols="50"></textarea>&nbsp;
</td>
</tr>
<tr class=a4>
<td height="1" align="right" valign="middle" width="20%">Сͼ��URL��</td>
<td height="1" align="Left" valign="middle" width="78%">
&nbsp;
<input size="30" name="icon">����ʾ����̳�����ұ�</td>
</tr>
<tr class=a3>
<td align="right" valign="bottom" width="20%">��ͼ��URL��</td>
<td align="Left" valign="bottom" width="78%">
&nbsp;
<input size="30" name="ForumLogo">����ʾ����̳���Ͻ�</td>
</tr>
<tr class=a3>
<td align="right" width="100%" colspan="2">
<input type="submit" value=" ȷ �� &gt;&gt;�� һ �� "> </td>
</tr>
</table>

</form>
<%
htmlend
%>