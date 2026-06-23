<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
youip="127.0.0.1" '��������IP��ַ

if SiteSettings("AdminPassword")<>"" and Request.ServerVariables("REMOTE_ADDR")<>youip then error("<li>Ϊ�˰�ȫ�������༭ <font color=red>Install.asp</font> �ڱ�����IP��ַ<li>��������ó� <font color=red>"&Request.ServerVariables("REMOTE_ADDR")&"</font>")

top

if Request("menu")="ok" then
Administrators=Request.Form("Administrators")
if Administrators="" then error("<li>��û�����ù���Ա")
if Request("Adminpassword")="" then error("<li>��û�����ù�������")
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&Administrators&"'" ).eof Then error("<li>"&Administrators&"���û����ϻ�δ<a href=CreateUser.asp>ע��</a>")

Conn.execute("update [BBSXP_SiteSettings] set Adminpassword='"&SqlString(MD5(Request("Adminpassword")))&"'")
Conn.execute("update [BBSXP_Users] set membercode=5 where UserName='"&SqlString(Administrators)&"'")

Message=Message&"<li>��װ�ɹ�<li><a href=Admin.asp target=_top>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Default.asp>")
end if

%>
<title>��װBBSXP... - Powered By BBSXP</title><br><form method="POST"><input type="hidden" value="ok" name="menu">


<table cellpadding="1" cellspacing="1" border="0" class=a2 align=center width="380">
<tr class=a1>
<td align=center colspan="2" height="25"><p><b>����������������</b></p></td>
</tr>


<tr class=a3><td align=right width="30%" height="20">�� �� Ա��</td>
<td height="20">
<input name=Administrators size="30"></td></tr>
<tr class=a4><td align=right width="30%" height="10">�������룺</td>
<td height="10">
<input name=Adminpassword type=password size="30"></td></tr>
<tr class=a3><td align=center height="10" colspan="2"><input type="submit" value=" ��һ�� "></td>
</tr>
</table>
</FORM>
<center>ע��<a href="CreateUser.asp">�������Ա�û�����û��ע�������������ע��</a></center>
<%htmlend%>