<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
UserName=Trim(HTMLEncode(Request("UserName")))

if Request.Form("menu")="ok" then

if membercode < 4 then error("<li>����Ȩ�޲������޷�ץ��������")

If Conn.Execute("Select UserName From [BBSXP_Users] where UserName='"&UserName&"'" ).eof Then error("<li>"&UserName&"�����ϲ�����")
If not Conn.Execute("Select UserName From [BBSXP_Prison] where UserName='"&UserName&"'" ).eof Then error("<li>"&UserName&"�Ѿ����ؽ�����")


if UserName="" then error("<li>���˵�����û����д")


causation=HTMLEncode(Request("causation"))
PrisonDay=HTMLEncode(Request("PrisonDay"))

if causation="" then error("<li>��û����������ԭ��")
if PrisonDay>1000 then error("<li>����ʱ�䲻�ܳ���1000��")

sql="insert into [BBSXP_Prison] (UserName,causation,constable,PrisonDay) values ('"&UserName&"','"&causation&"','"&CookieUserName&"','"&PrisonDay&"')"
Conn.Execute(SQL)



end if

if Request("menu")="release" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")

if membercode < 4 then error("<li>����Ȩ�޲������޷��ͷŷ��ˣ�")

Conn.execute("Delete from [BBSXP_Prison] where UserName='"&UserName&"'")

Log("�� "&UserName&" �ͷų�������")

error2("�Ѿ��� "&UserName&" �ͷų�������")
end if


if Request("menu")="look" then
sql="select * from [BBSXP_Prison] where UserName='"&UserName&"'"
Set Rs=Conn.Execute(sql)


%>
<title>̽ �� - Powered By BBSXP</title>
<b><%=Rs("UserName")%></b>
<SCRIPT>
var tips=["б���۾����һ�ۿ���,�����:����ĵ��̫��!","����������˵:�����Ҳ���!�Բ�������!","����¶�������Ц��:�ٺ�!Ҫ��Ҫ��������!","�п���ֵ�:һʧ��,��ǧ�ź�!��һ����������!","���Ų�����������˿���ĸ�ǽ,ҡͷ̾Ϣ��!"]
index = Math.floor(Math.random() * tips.length);
document.write("" + tips[index] + "");
  </SCRIPT><br><br>
����ԭ��<%=Rs("causation")%><br><br>
����ʱ�䣺<%=Rs("cometime")%><br><br>
����ʱ�䣺<%=Rs("cometime")+Rs("PrisonDay")%><br><br>
ִ�о��٣�<%=Rs("constable")%>

<%
Rs.close
CloseDatabase

end if

top

Conn.execute("Delete from [BBSXP_Prison] where cometime<"&SqlNowString&"-PrisonDay")
%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
<a href="Prison.asp">��������</a></td>
</tr>
</table><br>



<TABLE cellSpacing=1 cellPadding=3 border=0 width=100% align=center class=a2><TR class=a1 height="25">
	<TD align=center width="10%"><b>�û���</b></TD>
<TD align=center><b>����ԭ��</b></TD>
<TD align=center width="20%"><b>����ʱ��</b></TD>
<TD align=center width="10%"><b>ִ�о���</b></TD>
<TD align=center width="15%"><b>����</b></TD>
</TR>

<%
Rs.Open "[BBSXP_Prison] order by cometime Desc",Conn,1
PageSetup=20 '�趨ÿҳ����ʾ����
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��
i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1


response.write "<tr class=a4><TD align=center><a href=Profile.asp?UserName="&Rs("UserName")&">"&Rs("UserName")&"</a></TD><TD>"&Rs("causation")&"</TD><TD align=center>"&Rs("cometime")&"</TD><TD align=center><a href=Profile.asp?UserName="&Rs("constable")&">"&Rs("constable")&"</a></TD><TD align=center><a href=?menu=release&sessionid="&session.sessionid&"&UserName="&Rs("UserName")&">�� ��</a> | <a href=# onClick=javascript:open('Prison.asp?menu=look&UserName="&Rs("UserName")&"','','resizable,scrollbars,width=220,height=180')>̽ ��</a></TD></tr>"


Rs.MoveNext
loop
Rs.Close

%>

<form METHOD=Post><input type=hidden name=menu value=ok><tr>
  <TD align=center colspan="5" class=a3>��
<input name="UserName" size="13"> ץ�������������������<input name="PrisonDay" size="2" value="15"><br>
����ԭ��<input name="causation" size="33"> <input type="submit" value=" ȷ �� "></TD>
  			</tr></form>
</TABLE>

<table border=0 width=100% align=center><tr><td>
<%ShowPage()%>
</tr></td></table>



<%
htmlend
%>