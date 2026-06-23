<!-- #include file="Setup.asp" --><%

if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")

id=int(Request("id"))
url=HTMLEncode(Request("url"))
name=HTMLEncode(Request("name"))

top


select case Request("menu")
case "add"

If not Conn.Execute("Select id From [BBSXP_Favorites] where UserName='"&SqlString(CookieUserName)&"' and name='"&name&"' and url='"&url&"'" ).eof Then error("<li>�ղؼ����Ѿ����ڴ�����")
Conn.execute("insert into [BBSXP_Favorites](UserName,name,url)values('"&CookieUserName&"','"&name&"','"&url&"')")
Message="<li>���ӳɹ�<li><a href="&Request.ServerVariables("http_referer")&">"&Request.ServerVariables("http_referer")&"</a><li><a href=Default.asp>����������ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url="&Request.ServerVariables("http_referer")&">")


case "Delweb"
Conn.execute("Delete from [BBSXP_Favorites] where UserName='"&SqlString(CookieUserName)&"' and id="&id&"")
Message="<li>ɾ���ɹ�<li><a href="&Request.ServerVariables("http_referer")&">"&Request.ServerVariables("http_referer")&"</a><li><a href=Default.asp>����������ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url="&Request.ServerVariables("http_referer")&">")


case "Del"
Conn.execute("Delete from [BBSXP_Favorites] where UserName='"&SqlString(CookieUserName)&"' and url='"&url&"' and name='"&name&"'")
Message="<li>ɾ���ɹ�<li><a href="&Request.ServerVariables("http_referer")&">"&Request.ServerVariables("http_referer")&"</a><li><a href=Default.asp>����������ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url="&Request.ServerVariables("http_referer")&">")


end select

%>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
�ղؼ�</td>
</tr>
</table><br>


<%

PageSetup=10 '�趨ÿҳ����ʾ����

select case Request("menu")
case ""
%>

<table width="100%" border=0 align=center cellPadding=3 cellSpacing=1 class=a2>
<tr class=a1><td width="69%" align="center" height="25"><b>�� ��</b></td>
	<td width="20%" align="center" height="25">
		<b>����ʱ��</b></td>
	<td width="11%" align="center" height="25"><b>����</b></td></tr><%

sql="select * from [BBSXP_Favorites] where UserName='"&SqlString(CookieUserName)&"' and url<>'Topic' and url<>'forum' order by id Desc"
Rs.Open sql,Conn,1

Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��

i=0 
Do While Not Rs.EOF and i<PageSetup 
i=i+1 

%> <tr class=a3><td><img src=images/ie.gif> <a target=_blank href="<%=Rs("url")%>"><%=Rs("name")%></a></td><td align=center><%=Rs("DateCreated")%><td align=center><a href="?menu=Delweb&id=<%=Rs("id")%>"><img src=images/Del.gif border=0></a></td></tr><%          
Rs.MoveNext
loop
Rs.Close      
%> <tr><td colSpan="3"  class=a1 align="center" height="25"><b>&gt;&gt; ���������� &lt;&lt;</b></td></tr><tr class=a3><td align="center" colspan="3"><form method=Post name=form action=MyFavorites.asp><input type=hidden name=menu value=add>
<b>���ƣ�</b><INPUT size=20 name=name>��<b>���ӵ�ַ��</b><INPUT size=40 name=url value="http://">��<input type=submit value="�� ��"> </td></form></tr></table>


<%

case "Topic"
%>
<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">����</td>
<td align="center" width="10%">����</td>
<td align="center" width="6%">�ظ�</td>
<td align="center" width="6%">���</td>
<td align="center" width="25%">������</td>
</tr>
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql="select * from [BBSXP_Favorites] where UserName='"&SqlString(CookieUserName)&"' and url='Topic' order by id Desc"
Rs1.Open sql,Conn,1
PageSetup=SiteSettings("ThreadsPerPage") '�趨ÿҳ����ʾ����
Rs1.Pagesize=PageSetup
TotalPage=Rs1.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs1.absolutePage=PageCount '��ת��ָ��ҳ��


i=0
Do While Not RS1.EOF and i<pagesetup
i=i+1


if not isnumeric(Rs1("name")) then Conn.execute("Delete from [BBSXP_Favorites] where id="&Rs1("id")&"")

sql="Select * From [BBSXP_Threads] where id="&Rs1("name")&""
Set Rs=Conn.Execute(sql)
if Rs.eof then Conn.execute("Delete from [BBSXP_Favorites] where id="&Rs1("id")&"")
ShowThread()
Set Rs = Nothing


Rs1.MoveNext
loop
Rs1.Close

%>
</table>

<%


case "Forum"

%>
<table width="100%" align="center" border="0" class="a2" cellspacing=1>
<%
sql="select * from [BBSXP_Favorites] where UserName='"&SqlString(CookieUserName)&"' and url='forum' order by id Desc"
Rs.Open sql,Conn,1
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��


i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1

if not isnumeric(Rs("name")) then Conn.execute("Delete from [BBSXP_Favorites] where id="&Rs("id")&"")

sql="Select * From [BBSXP_Forums] where id="&Rs("name")&""
Set Rs1=Conn.Execute(sql)
if Rs1.eof then Conn.execute("Delete from [BBSXP_Favorites] where id="&Rs("id")&"")
ShowForum()
Set Rs1 = Nothing
Rs.MoveNext
loop
Rs.Close
%>
</table>
<%
end select
%>
<table border=0 width=100% cellspacing=0><tr><td>
<%ShowPage()%></td></tr></table>
<%
htmlend
%>