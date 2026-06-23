<!-- #include file="Setup.asp" -->
<title>������ - Powered By BBSXP</title>
<%
top
%>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
<a href="Affiche.asp">��������</a></td>
</tr>
</table>
<%
if Request("id")="" then
sql="select * from [BBSXP_Affiche] order by Posttime Desc"
else
sql="select * from [BBSXP_Affiche] where id="&RequestInt("id")&""
end if


Rs.Open sql,Conn,1
PageSetup=10 '�趨ÿҳ����ʾ����
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��
i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1
%><br>
<table border="0" width="100%" cellspacing="1" cellpadding="5" height="64" class="a2" align=center>
	<tr>
		<td width="100%" height="20" class="a1">
		<p align="center"><%=Rs("title")%></p>
		</td>
	</tr>
	<tr>
		<td width="100%" class="a3"><%=Rs("content")%>
		</td>
	</tr>
	<tr>
		<td width="100%" class="a4" height="18">
		<p align="right">������ <%=Rs("UserName")%>������ʱ�� 
		<font style="family:arial; font-size: 7pt"><%=Rs("Posttime")%></font> </p>
		</td>
	</tr>
</table>
<%
Rs.MoveNext
loop
Rs.Close
%>
<table border=0 width=100% align=center><tr><td>
<%ShowPage()%>
</tr></td></table>

<%
htmlend
%>