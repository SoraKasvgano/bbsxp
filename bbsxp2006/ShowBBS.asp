<!-- #include file="Setup.asp" -->
<%
UserName=HTMLEncode(Request("UserName"))
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
%>


<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <span id=bbsxpname></span></td>
</tr>
</table><br>

<table cellspacing="1" cellpadding="1" width="100%" align="center" border="0" class="a2">
  <TR id=TableTitleLink class=a1 height="25">
      <Td width="14%" align="center"><a href="?">��������</a></td>
      <TD width="14%" align="center"><a href="?menu=1">������������</a></td>
      <TD width="14%" align="center"><a href="?menu=2">������������</a></td>
      <TD width="14%" align="center"><a href="?menu=3">��������</a></td>
      <TD width="14%" align="center"><a href="?menu=4">ͶƱ����</a></td>
      <TD width="14%" align="center"><a href="?menu=5">�ҵ�����</a></td>
      </TR></TABLE>


<br>


<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">����</td>
<td align="center" width="10%">����</td>
<td align="center" width="6%">�ظ�</td>
<td align="center" width="6%">���</td>
<td align="center" width="25%">������</td>
</tr>
<%
select case Request("menu")
case ""
sql="select top 200 * from [BBSXP_Threads] where IsDel=0 order by id Desc"
%><SCRIPT>bbsxpname.innerText="��������"</SCRIPT><%
case "1"
sql="select top 200 * from [BBSXP_Threads] where IsDel=0 and PostTime>"&SqlNowString&"-7 order by Views Desc"
%><SCRIPT>bbsxpname.innerText="������������"</SCRIPT><%
case "2"
sql="select top 200 * from [BBSXP_Threads] where IsDel=0 and PostTime>"&SqlNowString&"-7 order by replies Desc"
%><SCRIPT>bbsxpname.innerText="������������"</SCRIPT><%
case "3"
sql="select top 200 * from [BBSXP_Threads] where IsGood=1 and IsDel=0 order by id Desc"
%><SCRIPT>bbsxpname.innerText="��������"</SCRIPT><%
case "4"
sql="select top 200 * from [BBSXP_Threads] where IsVote=1 and IsDel=0 order by id Desc"
%><SCRIPT>bbsxpname.innerText="ͶƱ����"</SCRIPT><%
case "5"
sql="select top 200 * from [BBSXP_Threads] where UserName='"&UserName&"' and IsDel=0 order by id Desc"
%><SCRIPT>bbsxpname.innerText="<%=UserName%> ������"</SCRIPT><%
end select

Rs.Open sql,Conn,1

PageSetup=SiteSettings("ThreadsPerPage") '�趨ÿҳ����ʾ����
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1
ShowThread()
Rs.MoveNext
loop
Rs.Close
%></table>

<table border=0 width=100% align=center><tr><td>
<%ShowPage()%></td></tr></table>
<%
htmlend
%>