<!-- #include file="Setup.asp" -->
<%
top
%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
<a href="?">��Ա�б�</a> (TOP 200)</td>
</tr>
</table><br>



<TABLE cellSpacing=1 cellPadding=0 width=100% border=0 class=a2 align=center>
<TR align=middle id=TableTitleLink class=a1 height="25">
<TD><a href="?order=UserName">�û���</a></TD>
<TD>��ѶϢ</TD>
<TD><a href="?order=UserRegTime">ע��ʱ��</a></TD>
<TD><a href="?order=PostTopic">��������</a></TD>
<TD><a href="?order=Postrevert">�ظ�����</a></TD>
<TD><a href="?order=UserMoney">�������</a></TD>
<TD><a href="?order=experience">����ֵ</a></TD>
<TD><a href="?order=UserLandTime">����¼ʱ��</a></TD>
<TD><a href="?order=UserDegree">��¼����</a></TD>
</TR>
<%

order=Request("order")
if Len(order)>15 then error2("�Ƿ�����")
if Order<>"" then SqlOrder=" order by "&order&" Desc"

sql="select top 200 * from [BBSXP_Users] "&SqlOrder&""
Rs.Open sql,Conn,1


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



if Rs("Userhome")<>"http://" then
Userhome="<a href="&Rs("Userhome")&" target=_blank><img border=0 src=images/home.gif></a>"
else
Userhome=""
end if

%>

<TR align=middle height="25">
<TD class=a4><a href=Profile.asp?UserName=<%=Rs("UserName")%>><%=Rs("UserName")%></a></TD>
<TD class=a3><a style=cursor:hand onclick="javascript:open('Friend.asp?menu=Post&incept=<%=Rs("UserName")%>','','width=320,height=170')">
<img border="0" src="images/message.gif"></a></TD>
<TD class=a4><%=Rs("UserRegTime")%></TD>
<TD class=a3><%=Rs("PostTopic")%></TD>
<TD class=a4><%=Rs("Postrevert")%></TD>
<TD class=a3><%=Rs("UserMoney")%></TD>
<TD class=a4><%=Rs("experience")%></TD>
<TD class=a3><%=Rs("UserLandTime")%></TD>
<TD class=a3><%=Rs("UserDegree")%></TD></TR>
<%


Rs.MoveNext
loop
Rs.Close

%>
</TABLE>


<table border="0" width="100%" align=center>
		<tr>
			<td width="50%">
<%ShowPage()%>
			</td>
			<td width="50%" align="right">
<form action="Profile.asp" method="POST">��ѯ�û�����<input size="15" name="UserName"> <input type="submit" value=" ȷ�� "></td></form>
</td>
		</tr>
</table>
<%
htmlend
%>