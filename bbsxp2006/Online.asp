<!-- #include file="Setup.asp" --><%
top
count=Conn.execute("Select count(sessionid)from [BBSXP_UsersOnline]")(0)
regOnline=Conn.execute("Select count(sessionid)from [BBSXP_UsersOnline] where UserName<>''")(0)
ForumThreads=Conn.execute("Select SUM(ForumThreads)from [BBSXP_Forums]")(0)
tolReTopic=Conn.execute("Select SUM(ForumPosts)from [BBSXP_Forums]")(0)
%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		�� �鿴��̳״̬</td>
	</tr>
</table>
<br>
<table cellspacing="1" cellpadding="4" width="100%" align="center" border="0" class="a2">
	<tr class="a1" id="TableTitleLink">
		<td width="16%" align="center" height="10"><b><a href="Online.asp">�������</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=cutline">����ͼ��</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=UserSex">�Ա�ͼ��</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=TodayPage">����ͼ��</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=board">����ͼ��</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=ForumPosts">����ͼ��</a></b></td>
	</tr>
	<tr class="a3">
		<td width="48%" align="center" height="10" colspan="3">������ <%=tolReTopic%> 
		ƪ���������� <%=ForumThreads%> ƪ������ <%=tolReTopic-ForumThreads%> ƪ��</td>
		<td width="48%" align="center" height="10" colspan="3">������ <%=count%> �ˡ�����ע���û� 
		<%=regOnline%> �ˣ��ÿ� <%=count-regOnline%> �ˡ�</td>
	</tr>
</table>
<br>
<%

select case Request("menu")
case ""
index
case "cutline"
cutline

case "board"
board

case "ForumPosts"
ForumPosts

case "TodayPage"
TodayPage

case "UserSex"
UserSex

end select

sub index
Key=HTMLEncode(Request.Form("Key"))
Find=HTMLEncode(Request.Form("Find"))
if Key<>empty then SqlFind=" where "&Find&"='"&Key&"' and eremite=0"
sql="select * from [BBSXP_UsersOnline] "&SqlFind&" order by lasttime Desc"
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




if membercode<4 then
ips=split(Rs("ip"),".")
ShowIP=""&ips(0)&"."&ips(1)&".*.*"
else
ShowIP=""&Rs("ip")&""
end if


if Rs("UserName")="" then
UserName="�� ��"
elseif Rs("eremite")=1 and membercode<4 then
UserName="�� ��"
else
UserName="<a href=Profile.asp?UserName="&Rs("UserName")&">"&Rs("UserName")&"</a>"
end if



place2=""
if Rs("act")<>"" then
place2 = "<a href="&Rs("acturl")&">"&Rs("act")&"</a>"
place = "�� "&Rs("ForumName")&" ��"
else
place = "�� <a href="&Rs("acturl")&">"&Rs("ForumName")&"</a> ��"
end if
allline=""&allline&"<TR align=middle class=a4><TD height=24>"&ShowIP&"</TD><TD height=24>"&Rs("cometime")&"</TD><TD height=24>"&UserName&"</TD><TD height=24>"&place&"</TD><TD height=24>"&place2&"</TD><TD height=24>"&Rs("lasttime")&"</TD></TR>"

Rs.Movenext
loop
Rs.close



%>
<table cellspacing="1" cellpadding="1" width="100%" align="center" border="0" class="a2">
	<tr align="middle" class="a1" height="23">
		<td>IP��ַ</td>
		<td>��¼ʱ��</td>
		<td>�û���</td>
		<td>������̳</td>
		<td>��������</td>
		<td>�ʱ��</td>
	</tr>
	<%=allline%>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td valign="top"><%ShowPage()%></td>

		<td align="right">
		<form action="Online.asp" method="POST">
<select name=Find>
<option value="UserName">��ѯ�û�</option>
<option value="IP">��ѯ�ɣ�</option>
</select>		
<input size="15" value="<%=Key%>" name="Key">
			<input type="submit" value=" ȷ�� "></form>
		</td>
	</tr>
</table>
<%


end sub


sub cutline


sql="select * from [BBSXP_Forums] where followid<>0 and ForumHide=0"
Set Rs=Conn.Execute(sql)



%>
<table class="a2" cellspacing="1" cellpadding="4" width="100%" align="center" border="0">
	<tr>
		<td class="a1" valign="bottom" align="middle" height="20">��̳����</td>
		<td class="a1" valign="bottom" align="middle" height="20">ͼ�α���</td>
		<td class="a1" valign="bottom" align="middle" height="20">��������</td>
	</tr>
	<%

i=0
Do While Not Rs.EOF
Onlinemany=Conn.execute("Select count(sessionid)from [BBSXP_UsersOnline] where ForumID="&Rs("id")&"")(0)

%>
	<tr class="a3">
		<td width="21%" height="2" align="center">
		<a href="ShowForum.asp?ForumID=<%=Rs("id")%>"><%=Rs("ForumName")%></a></td>
		<td width="65%" height="2">
		<img height="8" src="images/bar/<%=i%>.gif" width="<%=Onlinemany/count*100%>%"></td>
		<td align="center" width="12%" height="2"><%=Onlinemany%></td>
	</tr>
	<%
			
i=i+1
if i=10 then i=0

Rs.MoveNext
loop
Rs.Close   


%></table>
<%


end sub

sub board

sql="select * from [BBSXP_Forums] where followid<>0 and ForumHide=0 order by ForumThreads Desc"
Set Rs=Conn.Execute(sql)
%>
<table class="a2" cellspacing="1" cellpadding="4" width="100%" align="center" border="0">
	<tr>
		<td class="a1" valign="bottom" align="middle" height="20">��̳����</td>
		<td class="a1" valign="bottom" align="middle" height="20">ͼ�α���</td>
		<td class="a1" valign="bottom" align="middle" height="20">������</td>
	</tr>
	<%
i=0
Do While Not Rs.EOF
%>
	<tr class="a3">
		<td width="21%" height="2" align="center">
		<a href="ShowForum.asp?ForumID=<%=Rs("id")%>"><%=Rs("ForumName")%></a></td>
		<td width="65%" height="2">
		<img height="8" src="images/bar/<%=i%>.gif" width="<%=Rs("ForumThreads")/ForumThreads*100%>%"></td>
		<td align="center" width="12%" height="2"><%=Rs("ForumThreads")%></td>
	</tr>
	<%


i=i+1
if i=10 then i=0


			
Rs.MoveNext
loop
Rs.Close   


%></table>
<%
end sub

sub ForumPosts

sql="select * from [BBSXP_Forums] where followid<>0 and ForumHide=0 order by ForumPosts Desc"
Set Rs=Conn.Execute(sql)
%>
<table class="a2" cellspacing="1" cellpadding="4" width="100%" align="center" border="0">
	<tr>
		<td class="a1" valign="bottom" align="middle" height="20">��̳����</td>
		<td class="a1" valign="bottom" align="middle" height="20">ͼ�α���</td>
		<td class="a1" valign="bottom" align="middle" height="20">������</td>
	</tr>
	<%
i=0
Do While Not Rs.EOF
%>
	<tr class="a3">
		<td width="21%" height="2" align="center">
		<a href="ShowForum.asp?ForumID=<%=Rs("id")%>"><%=Rs("ForumName")%></a></td>
		<td width="65%" height="2">
		<img height="8" src="images/bar/<%=i%>.gif" width="<%=Rs("ForumPosts")/tolReTopic*100%>%"></td>
		<td align="center" width="12%" height="2"><%=Rs("ForumPosts")%></td>
	</tr>
	<%
			
i=i+1
if i=10 then i=0

			
Rs.MoveNext
loop
Rs.Close   


%></table>
<%
end sub



sub TodayPage
tolForumToday=Conn.execute("Select SUM(ForumToday)from [BBSXP_Forums]")(0)
if tolForumToday=0 then tolForumToday=1
sql="select * from [BBSXP_Forums] where followid<>0 and ForumHide=0 order by ForumToday Desc"
Set Rs=Conn.Execute(sql)
%>
<table class="a2" cellspacing="1" cellpadding="4" width="100%" align="center" border="0">
	<tr>
		<td class="a1" valign="bottom" align="middle" height="20">��̳����</td>
		<td class="a1" valign="bottom" align="middle" height="20">ͼ�α���</td>
		<td class="a1" valign="bottom" align="middle" height="20">��������</td>
	</tr>
	<%
i=0
Do While Not Rs.EOF
%>
	<tr class="a3">
		<td width="21%" height="2" align="center">
		<a href="ShowForum.asp?ForumID=<%=Rs("id")%>"><%=Rs("ForumName")%></a></td>
		<td width="65%" height="2">
		<img height="8" src="images/bar/<%=i%>.gif" width="<%=Rs("ForumToday")/tolForumToday*100%>%"></td>
		<td align="center" width="12%" height="2"><%=Rs("ForumToday")%></td>
	</tr>
	<%
			
i=i+1
if i=10 then i=0

			
Rs.MoveNext
loop
Rs.Close   


%></table>
<%
end sub





sub UserSex
count=Conn.execute("Select count(id)from [BBSXP_Users]")(0)
male=Conn.execute("Select count(id)from [BBSXP_Users] where UserSex='male'")(0)
female=Conn.execute("Select count(id)from [BBSXP_Users] where UserSex='female'")(0)

%>
<table class="a2" cellspacing="1" cellpadding="4" width="100%" align="center" border="0">
	<tr>
		<td class="a1" valign="bottom" align="middle" height="20">�Ա�</td>
		<td class="a1" valign="bottom" align="middle" height="20">ͼ�α���</td>
		<td class="a1" valign="bottom" align="middle" height="20">����</td>
	</tr>
	<tr class="a3">
		<td width="10%" height="2" align="center"><img src="images/male.gif"></td>
		<td width="75%" height="2">
		<img height="8" src="images/bar/7.gif" width="<%=male/count*100%>%"></td>
		<td align="center" width="12%" height="2"><%=male%></td>
	</tr>
	<tr class="a3">
		<td width="10%" height="2" align="center"><img src="images/female.gif"></td>
		<td width="75%" height="2">
		<img height="8" src="images/bar/0.gif" width="<%=female/count*100%>%"></td>
		<td align="center" width="12%" height="2"><%=female%></td>
	</tr>
	<tr class="a3">
		<td width="10%" height="2" align="center">δ֪</td>
		<td width="75%" height="2">
		<img height="8" src="images/bar/2.gif" width="<%=(count-male-female)/count*100%>%"></td>
		<td align="center" width="12%" height="2"><%=count-male-female%></td>
	</tr>
</table>
<%
end sub


htmlend
%>