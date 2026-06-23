<!-- #include file="Setup.asp" --><%

UserName=HTMLEncode(Request("UserName"))
id=RequestInt("id")

top

if id<>empty then
sql="select * from [BBSXP_Calendar] where id="&id&" order by id Desc"
else
sql="select * from [BBSXP_Calendar] where (hide=0 or UserName='"&SqlString(CookieUserName)&"') and UserName='"&SqlString(UserName)&"' order by id Desc"
end if

Rs.Open sql,Conn,1

if Rs.eof then error("<li>���û���ʱû���κ���־")

%>
<table class="a2" cellSpacing="1" cellPadding="4" width="100%" align="center" border="0">
	<tr class="a3">
		<td colSpan="2">
		<table cellSpacing="0" cellPadding="0" width="100%" border="0" id="table2">
			<tr>
				<td height="18">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> �� <a href="Calendar.asp">������־</a> �� <span id=title><a href="Blog.asp">������־</a></span></td>
				<td align="right" height="18">
				<img src="images/jt.gif"> <a href=Calendar.asp?menu=NewCalendar>������־</a></td>
			</tr>
		</table>
		</td>
	</tr>
</table>

<br>
<table border="0" width="100%" align="center">
	<tr>
		<td>
		<table class="a2" style="WIDTH:100%" height="100%" cellspacing="1" cellpadding="3" border="0">
			<tr>
				<td class="a1" align="middle" height="25"><b>������־</b></td>
			</tr>
			<tr align="middle" class=a3>
				<td align="Left" height="100">
<%




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
UserName=Rs("UserName")
content=Rs("content")

if id="" then
content=ReplaceText(content,"<[^>]*>","")
if len(content)>200 then content=Left(""&content&"",200)&"..."
end if


%>
				<table border="0" width="100%" cellspacing="10" style=TABLE-LAYOUT:fixed>
					<tr>
						<td style=word-break:break-all>
						<b><font size="4"><%=Rs("title")%></font></b><br>
						<br><%=content%>
<%if Rs("hide")=1 then%><br><br>ע��<font color="#FF0000">��ƪ��־Ϊ����״̬</font><%end if%>
<hr>
<%if id="" then%>
<a href="?id=<%=Rs("id")%>">�Ķ�ȫ��</a>
<%else%>
<font color="#C0C0C0">�Ķ�ȫ��</font>
<%end if%>
 | 
<%if CookieUserName=UserName then%>
<a href="Calendar.asp?menu=NewCalendar&id=<%=Rs("id")%>">�༭</a>
<%else%>
<font color="#C0C0C0">�༭</font>
<%end if%>
 | 
<%if UserName=CookieUserName or membercode > 3 then%>
<a href="Calendar.asp?menu=del&id=<%=Rs("id")%>" onclick="checkclick('��ȷ��Ҫɾ��������־?')">ɾ��</a>
<%else%>
<font color="#C0C0C0">ɾ��</font>
<%end if%>
 | <%=Rs("DateCreated")%> by <a href="Profile.asp?UserName=<%=UserName%>"><%=UserName%></a></td>
					</tr>
				</table>
<%
Rs.MoveNext
loop

Rs.Close  
%>
				</td>
			</tr>
		</table>
<%ShowPage()%>
		</td>
		<td width="200" align="right" valign="top">
		<table class="a2" style="WIDTH:95%" cellspacing="1" cellpadding="3" border="0">
			<tr class="a1">
				<td align="middle" height="25"><b>�����ļ�</b></td>
			</tr>
			<tr align="middle" class=a3>
				<td>
<%
sql="select * from [BBSXP_Users] where UserName='"&UserName&"'"
Set Rs=Conn.Execute(sql)

select case Rs("UserSex")
case "male"
UserSex="��"
case "female"
UserSex="Ů"
end select

if Rs("birthday")<>"" then birthyear=year(now)-split(Rs("birthday"),"-")(0)

Userphoto=Rs("Userphoto")
UserInfo=split(Rs("UserInfo"),"\")
realname=UserInfo(0)
country=UserInfo(1)
province=UserInfo(2)
city=UserInfo(3)
Postcode=UserInfo(4)
blood=UserInfo(5)
belief=UserInfo(6)
occupation=UserInfo(7)
marital=UserInfo(8)
education=UserInfo(9)
college=UserInfo(10)
address=UserInfo(11)
phone=UserInfo(12)
character=UserInfo(13)
personal=UserInfo(14)
Rs.close
%>
<script>
if("<%=Userphoto%>"!=""){
document.write("<a target=_blank href=<%=Userphoto%>><img src=<%=Userphoto%> border=0 width=150></a>")
}
</script>
				<table cellspacing="2" cellpadding="2" width="100%" border="0">
					<tr>
						<td valign="top">������</td>
						<td width="120" ><%=UserName%></td>
					</tr>
					<tr>
						<td valign="top">���䣺</td>
						<td width="120"><%=birthyear%>
						</td>
					</tr>
					<tr>
						<td valign="top">�Ա�</td>
						<td width="120"><%=UserSex%></td>
					</tr>
					<tr>
						<td valign="top">ְҵ��</td>
						<td width="120"><%=occupation%></td>
					</tr>
					<tr>
						<td valign="top"><font class="bold">λ�ã�</font></td>
						<td width="120"><%=country%><br>
						<%=province%><br>
						<%=city%></td>
					</tr>
					<tr>
						<td colspan="2"><%=personal%></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr align="middle" class=a3>
				<td align="right">
				<a href="Profile.asp?UserName=<%=UserName%>">�鿴�����ļ���ϸ��Ϣ</a></td>
			</tr>
		</table>
		<br>
		<table class="a2" style="WIDTH:95%" cellspacing="1" cellpadding="3" border="0">
			<tr>
				<td class="a1" align="middle" height="25"><b>����������־</b></td>
			</tr>
			<tr class=a3>
				<td>
<%
sql="select top 10 * from [BBSXP_Calendar] where hide=0 and UserName='"&UserName&"' order by id Desc"

Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF
%><li><a href=?id=<%=Rs("id")%>><%=Rs("title")%></a> (<%=Rs("adddate")%>)</li><%
Rs.MoveNext
loop
Rs.Close   
%>
</td></tr>
		</table>
		��</td>
	</tr>
</table>

<script>title.innerHTML='<a href="Blog.asp?UserName=<%=UserName%>">��<%=UserName%>��������־</a>'</script>

<%

htmlend
%>