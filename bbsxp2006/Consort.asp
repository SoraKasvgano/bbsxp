<!-- #include file="Setup.asp" --><%
id=RequestInt("id")

content=HTMLEncode(Request.Form("content"))

top

if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")


Consort=Conn.Execute("Select Consort From [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'")(0)



select case Request("menu")
case "add"
aim=HTMLEncode(Request("aim"))
if content=empty then error("<li>�������ݲ���Ϊ��")
if aim=CookieUserName then error("<li>�����Լ�׷���Լ���")

If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&aim&"'" ).eof Then error("<li>ϵͳ������"&aim&"������")

If Not Conn.Execute("Select id From [BBSXP_Consort] where UserName='"&CookieUserName&"' and aim='"&aim&"'" ).eof Then error("<li>"&aim&"�Ѿ�������׷���б�����")

sql="insert into [BBSXP_Consort] (UserName,aim,unburden) values ('"&CookieUserName&"','"&aim&"','"&content&"')"
Conn.Execute(SQL)


sql="insert into [BBSXP_Messages](UserName,incept,content) values ('"&CookieUserName&"','"&aim&"','<font color=0000FF>��������ż����"&content&"</font>')"
Conn.Execute(SQL)


Conn.execute("update [BBSXP_Users] set NewMessage=NewMessage+1 where UserName='"&aim&"'")

case "accept"

if Consort<>empty then error("<li>����ǰ������ż")

aim=Conn.Execute("Select aim From [BBSXP_Consort] where id="&id&"")(0)
if aim<>CookieUserName then error("<li>�Ƿ�����")

Consort=Conn.Execute("Select UserName From [BBSXP_Consort] where id="&id&"")(0)
if Conn.Execute("Select Consort From [BBSXP_Users] where UserName='"&Consort&"'")(0)<>empty then error("<li>"&Consort&"�Ѿ�����ż��")

Conn.execute("update [BBSXP_Users] set Consort='"&aim&"' where UserName='"&Consort&"'")
Conn.execute("update [BBSXP_Users] set Consort='"&Consort&"' where UserName='"&aim&"'")
Conn.execute("Delete from [BBSXP_Consort] where id="&id&"")
succeed("<li>���Ѿ�������"&Consort&"��׷��<li><a href=Consort.asp>����������ż</a><meta http-equiv=refresh content=3;url=Consort.asp>")



case "Del"
Conn.execute("Delete from [BBSXP_Consort] where id="&id&" and (UserName='"&CookieUserName&"' or aim='"&CookieUserName&"') ")

case "part"
Conn.execute("update [BBSXP_Users] set Consort='' where UserName='"&Consort&"'")
Conn.execute("update [BBSXP_Users] set Consort='' where UserName='"&CookieUserName&"'")
succeed("<li>���ֳɹ�<li><a href=Consort.asp>����������ż</a><meta http-equiv=refresh content=3;url=Consort.asp>")

end select

%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		�� <a href="Consort.asp">������ż</a></td>
	</tr>
</table>
<br>
<center>

<table cellspacing="1" cellpadding="6" width="100%" border="0" class="a2">
	<tr>
		<td width="100%" height="10" align="middle" class="a1" colspan="4">
		<font face="����">�ҵ�׷����</font></td>
	</tr>
	<tr class="a3">
		<td width="10%" height="5" align="middle">�û���</td>
		<td height="5" align="middle">���ı���</td>
		<td width="20%" height="5" align="middle"><font face="����">׷��</font>ʱ��</td>
		<td width="15%" height="5" align="middle">����</td>
	</tr>
	<%
sql="select * from [BBSXP_Consort] where aim='"&CookieUserName&"' order by id Desc"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF 
%>
	<tr class="a4">
		<td height="5" align="middle">
		<a href="Profile.asp?UserName=<%=Rs("UserName")%>"><%=Rs("UserName")%></a></td>
		<td height="5" align="middle"><%=Rs("unburden")%></td>
		<td height="5" align="middle"><%=Rs("DateCreated")%></td>
		<td height="5" align="middle"><%if Consort=Rs("UserName") then%>
		<a href="?menu=part">�� ��</a> <%else%>
		<a href="?menu=accept&id=<%=Rs("id")%>">����</a>
		<a href="?menu=Del&id=<%=Rs("id")%>">�ܾ�</a> <%end if%></td>
	</tr>
	<%
Rs.MoveNext
loop
Rs.Close
%>
</table>

<br>

<table cellspacing="1" cellpadding="6" width="100%" border="0" class="a2">
	<tr class="a1">
		<td width="100%" height="10" align="middle" colspan="4">��׷�����</td>
	</tr>
	<tr class="a3">
		<td width="10%" height="5" align="middle">�û���</td>
		<td height="5" align="middle">���ı���</td>
		<td width="20%" height="5" align="middle"><font face="����">׷��</font>ʱ��</td>
		<td width="15%" height="5" align="middle">����</td>
	</tr>
	<%
sql="select * from [BBSXP_Consort] where UserName='"&CookieUserName&"' order by id Desc"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF 
%>
	<tr class="a4">
		<td height="5" align="middle">
		<a href="Profile.asp?UserName=<%=Rs("aim")%>"><%=Rs("aim")%></a></td>
		<td height="5" align="middle"><%=Rs("unburden")%></td>
		<td height="5" align="middle"><%=Rs("DateCreated")%></td>
		<td height="5" align="middle"><%if Consort=Rs("aim") then%>
		<a href="?menu=part">�� ��</a> <%else%>
		<a href="?menu=Del&id=<%=Rs("id")%>">ȡ��׷��</a> <%end if%> </td>
	</tr>
	<%
Rs.MoveNext
loop
Rs.Close

%>
</table>


<br>

<%if Consort<>empty then

sql="select * from [BBSXP_Users] where UserName='"&Consort&"'"
Set Rs=Conn.Execute(sql)

if Rs.eof then Conn.execute("update [BBSXP_Users] set Consort='' where UserName='"&CookieUserName&"'")

select case Rs("UserSex")
case "male"
UserSex="��"
case "female"
UserSex="Ů"
end select

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
Userphoto=Rs("Userphoto")
Rs.close
%>
<table cellspacing="1" cellpadding="6" width="100%" border="0" class="a2">
	<tr class="a1" id="TableTitleLink">
		<td align="middle" class="a1" colspan="7">�ҵ���ż</td>
	</tr>
	<tr>
		<td width="20%" align="Left" class="a4" rowspan="3"><script>
if("<%=Userphoto%>"!=""){
document.write("<img src=<%=Userphoto%> border=0 onload='javascript:if(this.width>200)this.width=200'>")
}
</script>
		
		</td>
		<td width="80" align="Left" class="a4">�ǳƣ�</td>
		<td align="Left" class="a4" width="15%"><a href="Profile.asp?UserName=<%=Consort%>">
		<%=Consort%></a></td>
		<td width="80" align="Left" class="a4">������</td>
		<td align="Left" class="a4" width="15%"><%=realname%></td>
		<td width="80" align="Left" class="a4">�Ա�</td>
		<td align="Left" class="a4" width="15%"><%=UserSex%></td>
	</tr>
	<tr>
		<td width="80" align="Left" class="a4" valign="top">���ң�</td>
		<td align="Left" class="a4" width="15%"><%=country%></td>
		<td width="80" align="Left" class="a4">ʡ�ݣ�</td>
		<td align="Left" class="a4" width="15%"><%=province%></td>
		<td width="80" align="Left" class="a4">���У�</td>
		<td align="Left" class="a4" width="15%"><%=city%></td>
	</tr>
	<tr>
		<td width="80" align="Left" class="a4" valign="top">����˵����</td>
		<td align="Left" class="a4" colspan="5"><%=personal%></td>
	</tr>
	<tr>
		<td align="right" class="a4" valign="top" colspan="7">
		<a onclick="checkclick('��ȷ��Ҫ�뵱ǰ��ż���֣�')" href="?menu=part">�뵱ǰ��ż����</a></td>
	</tr>
</table>
<%else%>
<table cellspacing="1" cellpadding="6" width="100%" border="0" class="a2">
	<form action="Consort.asp" method="POST">
		<input type="hidden" value="add" name="menu">
		<tr>
			<td width="77%" height="2" align="middle" class="a1" colspan="2">��������׷�����</td>
		</tr>
		<tr>
			<td width="12%" height="2" align="Left" class="a4">�Է��û�����</td>
			<td width="64%" height="2" align="Left" class="a4">
			<input maxlength="30" size="15" name="aim"></td>
		</tr>
		<tr>
			<td width="12%" height="1" align="Left" class="a4" valign="top">���ı��ף�</td>
			<td width="64%" height="1" align="Left" class="a4">
			<textarea name="content" rows="5" style="width:95%"></textarea></td>
		</tr>
		<tr>
			<td width="77%" height="1" align="center" class="a4" colspan="2">
			<input type="submit" value=" ȷ �� ">
			<input onclick="checkclick('�������Ҫ���ȫ�������ݣ���ȷ��Ҫ�����?');" type="reset" value=" �� д "></td>
		</tr>
	</form>
</table>
<%end if%>

</center><%htmlend%>