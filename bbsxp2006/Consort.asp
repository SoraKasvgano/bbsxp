<!-- #include file="Setup.asp" --><%
id=int(Request("id"))

content=HTMLEncode(Request.Form("content"))

top

if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")


Consort=Conn.Execute("Select Consort From [BBSXP_Users] where UserName='"&CookieUserName&"'")(0)



select case Request("menu")
case "add"
aim=HTMLEncode(Request("aim"))
if content=empty then error("<li>表白内容不能为空")
if aim=CookieUserName then error("<li>不能自己追求自己！")

If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&aim&"'" ).eof Then error("<li>系统不存在"&aim&"的资料")

If Not Conn.Execute("Select id From [BBSXP_Consort] where UserName='"&CookieUserName&"' and aim='"&aim&"'" ).eof Then error("<li>"&aim&"已经在您的追求列表当中")

sql="insert into [BBSXP_Consort] (UserName,aim,unburden) values ('"&CookieUserName&"','"&aim&"','"&content&"')"
Conn.Execute(SQL)


sql="insert into [BBSXP_Messages](UserName,incept,content) values ('"&CookieUserName&"','"&aim&"','<font color=0000FF>【社区配偶】："&content&"</font>')"
Conn.Execute(SQL)


Conn.execute("update [BBSXP_Users] set NewMessage=NewMessage+1 where UserName='"&aim&"'")

case "accept"

if Consort<>empty then error("<li>您当前已有配偶")

aim=Conn.Execute("Select aim From [BBSXP_Consort] where id="&id&"")(0)
if aim<>CookieUserName then error("<li>非法操作")

Consort=Conn.Execute("Select UserName From [BBSXP_Consort] where id="&id&"")(0)
if Conn.Execute("Select Consort From [BBSXP_Users] where UserName='"&Consort&"'")(0)<>empty then error("<li>"&Consort&"已经有配偶了")

Conn.execute("update [BBSXP_Users] set Consort='"&aim&"' where UserName='"&Consort&"'")
Conn.execute("update [BBSXP_Users] set Consort='"&Consort&"' where UserName='"&aim&"'")
Conn.execute("Delete from [BBSXP_Consort] where id="&id&"")
succeed("<li>您已经接受了"&Consort&"的追求<li><a href=Consort.asp>返回社区配偶</a><meta http-equiv=refresh content=3;url=Consort.asp>")



case "Del"
Conn.execute("Delete from [BBSXP_Consort] where id="&id&" and (UserName='"&CookieUserName&"' or aim='"&CookieUserName&"') ")

case "part"
Conn.execute("update [BBSXP_Users] set Consort='' where UserName='"&Consort&"'")
Conn.execute("update [BBSXP_Users] set Consort='' where UserName='"&CookieUserName&"'")
succeed("<li>分手成功<li><a href=Consort.asp>返回社区配偶</a><meta http-equiv=refresh content=3;url=Consort.asp>")

end select

%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		→ <a href="Consort.asp">社区配偶</a></td>
	</tr>
</table>
<br>
<center>

<table cellspacing="1" cellpadding="6" width="100%" border="0" class="a2">
	<tr>
		<td width="100%" height="10" align="middle" class="a1" colspan="4">
		<font face="宋体">我的追求者</font></td>
	</tr>
	<tr class="a3">
		<td width="10%" height="5" align="middle">用户名</td>
		<td height="5" align="middle">爱的表白</td>
		<td width="20%" height="5" align="middle"><font face="宋体">追求</font>时间</td>
		<td width="15%" height="5" align="middle">操作</td>
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
		<a href="?menu=part">分 手</a> <%else%>
		<a href="?menu=accept&id=<%=Rs("id")%>">接受</a>
		<a href="?menu=Del&id=<%=Rs("id")%>">拒绝</a> <%end if%></td>
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
		<td width="100%" height="10" align="middle" colspan="4">我追求的人</td>
	</tr>
	<tr class="a3">
		<td width="10%" height="5" align="middle">用户名</td>
		<td height="5" align="middle">爱的表白</td>
		<td width="20%" height="5" align="middle"><font face="宋体">追求</font>时间</td>
		<td width="15%" height="5" align="middle">操作</td>
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
		<a href="?menu=part">分 手</a> <%else%>
		<a href="?menu=Del&id=<%=Rs("id")%>">取消追求</a> <%end if%> </td>
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
UserSex="男"
case "female"
UserSex="女"
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
		<td align="middle" class="a1" colspan="7">我的配偶</td>
	</tr>
	<tr>
		<td width="20%" align="Left" class="a4" rowspan="3"><script>
if("<%=Userphoto%>"!=""){
document.write("<img src=<%=Userphoto%> border=0 onload='javascript:if(this.width>200)this.width=200'>")
}
</script>
		
		</td>
		<td width="80" align="Left" class="a4">昵称：</td>
		<td align="Left" class="a4" width="15%"><a href="Profile.asp?UserName=<%=Consort%>">
		<%=Consort%></a></td>
		<td width="80" align="Left" class="a4">姓名：</td>
		<td align="Left" class="a4" width="15%"><%=realname%></td>
		<td width="80" align="Left" class="a4">性别：</td>
		<td align="Left" class="a4" width="15%"><%=UserSex%></td>
	</tr>
	<tr>
		<td width="80" align="Left" class="a4" valign="top">国家：</td>
		<td align="Left" class="a4" width="15%"><%=country%></td>
		<td width="80" align="Left" class="a4">省份：</td>
		<td align="Left" class="a4" width="15%"><%=province%></td>
		<td width="80" align="Left" class="a4">城市：</td>
		<td align="Left" class="a4" width="15%"><%=city%></td>
	</tr>
	<tr>
		<td width="80" align="Left" class="a4" valign="top">个人说明：</td>
		<td align="Left" class="a4" colspan="5"><%=personal%></td>
	</tr>
	<tr>
		<td align="right" class="a4" valign="top" colspan="7">
		<a onclick="checkclick('您确定要与当前配偶分手？')" href="?menu=part">与当前配偶分手</a></td>
	</tr>
</table>
<%else%>
<table cellspacing="1" cellpadding="6" width="100%" border="0" class="a2">
	<form action="Consort.asp" method="POST">
		<input type="hidden" value="add" name="menu">
		<tr>
			<td width="77%" height="2" align="middle" class="a1" colspan="2">添加我想追求的人</td>
		</tr>
		<tr>
			<td width="12%" height="2" align="Left" class="a4">对方用户名：</td>
			<td width="64%" height="2" align="Left" class="a4">
			<input maxlength="30" size="15" name="aim"></td>
		</tr>
		<tr>
			<td width="12%" height="1" align="Left" class="a4" valign="top">爱的表白：</td>
			<td width="64%" height="1" align="Left" class="a4">
			<textarea name="content" rows="5" style="width:95%"></textarea></td>
		</tr>
		<tr>
			<td width="77%" height="1" align="center" class="a4" colspan="2">
			<input type="submit" value=" 确 定 ">
			<input onclick="checkclick('该项操作要清除全部的内容，您确定要清除吗?');" type="reset" value=" 重 写 "></td>
		</tr>
	</form>
</table>
<%end if%>

</center><%htmlend%>