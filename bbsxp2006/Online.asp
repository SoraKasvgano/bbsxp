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
		→ 查看论坛状态</td>
	</tr>
</table>
<br>
<table cellspacing="1" cellpadding="4" width="100%" align="center" border="0" class="a2">
	<tr class="a1" id="TableTitleLink">
		<td width="16%" align="center" height="10"><b><a href="Online.asp">在线情况</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=cutline">在线图例</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=UserSex">性别图例</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=TodayPage">今日图例</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=board">主题图例</a></b></td>
		<td width="16%" align="center" height="10"><b>
		<a href="Online.asp?menu=ForumPosts">帖子图例</a></b></td>
	</tr>
	<tr class="a3">
		<td width="48%" align="center" height="10" colspan="3">总帖数 <%=tolReTopic%> 
		篇。其中主题 <%=ForumThreads%> 篇，回帖 <%=tolReTopic-ForumThreads%> 篇。</td>
		<td width="48%" align="center" height="10" colspan="3">总在线 <%=count%> 人。其中注册用户 
		<%=regOnline%> 人，访客 <%=count-regOnline%> 人。</td>
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

PageSetup=20 '设定每页的显示数量
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
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
UserName="访 客"
elseif Rs("eremite")=1 and membercode<4 then
UserName="隐 身"
else
UserName="<a href=Profile.asp?UserName="&Rs("UserName")&">"&Rs("UserName")&"</a>"
end if



place2=""
if Rs("act")<>"" then
place2 = "<a href="&Rs("acturl")&">"&Rs("act")&"</a>"
place = "『 "&Rs("ForumName")&" 』"
else
place = "『 <a href="&Rs("acturl")&">"&Rs("ForumName")&"</a> 』"
end if
allline=""&allline&"<TR align=middle class=a4><TD height=24>"&ShowIP&"</TD><TD height=24>"&Rs("cometime")&"</TD><TD height=24>"&UserName&"</TD><TD height=24>"&place&"</TD><TD height=24>"&place2&"</TD><TD height=24>"&Rs("lasttime")&"</TD></TR>"

Rs.Movenext
loop
Rs.close



%>
<table cellspacing="1" cellpadding="1" width="100%" align="center" border="0" class="a2">
	<tr align="middle" class="a1" height="23">
		<td>IP地址</td>
		<td>登录时间</td>
		<td>用户名</td>
		<td>所在论坛</td>
		<td>所在主题</td>
		<td>活动时间</td>
	</tr>
	<%=allline%>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td valign="top"><%ShowPage()%></td>

		<td align="right">
		<form action="Online.asp" method="POST">
<select name=Find>
<option value="UserName">查询用户</option>
<option value="IP">查询ＩＰ</option>
</select>		
<input size="15" value="<%=Key%>" name="Key">
			<input type="submit" value=" 确定 "></form>
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
		<td class="a1" valign="bottom" align="middle" height="20">论坛名称</td>
		<td class="a1" valign="bottom" align="middle" height="20">图形比例</td>
		<td class="a1" valign="bottom" align="middle" height="20">在线人数</td>
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
		<td class="a1" valign="bottom" align="middle" height="20">论坛名称</td>
		<td class="a1" valign="bottom" align="middle" height="20">图形比例</td>
		<td class="a1" valign="bottom" align="middle" height="20">主题数</td>
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
		<td class="a1" valign="bottom" align="middle" height="20">论坛名称</td>
		<td class="a1" valign="bottom" align="middle" height="20">图形比例</td>
		<td class="a1" valign="bottom" align="middle" height="20">帖子数</td>
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
		<td class="a1" valign="bottom" align="middle" height="20">论坛名称</td>
		<td class="a1" valign="bottom" align="middle" height="20">图形比例</td>
		<td class="a1" valign="bottom" align="middle" height="20">今日帖数</td>
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
		<td class="a1" valign="bottom" align="middle" height="20">性别</td>
		<td class="a1" valign="bottom" align="middle" height="20">图形比例</td>
		<td class="a1" valign="bottom" align="middle" height="20">人数</td>
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
		<td width="10%" height="2" align="center">未知</td>
		<td width="75%" height="2">
		<img height="8" src="images/bar/2.gif" width="<%=(count-male-female)/count*100%>%"></td>
		<td align="center" width="12%" height="2"><%=count-male-female%></td>
	</tr>
</table>
<%
end sub


htmlend
%>