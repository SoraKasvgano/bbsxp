<!-- #include file="Setup.asp" --><%
HtmlTop
count=Execute("Select count(sessionid)from ["&TablePrefix&"UserOnline]")(0)
regOnline=Execute("Select count(sessionid)from ["&TablePrefix&"UserOnline] where UserName<>''")(0)
TotalThreads=Execute("Select SUM(TotalThreads)from ["&TablePrefix&"Forums]")(0)
TolReTopic=Execute("Select SUM(TotalPosts)from ["&TablePrefix&"Forums]")(0)
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 查看论坛状态</div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td width="16%" align="center"><a href="?">在线情况</a></td>
		<td width="16%" align="center"><a href="?menu=cutline">在线图例</a></td>
		<td width="16%" align="center"><a href="?menu=UserSex">性别图例</a></td>
		<td width="16%" align="center"><a href="?menu=TodayPage">今日图例</a></td>
		<td width="16%" align="center"><a href="?menu=board">主题图例</a></td>
		<td width="16%" align="center"><a href="?menu=TotalPosts">帖子图例</a></td>
	</tr>
	<tr class="CommonListCell">
		<td width="48%" align="center" colspan="3">总在线 <%=count%> 人。其中注册用户 <%=regOnline%> 人，访客 <%=count-regOnline%> 人。</td>
		<td width="48%" align="center" colspan="3">总帖数 <%=TolReTopic%> 篇。其中主题 <%=TotalThreads%> 篇，回帖 <%=tolReTopic-TotalThreads%> 篇。</td>
	</tr>
</table>
<br />
<%
select case Request("menu")
	case ""
		default
	case "cutline"
		cutline
	case "board"
		board
	case "TotalPosts"
		TotalPosts
	case "TodayPage"
		TodayPage
	case "UserSex"
		UserSex
end select

Sub default
	if Request_Method = "POST" and BestRole<>1 then error("只有超级版主与管理员才能使用查询功能")
	Key=HTMLEncode(Request.Form("Key"))
	Find=HTMLEncode(Request.Form("Find"))

	if Len(Find)>10 then error("非法操作")
	if Key<>empty then SqlFind=" where "&Find&"='"&Key&"'"
	sql="Select * from ["&TablePrefix&"UserOnline] "&SqlFind&" order by LastTime Desc"
	Rs.Open sql,Conn,1
		PageSetup=20 '设定每页的显示数量
		Rs.Pagesize=PageSetup
		TotalPage=Rs.Pagecount  '总页数
		PageCount = RequestInt("PageIndex")
		if PageCount <1 then PageCount = 1
		if PageCount > TotalPage then PageCount = TotalPage
		if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
		i=0
		Do While Not Rs.EOF and i<PageSetup
			i=i+1
			if BestRole<>1 then
				ips=split(Rs("IPAddress"),".")
				ShowIP=""&ips(0)&"."&ips(1)&".*.*"
			else
				ShowIP=""&Rs("IPAddress")&""
			end if


			if ""&Rs("UserName")&""="" then
				UserName="<FONT COLOR=#C0C0C0>"&Rs("SessionID")&"</FONT>"
			else
				if Rs("IsInvisible")=0 or BestRole=1 then UserName="<a href=Profile.asp?UserName="&Rs("UserName")&">"&Rs("UserName")&"</a>"
				if Rs("IsInvisible")=1 then UserName=UserName&"(隐身)"
			end if


			place2=Rs("Topic")
			place = "『"&Rs("ForumName")&"』"
			if Rs("ThreadID")>0 then place2 = "<a href=ShowPost.asp?ThreadID="&Rs("ThreadID")&">"&place2&"</a>"
			if Rs("ForumID")>0 then place = "<a href=ShowForum.asp?ForumID="&Rs("ForumID")&">"&place&"</a>"

			allline=""&allline&"<TR align=center class=CommonListCell><TD width=120>"&ShowIP&"</TD><TD width=120>"&Rs("cometime")&"</TD><TD width=120>"&UserName&"</TD><TD width=140>"&place&"</TD><TD height=24>"&place2&"</TD><TD width=120>"&Rs("lasttime")&"</TD></TR>"

			Rs.Movenext
		loop
	Rs.close
%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr align="center" class=CommonListTitle>
		<td width=120>IP地址</td>
		<td width=120>登录时间</td>
		<td width=120>用户</td>
		<td width=140>所在论坛</td>
		<td>所在主题</td>
		<td width=120>活动时间</td>
	</tr>
	<%=allline%>
</table>
<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
		<td valign="top"><%ShowPage()%></td>

		<td align="right">
		<form action="ViewOnline.asp" method="POST">
			<select name=Find>
				<option value="UserName">查询用户</option>
				<option value="IPAddress">查询ＩＰ</option>
			</select>	<input size="15" value="<%=Key%>" name="Key"> <input type="submit" value=" 确定 ">
		</form>
		</td>
	</tr>
</table>
<%
End Sub

Sub cutline
	sql="Select * from ["&TablePrefix&"Forums] where SortOrder>0 and IsActive=1"
	Set Rs=Execute(sql)
%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td valign="bottom" align="center">论坛名称</td>
		<td valign="bottom" align="center">图形比例</td>
		<td valign="bottom" align="center">在线人数</td>
	</tr>
<%
	i=0
		Do While Not Rs.EOF
			Onlinemany=Execute("Select count(sessionid)from ["&TablePrefix&"UserOnline] where ForumID="&Rs("ForumID")&"")(0)
%>
	<tr class="CommonListCell">
		<td width="21%" align="center"><a href="ShowForum.asp?ForumID=<%=Rs("ForumID")%>"><%=Rs("ForumName")%></a></td>
		<td width="65%"><%PercentBar(Onlinemany/count*100)%></td>
		<td align="center" width="12%"><%=Onlinemany%></td>
	</tr>
	<%		
			i=i+1
			if i=10 then i=0
			Rs.MoveNext
		loop
	Rs.Close   
%>
</table>
<%
End Sub

Sub board
	if TotalThreads=0 then TotalThreads=1
	sql="Select * from ["&TablePrefix&"Forums] where SortOrder>0 and IsActive=1 order by TotalThreads Desc"
	Set Rs=Execute(sql)
%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td valign="bottom" align="center">论坛名称</td>
		<td valign="bottom" align="center">图形比例</td>
		<td valign="bottom" align="center">主题数</td>
	</tr>
	<%
	i=0
	Do While Not Rs.EOF
%>
	<tr class="CommonListCell">
		<td width="21%" align="center"><a href="ShowForum.asp?ForumID=<%=Rs("ForumID")%>"><%=Rs("ForumName")%></a></td>
		<td width="65%"><%PercentBar(Rs("TotalThreads")/TotalThreads*100)%></td>
		<td align="center" width="12%"><%=Rs("TotalThreads")%></td>
	</tr>
	<%
		i=i+1
		if i=10 then i=0		
		Rs.MoveNext
	loop
	Rs.Close   
%>
</table>
<%
End Sub

Sub TotalPosts
	if tolReTopic=0 then tolReTopic=1
	sql="Select * from ["&TablePrefix&"Forums] where SortOrder>0 and IsActive=1 order by TotalPosts Desc"
	Set Rs=Execute(sql)
%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td valign="bottom" align="center">论坛名称</td>
		<td valign="bottom" align="center">图形比例</td>
		<td valign="bottom" align="center">帖子数</td>
	</tr>
	<%
	i=0
	Do While Not Rs.EOF
%>
	<tr class="CommonListCell">
		<td width="21%" align="center"><a href="ShowForum.asp?ForumID=<%=Rs("ForumID")%>"><%=Rs("ForumName")%></a></td>
		<td width="65%"><%PercentBar(Rs("TotalPosts")/tolReTopic*100)%></td>
		<td align="center" width="12%"><%=Rs("TotalPosts")%></td>
	</tr>
	<%		
		i=i+1
		if i=10 then i=0		
		Rs.MoveNext
	loop
	Rs.Close
%>
</table>
<%
End Sub

Sub TodayPage
	tolTodayPosts=Execute("Select SUM(TodayPosts)from ["&TablePrefix&"Forums]")(0)
	if tolTodayPosts=0 then tolTodayPosts=1
	sql="Select * from ["&TablePrefix&"Forums] where SortOrder>0 and IsActive=1 order by TodayPosts Desc"
	Set Rs=Execute(sql)
%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td valign="bottom" align="center">论坛名称</td>
		<td valign="bottom" align="center">图形比例</td>
		<td valign="bottom" align="center">今日帖数</td>
	</tr>
	<%
	i=0
	Do While Not Rs.EOF
%>
	<tr class="CommonListCell">
		<td width="21%" align="center"><a href="ShowForum.asp?ForumID=<%=Rs("ForumID")%>"><%=Rs("ForumName")%></a></td>
		<td width="65%"><%PercentBar(Rs("TodayPosts")/tolTodayPosts*100)%></td>
		<td align="center" width="12%"><%=Rs("TodayPosts")%></td>
	</tr>
	<%		
		i=i+1
		if i=10 then i=0			
		Rs.MoveNext
	loop
	Rs.Close   
%>
</table>
<%
End Sub

Sub UserSex
	count=Execute("Select count(UserID) from ["&TablePrefix&"Users]")(0)
	male=Execute("Select count(UserID) from ["&TablePrefix&"Users] where UserSex=1")(0)
	female=Execute("Select count(UserID) from ["&TablePrefix&"Users] where UserSex=2")(0)
%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td valign="bottom" align="center">性别</td>
		<td valign="bottom" align="center">图形比例</td>
		<td valign="bottom" align="center">人数</td>
	</tr>
	<tr class="CommonListCell">
		<td width="10%" align="center"><img src="images/Sex_1.gif" /></td>
		<td width="75%"><%PercentBar(male/count*100)%></td>
		<td align="center" width="12%"><%=male%></td>
	</tr>
	<tr class="CommonListCell">
		<td width="10%" align="center"><img src="images/Sex_2.gif" /></td>
		<td width="75%"><%PercentBar(female/count*100)%></td>
		<td align="center" width="12%"><%=female%></td>
	</tr>
	<tr class="CommonListCell">
		<td width="10%" align="center">未知</td>
		<td width="75%"><%PercentBar((count-male-female)/count*100)%></td>
		<td align="center" width="12%"><%=count-male-female%></td>
	</tr>
</table>
<%
End Sub
function PercentBar(PercentWidth)
	response.write("<div class='percent'><div style='width:"&int(PercentWidth)&"%'></div></div>")
end function
HtmlBottom
%>