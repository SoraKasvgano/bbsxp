<!-- #include file="Setup.asp" -->
<%
GroupID=RequestInt("GroupID")
HtmlTop
%>
<div class="CommonBreadCrumbArea">
	<div style="float:left"><%=ClubTree%></div>
<%if BestRole=1 then%>
	<div style="float:right"><a href="Moderation.asp?menu=ForumsFix" title="修复首页‘最后发表’栏目" onclick="return window.confirm('确定执行：修复首页‘最后发表’栏目 操作？')">修复统计</a></div>
<%end if%>
</div>

<%
sql="["&TablePrefix&"Threads] where Visible=1 and ThreadTop=2 order by ThreadID Desc"
Set Rs=Execute(sql)
if Not Rs.Eof Then 
%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>公告</td>
	</tr>
	<tr class="CommonListCell">
		<td>
			<marquee onMouseOver="this.stop();" onMouseOut="this.start();" scrollAmount="3">
	<%
		Do While Not Rs.Eof
			Response.Write "<a href='ShowPost.asp?ThreadID="&Rs("ThreadID")&"'><b>"&Rs("Topic")&"</b> ("&FormatDateTime(Rs("PostTime"),2)&")</a>　　"
			Rs.MoveNext
		Loop
	%>	
			</marquee>
		</td>
	</tr>
</table>

<%
End If
Rs.Close


if SiteConfig("DisplayTags")=1 then
sql="select top 14 * From ["&TablePrefix&"PostTags] where IsEnabled=1 and DateDiff("&SqlChar&"m"&SqlChar&",MostRecentPostDate,"&SqlNowString&")<1 order by TotalPosts Desc"
Set Rs=Execute(sql)
if Not Rs.Eof Then 
%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td><a href="Tags.asp">热门标签</a></td>
	</tr>
	<tr class="CommonListCell">
		<td class="TagList">
	<%
		i=0
		Do While Not Rs.Eof and i<14
			Response.Write "<li><a href='Tags.asp?TagID="&Rs("TagID")&"'>"&Rs("TagName")&"</a>("&Rs("TotalPosts")&")</li>"
			i=i+1
			Rs.MoveNext
		Loop
	%>
		</td>
	</tr>
</table>
<%
End If
Rs.Close
end if




if GroupID>0 then GroupIDSQL=" and GroupID="&GroupID&""
GroupGetRow=FetchEmploymentStatusList("Select GroupID,GroupName,GroupDescription,ForumColumns,Moderated From ["&TablePrefix&"Groups] where SortOrder>0 "&GroupIDSQL&" order by SortOrder")

if IsArray(GroupGetRow) then

	for j=0 to Ubound(GroupGetRow,2)

		if RequestCookies("ForumGroupDisplay"&GroupGetRow(0,j)&"")="none" then
			ForumGroupDisplay="style='display:none;'"
			ForumGroupDisplayImg="group_expand.gif"
		else
			ForumGroupDisplay=""
			ForumGroupDisplayImg="group_collapse.gif"
		end if

		GroupModeratedList=""
		if GroupGetRow(4,j)<>empty then
			GroupModeratedList="版主:"
			filtrate=split(GroupGetRow(4,j),"|")
			for i = 0 to ubound(filtrate)
				if ""&filtrate(i)&""<>"" then GroupModeratedList=GroupModeratedList&"<a href='Profile.asp?UserName="&filtrate(i)&"'>"&filtrate(i)&"</a> "
			next
		end if
	
		sql="Select * From ["&TablePrefix&"Forums] where GroupID="&GroupGetRow(0,j)&" and ParentID=0 and SortOrder>0 and IsActive=1 order by SortOrder"
		Set Rs=Execute(sql)
%>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
<%
		if GroupGetRow(3,j)<1 then	'竖排显示
%>
	<tr class=CommonListTitle>
		<td colspan="7">
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td><a href="default.asp?GroupID=<%=GroupGetRow(0,j)%>" title="<%=GroupGetRow(2,j)%>"><%=GroupGetRow(1,j)%></a></td>
					<td align="right"><%=GroupModeratedList%> <img style="CURSOR: pointer" onMouseDown="ForumGroupToggleCollapsed('<%=GroupGetRow(0,j)%>')" src="images/<%=ForumGroupDisplayImg%>" id="ForumGroupImg<%=GroupGetRow(0,j)%>" /></td>
				</tr>
			</table>
		</td>
	</tr>
	<tbody id="ForumGroup<%=GroupGetRow(0,j)%>" <%=ForumGroupDisplay%>>
	<tr class=CommonListHeader align="center">
		<td width="30"></td>
		<td>版块</td>
		<td width="50">主题</td>
		<td width="50">帖子</td>
		<td width="200">最后发表</td>
		<td width="100">版主</td>
	</tr>
<%
			do while not Rs.eof
	    		ShowForum
				Rs.Movenext
			loop
		else	'横排显示
			ki=0
			WidthPercent=100/GroupGetRow(3,j)
%>
	<tr class=CommonListTitle>
		<td colspan="<%=GroupGetRow(3,j)%>">
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td><a href="default.asp?GroupID=<%=GroupGetRow(0,j)%>" title="<%=GroupGetRow(2,j)%>"><%=GroupGetRow(1,j)%></a></td>
					<td align="right"><%=GroupModeratedList%> <img style="CURSOR: pointer" onMouseDown="ForumGroupToggleCollapsed('<%=GroupGetRow(0,j)%>')" src="images/<%=ForumGroupDisplayImg%>" id="ForumGroupImg<%=GroupGetRow(0,j)%>" /></td>
				</tr>
			</table>
		</td>
	</tr>
	<tbody id="ForumGroup<%=GroupGetRow(0,j)%>" <%=ForumGroupDisplay%>>
<%
			do while not Rs.eof
    			ShowForumFloor
				Rs.Movenext
			loop
			if ki<>0 and GroupGetRow(3,j)<>ki then
				for ki = ki to GroupGetRow(3,j)-1
					response.write "<td width='"&WidthPercent&"'></td>"
				next
			end if
		end if

		Rs.close

		response.write "</tbody></table>"
	next

end if
GroupGetRow=null




ForumName=SiteConfig("SiteName")
ForumID=0


%>


<br />
<!-- #include file="Utility/OnLine.asp" -->
<%
regOnline=Execute("Select count(sessionid) from ["&TablePrefix&"UserOnline] where UserName<>''")(0)
if BestOnline < Onlinemany then
	Execute("update ["&TablePrefix&"SiteSettings] Set BestOnline="&Onlinemany&",BestOnlineTime="&SqlNowString&"")
end if
if SiteConfig("DisplayWhoIsOnline")=1 then
%>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2">用户在线信息</td>
	</tr>
	<tr class="CommonListCell">
    	<td width="30" align="center"><img src="images/whos_online.gif" /></td>
		<td>
		<img src="images/plus.gif" id="followImg" style="cursor:pointer;" onClick="loadThreadFollow('ForumID=<%=ForumID%>')" />
		<%Response.Write("目前论坛总共 有 <b>"&Onlinemany&"</b> 人在线。其中注册用户 <b>"&regOnline&"</b> 人，访客 <b>"&Onlinemany-regOnline&"</b> 人。最高在线 <font color='red'><b>"&BestOnline&"</b></font> 人，发生在 <b>"&BestOnlineTime&"</b>")%>
		<div style="display:none" id="follow">
			<hr width="90%" SIZE="1" align="left"><span id=followTd class="UserList"><img src=images/loading.gif />正在加载...</span>
		</div>
		</td>
	</tr>
</table>

<br />
<%
end if
if SiteConfig("DisplayStatistics")=1 then
%>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2">统计信息</td>
	</tr>
	<tr class="CommonListCell">
		<td width="30" align="center"><img src="images/stats.gif" width="30" height="30" /></td>
		<td><%
			SQL="Select top 1 * from ["&TablePrefix&"Statistics] order by DateCreated desc"
			Set Rs=Execute(sql)
			'if Rs.eof then response.redirect "Install.asp"
		%>
		总共 <b><%=Rs("TotalTopics")%></b> 个主题 / <b><%=Rs("TotalPosts")%></b> 篇帖子 / <b><%=Rs("TotalUsers")%></b> 位用户<br />
		<%=FormatDateTime(Rs("DateCreated"),1)%>新增 <b><%=Rs("DaysTopics")%></b> 个新主题 / <b><%=Rs("DaysPosts")%></b> 篇新帖子 / <b><%=Rs("DaysUsers")%></b> 位新用户<br />
		欢迎新用户 <a href='Profile.asp?UserName=<%=Rs("NewestUserName")%>'><b><%=Rs("NewestUserName")%></b></a>
		<%Rs.close%>
		</td>
	</tr>
</table>
<br />
<%
end if
if SiteConfig("DisplayBirthdays")=1 then
%>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2">今天过生日的用户</td>
	</tr>
	<tr class="CommonListCell">
        <td width="30" align="center"><img src="images/birthday.gif" width="30" height="30" /></td>
        <td class="UserList">
<%
UserBirthdayGetRow=RequestApplication("UserBirthday")

if IsArray(UserBirthdayGetRow) then
	For i=0 To Ubound(UserBirthdayGetRow,2)
		%><li><a href="Profile.asp?UID=<%=UserBirthdayGetRow(0,i)%>"><%=UserBirthdayGetRow(1,i)%></a></li><%
	Next
	UserBirthdayGetRow=empty
Else
	Response.Write "今天没有过生日的用户" 
End if
UserBirthdayGetRow=null
%>
		</td>
	</tr>
</table><br />
<%
end if
if SiteConfig("DisplayLink")=1 then
%>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class="CommonListTitle">
		<td colspan="2" align="left">友情链接</td>
	</tr>
<%
LinksGetRow=RequestApplication("Links")		'直接读取Application缓存
if IsArray(LinksGetRow) then
	For i=0 To Ubound(LinksGetRow,2)
	
		if ""&LinksGetRow(1,i)&""="" or LinksGetRow(1,i)="http://" then
			Link1=Link1+"<li><a title='"&LinksGetRow(3,i)&"' href="&LinksGetRow(2,i)&" target=_blank>"&LinksGetRow(0,i)&"</a></li>"
		else
			Link2=Link2+"<a title='"&LinksGetRow(0,i)&""&chr(10)&""&LinksGetRow(3,i)&"' href="&LinksGetRow(2,i)&" target=_blank><img src="&LinksGetRow(1,i)&" border=0 width=88 height=31 /></a> "
		end if
	Next
End if
%>
	<tr class="CommonListCell">
		<td width="30" rowspan="2"><img src="images/shareforum.gif" /></td>
		<td class="LinkList"><%=Link1%></td>
	</tr>
	<tr class="CommonListCell">
		<td><%=Link2%></td>
	</tr>
</table>

<%
end if
LinksGetRow=null
%>
<br />

<table width="50%" align="center">
	<tr>
		<td align="center"><img src="images/forum_status_new.gif" border="0" title="一天内有新帖子的论坛" /> 有新帖的论坛</td>
		<td align="center"><img src="images/forum_status.gif" border="0" title="一天内无新帖子的论坛" /> 无新帖的论坛</td>
		<td align="center"><img src="images/forum_link.gif" border="0" title="论坛直接转向设定的链接" /> 带链接的论坛</td>
	</tr>
</table>

<%
HtmlBottom


Sub ShowForumFloor()
	if Rs("MostRecentPostSubject")<>empty then MostRecent="<a href='ShowPost.asp?ThreadID="&Rs("MostRecentThreadID")&"' title='"&Rs("MostRecentPostDate")&" by:"&Rs("MostRecentPostAuthor")&"'>"&Rs("MostRecentPostSubject")&"</a>"
	if Rs("ForumUrl")<>"" then
		ForumIcon="<img src=images/forum_link.gif>"
	elseif int(DateDiff("d",Rs("MostRecentPostDate"),Now())) < 2 then
		ForumIcon="<img src=images/forum_status_new.gif>"
	else
		ForumIcon="<img src=images/forum_status.gif>"
	end if
	
	ki=ki+1
	if ki=1 then response.write "<tr class=CommonListCell>"
%>
	<td width="<%=WidthPercent%>%">
		<table border=0 width="100%" cellspacing=0 cellpadding=5>
			<tr class="CommonListCell">
				<td rowspan="3" width="32"><%=ForumIcon%></td>
				<td colspan="3" title='<%=ReplaceText(BBCode(Rs("ForumDescription")),"<[^>]*>","")%>'>『 <a href="ShowForum.asp?ForumID=<%=Rs("ForumID")%>"><%=Rs("ForumName")%></a> 』</td>
			</tr>
			<tr class="CommonListCell">
				<td>今日 <font color="#F00"><%=Rs("TodayPosts")%></font></td>
				<td>主题 <%=Rs("TotalThreads")%></td>
				<td>帖子 <%=Rs("TotalPosts")%></td>
			</tr>
			<tr class="CommonListCell"><td colspan="3"><%=MostRecent%></td></tr>
		</table>
	</td>
<%
	if ki=GroupGetRow(3,j) then ki=0:response.write "</tr>"
End Sub
%>