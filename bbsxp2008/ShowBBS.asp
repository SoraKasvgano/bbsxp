<!-- #include file="Setup.asp" -->
<%
UserName=HTMLEncode(Request("UserName"))
TempStr="MyTopic|MyPost|MySubscriptions"
if instr("|"&TempStr&"|","|"&Request("menu")&"|")>0 and CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")

HtmlTop
select case Request("menu")
	case ""
		ForumTitle="最新主题"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 order by ThreadID Desc"
	case "HotViews"
		ForumTitle="人气主题"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and DateDiff("&SqlChar&"d"&SqlChar&",PostTime,"&SqlNowString&") < "&SiteConfig("PopularPostThresholdDays")&" order by TotalViews Desc"
	case "HotReplies"
		ForumTitle="热门主题"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and DateDiff("&SqlChar&"d"&SqlChar&",PostTime,"&SqlNowString&") < "&SiteConfig("PopularPostThresholdDays")&" order by TotalReplies Desc"
	case "GoodTopic"
		ForumTitle="精华主题"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and IsGood=1 order by ThreadID Desc"
	case "VoteTopic"
		ForumTitle="投票主题"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and IsVote=1 order by ThreadID Desc"
	case "NoReplies"
		ForumTitle="未回复帖"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and TotalReplies=0 order by ThreadID Desc"
	case "MyTopic"
		ForumTitle="我的主题"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and PostAuthor='"&UserName&"' order by ThreadID Desc"
	case "MyPost"
		ForumTitle="我参与的主题"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and ThreadID in (Select distinct ThreadID from ["&TablePrefix&"Posts] where PostAuthor='"&UserName&"' and Visible=1 and ParentID>0) order by ThreadID Desc"
	case "MySubscriptions"
		ForumTitle="我的订阅"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and ThreadID in (Select distinct ThreadID from ["&TablePrefix&"Subscriptions] where UserName='"&UserName&"') order by ThreadID Desc"
	case else
		error("错误参数！")
end select

Rs.Open sql,Conn,1
%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> → <%=ForumTitle%></div>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
      <td align="center"><a href="?">最新主题</a></td>
      <td align="center"><a href="?menu=HotViews">人气主题</a></td>
      <td align="center"><a href="?menu=HotReplies">热门主题</a></td>
      <td align="center"><a href="?menu=NoReplies">未回复帖</a></td>
      <td align="center"><a href="?menu=GoodTopic">精华主题</a></td>
      <td align="center"><a href="?menu=VoteTopic">投票主题</a></td>
      <td align="center"><a href="?menu=MyTopic&UserName=<%=CookieUserName%>">我的主题</a></td>
      <td align="center"><a href="?menu=MyPost&UserName=<%=CookieUserName%>">我的回帖</a></td>
      <td align="center"><a href="?menu=MySubscriptions&UserName=<%=CookieUserName%>">我的订阅</a></td>
	</tr>
</table>
<br />
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="5">主题</td>
	</tr>
	<tr class=CommonListHeader align=center>
		<td>标题</td>
		<td>作者</td>
		<td>回复</td>
		<td>查看</td>
		<td>最后更新</td>
	</tr>
<%

	PageSetup=SiteConfig("ThreadsPerPage") '设定每页的显示数量
	Rs.Pagesize=PageSetup
	TotalPage=Rs.Pagecount  '总页数
	PageCount = RequestInt("PageIndex")
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
	i=0
	Do While Not RS.EOF and i<pagesetup
		i=i+1
		ShowThread()
		Rs.MoveNext
	loop
Rs.Close
%>
</table>
<table cellspacing=0 cellpadding=0 border=0 width=100%><tr><td><%ShowPage()%></td></tr></table>
<%
HtmlBottom
%>