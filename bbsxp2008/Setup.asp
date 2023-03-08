<!-- #include file="Conn.asp" -->
<%
ii=0

if BBSxpDataBaseVersion<>ForumsDataBaseVersion then error("程序与数据库版本不符合，请更新数据库版本到 "&ForumsDataBaseVersion&" ！")

if SiteConfig("EnableBannedUsersToLogin")=0 and CookieUserAccountStatus=2 then error("您的帐号已被禁用！")

if SiteConfig("BannedIP")<>"" then
	filtrate=split(SiteConfig("BannedIP"),"|")
	for i = 0 to ubound(filtrate)
		if instr("|"&REMOTE_ADDR&"","|"&filtrate(i)&"") > 0 then error(""&REMOTE_ADDR&"被禁止进入论坛！")
	next
end if

if SiteConfig("SiteDisabled")=1 then
	if instr(Script_Name,"admin_") = 0 and instr(Script_Name,"login.asp")=0 and instr(Script_Name,"install.asp")=0 then error(""&SiteConfig("SiteDisabledReason")&"")
end if

if Request_Method = "POST" then
	if instr(""&Http_Referer&"","http://"&Server_Name&"") = 0 and instr(Script_Name,"login.asp")=0 then error("<li>引用页错误！"&Http_Referer&"<li>系统无法识别您的引用页！<li>建议您关闭防火墙后再提交此信息！")
end if

Sub HtmlHead
	MetaDescription=SiteConfig("MetaDescription")
	MetaKeywords=SiteConfig("MetaKeywords")
	if ""&HtmlHeadTitle&""<>"" then
		MetaDescription=HtmlHeadTitle
		MetaKeywords=HtmlHeadTitle&","&MetaKeywords
	end if
	if ""&HtmlHeadDescription&"">"" then MetaDescription=ThreadDescription
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<%if SiteConfig("NoCacheHeaders")=1 then%>
	<!-- no cache headers -->
	<meta http-equiv="Pragma" content="no-cache" />
	<meta http-equiv="Expires" content="-1" />
	<meta http-equiv="Cache-Control" content="no-cache" />
	<!-- end no cache headers -->
	<%end if%>
	<meta http-equiv="Content-Type" content="text/html; charset=<%=BBSxpCharset%>" />

	<meta name="keywords" content="<%=MetaKeywords%>" />
	<meta name="description" content="<%=MetaDescription%>" />
	<meta name="GENERATOR" content="<%=ForumsProgram%> (Build: <%=ForumsBuild%>)" />
	<link rel="shortcut icon" type="image/x-icon" href="images/yuzi.ico" />
<%
	if SiteConfig("EnableForumsRSS")=1 then
		if instr(Script_Name,"showforum.asp") > 0 then response.write "	<link rel=alternate type=application/rss+xml href=rss.asp?"&Query_String&" />"&CHR(10)
	end if
	
	if HtmlHeadTitle<>"" then HtmlHeadTitle=HtmlHeadTitle&" - "
	response.write "	<title>"&HtmlHeadTitle&SiteConfig("SiteName")&" - Powered By BBSXP</title>"

	
%>
</head>

<script language="JavaScript" type="text/javascript">var CookiePath="<%=SiteConfig("CookiePath")%>",CookieDomain="<%=SiteConfig("CookieDomain")%>";</script>
<script language="JavaScript" type="text/javascript" src="Utility/global.js"></script>
<script language="JavaScript" type="text/javascript" src="Utility/BBSXP_Modal.js"></script>
<script language="JavaScript" type="text/javascript">InitThemes('<%=SiteConfig("DefaultSiteStyle")%>');</script>


<%
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HtmlTop
HtmlHead
IsResponseTop=1
%>
<body>
<%=GenericHeader%>

<div class="CommonHeader">
<div class="CommonTop">

<div class="CommonTopLogo" onClick="window.location.href='default.asp'" title="<%=SiteConfig("SiteName")%>"></div>
<div class="CommonTopBanner"><%=GenericTop%></div>
<div class="CommonTopFrameText"><script language="JavaScript" type="text/javascript">if (top !== self) {document.write('<a target="_top" href="'+location.href+'">平板模式</a>');}else {document.write('<a target="_top" href="Frame.asp">分栏模式</a>');}</script></div>
<div class="CommonTopMenuText">


<%
if CookieUserName=empty then
%><a href="javascript:BBSXP_Modal.Open('Login.asp',380,170);">登录</a> | <a href="CreateUser.asp">注册</a> | <a href="ViewOnline.asp">在线情况</a> | <a href="Search.asp?ForumID=<%=RequestInt("ForumID")%>">搜索</a> | <a href="Help.asp">帮助</a><%
else
%>&raquo;　<a href="Profile.asp?UID=<%=CookieUserID%>"><%=CookieUserName%></a> | <a onmouseover="showmenu(event,'<div class=menuitems><a href=javascript:UrlPost(&quot;Cookies.asp?menu=Online&quot;)>上线</a></div><div class=menuitems><a href=javascript:UrlPost(&quot;Cookies.asp?menu=Invisible&quot;)>隐身</a></div>')" style="cursor:default">
<%
	if RequestCookies("Invisible")="1" then
		response.write "隐身"
	else
		response.write "上线"
	end if
	
	response.write("</a> | ")

	if SiteConfig("AccountActivation")=2 then response.write("<a href=InviteUser.asp>邀请用户</a> | ")
%>
			<a onMouseOver="showmenu(event,'<div class=menuitems><a href=EditProfile.asp>编辑资料</a></div><div class=menuitems><a href=EditProfile.asp?menu=pass>密码修改</a></div><div class=menuitems><a href=MyUpFiles.asp>上传管理</a></div><div class=menuitems><a href=MyFavorites.asp>收 藏 夹</a></div><div class=menuitems><a href=MyMessage.asp>短信服务</a></div>')" style="cursor:default">控制面板</a> |
			<a onMouseOver="showmenu(event,'<div class=menuitems><a href=ShowBBS.asp>最新主题</a></div><div class=menuitems><a href=ShowBBS.asp?menu=HotViews>人气主题</a></div><div class=menuitems><a href=ShowBBS.asp?menu=HotReplies>热门主题</a></div><div class=menuitems><a href=ShowBBS.asp?menu=NoReplies>未回复帖</a></div><div class=menuitems><a href=ShowBBS.asp?menu=GoodTopic>精华主题</a></div><div class=menuitems><a href=ShowBBS.asp?menu=VoteTopic>投票主题</a></div><div class=menuitems><a href=ShowBBS.asp?menu=MyTopic&amp;UserName=<%=CookieUserName%>>我的主题</a></div><div class=menuitems><a href=ShowBBS.asp?menu=MyPost&amp;UserName=<%=CookieUserName%>>我的回帖</a></div><div class=menuitems><a href=ShowBBS.asp?menu=MySubscriptions&amp;UserName=<%=CookieUserName%>>我的订阅</a></div>')" style="cursor:default">主题索引</a> |
			<a onMouseOver="showmenu(event,'<div class=menuitems><a href=ViewOnline.asp>在线情况</a></div><div class=menuitems><a href=ViewOnline.asp?menu=cutline>在线图例</a></div><div class=menuitems><a href=ViewOnline.asp?menu=UserSex>性别图例</a></div><div class=menuitems><a href=ViewOnline.asp?menu=TodayPage>今日图例</a></div><div class=menuitems><a href=ViewOnline.asp?menu=board>主题图例</a></div><div class=menuitems><a href=ViewOnline.asp?menu=TotalPosts>帖子图例</a></div><%if SiteConfig("PublicMemberList")=1 then%><div class=menuitems><a href=Members.asp>成员列表</a></div><%end if%>')" style="cursor:default">统计</a> | 
			<a href="Search.asp?ForumID=<%=RequestInt("ForumID")%>">搜索</a><span id="MenuListID"></span>
			<script language="javascript" type="text/javascript">ShowMenuList("Xml/Menu.xml");</script>
<%
	response.write " | <a href=Help.asp>帮助</a> | <a onclick='return log_out()'>退出</a>"
	
	if BestRole=1 then
	%> | <a onMouseOver="showmenu(event,'<div class=menuitems><a href=Moderation.asp>主题管理</a></div><div class=menuitems><a href=Moderation.asp?menu=Visible>审 核 区</a></div><div class=menuitems><a href=Moderation.asp?menu=Recycle>回 收 站</a></div>')" style="cursor:default">审核</a><%
	end if
	if CookieUserRoleID = 1 then response.write " | <a href=Admin_Default.asp target=_top>管理</a>"
end if
%>
</div>

</div>

<div class="CommonBody">
<%

End Sub
''''''''''''''''''''''''''''''
Sub HtmlBottom
	ProcessTime=FormatNumber((timer()-startime))
	if ProcessTime<1 then ProcessTime="0"&ProcessTime
%>
	<table cellspacing=0 cellpadding=0 width=100% border="0">
		<tr>
	        <td align="center"><br /><%=GenericBottom%></td>
        </tr>
		<tr>
        	<td align="right"><script type="text/javascript">loadThemes("<%=SiteConfig("DefaultSiteStyle")%>");</script></td>
        </tr>
	</table>
</div>
	<div class="CommonFooter">
        <li style="text-align:left;width:30%;"><a href="http://www.bbsxp.com/" target="_blank"><img src="images/PoweredByBBSxp.gif" border="0" /></a></li>
        <li style="text-align:center;width:40%">
			<a href="SendMessage.asp">联系我们</a> - <a href="<%=SiteConfig("CompanyURL")%>" target="_blank"><%=SiteConfig("CompanyName")%></a> - <a href="Archiver.asp?<%=Query_String%>">论坛存档</a> - <a href="#top">返回顶部</a><br />
			Powered by <a target="_blank" href="http://www.bbsxp.com"><%=ForumsProgram%></a> &copy; 1998-<%=year(now)%> <a href="http://www.yuzi.net/" target="_blank">Yuzi.Net</a>
        </li>
        <li style="text-align:right;width:29%;">
            Processed in <%=ProcessTime%> second(s)<br />Server Time <%=now()%>
        </li>
	</div>
</div>
<%
response.write "<!-- SQL查询("&SqlQueryNum&") -->"

if CookieUserName<>"" then
	if CookieNewMessage > 0 then
		AddApplication "Message_"&CookieUserName,"【讯息通知】<a target='_blank' href='MyMessage.asp'>您有 "&CookieNewMessage&" 条新讯息，请注意查收</a>"
		Execute("update ["&TablePrefix&"Users] Set NewMessage=0 where UserID="&CookieUserID&"")
	end if

if RequestApplication("Message_"&CookieUserName)<>"" then
%>
<div id="ApplicationMsg" style=" Z-INDEX:1;WIDTH:250px;VISIBILITY:hidden;position:absolute;" onClick="if(typeof(IsReadMsg) == 'undefined')Ajax_CallBack(false,'MsgTitle','Loading.asp?menu=CloseMsg');IsReadMsg=true;">
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
<tr class=CommonListTitle><td><div style=float:left id='MsgTitle'>临时讯息</div><div style=float:right><a href="javascript:MsgClose()">Ｘ</a></div></td></tr>
<tr class="CommonListCell"><td><%=RequestApplication("Message_"&CookieUserName)%></td></tr>
</table>
</div>
<script language="JavaScript" type="text/javascript">
var MsgDivID="ApplicationMsg"
MsgGet();
</script>

<%
end if


end if

%>
</body>
</html>

<%
	Response.End
End Sub

''''''''''''''''''''''''''''''''
Sub AdminTop
	HtmlHead
	response.write "<body><div class=CommonListCell  style='PADDING-TOP:15px; PADDING-RIGHT:5px; PADDING-LEFT:5px; PADDING-BOTTOM:15px;text-align:center;'>"
End Sub



Sub AdminFooter
	ProcessTime=FormatNumber((timer()-startime))
	if ProcessTime<1 then ProcessTime="0"&ProcessTime
	response.write "<br /><br /><input onclick='history.go(-1)' type='button' value=' &lt;&lt; 返 回 '>　　<input onclick='history.go(1)' type='button' value='  前 进 &gt;&gt;'><br />"
	response.write "<hr width=80% SIZE=0 class=CommonListArea />Powered by <a target=_blank href=http://www.bbsxp.com>"&ForumsProgram&"</a> &nbsp;&copy; 1998-"&year(now)&"<br />Server Time "&now()&"<br />Processed in "&ProcessTime&" second(s)</div></body></html>"
End Sub
''''''''''''''''''''''''''''''''
Function succeed(Message,Url)

if IsResponseTop<>1 then HtmlTop
if Left(Message,4)<>"<li>" then Message="<li>"&Message
%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> → 论坛信息</div>
<table cellspacing="1" cellpadding="5" width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td width="100%" align="center">论坛提示信息</td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top" align="Left" height="122">
		<table cellspacing="0" cellpadding="5" width="100%" border="0">
			<tr>
				<td width="83%" valign="top">
					<b><span id="yu">0</span>秒钟后系统将自动返回...</b>
					<ul>
					<%=Message%>
					<li><a href=Default.asp>返回论坛首页</a></li>
					<script language="javascript" type="text/javascript">
					document.write("<li><a href="+document.referrer+">"+document.referrer+"</a></li>");
					</script>
					</ul>
				</td>
				<td width="17%"><img height="97" src="images/succ.gif" width="95" /></td>
			</tr>
		</table>
		</td>
	</tr>
</table>

<script language="JavaScript" type="text/javascript">
Url="<%=Url%>";
if (!Url){Url=document.referrer}
function countDown(secs){
	$("yu").innerHTML=secs;
	if(--secs>0) {
		setTimeout("countDown("+secs+")",1000);
	}
	else {
		window.location.href=Url;
	}
}
countDown(5);
</script>
<%
HtmlBottom
End Function

Sub error(Message)
if IsResponseTop<>1 then HtmlTop
if Left(Message,4)<>"<li>" then Message="<li>"&Message
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 论坛信息</div>
<table cellspacing="1" cellpadding="5" width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td width="100%" colspan="2" align="center">论坛提示信息</td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top" align="Left" colspan="2" height="122">
		<table cellspacing="0" cellpadding="5" width="100%" border="0">
			<tr>
				<td width="83%" valign="top"><b>操作不成功的可能原因或建议：</b><ul><%=Message%></ul></td>
				<td width="17%"><img height="97" src="images/err.gif" width="95" /></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr align="center" class="CommonListCell">
		<td valign="center" colspan="2"><input onClick="history.back()" type="submit" value=" << 返 回 上 一 页 " /></td>
	</tr>
</table>
<%
HtmlBottom
End Sub

Sub ShowForum()
	TodayPostsHtml=""
	ModeratedList=""
	if Rs("Moderated")<>empty then
		filtrate=split(Rs("Moderated"),"|")
		for i = 0 to ubound(filtrate)
			if ""&filtrate(i)&""<>"" then ModeratedList=ModeratedList&"<a href='Profile.asp?UserName="&filtrate(i)&"'>"&filtrate(i)&"</a> "
		next
	end if


	ForumUrlStr="ShowForum.asp?ForumID="&Rs("ForumID")&""
	if Rs("ForumUrl")<>"" then ForumUrlStr=Rs("ForumUrl")
	
	
	If SiteConfig("IsShowSonForum") = 1 then
		SubForumStrings=""
		SubForumGetRow=FetchEmploymentStatusList("Select ForumID,ForumName,ForumUrl from ["&TablePrefix&"Forums] where ParentID="&Rs("ForumID")&" and SortOrder>0 and IsActive=1 order by SortOrder")
		if IsArray(SubForumGetRow) then
		SubForumStrings="<br /><b>子版面："
		for iii=0 to Ubound(SubForumGetRow,2)
			if SubForumGetRow(2,iii)<>"" then
				SubForumStrings=SubForumStrings&"<a href="&SubForumGetRow(2,iii)&">"&SubForumGetRow(1,iii)&"</a>　"
			else
				SubForumStrings=SubForumStrings&"<a href=ShowForum.asp?ForumID="&SubForumGetRow(0,iii)&">"&SubForumGetRow(1,iii)&"</a>　"
			end if
		next
		SubForumStrings=SubForumStrings&"</b>"
		end if
		SubForumGetRow=null
	end if
	
%>
	<tr class="CommonListCell">
		<td align="center" width="40">
<%
	if Rs("ForumUrl")<>"" then
		response.write "<img src=images/forum_link.gif>"
	elseif int(DateDiff("d",Rs("MostRecentPostDate"),Now())) < 2 then
		response.write "<img src=images/forum_status_new.gif>"
	else
		response.write "<img src=images/forum_status.gif>"
	end if
	
if Rs("TodayPosts") > 0 then TodayPostsHtml="<font color='red'>("&Rs("TodayPosts")&")</font>"
%>
		</td>
		<td><a href="<%=ForumUrlStr%>"><%=Rs("ForumName")%></a> <%=TodayPostsHtml%><br /><%=BBCode(Rs("ForumDescription"))%><%=SubForumStrings%></td>
		<td align="center"><%=Rs("TotalThreads")%></td>
		<td align="center"><%=Rs("TotalPosts")%></td>
		<td width="200"><a href=ShowPost.asp?ThreadID=<%=Rs("MostRecentThreadID")%>><b><%=Left(Rs("MostRecentPostSubject"),14)%></b></a><br />by <a href="Profile.asp?UserName=<%=Rs("MostRecentPostAuthor")%>"><%=Rs("MostRecentPostAuthor")%></a><br /><%=Rs("MostRecentPostDate")%></td><td align="center" width="100"><%=ModeratedList%></td>
	</tr>
<%
End Sub

Sub ShowThread()
	if Rs("ThreadTop")=2 then
		IconImage="topic-announce.gif alt='公告主题'"
	elseif Rs("ThreadTop")=1 then
		IconImage="topic-pinned.gif alt='置顶主题'"
	elseif Rs("IsGood")=1 then
		IconImage="topic-popular.gif alt='精华主题'"
	elseif Rs("IsLocked")=1 then
		IconImage="topic-locked.gif alt='主题锁定'"
	elseif Rs("IsVote")=1 then
		IconImage="topic-poll.gif alt='投票主题'"
	elseif DateDiff("d",Rs("PostTime"),Now()) <= SiteConfig("PopularPostThresholdDays")   and  (  Rs("TotalReplies")=>SiteConfig("PopularPostThresholdPosts") or  Rs("TotalViews")=>SiteConfig("PopularPostThresholdViews") ) then
		IconImage="topic-hot.gif alt='热门主题'"
	else
		IconImage="topic.gif alt='普通主题'"
	end if
	
	if Rs("TotalReplies")=0 then
		replies="-"
	else
		replies=Rs("TotalReplies")
	end if
	
	if Rs("Category")<>"" then
		CategoryHtml="[<a href=ShowForum.asp?ForumID="&Rs("ForumID")&"&Category="&Rs("Category")&">"&Rs("Category")&"</a>] "
	else
		CategoryHtml=""
	end if
	if Rs("ThreadEmoticonID")>0 then
		ThreadEmoticonID="<img src=images/Emoticons/"&Rs("ThreadEmoticonID")&".gif> "
	else
		ThreadEmoticonID=""
	end if
	
	if SiteConfig("DisplayThreadStatus")=1 then
		if Rs("ThreadStatus")=1 then
			ThreadStatus="<img src=images/status_Answered.gif align=middle title='主题状态：已解决'>"
		elseif Rs("ThreadStatus")=2 then
			ThreadStatus="<img src=images/status_NotAnswered.gif align=middle title='主题状态：未解决'>"
		else
			ThreadStatus="<img src=images/status_NotSet.gif align=middle>"
		end if
	end if
	if Rs("HiddenCount")>0 then
		ThreadHidden="&nbsp;<img src='images/InVisible.gif' align=middle alt='"&Rs("HiddenCount")&"未审核帖子' />"
	end if
	if Rs("DeletedCount")>0 then
		ThreadDel="&nbsp;<img src='images/recycle.gif' align=middle alt='"&Rs("DeletedCount")&"已删除帖子' />"
	end if
	
	if DateDiff("d",Rs("PostTime"),Now()) < 2 then
		NewHtml=" <img title='一天内新发表的主题' src=images/new.gif align=absmiddle>"
	else
		NewHtml=""
	end if
	if Rs("TotalRatings")>0 then StarHtml="<a style=CURSOR:pointer onclick="&CHR(34)&"OpenWindow('PostRating.asp?ThreadID="&Rs("ThreadID")&"')"&CHR(34)&" ><img border=0 src=Images/Star/"&cint(Rs("RatingSum")/Rs("TotalRatings"))&".gif align=middle></a>&nbsp;"
	if Rs("TotalReplies")=>SiteConfig("PostsPerPage") then
		MaxPostPage=fix(Rs("TotalReplies")/SiteConfig("PostsPerPage"))+1 '共多少页
		ShowPostPage="( <img src=images/multiPage.gif> "
		For PostPage = 1 To MaxPostPage
			if PostPage<11 or MaxPostPage=PostPage then ShowPostPage=""&ShowPostPage&"<a href=ShowPost.asp?PageIndex="&PostPage&"&ThreadID="&Rs("ThreadID")&"><b>"&PostPage&"</b></a> "
		Next
		ShowPostPage=""&ShowPostPage&")"
	else
		ShowPostPage=""
	end if
%>
	<tr class="CommonListCell" id="Thread<%=Rs("ThreadID")%>">
		<td width="55%">

				<table width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="30"><a target="_blank" href="ShowPost.asp?ThreadID=<%=Rs("ThreadID")%>"><img src=images/<%=IconImage%> border=0 /></a></td>
						<td<%if SiteConfig("EnablePostPreviewPopup")=1 then%> title='<%=Rs("Description")%>'<%end if%>><%=checkboxHtml%><%
						
						if Rs("Visible")=0 then
							Response.write("等待审核：")
						elseif Rs("Visible")=2 then
							Response.write("已被删除：")
					    end if
						
						
						%><%=ThreadEmoticonID&CategoryHtml%><a href="ShowPost.asp?ThreadID=<%=Rs("ThreadID")%>"><span style="<%=Rs("ThreadStyle")%>"><%=Rs("Topic")%></span></a><%=ShowPostPage%><%=NewHtml%></td>
						<td align="right"><%=StarHtml&ThreadStatus%><%if PermissionManage=1 then Response.write(ThreadHidden&ThreadDel)%></td>
					</tr>
				</table>

		</td>
		
		<td align="center" width="13%"><a href="Profile.asp?UserName=<%=Rs("PostAuthor")%>"><%=Rs("PostAuthor")%></a><br /><%=FormatDateTime(Rs("PostTime"),2)%></td>
		<td align="center" width="7%"><%=replies%></td>
		<td align="center" width="7%"><%=Rs("TotalViews")%></td>
		<td align="center" width="18%"><%=Rs("lasttime")%><br />by <a href="Profile.asp?UserName=<%=Rs("lastname")%>"><%=Rs("lastname")%></a></td>
		
		<%
		  if PermissionManage=1 then
		  	if Rs("ThreadTop")=2 and BestRole<>1 then
		  		response.write("<td width=50>&nbsp;</td>")
			else
		  		response.write("<td width=50><input type=checkbox value="&Rs("ThreadID")&" name=ThreadID onclick="&chr(34)&"CheckSelected(this.form,this.checked,'Thread"&Rs("ThreadID")&"')"&chr(34)&"></td>")
			end if
		  end if
		%>
	</tr>
<%
End Sub

''''''''''''''''''''''''''''''''
Sub Alert(Message)
%>
	<script language="JavaScript" type="text/javascript">
	alert('<%=Message%>');
	window.history.back();
	</script>
	<script language="JavaScript" type="text/javascript">window.close();</script>
<%
	Response.End
End Sub

''''''''''''''''''''''''''''''''

Sub AlertForModal(Message)
%>
	<script language="JavaScript" type="text/javascript">
	alert('<%=Message%>');
	<%if Request_Method <> "POST" then%>
	parent.BBSXP_Modal.Close();
	<%else%>
	window.history.back();
	<%end if%>
	</script>
<%
	Response.End
End Sub

''''''''''''''''''''''''''''''''

Sub ShowPage()
	PageUrl=ReplaceText(Request.QueryString,"PageIndex=([0-9]*)&","")
	if Request.Form<>empty then PageUrl=""&PageUrl&"&"&Request.Form&""
	%><script language="JavaScript" type="text/javascript">ShowPage(<%=TotalPage%>,<%=PageCount%>,"<%=PageUrl%>")</script><%
End Sub
%>