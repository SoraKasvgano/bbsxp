<!-- #include file="Setup.asp" -->
<%
HtmlTop
if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")
if CookieUserAccountStatus<>1 then error("您的帐号未通过审核！")


	if CookieReputation < SiteConfig("InPrisonReputation") then error("您的声望低于"&SiteConfig("InPrisonReputation")&"，无法发表帖子！")

	
	
	
	if SiteConfig("RegUserTimePost") > 0 then
		StopPostTime=int(DateDiff("n",CookieUserRegisterTime,Now()))
		if StopPostTime < SiteConfig("RegUserTimePost") then error("<li>新注册用户必须等待 "&SiteConfig("RegUserTimePost")&" 分钟后才能发帖！<li>您必须再等待 "&SiteConfig("RegUserTimePost")-StopPostTime&" 分钟！")
	end if

	if SiteConfig("PostInterval") > 0 then
		StopPostTime=int(DateDiff("s",CookieUserPostTime,Now()))
		if StopPostTime < SiteConfig("PostInterval") then error("论坛限制一个人两次发帖间隔必须大于 "&SiteConfig("PostInterval")&" 秒！<li>您必须再等待 "&SiteConfig("PostInterval")-StopPostTime&" 秒！")
	end if

ThreadID=RequestInt("ThreadID")
PostID=RequestInt("PostID")
Quote=RequestInt("Quote")

sql="Select * From ["&TablePrefix&"Threads] where ThreadID="&ThreadID&""
Rs.Open sql,Conn,1
	if Rs.Eof then error("系统不存在此主题的资料")
	if Rs("Visible")=0 then error("此主题目前是审核状态，不接受新的回复")
	if Rs("Visible")=2 then error("此主题已经删除，不接受新的回复")
	if Rs("IsLocked")=1 then error("此主题已经关闭，不接受新的回复")

	
	ForumID=Rs("ForumID")
	Topic=Rs("Topic")
	PostAuthor=Rs("PostAuthor")
Rs.close

'if Execute("Select PostID from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&" and PostID="&PostID&"").Eof then error("系统没有找到回复帖ID")


%><!-- #include file="Utility/ForumPermissions.asp" --><%


if Request_Method = "POST" then
	if Request.Form = RequestApplication("LastPost") and SiteConfig("AllowDuplicatePosts")=0 then error("请不要提交重复数据")



	Subject=HTMLEncode(Request.Form("PostSubject"))
	Body=BodyEncode(Request.Form("Body"))
	Tags=Request.Form("Tags")
	TagArray=split(Tags,",")
	if Ubound(TagArray)>5 then Message=Message&"<li>标签不能超过5个"
	

	'if Request.Form("DisableBBCode")<>1 then Body=BBCode(Body)
	if Request.Form("DisableBBCode")=1 then Body=Replace(Body,CHR(91),"&#91")
	if Len(Body)<2 then Message=Message&"<li>文章内容不能小于 2 字符"
	if Message<>"" then error(""&Message&"")


	Rs.open "select * from ["&TablePrefix&"Users] where UserID="&CookieUserID&"",Conn,1,3
		Rs("TotalPosts")=Rs("TotalPosts")+1
		Rs("UserMoney")=Rs("UserMoney")+SiteConfig("IntegralAddPost")
		Rs("experience")=Rs("experience")+SiteConfig("IntegralAddPost")
		Rs("UserPostTime")=Now()
		Rs("UserRank")=UpUserRank()
	Rs.Update
	Rs.Close


	
	UpdateStatistics 0,0,1
	
	
	Rs.Open "Select top 1 * from ["&TablePrefix&"Posts]",Conn,1,3
	Rs.addNew
		Rs("ThreadID")=ThreadID
		Rs("ParentID")=PostID
		Rs("PostAuthor")=CookieUserName
		Rs("Subject")=Subject
		Rs("Body")=Body
		Rs("IPAddress")=REMOTE_ADDR
		if ModerateNewPost=1 and CookieModerationLevel=0 then
			Rs("Visible")=0
			HiddenCount=1
		else
			HiddenCount=0
		end if
	Rs.update
	PostID=Rs("PostID")
	Rs.close
	
	
	if Request.Form("UpFileID")<>"" then
		UpFileID=split(Request.Form("UpFileID"),",")
		for i = 0 to ubound(UpFileID)-1
			Execute("update ["&TablePrefix&"PostAttachments] Set Description='"&Subject&"',PostID='"&PostID&"' where UpFileID="&int(UpFileID(i))&" and UserName='"&CookieUserName&"'")
		next
	end if
	
	if SiteConfig("DisplayPostTags")=1 then AddTags()
	
	Execute("update ["&TablePrefix&"Threads] Set LastName='"&CookieUserName&"',TotalReplies=TotalReplies+1,HiddenCount=HiddenCount+"&HiddenCount&",LastTime="&SqlNowString&" where ThreadID="&ThreadID&"")
	Execute("update ["&TablePrefix&"Forums] Set MostRecentPostSubject='"&Topic&"',MostRecentPostAuthor='"&CookieUserName&"',MostRecentPostDate="&SqlNowString&",TodayPosts=TodayPosts+1,TotalPosts=TotalPosts+1,MostRecentThreadID="&ThreadID&" where ForumID="&ForumID&" or ForumID="&ParentID&"")

	ResponseApplication "LastPost",Request.Form



'注意：回帖需要审核时，就不发送主题订阅邮件通知！
	if ModerateNewPost=1 and CookieModerationLevel=0 then
		EnableCensorship="由于论坛设置新帖子审核制度，您发表的帖子需要等待激活才能显示。"
	else
		if SiteConfig("SelectMailMode")<>"" then

			Set Rs=Execute("Select top 100 * from ["&TablePrefix&"Subscriptions] where ThreadID="&ThreadID&" order by SubscriptionID")
			do while not Rs.eof
				MailAddRecipient=""&Rs("Email")&";"&MailAddRecipient
				Rs.movenext
			loop
			Rs.close
			

			LoadingEmailXml("NewMessagePostedToThread")
			MailBody=Replace(MailBody,"[Topic]",Topic)
			MailBody=Replace(MailBody,"[ThreadURL]","<a target=_blank href="&SiteConfig("SiteUrl")&"/ShowPost.asp?ThreadID="&ThreadID&">"&SiteConfig("SiteUrl")&"/ShowPost.asp?ThreadID="&ThreadID&"</a>")
			MailBody=Replace(MailBody,"[UserName]",CookieUserName)
			MailBody=Replace(MailBody,"[PostTime]",now())
			MailBody=Replace(MailBody,"[Subject]",Subject)
			MailBody=Replace(MailBody,"[body]",Body)
			SendMail MailAddRecipient,MailSubject,MailBody
		end if	

		EnableCensorship="<a href=ShowPost.asp?ThreadID="&ThreadID&">返回主题</a>"
	end if



	Message=Message&"<li>回复主题成功<li>"&EnableCensorship&"<li><a href=ShowForum.asp?ForumID="&ForumID&">返回论坛</a>"
	succeed Message,"ShowForum.asp?ForumID="&ForumID&""
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql="Select * from ["&TablePrefix&"Posts] where PostID="&PostID&""
	Set Rs=Execute(sql)
	if Rs.eof then error("PostID("&PostID&")不存在！")
	Body=BBCode(Rs("Body"))
	if Quote=1 then
		ReBody="[quote user="&chr(34)&Rs("PostAuthor")&chr(34)&"]"&Body&"[/quote]"
	else
		ReBody="<table width=100% cellspacing=1 cellpadding=5 class=CommonListArea>"
		ReBody=ReBody&"<tr class=CommonListTitle><td colspan=2><span id=RePostAuthor>"&Rs("PostAuthor")&"</span> 发表于："&Rs("PostDate")&"</td></tr>"
		ReBody=ReBody&"<tr class=CommonListCell><td><span id=ReBody>"&Body&"</span><br><br>"
		ReBody=ReBody&"<a class='CommonImageTextButton' style='BACKGROUND-IMAGE:url(images/Quote.gif)' title='引用回复这个帖子' href='javascript:ReplyQuote();'> 引用</a>"
		ReBody=ReBody&"</td></tr></table>"
	end if
	Rs.close
%>
<script language="JavaScript" type="text/javascript">
function ReplyQuote() { //Edit at 2007-05-26
	var QuoteValue = '[quote user="' + $("RePostAuthor").innerHTML + '"]';
	QuoteValue += $("ReBody").innerHTML + '[/quote]';
	BBSXPExecute('YuZi_QUOTE', QuoteValue);
}
</script>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → <%=ForumTree(ParentID)%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> → <a href="ShowPost.asp?ThreadID=<%=ThreadID%>"><%=Topic%></a> → 回复帖子</div>
<%if Quote<>1 then response.Write(ReBody)%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
<form name="form" method="post" onsubmit="return CheckForm(this);">
<input name="Body" type="hidden"<%if Quote=1 then%> value='<%=server.htmlencode(ReBody)%>'<%end if%> />
<input type=hidden name=ThreadID value=<%=ThreadID%> />
<input type=hidden name=PostID value=<%=PostID%> />
<input name="UpFileID" type="hidden" />
	<tr class=CommonListTitle>
		<td valign=Left colspan=2><b>回复帖子</b></td>
	</tr>
	<tr class="CommonListCell">
		<td width=180><b>标题 </b> </td>
		<td><input type="text" size="60" name="PostSubject" /></td>
	</tr>
<%if SiteConfig("UpFileOption")<>empty and PermissionAttachment=1 then%>
	<tr class="CommonListCell">
		<td valign=top><b>上传附件</b></td>
		<td><iframe id="UpLoadIframe" name="UpLoadIframe" src="UploadAttachment.asp" frameborder="0" width="100%" height="20" scrolling="no"></iframe></td>
	</tr>
<%end if%>
	<tr class="CommonListCell">
		<td valign=top>
			<br /><b>内容</b><br />（<a href="javascript:CheckLength();">查看内容长度</a>）<br /><br /><input id=DisableBBCode name=DisableBBCode type=checkbox value=1 /><label for=DisableBBCode> 禁用 BB 代码</label>
			<%if SiteConfig("UpFileOption")<>empty and PermissionAttachment=1 then%>
					<br /><br /><span id=UpFile></span>
			<%end if%>
		</td>
		<td height=250><script type="text/javascript" src="Editor/Post.js"></script></td>
	</tr>
<%if SiteConfig("DisplayPostTags")=1 then%>
	<tr class="CommonListCell">
		<td><b>标签<br /></b>以逗号“,”分隔</td>
		<td><input type="text" name="Tags" size="80" /> <a href="javascript:BBSXP_Modal.Open('Tags.asp?menu=SelectTags',500,420);" class="CommonTextButton">选择标签</a></td>
	</tr>
<%end if%>
	<tr class="CommonListCell">
		<td align=center colspan=2><input accesskey="s" title="(Alt + S)" type=submit value=" 回复 " name=EditSubmit />　<input type="Button" value=" 预览 " onclick="Preview()" />　<input onclick="history.back()" type="button" value=" 取消 " />　<input type="button" id="recoverdata" onclick="RestoreData()" title="恢复上次自动保存的数据" value="恢复数据" /></td>
	</tr>
</form>
</table>
<div name="Preview" id="Preview"></div>
<%
HtmlBottom
%>