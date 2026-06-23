<!-- #include file="Setup.asp" -->
<%
HtmlTop
%>
<!-- Markdown Support -->
<script src="js/marked.min.js"></script>
<script src="js/dompurify.min.js"></script>
<script src="js/markdown-handler.js"></script>
<link rel="stylesheet" href="css/markdown-content.css">
<%
if CookieUserName=empty then error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")
if CookieUserAccountStatus<>1 then error("�����ʺ�δͨ����ˣ�")


	if CookieReputation < SiteConfig("InPrisonReputation") then error("������������"&SiteConfig("InPrisonReputation")&"���޷��������ӣ�")




	if SiteConfig("RegUserTimePost") > 0 then
		StopPostTime=int(DateDiff("n",CookieUserRegisterTime,Now()))
		if StopPostTime < SiteConfig("RegUserTimePost") then error("<li>��ע���û�����ȴ� "&SiteConfig("RegUserTimePost")&" ���Ӻ���ܷ�����<li>�������ٵȴ� "&SiteConfig("RegUserTimePost")-StopPostTime&" ���ӣ�")
	end if

	if SiteConfig("PostInterval") > 0 then
		StopPostTime=int(DateDiff("s",CookieUserPostTime,Now()))
		if StopPostTime < SiteConfig("PostInterval") then error("��̳����һ�������η������������� "&SiteConfig("PostInterval")&" �룡<li>�������ٵȴ� "&SiteConfig("PostInterval")-StopPostTime&" �룡")
	end if

ThreadID=RequestInt("ThreadID")
PostID=RequestInt("PostID")
Quote=RequestInt("Quote")

sql="Select * From ["&TablePrefix&"Threads] where ThreadID="&ThreadID&""
Rs.Open sql,Conn,1
	if Rs.Eof then error("ϵͳ�����ڴ����������")
	if Rs("Visible")=0 then error("������Ŀǰ�����״̬���������µĻظ�")
	if Rs("Visible")=2 then error("�������Ѿ�ɾ�����������µĻظ�")
	if Rs("IsLocked")=1 then error("�������Ѿ��رգ��������µĻظ�")


	ForumID=Rs("ForumID")
	Topic=Rs("Topic")
	PostAuthor=Rs("PostAuthor")
Rs.close

'if Execute("Select PostID from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&" and PostID="&PostID&"").Eof then error("ϵͳû���ҵ��ظ���ID")


%><!-- #include file="Utility/ForumPermissions.asp" --><%


if Request_Method = "POST" then
	if Request.Form = RequestApplication("LastPost") and SiteConfig("AllowDuplicatePosts")=0 then error("�벻Ҫ�ύ�ظ�����")



	Subject=HTMLEncode(Request.Form("PostSubject"))
	Body=BodyEncode(Request.Form("Body"))
	Tags=Request.Form("Tags")
	TagArray=split(Tags,",")
	if Ubound(TagArray)>5 then Message=Message&"<li>��ǩ���ܳ���5��"


	'if Request.Form("DisableBBCode")<>1 then Body=BBCode(Body)
	if Request.Form("DisableBBCode")=1 then Body=Replace(Body,CHR(91),"&#91")
	if Len(Body)<2 then Message=Message&"<li>�������ݲ���С�� 2 �ַ�"
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
			UpFileItem=SafeLongValue(UpFileID(i),0)
			if UpFileItem>0 then Execute("update ["&TablePrefix&"PostAttachments] Set Description='"&SqlString(Subject)&"',PostID="&PostID&" where UpFileID="&UpFileItem&" and UserName='"&SqlString(CookieUserName)&"'")
		next
	end if

	if SiteConfig("DisplayPostTags")=1 then AddTags()

	Execute("update ["&TablePrefix&"Threads] Set LastName='"&SqlString(CookieUserName)&"',TotalReplies=TotalReplies+1,HiddenCount=HiddenCount+"&HiddenCount&",LastTime="&SqlNowString&" where ThreadID="&ThreadID&"")
	Execute("update ["&TablePrefix&"Forums] Set MostRecentPostSubject='"&SqlString(Topic)&"',MostRecentPostAuthor='"&SqlString(CookieUserName)&"',MostRecentPostDate="&SqlNowString&",TodayPosts=TodayPosts+1,TotalPosts=TotalPosts+1,MostRecentThreadID="&ThreadID&" where ForumID="&ForumID&" or ForumID="&ParentID&"")

	ResponseApplication "LastPost",Request.Form



'ע�⣺������Ҫ���ʱ���Ͳ��������ⶩ���ʼ�֪ͨ��
	if ModerateNewPost=1 and CookieModerationLevel=0 then
		EnableCensorship="������̳��������������ƶȣ���������������Ҫ�ȴ����������ʾ��"
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

		EnableCensorship="<a href=ShowPost.asp?ThreadID="&ThreadID&">��������</a>"
	end if



	Message=Message&"<li>�ظ�����ɹ�<li>"&EnableCensorship&"<li><a href=ShowForum.asp?ForumID="&ForumID&">������̳</a>"
	succeed Message,"ShowForum.asp?ForumID="&ForumID&""
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql="Select * from ["&TablePrefix&"Posts] where PostID="&PostID&""
	Set Rs=Execute(sql)
	if Rs.eof then error("PostID("&PostID&")�����ڣ�")
	Body=BBCode(Rs("Body"))
	if Quote=1 then
		ReBody="[quote user="&chr(34)&Rs("PostAuthor")&chr(34)&"]"&Body&"[/quote]"
	else
		ReBody="<table width=100% cellspacing=1 cellpadding=5 class=CommonListArea>"
		ReBody=ReBody&"<tr class=CommonListTitle><td colspan=2><span id=RePostAuthor>"&Rs("PostAuthor")&"</span> �����ڣ�"&Rs("PostDate")&"</td></tr>"
		ReBody=ReBody&"<tr class=CommonListCell><td><span id=ReBody>"&Body&"</span><br><br>"
		ReBody=ReBody&"<a class='CommonImageTextButton' style='BACKGROUND-IMAGE:url(images/Quote.gif)' title='���ûظ��������' href='javascript:ReplyQuote();'> ����</a>"
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
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� <%=ForumTree(ParentID)%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> �� <a href="ShowPost.asp?ThreadID=<%=ThreadID%>"><%=Topic%></a> �� �ظ�����</div>
<%if Quote<>1 then response.Write(ReBody)%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
<form name="form" method="post" onsubmit="return CheckForm(this);">
<input name="Body" type="hidden"<%if Quote=1 then%> value="<%=Server.HTMLEncode(ReBody)%>"<%end if%> />
<input type=hidden name=ThreadID value=<%=ThreadID%> />
<input type=hidden name=PostID value=<%=PostID%> />
<input name="UpFileID" type="hidden" />
	<tr class=CommonListTitle>
		<td valign=Left colspan=2><b>�ظ�����</b></td>
	</tr>
	<tr class="CommonListCell">
		<td width=180><b>���� </b> </td>
		<td><input type="text" size="60" name="PostSubject" /></td>
	</tr>
<%if SiteConfig("UpFileOption")<>empty and PermissionAttachment=1 then%>
	<tr class="CommonListCell">
		<td valign=top><b>�ϴ�����</b></td>
		<td><iframe id="UpLoadIframe" name="UpLoadIframe" src="UploadAttachment.asp" frameborder="0" width="100%" height="20" scrolling="no"></iframe></td>
	</tr>
<%end if%>
	<tr class="CommonListCell">
		<td valign=top>
			<br /><b>����</b><br />��<a href="javascript:CheckLength();">�鿴���ݳ���</a>��<br /><br /><input id=DisableBBCode name=DisableBBCode type=checkbox value=1 /><label for=DisableBBCode> ���� BB ����</label>
			<%if SiteConfig("UpFileOption")<>empty and PermissionAttachment=1 then%>
					<br /><br /><span id=UpFile></span>
			<%end if%>
		</td>
		<td height=250>
			<script type="text/javascript" src="Editor/Post.js"></script>
			<!-- Initialize Markdown editor after BBCode editor loads -->
			<script>
			// Add Markdown toolbar to the editor
			setTimeout(function() {
				if (typeof initMarkdownEditor === 'function' && window.YuZi_EDIT) {
					// Markdown support is available
					console.log('Markdown support enabled');
				}
			}, 100);
			</script>
		</td>
	</tr>
<%if SiteConfig("DisplayPostTags")=1 then%>
	<tr class="CommonListCell">
		<td><b>��ǩ<br /></b>�Զ��š�,���ָ�</td>
		<td><input type="text" name="Tags" size="80" /> <a href="javascript:BBSXP_Modal.Open('Tags.asp?menu=SelectTags',500,420);" class="CommonTextButton">ѡ���ǩ</a></td>
	</tr>
<%end if%>
	<tr class="CommonListCell">
		<td align=center colspan=2><input accesskey="s" title="(Alt + S)" type=submit value=" �ظ� " name=EditSubmit />��<input type="Button" value=" Ԥ�� " onclick="Preview()" />��<input onclick="history.back()" type="button" value=" ȡ�� " />��<input type="button" id="recoverdata" onclick="RestoreData()" title="�ָ��ϴ��Զ����������" value="�ָ�����" /></td>
	</tr>
</form>
</table>
<div name="Preview" id="Preview"></div>
<%
HtmlBottom
%>