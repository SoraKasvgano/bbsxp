<!-- #include file="Setup.asp" -->
<%
HtmlTop
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


ForumID=RequestInt("ForumID")

%><!-- #include file="Utility/ForumPermissions.asp" --><%


if Request_Method = "POST" then
	if Request.Form=RequestApplication("LastPost") and SiteConfig("AllowDuplicatePosts")=0 then error("�벻Ҫ�ύ�ظ�����")
	if SiteConfig("EnableAntiSpamTextGenerateForPost")=1 then
		if Request.Form("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then Message=Message&"<li>��֤�����</li>"
	end if
	Subject=HTMLEncode(Request.Form("Subject"))
	Category=HTMLEncode(Request.Form("Category"))
	Body=BodyEncode(Request.Form("Body"))
	Description=Left(HTMLEncode(Request.Form("Description")),200)


	Tags=Request.Form("Tags")
	ThreadEmoticonID=RequestInt("ThreadEmoticonID")
	StickyDate=RequestInt("StickyDate")

	if PermissionManage<>1 then StickyDate=0

	if Request.Form("DisableBBCode")=1 then Body=Replace(Body,CHR(91),"&#91;")
	
	if Len(Subject)<2 then Message=Message&"<li>�������ⲻ��С�� 2 �ַ�"
	if Len(Body)<2 then Message=Message&"<li>�������ݲ���С�� 2 �ַ�"
	
	if PermissionCreatePoll=1 and RequestInt("IsVote")=1 then
		j=0
		for each formElement in request.form
			if instr(formElement,"Bbsxp_Input_")>0 then
				if trim(request.form(""&formElement&""))<>"" then
					allpollTopic=""&allpollTopic&""&request.form(""&formElement&"")&"|"
					j=j+1
				end if
			end if
		next
		if j<SiteConfig("MinVoteOptions") then Message=Message&"<li>ͶƱѡ������� "&SiteConfig("MinVoteOptions")&" ��"
		if j>SiteConfig("MaxVoteOptions") then Message=Message&"<li>ͶƱѡ��ܳ��� "&SiteConfig("MaxVoteOptions")&" ��"
		for y = 1 to j
			Votenum=""&Votenum&"0|"
		next
	end if

	IsMultiplePoll=RequestInt("IsMultiplePoll")
	VoteItems=HTMLEncode(allpollTopic)
	VoteExpiry=now()+RequestInt("VoteExpiry")

	if Not Isdate(VoteExpiry) then Message=Message&"<li>ͶƱ����ʱ�����"
	
	TagArray=split(Tags,",")
	if Ubound(TagArray)>5 then Message=Message&"<li>��ǩ���ܳ���5��"

	if Message<>"" then error(""&Message&"")

	Rs.open "select * from ["&TablePrefix&"Users] where UserID="&CookieUserID&"",Conn,1,3
		Rs("TotalPosts")=Rs("TotalPosts")+1
		Rs("UserMoney")=Rs("UserMoney")+SiteConfig("IntegralAddThread")
		Rs("experience")=Rs("experience")+SiteConfig("IntegralAddThread")
		Rs("UserPostTime")=Now()
		Rs("UserRank")=UpUserRank()
	Rs.Update
	Rs.Close


	Rs.Open "Select top 1 * from ["&TablePrefix&"Threads]",Conn,1,3
	Rs.addNew
		if PermissionManage=1 then Rs("ThreadStyle")=HTMLEncode(Request.Form("ThreadStyle"))
		Rs("PostAuthor")=CookieUserName
		Rs("PostTime")=now()
		Rs("lastname")=CookieUserName
		Rs("lasttime")=now()
		Rs("Topic")=Subject
		Rs("Description")=Description
		Rs("ForumID")=ForumID
		Rs("Category")=Category
		Rs("ThreadEmoticonID")=ThreadEmoticonID
		if StickyDate=999 then
			Rs("ThreadTop")=2
		elseif StickyDate>0 and StickyDate<999 then
			Rs("ThreadTop")=1
		end if
		Rs("StickyDate")=DateAdd("d", StickyDate, now())
		if PermissionCreatePoll=1 then Rs("IsVote")=RequestInt("IsVote")
		if Request("IsLocked")=1 then Rs("IsLocked")=1
		if ModerateNewThread=1 and CookieModerationLevel=0 then
			Rs("Visible")=0
			Rs("HiddenCount")=1
		end if
	Rs.update
	ThreadID=Rs("ThreadID")
	Rs.close

	
	if PermissionCreatePoll=1 and RequestInt("IsVote")=1 then
		Execute("insert into ["&TablePrefix&"Votes] (ThreadID,IsMultiplePoll,Items,Result,Expiry) values ('"&ThreadID&"',"&IsMultiplePoll&",'"&VoteItems&"','"&Votenum&"','"&VoteExpiry&"')")
	end if
	
	UpdateStatistics 0,1,1

	Rs.Open "Select top 1 * from ["&TablePrefix&"Posts]",Conn,1,3
	Rs.addNew
		Rs("ThreadID")=ThreadID
		Rs("PostAuthor")=CookieUserName
		Rs("Subject")=Subject
		Rs("Body")=Body
		if ModerateNewThread=1 and CookieModerationLevel=0 then Rs("Visible")=0
		Rs("IPAddress")=REMOTE_ADDR
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
	

	Execute("update ["&TablePrefix&"Forums] Set MostRecentPostSubject='"&Subject&"',MostRecentPostAuthor='"&CookieUserName&"',MostRecentPostDate="&SqlNowString&",TodayPosts=TodayPosts+1,TotalThreads=TotalThreads+1,TotalPosts=TotalPosts+1,MostRecentThreadID="&ThreadID&" where ForumID="&ForumID&" or ForumID="&ParentID&"")


	Session("VerifyCode")=""
	ResponseApplication "LastPost",Request.Form


	if ModerateNewThread=1 and CookieModerationLevel=0 then
		EnableCensorship="������̳��������������ƶȣ��������������Ҫ�ȴ����������ʾ��"
	else
		EnableCensorship="<a href=ShowPost.asp?ThreadID="&ThreadID&">��������</a>"
	end if

	Message="<li>�����ⷢ��ɹ�</li><li>"&EnableCensorship&"</li><li><a href=ShowForum.asp?ForumID="&ForumID&">������̳</a></li>"
	succeed Message,"ShowForum.asp?ForumID="&ForumID&""

end if
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� <%=ForumTree(ParentID)%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> �� ��������</div>
<form name="form" method="post" onSubmit="return CheckForm(this);" />
<input name="Body" type="hidden" />
<input name="Description" type="hidden" />
<input name="UpFileID" type="hidden" />
<input name="ThreadStyle" id="ThreadStyle" type="hidden" />
<input type=hidden name=ForumID value=<%=ForumID%> />
<input type=hidden name=IsVote value=<%=RequestInt("Poll")%> />
<table cellspacing=1 cellpadding=5 border=0 align=center class=CommonListArea  width=100%>
	<tr class=CommonListTitle>
		<td valign=Left colspan=2>��������</td>
	</tr>
	<tr class="CommonListCell">
		<td width="180"><b>����</b><%if PermissionManage=1 then%>��<a href="javascript:BBSXP_Modal.Open('Utility/SelectStyle.htm',500,420);">����</a>��<%end if%></td>
		<td><input type="text" size="60" id="Subject" name="Subject" /></td>
	</tr>
	<tr class="CommonListCell">
		<td width="180"><b>���</b></td>
		<td>
			<select name=Category size=1><option value="" selected="selected">��</option>
<%
	if TotalCategorys<>empty then
		filtrate=split(TotalCategorys,"|")
		for i = 0 to ubound(filtrate)
			response.write "<OPTION value='"&filtrate(i)&"'>"&filtrate(i)&"</OPTION>"
		next
	end if
%>
			</select> <a href="javascript:BBSXP_Modal.Open('Utility/AddCategory.htm', 500, 100);">������</a></td>
	</tr>
	<tr class="CommonListCell">
		<td valign=top align=Left><b>����</b></td>
		<td><input type=radio value=0 name=ThreadEmoticonID <%if ThreadEmoticonID=0 then%>checked<%end if%> />��<br />
		<script language="JavaScript" type="text/javascript">
		for(i=1;i<=30;i++) {
			document.write("<input type=radio value="+i+" name=ThreadEmoticonID><IMG src=images/Emoticons/"+i+".gif width=23 height=23 />��")
			if (i ==10 || i ==20){document.write("<br />")}
		}
		</script>		</td>
	</tr>
<%if PermissionCreatePoll=1 and Request("Poll")=1 then%>
	<script language="javascript" type="text/javascript" src="Utility/LabelDom.js"></script>
	<tr class="CommonListCell">
		<td valign=top align=Left><b>ͶƱ</b><br />ÿ��һ��ͶƱ��Ŀ<br />
<input type=radio checked="checked" value=0 name=IsMultiplePoll id=IsMultiplePoll /><label for=IsMultiplePoll>��ѡͶƱ</label><br /><input type=radio value=1 name=IsMultiplePoll id=IsMultiplePoll_1 /><label for=IsMultiplePoll_1>��ѡͶƱ</label></font> <br />����ʱ�� <input type="text" size="2" name="VoteExpiry" value="7" onkeyup=if(isNaN(this.value))this.value='7' /> ���
			
			
</td>
		<td valign="top">
		<div id="VoteOptionList"></div><a class="CommonTextButton" href="javascript:AddLabel('VoteOptionList')">�����Ŀ</a>
		<script language="javascript" type="text/javascript">
			var MinCount=parseInt('<%=SiteConfig("MinVoteOptions")%>');
			var MaxCount=parseInt('<%=SiteConfig("MaxVoteOptions")%>');
			var CurrentCount=0;
			var RealCount=0;
			init();
		</script>
		</td>
	</tr>
<%
end if
if SiteConfig("UpFileOption")<>empty and PermissionAttachment=1 then%>
	<tr class="CommonListCell">
		<td valign=top><b>�ϴ�����</b></td>
		<td><iframe id="UpLoadIframe" name="UpLoadIframe" src="UploadAttachment.asp" frameborder="0" width="100%" height="20" scrolling="no"></iframe></td>
	</tr>
<%end if%>
	<tr class="CommonListCell">
		<td valign=top>
		<br /><b>����</b><br />��<a href="javascript:CheckLength();">�鿴���ݳ���</a>��<br /><br />
		<input id=LockMyPost name=IsLocked type=checkbox value="1" /><label for=LockMyPost> ��������</label><br />
        <input id=DisableBBCode name=DisableBBCode type=checkbox value=1 /><label for=DisableBBCode> ���� BB ����</label>
        <br /><br />
		<%if PermissionManage=1 then%><br />
		<select name=StickyDate size=1>
			<option value="0" selected="selected">�ö�</option>
			<option value="1">1 ��</option>
			<option value="3">3 ��</option>
			<option value="7">1 ��</option>
			<option value="14">2 ��</option>
			<option value="30">1 ��</option>
			<option value="90">3 ��</option>
			<option value="180">6 ��</option>
			<option value="366">1 ��</option>
			<%if BestRole=1 then%><option value="999">����</option><%end if%>
		</select>
		<%end if%>
		
		
		<%if SiteConfig("UpFileOption")<>empty and PermissionAttachment=1 then%>
			<br /><br /><span id=UpFile></span>
		<%end if%>

		</td>
		<td height=250 valign="top"><script type="text/javascript" src="Editor/Post.js"></script></td>
	</tr>
<%if SiteConfig("DisplayPostTags")=1 then%>
	<tr class="CommonListCell">
		<td><b>��ǩ</b><br />�Զ��š�,���ָ�</td>
		<td><input type="text" name="Tags" size="80" id="Tags" /> <a href="javascript:BBSXP_Modal.Open('Tags.asp?menu=SelectTags',500,420);" class="CommonTextButton">ѡ���ǩ</a></td>
	</tr>
<%end if%>
<%if SiteConfig("EnableAntiSpamTextGenerateForPost")=1 then%>
	<tr class="CommonListCell">
		<td><b>��֤��</b></td>
		<td><input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">���������ȡ��֤��</span> <span id="CheckVerifyCode" style="color:red"></span></td>
	</tr>
<%end if%>
	<tr class="CommonListCell">
		<td align=center colspan=2><input type=submit accesskey="s" title="(Alt + S)" value=" ���� " name=EditSubmit />��<input type="Button" value=" Ԥ�� " onClick="Preview();" />��<input onClick="history.back()" type="button" value=" ȡ�� " />��<input type="button" id="recoverdata" onclick="RestoreData()" title="�ָ��ϴ��Զ����������" value="�ָ�����" /></td>
	</tr>
</table>
</form>
<div name="Preview" id="Preview"></div>
<%
HtmlBottom
%>