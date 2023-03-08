<!-- #include file="Setup.asp" --><%
HtmlTop
if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")

ThreadID=RequestInt("ThreadID")
PostID=RequestInt("PostID")
sql="Select * From ["&TablePrefix&"Threads] where ThreadID="&ThreadID&""
Rs.Open sql,Conn,1
	if Rs.eof or Rs.bof then error"<li>系统不存在该帖子的资料"
	ForumID=Rs("ForumID")
	Topic=Rs("Topic")
	Visible=Rs("Visible")
Rs.close

if Visible=2 then  error("该帖子的主题目前是删除状态！")
if Visible=0 then  error("该帖子的主题目前是审核状态！")

%><!-- #include file="Utility/ForumPermissions.asp" --><%

if PostID<1 then error("参数错误，未指定操作帖子ID")

if Execute("Select Visible From ["&TablePrefix&"Posts] where PostID="&PostID&"")(0)=2 then error("此帖子已经在回收站了")
	
sql="Select * from ["&TablePrefix&"Posts] where PostID="&PostID&""
Rs.Open sql,Conn,1,3
	if Rs.eof or Rs.bof then error"<li>系统不存在该帖子的资料"
	if Rs("PostAuthor")<>CookieUserName and PermissionManage=0 then error("您的<a href=ShowForumPermissions.asp?ForumID="&ForumID&">权限</a>不够！")
	Subject=Rs("Subject")
	PostBody=Rs("Body")
Rs.close
if isNull(Subject) then Subject=""
if Request_Method <> "POST" then

%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> → <%=ForumTree(ParentID)%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> → <a href="ShowPost.asp?ThreadID=<%=ThreadID%>"><%=Topic%></a> → 删除帖子</div>

<script language="JavaScript" type="text/javascript">
var XmlDom = GetXmlDom();
XmlDom.async = false;
XmlDom.load("Xml/Templates.xml");
var XmlDomRoot = XmlDom.documentElement;
TemplateNode = XmlDomRoot.getElementsByTagName('template');

function PostDelReason(i) {  //原因
	var ElementNode = TemplateNode[i].getElementsByTagName('body')[0];
	ReasonText = GetNodeValue(ElementNode);
	ReasonText = ReasonText.replace('\[SiteName\]','<%=SiteConfig("SiteName")%>');
	$("ReasonBody").value = ReasonText;
}

function DelReasonOptions() { //原因下拉框
	for( var i=0;i<TemplateNode.length;i++) {
		var ElementNode = TemplateNode[i].getElementsByTagName('title')[0];
		document.write("<option value="+TemplateNode[i].getAttributeNode('id').nodeValue+">"+GetNodeValue(ElementNode)+"</option>");
	}
}
</script>
<form method="Post" action="?<%=Request.ServerVariables("Query_String")%>">
	<table cellspacing="1" cellpadding="5" width=100% class=CommonListArea>
		<tr class=CommonListTitle>
			<td colspan="2" align="center">删除帖子</td>
		</tr>
		<tr class="CommonListCell">
			<td width="200">标　题：</td>
			<td><%=Subject%></td>
		</tr>
		<tr class="CommonListCell">
			<td width="200">删除人：</td>
			<td><%=CookieUserName%></td>
		</tr>
		<tr class="CommonListCell">
			<td>删除帖子的原因：</td>
			<td><select name="PostDeleteReason" onchange="PostDelReason(this.options[this.selectedIndex].value)"><script language="JavaScript" type="text/javascript">DelReasonOptions()</script></select></td>
		</tr>
		<tr class="CommonListCell">
			<td>原因：</td>
			<td><textarea id="ReasonBody" name="ReasonBody" rows="12" cols="80"></textarea></td>
		</tr>
		<tr class="CommonListCell">
			<td colspan="2" align="center"><input type="submit" value=" 确定 " />　　　<input onclick="history.back()" type="button" value=" 取消 " /></td>
		</tr>
	</table>
</form>
<%
HtmlBottom
end if

sql="Select * from ["&TablePrefix&"Posts] where PostID="&PostID&""
Rs.Open sql,Conn,1
	if Rs.eof or Rs.bof then error"<li>系统不存在该帖子的资料"
	if Rs("PostAuthor")<>CookieUserName and PermissionManage=0 then error("您的<a href=ShowForumPermissions.asp?ForumID="&ForumID&">权限</a>不够！")
	PostID=Rs("PostID")
	ThreadID=Rs("ThreadID")
	ParentID=Rs("ParentID")
	PostAuthor=Rs("PostAuthor")
	PostDate=Rs("PostDate")
Rs.close
	
	if ParentID=0 then
		succtitle="删除主题成功"
		Execute("update ["&TablePrefix&"Users] Set TotalPosts=TotalPosts-1,UserMoney=UserMoney+"&SiteConfig("IntegralDeleteThread")&",experience=experience+"&SiteConfig("IntegralDeleteThread")&" where UserName='"&PostAuthor&"'")
		Execute("update ["&TablePrefix&"Threads] Set ThreadTop=0,Visible=2,DeletedCount=DeletedCount+1,LastName='"&CookieUserName&"',LastTime="&SqlNowString&" where ThreadID="&ThreadID&"")
		Execute("update ["&TablePrefix&"Posts] Set Visible=2 where PostID="&PostID&" and ParentID=0")
		Execute("update ["&TablePrefix&"Forums] Set TotalThreads=TotalThreads-1,TotalPosts=TotalPosts-1 where ForumID="&ForumID&"")
	else
		succtitle="删除回帖成功"
		
		

		
		Execute("update ["&TablePrefix&"Posts] Set Visible=2 where PostID="&PostID&" and ParentID>0")



		sql="Select top 1 * from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&" and Visible=1 order by PostID DESC"
		Rs.Open sql,Conn,1
			LastName=Rs("PostAuthor")
			LastTime=Rs("PostDate")
		Rs.close		

		Execute("update ["&TablePrefix&"Threads] Set TotalReplies=TotalReplies-1,DeletedCount=DeletedCount+1,LastName='"&LastName&"',LastTime='"&LastTime&"' where ThreadID="&ThreadID&"")
		Execute("update ["&TablePrefix&"Forums] Set TotalPosts=TotalPosts-1 where ForumID="&ForumID&"")
		
		Execute("update ["&TablePrefix&"Users] Set TotalPosts=TotalPosts-1,UserMoney=UserMoney+"&SiteConfig("IntegralDeletePost")&",experience=experience+"&SiteConfig("IntegralDeletePost")&" where UserName='"&PostAuthor&"'")
		
	end if
	
	UpForumMostRecent(ForumID)

		MailAddRecipient=Execute("Select UserEmail from ["&TablePrefix&"Users] where UserName='"&PostAuthor&"'")(0)
		LoadingEmailXml("MessageDeleted")
		MailBody=Replace(MailBody,"[Moderator]",CookieUserName)
		MailBody=Replace(MailBody,"[Topic]",Topic)
		MailBody=Replace(MailBody,"[Subject]",Subject)
		MailBody=Replace(MailBody,"[Body]",PostBody)
		MailBody=Replace(MailBody,"[DeleteReasons]",HTMLEncode(Request("ReasonBody")))
		SendMail MailAddRecipient,MailSubject,MailBody

if succtitle="" then error("无效命令")
Log(""&succtitle&"，标题："&Subject&"，主题ID："&ThreadID&"，帖子ID："&PostID&"")
Message="<li>"&succtitle&"<li><a href=ShowForum.asp?ForumID="&ForumID&">返回论坛</a>"
succeed Message,"ShowForum.asp?ForumID="&ForumID&""
%>