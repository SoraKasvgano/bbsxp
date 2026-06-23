<!-- #include file="Setup.asp" --><%
ForumID=RequestInt("ForumID")
ThreadID=RequestInt("ThreadID")
PostID=RequestInt("PostID")
PostAuthor=HTMLEncode(Request("PostAuthor"))

if Request("ViewMode")<>"" then  ResponseCookies "ViewMode",RequestInt("ViewMode"),"999"
if RequestCookies("ViewMode")="" then ResponseCookies "ViewMode",SiteConfig("ViewMode"),"999"
ViewMode=RequestCookies("ViewMode")


if PostID>0 then
	Set Rs=Execute("Select top 1 ThreadID,Visible from ["&TablePrefix&"Posts] where PostID="&PostID&"")
		if Rs.eof or Rs.bof then error"<li>ϵͳ�����ڸ����ӵ�����"
		if Rs("Visible")<>1 and BestRole<>1 then error"<li>�������ڻ���վ�У�"
		ThreadID=Rs("ThreadID")
	Rs.Close
	PostSql=" and PostID="&PostID&""
elseif PostAuthor>"" then
	PostSql=" and PostAuthor='"&SqlString(PostAuthor)&"'"
end if

if Request("menu")="Next" then
	sql="Select top 1 * from ["&TablePrefix&"Threads] where ThreadID > "&ThreadID&" and ForumID="&ForumID&" and Visible=1 order by ThreadID"
elseif Request("menu")="Previous" then
	sql="Select top 1 * from ["&TablePrefix&"Threads] where ThreadID < "&ThreadID&" and ForumID="&ForumID&" and Visible=1 order by ThreadID Desc"
else
	sql="Select top 1 * from ["&TablePrefix&"Threads] where ThreadID="&ThreadID&""
end if

Set Rs=Execute(sql)
	if Rs.eof or Rs.bof then error"<li>ϵͳ�����ڸ����ӵ�����"
	Topic=ReplaceText(Rs("Topic"),"<[^>]*>","")
	ThreadDescription=Rs("Description")
	TotalReplies=Rs("TotalReplies")
	TotalViews=Rs("TotalViews")
	IsVote=Rs("IsVote")
	IsGood=Rs("IsGood")
	IsLocked=Rs("IsLocked")
	Visible=Rs("Visible")
	Category=Rs("Category")
	ThreadID=Rs("ThreadID")
	ForumID=Rs("ForumID")
	ThreadStatus=Rs("ThreadStatus")
	UserName=Rs("PostAuthor")
	lastname=Rs("lastname")
	lasttime=Rs("lasttime")
	RatingSum=Rs("RatingSum")
	TotalRatings=Rs("TotalRatings")
	HiddenCount=Rs("HiddenCount")
	DeletedCount=Rs("DeletedCount")
	ThreadTop=Rs("ThreadTop")
	if DateDiff("d", Rs("StickyDate"), now())=0 then ThreadTop=0
Rs.close
UserNameHtml=Server.HTMLEncode(UserName)
UserNameUrl=Server.URLEncode(UserName)
LastNameHtml=Server.HTMLEncode(lastname)
LastNameUrl=Server.URLEncode(lastname)
TopicUrl=Server.URLEncode(Topic)
SortQueryString=Server.HTMLEncode(SafeJsString(ReplaceText(Request.QueryString,"&SortOrder=([0-9]*)","")))

if TotalRatings>0 then
	VoteRatings=RatingSum/TotalRatings
else
	VoteRatings=0
end if

%><!-- #include file="Utility/ForumPermissions.asp" --><%
HtmlHeadTitle=Topic
HtmlHeadDescription=ThreadDescription
HtmlTop

if Visible=2 and PermissionManage=0 then error"<li>�������ڻ���վ�У�"
if Visible=0 and PermissionManage=0 then error"<li>��������������У�"

Execute("update ["&TablePrefix&"Threads] Set TotalViews=TotalViews+1,LastViewedDate="&SqlNowString&",ThreadTop="&ThreadTop&" where ThreadID="&ThreadID&"")


AdvertisementGetRow=RequestApplication("Advertisements")		'ֱ�Ӷ�ȡApplication����

if IsArray(AdvertisementGetRow) then
	Randomize
	AdvertisementNum=Ubound(AdvertisementGetRow,2)+1
End if

%>
<!-- Markdown Support -->
<script src="js/marked.min.js"></script>
<script src="js/dompurify.min.js"></script>
<script src="js/markdown-handler.js"></script>
<link rel="stylesheet" href="css/markdown-content.css">

<script type="text/javascript" src="Utility/PopupMenu.js"></script>


<div class="CommonBreadCrumbArea">
	<div style="float:left"><%=ClubTree%> �� <%=ForumTree(ParentID)%><a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a> �� <a href="?ThreadID=<%=ThreadID%>"><%=Topic%></a></div>
	<div style="float:right">
		<a href="javascript:window.external.AddFavorite(location.href,document.title)" onmouseover="MouseOverOpen('FavoriteAllItem',this.id);" id="FavoriteAll"><img title="���ӵ��ղؼ�" src="images/favs.gif" border="0" /></a>��<script language="JavaScript" type="text/javascript">
		document.write("<a target=_blank href='Mailto:?subject="+document.title+"&body="+encodeURIComponent(location.href)+"'>");</script><img title="ͨ�������ʼ����ͱ�ҳ��" src="images/mail.gif" border="0" /></a>��<a href="javascript:window.print();"><img title="��ӡ��ҳ" src="images/Print.gif" border="0" /></a>��<a href="?menu=Previous&ForumID=<%=ForumID%>&ThreadID=<%=ThreadID%>"><img title="�����һƪ����" src="images/previous.gif" border="0" /></a>��<a href="?menu=Next&ForumID=<%=ForumID%>&ThreadID=<%=ThreadID%>"><img title="�����һƪ����" src="images/next.gif" border="0" /></a>
	</div>��
</div>


<div class="PopupMenu" id="FavoriteAllItem" style="display: none;">
	<table cellspacing="0" cellpadding="1">
		<tr>
			<td><a style="background-image:url(images/favorite.gif)" href="javascript:window.external.AddFavorite(location.href,document.title)">�����ղ�</a></td>
		</tr>
		<tr>
			<td><a style="background-image:url(images/qq_favorite.gif)" href="javascript:window.open('http://shuqian.qq.com/post?title='+encodeURIComponent(document.title)+'&uri='+encodeURIComponent(document.location.href)+'&jumpback=2&noui=1','favit','width=960,height=600,left=50,top=50,toolbar=no,menubar=no,location=no,scrollbars=yes,status=yes,resizable=yes');void(0);">�ѣ���ǩ</a></td>
		</tr>
		<tr>
			<td><a style="background-image:url(images/baidu_favorite.jpg)" href="javascript:window.open('http://cang.baidu.com/do/add?it='+encodeURIComponent(document.title.substring(0,76))+'&iu='+encodeURIComponent(location.href)+'&fr=ien#nw=1','_blank','scrollbars=no,width=600,height=450,left=75,top=20,status=no,resizable=yes');void(0);">�ٶ��Ѳ�</a></td>
		</tr>
		<tr>
			<td><a style="background-image:url(images/yahoo_favorite.gif)" href="javascript:window.open('http://myweb.cn.yahoo.com/popadd.html?url='+encodeURIComponent(document.location.href)+'&title='+encodeURIComponent(document.title), 'Yahoo','scrollbars=yes,width=780,height=550,left=80,top=80,status=yes,resizable=yes');void(0);">�Ż��ղ�</a></td>
		</tr>
	</table>
</div>

<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="2"><%=Topic%></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" width="5%"><img src="images/totel.gif" /></td>
		<td>�����ˣ�<a href="Profile.asp?UserName=<%=UserNameUrl%>"><%=UserNameHtml%></a>�����ظ�����<b><%=TotalReplies%></b>�����������<b><%=TotalViews%></b>���������£�<%=lasttime%>
		by <a href="Profile.asp?UserName=<%=LastNameUrl%>"><%=LastNameHtml%></a></td>
	</tr>
</table>
<br />

<div class="PopupMenu" id="View" style="DISPLAY: none">
	<table cellspacing="0" cellpadding="1">
		<tr>
			<td><a href="?ThreadID=<%=ThreadID%>&ViewMode=0">���ģʽ</a></td>
		</tr>
		<tr>
			<td><a href="?ThreadID=<%=ThreadID%>&ViewMode=1">����ģʽ</a></td>
		</tr>
	</table>
</div>

<table cellspacing="0" cellpadding="0" border="0" width="100%">
		<tr>
			<td>
			<%if CookieUserName<>"" then%>
			<a class="CommonTextButton"><script type="text/javascript" src="Utility/vote.js"></script><script language="JavaScript" type="text/javascript">ThreadID=<%=ThreadID%>;showData("<%=VoteRatings%>");</script></a>
			<%
				if SiteConfig("SelectMailMode")<>"" then
					if Execute("Select UserName from ["&TablePrefix&"Subscriptions] where UserName='"&SqlString(CookieUserName)&"' and ThreadID="&ThreadID&"").eof then
						BgImage="tracktopic.gif"
						ButtonText="��������"
					else
						BgImage="tracktopic-on.gif"
						ButtonText="ȡ������"
					end if
					%>
					<span id="ThreadSubscription"><a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/<%=BgImage%>)" href="javascript:Ajax_CallBack(false,'ThreadSubscription','Loading.asp?menu=Subscription&ThreadID=<%=ThreadID%>')"><%=ButtonText%></a></span>
					<%
				end if

			end if
			%>
			<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/view.gif)" onmouseover="MouseOverOpen('View',this.id);" id="View1">ѡ��鿴</a>
			<%if PermissionPost=1 then%><a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/NewPost.gif)" href="AddTopic.asp?ForumID=<%=ForumID%>">��������</a> <%end if%>
			</td>
			<td align="right" valign="bottom"><a href="http://www.duoci.com/Search/?Charset=<%=BBSxpCharset%>&word=<%=TopicUrl%>" target="_blank" title="�ڸ�����վ��������������">���������������</a>
			<%if SiteConfig("DisplayThreadStatus")=1 and (PermissionManage=1 or UserName=CookieUserName) then%>
			������״̬��<select onchange="javascript:if(this.options[this.selectedIndex].value)Ajax_CallBack(false,false,'loading.asp?menu=ThreadStatus&amp;ThreadID=<%=ThreadID%>&amp;ThreadStatus='+this.options[this.selectedIndex].value)">
			<option value="0" <%if threadstatus=0 then%>selected<%end if%>>--
			</option>
			<option value="1" <%if threadstatus=1 then%>selected<%end if%>>�ѽ��
			</option>
			<option value="2" <%if threadstatus=2 then%>selected<%end if%>>δ���
			</option>
			</select> <%end if%> ����������<select onchange="javascript:if(this.options[this.selectedIndex].value)window.location.href='ShowPost.asp?<%=SortQueryString%>&amp;SortOrder='+this.options[this.selectedIndex].value">
			<option value="0" <%if Request("sortorder")="0" then%>selected<%end if%>>
			�Ӿɵ���</option>
			<option value="1" <%if Request("sortorder")="1" then%>selected<%end if%>>
			���µ���</option>
			</select> </td>
		</tr>

</table>


<%

'''''''ͶƱ''''''''
if IsVote=1 then
%>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td width="40%" align="center">ѡ��</td>
		<td width="10%" align="center">Ʊ��</td>
		<td width="50%" align="center" colspan="2">�ٷֱ�</td>
	</tr>
	<form action="PostVote.asp?ThreadID=<%=ThreadID%>" method="Post">
		<%
sql="Select top 1 * from ["&TablePrefix&"Votes] where ThreadID="&ThreadID&""
Set Rs=Execute(sql)
	if Rs("IsMultiplePoll")=1 then
		multiplicity="checkbox"
	else
		multiplicity="radio"
	end if
	allticket=0
	Result=split(Rs("Result"),"|")
	for i = 0 to ubound(Result)
		if not Result(i)="" then allticket=Result(i)+allticket
	next
	Vote=split(Rs("Items"),"|")
	for i = 0 to ubound(Vote)
		if not Vote(i)="" then
			if allticket=0 then
				Votepercent=0
			else
				Votepercent=Formatpercent(result(i)/allticket)
			end if
%>
		<tr class="CommonListCell">
			<td width="40%">
			<input type="<%=multiplicity%>" value="<%=i%>" name="PostVote" id="PostVote<%=i%>" /><label for="PostVote<%=i%>"><%=Vote(i)%></label></td>
			<td width="10%" align="center"><%=Result(i)%></td>
			<td width="40%">
			<div class="percent">
				<div style="width:<%=Votepercent%>">
				</div>
			</div>
			</td>
			<td width="10%" align="center"><%=Votepercent%></td>
		</tr>
    <%
		end if
next
%>
		<tr class="CommonListCell">
			<td align="center"><%
	if  PermissionVote=0 then
		response.write "��û��Ȩ��ͶƱ"
	elseif CookieUserName=empty then
		response.write "���¼����ͶƱ"
	elseif instr("|"&Rs("BallotUserList")&"|","|"&CookieUserName&"|")>0 then
		response.write "���Ѿ�Ͷ��Ʊ��"
	elseif instr("|"&Rs("BallotIPList")&"|","|"&REMOTE_ADDR&"|")>0 then
		response.write "��IP�Ѿ�Ͷ��Ʊ��"
	elseif Rs("Expiry")< now() then
		response.write "ͶƱ�ѹ���"
	else
		response.write "<INPUT type=submit value='Ͷ��Ʊ'>"
	end if
%> </td>
			<td align="center">��Ʊ����<%=allticket%></td>
			<td colspan="2" align="center">��ֹͶƱʱ�䣺<%=Rs("Expiry")%></td>
		</tr>
	</form>
</table>
<%
Rs.Close
end if
'''''''ͶƱ END''''''''

if PermissionManage=1 then
	response.write("<form style='margin:0px;' method=POST action=Manage.asp>")
	response.write("<input type=hidden name='ThreadID' value='"&ThreadID&"' />")
	VisibleSql=""
	TotalCount=TotalReplies+DeletedCount+HiddenCount+1
else
	VisibleSql=" and Visible=1"
	TotalCount=TotalReplies+1
end if
if RequestInt("PostID")>0 then
	TotalCount=1		'�������������Ӳ���ҳ
elseif PostAuthor>"" then
	TotalCount=Execute("Select Count(PostID) from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&PostSql&VisibleSql&"")(0)
end if
if TotalCount<1 then
	IsResponseTop=0
	response.clear()
	error"<li>ϵͳ�����ڸ�������������"
end if

PageSetup=SiteConfig("PostsPerPage") '�趨ÿҳ����ʾ����
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
PageCount = RequestInt("PageIndex") '��ȡ��ǰҳ
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage

if Request("SortOrder")="1" then
	SqlSortOrder="Desc"
else
	SqlSortOrder=""
end if

'������˹���
	sql="Select top "&PageCount*pagesetup&" PostID,ThreadID,ParentID,PostAuthor,Subject,Body,IPAddress,PostDate,Visible from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&PostSql&VisibleSql&" order by PostID "&SqlSortOrder&""
	Set Rs=Execute(sql)

	If Rs.Eof and ""&PostAuthor&""="" then Execute("Delete from ["&TablePrefix&"Threads] where ThreadID="&ThreadID&"")
	If TotalPage>1 then RS.Move (PageCount-1) * pagesetup
	IF Not Rs.Eof then PostGetRows=Rs.GetRows(pagesetup)
	Rs.close



if IsArray(PostGetRows) then
	For i=0 To Ubound(PostGetRows,2)
		PostID=PostGetRows(0,i)
		ThreadID=PostGetRows(1,i)
		ParentID=PostGetRows(2,i)
		PostAuthor=PostGetRows(3,i)
		Subject=PostGetRows(4,i)
		Body=PostGetRows(5,i)
		IPAddress=PostGetRows(6,i)
		PostDate=PostGetRows(7,i)
		Visible=PostGetRows(8,i)


		if ViewMode=0 then
			ShowPostSimple
		else
			ShowPost
		end if

	Next
End if
AdvertisementGetRow=null
PostGetRows=null



if PermissionManage=1 then%>
<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
		<td align="right" style="Padding-Top:8px">
        ������
		<select name=menu size=1>
			<optgroup label="ѡ��">
					<option value="DelPost">ɾ������</option>
					<option value="UnDelPost">��ɾ������</option>
					<option value="PostVisible">ͨ���������</option>
					<option value="PostInVisible">ȡ���������</option>
				</optgroup>
		</select>��<input type="submit" value=" ִ �� " onclick="return VerifyRadio('Item');" />��<input type=checkbox name=chkall onclick='CheckAll(this.form)' value=ON>
        </td>
	</tr>
</table>
<%
response.write("</form>")
end if
%>
<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr><td><%ShowPage()%></td></tr>
</table>

<%
if IsLocked=0 and PermissionReply=1 and CookieUserName<>empty then
%>
<a name="QuickReply"></a>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<form name="form" method="POST" action="AddPost.asp" onsubmit="return CheckForm(this);">
	<input type="hidden" value="<%=ThreadID%>" name="ThreadID">
	<input type="hidden" value="<%=PostID%>" name="PostID">
	<input name="Body" type="hidden">
	<tr class="CommonListTitle">
		<td colspan="2"><b>���ٻظ�</b></td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top" height="50%" width="180">
				<br /><b>��������</b><br />��<a href="javascript:CheckLength();">�鿴���ݳ���</a>��<br /><br />
				<input id="DisableBBCode" name="DisableBBCode" type="checkbox" value="1"><label for="DisableBBCode"> ���� BB ����</label>
		</td>
		<td height="200"><script type="text/javascript" src="Editor/Post.js"></script></td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top" height="50%" colspan="2" align="center">
		<input type="submit" accesskey="s" title="(Alt + S)" value=" �ظ� " name="EditSubmit">��<input type="Button" value=" Ԥ�� " onclick="Preview()">��<input onclick="history.back()" type="button" value=" ȡ�� ">��<input type="button" id="recoverdata" onclick="RestoreData()" title="�ָ��ϴ��Զ����������" value="�ָ�����" /></td>
	</tr>
	</form>
</table>
<div name="Preview" id="Preview"></div>
<script language="JavaScript" type="text/javascript">
function QuickReply(NO){
	document.form.PostID.value =NO;
	window.location="#QuickReply";
	BBSXPEditorForm.focus();
}
function VerifyRadio() {
	objYN=false;
	if (window.confirm('��ȷ��ִ�б��β���?')){
		for (i=0;i<document.getElementsByName("PostID").length;i++) {
			if (document.getElementsByName("PostID")[i].checked) {objYN= true;}
		}
		if (objYN==false) {alert ('��ѡ����Ҫ���������ӣ�');return false;}
	}
	return objYN;
}
function showPostText(PostID){
	ToggleMenuOnOff('PostContent_'+PostID);
	ToggleMenuOnOff('PostContent'+PostID);
	ToggleMenuOnOff('PostTitle'+PostID);
}
</script>
<%
end if

%><!-- #include file="Utility/OnLine.asp" -->
<%
if SiteConfig("DisplayThreadUsers")=1 then
	ForumIDOnline=Execute("Select count(sessionid) from ["&TablePrefix&"UserOnline] where ForumID="&ForumID&"")(0)
	ThreadIDOnline=Execute("Select count(sessionid) from ["&TablePrefix&"UserOnline] where ThreadID="&ThreadID&"")(0)
	regThreadIDOnline=Execute("Select count(sessionid) from ["&TablePrefix&"UserOnline] where ThreadID="&ThreadID&" and UserName<>''")(0)
%>

<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td>�û�������Ϣ</td>
	</tr>
	<tr class="CommonListCell">
	<td>
		<img src="images/plus.gif" id="followImg" style="cursor:pointer;" onclick="loadThreadFollow('ThreadID=<%=ThreadID%>')" />
		��ǰ�鿴������Ļ�Ա: <b><%=ThreadIDOnline%></b> �ˡ�����ע���û� <b><%=regThreadIDOnline%></b> �ˣ��ÿ� <b><%=ThreadIDOnline-regThreadIDOnline%></b>
		�ˡ�<div style="display:none" id="follow">
			<hr width="90%" size="1" align="left"><span id="followTd" class="UserList">
			<img src="images/loading.gif" />���ڼ���...</span></div>
		</td>
	</tr>
</table>
<br />
<%
end if

if PermissionManage=1 then
	response.write "<br><table cellspacing=0 cellpadding=0 width=100% align=center border=0><tr><td align=center>����ѡ�"
	if ThreadTop=2 then
		response.write "<a href=javascript:UrlPost('Manage.asp?menu=UnTop&ThreadID="&ThreadID&"')>ȡ������</a> | "
	else
		response.write "<a href=javascript:UrlPost('Manage.asp?menu=Top&ThreadID="&ThreadID&"')>��Ϊ����</a> | "
	end if

	if ThreadTop=1 then
		response.write "<a href=javascript:UrlPost('Manage.asp?menu=DelTop&ThreadID="&ThreadID&"')>ȡ���ö�</a> | "
	else
		response.write "<a href=javascript:UrlPost('Manage.asp?menu=ThreadTop&ThreadID="&ThreadID&"')>�ö�����</a> | "
	end if
	response.write "<a href=javascript:UrlPost('Manage.asp?menu=MoveNew&ThreadID="&ThreadID&"')>��ǰ����</a> | "
	if IsLocked=1 then
		response.write "<a href=javascript:UrlPost('Manage.asp?menu=DelIsLocked&ThreadID="&ThreadID&"')>��������</a> | "
	else
		response.write "<a href=javascript:UrlPost('Manage.asp?menu=IsLocked&ThreadID="&ThreadID&"')>��������</a> | "
	end if
	if IsGood=1 then
		response.write "<a href=javascript:UrlPost('Manage.asp?menu=DelIsGood&ThreadID="&ThreadID&"')>ȡ����������</a>"
	else
		response.write "<a href=javascript:UrlPost('Manage.asp?menu=IsGood&ThreadID="&ThreadID&"')>��Ϊ��������</a>"
	end if
	response.write " | <a href=MoveThread.asp?ThreadID="&ThreadID&">�ƶ�����</a> | <a title='�޸����ӵĻظ���' href=javascript:UrlPost('Manage.asp?menu=Fix&ThreadID="&ThreadID&"')>�޸�����</a></td></tr></table>"
end if


HtmlBottom



Sub ShowPostSimple()
	PostAuthorHtml=Server.HTMLEncode(PostAuthor)
	PostAuthorUrl=Server.URLEncode(PostAuthor)
	IPAddressHtml=Server.HTMLEncode(IPAddress)

%>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea" style="TABLE-LAYOUT:fixed;">
	<tr class="CommonListHeader">
		<td>
			<div style=float:left><b><a target="_blank" href="Profile.asp?UserName=<%=PostAuthorUrl%>"><%=PostAuthorHtml%></a></b> ������ <%=PostDate%></div>
			<div style=float:right>
						<%if IsLocked=1 then%>
						<a onclick="window.alert('�����������������ظ���');">����</a>
						<%elseif PermissionReply=1 then%>
						<a href="AddPost.asp?ThreadID=<%=ThreadID%>&PostID=<%=PostID%>" title="�ظ�����">�ظ�</a>
    <%end if%>
						<a href="EditPost.asp?ThreadID=<%=ThreadID%>&PostID=<%=PostID%>" title="�༭����">�༭</a>
						<a href="DelPost.asp?ThreadID=<%=ThreadID%>&PostID=<%=PostID%>" title="ɾ������">ɾ��</a>
			</div>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td>
<%
		if Subject<>"" then response.write "<div class=ForumPostTitle>"&Subject&"</div>"
		response.write "<div class='ForumPostContentText markdown-content'>"&BBCode(Body)&"</div>"
		%>

				<div style="float:right">
				<%if PermissionManage=1 then response.write "IP��"&IPAddressHtml&"��"%>
				<%if IsLocked=0 and PermissionReply=1 then response.write "<a onclick=javascript:QuickReply("&PostID&")>���ٻظ�</a>"%>
				</div>


		</td>
	</tr>
</table>
<%


End Sub




Sub ShowPost()

	Set Rs=Execute("Select top 1 * from ["&TablePrefix&"Users] where UserName='"&SqlString(PostAuthor)&"'")
	if Rs.EOF then
		Rs.close
		Exit Sub
	End if
	UserFaceUrl=SafeUrl(Rs("UserFaceUrl"))
	WebAddress=SafeUrl(Rs("WebAddress"))
	WebLog=SafeUrl(Rs("WebLog"))
	WebGallery=SafeUrl(Rs("WebGallery"))
	PostAuthorHtml=Server.HTMLEncode(PostAuthor)
	PostAuthorUrl=Server.URLEncode(PostAuthor)
	PostUserName=Rs("UserName")
	PostUserNameHtml=Server.HTMLEncode(PostUserName)
	PostUserNameUrl=Server.URLEncode(PostUserName)
	PostUserEmailHtml=Server.HTMLEncode(Rs("UserEmail"))
	WebAddressHtml=Server.HTMLEncode(WebAddress)
	WebLogHtml=Server.HTMLEncode(WebLog)
	WebGalleryHtml=Server.HTMLEncode(WebGallery)
	IPAddressHtml=Server.HTMLEncode(IPAddress)
ReportSubject=Server.URLEncode("�������ӱ���")
ReportBody=Server.URLEncode("���������ӡ���"&SiteConfig("SiteUrl")&"/ShowPost.asp?PostID="&PostID)
%>
<div class="PopupMenu" id="ContactMenu<%=PostID%>" style="DISPLAY: none">
	<table cellspacing="0" cellpadding="1">
		<tr>
			<td><a style="BACKGROUND-IMAGE:url(images/profile.gif)" href="Profile.asp?UID=<%=Rs("UserID")%>">�鿴 <%=PostUserNameHtml%> ������</a></td>
		</tr>
		<%if CookieUserName<>"" then%><tr>
			<td><a style="BACKGROUND-IMAGE:url(images/privatemessage.gif)" href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&RecipientUserName=<%=PostUserNameUrl%>', 600, 350);">�� <%=PostUserNameHtml%> ����ѶϢ</a></td>
		</tr>
		<%end if%>
		<tr>
			<td><a style="BACKGROUND-IMAGE:url(images/email.gif)" href="Mailto:<%=PostUserEmailHtml%>">�� <%=PostUserNameHtml%> �����ʼ�</a></td>
		</tr>
		<%if WebAddress<>"" then%><tr>
			<td><a style="BACKGROUND-IMAGE:url(images/homepage.gif)" href="<%=WebAddressHtml%>" target="_blank">��� <%=PostUserNameHtml%> ����ҳ</a></td>
		</tr>
		<%end if%> <%if WebLog<>"" then%><tr>
			<td><a style="BACKGROUND-IMAGE:url(images/weblog.gif)" href="<%=WebLogHtml%>" target="_blank">��� <%=PostUserNameHtml%> �Ĳ���</a></td>
		</tr>
		<%end if%> <%if WebGallery<>"" then%><tr>
			<td><a style="BACKGROUND-IMAGE:url(images/webgallery.gif)" href="<%=WebGalleryHtml%>" target="_blank">��� <%=PostUserNameHtml%> �����</a></td>
		</tr>
		<%end if%>
		<tr>
			<td><a style="BACKGROUND-IMAGE:url(images/search.gif)" href="ShowBBS.asp?menu=MyTopic&UserName=<%=PostUserNameUrl%>">���� <%=PostUserNameHtml%> ������</a></td>
		</tr>
	</table>
</div>
<%if ""&CookieUserName&""<>"" then%>
<div class="PopupMenu" id="FavoriteMenu<%=PostID%>" style="DISPLAY: none">
	<table cellspacing="0" cellpadding="1">
		<tr>
			<td><a style="BACKGROUND-IMAGE:url(images/favorite.gif)" href="javascript:Ajax_CallBack(false,false,'MyFavorites.asp?menu=FavoriteFriend&FriendUserName=<%=PostUserNameUrl%>',true);">�� <%=PostUserNameHtml%> ��Ϊ����</a></td>
		</tr>
		<tr>
			<td><a style="BACKGROUND-IMAGE:url(images/favorite.gif)" href="javascript:Ajax_CallBack(false,false,'MyFavorites.asp?menu=FavoritePost&PostID=<%=PostID%>',true);">�������Ӽ����ղؼ�</a></td>
		</tr>
		<tr>
			<td><a style="BACKGROUND-IMAGE:url(images/favorite.gif)" href="javascript:Ajax_CallBack(false,false,'MyFavorites.asp?menu=FavoriteForums&ForumID=<%=ForumID%>',true);">������̳�����ղؼ�</a></td>
		</tr>
	</table>
</div>
<%end if%>

<a name="<%=PostID%>"></a>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td>
			<div style=float:left><img src="images/icon_post_show.gif" /> <%=PostDate%></div>
			<div style=float:right>[<a href="?ThreadID=<%=ThreadID%>&PostAuthor=<%=PostAuthorUrl%>" title="ֻ�������ߵ�����">ֻ��������</a>] <a href="?PostID=<%=PostID%>" title="ֻ��������">#<%=i+(PageCount-1)*PageSetup+1%></a><%if PermissionManage=1 then response.write("<input type=checkbox value="&PostID&" name=PostID onclick="&chr(34)&"CheckSelected(this.form,this.checked,'Post"&PostID&"')"&chr(34)&">")%></div>
		</td>
	</tr>
	<tr class="CommonListCell" id="Post<%=PostID%>">
		<td>


		<table width="100%" border="0" cellpadding="0" cellspacing="0" style="padding-left:5px;TABLE-LAYOUT:fixed;">
			<tr>
				<td rowspan="2" valign="top" class="ForumPostUserArea">

				<div style="text-align:left;width:90%;">
					<div style="float:left">
						<%if DateDiff("n",Rs("UserActivityTime"),Now()) < SiteConfig("UserOnlineTime") then%>
						<img title="<%=PostUserNameHtml%> ����. ���ʱ��:<%=Server.HTMLEncode(Rs("UserActivityTime"))%>" src="Images/user_IsOnline.gif" border="0" />
				    <%end if%>
						<font style="font-size:10pt"><b><%=PostUserNameHtml%></b></font><br /><%=Server.HTMLEncode(Rs("UserTitle"))%>
					</div>
					<%if SiteConfig("EnableReputation")=1 then%>
					<div style="float:right">
						<a href="javascript:BBSXP_Modal.Open('Reputation.asp?CommentFor=<%=PostUserNameUrl%>',550,200);"><img title="�� <%=PostUserNameHtml%> ������������" src="Images/reputation.gif" border="0" align="absmiddle" /></a>
					</div>
				<%
					end if
				response.write "<br /><br /><br /><div style='text-align:center;'>"

				if SiteConfig("EnableAvatars")=1 and SiteConfig("AllowAvatars")=1 and UserFaceUrl<>"" then response.write "<img src='"&UserFaceUrl&"' style='max-width:"&SiteConfig("AvatarWidth")&"px;max-height:"&SiteConfig("AvatarHeight")&"px;'  /><br />"
				if Rs("UserRank")<>"" then response.write "<br />"&Rs("UserRank")&"<br />"

				response.write "<br /></div>�ǡ���ɫ��"

				if instr("|"&Moderated&"|","|"&Rs("UserName")&"|") > 0 then
					Response.Write "����"
				else
					response.write ShowRole(Rs("UserRoleID"))
				end if

				if Rs("UserMate")<>"" then response.write "<br />�䡡��ż��"&Server.HTMLEncode(Rs("UserMate"))&""
				response.write "<br />�� �� ����"&Rs("TotalPosts")&""
				response.write "<br />�� �� ֵ��"&Rs("experience")&""
				response.write "<br />ע��ʱ�䣺"&FormatDateTime(Rs("UserRegisterTime"),2)&"<br />"

				response.write "<img src=images/money.gif title='���:"&Rs("UserMoney")&"'> "&ShowReputation(Rs("Reputation"))&" "&ShowUserActivityDay(Rs("UserActivityDay"))&" "&Horoscope(Rs("birthday"))&" "&ShowUserSex(Rs("UserSex"))&""

				%>
				</div>

				</td>
				<td valign="top">
				<div class=ForumPostButtons>
					<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/contact.gif)" onmouseover="MouseOverOpen('ContactMenu<%=PostID%>',this.id);" id="Contact<%=PostID%>">��ϵ</a>
					<%if ""&CookieUserName&""<>"" then%><a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/favorite.gif)" onmouseover="MouseOverOpen('FavoriteMenu<%=PostID%>',this.id);" id="Favorite<%=PostID%>">�ղ�</a> <%end if%>
					<%if IsLocked=1 then%>
					<a class="CommonImageTextButton" style="BACKGROUND-IMAGE:url(images/locked.gif)" onclick="window.alert('�����������������ظ���');">����</a>
					<%elseif PermissionReply=1 then%>
					<a class="CommonImageTextButton" style="BACKGROUND-IMAGE:url(images/NewPost.gif)" href="AddPost.asp?ThreadID=<%=ThreadID%>&PostID=<%=PostID%>" title="�ظ�����">�ظ�</a>
		    <%end if%>
						<a class="CommonImageTextButton" style="BACKGROUND-IMAGE:url(images/edit.gif)" href="EditPost.asp?ThreadID=<%=ThreadID%>&PostID=<%=PostID%>" title="�༭����">�༭</a>
						<a class="CommonImageTextButton" style="BACKGROUND-IMAGE:url(images/delete.gif)" href="DelPost.asp?ThreadID=<%=ThreadID%>&PostID=<%=PostID%>" title="ɾ������">ɾ��</a>
				</div>

				<div class="ForumPostBodyArea">


				<%
				if Rs("UserAccountStatus")=2 then
					Response.Write "==============================<br /><font color=RED>���û��ʺ��ѱ�����</font><br />=============================="
				elseif Rs("Reputation") < SiteConfig("InPrisonReputation") then
					Response.Write "<div id='PostContent_"&PostID&"'>=============================================<br /><font color=RED>���û�����С��"&SiteConfig("InPrisonReputation")&"�����������ѱ�����.</font>��<a onclick='showPostText("&PostID&")'>����鿴</a><br />=============================================</div>"
					response.write "<div class=ForumPostTitle id='PostTitle"&PostID&"' style='display:none'>"&Subject&"</div><div class='ForumPostContentText markdown-content' id='PostContent"&PostID&"' style='display:none'>"&BBCode(Body)&"</div>"
				elseif Visible=2 then
					Response.Write "<div id='PostContent_"&PostID&"'>========================<br /><font color=RED>�����ѱ�ɾ����</font>��<a onclick='showPostText("&PostID&")'>����鿴</a><br />========================</div>"
					response.write "<div class=ForumPostTitle id='PostTitle"&PostID&"' style='display:none'>"&Subject&"</div><div class='ForumPostContentText markdown-content' id='PostContent"&PostID&"' style='display:none'>"&BBCode(Body)&"</div>"
				else
					response.write "<div class=ForumPostTitle>"&Subject&"</div><div class='ForumPostContentText markdown-content'>"&BBCode(Body)&"</div>"

					if SiteConfig("EnableSignatures")=1 and SiteConfig("AllowSignatures")=1 then
						if Rs("UserSign")<>"" then response.write "<div class=ForumPostSignature>"&BBCode(Rs("UserSign"))&"</div>"
					end if


					sql="select * from ["&TablePrefix&"PostInTags] where PostID="&PostID&""
					Set RsTag=Execute(sql)
					do while not RsTag.eof
						Tags=Tags&",<a href='Tags.asp?TagID="&RsTag("TagID")&"'>"&Server.HTMLEncode(Execute("Select TagName from ["&TablePrefix&"PostTags] where TagID="&RsTag("TagID")&"")(0))&"</a>"
						RsTag.movenext
					Loop
					RsTag.Close
					Set RsTag = Nothing
					if ""&Tags&""<>"" then Response.Write("<p>��ǩ��"&Mid(Tags,2))&"</p>"


					if SiteConfig("DisplayEditNotes")=1 then
						Set EditNotesRs=Execute("Select * from ["&TablePrefix&"PostEditNotes] where PostID="&PostID&"")
						If Not EditNotesRs.eof Then EditNotesRecordset=EditNotesRs("EditNotes")
						EditNotesRs.Close
						Set EditNotesRs = Nothing
						Response.Write("<p>"&Replace(Server.HTMLEncode(EditNotesRecordset),vbCrLf,"<br />")&"</p>")
					end if


				end if
				%>
				</div>

				</td>
			</tr>
			<tr>
				<td valign="bottom">
				<table width="100%">
				<tr>
					<td valign="bottom">
				<%
					If AdvertisementNum>0 then
						RndValue=Int(AdvertisementNum * Rnd)
						Response.Write(AdvertisementGetRow(0,RndValue))
					end if
				%>
					</td>
					<td align="right" valign="bottom"><%
					if CookieUserName<>empty then%>
						<a onclick="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&ForumID=<%=ForumID%>&subject=<%=ReportSubject%>&Body=<%=ReportBody%>', 600, 350);"><img title="���汾��" src="images/feedback.gif" border="0" /></a>��<%
					end if
					if PermissionManage=1 then
						if Visible=0 then response.write("<img src='images/InVisible.gif' border=0 alt='����δͨ�����' title='����δͨ�����' />��")
						if Visible=2 then response.write("<img src='images/recycle.gif' border=0 alt='������ɾ��' title='������ɾ��' />��")
						response.write("<img src='images/IP.gif' border=0 alt='"&IPAddressHtml&"' title='"&IPAddressHtml&"' />��")
					end if
					if IsLocked=0 and PermissionReply=1 and CookieUserName<>empty then response.write "<a href=AddPost.asp?ThreadID="&ThreadID&"&PostID="&PostID&"&Quote=1><img src=images/Quote.gif alt='���ûظ�' title='���ûظ�' border=0 /></a>��<a onclick=javascript:QuickReply("&PostID&")><img src=images/QuickReply.gif alt='���ٻظ�' title='���ٻظ�' /></a>"%>
					</td>
				</tr>
				</table>
				</td>
			</tr>
		</table>

		</td>
	</tr>
</table>

<%
Rs.close
End Sub
%>

<!-- Render Markdown content -->
<script>
if (typeof renderMarkdown === 'function') {
    renderMarkdown('.markdown-content');
}
</script>
