<!-- #include file="Conn.asp" -->
<%
ii=0

if BBSxpDataBaseVersion<>ForumsDataBaseVersion then error("���������ݿ�汾�����ϣ���������ݿ�汾�� "&ForumsDataBaseVersion&" ��")

if SiteConfig("EnableBannedUsersToLogin")=0 and CookieUserAccountStatus=2 then error("�����ʺ��ѱ����ã�")

if SiteConfig("BannedIP")<>"" then
	filtrate=split(SiteConfig("BannedIP"),"|")
	for i = 0 to ubound(filtrate)
		if instr("|"&REMOTE_ADDR&"","|"&filtrate(i)&"") > 0 then error(""&REMOTE_ADDR&"����ֹ������̳��")
	next
end if

if SiteConfig("SiteDisabled")=1 then
	if instr(Script_Name,"admin_") = 0 and instr(Script_Name,"login.asp")=0 and instr(Script_Name,"install.asp")=0 then error(""&SiteConfig("SiteDisabledReason")&"")
end if

if Request_Method = "POST" then
	if instr(""&Http_Referer&"","http://"&Server_Name&"") = 0 and instr(Script_Name,"login.asp")=0 then error("<li>����ҳ����"&Http_Referer&"<li>ϵͳ�޷�ʶ����������ҳ��<li>�������رշ���ǽ�����ύ����Ϣ��")
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
<div class="CommonTopFrameText"><script language="JavaScript" type="text/javascript">if (top !== self) {document.write('<a target="_top" href="'+location.href+'">ƽ��ģʽ</a>');}else {document.write('<a target="_top" href="Frame.asp">����ģʽ</a>');}</script></div>
<div class="CommonTopMenuText">


<%
if CookieUserName=empty then
%><a href="javascript:BBSXP_Modal.Open('Login.asp',380,170);">��¼</a> | <a href="CreateUser.asp">ע��</a> | <a href="ViewOnline.asp">�������</a> | <a href="Search.asp?ForumID=<%=RequestInt("ForumID")%>">����</a> | <a href="Help.asp">����</a><%
else
%>&raquo;��<a href="Profile.asp?UID=<%=CookieUserID%>"><%=CookieUserName%></a> | <a onmouseover="showmenu(event,'<div class=menuitems><a href=javascript:UrlPost(&quot;Cookies.asp?menu=Online&quot;)>����</a></div><div class=menuitems><a href=javascript:UrlPost(&quot;Cookies.asp?menu=Invisible&quot;)>����</a></div>')" style="cursor:default">
<%
	if RequestCookies("Invisible")="1" then
		response.write "����"
	else
		response.write "����"
	end if
	
	response.write("</a> | ")

	if SiteConfig("AccountActivation")=2 then response.write("<a href=InviteUser.asp>�����û�</a> | ")
%>
			<a onMouseOver="showmenu(event,'<div class=menuitems><a href=EditProfile.asp>�༭����</a></div><div class=menuitems><a href=EditProfile.asp?menu=pass>�����޸�</a></div><div class=menuitems><a href=MyUpFiles.asp>�ϴ�����</a></div><div class=menuitems><a href=MyFavorites.asp>�� �� ��</a></div><div class=menuitems><a href=MyMessage.asp>���ŷ���</a></div>')" style="cursor:default">�������</a> |
			<a onMouseOver="showmenu(event,'<div class=menuitems><a href=ShowBBS.asp>��������</a></div><div class=menuitems><a href=ShowBBS.asp?menu=HotViews>��������</a></div><div class=menuitems><a href=ShowBBS.asp?menu=HotReplies>��������</a></div><div class=menuitems><a href=ShowBBS.asp?menu=NoReplies>δ�ظ���</a></div><div class=menuitems><a href=ShowBBS.asp?menu=GoodTopic>��������</a></div><div class=menuitems><a href=ShowBBS.asp?menu=VoteTopic>ͶƱ����</a></div><div class=menuitems><a href=ShowBBS.asp?menu=MyTopic&amp;UserName=<%=CookieUserName%>>�ҵ�����</a></div><div class=menuitems><a href=ShowBBS.asp?menu=MyPost&amp;UserName=<%=CookieUserName%>>�ҵĻ���</a></div><div class=menuitems><a href=ShowBBS.asp?menu=MySubscriptions&amp;UserName=<%=CookieUserName%>>�ҵĶ���</a></div>')" style="cursor:default">��������</a> |
			<a onMouseOver="showmenu(event,'<div class=menuitems><a href=ViewOnline.asp>�������</a></div><div class=menuitems><a href=ViewOnline.asp?menu=cutline>����ͼ��</a></div><div class=menuitems><a href=ViewOnline.asp?menu=UserSex>�Ա�ͼ��</a></div><div class=menuitems><a href=ViewOnline.asp?menu=TodayPage>����ͼ��</a></div><div class=menuitems><a href=ViewOnline.asp?menu=board>����ͼ��</a></div><div class=menuitems><a href=ViewOnline.asp?menu=TotalPosts>����ͼ��</a></div><%if SiteConfig("PublicMemberList")=1 then%><div class=menuitems><a href=Members.asp>��Ա�б�</a></div><%end if%>')" style="cursor:default">ͳ��</a> | 
			<a href="Search.asp?ForumID=<%=RequestInt("ForumID")%>">����</a><span id="MenuListID"></span>
			<script language="javascript" type="text/javascript">ShowMenuList("Xml/Menu.xml");</script>
<%
	response.write " | <a href=Help.asp>����</a> | <a onclick='return log_out()'>�˳�</a>"
	
	if BestRole=1 then
	%> | <a onMouseOver="showmenu(event,'<div class=menuitems><a href=Moderation.asp>�������</a></div><div class=menuitems><a href=Moderation.asp?menu=Visible>�� �� ��</a></div><div class=menuitems><a href=Moderation.asp?menu=Recycle>�� �� վ</a></div>')" style="cursor:default">���</a><%
	end if
	if CookieUserRoleID = 1 then response.write " | <a href=Admin_Default.asp target=_top>����</a>"
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
			<a href="SendMessage.asp">��ϵ����</a> - <a href="<%=SiteConfig("CompanyURL")%>" target="_blank"><%=SiteConfig("CompanyName")%></a> - <a href="Archiver.asp?<%=Query_String%>">��̳�浵</a> - <a href="#top">���ض���</a><br />
			Powered by <a target="_blank" href="http://www.bbsxp.com"><%=ForumsProgram%></a> &copy; 1998-<%=year(now)%> <a href="http://www.yuzi.net/" target="_blank">Yuzi.Net</a>
        </li>
        <li style="text-align:right;width:29%;">
            Processed in <%=ProcessTime%> second(s)<br />Server Time <%=now()%>
        </li>
	</div>
</div>
<%
response.write "<!-- SQL��ѯ("&SqlQueryNum&") -->"

if CookieUserName<>"" then
	if CookieNewMessage > 0 then
		AddApplication "Message_"&CookieUserName,"��ѶϢ֪ͨ��<a target='_blank' href='MyMessage.asp'>���� "&CookieNewMessage&" ����ѶϢ����ע�����</a>"
		Execute("update ["&TablePrefix&"Users] Set NewMessage=0 where UserID="&CookieUserID&"")
	end if

if RequestApplication("Message_"&CookieUserName)<>"" then
%>
<div id="ApplicationMsg" style=" Z-INDEX:1;WIDTH:250px;VISIBILITY:hidden;position:absolute;" onClick="if(typeof(IsReadMsg) == 'undefined')Ajax_CallBack(false,'MsgTitle','Loading.asp?menu=CloseMsg');IsReadMsg=true;">
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
<tr class=CommonListTitle><td><div style=float:left id='MsgTitle'>��ʱѶϢ</div><div style=float:right><a href="javascript:MsgClose()">��</a></div></td></tr>
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
	response.write "<br /><br /><input onclick='history.go(-1)' type='button' value=' &lt;&lt; �� �� '>����<input onclick='history.go(1)' type='button' value='  ǰ �� &gt;&gt;'><br />"
	response.write "<hr width=80% SIZE=0 class=CommonListArea />Powered by <a target=_blank href=http://www.bbsxp.com>"&ForumsProgram&"</a> &nbsp;&copy; 1998-"&year(now)&"<br />Server Time "&now()&"<br />Processed in "&ProcessTime&" second(s)</div></body></html>"
End Sub
''''''''''''''''''''''''''''''''
Function succeed(Message,Url)

if IsResponseTop<>1 then HtmlTop
if Left(Message,4)<>"<li>" then Message="<li>"&Message
%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> �� ��̳��Ϣ</div>
<table cellspacing="1" cellpadding="5" width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td width="100%" align="center">��̳��ʾ��Ϣ</td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top" align="Left" height="122">
		<table cellspacing="0" cellpadding="5" width="100%" border="0">
			<tr>
				<td width="83%" valign="top">
					<b><span id="yu">0</span>���Ӻ�ϵͳ���Զ�����...</b>
					<ul>
					<%=Message%>
					<li><a href=Default.asp>������̳��ҳ</a></li>
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
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� ��̳��Ϣ</div>
<table cellspacing="1" cellpadding="5" width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td width="100%" colspan="2" align="center">��̳��ʾ��Ϣ</td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top" align="Left" colspan="2" height="122">
		<table cellspacing="0" cellpadding="5" width="100%" border="0">
			<tr>
				<td width="83%" valign="top"><b>�������ɹ��Ŀ���ԭ����飺</b><ul><%=Message%></ul></td>
				<td width="17%"><img height="97" src="images/err.gif" width="95" /></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr align="center" class="CommonListCell">
		<td valign="center" colspan="2"><input onClick="history.back()" type="submit" value=" << �� �� �� һ ҳ " /></td>
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
		SubForumStrings="<br /><b>�Ӱ��棺"
		for iii=0 to Ubound(SubForumGetRow,2)
			if SubForumGetRow(2,iii)<>"" then
				SubForumStrings=SubForumStrings&"<a href="&SubForumGetRow(2,iii)&">"&SubForumGetRow(1,iii)&"</a>��"
			else
				SubForumStrings=SubForumStrings&"<a href=ShowForum.asp?ForumID="&SubForumGetRow(0,iii)&">"&SubForumGetRow(1,iii)&"</a>��"
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
		IconImage="topic-announce.gif alt='��������'"
	elseif Rs("ThreadTop")=1 then
		IconImage="topic-pinned.gif alt='�ö�����'"
	elseif Rs("IsGood")=1 then
		IconImage="topic-popular.gif alt='��������'"
	elseif Rs("IsLocked")=1 then
		IconImage="topic-locked.gif alt='��������'"
	elseif Rs("IsVote")=1 then
		IconImage="topic-poll.gif alt='ͶƱ����'"
	elseif DateDiff("d",Rs("PostTime"),Now()) <= SiteConfig("PopularPostThresholdDays")   and  (  Rs("TotalReplies")=>SiteConfig("PopularPostThresholdPosts") or  Rs("TotalViews")=>SiteConfig("PopularPostThresholdViews") ) then
		IconImage="topic-hot.gif alt='��������'"
	else
		IconImage="topic.gif alt='��ͨ����'"
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
			ThreadStatus="<img src=images/status_Answered.gif align=middle title='����״̬���ѽ��'>"
		elseif Rs("ThreadStatus")=2 then
			ThreadStatus="<img src=images/status_NotAnswered.gif align=middle title='����״̬��δ���'>"
		else
			ThreadStatus="<img src=images/status_NotSet.gif align=middle>"
		end if
	end if
	if Rs("HiddenCount")>0 then
		ThreadHidden="&nbsp;<img src='images/InVisible.gif' align=middle alt='"&Rs("HiddenCount")&"δ�������' />"
	end if
	if Rs("DeletedCount")>0 then
		ThreadDel="&nbsp;<img src='images/recycle.gif' align=middle alt='"&Rs("DeletedCount")&"��ɾ������' />"
	end if
	
	if DateDiff("d",Rs("PostTime"),Now()) < 2 then
		NewHtml=" <img title='һ�����·��������' src=images/new.gif align=absmiddle>"
	else
		NewHtml=""
	end if
	if Rs("TotalRatings")>0 then StarHtml="<a style=CURSOR:pointer onclick="&CHR(34)&"OpenWindow('PostRating.asp?ThreadID="&Rs("ThreadID")&"')"&CHR(34)&" ><img border=0 src=Images/Star/"&cint(Rs("RatingSum")/Rs("TotalRatings"))&".gif align=middle></a>&nbsp;"
	if Rs("TotalReplies")=>SiteConfig("PostsPerPage") then
		MaxPostPage=fix(Rs("TotalReplies")/SiteConfig("PostsPerPage"))+1 '������ҳ
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
							Response.write("�ȴ���ˣ�")
						elseif Rs("Visible")=2 then
							Response.write("�ѱ�ɾ����")
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