<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then
	if instr("|FavoriteFriend|FavoritePost|FavoriteForums|","|"&Request("menu")&"|")>0 then
		response.Write("���ȵ�¼��̳")
	else
		error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")
	end if
end if

if Request("menu")<>"" then
	if Request_Method <> "POST" then ResponseWrite("<li>�ύ��ʽ����</li><li>������ʹ�õ���"&Request_Method&"�ύ��ʽ��</li>")
end if

if Http_Referer="" then Http_Referer="Default.asp"
FavoriteID=RequestInt("FavoriteID")
PostID=RequestInt("PostID")
ForumID=RequestInt("ForumID")
FriendUserName=HTMLEncode(Request("FriendUserName"))


select case Request("menu")
	case "Favorites"
		PageSetup=20 '�趨ÿҳ����ʾ����
		PageCount = RequestInt("PageIndex") '��ȡ��ǰҳ
		if PageCount <1 then PageCount = 1

		FavoriteAreaStr="<table cellspacing=0 cellpadding=0 width='100%' class='PannelBody' style='table-layout: fixed'>"
		select case Request("Item")
		case "FavoriteUsers"
			TotalCount=Execute("Select count(FavoriteID) from ["&TablePrefix&"FavoriteUsers] where OwnerUserName='"&CookieUserName&"'")(0)
			TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
			if PageCount > TotalPage then PageCount = TotalPage
			colspan=7
			FavoriteUsers("Select FriendUserName from ["&TablePrefix&"FavoriteUsers] where OwnerUserName='"&CookieUserName&"' order by FavoriteID desc")

		case "WatchingYou"
			TotalCount=Execute("Select count(FavoriteID) from ["&TablePrefix&"FavoriteUsers] where FriendUserName='"&CookieUserName&"'")(0)
			TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
			if PageCount > TotalPage then PageCount = TotalPage
			colspan=7
			FavoriteUsers("Select OwnerUserName from ["&TablePrefix&"FavoriteUsers] where FriendUserName='"&CookieUserName&"' order by FavoriteID desc")

		case "FavoritePosts"
			TotalCount=Execute("Select count(FavoriteID) from ["&TablePrefix&"FavoritePosts] where OwnerUserName='"&CookieUserName&"'")(0)
			TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
			if PageCount > TotalPage then PageCount = TotalPage
			colspan=4
			FavoritePosts

		case "FavoriteForums"
			TotalCount=Execute("Select count(FavoriteID) from ["&TablePrefix&"FavoriteForums] where OwnerUserName='"&CookieUserName&"'")(0)
			TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
			if PageCount > TotalPage then PageCount = TotalPage
			colspan=3
			FavoriteForums
		end select
	
		FavoriteAreaStr=FavoriteAreaStr&"<tr><td colspan='"&colspan&"' align=right>"&AjaxShowPage(TotalPage,PageCount,"MyFavorites.asp?menu="&Request("menu")&"&Item="&Request("Item")&"")&"</td></tr>"
		FavoriteAreaStr=FavoriteAreaStr&"</table>"
		response.Write(FavoriteAreaStr)


	case "FavoriteFriend"
		if Ucase(FriendUserName)=Ucase(CookieUserName) then ResponseWrite("��������Լ���")
		if Execute("Select UserName from ["&TablePrefix&"Users] where UserName='"&FriendUserName&"'").eof then ResponseWrite("ϵͳ��û�� "&FriendUserName&" �����ϣ�")

		Rs.Open "Select * from ["&TablePrefix&"FavoriteUsers] where OwnerUserName='"&CookieUserName&"' and FriendUserName='"&FriendUserName&"'",Conn,1,3
		if Rs.eof then Rs.addNew
			Rs("OwnerUserName")=CookieUserName
			Rs("FriendUserName")=FriendUserName
		Rs.update
		Rs.close

	case "FavoritePost"
		if Execute("Select PostID from ["&TablePrefix&"Posts] where PostID="&PostID&"").eof then ResponseWrite("ϵͳ��û��IDΪ "&PostID&" �����ӣ�")

		Rs.Open "Select * from ["&TablePrefix&"FavoritePosts] where OwnerUserName='"&CookieUserName&"' and PostID="&PostID&"",Conn,1,3
		if Rs.eof then Rs.addNew
			Rs("OwnerUserName")=CookieUserName
			Rs("PostID")=PostID
		Rs.update
		Rs.close
		
	case "FavoriteForums"
		if Execute("Select ForumID from ["&TablePrefix&"Forums] where ForumID="&ForumID&"").eof then ResponseWrite("ϵͳ��û��IDΪ "&ForumID&" ����̳��")

		Rs.Open "Select * from ["&TablePrefix&"FavoriteForums] where OwnerUserName='"&CookieUserName&"' and ForumID="&ForumID&"",Conn,1,3
		if Rs.eof then Rs.addNew
			Rs("OwnerUserName")=CookieUserName
			Rs("ForumID")=ForumID
		Rs.update
		Rs.close


	case "DelFavoriteForums"
		Execute("delete from ["&TablePrefix&"FavoriteForums] where OwnerUserName='"&CookieUserName&"' and ForumID="&ForumID&"")
		succeed "ɾ���ɹ�",""
				
	case "DelFavoritePosts"
		Execute("delete from ["&TablePrefix&"FavoritePosts] where OwnerUserName='"&CookieUserName&"' and PostID="&PostID&"")
		succeed "ɾ���ɹ�",""
		
	case "DelFavoriteFriend"
		Execute("delete from ["&TablePrefix&"FavoriteUsers] where OwnerUserName='"&CookieUserName&"' and FriendUserName='"&FriendUserName&"'")
		succeed "ɾ���ɹ�",""
	case else
		default
end select
	
	
	
Sub Default
	HtmlTop
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� �ղؼ�</div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle align="center">
		<td><a href="EditProfile.asp">�����޸�</a></td>
		<td><a href="EditProfile.asp?menu=pass">�����޸�</a></td>
		<td><a href="MyUpFiles.asp">�ϴ�����</a></td>
		<td><a href="MyFavorites.asp">�� �� ��</a></td>
		<td><a href="MyMessage.asp">���ŷ���</a></td>
	</tr>
</table>
<br />

<div class="MenuTag" ajaxurl="MyFavorites.asp?menu=Favorites"><li id="CommentArea_btn_0" onclick='AjaxShowPannel(this)' menu="item=FavoriteUsers" class="NowTag">�����б�</li><li id="CommentArea_btn_1" onclick='AjaxShowPannel(this)' menu="item=WatchingYou">��ע�����û�</li><li id="CommentArea_btn_2" onclick='AjaxShowPannel(this)' menu="item=FavoriteForums">��̳�ղؼ�</li><li id="CommentArea_btn_3" onclick='AjaxShowPannel(this)' menu="item=FavoritePosts">�����ղؼ�</li></div>

<div id=CommentArea>
	<table cellspacing=0 cellpadding=0 width="100%" class="PannelBody"><tr><td><img src="images/loading.gif" border=0 /><script language="javascript" type="text/javascript">Ajax_CallBack(false,'CommentArea','MyFavorites.asp?menu=Favorites&Item=FavoriteUsers');</script></td></tr></table>
</div>
<%
	HtmlBottom
End Sub

Sub ResponseWrite(Message)
	Response.Write(Message)
	Response.End()
End Sub







Sub FavoriteUsers(Sql)
	FavoriteAreaStr=FavoriteAreaStr&"<tr align='center'><td>�û���</td><td width='10%'>������</td><td width='10%'>������</td><td width='10%'>����</td><td width='10%'>��Ǯ</td><td width='25%'>����</td><td width='10%'>����</td></tr>"	
	
	Set Rs=Execute(sql)
	If TotalPage>1 then Rs.Move (PageCount-1) * PageSetup
	IF Not Rs.Eof then UserNameArray=Rs.GetRows(PageSetup)
	Rs.close
		
	if IsArray(UserNameArray) then
		For i=0 To Ubound(UserNameArray,2)
			Set Rs=Execute("["&TablePrefix&"Users] where UserName='"&UserNameArray(0,i)&"'")
			if Rs.eof then Execute("delete from ["&TablePrefix&"FavoriteUsers] where FriendUserName="&UserNameArray(0,i)&"")
			FavoriteAreaStr=FavoriteAreaStr&"<tr align=center>"
			FavoriteAreaStr=FavoriteAreaStr&"<td valign=center><a target='_blank' href='Profile.asp?UID="&Rs("UserID")&"'>"&Rs("UserName")&"</a>"
			FavoriteAreaStr=FavoriteAreaStr&"<td><a onclick=""javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&RecipientUserName="&Rs("UserName")&"', 600, 350);""><img border='0' src='Images/message.gif' /></a></td>"
			FavoriteAreaStr=FavoriteAreaStr&"<td>"&Rs("TotalPosts")&"</td>"
			FavoriteAreaStr=FavoriteAreaStr&"<td>"&Rs("Experience")&"</td>"
			FavoriteAreaStr=FavoriteAreaStr&"<td>"&Rs("UserMoney")&"</td>"
			FavoriteAreaStr=FavoriteAreaStr&"<td>"&Rs("UserActivityTime")&"</td>"
			FavoriteAreaStr=FavoriteAreaStr&"<td><a onclick=""return window.confirm('ȷʵҪ�� "&Rs("UserName")&" �Ӻ����б���ɾ��?')"" href=""javascript:UrlPost('?menu=DelFavoriteFriend&FriendUserName="&Rs("UserName")&"')""><img src=images/delete.gif border=0 /></a></td>"
			FavoriteAreaStr=FavoriteAreaStr&"</tr>"
			Set Rs = Nothing
		Next
	end if
end Sub


Sub FavoriteForums
	FavoriteAreaStr=FavoriteAreaStr&"<tr align=center><td width='30%'>�������</td><td>����</td><td width='10%'>����</td></tr>"

	Set Rs=Execute("Select ForumID from ["&TablePrefix&"FavoriteForums] where OwnerUserName='"&CookieUserName&"' order by FavoriteID desc")
	If TotalPage>1 then Rs.Move (PageCount-1) * PageSetup
	IF Not Rs.Eof then SQL=Rs.GetRows(PageSetup)
	Rs.close

	if IsArray(SQL) then
		For i=0 To Ubound(SQL,2)
			Set Rs=Execute("Select * from ["&TablePrefix&"Forums] where ForumID="&SQL(0,i)&"")
				if Rs.eof then Execute("delete from ["&TablePrefix&"FavoriteForums] where ForumID="&SQL(0,i)&"")
				
				FavoriteAreaStr=FavoriteAreaStr&"<tr>"
				FavoriteAreaStr=FavoriteAreaStr&"<td><a href='ShowForum.asp?ForumID="&Rs("ForumID")&"' target=_blank>"&Rs("ForumName")&"</a></td>"
				FavoriteAreaStr=FavoriteAreaStr&"<td>"&BBCode(Rs("ForumDescription"))&"</td>"
				FavoriteAreaStr=FavoriteAreaStr&"<td align=center><a onclick=""return window.confirm('��ȷ��Ҫɾ��������?')"" href=""javascript:UrlPost('?menu=DelFavoriteForums&ForumID="&Rs("ForumID")&"')""><img src=images/delete.gif border=0 /></a></td>"
				FavoriteAreaStr=FavoriteAreaStr&"</tr>"
			Set Rs = Nothing
		Next
	end if
End Sub



Sub FavoritePosts
	FavoriteAreaStr=FavoriteAreaStr&"<tr align=center><td width='40%'>����</td><td width='10%'>����</td><td width='10%'>�ظ�ʱ��</td><td width='10%'>����</td></tr>"

	Set Rs=Execute("Select PostID from ["&TablePrefix&"FavoritePosts] where OwnerUserName='"&CookieUserName&"' order by FavoriteID desc")
	If TotalPage>1 then Rs.Move (PageCount-1) * PageSetup
	IF Not Rs.Eof then SQL=Rs.GetRows(PageSetup)
	Rs.close

	if IsArray(SQL) then
		For i=0 To Ubound(SQL,2)
			Set Rs=Execute("Select * from ["&TablePrefix&"Posts] where PostID="&SQL(0,i)&"")
			if Rs.eof then Execute("delete from ["&TablePrefix&"FavoritePosts] where PostID="&SQL(0,i)&"")
			Subject=Rs("Subject")
			if ""&Subject&""="" then Subject=Left(ReplaceText(""&Rs("Body")&"","<[^>]*>",""),10)
			FavoriteAreaStr=FavoriteAreaStr&"<tr>"
			FavoriteAreaStr=FavoriteAreaStr&"<td><a href='ShowPost.asp?PostID="&Rs("PostID")&"' target=_blank>"&Subject&"</a></td>"
			FavoriteAreaStr=FavoriteAreaStr&"<td align=center><a href='Profile.asp?UserName="&Rs("PostAuthor")&"' target='_blank'>"&Rs("PostAuthor")&"</a></td>"
			FavoriteAreaStr=FavoriteAreaStr&"<td align=center>"&FormatDateTime(Rs("PostDate"),2)&"</td>"
			FavoriteAreaStr=FavoriteAreaStr&"<td align=center><a onclick=""return window.confirm('��ȷ��Ҫɾ��������?')"" href=""javascript:UrlPost('?menu=DelFavoritePosts&PostID="&Rs("PostID")&"')""><img src=images/delete.gif border=0 /></a></td>"
			FavoriteAreaStr=FavoriteAreaStr&"</tr>"
			FavoriteAreaStr=FavoriteAreaStr&""
			Set Rs = Nothing
		Next
	end if
End Sub
%>