<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")

if Request_Method <> "POST" then error("<li>提交方式错误！</li><li>您本次使用的是"&Request_Method&"提交方式！</li>")

HtmlTop

ForumID=RequestInt("ForumID")
ThreadID=Request("ThreadID")
PostID=Request.Form("PostID")


If IsNumeric(ThreadID) then
	ForumID=Execute("Select ForumID From ["&TablePrefix&"Threads] where ThreadID="&ThreadID&"")(0)
else
	for each ho in Request("ThreadID")
		if Not IsNumeric(ho) then error(ThreadID&"非法操作")
	next
End If

If Not IsNumeric(PostID) then
	for each ho in Request("PostID")
		if Not IsNumeric(ho) then error(PostID&"非法操作")
	next
End If


if BestRole<>1 then
%><!-- #include file="Utility/ForumPermissions.asp" --><%
end if

if BestRole=1 and ForumID<1 then
	ForumIDSql=""
else
	ForumIDSql=" and ForumID="&ForumID&""
end if

select case Request("menu")
	case "Top"
		if BestRole = 1 then
			for each ho in Request("ThreadID")
				ho=int(ho)
				Execute("update ["&TablePrefix&"Threads] Set ThreadTop=2,StickyDate=DateAdd("&SqlChar&"yyyy"&SqlChar&", 3, "&SqlNowString&") where ThreadID="&ho&ForumIDSql&"")
			next
			succtitle="批量公告主题，主题ID："&ThreadID&""
		else
			error("您的权限不够")
		end if

	case "UnTop"
		if BestRole = 1 then
			for each ho in Request("ThreadID")
				ho=int(ho)
				Execute("update ["&TablePrefix&"Threads] Set ThreadTop=0,StickyDate="&SqlNowString&" where ThreadID="&ho&ForumIDSql&"")
			next
			succtitle="批量取消公告，主题ID："&ThreadID&""
		else
			error("您的权限不够")
		end if
		
	case "Fix"
		for each ho in Request("ThreadID")
			ho=int(ho)
			UpdateThreadStatic(ho)
		next
		succtitle="批量修复主题，主题ID："&ThreadID&""
		
	case "MoveNew"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set LastTime="&SqlNowString&" where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="批量拉前主题，主题ID："&ThreadID&""
		
	case "IsGood"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Rs.open "Select ThreadID,IsGood,PostAuthor From ["&TablePrefix&"Threads] where ThreadID="&ho&ForumIDSql&"",conn,1,3
			if not Rs.eof and Rs("IsGood")=0 then
				Rs("IsGood")=1
				Rs.update
				Execute("update ["&TablePrefix&"Users] Set UserMoney=UserMoney+"&SiteConfig("IntegralAddValuedPost")&",experience=experience+"&SiteConfig("IntegralAddValuedPost")&" where UserName='"&Rs("PostAuthor")&"'")
			end if
			Rs.close
		next
		succtitle="批量精华主题，主题ID："&ThreadID&""
	case "DelIsGood"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Rs.open "Select ThreadID,IsGood,PostAuthor From ["&TablePrefix&"Threads] where ThreadID="&ho&ForumIDSql&"",conn,1,3
			if not Rs.eof and Rs("IsGood")=1 then
				Rs("IsGood")=0
				Rs.update
				Execute("update ["&TablePrefix&"Users] Set UserMoney=UserMoney+"&SiteConfig("IntegralDeleteValuedPost")&",experience=experience+"&SiteConfig("IntegralDeleteValuedPost")&" where UserName='"&Rs("PostAuthor")&"'")
			end if
			Rs.close
		next
		succtitle="批量取消精华，主题ID："&ThreadID&""
	case "ThreadTop"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set ThreadTop=1,StickyDate=DateAdd("&SqlChar&"yyyy"&SqlChar&", 1, "&SqlNowString&") where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="批量置顶主题，主题ID："&ThreadID&""
	case "DelTop"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set ThreadTop=0,StickyDate="&SqlNowString&" where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="批量取消置顶，主题ID："&ThreadID&""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	case "IsLocked"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set IsLocked=1 where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="批量锁定，主题ID："&ThreadID&""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	case "DelIsLocked"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set IsLocked=0 where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="批量解锁，主题ID："&ThreadID&""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	case "Visible"
		for each ho in Request("ThreadID")
			ho=int(ho)
			sql="select * from ["&TablePrefix&"Threads] where ThreadID="&ho&ForumIDSql&""
			Rs.open sql,Conn,1,3
			if not Rs.eof and Rs("Visible")<>1 then
				if Rs("Visible")=0 then
					Rs("HiddenCount")=Rs("HiddenCount")-1
				elseif Rs("Visible")=2 then
					Rs("DeletedCount")=Rs("DeletedCount")-1
				end if
				Rs("Visible")=1
				Rs("LastTime")=now()
				Rs("LastName")=CookieUserName
				Rs.update
				Execute("update ["&TablePrefix&"Posts] Set Visible=1 where ThreadID="&ho&" and ParentID=0")
			end if
			Rs.close
		next
		succtitle="批量审核，主题ID："&ThreadID&""

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	case "InVisible"
		for each ho in Request("ThreadID")
			ho=int(ho)
			sql="select * from ["&TablePrefix&"Threads] where ThreadID="&ho&ForumIDSql&""
			Rs.open sql,Conn,1,3
			if not Rs.eof and Rs("Visible")<>0 then
				if Rs("Visible")=2 then Rs("DeletedCount")=Rs("DeletedCount")-1
				Rs("Visible")=0
				Rs("HiddenCount")=Rs("HiddenCount")+1
				Rs("LastTime")=now()
				Rs("LastName")=CookieUserName
				Rs.update
				Execute("update ["&TablePrefix&"Posts] Set Visible=0 where ThreadID="&ho&" and ParentID=0")
			end if
			Rs.close
		next
		succtitle="批量取消审核，主题ID："&ThreadID&""
		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	case "DelThread"
		for each ho in Request("ThreadID")
			ho=int(ho)
			sql="select * from ["&TablePrefix&"Threads] where ThreadID="&ho&ForumIDSql&""
			Rs.open sql,Conn,1,3
			if not Rs.eof and Rs("Visible")<>2 then
				if Rs("Visible")=0 then Rs("HiddenCount")=Rs("HiddenCount")-1
				Rs("DeletedCount")=Rs("DeletedCount")+1
				Rs("Visible")=2
				Rs("LastTime")=now()
				Rs("LastName")=CookieUserName
				Rs.update
				Execute("update ["&TablePrefix&"Posts] Set Visible=2 where ThreadID="&ho&" and ParentID=0")
			end if
			Rs.close
		next
		UpForumMostRecent(ForumID)
		succtitle="批量删除主题，主题ID："&ThreadID&""
		if instr(Http_Referer,"Search.asp")>0 then succReturnUrl="Search.asp"
	case "UnDelThread"
		for each ho in Request("ThreadID")
			ho=int(ho)
			sql="select * from ["&TablePrefix&"Threads] where ThreadID="&ho&ForumIDSql&""
			Rs.open sql,Conn,1,3
			if not Rs.eof and Rs("Visible")=2 then
				Rs("DeletedCount")=Rs("DeletedCount")-1
				Rs("Visible")=1
				Rs("LastTime")=now()
				Rs("LastName")=CookieUserName
				Rs.update
				Execute("update ["&TablePrefix&"Posts] Set Visible=1 where ThreadID="&ho&" and ParentID=0")
			end if
			Rs.close
		next
		UpForumMostRecent(ForumID)
		succtitle="批量取消删除，主题ID："&ThreadID&""

	case "MoveThread"
		Response.Redirect("MoveThread.asp?ThreadID="&Request("ThreadID"))
	case "MoveThreadUp"
		AimForumID=RequestInt("AimForumID")
		if AimForumID=0 then error("您没有选择要将主题移动哪个论坛")
		ThreadIDArray=split(ThreadID,",")
		ReturnThreadID=0
		for i=0 to ubound(ThreadIDArray)
			ho=int(trim(ThreadIDArray(i)))
			if ReturnThreadID=0 then ReturnThreadID=ho
			Execute("update ["&TablePrefix&"Threads] Set ForumID="&AimForumID&",ThreadTop=0,IsGood=0,IsLocked=0,ThreadStyle='' where ThreadID="&ho&ForumIDSql&"")
		next
		UpForumMostRecent(ForumID)
		succtitle="批量移动主题，主题ID："&ThreadID&""
		succReturnUrl="ShowPost.asp?ThreadID="&ReturnThreadID&""
		

'''''''''''''''''''''''''''''''''''帖子管理 Start''''''''''''''''''''''''''''''''''
	case "PostVisible"
		for each ho in Request.Form("PostID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Posts] Set Visible=1 where ThreadID="&ThreadID&" and PostID="&ho&"")
		next
		UpdateThreadStatic(ThreadID)
		succtitle="批量审核帖子，帖子ID："&PostID&""
		
	case "PostInVisible"
		for each ho in Request.Form("PostID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Posts] Set Visible=0 where ThreadID="&ThreadID&" and PostID="&ho&"")
		next
		UpdateThreadStatic(ThreadID)
		succtitle="批量取消审核，帖子ID："&PostID&""
		
	case "DelPost"
		for each ho in Request.Form("PostID")
			ho=int(ho)
			Rs.open "select * from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&" and PostID="&ho&"",Conn,1,3
			if not Rs.eof then
				Rs("Visible")=2
				Rs.update
				if Rs("ParentID")=0 then Execute("update ["&TablePrefix&"Threads] Set Visible=2 where ThreadID="&Rs("ThreadID")&"")
			end if
			Rs.close
		next
		UpdateThreadStatic(ThreadID)
		succtitle="批量删除帖子，帖子ID："&PostID&""
		
	case "UnDelPost"
		for each ho in Request.Form("PostID")
			ho=int(ho)
			Rs.open "select * from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&" and PostID="&ho&"",Conn,1,3
			if not Rs.eof then
				Rs("Visible")=1
				Rs.update
				if Rs("ParentID")=0 then Execute("update ["&TablePrefix&"Threads] Set Visible=1 where ThreadID="&Rs("ThreadID")&"")
			end if
			Rs.close
		next
		UpdateThreadStatic(ThreadID)
		succtitle="批量还原帖子，帖子ID："&PostID&""
		
'''''''''''''''''''''''''''''''''''帖子管理 End''''''''''''''''''''''''''''''''''
end select
if succtitle="" then error("无效指令")

Log(""&succtitle&"")
Message="<li>"&succtitle&"</li>"
succeed Message,""&succReturnUrl&""

%>