<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")

if Request_Method <> "POST" then error("<li>�ύ��ʽ����</li><li>������ʹ�õ���"&Request_Method&"�ύ��ʽ��</li>")

HtmlTop

ForumID=RequestInt("ForumID")
ThreadID=Request("ThreadID")
PostID=Request.Form("PostID")


If IsNumeric(ThreadID) then
	ForumID=Execute("Select ForumID From ["&TablePrefix&"Threads] where ThreadID="&ThreadID&"")(0)
else
	for each ho in Request("ThreadID")
		if Not IsNumeric(ho) then error(ThreadID&"�Ƿ�����")
	next
End If

If Not IsNumeric(PostID) then
	for each ho in Request("PostID")
		if Not IsNumeric(ho) then error(PostID&"�Ƿ�����")
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
			succtitle="�����������⣬����ID��"&ThreadID&""
		else
			error("����Ȩ�޲���")
		end if

	case "UnTop"
		if BestRole = 1 then
			for each ho in Request("ThreadID")
				ho=int(ho)
				Execute("update ["&TablePrefix&"Threads] Set ThreadTop=0,StickyDate="&SqlNowString&" where ThreadID="&ho&ForumIDSql&"")
			next
			succtitle="����ȡ�����棬����ID��"&ThreadID&""
		else
			error("����Ȩ�޲���")
		end if
		
	case "Fix"
		for each ho in Request("ThreadID")
			ho=int(ho)
			UpdateThreadStatic(ho)
		next
		succtitle="�����޸����⣬����ID��"&ThreadID&""
		
	case "MoveNew"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set LastTime="&SqlNowString&" where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="������ǰ���⣬����ID��"&ThreadID&""
		
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
		succtitle="�����������⣬����ID��"&ThreadID&""
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
		succtitle="����ȡ������������ID��"&ThreadID&""
	case "ThreadTop"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set ThreadTop=1,StickyDate=DateAdd("&SqlChar&"yyyy"&SqlChar&", 1, "&SqlNowString&") where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="�����ö����⣬����ID��"&ThreadID&""
	case "DelTop"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set ThreadTop=0,StickyDate="&SqlNowString&" where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="����ȡ���ö�������ID��"&ThreadID&""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	case "IsLocked"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set IsLocked=1 where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="��������������ID��"&ThreadID&""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	case "DelIsLocked"
		for each ho in Request("ThreadID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Threads] Set IsLocked=0 where ThreadID="&ho&ForumIDSql&"")
		next
		succtitle="��������������ID��"&ThreadID&""
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
		succtitle="������ˣ�����ID��"&ThreadID&""

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
		succtitle="����ȡ����ˣ�����ID��"&ThreadID&""
		
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
		succtitle="����ɾ�����⣬����ID��"&ThreadID&""
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
		succtitle="����ȡ��ɾ��������ID��"&ThreadID&""

	case "MoveThread"
		Response.Redirect("MoveThread.asp?ThreadID="&Request("ThreadID"))
	case "MoveThreadUp"
		AimForumID=RequestInt("AimForumID")
		if AimForumID=0 then error("��û��ѡ��Ҫ�������ƶ��ĸ���̳")
		ThreadIDArray=split(ThreadID,",")
		ReturnThreadID=0
		for i=0 to ubound(ThreadIDArray)
			ho=int(trim(ThreadIDArray(i)))
			if ReturnThreadID=0 then ReturnThreadID=ho
			Execute("update ["&TablePrefix&"Threads] Set ForumID="&AimForumID&",ThreadTop=0,IsGood=0,IsLocked=0,ThreadStyle='' where ThreadID="&ho&ForumIDSql&"")
		next
		UpForumMostRecent(ForumID)
		succtitle="�����ƶ����⣬����ID��"&ThreadID&""
		succReturnUrl="ShowPost.asp?ThreadID="&ReturnThreadID&""
		

'''''''''''''''''''''''''''''''''''���ӹ��� Start''''''''''''''''''''''''''''''''''
	case "PostVisible"
		for each ho in Request.Form("PostID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Posts] Set Visible=1 where ThreadID="&ThreadID&" and PostID="&ho&"")
		next
		UpdateThreadStatic(ThreadID)
		succtitle="����������ӣ�����ID��"&PostID&""
		
	case "PostInVisible"
		for each ho in Request.Form("PostID")
			ho=int(ho)
			Execute("update ["&TablePrefix&"Posts] Set Visible=0 where ThreadID="&ThreadID&" and PostID="&ho&"")
		next
		UpdateThreadStatic(ThreadID)
		succtitle="����ȡ����ˣ�����ID��"&PostID&""
		
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
		succtitle="����ɾ�����ӣ�����ID��"&PostID&""
		
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
		succtitle="������ԭ���ӣ�����ID��"&PostID&""
		
'''''''''''''''''''''''''''''''''''���ӹ��� End''''''''''''''''''''''''''''''''''
end select
if succtitle="" then error("��Чָ��")

Log(""&succtitle&"")
Message="<li>"&succtitle&"</li>"
succeed Message,""&succReturnUrl&""

%>