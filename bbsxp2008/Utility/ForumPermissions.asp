<%
sql="Select * from ["&TablePrefix&"Forums] where ForumID="&ForumID&""
Set Rs=Execute(sql)
	if Rs.eof then error"<li>�Ҳ����ð�������"
	ForumName=Rs("ForumName")
	Moderated=Rs("Moderated")
	ParentID=Rs("ParentID")
	GroupID=Rs("GroupID")
	ForumUrl=Rs("ForumUrl")
	IsActive=Rs("IsActive")
	ForumRules=BBCode(Rs("ForumRules"))
	ModerateNewThread=Rs("ModerateNewThread")
	ModerateNewPost=Rs("ModerateNewPost")
	TotalCategorys=Rs("TotalCategorys")
	TotalThreads=Rs("TotalThreads")
Rs.Close

GroupModerated=Execute("Select Moderated from ["&TablePrefix&"Groups] where GroupID="&GroupID&"")(0)
AllModerated=Moderated &"|"& GroupModerated


if BestRole=1 or (CookieUserName<>empty and instr("|"&AllModerated&"|","|"&CookieUserName&"|")>0) then

	PermissionView=1
	PermissionRead=1
	PermissionPost=1
	PermissionReply=1
	PermissionEdit=1
	PermissionDelete=1
	PermissionCreatePoll=1
	PermissionVote=1
	PermissionAttachment=1
	PermissionManage=1

else

	SQL="Select * from ["&TablePrefix&"ForumPermissions] where ForumID="&ForumID&" and RoleID="&CookieUserRoleID&""
	Rs.Open sql,Conn,1,3
		if Rs.eof then

			Rs.addNew

			if CookieUserRoleID = 0 then
				Rs("PermissionView")=1
				Rs("PermissionRead")=1
				Rs("PermissionPost")=0
				Rs("PermissionReply")=0
				Rs("PermissionEdit")=0
				Rs("PermissionDelete")=0
				Rs("PermissionCreatePoll")=0
				Rs("PermissionVote")=0
				Rs("PermissionAttachment")=0
				Rs("PermissionManage")=0
			elseif CookieUserRoleID > 2 then
				Rs("PermissionView")=1
				Rs("PermissionRead")=1
				Rs("PermissionPost")=1
				Rs("PermissionReply")=1
				Rs("PermissionEdit")=1
				Rs("PermissionDelete")=0
				Rs("PermissionCreatePoll")=1
				Rs("PermissionVote")=1
				Rs("PermissionAttachment")=1
				Rs("PermissionManage")=0
			end if
			Rs("ForumID")=ForumID
			Rs("RoleID")=CookieUserRoleID
			Rs.update
			Rs.close
			succeed "ϵͳ�����Զ�����Ȩ��","ShowForum.asp?ForumID="&ForumID&""


		else
			PermissionView=Rs("PermissionView")
			PermissionRead=Rs("PermissionRead")
			PermissionPost=Rs("PermissionPost")
			PermissionReply=Rs("PermissionReply")
			PermissionEdit=Rs("PermissionEdit")
			PermissionDelete=Rs("PermissionDelete")
			PermissionCreatePoll=Rs("PermissionCreatePoll")
			PermissionVote=Rs("PermissionVote")
			PermissionAttachment=Rs("PermissionAttachment")
			PermissionManage=Rs("PermissionManage")
		end if
	Rs.close

end if


if ForumUrl<>"" then response.redirect ForumUrl
if IsActive=0 and PermissionManage=0 then error"<li>�ð���Ѿ��رգ�"

if instr(Script_Name,"showforum.asp") > 0 then
	if PermissionView=0 then error("��û��<a href=ShowForumPermissions.asp?ForumID="&ForumID&">Ȩ��</a>��������б�")
	
elseif instr(Script_Name,"showpost.asp") > 0 then
	if PermissionRead=0 then error"<li>��û��<a href=ShowForumPermissions.asp?ForumID="&ForumID&">Ȩ��</a>�Ķ�����"
	
elseif instr(Script_Name,"addtopic.asp") > 0 then
	if PermissionPost=0 then error("��û��<a href=ShowForumPermissions.asp?ForumID="&ForumID&">Ȩ��</a>��������")
	
elseif instr(Script_Name,"addpost.asp") > 0 then
	if PermissionReply=0 then error("��û��<a href=ShowForumPermissions.asp?ForumID="&ForumID&">Ȩ��</a>�ظ�����")

elseif instr(Script_Name,"editpost.asp") > 0 then
	if PermissionEdit=0 then error("��û��<a href=ShowForumPermissions.asp?ForumID="&ForumID&">Ȩ��</a>�༭����")

elseif instr(Script_Name,"delpost.asp") > 0 then
	if PermissionDelete=0 then error("��û��<a href=ShowForumPermissions.asp?ForumID="&ForumID&">Ȩ��</a>ɾ������")
	
elseif instr(Script_Name,"forummanage.asp") or instr(Script_Name,"manage.asp") > 0 or instr(Script_Name,"moderation.asp") > 0 or instr(Script_Name,"movethread.asp") > 0 then	
	if PermissionManage=0 then error("��û��<a href=ShowForumPermissions.asp?ForumID="&ForumID&">Ȩ��</a>�������")
end if
%>