<!-- #include file="Setup.asp" --><%
HtmlTop
ForumID=RequestInt("ForumID")
sql="Select * from ["&TablePrefix&"Forums] where ForumID="&ForumID&""
Set Rs=Execute(sql)
	if Rs.eof then error"<li>系统找不到该论坛的资料"
	ForumName=Rs("ForumName")
	Moderated=Rs("Moderated")
	ParentID=Rs("ParentID")
	GroupID=Rs("GroupID")
Rs.close
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → <%=ForumTree(ParentID)%><a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a> → 权限列表</div>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td align="center" width="9%">角色</td>
		<td align="center" width="9%">浏览</td>
		<td align="center" width="9%">阅读</td>
		<td align="center" width="9%">发帖</td>
		<td align="center" width="9%">回复</td>
		<td align="center" width="9%">编辑</td>
		<td align="center" width="9%">删除</td>
		<td align="center" width="9%">发起投票</td>
		<td align="center" width="9%">投票</td>
		<td align="center" width="9%">附件</td>
		<td align="center" width="9%">管理</td>
	</tr>
<%
sql="Select * from ["&TablePrefix&"ForumPermissions] where ForumID="&ForumID&" order by RoleID"
Set Rs=Execute(sql)
	Do While Not Rs.EOF 
%>
	<tr class="CommonListCell">
		<td align="center"><%=Execute("Select Name From ["&TablePrefix&"Roles] where RoleID="&RS("RoleID")&"")(0)%></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionView")%>.gif" /></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionRead")%>.gif" /></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionPost")%>.gif" /></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionReply")%>.gif" /></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionEdit")%>.gif" /></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionDelete")%>.gif" /></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionCreatePoll")%>.gif" /></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionVote")%>.gif" /></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionAttachment")%>.gif" /></td>
		<td align="center"><img src="Images/ForumPermissions<%=RS("PermissionManage")%>.gif" /></td>
	</tr>
<%
		Rs.MoveNext
	loop
Rs.Close
%>
</table>
<center><br /><a href="javascript:history.back()">&lt; 返 回 &gt;</a>
<%
HtmlBottom
%>