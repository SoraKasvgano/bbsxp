<!-- #include file="Setup.asp" -->
<%
AdminTop
if RequestCookies("UserPassword")="" or RequestCookies("UserPassword")<>session("pass") then response.redirect "Admin_Default.asp"
Log("")
UserName=HTMLEncode(Request("UserName"))
RoleID=RequestInt("RoleID")
ForumID=RequestInt("ForumID")
GroupID=RequestInt("GroupID")
TimeLimit=RequestInt("TimeLimit")

ForumIDList=""
select case Request("menu")
	case ""
		Alert("请选择您要操作的项目")
	case "ApplyManage"
		ApplyManage
	case "bbsManage"
		bbsManage
	case "ViewForumPermissions"
		ViewForumPermissions
	case "DelPermissions"
		Execute("Delete from ["&TablePrefix&"ForumPermissions] where ForumID="&ForumID&" and RoleID="&RoleID&"")
		response.redirect "?menu=ForumAdd&ForumID="&ForumID&""
	case "RoleAdd"
		if Execute("Select RoleID From ["&TablePrefix&"ForumPermissions] where ForumID="&ForumID&" and RoleID="&RoleID&"" ).Eof then
			if RoleID=0 then
				Execute("insert into ["&TablePrefix&"ForumPermissions] (ForumID,PermissionPost,PermissionReply,PermissionEdit,PermissionCreatePoll,PermissionVote,PermissionAttachment) values ('"&ForumID&"','0','0','0','0','0','0')")
			else
				Execute("insert into ["&TablePrefix&"ForumPermissions] (ForumID,RoleID) values ('"&ForumID&"','"&RoleID&"')")
			end if
		end if
		response.redirect "?menu=ForumAdd&ForumID="&ForumID&""
	case "ForumPermissionsUP"

		Rs.Open "["&TablePrefix&"ForumPermissions] where ForumID="&ForumID&" and RoleID="&RoleID&"",Conn,1,3
			for each ho in Request.Form
				Rs(ho)=Request(ho)
			next
		Rs.update
		Rs.close
		response.redirect "?menu=ForumAdd&ForumID="&ForumID&""

	case "ManageGroups"
		ManageGroups


	case "AddGroup"
		AddGroup
	case "AddGroupUp"
		AddGroupUp
	case "GroupManageDel"
		Execute("Delete from ["&TablePrefix&"Groups] where GroupID="&GroupID&"")
		Alert("已经将该论坛组的所有数据删除了！")
	case "ForumAdd"
		ForumAdd
	case "ForumAddUp"
		ForumAddUp
	case "bbsManageDel"
		Execute("Delete from ["&TablePrefix&"Forums] where ForumID="&ForumID&"")				'只删除该论坛，下属论坛仍在
		Alert("已经将该论坛的所有数据删除了！")
	case "DelThreads"
		if ForumID>0 then ForumIDList="and ForumID="&ForumID&""
		Execute("Delete from ["&TablePrefix&"Threads] where DateDiff("&SqlChar&"d"&SqlChar&",LastViewedDate,"&SqlNowString&") > "&TimeLimit&" "&ForumIDList&"")
		Alert("已经将"&TimeLimit&"天没有人查看过的主题删除了！")
	case "DelUserTopicok"
		if UserName="" then Alert("您没有输入用户名！")
		if ForumID>0 then ForumIDList=" and ForumID="&ForumID&""
		Execute("Delete from ["&TablePrefix&"Threads] where PostAuthor='"&UserName&"'"&ForumIDList&"")
		Alert("已经将"&UserName&"发表的主题删除了！")
	case "DelUserPost"
		if UserName="" then Alert("您没有输入用户名！")
		Execute("Delete from ["&TablePrefix&"Posts] where PostAuthor='"&UserName&"' and ParentID>0")
		Alert("已经将 "&UserName&" 发表的回帖删除了！")
		
	case "ClearRecyclePost"
		TimeLimit=RequestInt("TimeLimit")
		if TimeLimit < 1 then error("只能清空24小时之前的回帖")
		Execute("Delete from ["&TablePrefix&"Posts] where Visible=2 and DateDiff("&SqlChar&"d"&SqlChar&",PostDate,"&SqlNowString&") > "&TimeLimit&"")
		Alert("已经将 "&TimeLimit&" 天以前删除的回帖清空了！")
		
	case "DellikeTopicok"
		Topic=HTMLEncode(Request("Topic"))
		if Topic="" then Alert("您没有输入字符！")
		if ForumID>0 then ForumIDList="and ForumID="&ForumID&""
		Execute("Delete from ["&TablePrefix&"Threads] where Topic like '%"&Topic&"%' "&ForumIDList&" ")
		Alert("已经将标题里包含有 "&Topic&" 的帖子全部删除了！")
		
	case "DellikeTopicPost"
		Topic=HTMLEncode(Request("Topic"))
		if Topic="" then Alert("您没有输入字符！")
		Execute("Delete from ["&TablePrefix&"Posts] where Subject like '%"&Topic&"%' and ParentID>0")
		Alert("已经将标题里包含有 "&Topic&" 的回帖删除了！")

	case "uniteok"

		ForumSource=RequestInt("ForumSource")
		ForumTarget=RequestInt("ForumTarget")
		if ForumSource=0 or ForumTarget=0 then Alert("该操作不能选择论坛组！")
		if ForumSource = ForumTarget then Alert("不能选择相同论坛！")
		if UserName<>"" then UserNamelist="and PostAuthor='"&UserName&"'"
		Execute("update ["&TablePrefix&"Threads] Set ForumID="&ForumTarget&" where ForumID="&ForumSource&" and DateDiff("&SqlChar&"d"&SqlChar&",lasttime,"&SqlNowString&") > "&TimeLimit&" "&UserNamelist&"")
		Alert("移动论坛资料成功！")

	case "TreeView"
		FroumTreeView



	case "Tags"
		ShowTags
	case "EditTagsUp"
		for each ho in Request.Form("TagID")
			Execute("update ["&TablePrefix&"PostTags] set TagName='"&HTMLEncode(Request.Form("TagName"&ho))&"',IsEnabled='"&Request.Form("IsEnabled"&ho)&"' where TagID="&ho&"")
		next
		response.write "标签更新成功<br><br><a href=javascript:history.back()>返 回</a>"
	case "DelTag"
		TagID=RequestInt("TagID")
		Execute("delete from ["&TablePrefix&"PostTags] where TagID="&TagID&"")
		response.write "标签删除成功<br><br><a href=javascript:history.back()>返 回</a>"
end select





Sub ApplyManage
%>
<table cellspacing="1" cellpadding="1" width="100%" border="0" align="center">
	<tr>
		<td align=right>
			<select name=Groups onchange="window.location='?menu=ApplyManage&GroupID='+this.options[this.selectedIndex].value">
			<option value="">所有的论坛组</option>
<%
	sql="Select * from ["&TablePrefix&"Groups] order by GroupID desc"
	Set Rs=Execute(sql)
		Do While Not Rs.EOF
			selec=""
			if Rs("GroupID")=GroupID then selec="selected"
			Response.Write("<option value='"&Rs("GroupID")&"'"&selec&">"&Rs("GroupName")&"</option>")
			Rs.MoveNext
		loop
	Rs.Close
%>
			</select>
		</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<tr class=CommonListTitle>
		<td align="center"><a href="?menu=ApplyManage&fashion=ForumID">ID</a></td>
		<td width="20%" align="center"><a href="?menu=ApplyManage&fashion=ForumName">论坛</a></td>
		<td align="center"><a href="?menu=ApplyManage&fashion=TodayPosts">今日</a></td>
		<td align="center"><a href="?menu=ApplyManage&fashion=TotalThreads">主题</a></td>
		<td align="center"><a href="?menu=ApplyManage&fashion=TotalPosts">帖子</a></td>
		<td align="center"><a href="?menu=ApplyManage&fashion=Moderated">版主</a></td>
		<td align="center"><a href="?menu=ApplyManage&fashion=IsActive">已启用</a></td>
		<td align="center" width="20%">操作</td>
	</tr>
<%
	fashion=HTMLEncode(Request("fashion"))
	if fashion="" then fashion="TotalPosts"
	
	if GroupID>0 then GroupSql=" where GroupID="&GroupID&""
	sql="Select * from ["&TablePrefix&"Forums] "&GroupSql&" order by "&fashion&" Desc"
	
	Rs.Open sql,Conn,1
	PageSetup=20 '设定每页的显示数量
	Rs.Pagesize=PageSetup
	TotalPage=Rs.Pagecount  '总页数
	PageCount = RequestInt("PageIndex")
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
	i=0
	Do While Not Rs.EOF and i<PageSetup
		i=i+1
%>
	<tr class="CommonListCell">
		<td align="center"><%=Rs("ForumID")%></td>
		<td><a href="?menu=ForumAdd&ForumID=<%=Rs("ForumID")%>&GroupID=<%=Rs("GroupID")%>"><%=Rs("ForumName")%></a></td>
		<td align="center"><b><font color="red"><%=Rs("TodayPosts")%></font></b></td>
		<td align="center"><b><font color="red"><%=Rs("TotalThreads")%></font></b></td>
		<td align="center"><b><font color="red"><%=Rs("TotalPosts")%></font></b></td>
		<td align="center" width="200"><%=Rs("Moderated")%></td>
		<td align="center"><%if Rs("IsActive")=1 then%><font color="#FF0000"><img src="images/ForumPermissions1.gif" /></font><%else%><img src="images/ForumPermissions0.gif" /><%end if%></td>
		<td align="center"><a href="?menu=ForumAdd&ForumID=<%=Rs("ForumID")%>&GroupID=<%=Rs("GroupID")%>" class="CommonTextButton">编辑</a> <a target="_blank" href="ShowForum.asp?ForumID=<%=Rs("ForumID")%>" class="CommonTextButton">浏览</a> <a onclick="return window.confirm('您确定要删除该论坛的所有资料?')" href="?menu=bbsManageDel&ForumID=<%=Rs("ForumID")%>" class="CommonTextButton">删除</a> </td>
<%
		Rs.MoveNext
	loop
	Rs.Close
%>
	</tr>
</table>
<table border="0" width="100%" align="center">
	<tr>
		<td><%ShowPage()%></td>
		<td align=right><a href="?menu=ForumAdd&GroupID=<%=GroupID%>" class="CommonTextButton">新建论坛</a> <a href="?menu=TreeView" class="CommonTextButton">切换到树型模式</a></td>
	</tr>
</table>
<%
End Sub

Sub AddGroup
	SortOrder=1
	if GroupID>0 then
		sql="Select * from ["&TablePrefix&"Groups] where GroupID="&GroupID&""
		Set Rs=Execute(sql)
		if Rs.EOF then Alert("系统不存在该组论坛组的资料")
			GroupName=Rs("GroupName")
			GroupDescription=Rs("GroupDescription")
			SortOrder=Rs("SortOrder")
			ForumColumns=Rs("ForumColumns")
			Moderated=Rs("Moderated")
		Rs.close
	end if
%>
<table cellspacing="1" cellpadding="5" width="90%" border="0" class=CommonListArea align="center">
<form name="form" method="POST" action="?menu=AddGroupUp&GroupID=<%=GroupID%>">
	<tr class=CommonListTitle>
		<td align=center colspan="2">新建/编辑论坛组</td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top"><b>名称</b><br />论坛组名称</td>
		<td><input name="GroupName" value="<%=GroupName%>" /></td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top"><b>排序</b><br />从小到大排序设置，为“0”则隐藏此论坛组</td>
		<td><input name="SortOrder" value="<%=SortOrder%>" /></td>
	</tr>
	<tr class="CommonListCell">
		<td><b>论坛横排</b><br />设置论坛组下的所有论坛横排时每行版块数量，如果设置为 0，则按正常方式排列</td>
		<td><input type="text" name="ForumColumns" value="<%=ForumColumns%>" /></td>
	</tr>
	<tr class="CommonListCell">
		<td><b>论坛组版主</b><br />用于管理论坛组下的所有论坛。多版主添加请用“|”分隔，如：yuzi|裕裕</td>
		<td><input type="text" name="Moderated" value="<%=Moderated%>" /></td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top"><b>论坛组描述</b><br />简单描述一下此论坛组</td>
		<td><textarea cols=54 rows=3 name=GroupDescription><%=GroupDescription%></textarea></td>
	</tr>
	<tr class="CommonListCell">
		<td colspan="2" align="center"><input type="submit" value=" 保存 " /></td>
	</tr>
</form>
</table>
<%
End Sub


Sub ManageGroups
%>
<table cellspacing="1" cellpadding="1" width="90%" border="0" align="center"><tr><td align=right><a href="?menu=AddGroup" class="CommonTextButton">新建论坛组</a></td></tr></table>
<table cellspacing="1" cellpadding="5" width="90%" border="0" class=CommonListArea align="center">
<tr class=CommonListTitle>
	<td width="20%">名称</td>
	<td align=center>描述</td>
	<td align="center" width="25%">动作</td>
</tr>
<%
sql="Select * from ["&TablePrefix&"Groups] order by SortOrder"
Rs.open sql,conn,1
do while not rs.eof
%>
<tr class="CommonListCell">
	<td><a href="?menu=AddGroup&GroupID=<%=rs("GroupID")%>"><%=Rs("GroupName")%></a></td>
	<td><%=Rs("GroupDescription")%></td>
	<td align=center><a href="?menu=AddGroup&GroupID=<%=rs("GroupID")%>" class="CommonTextButton">编辑</a> 
	<a href="?menu=ApplyManage&GroupID=<%=Rs("GroupID")%>" class="CommonTextButton">管理</a> 
	<a onclick="return window.confirm('您确定要删除该论坛组内的所有论坛?')" href="?menu=GroupManageDel&GroupID=<%=rs("GroupID")%>" class="CommonTextButton">删除</a></td>

</tr>
<%
	Rs.movenext
loop
%>
</table>
<table border="0" width="90%" align="center">
	<tr>
		<td align=right><a href="?menu=TreeView" class="CommonTextButton">切换到树型模式</a></td>
	</tr>
</table>
<%
End Sub



Sub AddGroupUp
	GroupName=HTMLEncode(Request("GroupName"))
	SortOrder=RequestInt("SortOrder")
	GroupDescription=BodyEncode(Request.Form("GroupDescription"))
	Moderated=HTMLEncode(Request("Moderated"))
	if GroupName="" then Alert("请输入论坛组名称")

	master=split(Moderated,"|")
	for i = 0 to ubound(master)
		If Execute("Select UserID From ["&TablePrefix&"Users] where UserName='"&master(i)&"'" ).eof and master(i)<>"" Then Alert(""&master(i)&"的用户资料不存在")
	next
	
	sql="Select * from ["&TablePrefix&"Groups] where GroupID="&GroupID&""
	Rs.open sql,conn,1,3
		if Rs.eof then Rs.addnew
		Rs("GroupName")=GroupName
		Rs("SortOrder")=SortOrder
		Rs("GroupDescription")=GroupDescription
		Rs("ForumColumns")=RequestInt("ForumColumns")
		Rs("Moderated")=Moderated
	Rs.update
	Rs.close
	Response.Write("添加/编辑 论坛组成功")
End Sub

Sub ForumAddUp
	
	ParentID=RequestInt("ParentID")
	SortOrder=RequestInt("SortOrder")
	ModerateNewPost=RequestInt("ModerateNewPost")
	ModerateNewThread=RequestInt("ModerateNewThread")
	IsActive=RequestInt("IsActive")
	GroupID=RequestInt("GroupID")

	ForumName=HTMLEncode(Request.Form("ForumName"))
	Moderated=HTMLEncode(Request.Form("Moderated"))
	TotalCategorys=HTMLEncode(Request.Form("TotalCategorys"))
	ForumDescription=HTMLEncode(Request.Form("ForumDescription"))
	ForumRules=HTMLEncode(Request.Form("ForumRules"))
	ForumUrl=HTMLEncode(Request.Form("ForumUrl"))

	if GroupID=0 then Alert("请先建立论坛组才能添加论坛！")
	if ForumName="" then Alert("请输入论坛名称！")

	master=split(Moderated,"|")
	for i = 0 to ubound(master)
		If Execute("Select UserID From ["&TablePrefix&"Users] where UserName='"&master(i)&"'" ).eof and master(i)<>"" Then Alert(""&master(i)&"的用户资料不存在")
	next
	sql="Select * from ["&TablePrefix&"Forums] where ForumID="&ForumID&""
	Response.Write("")
	Rs.Open sql,Conn,1,3
		if Rs.eof then Rs.addNew

		if Rs("ForumID")=ParentID  then ParentID=0

		Rs("ParentID")=ParentID
		Rs("SortOrder")=SortOrder
		Rs("ModerateNewPost")=ModerateNewPost
		Rs("ModerateNewThread")=ModerateNewThread
		Rs("IsActive")=IsActive
		Rs("GroupID")=GroupID
		Rs("ForumName")=ForumName
		Rs("Moderated")=Moderated
		Rs("TotalCategorys")=TotalCategorys
		Rs("ForumDescription")=ForumDescription
		Rs("ForumRules")=ForumRules
		Rs("ForumUrl")=ForumUrl
		Rs("DateCreated")=now()
	Rs.update
	ForumID=Rs("ForumID")
	Rs.close
	if Execute("Select RoleID From ["&TablePrefix&"ForumPermissions] where ForumID="&ForumID&" and RoleID=0").Eof then
		Execute("insert into ["&TablePrefix&"ForumPermissions] (ForumID,PermissionPost,PermissionReply,PermissionEdit,PermissionCreatePoll,PermissionVote,PermissionAttachment) values ('"&ForumID&"','0','0','0','0','0','0')")
	end if

	UpdateForumParentID GroupID,ForumID
	Response.Write("添加/编辑成功")
End Sub

Sub UpdateForumParentID(GroupID,ParentID)
	sql="Select ForumID,GroupID,ParentID from ["&TablePrefix&"Forums] where ParentID="&ForumID&""
	rs.open sql,conn,1,3
	if rs.eof then
		exit sub
	else
	do while not rs.eof
		ParentID=Rs("ForumID")
		rs("GroupID")=GroupID
		rs.update
		rs.movenext
	loop
	end if
	rs.close
	UpdateForumParentID GroupID,ParentID
End Sub

Sub ForumAdd
	IsActive=1
	SortOrder=1
	ParentID=RequestInt("ParentID")
	if ForumID > 0 then
		sql="Select * from ["&TablePrefix&"Forums] where ForumID="&ForumID&""
		Set Rs=Execute(sql)
		if Rs.EOF then Alert("系统不存在该论坛的资料")
			ForumDescription=replace(""&Rs("ForumDescription")&"","<br>",vbCrlf)
			ForumRules=replace(""&Rs("ForumRules")&"","<br>",vbCrlf)
			ForumName=Rs("ForumName")
			SortOrder=Rs("SortOrder")
			ParentID=Rs("ParentID")
			Moderated=Rs("Moderated")
			TotalCategorys=Rs("TotalCategorys")
			ModerateNewPost=Rs("ModerateNewPost")
			ModerateNewThread=Rs("ModerateNewThread")
			IsActive=Rs("IsActive")
			ForumUrl=Rs("ForumUrl")
		Rs.close
	end if
	if GroupID=0 then
		Set Rs=Execute("Select top 1 * from ["&TablePrefix&"Groups] order by GroupID")
		if Rs.EOF then Alert("请先创建论坛组")
		GroupID=Rs("GroupID")
		Rs.close
	end if
%>
<form method="POST" action="?menu=ForumAddUp&ForumID=<%=ForumID%>" name="form">
	<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
		<tr class=CommonListTitle>
			<td align=center colspan="2">添加/编辑论坛资料</td>
		</tr>
		<tr class="CommonListCell">
			<td width="40%"><b>名称</b><br />版块名称</td>
			<td><input size="30" name="ForumName" value="<%=ForumName%>" /> </td>
		</tr>
		<tr class="CommonListCell">
			<td width="40%"> <b>排序</b><br />
			从小到大排序设置，为“0”则隐藏此论坛</td>
			<td> 
			<input name="SortOrder" value="<%=SortOrder%>"></td>
		</tr>
		<tr class="CommonListCell">
			<td><b>版面链接</b><br />在这里输入一个网址，每当用户点击这个版面时就会链接到那个地址</td>
			<td><input type="text" name="ForumUrl" size="30" value="<%=ForumUrl%>" /></td>
		</tr>
		<tr class="CommonListCell">
			<td><b>论坛组</b><br />选择该论坛的论坛组</td>
			<td>
				<select name="GroupID" onchange="Ajax_CallBack(false,'ParentForum','loading.asp?menu=ForumList&GroupID='+this.options[this.selectedIndex].value)">


<%
	sql="Select * from ["&TablePrefix&"Groups] order by GroupID"
	Rs.open sql,conn,1
		do while not Rs.eof
			if Rs("GroupID")=GroupID then
				response.write "<option value="&Rs("GroupID")&" selected>"&rs("GroupName")&"</option>"
			else
				response.write "<option value="&Rs("GroupID")&">"&rs("GroupName")&"</option>"
			end if
		Rs.movenext
		loop
	Rs.close
	ForumList GroupID,0,ParentID
%>
				</select>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td><b>父版块</b><br />选择该论坛的父版块</td>
			<td>
				<span id=ParentForum>
				<select name="ParentID" id="ParentID">
					<option value="0">--</option>
					<%=ForumsList%>
				</select></span>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td><b>论坛版主</b><br />多版主添加请用“|”分隔，如：yuzi|裕裕 </td>
			<td><input size="40" name="Moderated" value="<%=Moderated%>" /> </td>
		</tr>
		<tr class="CommonListCell">
			<td><b>帖子类别</b><br />添加请用“|”分隔，如：原创|转帖</td>
			<td align="Left" valign="middle"><input size="40" name="TotalCategorys" value="<%=TotalCategorys%>" /></td>
		</tr>
		<tr class="CommonListCell">
			<td valign="top"><b>描述</b><br />简单描述一下该版，可以使用BBCode语法</td>
			<td><textarea rows="5" name="ForumDescription" cols="50"><%=ForumDescription%></textarea></td>
		</tr>
		<tr class="CommonListCell">
			<td valign="top"><b>规则与导读</b><br />简单描述一下该版的规则与导读，可以使用BBCode语法</td>
			<td><textarea rows="5" name="ForumRules" cols="50"><%=ForumRules%></textarea></td>
		</tr>
		<tr class="CommonListCell">
			<td><b>启用这个论坛</b><br />禁用此论坛后，该论坛将无法访问</td>
			<td><input type="radio" <%if IsActive=1 then%>checked<%end if%> value="1" name="IsActive" />是 <input type="radio" <%if IsActive=0 then%>checked<%end if%> value="0" name="IsActive" />否</td>
		</tr>
		<tr class="CommonListCell">
			<td><b>主题需要审批</b><br />启用该功能后，非信任用户该版的新主题需要审批通过了才会显示出来</td>
			<td><input type="radio" <%if ModerateNewThread=1 then%>checked<%end if%> value="1" name="ModerateNewThread" />是 <input type="radio" <%if ModerateNewThread=0 then%>checked<%end if%> value="0" name="ModerateNewThread" />否</td>
		</tr>
		<tr class="CommonListCell">
			<td><b>帖子需要审批</b><br />启用该功能后，非信任用户该版的新帖子需要审批通过了才会显示出来</td>
			<td><input type="radio" <%if ModerateNewPost=1 then%>checked<%end if%> value="1" name="ModerateNewPost" />是 <input type="radio" <%if ModerateNewPost=0 then%>checked<%end if%> value="0" name="ModerateNewPost" />否</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center" colspan="2"><input type="submit" value=" 保存 " /></td>
		</tr>
	</table>
</form>
<%	if ForumID<>0 then%>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<tr class=CommonListTitle>
		<td>版块权限</td>
		<td align="center" width="8%">浏览</td>
		<td align="center" width="8%">阅读</td>
		<td align="center" width="8%">发帖</td>
		<td align="center" width="8%">回复</td>
		<td align="center" width="8%">编辑</td>
		<td align="center" width="8%">删除</td>
		<td align="center" width="8%">发起投票</td>
		<td align="center" width="8%">投票</td>
		<td align="center" width="8%">附件</td>
		<td align="center" width="8%">管理</td>
		<td align="center" width="8%">动作</td>
	</tr>
<form method="POST" action="?menu=RoleAdd&ForumID=<%=ForumID%>" name="form">
	<tr class="CommonListCell">
		<td>
			<select name="RoleID">
<%
		sql="Select * from ["&TablePrefix&"Roles] order by RoleID"
		Set Rs=Execute(sql)
			Do While Not Rs.EOF
				Response.write("<option value='"&Rs("RoleID")&"'>"&Rs("Name")&"</option>")	
				Rs.MoveNext
			loop
		Rs.Close
%>	
			</select>
		</td>
		<td colspan="10">操作提示：您可以点击“角色名称”链接来设置该角色在该版块的权限</td>
		<td align="center"><input type="submit" value="增加" /></td>
	</tr>
</form>
<%
		sql="Select * from ["&TablePrefix&"ForumPermissions] where ForumID="&ForumID&" order by RoleID"
		Set Rs=Execute(sql)
			Do While Not Rs.EOF		
%>
	<tr class="CommonListCell">
		<td><a href="?menu=ViewForumPermissions&ForumID=<%=ForumID%>&RoleID=<%=Rs("RoleID")%>"><%=Execute("Select Name From ["&TablePrefix&"Roles] where RoleID="&Rs("RoleID")&"")(0)%></a></td>
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
		<td align="center"><input type="button" value="删除" onclick="document.location.href='?menu=DelPermissions&ForumID=<%=ForumID%>&RoleID=<%=RS("RoleID")%>'" /></td>
	</tr>
<%
				Rs.MoveNext
			loop
		Rs.Close
%>
</table>
<%
	end if
End Sub

Sub ViewForumPermissions
	sql="Select * from ["&TablePrefix&"ForumPermissions] where ForumID="&ForumID&" and RoleID="&RoleID&""
	Set Rs=Execute(sql)
%>
<b>版块权限管理</b><br />为一个给定的角色设置个人论坛权限。  
<form method="POST" action="?menu=ForumPermissionsUP&ForumID=<%=ForumID%>&RoleId=<%=RoleId%>" name="form">
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2">论坛权限</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>版块名称</b></td>
		<td width="50%"><a href="?menu=ForumAdd&ForumID=<%=ForumID%>"><%=Execute("Select ForumName From ["&TablePrefix&"Forums] where ForumID="&ForumID&"")(0)%></a></td>
	</tr>
	<tr class="CommonListCell">
		<td><b>角色名称</b></td>
		<td><a href="Admin_User.asp?menu=ViewRole&RoleID=<%=RoleID%>"><%=Execute("Select Name From ["&TablePrefix&"Roles] where RoleID="&RoleID&"")(0)%></a></td>
	</tr>
	<tr class="CommonListCell">
		<td><b>浏览</b><br />角色里的成员能够浏览帖子列表 </td>
		<td><input type="radio" value="1" name="PermissionView" <%if RS("PermissionView")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionView" <%if RS("PermissionView")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>阅读</b><br />角色里的成员能够阅读个人的帖子</td>
		<td><input type="radio" value="1" name="PermissionRead" <%if RS("PermissionRead")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionRead" <%if RS("PermissionRead")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>发帖</b><br />角色里的成员能够在论坛上发表新帖</td>
		<td><input type="radio" value="1" name="PermissionPost" <%if RS("PermissionPost")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionPost" <%if RS("PermissionPost")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>回复</b><br />角色里的成员能够回复帖子 </td>
		<td><input type="radio" value="1" name="PermissionReply" <%if RS("PermissionReply")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionReply" <%if RS("PermissionReply")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>编辑</b><br />角色里的成员能够编辑他们自己的帖子</td>
		<td><input type="radio" value="1" name="PermissionEdit" <%if RS("PermissionEdit")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionEdit" <%if RS("PermissionEdit")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>删除</b><br />角色里的成员能够删除他们自己的帖子</td>
		<td><input type="radio" value="1" name="PermissionDelete" <%if RS("PermissionDelete")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionDelete" <%if RS("PermissionDelete")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>发起投票</b><br />角色里的成员能够创建投票的选举</td>
		<td><input type="radio" value="1" name="PermissionCreatePoll" <%if RS("PermissionCreatePoll")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionCreatePoll" <%if RS("PermissionCreatePoll")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>投票</b><br />角色里的成员能够在已有的选举上投票</td>
		<td><input type="radio" value="1" name="PermissionVote" <%if RS("PermissionVote")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionVote" <%if RS("PermissionVote")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>附件</b><br />角色里的成员能够添加附件到本论坛的帖子中</td>
		<td><input type="radio" value="1" name="PermissionAttachment" <%if RS("PermissionAttachment")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionAttachment" <%if RS("PermissionAttachment")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td><b>管理</b><br />角色里的成员能够管理本版块</td>
		<td><input type="radio" value="1" name="PermissionManage" <%if RS("PermissionManage")=1 then%>checked<%end if%> />允许 <input type="radio" value="0" name="PermissionManage" <%if RS("PermissionManage")=0 then%>checked<%end if%> />拒绝</td>
	</tr>
	<tr class="CommonListCell">
		<td colspan="2" align="right"><input type="submit" value="保存" /></td>
	</tr>
</table></form>
<%
	Rs.Close
End Sub

Sub bbsManage
	ForumsList=GroupList(0)
%> 论坛数：<b><font color="red"><%=Execute("Select count(ForumID) from ["&TablePrefix&"Forums]")(0)%></font></b>　　主题数：<b><font color="red"><%=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads]")(0)%></font></b><br />
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<tr class=CommonListTitle>
		<td align=center colspan="3">批量删除主题</td>
	</tr>
	<form method="POST" action="?menu=DelThreads">
		<tr class="CommonListCell">
			<td align="center" width="50%">删除 
			<select name="TimeLimit" size="1">
				<option value="30">30</option>
				<option value="60">60</option>
				<option value="90" selected="selected">90</option>
				<option value="180">180</option>
				<option value="360">360</option>
			</select> 天没有人查看过的主题
			</td>
			<td align="center"><select name="ForumID"><option value="">所有论坛</option><%=ForumsList%></select></td>
			<td align="center"><input type="submit" value=" 确 定 " /></td>
		</tr>
	</form>
	<form method="POST" action="?menu=DelUserTopicok">
	<tr class="CommonListCell">
		<td align="center">删除 <input size="10" name="UserName" /> 发表的所有主题 
		</td>
		<td align="center"><select name="ForumID"><option value="">所有论坛</option><%=ForumsList%></select></td>
		<td align="center"><input type="submit" value=" 确 定 " /></td>
	</tr>
	</form>
	<form method="POST" action="?menu=DellikeTopicok">
	<tr class="CommonListCell">
		<td align="center">删除标题里包含有 <input size="10" name="Topic" /> 的所有主题 
		</td>
		<td align="center"><select name="ForumID"><option value="">所有论坛</option><%=ForumsList%></select></td>
		<td align="center"><input type="submit" value=" 确 定 " /></td>
	</tr>
	</form>
</table>
<br />
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<tr class=CommonListTitle>
		<td align=center colspan="2">批量删除回帖</td>
	</tr>
	<form method="POST" action="?menu=DelUserPost">
	<tr class="CommonListCell">
		<td align="center">删除 <input size="10" name="UserName" /> 发表的所有回帖</td>
		<td align="center"><input type="submit" value=" 确 定 " /></td>
	</tr>
	</form>
	<form method="POST" action="?menu=ClearRecyclePost">
	<tr class="CommonListCell">
		<td align="center">清空 <input size="10" name="TimeLimit" /> 天前删除的回帖</td>
		<td align="center"><input type="submit" value=" 确 定 " /></td>
	</tr>
	</form>
	<form method="POST" action="?menu=DellikeTopicPost">
	<tr class="CommonListCell">
		<td align="center">删除标题里包含有 <input size="10" name="Topic" /> 的所有回帖 </td>
		<td align="center"><input type="submit" value=" 确 定 " /></td>
	</tr>
	</form>
</table>
<br />
<form method="POST" action="?menu=uniteok">
	<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
		<tr class=CommonListTitle>
			<td align=center colspan="2">移动论坛帖子</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center">将 <select name="ForumSource"><%=ForumsList%></select> 移动到 <select name="ForumTarget"><%=ForumsList%></select> </td>
			<td align="center" rowspan="2"><input type="submit" value=" 确 定 " /></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center">仅移动 <input size="2" name="TimeLimit" value="0" /> 天前的帖子　　　仅移动 <input size="8" name="UserName" /> 发表的帖子</td>
		</tr>
		</table>
</form>
<%
End Sub



Sub FroumTreeView
%>

<table cellspacing="1" cellpadding="1" width="90%" border="0" align="center"><tr>
	<td align="right"><a href="?menu=AddGroup" class="CommonTextButton">新建论坛组</a></td></tr></table>

<table cellspacing="1" cellpadding="5" width="90%" border="0" class=CommonListArea align="center">
<%SelectGroup()%>
</table>
<%
End Sub
Sub SelectGroup()



	GroupGetRow = FetchEmploymentStatusList("Select GroupID,GroupName From ["&TablePrefix&"Groups] order by SortOrder")
	if IsArray(GroupGetRow) then
	For i=0 To Ubound(GroupGetRow,2)


%>
	<tr class=CommonListTitle>
		<td>　<a target="_blank" href="Default.asp?GroupID=<%=GroupGetRow(0,i)%>"><%=GroupGetRow(1,i)%></a></td>
		<td align="right" width="250">
			<a href="?menu=ForumAdd&ParentID=0&GroupID=<%=GroupGetRow(0,i)%>">新建论坛</a> | <a href="?menu=AddGroup&GroupID=<%=GroupGetRow(0,i)%>">编辑论坛组</a> | <a onclick="return window.confirm('您确定要删除该论坛组内的所有论坛?')" href="?menu=GroupManageDel&GroupID=<%=GroupGetRow(0,i)%>">删除论坛组</a>
		</td>
	</tr>
<%
		ii=ii+1
		sort GroupGetRow(0,i),0
		ii=ii-1
	Next
	end if
	%>
	
	
	<table cellspacing="1" cellpadding="1" width="90%" border="0" align="center"><tr><td align=right><a href="?menu=ManageGroups" class="CommonTextButton">切换到网格模式</a></td></tr></table>

	
	<%
	
	
End Sub

Sub sort(selec,ParentID)
	if ParentID=0 then
		sql="GroupID="&selec&" and ParentID=0"
	else
		sql="ParentID="&selec&""
	end if
	sql="Select * From ["&TablePrefix&"Forums] where "&sql&" order by SortOrder"
	Set Rs1=Execute(sql)
	do while not rs1.eof
%>
	<tr class="CommonListCell">
		<td>　<%=string(ii*2,"　")%><a target="_blank" href="ShowForum.asp?ForumID=<%=rs1("ForumID")%>"><%=rs1("ForumName")%></a></td>
		<td align="right"><a href="?menu=ForumAdd&ParentID=<%=rs1("ForumID")%>&GroupID=<%=rs1("GroupID")%>">新建子论坛</a> | <a href="?menu=ForumAdd&ForumID=<%=rs1("ForumID")%>&GroupID=<%=rs1("GroupID")%>">编辑论坛</a> | <a onclick="return window.confirm('您确定要删除该论坛的所有资料?')" href="?menu=bbsManageDel&ForumID=<%=rs1("ForumID")%>">删除论坛</a></td>
	</tr>
<%
		ii=ii+1
		sort rs1("ForumID"),1
		ii=ii-1
	rs1.Movenext
	loop
	Set Rs1 = Nothing
End Sub


Sub ShowTags
%>
<form method=post action=?menu=EditTagsUp style="margin:0px">
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle align="center">
		<td width="10%">ID</td>
		<td width="25%">标签名</td>
		<td width="17%">创建日期</td>
		<td width="15%">是否启用</td>
		<td width="10%">总帖子数</td>
		<td width="17%">最后更新日期</td>
		<td width="10%">动作</td>
	</tr>
<%
	sql="["&TablePrefix&"PostTags]"
	
	TotalCount=Execute("Select count(TagID) From "&sql&"")(0) '获取数据数量
	PageSetup=20 '设定每页的显示数量
	TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
	PageCount = RequestInt("PageIndex") '获取当前页
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	
	
	Rs.open sql,conn,1
	if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
	i=0
	Do While Not RS.EOF and i<PageSetup
		i=i+1
%>
	<tr class="CommonListCell" align="center">
		<td><%=Rs("TagID")%><input type=hidden name=TagID value=<%=Rs("TagID")%>></td>
		<td><input type=text name="TagName<%=Rs("TagID")%>" value=<%=Rs("TagName")%>></td>
		<td><%=Rs("DateCreated")%></td>
		<td>
			<input type="radio" name="IsEnabled<%=Rs("TagID")%>" value="1" <%if Rs("IsEnabled")=1 then%>checked<%end if%>>是
			<input type="radio" name="IsEnabled<%=Rs("TagID")%>" value="0" <%if Rs("IsEnabled")=0 then%>checked<%end if%>>否
		</td>
		<td><%=Rs("TotalPosts")%></td>
		<td><%=Rs("MostRecentPostDate")%></td>
		<td><a href="?menu=DelTag&TagID=<%=Rs("TagID")%>" onclick="return window.confirm('确定执行此操作？')">删除</a></td>
	</tr>
<%
		Rs.movenext
	loop
%>
</table>
<table cellspacing="0" cellpadding="0" width="100%" border="0" align="center">
	<tr class="CommonListCell">
		<td><input value=" 更 新 " type="submit" /></td>
</form>
		<td align="right"><%ShowPage()%></td>
	</tr>
</table>
<%
End Sub



AdminFooter
%>