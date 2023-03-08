<!-- #include file="Setup.asp" -->
<%
HtmlTop
if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")

if BestRole<>1 then error("你的权限不够！")
PermissionManage=1

ThreadID=Request.Form("ThreadID")


if Request_Method = "POST" then
	select case Request.Form("Menu_Item")
		case "ClearRecycle"
			TimeLimit=RequestInt("TimeLimit")
			if TimeLimit < 1 then error("只能清空24小时之前的文章")

			ThreadGetRow=FetchEmploymentStatusList("select ThreadID from ["&TablePrefix&"Threads] where Visible=2 and DateDiff("&SqlChar&"d"&SqlChar&",lasttime,"&SqlNowString&") > "&TimeLimit&"")

			if IsArray(ThreadGetRow) then
				for i=0 to ubound(ThreadGetRow,2)
					sql="select PostID from ["&TablePrefix&"Posts] where ThreadID="&ThreadGetRow(0,i)&""
					Rs.open sql,Conn,1,1
						do while not Rs.eof
							DelAttachments"Select * from ["&TablePrefix&"PostAttachments] where PostID="&Rs("PostID")&""
							Rs.movenext
						loop
					Rs.Close
				next
			end if
			ThreadGetRow=null
			
			
			Execute("Delete from ["&TablePrefix&"Threads] where Visible=2 and DateDiff("&SqlChar&"d"&SqlChar&",lasttime,"&SqlNowString&") > "&TimeLimit&"")
			succtitle="清空回收站内 "&TimeLimit&" 天以前的主题！"
		

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		case "Visible"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				sql="select * from ["&TablePrefix&"Threads] where ThreadID="&ho&""
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

		case "InVisible"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				sql="select * from ["&TablePrefix&"Threads] where ThreadID="&ho&""
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
		case "IsLocked"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				Execute("update ["&TablePrefix&"Threads] Set IsLocked=1 where ThreadID="&ho&"")
			next
			succtitle="批量锁定，主题ID："&ThreadID&""
		case "DelIsLocked"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				Execute("update ["&TablePrefix&"Threads] Set IsLocked=0 where ThreadID="&ho&"")
			next
			succtitle="批量解锁，主题ID："&ThreadID&""
		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		case "IsGood"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				Rs.open "Select ThreadID,IsGood,PostAuthor From ["&TablePrefix&"Threads] where ThreadID="&ho&"",conn,1,3
				if not Rs.eof and Rs("IsGood")=0 then
					Rs("IsGood")=1
					Rs.update
					Execute("update ["&TablePrefix&"Users] Set UserMoney=UserMoney+"&SiteConfig("IntegralAddValuedPost")&",experience=experience+"&SiteConfig("IntegralAddValuedPost")&" where UserName='"&Rs("PostAuthor")&"'")
				end if
				Rs.close
			next
			succtitle="批量精华主题，主题ID："&ThreadID&""
		case "DelIsGood"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				Rs.open "Select ThreadID,IsGood,PostAuthor From ["&TablePrefix&"Threads] where ThreadID="&ho&"",conn,1,3
				if not Rs.eof and Rs("IsGood")=1 then
					Rs("IsGood")=0
					Rs.update
					Execute("update ["&TablePrefix&"Users] Set UserMoney=UserMoney+"&SiteConfig("IntegralDeleteValuedPost")&",experience=experience+"&SiteConfig("IntegralDeleteValuedPost")&" where UserName='"&Rs("PostAuthor")&"'")
				end if
				Rs.close
			next
			succtitle="批量取消精华，主题ID："&ThreadID&""

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		case "DelThread"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				Rs.open "Select * From ["&TablePrefix&"Threads] where ThreadID="&ho&"",Conn,1,3
				if not Rs.eof and Rs("Visible")<>2 then
					if Rs("Visible")=0 then Rs("HiddenCount")=Rs("HiddenCount")-1
					Rs("DeletedCount")=Rs("DeletedCount")+1
					Rs("Visible")=2
					Rs("LastTime")=now()
					Rs("LastName")=CookieUserName
					Rs.update
					Execute("update ["&TablePrefix&"Posts] Set Visible=2 where ThreadID="&ho&" and ParentID=0")
					UpForumMostRecent(Rs("ForumID"))	'更新版块的最新主题
				end if
				Rs.close
			next
			succtitle="批量删除主题，主题ID："&ThreadID&""
		
		case "UnDelThread"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				
				Rs.open "Select * From ["&TablePrefix&"Threads] where ThreadID="&ho&"",Conn,1,3
				if not Rs.eof and Rs("Visible")<>1 then
					if Rs("Visible")=0 then Rs("HiddenCount")=Rs("HiddenCount")-1
					if Rs("Visible")=2 then Rs("DeletedCount")=Rs("DeletedCount")-1
					Rs("Visible")=1
					Rs("LastTime")=now()
					Rs("LastName")=CookieUserName
					Rs.update
					Execute("update ["&TablePrefix&"Posts] Set Visible=1 where ThreadID="&ho&" and ParentID=0")
				end if
				Rs.close
			next
			succtitle="批量还原主题，主题ID："&ThreadID&""
			
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		

		case "Fix"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				UpdateThreadStatic(ho)
			next
			succtitle="批量修复主题，主题ID："&ThreadID&""
			

		case "MoveNew"
			for each ho in Request.Form("ThreadID")
				ho=int(ho)
				Execute("update ["&TablePrefix&"Threads] Set LastTime="&SqlNowString&" where ThreadID="&ho&"")
			next
			succtitle="批量拉前主题，主题ID："&ThreadID&""


		'移动主题
		case "MoveThread"
			Response.Redirect("MoveThread.asp?ThreadID="&Request.Form("ThreadID"))
		case "MoveThreadUp"
			AimForumID=RequestInt("AimForumID")
			if AimForumID=0 then error("您没有选择要将主题移动哪个论坛")
			ThreadIDArray=split(ThreadID,",")
			for i=0 to ubound(ThreadIDArray)
				ho=int(trim(ThreadIDArray(i)))
				Execute("update ["&TablePrefix&"Threads] Set ForumID="&AimForumID&",ThreadTop=0,IsGood=0,IsLocked=0,ThreadStyle='' where ThreadID="&ho&"")
			next
			succtitle="批量移动主题，主题ID："&ThreadID&""

		
		
		
		case "ThreadVisible"
			ThreadIDStr=""
			for each ElementName in request.Form
				ElementNameArray=split(ElementName,"_")
				if ElementNameArray(0)="ThreadID" then
					ThreadID=Int(ElementNameArray(1))
					Visible=Int(Request.Form(""&ElementName&""))
					if Visible>0 then
					Rs.open "select * from ["&TablePrefix&"Threads] where ThreadID="&ThreadID&"",Conn,1,3
					if not Rs.eof and Rs("Visible")<>Visible then
						if Visible=1 then		'通过审核
							if Rs("Visible")=0 then
								Rs("HiddenCount")=Rs("HiddenCount")-1
							else
								Rs("DeletedCount")=Rs("DeletedCount")-1
							end if
						elseif Visible=2 then	'删除主题
							if Rs("Visible")=0 then Rs("HiddenCount")=Rs("HiddenCount")-1
							Rs("DeletedCount")=Rs("DeletedCount")+1
						end if
						Rs("Visible")=Visible
						Rs.update
						Execute("update ["&TablePrefix&"Posts] Set Visible="&Visible&" where ThreadID="&ThreadID&" and ParentID=0")
						ThreadIDStr=ThreadIDStr&","&ThreadID
					end if
					Rs.Close
					end if
				end if
			next

			succtitle="批量操作成功，主题ID："&ThreadIDStr&""
		'''''''''''''''''''''''''''''''''''帖子管理''''''''''''''''''''''''''''''''''
		case "PostVisible"
			PostIDStr=""
			for each FormElementName in request.Form
				FormElementNameArray=split(""&FormElementName&"","_")
				if FormElementNameArray(0)="ThreadID" then ThreadID=RequestInt(""&FormElementName&"")			
				for each ElementName in request.Form
					ElementNameArray=split(ElementName,"_")
					if Ubound(ElementNameArray)>1 then
					if ElementNameArray(0)="PostID" and ElementNameArray(1)=""&ThreadID&"" then
						PostID=Int(ElementNameArray(2))
						Visible=Int(Request.Form(""&ElementName&""))
						if Visible>0 then
						Rs.open "Select top 1 * from ["&TablePrefix&"Posts] where PostID="&PostID&"",Conn,1,3
						if not Rs.eof and Rs("Visible")<>Visible then
							Rs("Visible")=Visible
							Rs.update
							if Rs("ParentID")=0 then Execute("update ["&TablePrefix&"Threads] Set Visible="&Visible&" where ThreadID="&Rs("ThreadID")&"")
							PostIDStr=PostIDStr&","&PostID
						end if
						Rs.close
						end if
					end if
					end if
				next
				if ThreadID>0 then UpdateThreadStatic(ThreadID)
			next

			succtitle="批量操作成功，帖子ID："&PostIDStr&""

		
		case "ClearRecyclePost"
			TimeLimit=RequestInt("TimeLimit")
			if TimeLimit < 1 then error("只能清空24小时之前的回帖")

			ThreadGetRow=FetchEmploymentStatusList("Select ThreadID from ["&TablePrefix&"Threads] where Visible=1 and DeletedCount>0 order by PostTime Desc")

			if IsArray(ThreadGetRow) then
				for i=0 to ubound(ThreadGetRow,2)
					sql="select PostID from ["&TablePrefix&"Posts] where ThreadID="&ThreadGetRow(0,i)&" and ParentID>0 and Visible=2 and DateDiff("&SqlChar&"d"&SqlChar&",PostDate,"&SqlNowString&") > "&TimeLimit&""
					Rs.open sql,Conn,1,3
						do while not Rs.eof
							DelAttachments"Select * from ["&TablePrefix&"PostAttachments] where PostID="&Rs("PostID")&""
							
							Rs.Delete
							Rs.Update
							Rs.movenext
						loop
					Rs.Close
					UpdateThreadStatic(ThreadGetRow(0,i))	'主题需要更新吗？
				next
			end if
			ThreadGetRow=null
			
			succtitle="清空回收站内 "&TimeLimit&" 天以前的回帖！"

		'''''''''''''''''''''''''''''''''''帖子管理 End''''''''''''''''''''''''''''''''''
	end select

	if succtitle="" then error("无效指令")
	Log(""&succtitle&"")
	succeed succtitle,""
end if



select case Request("menu")
	case "ForumsFix"
		Rs.open "select * from ["&TablePrefix&"Forums]",Conn,1,3
		do while not Rs.eof
			allarticle=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads] where Visible=1 and ForumID="&Rs("ForumID")&"")(0)
			if allarticle>0 then
				allrearticle=Execute("Select sum(TotalReplies) from ["&TablePrefix&"Threads] where Visible=1 and ForumID="&Rs("ForumID")&"")(0)
			else
				allrearticle=0
			end if
			Rs("TotalThreads")=allarticle
			Rs("TotalPosts")=allarticle+allrearticle
			Rs.update
			UpForumMostRecent(Rs("ForumID"))
			Rs.movenext
		loop
		Rs.close
		succeed "修复论坛统计数据成功",""
	
	case "Visible"
		ForumTitle="等待审核的主题"
		ItemValue="ThreadVisible"
		Sql=" Visible=0"
		ManageThreadPost
	case "PostVisible"
		ForumTitle="等待审核的帖子"
		ItemValue="PostVisible"
		Sql=" Visible=1 and HiddenCount>0"
		ManageThreadPost
	case "Recycle"
		ForumTitle="主题回收站"
		ItemValue="UnDelThread"
		Sql="["&TablePrefix&"Threads] where Visible=2"
		ManageThread
	case "PostRecycle"
		ForumTitle="帖子回收站"
		ItemValue="PostVisible"
		Sql="where Visible=1 and DeletedCount>0"
		PostRecycle
	case else
		sql="["&TablePrefix&"Threads] where Visible=1"
		ForumTitle="主题管理"
		ManageThread
end select




Sub CommonBread
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → <a href="?menu=<%=Request("menu")%>"><%=ForumTitle%></a></div>
<table border="0" width="100%" align="center">
	<tr>
		<td align=right>
			<select name="menu" onchange="window.location.href='?menu='+this.options[this.selectedIndex].value">
            <optgroup label="审核区">
            <option value="Visible"<%if Request("menu")="Visible" then Response.Write(" selected")%>>主题审核</option>
            <option value="PostVisible"<%if Request("menu")="PostVisible" then Response.Write(" selected")%>>帖子审核</option>
			</optgroup>
            <optgroup label="回收站">
            <option value="Recycle"<%if Request("menu")="Recycle" then Response.Write(" selected")%>>主题回收站</option>
            <option value="PostRecycle"<%if Request("menu")="PostRecycle" then Response.Write(" selected")%>>帖子回收站</option>
			</optgroup>
			</select>
		</td>
	</tr>
</table>
<%
End Sub






Sub PostRecycle
	CommonBread
%>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="6"><%=ForumTitle%></td>
	</tr>
	<tr class="CommonListHeader" align="center">
		<td>标题</td>
		<td>作者</td>
		<td>回复</td>
		<td>查看</td>
		<td>最后更新</td>
	</tr>
<%
	TotalCount=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads]"&Sql)(0) '获取数据数量
	PageSetup=SiteConfig("ThreadsPerPage") '设定每页的显示数量
	TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
	PageCount = RequestInt("PageIndex") '获取当前页
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	Sql="Select * from ["&TablePrefix&"Threads]"&Sql&" order by LastTime desc"
	if PageCount<11 then
		Set Rs=Execute(Sql)
	else
		rs.Open Sql,Conn,1
	end if
	
	if TotalPage>1 then Rs.Move (PageCount-1) * PageSetup

	If Rs.Eof Then Rs.close:Response.write("</table>"):Exit Sub
	i=0
	
		do while not Rs.eof and i<PageSetup
			ShowThreadForPostRecycle()
			i=i+1
			Rs.movenext
		loop
		Rs.close
		Set Rs = Nothing
%>
</table>
</form>
<table border="0" width="100%" align="center">	<tr>
		<td valign="top"><%ShowPage()%></td>
		<form name="form" method="POST" action="Moderation.asp">
		<td align="right">
        <input type="hidden" name="Menu_Item" value="ClearRecyclePost" />清空 <input size="1" value="7" name="TimeLimit" /> 天以前删除的帖子 <input type="submit" onclick="return window.confirm('执行本操作将清空回收站的帖子?');" value="确定" />
		</td>
		</form>
	</tr>
</table>
<%

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ManageThreadPost
	CommonBread
%>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
<form method="POST" action="Moderation.asp" style="margin:0px">
	<tr class="CommonListTitle">
		<td colspan="2"><input type="hidden" name="Menu_Item" value="<%=ItemValue%>" /><%=ForumTitle%></td>
	</tr>
	<tr class="CommonListHeader">
		<td colspan="2" align="center">
        	<label for="Visible"><input type="button" id="Visible" onclick="CheckRadioAll(this.form,'1')" value=确认 /></label> 
        	<label for="Delete"><input type="button" id="Delete" onclick="CheckRadioAll(this.form,'2')" value=删除 /></label> 
        	<label for="Cancel"><input type="button" id="Cancel" onclick="CheckRadioAll(this.form,'0')" value=忽略 /></label>
        </td>
	</tr>
<%
	TotalCount=Execute("Select count(ThreadID) From ["&TablePrefix&"Threads] where"&Sql&"")(0) '获取数据数量
	PageSetup=SiteConfig("ThreadsPerPage")			'设定每页的显示数量
	TotalPage=Abs(Int(TotalCount/PageSetup*(-1)))	'总页数
	PageCount = RequestInt("PageIndex")				'获取当前页
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage

	sql="Select ThreadID,PostTime,Visible,ForumID,Topic from ["&TablePrefix&"Threads] where"&Sql&" order by PostTime Desc"
	Rs.open sql,Conn,1,1
	
	if TotalPage>1 then Rs.Move (PageCount-1) * PageSetup

	If Rs.Eof Then Rs.close:Response.write("</table>"):Exit Sub
	ThreadGetRow=Rs.GetRows(PageSetup)
	Rs.Close


	if IsArray(ThreadGetRow) then
		InputHidden=""
		for i=0 to ubound(ThreadGetRow,2)
			if ItemValue="PostVisible" then
				set Rs=Execute("select * from ["&TablePrefix&"Posts] where ThreadID="&ThreadGetRow(0,i)&" and Visible=0 and ParentID>0")
				If Rs.eof then
					UpdateThreadStatic(ThreadGetRow(0,i))
				else
					InputHidden=InputHidden&"<input type=hidden name='ThreadID_"&ThreadGetRow(0,i)&"' value="&ThreadGetRow(0,i)&" />"&vbnewline
					
					do while not Rs.eof
						ShowInvisiblePost"PostID",ThreadGetRow(0,i),ThreadGetRow(4,i)
						Rs.movenext
					loop
				end if
			elseif ItemValue="ThreadVisible" then
				set Rs=Execute("select * from ["&TablePrefix&"Posts] where ThreadID="&ThreadGetRow(0,i)&" and Visible=0 and ParentID=0")
				If Rs.eof then
					UpdateThreadStatic(ThreadGetRow(0,i))
				else
					ShowInvisiblePost"ThreadID",0,""
				end if
			end if
		next
		Rs.Close
	end if
	ThreadGetRow=null
%>
	<tr class="CommonListHeader">
	  <td colspan="2" align="center"><%=InputHidden%><input onclick="return window.confirm('您确定执行本次操作?');" type="submit" value=" 确 定 " /></td>
	</tr>
</form>
</table>

<table border="0" width="100%" align="center">
	<tr>
		<td><%ShowPage()%></td>
		<form name="form" method="POST" action="Moderation.asp">
		<td align="right">
		<%if Request("menu")="Recycle" then%>
        <input type="hidden" name="Menu_Item" value="ClearRecycle" />	清空 <input size="1" value="7" name="TimeLimit" /> 天以前删除的主题 <input type="submit" onclick="return window.confirm('执行本操作将清空回收站的主题?');" value="确定" />
        <%elseif Request("menu")="PostRecycle" then%>
        <input type="hidden" name="Menu_Item" value="ClearRecyclePost" />清空 <input size="1" value="7" name="TimeLimit" /> 天以前删除的帖子 <input type="submit" onclick="return window.confirm('执行本操作将清空回收站的帖子?');" value="确定" />
		<%end if%>
		</td>
		</form>
	</tr>
</table>

<%
End Sub

Sub ShowInvisiblePost(FieldName,ThreadID,Topic)
	if int(ThreadID)>0 then ThreadIDStr="_"&ThreadID
%>
	<tr class="CommonListCell">
		<td align=right width=20%>作　者：</td>
		<td><a href="Profile.asp?UserName=<%=Rs("PostAuthor")%>"><%=Rs("PostAuthor")%></a></td>
	</tr>
	<tr class="CommonListCell">
		<td align=right>时　间：</td>
		<td><%=Rs("PostDate")%></td>
	</tr>
    <%if Topic<>"" then%>
	<tr class="CommonListCell">
		<td align=right>主　题：</td>
		<td><a href="ShowPost.asp?ThreadID=<%=ThreadID%>" target="_blank"><%=Topic%></a></td>
	</tr>
    <%end if%>
	<tr class="CommonListCell">
		<td align=right>标　题：</td>
		<td><input type=text value="<%=Rs("Subject")%>" size="60" readonly="readonly" /></td>
	</tr>
	<tr class="CommonListCell">
		<td align=right>内　容：</td>
		<td><textarea cols=80 rows=5 readonly="readonly"><%=Rs("Body")%></textarea></td>
	</tr>

	<tr class="CommonListCell">
		<td align=right>操　作：</td>
		<td>
        	<label for="Visible<%=Rs(""&FieldName&"")%>"><input type="radio" name="<%=FieldName&ThreadIDStr&"_"&Rs(""&FieldName&"")%>" id="Visible<%=Rs(""&FieldName&"")%>" value=1 />确认</label> 
        	<label for="Delete<%=Rs(""&FieldName&"")%>"><input type="radio" name="<%=FieldName&ThreadIDStr&"_"&Rs(""&FieldName&"")%>" id="Delete<%=Rs(""&FieldName&"")%>" value=2 />删除</label> 
        	<label for="Cancel<%=Rs(""&FieldName&"")%>"><input type="radio" name="<%=FieldName&ThreadIDStr&"_"&Rs(""&FieldName&"")%>" id="Cancel<%=Rs(""&FieldName&"")%>" value=0 checked="checked" />忽略</label>
        </td>
	</tr>
	<tr class="CommonListHeader">
		<td colspan="2"></td>
	</tr>
<%
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ManageThread
	CommonBread
%>
<form method="POST" action="Moderation.asp" style="margin:0px">
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="6"><%=ForumTitle%></td>
	</tr>
	<tr class="CommonListHeader" align="center">
		<td>标题</td>
		<td>作者</td>
		<td>回复</td>
		<td>查看</td>
		<td>最后更新</td>
		<td width="50"><input type="checkbox" name="chkall" onclick="CheckAll(this.form)" value="ON" /></td>
	</tr>
<%
	TotalCount=Execute("Select count(ThreadID) From "&sql&" ")(0) '获取数据数量
	PageSetup=SiteConfig("ThreadsPerPage") '设定每页的显示数量
	TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
	PageCount = RequestInt("PageIndex") '获取当前页
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage

	sql="Select * from "&sql&" order by lasttime Desc"
	if PageCount<11 then
		Set Rs=Execute(sql)
	else
		rs.Open sql,Conn,1
	end if
	if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
	i=0
	Do While Not RS.EOF and i<PageSetup
		i=i+1
		ShowThread()
		Rs.MoveNext
	loop
	Rs.Close
%>
</table>
<table border="0" width="100%" align="center">
	<tr>
		<td align="right">
<%if Request("menu")="Recycle" then%>
		<input type="hidden" name="Menu_Item" value="<%=ItemValue%>" />
		<input onclick="return window.confirm('您确定执行本次操作?');" type="submit" value=" 还 原 " />
<%else%>管理选项：<select name=Menu_Item size=1>
			<optgroup label="审核">
				<option value="Visible">审核通过</option>
				<option value="InVisible">审核失败</option>
			</optgroup>
			<optgroup label="解锁">
				<option value="IsLocked">锁定主题</option>
				<option value="DelIsLocked">解锁主题</option>
				
			</optgroup>
			<optgroup label="精华">
				<option value="IsGood">精华主题</option>
				<option value="DelIsGood">取消精华</option>
			</optgroup>
			<optgroup label="删除">
				<option value="DelThread">删除主题</option>
				<option value="UnDelThread">还原主题</option>
			</optgroup>
			
			<optgroup label="其它">
				<option value="Fix">修复主题</option>
				<option value="MoveNew">拉前主题</option>
				<option value="MoveThread">移动主题</option>
			</optgroup>
		</select>　
		<input type="submit" value=" 执 行 " onclick="return VerifyRadio('Item');" />
<%end if%>
		</td>
	</tr>
</table>
</form>
<table border="0" width="100%" align="center">	<tr>
		<td valign="top"><%ShowPage()%></td>
		<%if Request("menu")="Recycle" then%>
		<form name="form" method="POST" action="Moderation.asp">
		<td align="right">
        <input type="hidden" name="Menu_Item" value="ClearRecycle" />	清空 <input size="1" value="7" name="TimeLimit" /> 天以前删除的主题 <input type="submit" onclick="return window.confirm('执行本操作将清空回收站的主题?');" value="确定" />
		</td>
		</form>
		<%end if%>
	</tr>
</table>
<%
End Sub

%>
<script language="JavaScript" type="text/javascript">
function VerifyRadio() {
	objYN=false;
	if (window.confirm('您确定执行本次操作?')){
		for (i=0;i<document.getElementsByName("ThreadID").length;i++) {
			if (document.getElementsByName("ThreadID")[i].checked) {objYN= true;}
		}
		if (objYN==false) {alert ('请选择您要操作的主题！');return false;}
	}
	return objYN;
}
function CheckRadioAll(FormObject,ElementValue) {
	for (var i =0; i < FormObject.elements.length; i++) {
		var elm = FormObject.elements[i];
		if (elm.type == 'radio' && elm.value == ElementValue){
				elm.checked = true;
		}
	}
}
</script>

<%
HtmlBottom



Sub ShowThreadForPostRecycle()
	if Rs("ThreadTop")=2 then
		IconImage="topic-announce.gif alt='公告主题'"
	elseif Rs("ThreadTop")=1 then
		IconImage="topic-pinned.gif alt='置顶主题'"
	elseif Rs("IsGood")=1 then
		IconImage="topic-popular.gif alt='精华主题'"
	elseif Rs("IsLocked")=1 then
		IconImage="topic-locked.gif alt='主题锁定'"
	elseif Rs("IsVote")=1 then
		IconImage="topic-poll.gif alt='投票主题'"
	elseif DateDiff("d",Rs("PostTime"),Now()) <= SiteConfig("PopularPostThresholdDays")   and  (  Rs("TotalReplies")=>SiteConfig("PopularPostThresholdPosts") or  Rs("TotalViews")=>SiteConfig("PopularPostThresholdViews") ) then
		IconImage="topic-hot.gif alt='热门主题'"
	else
		IconImage="topic.gif alt='普通主题'"
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
			ThreadStatus="<img src=images/status_Answered.gif align=middle title='主题状态：已解决'>"
		elseif Rs("ThreadStatus")=2 then
			ThreadStatus="<img src=images/status_NotAnswered.gif align=middle title='主题状态：未解决'>"
		else
			ThreadStatus="<img src=images/status_NotSet.gif align=middle>"
		end if
	end if
	if Rs("HiddenCount")>0 then
		ThreadHidden="&nbsp;<img src='images/InVisible.gif' align=middle alt='"&Rs("HiddenCount")&"未审核帖子' />"
	end if
	if Rs("DeletedCount")>0 then
		ThreadDel="&nbsp;<img src='images/recycle.gif' align=middle alt='"&Rs("DeletedCount")&"已删除帖子' />"
	end if
	
	if DateDiff("d",Rs("PostTime"),Now()) < 2 then
		NewHtml=" <img title='一天内新发表的主题' src=images/new.gif align=absmiddle>"
	else
		NewHtml=""
	end if
	if Rs("TotalRatings")>0 then StarHtml="<a style=CURSOR:pointer onclick="&CHR(34)&"OpenWindow('PostRating.asp?ThreadID="&Rs("ThreadID")&"')"&CHR(34)&" ><img border=0 src=Images/Star/"&cint(Rs("RatingSum")/Rs("TotalRatings"))&".gif align=middle></a>&nbsp;"
	if Rs("TotalReplies")=>SiteConfig("PostsPerPage") then
		MaxPostPage=fix(Rs("TotalReplies")/SiteConfig("PostsPerPage"))+1 '共多少页
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
						<td<%if SiteConfig("EnablePostPreviewPopup")=1 then%> title='<%=Rs("Description")%>'<%end if%>><%=checkboxHtml%><%=ThreadEmoticonID%><%=CategoryHtml%><%
						
						if Rs("Visible")=0 then
							Response.write("等待审核：")
						elseif Rs("Visible")=2 then
							 Response.write("已被删除：")
					    end if
						
						
						%><a href="ShowPost.asp?ThreadID=<%=Rs("ThreadID")%>"><span style="<%=Rs("ThreadStyle")%>"><%=Rs("Topic")%></span></a><%=ShowPostPage%><%=NewHtml%></td>
						<td align="right"><%=StarHtml&ThreadStatus%><%if PermissionManage=1 then Response.write(ThreadHidden&ThreadDel)%></td>
					</tr>
				</table>

		</td>
		
		<td align="center" width="13%"><a href="Profile.asp?UserName=<%=Rs("PostAuthor")%>"><%=Rs("PostAuthor")%></a><br /><%=FormatDateTime(Rs("PostTime"),2)%></td>
		<td align="center" width="7%"><%=replies%></td>
		<td align="center" width="7%"><%=Rs("TotalViews")%></td>
		<td align="center" width="18%"><%=Rs("lasttime")%><br />by <a href="Profile.asp?UserName=<%=Rs("lastname")%>"><%=Rs("lastname")%></a></td>
	</tr>
<%
End Sub
%>