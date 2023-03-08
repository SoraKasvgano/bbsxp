<!-- #include file="Setup.asp" -->
<%
AdminTop
if RequestCookies("UserPassword")="" or RequestCookies("UserPassword")<>session("pass") then response.redirect "Admin_Default.asp"
Log("")
LinkID=RequestInt("LinkID")
AdvertisementID=RequestInt("AdvertisementID")
TimeLimit=RequestInt("TimeLimit")
RoleID=RequestInt("RoleID")

UserName=HTMLEncode(Request("UserName"))
Body=HTMLEncode(Request("Body"))


select case Request("menu")
	case "Message"
		Message
	case "BatchSendMail"
		BatchSendMail
	case "BatchSendMailok"
		BatchSendMailok
	case "Messageok"
		if TimeLimit="" then error2("您没有选择日期！")
		Execute("Delete from ["&TablePrefix&"PrivateMessages] where DateDiff("&SqlChar&"d"&SqlChar&",CreateTime,"&SqlNowString&") > "&TimeLimit&"")
		Alert("已经将"&TimeLimit&"天以前的短讯息删除了！")
	case "DelMessageUser"
		if UserName="" then Alert("您没有输入用户名！")
		Execute("Delete from ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&UserName&"'")
		Execute("update ["&TablePrefix&"PrivateMessages] Set IsSenderDelete=1 where SenderUserName='"&UserName&"'")
		Alert("已经将"&UserName&"的短讯息全部删除了！")
	case "DelMessagekey"
		key=HTMLEncode(Request("key"))
		if key="" then Alert("您没有输入关键词！")
		Execute("Delete from ["&TablePrefix&"PrivateMessages] where Body like '%"&key&"%'")
		Alert("已经将内容中包含 "&key&" 的短讯息删除了！")
		
	case "UpStatic"
		UpdateStatic
	case "UpUserRanks"
		UpUserRanks
	case "upSiteSettingsok"
		upSiteSettingsok
	case "DelApplication"
		Application.contents.ReMoveAll()
		Alert("已经清除服务器上所有的application缓存！")
		
	case "Link"
		Link
	case "Linkok"
		Linkok
	case "DelLink"
		Execute("Delete from ["&TablePrefix&"Links] where LinkID="&LinkID&"")
		UpdateApplication"Links","Select name,Logo,url,Intro From ["&TablePrefix&"Links] where SortOrder>0 order by SortOrder"	'更新缓存
		Response.redirect"?menu=Link"
	case "AD"
		Advertisement
	case "Adok"
		Adok
	case "DelAd"
		Execute("Delete from ["&TablePrefix&"Advertisements] where AdvertisementID="&AdvertisementID&"")
		UpdateApplication"Advertisements","select Body from ["&TablePrefix&"Advertisements] where DateDiff("&SqlChar&"d"&SqlChar&","&SqlNowString&",ExpireDate) > 0"
		Response.redirect"?menu=AD"
end select

Sub BatchSendMail

if SiteConfig("SelectMailMode")="" then Alert("您没有开启发送邮件功能！")


%>

<form name="form" method="post" action="?menu=BatchSendMailok" onsubmit="return CheckForm(this);">
<input name="Body" type="hidden" />
<table cellspacing="1" cellpadding="5" width="90%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align=center colspan=2>群发邮件</td>
	</tr>
	<tr class="CommonListCell">
		<td align=Left width="12%">收件人：　</td>
		<td width="85%"><select name="RoleID">
<%
		sql="Select * from ["&TablePrefix&"Roles] where RoleID > 0 order by RoleID"
		Set Rs=Execute(sql)
			Do While Not Rs.EOF
			Response.Write("<option value='"&Rs("RoleID")&"'>"&Rs("Name")&"</option>")
				Rs.MoveNext
			loop
		Rs.Close
%>
</select>    
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align=Left>主题：</td>
		<td width="85%"><input size="60" name="Subject" /></td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top"> 内容：</td>
		<td height="250"> <script type="text/javascript" src="Editor/Post.js"></script>
		</td>
	</tr>
   <tr class="CommonListCell">
		<td align=center colspan=2> <input type="submit" value=" 发 送 " name=EditSubmit /> <input type="reset" value=" 重 填 " /><br /></td>
	</tr>
</table>
</form>

<%
End Sub

Sub BatchSendMailok
if Request("Subject")="" then Alert("请填写邮件主题")
if Request("Body")="" then Alert("请填写邮件内容")
	if RoleID>0 then
		sql="Select UserEmail from ["&TablePrefix&"Users] where UserRoleID="&RoleID&""
	else
		sql="Select UserEmail from ["&TablePrefix&"Users]"
	end if
	Set Rs=Execute(sql)
		do while not Rs.eof
			SendMail Rs("UserEmail"),Request("Subject"),Request("Body")
			Rs.Movenext
		loop
	Set Rs = Nothing
	Response.Write("邮件群发发送成功！")
End Sub

Sub Message
%>
数据库共 <%=Execute("Select count(MessageID) from ["&TablePrefix&"PrivateMessages]")(0)%> 条短讯息
<br /><br />
<table cellspacing="1" cellpadding="5" width="70%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align="center">批量删除短消息</td>
	</tr>
	<tr class="CommonListCell">
		<form method="POST" action="?menu=DelMessageUser"><td align="center">批量删除 <input size="13" name="UserName" /> 的短讯息 <input type="submit" value="确定" /></td></form>
	</tr>
	<tr class="CommonListCell">
		<form method="POST" action="?menu=DelMessagekey">
		<td align="center">批量删除内容含有 <input size="20" name="key" /> 的短讯息 <input type="submit" value="确定" />
		</td>
		</form>
	</tr>


	<tr class="CommonListCell">
		<form method="POST" action="?menu=Messageok"><td align="center">删除 <select name="TimeLimit">
			<option value="30">30</option>
			<option value="60">60</option>
			<option value="90">90</option>
			</select> 天以前的短讯息
<input type="submit" value="确定" />

		</td></form>
	</tr>

</table>
</form>
<%
End Sub

Sub Link
	if LinkID>0 then
		sql="Select * From ["&TablePrefix&"Links] where LinkID="&LinkID&""
		Set Rs=Execute(sql)
		if not Rs.eof then
			name=Rs("name")
			url=Rs("url")
			Logo=Rs("Logo")
			Intro=Rs("Intro")
			SortOrder=Rs("SortOrder")
		end if
		Rs.Close
	end if
%>
<form action=?menu=Linkok method=Post>
<input type=hidden name=LinkID value=<%=LinkID%> />
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea align=center>
		<tr class=CommonListTitle><td colspan="2">　■ 友情链接管理</td></tr>
		<tr class="CommonListCell"><td align=right>网站名称：</td><td><input size=40 name=name value="<%=name%>" /></td></tr>
		<tr class="CommonListCell"><td align=right>地址 URL：</td><td><input size=40 name=url <%if ""&URL&""<>"" then Response.write("value='"&URL&"'") else Response.write(" value='http://'") end if%> /></td></tr>
		<tr class="CommonListCell"><td align=right>图标 URL：</td><td><input size=40 name=Logo <%if ""&Logo&""<>"" then Response.write("value='"&Logo&"'") else Response.write(" value='http://'") end if%> /></td></tr>
		<tr class="CommonListCell"><td align=right>网站简介：</td><td><input size=40 name=Intro value="<%=Intro%>" /></td></tr>
		<tr class="CommonListCell"><td align=right>排　　序：</td><td><input size=5 name=SortOrder value="<%if ""&SortOrder&""<>"" then Response.write(SortOrder) else response.write("1") end if%>" /> 从小到大排序设置,为“0”则隐藏此友情链接</td></tr>
		<tr class="CommonListCell"><td colspan="2" align="center"><input type=submit value=" 确 定 " /> <input type="reset" value=" 重 填 " /></td></tr>
</table>
</form>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2">　■ 友情链接</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" width="5%"><img src="images/shareforum.gif" /></td>
		<td>
<%
	Rs.Open "["&TablePrefix&"Links] order by SortOrder",Conn
	do while not Rs.eof
		if ""&Rs("Logo")&""="" or Rs("Logo")="http://" then
			Link1=Link1+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=Link&LinkID="&Rs("LinkID")&">编辑</a></div><div class=menuitems><a href=?menu=DelLink&LinkID="&Rs("LinkID")&">删除</a></div>')"&Chr(34)&" title='"&Rs("Intro")&"' href="&Rs("url")&" target=_blank>"&Rs("name")&"</a>　　"
		else
			Link2=Link2+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=Link&LinkID="&Rs("LinkID")&">编辑</a></div><div class=menuitems><a href=?menu=DelLink&LinkID="&Rs("LinkID")&">删除</a></div>')"&Chr(34)&" title='"&Rs("name")&""&chr(10)&""&Rs("Intro")&"' href="&Rs("url")&" target=_blank><img src="&Rs("Logo")&" border=0 width=88 height=31></a>　　"
		end if
		Rs.Movenext
	loop
	Rs.close
%>
<%=Link1%>
<br /><br />
<%=Link2%>
		</td>
	</tr>
</table>
<%
End Sub


Sub Linkok
	if Request("url")="http://" or Request("url")="" then Alert("论坛URL没有填写")
	Rs.Open "["&TablePrefix&"Links]where LinkID="&LinkID&"",Conn,1,3
	If Rs.eof Then Rs.addNew
		Rs("name")=Request("name")
		Rs("url")=Request("url")
		Rs("Logo")=Request("Logo")
		Rs("Intro")=Request("Intro")
		Rs("SortOrder")=Requestint("SortOrder")
	Rs.update
	Rs.close
	UpdateApplication"Links","Select name,Logo,url,Intro From ["&TablePrefix&"Links] where SortOrder>0 order by SortOrder"	'更新缓存
	Response.redirect"?menu=Link"
End Sub


Sub Advertisement
	If AdvertisementID>0 Then
		Set Rs=Execute("Select * From ["&TablePrefix&"Advertisements] where AdvertisementID="&AdvertisementID&"")
		If Not Rs.eof Then
			Body=Rs("Body")
			ExpireDate=Rs("ExpireDate")
		End If
		Rs.close
	End If
	Body=replace(""&Body&"","<br>",vbCrlf)
	
if Not IsDate(ExpireDate) then ExpireDate=Date()+30

%>
<script type="text/javascript" src="Utility/calendar.js"></script>
<form name="form" action=?menu=Adok method="Post" onsubmit="return CheckForm(this);">
<input type=hidden name=AdvertisementID value=<%=AdvertisementID%> />

<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea align=center>
		<tr class=CommonListTitle><td colspan="2">　■ 帖间广告管理</td></tr>

		<tr class="CommonListCell">
			<td align=right width="100">广告代码：<br />（支持HTML）</td><td><textarea name="Body" cols="80" rows="8"><%=Body%></textarea></td>
		</tr>
		<tr class="CommonListCell">
			<td align=right>过期日期：</td><td><input onclick="showcalendar(event, this)" onfocus="showcalendar(event, this)" size=24 name=ExpireDate value="<%=ExpireDate%>" /></td>
		</tr>
		<tr class="CommonListCell"><td align="center" colspan="2"><input type=submit value=" 确 定 " name="EditSubmit" /> <input type="reset" value=" 重 填 " /></td></tr>
</table>
</form>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="4">　■ 帖间广告</td>
	</tr>
	<tr class=CommonListHeader align=center>
		<td>序号</td>
		<td width=50%>标题</td>
		<td>过期日期</td>
		<td>动作</td>
	</tr>
	<%
	Rs.Open "["&TablePrefix&"Advertisements] order by ExpireDate Desc",Conn
	do while not Rs.eof
	%>
	<tr class="CommonListCell">
		<td align=center><%=Rs("AdvertisementID")%></td>
		<td><%=Rs("Body")%></td>
		<td align=center><%=Rs("ExpireDate")%></td>
		<td align=center><a href="?menu=AD&AdvertisementID=<%=Rs("AdvertisementID")%>">编辑</a> | <a href="?menu=DelAd&AdvertisementID=<%=Rs("AdvertisementID")%>" onclick="return window.confirm('确认执行此操作？')">删除</a></td>
	</tr>
<%
		Rs.Movenext
	loop
	Rs.close
%>
</table>
<%
End Sub

Sub Adok
	Body=Request("Body")
	ExpireDate=Request("ExpireDate")
	if Body="" then  Alert("广告内容没有填写")

	if Not IsDate(ExpireDate) then Alert("请输入过期日期")
	Rs.Open "["&TablePrefix&"Advertisements] where AdvertisementID="&AdvertisementID&"",Conn,1,3
	If Rs.eof Then Rs.addNew
		Rs("Body")=Body
		Rs("ExpireDate")=ExpireDate
	Rs.update
	Rs.close
	UpdateApplication"Advertisements","select Body from ["&TablePrefix&"Advertisements] where DateDiff("&SqlChar&"d"&SqlChar&","&SqlNowString&",ExpireDate) > 0"
	Response.redirect"?menu=AD"
End Sub



Sub UpdateStatic
%>
<form action="?menu=UpUserRanks" method=post>
<table cellspacing="1" cellpadding="5" width="80%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle><td align="center">更新用户等级</td></tr>
	<tr class="CommonListCell"><td align="center">每次循环处理的用户数：<input type=text name=PerCircleNum value="1000" /></td></tr>
	<tr class="CommonListCell"><td align="center"><input type=submit value=" 更 新 " /></td></tr>
</table>
</form><br />
<table cellspacing="1" cellpadding="5" width="80%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align=center colspan="2">更新论坛资料</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" colspan="2">
			此操作将更新论坛资料，修复论坛统计的信息<br /><a href="?menu=upSiteSettingsok">点击这里更新论坛统计数据</a><br /><a href="?menu=DelApplication">清除服务器上的application缓存</a><br />
		</td>
	</tr>
</table>
<br />
<%
End Sub

Sub UpUserRanks
	PerCircleNum=RequestInt("PerCircleNum")
	i=RequestInt("i")
	TotalCount=RequestInt("TotalCount")
	j=0:IsEnd=false
	if PerCircleNum>0 then
		if i=0 or TotalCount=0 then TotalCount=Execute("select count(UserID) from ["&TablePrefix&"Users]")(0)/PerCircleNum
		Rs.Open "["&TablePrefix&"Users]",Conn,1,3
			Rs.move(PerCircleNum*i)
			do while j<PerCircleNum
				if Rs.eof then
					IsEnd=true
					exit do
				end if
				Rs("UserRank")=UpUserRank()
				j=j+1
				Rs.Update
				Rs.movenext
			loop
		Rs.Close
		
		if IsEnd then	'操作已完成
			Pause 0,0,0
		else			'继续执行循环更新操作
			Pause PerCircleNum,i+1,TotalCount
		end if
	end if
End Sub


Sub Pause(PerCircleNum,i,TotalCount)
	if PerCircleNum>0 then
		RefreshURL="?menu=UpUserRanks&PerCircleNum="&PerCircleNum&"&i="&i&"&TotalCount="&TotalCount&""
		response.Write("<br /><p style='text-align:center'><img src=images/loading.gif align=absmiddle />正在更新用户信息（<span style='color:#FF0000;font-weight:bold'>"&FormatPercent(i/TotalCount,,-1)&"</span>）...</p>")
%>
<script type="text/javascript">
	setTimeout('window.location.href="<%=RefreshURL%>"', 1000);
</script>
<%
	else
%>
<br />
<table cellspacing="1" cellpadding="5" width=80% class=CommonListArea>
	<tr class=CommonListTitle>
		<td width="100%" align="center">论坛提示信息</td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top" align="center">
			更新操作已完成！<br /><br />请<a href='?menu=UpStatic'>点击这里</a>返回。
		</td>
	</tr>
</table>
<%
	end if
End Sub


Sub upSiteSettingsok
	Rs.Open "["&TablePrefix&"Forums]",Conn
		do while not Rs.eof
			allarticle=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads] where Visible=1 and ForumID="&Rs("ForumID")&"")(0)
			if allarticle>0 then
				allrearticle=Execute("Select sum(TotalReplies) from ["&TablePrefix&"Threads] where Visible=1 and ForumID="&Rs("ForumID")&"")(0)
			else
				allrearticle=0
			end if
			Execute("update ["&TablePrefix&"Forums] Set TotalThreads="&allarticle&",TotalPosts="&allarticle+allrearticle&" where ForumID="&Rs("ForumID")&"")
			Rs.Movenext
		loop
	Rs.close
	Response.Write("操作成功！<br /><br />已经更新论坛所有论坛的统计数据<br />")
End Sub

AdminFooter
%>