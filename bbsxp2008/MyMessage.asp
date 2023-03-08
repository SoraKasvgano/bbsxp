<!-- #include file="Setup.asp" --><%

RecipientUserName=HTMLEncode(Request("RecipientUserName"))
UserName=HTMLEncode(Request("UserName"))
Subject=HTMLEncode(Request("Subject"))
Body=HTMLEncode(Request("Body"))
box=HTMLEncode(Request("box"))
MessageID=RequestInt("MessageID")
ForumID=RequestInt("ForumID")

MaxPrivateMessageSize=SiteConfig("MaxPrivateMessageSize")

if Request("menu")="Post" then
	if CookieUserName=empty then AlertForModal("您还未登录论坛")

	if MessageID > 0 then
		Rs.open "Select top 1 * from ["&TablePrefix&"PrivateMessages] where MessageID="&MessageID&"",conn,1
		if rs.eof then AlertForModal("引用/回复的短消息不存在")
		Subject=Rs("Subject")
		Body=replace(Rs("Body"),"<br>",vbcrlf)

		if RecipientUserName<>empty then	'回复
			Subject="RE: "&Subject
			Body=""
		else					'转发
			Body=vbcrlf&vbcrlf&vbcrlf&vbcrlf&"------------------------------------"&vbcrlf&"日期: "&Rs("CreateTime")&vbcrlf&"发件人: "&Rs("SenderUserName")&vbcrlf&"收件人: "&Rs("RecipientUserName")&vbcrlf&"主题: "&Subject&vbcrlf&vbcrlf&Body&""
			Subject="FW: "&Subject
		end if
		Rs.close
	end if
Response.clear
%>
<title>发送讯息</title>
<style type="text/css">body{FONT-SIZE:9pt;}</style>
<form name="form" action="MyMessage.asp" method="POST" onsubmit="return CheckMessage()">
<input type="hidden" name="menu" value="addPost" />
<input type="hidden" name="ForumID" value="<%=ForumID%>" />
发件人：<input type="text" name="SendUserName" value="<%=CookieUserName%>" readonly="readonly" /><br />
收件人：<input id="RecipientUserName" name="RecipientUserName" type="text" value="<%=RecipientUserName%>" />

<%
if ForumID > 0 then
	if RequestCookies("LastReportTime")<>"" then
		ReportInterval=SiteConfig("PostInterval")*3
		StopPostTime=int(DateDiff("s",RequestCookies("LastReportTime"),Now()))
		if StopPostTime < ReportInterval then AlertForModal("论坛限制一个人问题帖子两次报告间隔必须大于 "&ReportInterval&" 秒！您必须再等待 "&ReportInterval-StopPostTime&" 秒！")
	end if
%>
<select onchange="Javascript:document.getElementById('RecipientUserName').value=this.value">
<%
	ModeratedList=Execute("Select Moderated from ["&TablePrefix&"Forums] where ForumID="&ForumID&"")(0)
	if ModeratedList<>"" then
		Response.Write("<option value=''>版主列表</option>")
		Moderated=split(ModeratedList,"|")
		for i=0 to Ubound(Moderated)
			Response.Write("<option value="&Moderated(i)&">"&Moderated(i)&"</option>")
		next
	else
		Response.Write("<option value=''>管理员列表</option>")
		ModeratedList=(0)
		Set Rs=Execute("Select UserName from ["&TablePrefix&"Users] where (UserRoleID=1 or UserRoleID=2) and UserAccountStatus=1")
		do while not Rs.eof
			Response.Write("<option value="&Rs("UserName")&">"&Rs("UserName")&"</option>")
			Rs.movenext
		loop
	end if
%>
</select>
<%
end if
%><br />
主　题：<input size="60" name="Subject" value="<%=Subject%>" /><br />
<textarea name="Body" style="width:100%;height:200px"><%=Body%></textarea><br />


<input type="checkbox" name="IsSaveOutbox" id="IsSaveOutbox" value="1" /><label for=IsSaveOutbox>保存到已发送</label>

<div style="float:right"><input type="submit" value="发送讯息" />　</div>
</form>

<script language="JavaScript" type="text/javascript">
function CheckMessage(){
	var FormObject = document.form;
	if (FormObject.elements["RecipientUserName"].value == "")
	{
		alert("请输入收件人名字！");
		FormObject.elements["RecipientUserName"].focus()
		return false;
	}
	if ( FormObject.elements["Subject"].value.length < 2)
	{
		alert("请输入主题不能少于2个字节！");
		FormObject.elements["Subject"].focus()
		return false;
	}
	if (FormObject.elements["Body"].value.length < 2)
	{
		alert("请输入内容不能少于2个字节！");
		FormObject.elements["Body"].focus()
		return false;
	}
	if (FormObject.elements["RecipientUserName"].value == FormObject.elements["SendUserName"].value)
	{
		alert("不能自己给自己发送讯息！");
		FormObject.elements["RecipientUserName"].focus()
		return false;
	}
	return true;
}
</script>
<%


elseif Request("menu")="addPost" then
	if CookieUserName=empty then AlertForModal("您还未登录论坛")

	IsSaveOutbox=RequestInt("IsSaveOutbox")
	IsSenderDelete=0
	if IsSaveOutbox=0 then IsSenderDelete=1
	if Len(Subject)<2 then Alert("标题不能小于 2 字符")
	if Len(Body)<2 then Alert("内容不能小于 2 字符")
	if RecipientUserName="" or Lcase(RecipientUserName)=Lcase(CookieUserName) then Alert("请输入接收的对象且不能是自己")
	if Execute("Select UserID From ["&TablePrefix&"Users] where UserName='"&RecipientUserName&"'").eof Then Alert("系统不存在"&RecipientUserName&"的资料")
	
	if IsSaveOutbox=1 and Execute("Select count(MessageID) From ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&CookieUserName&"' and IsRecipientDelete=0 or SenderUserName='"&CookieUserName&"' and IsSenderDelete=0")(0) > MaxPrivateMessageSize Then
		AddApplication "Message_"&CookieUserName,"【系统讯息】您的短讯文件夹空间已满，请整理文件夹，否则您将不能保存任何短讯！"
		Alert("您的短信箱已满不能保存短讯！本次发送失败！")
	end if

	if Execute("Select count(MessageID) From ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&RecipientUserName&"' and IsRecipientDelete=0 or SenderUserName='"&RecipientUserName&"' and IsSenderDelete=0")(0)>MaxPrivateMessageSize Then
		AddApplication "Message_"&RecipientUserName,"【系统讯息】您的短讯文件夹空间已满，请整理文件夹，否则您将收不到任何短讯！"
		Alert("用户"&RecipientUserName&"的短信箱已满！本次发送失败！")
	end if
	
	sql="insert into ["&TablePrefix&"PrivateMessages](SenderUserName,RecipientUserName,Subject,Body,IsSenderDelete) values ('"&CookieUserName&"','"&RecipientUserName&"','"&Subject&"','"&Body&"',"&IsSenderDelete&")"
	Execute(sql)
	Execute("update ["&TablePrefix&"Users] Set NewMessage=NewMessage+1 where UserName='"&RecipientUserName&"'")
	if ForumID>0 then ResponseCookies "LastReportTime",""&now()&"","999"
%>
<script language="JavaScript" type="text/javascript">parent.BBSXP_Modal.Close();</script>
<%


else
	if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")
	HtmlTop
	RecipientCount=Execute("Select count(messageID) from ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&CookieUserName&"' and IsRecipientDelete=0 or SenderUserName='"&CookieUserName&"' and IsSenderDelete=0")(0)

UsedMessageSize=RecipientCount/MaxPrivateMessageSize*100

%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 短信服务</div>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td align="center"><a href="EditProfile.asp">资料修改</a></td>
		<td align="center"><a href="EditProfile.asp?menu=pass">密码修改</a></td>
		<td align="center"><a href="MyUpFiles.asp">上传管理</a></td>
		<td align="center"><a href="MyFavorites.asp">收 藏 夹</a></td>
		<td align="center"><a href="MyMessage.asp">短信服务</a></td>
	</tr>
</table>
<br />
<table border="0" width="100%" id="table14" cellspacing="0" cellpadding="0">
	<tr>
		<td width="150" valign="top">
		<table width="100%" cellspacing="1" cellpadding="5" border="0" class="CommonListArea">
			<tr class="CommonListTitle">
				<td><b>文件夹</b></td>
			</tr>
			<tr class="CommonListCell">
				<td><img src="images/i_inbox.gif" align="absmiddle" />
				<a href="?menu=Inbox">收件箱</a>[<b><%=Execute("Select count(messageID) from ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&CookieUserName&"' and IsRecipientDelete=0")(0)%></b>]</td>
			</tr>
			<tr class="CommonListCell">
				<td><img src="images/i_sent.gif" align="absmiddle" />
				<a href="?menu=Outbox">已发送</a>[<b><%=Execute("Select count(messageID) from ["&TablePrefix&"PrivateMessages] where SenderUserName='"&CookieUserName&"' and IsSenderDelete=0")(0)%></b>]</td>
			</tr>
		</table>
		<br />
		<table cellspacing="1" cellpadding="1" width="100%">
			<tr>
				<td>
				<div class="percent">
					<div style="width:<%=UsedMessageSize%>%">
					</div>
				</div>
				<%=UsedMessageSize%>% （<%=RecipientCount%>/<%=MaxPrivateMessageSize%>）</td>
			</tr>
		</table>
		</td>
		<td width="1%"></td>
		<td valign="top"><%
select case Request("menu")

	case "Read"
		ReadMessage

	case "Delete"
		for each ho in Request("MessageID")
			ho=int(ho)
			rs.open "Select * from ["&TablePrefix&"PrivateMessages] where ( RecipientUserName='"&CookieUserName&"' or SenderUserName='"&CookieUserName&"' ) and MessageID="&ho&"",Conn,1,3
				if LCASE(rs("RecipientUserName"))=LCASE(CookieUserName) then rs("IsRecipientDelete")=1
				if LCASE(rs("SenderUserName"))=LCASE(CookieUserName) then rs("IsSenderDelete")=1
			rs.update
			rs.close
		next
		Execute("Delete from ["&TablePrefix&"PrivateMessages] where IsRecipientDelete=1 and IsSenderDelete=1")

		response.redirect("?")

	case else
		InOutbox
end select
%> </td>
	</tr>
</table>
<br />
<%
Sub InOutbox
MessageMenu=Request("Menu")
if MessageMenu=empty then MessageMenu="Inbox"

	if MessageMenu="Inbox" then
		tdstr="<td align=center width='15%'>发件人</td>"
		sql="Select * from ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&CookieUserName&"' and IsRecipientDelete=0 order by MessageID Desc"
	else
		tdstr="<td align=center width='15%'>收件人</td>"
		sql="Select * from ["&TablePrefix&"PrivateMessages] where SenderUserName='"&CookieUserName&"' and IsSenderDelete=0 order by MessageID Desc"
	end if
%>
<form name="showMessage" action="?Menu=Delete" method="post">
	<table cellspacing="1" cellpadding="3" width="100%" class="CommonListArea">
		<tr class="CommonListTitle">
			<td colspan="5">
			<a href="javascript:BBSXP_Modal.Open('MyMessage.asp?Menu=Post',600,350);">
			<img src="images/i_new_msg.gif" border="0" align="absmiddle" />新建</a>　<a onclick="if(window.confirm('您确定执行本次操作?')){document.showMessage.submit();}"><img src="images/i_delete.gif" border="0" align="absmiddle" />删除</a></td>
		</tr>
		<tr class="CommonListHeader">
			<td align="center" width="6%">
			<input type="checkbox" name="chkall" onclick="CheckAll(this.form)" value="ON" /></td>
			<%=tdstr%>
			<td align="center" width="46%">主题</td>
			<td align="center" width="22%">日期</td>
			<td align="center" width="10%">大小</td>
		</tr>
		<%
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
		<tr class="CommonListCell" id="PrivateMessage<%=Rs("MessageID")%>" <%if rs("isread")=0 then response.write" style='font-weight:bold'"%>>
			<td valign="center" align="center" width="5%">
			<input type="checkbox" value="<%=Rs("MessageID")%>" name="MessageID" onclick="CheckSelected(this.form,this.checked,'PrivateMessage<%=Rs("MessageID")%>')" /></td>
			<%if MessageMenu="Inbox" then%>
			<td align="center">
			<a href="Profile.asp?UserName=<%=Rs("SenderUserName")%>" target="_blank">
			<%=Rs("SenderUserName")%></a></td>
			<%else%>
			<td align="center">
			<a href="Profile.asp?UserName=<%=Rs("RecipientUserName")%>" target="_blank">
			<%=Rs("RecipientUserName")%></a></td>
			<%end if%>
			<td width="10%">
			<a href="?Menu=Read&MessageID=<%=Rs("MessageID")%>&box=<%=MessageMenu%>">
			<%=Rs("Subject")%></a></td>
			<td align="center" width="15%"><%=Rs("CreateTime")%></td>
			<td align="center" width="10%"><%=CheckSize(Len(""&Rs("Body")&""))%></td>
		</tr>
		<%
		Rs.MoveNext
	loop
	Rs.Close
%>
	</table>
	<table cellspacing="1" cellpadding="1" width="100%">
		<tr>
			<td colspan="5"><%ShowPage()%></td>
		</tr>
	</table>
</form>
<%
End Sub

Sub ReadMessage

if box="Inbox" then
	FieldName="RecipientUserName"
	DeleteFieldName="IsRecipientDelete"
else
	FieldName="SenderUserName"
	DeleteFieldName="IsSenderDelete"
end if

if Request("pageIndex")="previous" then
	SQL="Select top 1 * from ["&TablePrefix&"PrivateMessages] where "&FieldName&"='"&CookieUserName&"' and "&DeleteFieldName&"=0 and MessageID>"&MessageID&" order by messageID"

elseif Request("pageIndex")="next" then
	SQL="Select top 1 * from ["&TablePrefix&"PrivateMessages] where "&FieldName&"='"&CookieUserName&"' and "&DeleteFieldName&"=0 and MessageID<"&MessageID&" order by messageID desc"

else
	SQL="Select top 1 * from ["&TablePrefix&"PrivateMessages] where (RecipientUserName='"&CookieUserName&"' or SenderUserName='"&CookieUserName&"') and MessageID="&MessageID&""
end if


Rs.Open SQL,Conn,1,3

if Http_Referer="" then Http_Referer="Default.asp"
if Rs.eof then response.redirect Http_Referer

%>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td>
		<div style="float:left">
			<a href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post', 600, 350);">
			<img src="images/i_new_msg.gif" border="0" align="absmiddle" />新建</a>　<a href="?menu=Delete&MessageID=<%=Rs("MessageID")%>" onclick="return window.confirm('您确定执行本次操作?');"><img src="images/i_delete.gif" border="0" align="absmiddle" />删除</a><%if Rs("SenderUserName")<>CookieUserName then%>　<img src="images/i_reply.gif" border="0" align="absmiddle" /><a href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&MessageID=<%=Rs("MessageID")%>&RecipientUserName=<%=Rs("SenderUserName")%>',600,350);">回复</a><%end if%>　<a href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&MessageID=<%=Rs("MessageID")%>',600,350);"><img src="images/i_forward.gif" border="0" align="absmiddle" />转发</a></div>
		<div style="float:right">
			<a href="?menu=Read&MessageID=<%=Rs("messageID")%>&pageIndex=previous&box=<%=box%>" title="上一条">
			<img src="images/i_previous.gif" border="0" align="absmiddle" /></a>
			<a href="?menu=Read&MessageID=<%=Rs("messageID")%>&pageIndex=next&box=<%=box%>" title="下一条">
			<img src="images/i_next.gif" border="0" align="absmiddle" /></a></div>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td width="10%" style="line-height:150%">
		<div style="float:left">
			发件人：<a target="_blank" href="Profile.asp?UserName=<%=Rs("SenderUserName")%>"><%=Rs("SenderUserName")%></a><%if Rs("SenderUserName")<>CookieUserName then%>　<a href="javascript:Ajax_CallBack(false,false,'MyFavorites.asp?menu=FavoriteFriend&FriendUserName=<%=Rs("SenderUserName")%>',true);">+添加好友</a><%end if%></div>
		<div style="float:right">
			<%=Rs("CreateTime")%></div>
		<br />
		收件人：<a target="_blank" href="Profile.asp?UserName=<%=Rs("RecipientUserName")%>"><%=Rs("RecipientUserName")%></a><%if Rs("RecipientUserName")<>CookieUserName then%>　<a href="javascript:Ajax_CallBack(false,false,'MyFavorites.asp?menu=FavoriteFriend&FriendUserName=<%=Rs("RecipientUserName")%>',true);">+添加好友</a><%end if%>
		<br />
		主　题：<%=Rs("Subject")%></td>
	</tr>
	<tr class="CommonListCell">
		<td><%=Rs("Body")%></td>
	</tr>
</table>
<%
	if LCASE(Rs("RecipientUserName"))=LCASE(CookieUserName) and Rs("IsRead")=0 then Rs("IsRead")=1
	Rs.update
	Rs.close


End Sub

HtmlBottom

end if
%>