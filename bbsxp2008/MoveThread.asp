<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")

ThreadID=Request("ThreadID")
ThreadIDArray=Split(ThreadID,",")
ThreadIDSql=0

if IsArray(ThreadIDArray) then
	for i=0 to Ubound(ThreadIDArray)
		if IsNumeric(ThreadIDArray(i)) then
			ho=int(trim(ThreadIDArray(i)))
			if Execute("Select ThreadID from ["&TablePrefix&"Threads] where ThreadID="&ho&"").eof then error"<li>系统不存在该帖子的资料"
			if ThreadIDSql=0 then ThreadIDSql=ho
		else
			error("参数错误。")
		end if
	next
else
	error("参数错误。")
end if


SQL="Select ForumID From ["&TablePrefix&"Threads] where ThreadID="&ThreadIDSql&""
Set Rs=Execute(sql)
if Rs.eof then
	error("系统不存在该主题的ID")
else
	ForumID=Rs("ForumID")
end if
Rs.close
%>
<!-- #include file="Utility/ForumPermissions.asp" -->
<%
HtmlTop
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 移动主题</div>

<form name=form method="POST" action="Manage.asp" onsubmit="return VerifyInput();">
	<input type="hidden" value="MoveThreadUp" name="menu" />
<%
if IsArray(ThreadIDArray) then
	for i=0 to Ubound(ThreadIDArray)%>
	<input type="hidden" value="<%=ThreadIDArray(i)%>" name="ThreadID" />
<%
	next
else
%>
	<input type="hidden" value="<%=ThreadID%>" name="ThreadID" />
<%
end if
%>
	
    <input type="hidden" value="<%=ForumID%>" name="ForumID" />
	<table cellspacing="1" cellpadding="5" width=100% class=CommonListArea>
		<tr class=CommonListTitle>
			<td align="center" colspan=2>移动帖子</td>
		</tr>
		<tr class="CommonListCell">
			<td align="right" width=30%>主题ID：</td>
			<td><%=ThreadID%></td>
		</tr>
		<tr class="CommonListCell">
			<td align="right">移动到：</td>
			<td>
			<select name="AimForumID" size=10 style="width:300px"><%=GroupList(0)%></select>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td colspan=2 align=center><input type="submit" value=" 确 定 " />　　<input onclick="history.back()" type="button" value=" 返 回 " /></td>
		</tr>
	</table>
</form>

<script language="JavaScript" type="text/javascript">
function VerifyInput()
{
	if (document.form.AimForumID.value == "")
	{
		alert("请选择您要移动到哪个论坛");
		document.form.AimForumID.focus();
		return false;
	}
}
</script>
<%
HtmlBottom
%>