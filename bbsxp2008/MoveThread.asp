<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")

ThreadID=Request("ThreadID")
ThreadIDArray=Split(ThreadID,",")
ThreadIDSql=0

if IsArray(ThreadIDArray) then
	for i=0 to Ubound(ThreadIDArray)
		if IsNumeric(ThreadIDArray(i)) then
			ho=int(trim(ThreadIDArray(i)))
			if Execute("Select ThreadID from ["&TablePrefix&"Threads] where ThreadID="&ho&"").eof then error"<li>ϵͳ�����ڸ����ӵ�����"
			if ThreadIDSql=0 then ThreadIDSql=ho
		else
			error("��������")
		end if
	next
else
	error("��������")
end if


SQL="Select ForumID From ["&TablePrefix&"Threads] where ThreadID="&ThreadIDSql&""
Set Rs=Execute(sql)
if Rs.eof then
	error("ϵͳ�����ڸ������ID")
else
	ForumID=Rs("ForumID")
end if
Rs.close
%>
<!-- #include file="Utility/ForumPermissions.asp" -->
<%
HtmlTop
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� �ƶ�����</div>

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
			<td align="center" colspan=2>�ƶ�����</td>
		</tr>
		<tr class="CommonListCell">
			<td align="right" width=30%>����ID��</td>
			<td><%=ThreadID%></td>
		</tr>
		<tr class="CommonListCell">
			<td align="right">�ƶ�����</td>
			<td>
			<select name="AimForumID" size=10 style="width:300px"><%=GroupList(0)%></select>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td colspan=2 align=center><input type="submit" value=" ȷ �� " />����<input onclick="history.back()" type="button" value=" �� �� " /></td>
		</tr>
	</table>
</form>

<script language="JavaScript" type="text/javascript">
function VerifyInput()
{
	if (document.form.AimForumID.value == "")
	{
		alert("��ѡ����Ҫ�ƶ����ĸ���̳");
		document.form.AimForumID.focus();
		return false;
	}
}
</script>
<%
HtmlBottom
%>