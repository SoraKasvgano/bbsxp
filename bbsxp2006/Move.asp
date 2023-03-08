<!-- #include file="Setup.asp" -->
<%

ThreadID=int(Request("ThreadID"))

ForumID=Conn.Execute("Select ForumID From [BBSXP_Threads] where id="&ThreadID)(0)

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumName=Rs("ForumName")
ForumLogo=Rs("ForumLogo")
followid=Rs("followid")
Rs.close


top
sql="select * from [BBSXP_Threads] where ID="&ThreadID&" and ForumID="&ForumID&""
Set Rs=Conn.Execute(sql)
%>
<script>
if ("<%=ForumLogo%>"!=''){Logo.innerHTML="<img border=0 src=<%=ForumLogo%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>
	<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> → <a href="ShowPost.asp?ThreadID=<%=ThreadID%>"><%=Rs("Topic")%></a> → 移动帖子</td>
		</tr>
	</table><br>

<form method="POST" action=Manage.asp>
<input type="hidden" value="Move" name=menu>
<input type="hidden" value="<%=ThreadID%>" name="ThreadID">

<table width="50%" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr><td class=a1 height="25" align="center">移动帖子</td></tr><tr>
<td height="19" valign="top" class=a4 align="center">
<p><br>
	<select name=AimForumID>
<option selected value="">将主题移动到...</option>

<%
Rs.close
%>
<%BBSList(0)%><%=ForumsList%>
</select>&nbsp;&nbsp;

<input type=submit value=" 确 定 "> <br>
<br></td></tr></table>
</form>





<%
htmlend
%>