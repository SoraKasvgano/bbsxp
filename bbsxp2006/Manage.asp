<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error2("请登录后再进行操作！")

if Request.ServerVariables("request_method") <> "POST" then
response.write "<form name=BBSXPPost method=Post action=Manage.asp?"&Request.ServerVariables("Query_String")&"></form><SCRIPT>if(confirm('您确定要执行该操作?')){returnValue=BBSXPPost.submit()}else{returnValue=history.back()}</SCRIPT>"
htmlend
end if


top

PostID=int(Request("PostID"))
ThreadID=int(Request("ThreadID"))
sql="Select * From [BBSXP_Threads] where ID="&ThreadID&""
Rs.Open sql,Conn,1
ForumID=Rs("ForumID")
PostsTableName=Rs("PostsTableName")
UserName=Rs("UserName")
ThreadID=Rs("id")
Replies=Rs("Replies")
Rs.close



if membercode > 3 then
pass=1
elseif instr("|"&Conn.Execute("Select moderated From [BBSXP_Forums] where id="&ForumID&" ")(0)&"|","|"&CookieUserName&"|")>0 then
pass=1
elseif Replies=0 and UserName=CookieUserName and Request("menu")="IsDel" then
pass=1
end if

if pass<>1 then error("<li>您的权限不够")



select case Request("menu")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "top"
if membercode > 3 then
Conn.execute("update [BBSXP_Threads] set IsTop=2 where id="&ThreadID&"")
succtitle="总固顶主题成功"
else
error("<li>您的权限不够")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "untop"
if membercode > 3 then
Conn.execute("update [BBSXP_Threads] set IsTop=0 where id="&ThreadID&"")
succtitle="取消总固顶主题成功"
else
error("<li>您的权限不够")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "Fix"
TotalCount=conn.Execute("Select count(ID) From [BBSXP_Posts"&PostsTableName&"] where ThreadID="&ThreadID&"")(0)
Conn.execute("update [BBSXP_Threads] set replies="&TotalCount&"-1 where id="&ThreadID&"")
succtitle="修复主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "MoveNew"
Conn.execute("update [BBSXP_Threads] set lasttime="&SqlNowString&" where id="&ThreadID&"")
succtitle="拉前主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "Move"
AimForumID=int(Request("AimForumID"))
if AimForumID="" then error("<li>您没有选择要将主题移动哪个论坛")
if Conn.Execute("Select ForumPass From [BBSXP_Forums] where id="&AimForumID&"")(0)=4 then error("<li>目标论坛为授权发帖状态")
Conn.execute("update [BBSXP_Threads] set ForumID="&AimForumID&",IsTop=0,IsGood=0,IsLocked=0 where id="&ThreadID&"")
succtitle="移动主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "IsDel"
if Conn.Execute("Select IsDel From [BBSXP_Threads] where id="&ThreadID&" ")(0)=1 then error("<li>此帖子已经在回收站了")
sql="select * from [BBSXP_Posts"&PostsTableName&"] where id="&PostID&""
Rs.Open sql,Conn,1,3
if Rs("IsTopic")=1 then
Conn.execute("update [BBSXP_Users] set DelTopic=DelTopic+1,UserMoney=UserMoney+"&SiteSettings("IntegralDeleteThread")&",experience=experience+"&SiteSettings("IntegralDeleteThread")&" where UserName='"&Rs("UserName")&"'")
Conn.execute("update [BBSXP_Threads] set IsTop=0,IsDel=1,lastname='"&CookieUserName&"',lasttime="&SqlNowString&" where id="&ThreadID&"")
Conn.execute("update [BBSXP_Forums] set ForumThreads=ForumThreads-1,ForumPosts=ForumPosts-1 where id="&ForumID&"")
succtitle="删除主题成功"
else
Conn.execute("update [BBSXP_Users] set DelTopic=DelTopic+1,UserMoney=UserMoney+"&SiteSettings("IntegralDeletePost")&",experience=experience+"&SiteSettings("IntegralDeletePost")&" where UserName='"&Rs("UserName")&"'")
Conn.execute("update [BBSXP_Threads] set replies=replies-1 where id="&ThreadID&"")
Conn.execute("update [BBSXP_Forums] set ForumPosts=ForumPosts-1 where id="&ForumID&"")
Rs.Delete
succtitle="删除回帖成功"
end if
Rs.update
Rs.close
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "IsGood"
if Conn.Execute("Select IsGood From [BBSXP_Threads] where id="&ThreadID&" ")(0)=1 then error("<li>此帖子已经加入精华区了，无需重复添加")
Conn.execute("update [BBSXP_Threads] set IsGood=1 where id="&ThreadID&"")
Conn.execute("update [BBSXP_Users] set goodTopic=goodTopic+1,UserMoney=UserMoney+"&SiteSettings("IntegralAddValuedPost")&",experience=experience+"&SiteSettings("IntegralAddValuedPost")&" where UserName='"&UserName&"'")
succtitle="加为精华帖成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "DelIsGood"
if Conn.Execute("Select IsGood From [BBSXP_Threads] where id="&ThreadID&" ")(0)=0 then error("<li>此帖子已经移出精华区了")
Conn.execute("update [BBSXP_Threads] set IsGood=0 where id="&ThreadID&"")
Conn.execute("update [BBSXP_Users] set goodTopic=goodTopic-1,UserMoney=UserMoney+"&SiteSettings("IntegralDeleteValuedPost")&",experience=experience+"&SiteSettings("IntegralDeleteValuedPost")&" where UserName='"&UserName&"'")
succtitle="取消精华帖成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "IsTop"
Conn.execute("update [BBSXP_Threads] set IsTop=1 where id="&ThreadID&"")
succtitle="固顶主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "DelIsTop"
Conn.execute("update [BBSXP_Threads] set IsTop=0 where id="&ThreadID&"")
succtitle="取消固顶主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "IsLocked"
Conn.execute("update [BBSXP_Threads] set IsLocked=1 where id="&ThreadID&"")
succtitle="锁定主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "DelIsLocked"
Conn.execute("update [BBSXP_Threads] set IsLocked=0 where id="&ThreadID&"")
succtitle="解锁主题成功"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "lookip"
sql="Select * From [BBSXP_Posts"&PostsTableName&"] where id="&PostID&""
Rs.Open sql,Conn,1
UserName=Rs("UserName")
Posttime=Rs("Posttime")
Postip=Rs("Postip")
Rs.close
%>


<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 查看IP</td>
</tr>
</table>
<br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr>
<td width="328" height="25" align="center" class=a1 colspan="2">
查看IP
</td></tr><tr>
<td height="7" width="164" valign="top" align="center" class=a3>
用户名</td>
<td height="7" width="164" valign="top" align="center" class=a3>
<%=UserName%></td></tr><tr>
<td height="6" width="164" valign="top" align="center" class=a3>
时间</td>
<td height="6" width="164" valign="top" align="center" class=a3>
<%=Posttime%></td></tr><tr>
<td height="6" width="164" valign="top" align="center" class=a3>
IP地址</td>
<td height="6" width="164" valign="top" align="center" class=a3>
<%=Postip%></td></tr></table>

<br>
<center>
<a href=ShowPost.asp?ThreadID=<%=ThreadID%>>BACK</a><br>
<%
htmlend

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
end select
if succtitle="" then error("<li>无效命令")

Log(""&succtitle&"，主题ID："&ThreadID&"")

Message="<li>"&succtitle&"<li><a href=ShowForum.asp?ForumID="&ForumID&">返回论坛</a><li><a href=Default.asp>返回社区首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?ForumID="&ForumID&">")
%>