<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error2("���¼���ٽ��в�����")

if Request.ServerVariables("request_method") <> "POST" then
response.write "<form name=BBSXPPost method=Post action=Manage.asp?"&Request.ServerVariables("Query_String")&"></form><SCRIPT>if(confirm('��ȷ��Ҫִ�иò���?')){returnValue=BBSXPPost.submit()}else{returnValue=history.back()}</SCRIPT>"
htmlend
end if


top

PostID=RequestInt("PostID")
ThreadID=RequestInt("ThreadID")
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
elseif instr("|"&Conn.Execute("Select moderated From [BBSXP_Forums] where id="&ForumID&" ")(0)&"|","|"&SqlString(CookieUserName)&"|")>0 then
pass=1
elseif Replies=0 and UserName=CookieUserName and Request("menu")="IsDel" then
pass=1
end if

if pass<>1 then error("<li>����Ȩ�޲���")



select case Request("menu")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "top"
if membercode > 3 then
Conn.execute("update [BBSXP_Threads] set IsTop=2 where id="&ThreadID&"")
succtitle="�̶ܹ�����ɹ�"
else
error("<li>����Ȩ�޲���")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "untop"
if membercode > 3 then
Conn.execute("update [BBSXP_Threads] set IsTop=0 where id="&ThreadID&"")
succtitle="ȡ���̶ܹ�����ɹ�"
else
error("<li>����Ȩ�޲���")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "Fix"
TotalCount=conn.Execute("Select count(ID) From [BBSXP_Posts"&PostsTableName&"] where ThreadID="&ThreadID&"")(0)
Conn.execute("update [BBSXP_Threads] set replies="&TotalCount&"-1 where id="&ThreadID&"")
succtitle="�޸�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "MoveNew"
Conn.execute("update [BBSXP_Threads] set lasttime="&SqlNowString&" where id="&ThreadID&"")
succtitle="��ǰ����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "Move"
AimForumID=RequestInt("AimForumID")
if AimForumID="" then error("<li>��û��ѡ��Ҫ�������ƶ��ĸ���̳")
if Conn.Execute("Select ForumPass From [BBSXP_Forums] where id="&AimForumID&"")(0)=4 then error("<li>Ŀ����̳Ϊ��Ȩ����״̬")
Conn.execute("update [BBSXP_Threads] set ForumID="&AimForumID&",IsTop=0,IsGood=0,IsLocked=0 where id="&ThreadID&"")
succtitle="�ƶ�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "IsDel"
if Conn.Execute("Select IsDel From [BBSXP_Threads] where id="&ThreadID&" ")(0)=1 then error("<li>�������Ѿ��ڻ���վ��")
sql="select * from [BBSXP_Posts"&PostsTableName&"] where id="&PostID&""
Rs.Open sql,Conn,1,3
if Rs("IsTopic")=1 then
Conn.execute("update [BBSXP_Users] set DelTopic=DelTopic+1,UserMoney=UserMoney+"&SiteSettings("IntegralDeleteThread")&",experience=experience+"&SiteSettings("IntegralDeleteThread")&" where UserName='"&Rs("UserName")&"'")
Conn.execute("update [BBSXP_Threads] set IsTop=0,IsDel=1,lastname='"&CookieUserName&"',lasttime="&SqlNowString&" where id="&ThreadID&"")
Conn.execute("update [BBSXP_Forums] set ForumThreads=ForumThreads-1,ForumPosts=ForumPosts-1 where id="&ForumID&"")
succtitle="ɾ������ɹ�"
else
Conn.execute("update [BBSXP_Users] set DelTopic=DelTopic+1,UserMoney=UserMoney+"&SiteSettings("IntegralDeletePost")&",experience=experience+"&SiteSettings("IntegralDeletePost")&" where UserName='"&Rs("UserName")&"'")
Conn.execute("update [BBSXP_Threads] set replies=replies-1 where id="&ThreadID&"")
Conn.execute("update [BBSXP_Forums] set ForumPosts=ForumPosts-1 where id="&ForumID&"")
Rs.Delete
succtitle="ɾ�������ɹ�"
end if
Rs.update
Rs.close
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "IsGood"
if Conn.Execute("Select IsGood From [BBSXP_Threads] where id="&ThreadID&" ")(0)=1 then error("<li>�������Ѿ����뾫�����ˣ������ظ�����")
Conn.execute("update [BBSXP_Threads] set IsGood=1 where id="&ThreadID&"")
Conn.execute("update [BBSXP_Users] set goodTopic=goodTopic+1,UserMoney=UserMoney+"&SiteSettings("IntegralAddValuedPost")&",experience=experience+"&SiteSettings("IntegralAddValuedPost")&" where UserName='"&UserName&"'")
succtitle="��Ϊ�������ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "DelIsGood"
if Conn.Execute("Select IsGood From [BBSXP_Threads] where id="&ThreadID&" ")(0)=0 then error("<li>�������Ѿ��Ƴ���������")
Conn.execute("update [BBSXP_Threads] set IsGood=0 where id="&ThreadID&"")
Conn.execute("update [BBSXP_Users] set goodTopic=goodTopic-1,UserMoney=UserMoney+"&SiteSettings("IntegralDeleteValuedPost")&",experience=experience+"&SiteSettings("IntegralDeleteValuedPost")&" where UserName='"&UserName&"'")
succtitle="ȡ���������ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "IsTop"
Conn.execute("update [BBSXP_Threads] set IsTop=1 where id="&ThreadID&"")
succtitle="�̶�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "DelIsTop"
Conn.execute("update [BBSXP_Threads] set IsTop=0 where id="&ThreadID&"")
succtitle="ȡ���̶�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "IsLocked"
Conn.execute("update [BBSXP_Threads] set IsLocked=1 where id="&ThreadID&"")
succtitle="��������ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "DelIsLocked"
Conn.execute("update [BBSXP_Threads] set IsLocked=0 where id="&ThreadID&"")
succtitle="��������ɹ�"
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
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� �鿴IP</td>
</tr>
</table>
<br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr>
<td width="328" height="25" align="center" class=a1 colspan="2">
�鿴IP
</td></tr><tr>
<td height="7" width="164" valign="top" align="center" class=a3>
�û���</td>
<td height="7" width="164" valign="top" align="center" class=a3>
<%=UserName%></td></tr><tr>
<td height="6" width="164" valign="top" align="center" class=a3>
ʱ��</td>
<td height="6" width="164" valign="top" align="center" class=a3>
<%=Posttime%></td></tr><tr>
<td height="6" width="164" valign="top" align="center" class=a3>
IP��ַ</td>
<td height="6" width="164" valign="top" align="center" class=a3>
<%=Postip%></td></tr></table>

<br>
<center>
<a href=ShowPost.asp?ThreadID=<%=ThreadID%>>BACK</a><br>
<%
htmlend

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
end select
if succtitle="" then error("<li>��Ч����")

Log(""&succtitle&"������ID��"&ThreadID&"")

Message="<li>"&succtitle&"<li><a href=ShowForum.asp?ForumID="&ForumID&">������̳</a><li><a href=Default.asp>����������ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?ForumID="&ForumID&">")
%>