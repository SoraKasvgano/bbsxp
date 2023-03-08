<!-- #include file="Setup.asp" -->
<%
ThreadID=int(Request("ThreadID"))

sql="Select * From [BBSXP_Threads] where ID="&ThreadID&""
Rs.Open sql,Conn,1
ForumID=Rs("ForumID")
PostsTableName=Rs("PostsTableName")
Topic=Rs("Topic")
Rs.close
Topic=ReplaceText(Topic,"<[^>]*>","")

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumName=Rs("ForumName")
moderated=Rs("moderated")
ForumPass=Rs("ForumPass")
ForumPassword=Rs("ForumPassword")
ForumUserList=Rs("ForumUserList")
Rs.close
%>
<!-- #include file="inc/Validate.asp" -->
<body onload="window.print()">
<title><%=Topic%></title>
<table cellpadding=0 cellspacing=0 border=0 width=100% align=center><tr><td>
- <b><%=SiteSettings("SiteName")%></b> ( <%=SiteSettings("SiteURL")%>Default.asp )<br>

-- <b><%=ForumName%></b> ( <%=SiteSettings("SiteURL")%>ShowForum.asp?ForumID=<%=ForumID%> )<br>

--- <b><%=Topic%></b> ( <%=SiteSettings("SiteURL")%>ShowPost.asp?ThreadID=<%=ThreadID%> )
<br><br>
<hr>
<%

sql="select * from [BBSXP_Posts"&PostsTableName&"] where ThreadID="&ThreadID&" order by id"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF
%><p>
作者：<%=Rs("UserName")%><br>
发表时间：<%=Rs("Posttime")%><br><br>
<%=Rs("content")%></p><hr>
<%
Rs.MoveNext
loop
Rs.Close


htmlend
%>