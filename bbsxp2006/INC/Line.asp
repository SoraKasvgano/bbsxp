<%
sql="select * from [BBSXP_UsersOnline] where ip='"&Request.ServerVariables("REMOTE_ADDR")&"' and UserName='"&CookieUserName&"' or sessionid='"&session.sessionid&"' "
Rs.Open sql,Conn,1,3
if Rs.eof then Rs.addNew
Rs("act")=""&act&""
Rs("acturl")=""&Request.ServerVariables("script_name")&"?"&Request.ServerVariables("Query_String")&""
Rs("ForumID")=ForumID
Rs("ForumName")=""&ForumName&""
Rs("lasttime")=now
if Request.Cookies("eremite")<>empty then Rs("eremite")=Request.Cookies("eremite")
Rs("Userface")=membercode
Rs("ip")=Request.ServerVariables("REMOTE_ADDR")
Rs("UserName")=CookieUserName
Rs("sessionid")=session.sessionid
Rs.update
Rs.close
Conn.execute("Delete from [BBSXP_UsersOnline] where DateDiff("&SqlChar&"n"&SqlChar&",lasttime,"&SqlNowString&")>"&SiteSettings("UserOnlineTime")&" ")
Onlinemany=Conn.execute("Select count(sessionid)from [BBSXP_UsersOnline]")(0)
%>