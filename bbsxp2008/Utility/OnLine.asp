<%
if CookieUserName<>empty then SqlLineUser=" and UserName='"&CookieUserName&"'"
sql="Select * from ["&TablePrefix&"UserOnline] where IPAddress='"&REMOTE_ADDR&"' "&SqlLineUser&" or sessionid='"&session.sessionid&"' "

if ThreadID<1 then ThreadID=0
Rs.Open sql,Conn,1,3
	if Rs.eof then Rs.addNew
	Rs("UserName")=CookieUserName
	Rs("ThreadID")=ThreadID
	Rs("Topic")=""&Topic&""
	Rs("ForumID")=ForumID
	Rs("ForumName")=""&ForumName&""
	Rs("LastTime")=now()
	if RequestCookies("Invisible")<>empty then Rs("IsInvisible")=RequestCookies("Invisible")
	Rs("IPAddress")=REMOTE_ADDR
	Rs("sessionid")=session.sessionid
Rs.update
Rs.close

if CookieUserID > 0 then Execute("update ["&TablePrefix&"Users] Set UserActivityTime="&SqlNowString&",UserActivityIP='"&REMOTE_ADDR&"' where UserID="&CookieUserID&"")

Execute("Delete from ["&TablePrefix&"UserOnline] where DateDiff("&SqlChar&"n"&SqlChar&",lasttime,"&SqlNowString&")>"&SiteConfig("UserOnlineTime")&" ")
Onlinemany=Execute("Select count(sessionid) from ["&TablePrefix&"UserOnline]")(0)
%>