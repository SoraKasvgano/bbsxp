<!-- #include file="conn.asp" -->
<%
if Request_Method <> "POST" then Response.End

RandomID=""&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&""

TargetUser=HTMLEncode(unescape(Request.QueryString("TargetUser")))

if CookieUserName="" then ConciseMsg("请登录后再进行操作")
select case Request("menu")
	case "MarryRequest"
		if CookieUserName=TargetUser then ConciseMsg("不能自己与自己结婚")
		if CookieUserMate<>"" then ConciseMsg("您已经有配偶了")
		if Execute("Select UserMate From ["&TablePrefix&"Users] where UserName='"&TargetUser&"'")(0)<>"" then  ConciseMsg(""&TargetUser&"已经有配偶了")
		AddApplication "Message_"&TargetUser,"【系统讯息】<span id="&RandomID&">"&CookieUserName&"向您求婚！请选择 <a href=javascript:Ajax_CallBack(false,'"&RandomID&"','UserMate.asp?menu=marry&TargetUser="&CookieUserName&"');>接受</a> 或 <a href=javascript:Ajax_CallBack(false,'"&RandomID&"','UserMate.asp?menu=refused&TargetUser="&CookieUserName&"');>拒绝</a> ！</span>"
		ConciseMsg("求婚请求已发送")

	case "marry"
		if CookieUserName=TargetUser then ConciseMsg("不能自己与自己结婚")
		if CookieUserMate<>"" then ConciseMsg("您已经有配偶了")	
		if Execute("Select UserMate From ["&TablePrefix&"Users] where UserName='"&TargetUser&"'")(0)<>"" then  ConciseMsg(""&TargetUser&"已经有配偶了")	
		Execute("Update ["&TablePrefix&"Users] Set UserMate='"&TargetUser&"' where UserName='"&CookieUserName&"'")
		Execute("Update ["&TablePrefix&"Users] Set UserMate='"&CookieUserName&"' where UserName='"&TargetUser&"'")
		AddApplication "Message_"&TargetUser,"【系统讯息】"&CookieUserName&"与您结婚了"
		ConciseMsg("您已经与 "&TargetUser&" 的结婚了")

	case "refused"
		AddApplication "Message_"&TargetUser,"【系统讯息】"&CookieUserName&"拒绝与您结婚"
		ConciseMsg("您已经拒绝了 "&TargetUser&" 的结婚请求")

	case "divorce"
		Execute("Update ["&TablePrefix&"Users] Set UserMate='' where UserName='"&CookieUserName&"' or UserName='"&CookieUserMate&"' ")
		AddApplication "Message_"&CookieUserMate,"【系统讯息】"&CookieUserName&"已经与您离婚"
		ConciseMsg("离婚成功")
	
end select
%>