<!-- #include file="conn.asp" -->
<%
if Request_Method <> "POST" then Response.End

RandomID=""&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&""

TargetUser=HTMLEncode(unescape(Request.QueryString("TargetUser")))
CookieUserNameHtml=Server.HTMLEncode(CookieUserName)
CookieUserNameUrl=Server.URLEncode(CookieUserName)
TargetUserHtml=Server.HTMLEncode(TargetUser)

if CookieUserName="" then ConciseMsg("请登录后再进行操作")
select case Request("menu")
	case "MarryRequest"
		if CookieUserName=TargetUser then ConciseMsg("不能自己与自己结婚")
		if CookieUserMate<>"" then ConciseMsg("您已经有配偶了")
		Set MateRs=Execute("Select UserMate From ["&TablePrefix&"Users] where UserName='"&SqlString(TargetUser)&"'")
		if MateRs.eof then ConciseMsg("用户不存在")
		if MateRs(0)<>"" then  ConciseMsg(""&TargetUserHtml&"已经有配偶了")
		AddApplication "Message_"&TargetUser,"【系统讯息】<span id="&RandomID&">"&CookieUserNameHtml&"向您求婚！请选择 <a href=javascript:Ajax_CallBack(false,'"&RandomID&"','UserMate.asp?menu=marry&TargetUser="&CookieUserNameUrl&"');>接受</a> 或 <a href=javascript:Ajax_CallBack(false,'"&RandomID&"','UserMate.asp?menu=refused&TargetUser="&CookieUserNameUrl&"');>拒绝</a> ！</span>"
		ConciseMsg("求婚请求已发送")

	case "marry"
		if CookieUserName=TargetUser then ConciseMsg("不能自己与自己结婚")
		if CookieUserMate<>"" then ConciseMsg("您已经有配偶了")
		Set MateRs=Execute("Select UserMate From ["&TablePrefix&"Users] where UserName='"&SqlString(TargetUser)&"'")
		if MateRs.eof then ConciseMsg("用户不存在")
		if MateRs(0)<>"" then  ConciseMsg(""&TargetUserHtml&"已经有配偶了")
		Execute("Update ["&TablePrefix&"Users] Set UserMate='"&SqlString(TargetUser)&"' where UserName='"&SqlString(CookieUserName)&"'")
		Execute("Update ["&TablePrefix&"Users] Set UserMate='"&SqlString(CookieUserName)&"' where UserName='"&SqlString(TargetUser)&"'")
		AddApplication "Message_"&TargetUser,"【系统讯息】"&CookieUserNameHtml&"与您结婚了"
		ConciseMsg("您已经与 "&TargetUserHtml&" 的结婚了")

	case "refused"
		AddApplication "Message_"&TargetUser,"【系统讯息】"&CookieUserNameHtml&"拒绝与您结婚"
		ConciseMsg("您已经拒绝了 "&TargetUserHtml&" 的结婚请求")

	case "divorce"
		Execute("Update ["&TablePrefix&"Users] Set UserMate='' where UserName='"&SqlString(CookieUserName)&"' or UserName='"&SqlString(CookieUserMate)&"' ")
		AddApplication "Message_"&CookieUserMate,"【系统讯息】"&CookieUserNameHtml&"已经与您离婚"
		ConciseMsg("离婚成功")

end select
%>