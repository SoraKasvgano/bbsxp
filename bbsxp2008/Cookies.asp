<!-- #include file="setup.asp" -->
<%
if Http_Referer="" then Http_Referer="Default.asp"
if Request_Method <> "POST" then error("<li>提交方式错误！</li><li>您本次使用的是"&Request_Method&"提交方式！</li>")

select case Request("menu")
	case "Invisible"
		ResponseCookies "Invisible","1","999"
	case "Online"
		ResponseCookies "Invisible","0","999"
end select
response.redirect Http_Referer
%>