<!-- #include file="Setup.asp" -->
<%
select case Request("menu")

case "skins"
Response.Cookies("skins")=""&Server.HTMLEncode(Request("no"))&""
Response.Cookies("skins").Expires=date+9999

case "eremite"
Response.Cookies("eremite")="1"
Response.Cookies("eremite").Expires=date+9999

case "Online"
Response.Cookies("eremite")="0"
Response.Cookies("eremite").Expires=date+9999

end select

url=Request.ServerVariables("http_referer")
if url<>empty then
response.redirect url
else
response.write "<SCRIPT>top.location='"&SiteSettings("SiteURL")&"';</SCRIPT>"
end if
%>