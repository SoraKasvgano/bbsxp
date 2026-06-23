<!-- #include file="Setup.asp" -->
<%
select case Request("menu")

case "skins"
Response.Cookies("skins")=SafeThemeName(Request("no"))
Response.Cookies("skins").Expires=date+9999

case "eremite"
Response.Cookies("eremite")="1"
Response.Cookies("eremite").Expires=date+9999

case "Online"
Response.Cookies("eremite")="0"
Response.Cookies("eremite").Expires=date+9999

end select

url=SafeRedirectUrl(Request.ServerVariables("http_referer"))
if url<>"" then
response.redirect url
else
response.write "<SCRIPT>top.location='"&SafeJsString(SiteSettings("SiteURL"))&"';</SCRIPT>"
end if
%>