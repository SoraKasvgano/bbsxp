<%
if ForumPass="0" then
error("<li>本论坛暂时关闭，不再接受访问！")
elseif ForumPass="2" then
if CookieUserName=empty then error("<li>只有<a href=Login.asp>登录</a>后才能浏览论坛")
elseif ForumPass="3" then
if instr("|"&ForumUserList&"|","|"&CookieUserName&"|")<=0 or CookieUserName=empty then error("<li>该论坛并未授权您访问！")
elseif ForumPass="4" then
if instr(""&Request.ServerVariables("script_name")&"","NewTopic.asp") > 0 and instr("|"&ForumUserList&"|","|"&CookieUserName&"|")<=0 then error("<li>该论坛并未授权您发帖！")
end if

if ForumPassword<>empty then
if ForumPassword<>Request.Cookies("password") then response.redirect "Login.asp?menu=password&url=http://"&Request.ServerVariables("server_name")&Request.ServerVariables("script_name")&"?"&Request.ServerVariables("Query_String")&""
end if
%>