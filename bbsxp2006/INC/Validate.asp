<%
if ForumPass="0" then
error("<li>����̳��ʱ�رգ����ٽ��ܷ��ʣ�")
elseif ForumPass="2" then
if CookieUserName=empty then error("<li>ֻ��<a href=Login.asp>��¼</a>����������̳")
elseif ForumPass="3" then
if instr("|"&ForumUserList&"|","|"&CookieUserName&"|")<=0 or CookieUserName=empty then error("<li>����̳��δ��Ȩ�����ʣ�")
elseif ForumPass="4" then
if instr(""&Request.ServerVariables("script_name")&"","NewTopic.asp") > 0 and instr("|"&ForumUserList&"|","|"&CookieUserName&"|")<=0 then error("<li>����̳��δ��Ȩ��������")
end if

if ForumPassword<>empty then
if ForumPassword<>Request.Cookies("password") then response.redirect "Login.asp?menu=password&url=http://"&Request.ServerVariables("server_name")&Request.ServerVariables("script_name")&"?"&Request.ServerVariables("Query_String")&""
end if
%>