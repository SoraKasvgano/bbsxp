<!-- #include file="Setup.asp" -->
<%
id=int(Request("id"))
ForumID=HTMLEncode(Request("ForumID"))


if id="0" then


if Request("ForumID")="0" then
sql="select * from [BBSXP_UsersOnline] where UserName<>'' and eremite<>1"
else
sql="select * from [BBSXP_UsersOnline] where ForumID="&ForumID&" and UserName<>'' and eremite<>1"
end if

Set Rs=Conn.Execute(sql)

do while not Rs.eof

if NO_count < 5 then
NO_count=NO_count+1
else
NO_count=1
end if


allline=""&allline&"<td width=16% style=word-break:break-all><img src=images/membercode/"&Rs("Userface")&".gif> <a href=Profile.asp?UserName="&Rs("UserName")&">"&Rs("UserName")&"</a></td>"

if NO_count = 5 then allline=""&allline&"</tr><tr>"

Rs.Movenext
loop
Rs.close

%>
<SCRIPT>
var parentfollowImg=parent.document.getElementById("followImg0")
var parentfollowTd=parent.document.getElementById("followTd0")
parentfollowImg.loaded="yes";
parentfollowTd.innerHTML="<TABLE cellSpacing=3 cellPadding=0 width=100% border=0><TR><%=allline%></TR></TABLE>";
</SCRIPT>
<%

else

sql="Select * From [BBSXP_Threads] where id="&id&""
Set Rs=Conn.Execute(sql)
PostsTableName=rs("PostsTableName")
ForumID=rs("ForumID")
rs.close
content=Conn.Execute("Select content From [BBSXP_Posts"&PostsTableName&"] where ThreadID="&id&" and IsTopic=1")(0)

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumPass=Rs("ForumPass")
ForumPassword=Rs("ForumPassword")
ForumUserList=Rs("ForumUserList")
Rs.close

if ForumPass="0" then
content="本论坛暂时关闭，不再接受访问！"
elseif ForumPass="2" then
if CookieUserName=empty then content="本论坛禁止游客访问！"
elseif ForumPass="3" then
if instr("|"&ForumUserList&"|","|"&CookieUserName&"|")<=0 or CookieUserName=empty then  content="本论坛需要授权才能访问！"
end if

if ForumPassword<>empty then
if ForumPassword<>Request.Cookies("password") then content="本论坛需要密码才能访问！"
end if
%>
<SCRIPT>
var parentfollowImg=parent.document.getElementById("followImg<%=id%>")
var parentfollowTd=parent.document.getElementById("followTd<%=id%>")
parentfollowImg.loaded="yes";
parentfollowTd.innerHTML='<%=content%>';
</SCRIPT>
<%
end if




CloseDatabase
%>