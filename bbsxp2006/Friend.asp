<!-- #include file="Setup.asp" -->


<%
if CookieUserName=empty then error2("请登录后才能使用此功能！")
incept=HTMLEncode(Request("incept"))
UserName=HTMLEncode(Request("UserName"))


select case Request("menu")
case "add"
add
case "bad"
bad
case "Del"
Del
case "Post"
Post
case "look"
look
case "loadLog"
loadLog
case "addPost"
addPost
case ""
index
end select



sub add
if UserName="" then error2("请输入您要添加的好友名字！")

if UserName=CookieUserName then error2("不能添加自己！")

If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&UserName&"'" ).eof Then error2("数据库不存在此用户的资料！")

sql="select UserFriend from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Rs.Open sql,Conn,1,3
master=split(Rs("UserFriend"),"|")
if ubound(master) > 20 then error2("系统允许最多添加20个好友，您已经达到上限")
if instr(Rs("UserFriend"),"|"&UserName&"|")>0 then error2("此好友已经添加！")
Rs("UserFriend")=""&Rs("UserFriend")&""&UserName&"|"
Rs.update
Rs.close
error2("已经成功添加好友!")
end sub


sub Del
sql="select UserFriend from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Rs.Open sql,Conn,1,3
Rs("UserFriend")=replace(Rs("UserFriend"),"|"&UserName&"|","|")
Rs.update
Rs.close
index
end sub

sub look

Conn.execute("update [BBSXP_Users] set NewMessage=0 where UserName='"&CookieUserName&"'")

Page=Request("Page")
if Page<1 then
disabled="disabled=true"
Page=0
end if
count=Conn.execute("Select count(id)from [BBSXP_Messages] where incept='"&CookieUserName&"'")(0)
sql="select UserName,content from [BBSXP_Messages] where incept='"&CookieUserName&"' order by id Desc"
Set Rs=Conn.Execute(sql)
if Count-Page<2 then disabled2="disabled=true"

if Rs.eof then error2("您没有短讯息！")

Rs.Move Page
%>


<title>查看消息 - Powered By BBSXP</title>
<body topmargin=0>
<style>
.bt {BORDER-RIGHT: 1px; BORDER-TOP: 1px; FONT-SIZE: 9pt; BORDER-Left: 1px; BORDER-BOTTOM: 1px;}
</style>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><TD>&nbsp;昵称：<input readOnly type="text" value="<%=Rs("UserName")%>" size="8"> Email：<input  readOnly type="text" value="<%=Conn.Execute("Select UserMail From [BBSXP_Users] where UserName='"&Rs("UserName")&"'")(0)%>" size="10">
</TD><TD align=right><a target=_blank href=Profile.asp?UserName=<%=Rs("UserName")%>><img border="0" src="<%=Conn.Execute("Select Userface From [BBSXP_Users] where UserName='"&Rs("UserName")&"'")(0)%>" width="32" height="32" alt=用户详细资料>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="2" bgcolor="F8F8F8">
<script>
document.write ('<iframe id="HtmlEditor"  MARGINHEIGHT=1 MARGINWIDTH=1 height="100" width=100% frameborder=1></iframe>')
frames.HtmlEditor.document.write ('<%=Rs("content")%>');
frames.HtmlEditor.document.body.style.fontSize="10pt";
frames.HtmlEditor.document.close();
</script>

</TD></TR></TABLE>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD><input type="button" value="回复讯息" onclick=javascript:open('Friend.asp?menu=Post&incept=<%=Rs("UserName")%>','_top','width=320,height=170')>
</td><TD><input type="button" value="<<上一条" <%=disabled%> onclick=javascript:open('Friend.asp?menu=look&Page=<%=Page-1%>','_top','')> </td><TD><input type="button" value="下一条>>" <%=disabled2%> onclick=javascript:open('Friend.asp?menu=look&Page=<%=Page+1%>','_top','')>
</td>
</TR></TABLE>
<%
end sub

sub loadLog
sql="select * from [BBSXP_Messages] where (UserName='"&CookieUserName&"' and incept='"&incept&"') or (UserName='"&incept&"' and incept='"&CookieUserName&"') order by id Desc"
Set Rs=Conn.Execute(sql)
do while not Rs.eof
content=content&"<font color=red>("&Rs("DateCreated")&")　　"&Rs("UserName")&"</font><br><font color=0000FF>"&Rs("content")&"</font><br>"
Rs.Movenext
loop
Rs.close
if content="" then content="(空)"
%>
<script>
parent.Log.innerHTML='<iframe id="HtmlEditor"  MARGINHEIGHT=1 MARGINWIDTH=1 height="100" width=100% frameborder=1></iframe>'
parent.frames.HtmlEditor.document.write ('<%=content%>');
parent.frames.HtmlEditor.document.body.style.fontSize="10pt";
parent.frames.HtmlEditor.document.close();
</script>
<%
end sub

sub Post

if incept="" then error2("对不起，您没有输入用户名称！")

sql="select UserName,Userface,UserMail from [BBSXP_Users] where UserName='"&incept&"'"
Set Rs=Conn.Execute(sql)
if Rs.eof then error2("系统不存在该用户的资料")


%>

<body topmargin=0>
<style>
.bt {BORDER-RIGHT: 1px; BORDER-TOP: 1px; FONT-SIZE: 9pt; BORDER-Left: 1px; BORDER-BOTTOM: 1px;}
</style><title>发送消息 - Powered By BBSXP</title>

<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><form name=form action="Friend.asp" method="POST">
<input type="hidden" name="menu" value="addPost">
<input type="hidden" name="incept" value="<%=Rs("UserName")%>">
<TD>
&nbsp;昵称：<input readOnly type="text" value="<%=Rs("UserName")%>" size="8"> Email：<input  readOnly type="text" value="<%=Rs("UserMail")%>" size="10">
</TD><TD align=right><a target=_blank href=Profile.asp?UserName=<%=Rs("UserName")%>><img border="0" src="<%=Rs("Userface")%>" width="32" height="32" alt=用户详细资料>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="2" bgcolor="F8F8F8">
    <textarea name="content" cols="39" rows="6" onkeydown=presskey() onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')"><%=HTMLEncode(Request("content"))%></textarea>
</TD></TR></TABLE><TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD><input type="button" value="聊天记录" onclick=lookLog(); name=LookLogButton>
</td><TD><input type="reset" value="取消发送" OnClick="window.close();"> </td><TD><input type="submit" value="发送讯息" name=okButton onclick="return check(this.form)" id='btnadd' disabled></td>
</TR></form>
</TABLE>

<span id=Log></span>
<script>
function lookLog(){
resizeTo(330,330)
document.frames["hiddenframe"].location.replace("?menu=loadLog&incept=<%=incept%>");
document.form.LookLogButton.disabled = true;
}
function check(theForm) {
if (theForm.content.value.lengtd > 200){
alert("对不起，您的讯息不能超过 200 个字节！");
return false;
}
}
function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.form.submit();}}
<%if Request("Log")=1 then%>lookLog()<%end if%>
</script><iframe height=0 width=0 name=hiddenframe></iframe>
<%
Rs.close


end sub

sub addPost
if Len(Request.Form("content"))<2 then error2("内容不能小于 2 字符")
if incept=CookieUserName then error2("不能给自己发送讯息！")
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&incept&"'" ).eof Then error2("系统不存在"&incept&"的资料")

sql="insert into [BBSXP_Messages](UserName,incept,content) values ('"&CookieUserName&"','"&incept&"','"&HTMLEncode(Request.Form("content"))&"')"
Conn.Execute(SQL)
Conn.execute("update [BBSXP_Users] set NewMessage=NewMessage+1 where UserName='"&incept&"'")
%>
发送成功！<script>close();</script>
<%
end sub


sub index

top


sql="select UserName From [BBSXP_UsersOnline] where Eremite=0"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF
OnlineUserList=OnlineUserList&""&Rs("UserName")&"|"
Rs.MoveNext
loop
Rs.close


%>

<SCRIPT>
var Onlinelist= "<%=Onlinelist%>";

function add(){
var id=prompt("请输入您要添加的好友名称！","");
if(id){
document.location='Friend.asp?menu=add&UserName='+id+'';
}
}
</SCRIPT>


<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 控制面板</td>
</tr>
</table><br>
<table cellspacing=1 cellpadding=1 width=100% align=center border=0 class=a2>
  <TR id=TableTitleLink class=a1 height="25">
      <Td align="center"><b><a href="UserCp.asp">控制面板</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp">资料修改</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp?menu=pass">密码修改</a></b></td>
      <TD align="center"><b><a href="MySettings.asp">个性设置</a></b></td>
      <TD align="center"><b><a href="MyAttachment.asp">附件管理</a></b></td>
      <TD align="center"><b><a href="Message.asp">短信服务</a></b></td>
      <TD align="center"><b><a href="Friend.asp">好友列表</a></b></td>
      </TR></TABLE>
<br>
<form method="POST">
<input type=hidden name="menu" value="Del">

<table width="100%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>
  <tr>
    <td class=a1 align="center">
    昵称 </th>
    <td class=a1 align="center">
    邮件 </th>
    <td class=a1 align="center">
    主页 </th>
    <td class=a1 align="center">
    状态 </th>
    <td class=a1 align="center">
    发短信 </th>
    <td class=a1 align="center">
    操作 </th>
  </tr>
<%

on error resume next '找不到好友资料时候忽略错误

sql="select UserFriend,Userface from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Set Rs=Conn.Execute(sql)
master=split(Rs("UserFriend"),"|")
for i = 1 to ubound(master)-1


Set Rs1=Conn.Execute("[BBSXP_Users] where UserName='"&master(i)&"'")
UserMail=Rs1("UserMail")
Userhome=Rs1("Userhome")
Set Rs1 = Nothing


%>
  <tr class=a3>
    <td vAlign=center align=middle><a href=Profile.asp?UserName=<%=master(i)%>><%=master(i)%></a>
    <td align=middle><a href=Mailto:<%=UserMail%>><%=UserMail%></a>
    <td><a href=<%=Userhome%> target=_blank><%=Userhome%></a>
    <td align=middle><%if instr("|"&OnlineUserList&"","|"&master(i)&"|")>0 then%> 在 线 <%else%> 离 线 <%end if%></td>
    <td align=middle><a style=cursor:hand onclick="javascript:open('Friend.asp?menu=Post&incept=<%=master(i)%>','','width=320,height=170')">发送</a></td>
    <td align=middle><INPUT type=radio value=<%=master(i)%> name=UserName></td>
  </tr>
  
  
  

<%
next
Rs.close
%>


  
  

  
  <tr class=a3>
    <td vAlign="center" align="right" colSpan="6">
    <input onclick="javascript:add();" type="button" value="添加好友" name="action">&nbsp;<input onclick="checkclick('确定删除选定的好友吗?');" type="submit" value="删除"></td>
  </tr></form>
</table>

<%

htmlend
end sub

CloseDatabase
%>